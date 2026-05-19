"""XML-level xlsx merge.

xlsx is a ZIP of XML parts (ECMA-376). This module reads source xlsx files,
remaps their style/sharedString indices into a single combined workbook, and
emits the merged file as raw zip bytes.

The summary sheet builder (Phase 2) will produce its own sheet XML that
references the combined styles/strings, then this module zips everything up.
"""
from __future__ import annotations

import re
import zipfile
from copy import deepcopy
from io import BytesIO
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NS_MAP = {"x": MAIN_NS, "r": REL_NS, "rel": PKG_REL_NS, "ct": CT_NS}

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def _q(ns_key: str, tag: str) -> str:
    return f"{{{NS_MAP[ns_key]}}}{tag}"


def _xml_bytes(root: ET.Element, standalone: bool = True) -> bytes:
    """Serialize an Element tree with the OOXML-friendly XML declaration."""
    body = ET.tostring(root, encoding="utf-8")
    decl = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        if standalone
        else b'<?xml version="1.0" encoding="UTF-8"?>\n'
    )
    return decl + body


class XlsxParts:
    """All parts of an xlsx file held in memory as raw bytes.

    Use `from_bytes` to load and `to_bytes` to emit. Parse/replace individual
    parts (XML or binary) via dict-like access on `parts`.
    """

    def __init__(self) -> None:
        self.parts: dict[str, bytes] = {}

    @classmethod
    def from_bytes(cls, data: bytes) -> "XlsxParts":
        self = cls()
        with zipfile.ZipFile(BytesIO(data)) as z:
            for name in z.namelist():
                self.parts[name] = z.read(name)
        return self

    def to_bytes(self) -> bytes:
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in self.parts.items():
                z.writestr(name, data)
        return buf.getvalue()

    def has(self, path: str) -> bool:
        return path in self.parts

    def get_xml(self, path: str) -> ET.Element | None:
        data = self.parts.get(path)
        if data is None:
            return None
        return ET.fromstring(data)

    def set_xml(self, path: str, root: ET.Element, standalone: bool = True) -> None:
        self.parts[path] = _xml_bytes(root, standalone=standalone)


CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def col_letter_to_index(letters: str) -> int:
    """A → 1, B → 2, AA → 27."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def col_index_to_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def parse_cell_ref(ref: str) -> tuple[int, int]:
    """'B5' → (col=2, row=5)."""
    m = CELL_REF_RE.match(ref)
    if not m:
        raise ValueError(f"bad cell ref: {ref}")
    return col_letter_to_index(m.group(1)), int(m.group(2))


# ---------------------------------------------------------------------------
# Styles / shared strings merge
# ---------------------------------------------------------------------------


def _elem_equal(a: ET.Element, b: ET.Element) -> bool:
    """Deep equality of two XML elements (text + attrib + children, order-sensitive)."""
    if a.tag != b.tag:
        return False
    if (a.text or "").strip() != (b.text or "").strip():
        return False
    if dict(a.attrib) != dict(b.attrib):
        return False
    ac, bc = list(a), list(b)
    if len(ac) != len(bc):
        return False
    return all(_elem_equal(x, y) for x, y in zip(ac, bc))


def _find_or_append(parent: ET.Element, child: ET.Element) -> int:
    """Append child to parent if no equal child exists. Return 0-based index."""
    for i, existing in enumerate(parent):
        if _elem_equal(existing, child):
            return i
    parent.append(deepcopy(child))
    return len(parent) - 1


def _update_count(parent: ET.Element) -> None:
    parent.set("count", str(len(parent)))


def _ensure_child(parent: ET.Element, tag_local: str) -> ET.Element:
    """Find or create a direct child by local tag (within main namespace)."""
    q = _q("x", tag_local)
    for c in parent:
        if c.tag == q:
            return c
    new = ET.SubElement(parent, q)
    return new


class StyleRegistry:
    """Owns the merged styles.xml. Each source's `register(styles_root)` returns a
    dict mapping the source's cellXfs index → merged cellXfs index."""

    def __init__(self, base_styles: ET.Element):
        self.root = deepcopy(base_styles)
        self.fonts = _ensure_child(self.root, "fonts")
        self.fills = _ensure_child(self.root, "fills")
        self.borders = _ensure_child(self.root, "borders")
        self.numfmts = _ensure_child(self.root, "numFmts")
        self.cell_xfs = _ensure_child(self.root, "cellXfs")
        # numFmt by id (custom IDs ≥ 164 are user-defined; built-ins 0–49 are implicit)
        self._numfmt_by_code: dict[str, int] = {}
        for nf in self.numfmts:
            nid = nf.get("numFmtId")
            code = nf.get("formatCode", "")
            if nid is not None:
                self._numfmt_by_code[code] = int(nid)

    def _next_numfmt_id(self) -> int:
        used = set(self._numfmt_by_code.values())
        n = 164
        while n in used:
            n += 1
        return n

    def _remap_numfmt_id(self, src_root: ET.Element, src_id_str: str | None) -> str | None:
        if src_id_str is None:
            return None
        try:
            sid = int(src_id_str)
        except ValueError:
            return src_id_str
        if sid < 164:
            return src_id_str  # built-in, identical across files
        # Find the source's numFmt element matching this id
        src_numfmts = src_root.find(_q("x", "numFmts"))
        if src_numfmts is None:
            return src_id_str
        target = None
        for nf in src_numfmts:
            if nf.get("numFmtId") == src_id_str:
                target = nf
                break
        if target is None:
            return src_id_str
        code = target.get("formatCode", "")
        if code in self._numfmt_by_code:
            return str(self._numfmt_by_code[code])
        new_id = self._next_numfmt_id()
        copy_nf = deepcopy(target)
        copy_nf.set("numFmtId", str(new_id))
        self.numfmts.append(copy_nf)
        self._numfmt_by_code[code] = new_id
        return str(new_id)

    def register(self, src_styles_root: ET.Element) -> dict[int, int]:
        """Merge another styles.xml into this registry. Return cellXfs index map."""
        if src_styles_root is self.root:
            # Base source: identity map
            return {i: i for i in range(len(self.cell_xfs))}

        src_fonts = src_styles_root.find(_q("x", "fonts"))
        src_fills = src_styles_root.find(_q("x", "fills"))
        src_borders = src_styles_root.find(_q("x", "borders"))
        src_cell_xfs = src_styles_root.find(_q("x", "cellXfs"))

        font_map: dict[int, int] = {}
        if src_fonts is not None:
            for i, e in enumerate(src_fonts):
                font_map[i] = _find_or_append(self.fonts, e)
        fill_map: dict[int, int] = {}
        if src_fills is not None:
            for i, e in enumerate(src_fills):
                fill_map[i] = _find_or_append(self.fills, e)
        border_map: dict[int, int] = {}
        if src_borders is not None:
            for i, e in enumerate(src_borders):
                border_map[i] = _find_or_append(self.borders, e)

        xf_map: dict[int, int] = {}
        if src_cell_xfs is not None:
            for i, xf in enumerate(src_cell_xfs):
                new_xf = deepcopy(xf)
                # Remap child refs
                if (v := xf.get("fontId")) is not None:
                    new_xf.set("fontId", str(font_map.get(int(v), int(v))))
                if (v := xf.get("fillId")) is not None:
                    new_xf.set("fillId", str(fill_map.get(int(v), int(v))))
                if (v := xf.get("borderId")) is not None:
                    new_xf.set("borderId", str(border_map.get(int(v), int(v))))
                if (v := xf.get("numFmtId")) is not None:
                    remapped = self._remap_numfmt_id(src_styles_root, v)
                    if remapped is not None:
                        new_xf.set("numFmtId", remapped)
                xf_map[i] = _find_or_append(self.cell_xfs, new_xf)

        return xf_map

    def finalize(self) -> ET.Element:
        for parent in (self.fonts, self.fills, self.borders, self.numfmts, self.cell_xfs):
            _update_count(parent)
        return self.root


class SharedStringRegistry:
    """Owns the merged sharedStrings.xml. Returns per-source string-index maps."""

    def __init__(self, base: ET.Element | None = None):
        if base is not None:
            self.root = deepcopy(base)
        else:
            self.root = ET.Element(_q("x", "sst"))
        self._existing: dict[str, int] = {}
        for i, si in enumerate(list(self.root)):
            key = ET.tostring(si, encoding="unicode")
            self._existing[key] = i

    def register(self, src_root: ET.Element | None) -> dict[int, int]:
        if src_root is None:
            return {}
        m: dict[int, int] = {}
        for i, si in enumerate(list(src_root)):
            key = ET.tostring(si, encoding="unicode")
            if key in self._existing:
                m[i] = self._existing[key]
            else:
                self.root.append(deepcopy(si))
                idx = len(self.root) - 1
                self._existing[key] = idx
                m[i] = idx
        return m

    def finalize(self) -> ET.Element:
        self.root.set("count", str(len(self.root)))
        self.root.set("uniqueCount", str(len(self.root)))
        return self.root


def rewrite_sheet_indices(
    sheet_root: ET.Element,
    style_map: dict[int, int],
    string_map: dict[int, int],
) -> ET.Element:
    """Return a copy of sheet_root with cell s= and t='s' v= values remapped."""
    new_root = deepcopy(sheet_root)
    for c in new_root.iter(_q("x", "c")):
        s = c.get("s")
        if s is not None:
            try:
                new_s = style_map.get(int(s))
                if new_s is not None:
                    c.set("s", str(new_s))
            except ValueError:
                pass
        if c.get("t") == "s":
            v = c.find(_q("x", "v"))
            if v is not None and v.text:
                try:
                    new_v = string_map.get(int(v.text))
                    if new_v is not None:
                        v.text = str(new_v)
                except ValueError:
                    pass
    return new_root


# ---------------------------------------------------------------------------
# Workbook assembly
# ---------------------------------------------------------------------------


def _read_workbook_sheets(parts: XlsxParts) -> list[tuple[str, str]]:
    """Return [(sheet_name, sheet_xml_path), ...] for the workbook."""
    wb_root = parts.get_xml("xl/workbook.xml")
    if wb_root is None:
        return []
    sheets_el = wb_root.find(_q("x", "sheets"))
    if sheets_el is None:
        return []
    rels_root = parts.get_xml("xl/_rels/workbook.xml.rels")
    rid_to_target: dict[str, str] = {}
    if rels_root is not None:
        for rel in rels_root:
            rid = rel.get("Id")
            target = rel.get("Target", "")
            if rid:
                rid_to_target[rid] = target
    result = []
    for sh in sheets_el:
        name = sh.get("name") or ""
        rid = sh.get(_q("r", "id")) or ""
        target = rid_to_target.get(rid, "")
        # Normalize: target is relative to xl/ (e.g., "worksheets/sheet1.xml")
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = "xl/" + target
        result.append((name, path))
    return result


def _unique_name(base: str, used: set[str]) -> str:
    name = base.strip() or "Sheet"
    for ch in "[]:*?/\\'":
        name = name.replace(ch, "_")
    name = name[:31]
    final = name
    n = 2
    while final in used:
        suffix = f"_{n}"
        final = name[: 31 - len(suffix)] + suffix
        n += 1
    used.add(final)
    return final


def merge_workbooks(
    sources: list[tuple[str, bytes]],
    summary_sheets: list[tuple[str, ET.Element]] | None = None,
) -> tuple[bytes, dict[tuple[str, str], str]]:
    """Merge multiple xlsx files into one.

    sources: list of (label, xlsx_bytes). The label is used to disambiguate sheet
      names ("planA_예산"). The first source becomes the base for workbook chrome
      (styles/sharedStrings template, content types, rels).
    summary_sheets: optional list of (name, sheet_xml_root) to inject as additional
      sheets ordered BEFORE all source sheets.

    Returns: (merged_xlsx_bytes, sheet_name_map) where sheet_name_map maps
      (source_label, original_sheet_name) → final_name_in_merged_workbook.
    """
    if not sources:
        raise ValueError("no source xlsx files")

    base_label, base_bytes = sources[0]
    out = XlsxParts.from_bytes(base_bytes)

    base_styles = out.get_xml("xl/styles.xml")
    if base_styles is None:
        # Minimal styles
        base_styles = ET.Element(_q("x", "styleSheet"))
    style_reg = StyleRegistry(base_styles)
    sst_reg = SharedStringRegistry(out.get_xml("xl/sharedStrings.xml"))

    used_sheet_names: set[str] = set()

    # Track per-source: list of (final_name, rewritten_sheet_xml)
    collected_sheets: list[tuple[str, ET.Element]] = []
    name_map: dict[tuple[str, str], str] = {}

    for label, src_bytes in sources:
        src = XlsxParts.from_bytes(src_bytes)
        src_styles = src.get_xml("xl/styles.xml")
        if src_styles is None:
            src_styles = ET.Element(_q("x", "styleSheet"))
        style_map = style_reg.register(src_styles)
        src_sst = src.get_xml("xl/sharedStrings.xml")
        sst_map = sst_reg.register(src_sst)

        for sheet_name, sheet_path in _read_workbook_sheets(src):
            sheet_root = src.get_xml(sheet_path)
            if sheet_root is None:
                continue
            rewritten = rewrite_sheet_indices(sheet_root, style_map, sst_map)
            stem = label.rsplit(".", 1)[0]
            final_name = _unique_name(f"{stem}_{sheet_name}", used_sheet_names)
            collected_sheets.append((final_name, rewritten))
            name_map[(label, sheet_name)] = final_name

    # Prepend summary sheets if provided
    if summary_sheets:
        injected = []
        for name, root in summary_sheets:
            final = _unique_name(name, used_sheet_names)
            injected.append((final, root))
        collected_sheets = injected + collected_sheets

    # Now rebuild the output workbook
    # 1) Drop any base-source sheet artifacts that we replaced
    base_sheets = _read_workbook_sheets(out)
    for _name, path in base_sheets:
        if path in out.parts:
            del out.parts[path]
    # Remove per-sheet rels too
    for key in list(out.parts):
        if key.startswith("xl/worksheets/_rels/"):
            del out.parts[key]

    # 2) Write each collected sheet
    sheet_entries: list[tuple[str, str, str]] = []  # (rid, name, path)
    for i, (name, root) in enumerate(collected_sheets, 1):
        rid = f"rId{i}"
        path = f"xl/worksheets/sheet{i}.xml"
        out.set_xml(path, root)
        sheet_entries.append((rid, name, path))

    # 3) Rebuild xl/workbook.xml
    wb_root = out.get_xml("xl/workbook.xml")
    if wb_root is None:
        wb_root = ET.Element(_q("x", "workbook"))
    sheets_el = wb_root.find(_q("x", "sheets"))
    if sheets_el is None:
        sheets_el = ET.SubElement(wb_root, _q("x", "sheets"))
    for child in list(sheets_el):
        sheets_el.remove(child)
    for sheet_id, (rid, name, _path) in enumerate(sheet_entries, 1):
        sh = ET.SubElement(sheets_el, _q("x", "sheet"))
        sh.set("name", name)
        sh.set("sheetId", str(sheet_id))
        sh.set(_q("r", "id"), rid)
    out.set_xml("xl/workbook.xml", wb_root)

    # 4) Rebuild xl/_rels/workbook.xml.rels
    rels_root = ET.Element(_q("rel", "Relationships"))
    next_rid = len(sheet_entries) + 1
    for rid, name, path in sheet_entries:
        rel = ET.SubElement(rels_root, _q("rel", "Relationship"))
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        # workbook.xml lives under xl/, so target is relative
        rel.set("Target", path[len("xl/"):])
    if out.has("xl/styles.xml"):
        rel = ET.SubElement(rels_root, _q("rel", "Relationship"))
        rel.set("Id", f"rId{next_rid}"); next_rid += 1
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
        rel.set("Target", "styles.xml")
    # sharedStrings: ensure present (write the merged one)
    sst_root = sst_reg.finalize()
    if len(sst_root):
        out.set_xml("xl/sharedStrings.xml", sst_root)
        rel = ET.SubElement(rels_root, _q("rel", "Relationship"))
        rel.set("Id", f"rId{next_rid}"); next_rid += 1
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
        rel.set("Target", "sharedStrings.xml")
    elif out.has("xl/sharedStrings.xml"):
        del out.parts["xl/sharedStrings.xml"]
    if out.has("xl/theme/theme1.xml"):
        rel = ET.SubElement(rels_root, _q("rel", "Relationship"))
        rel.set("Id", f"rId{next_rid}"); next_rid += 1
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        rel.set("Target", "theme/theme1.xml")
    ET.register_namespace("", PKG_REL_NS)
    out.parts["xl/_rels/workbook.xml.rels"] = _xml_bytes(rels_root)

    # 5) Update styles
    out.set_xml("xl/styles.xml", style_reg.finalize())

    # 6) Update [Content_Types].xml — keep declared types but ensure all sheets are listed
    ct_root = out.get_xml("[Content_Types].xml")
    if ct_root is not None:
        # Drop existing per-sheet Overrides
        for child in list(ct_root):
            part_name = child.get("PartName", "")
            if part_name.startswith("/xl/worksheets/"):
                ct_root.remove(child)
        for _rid, _name, path in sheet_entries:
            ov = ET.SubElement(ct_root, _q("ct", "Override"))
            ov.set("PartName", "/" + path)
            ov.set("ContentType",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        # Ensure sharedStrings override
        has_sst_override = any(
            c.get("PartName") == "/xl/sharedStrings.xml" for c in ct_root
        )
        if out.has("xl/sharedStrings.xml") and not has_sst_override:
            ov = ET.SubElement(ct_root, _q("ct", "Override"))
            ov.set("PartName", "/xl/sharedStrings.xml")
            ov.set("ContentType",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
        ET.register_namespace("", CT_NS)
        out.parts["[Content_Types].xml"] = _xml_bytes(ct_root)

    # 7) Drop calcChain (stale after our edits) and per-sheet rels
    for stale in ("xl/calcChain.xml",):
        if stale in out.parts:
            del out.parts[stale]

    # Restore main namespace as default for any subsequent serialization in caller
    ET.register_namespace("", MAIN_NS)

    return out.to_bytes(), name_map
