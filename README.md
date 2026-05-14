# cowork-llm-lab

Streamlit 기반의 대화형 LLM 작업 환경입니다. 다양한 모델(OpenAI, DeepMind 등)을 원격 GPU 서버(RTX 5090, Spark)에서 실행하고, 업로드한 엑셀 파일을 프롬프트로 가공·통합·연산할 수 있습니다.

## 구현 목표

- **대화형 프롬프트 UI** — Streamlit 기반 채팅 인터페이스로 모델과 멀티턴 대화
- **다중 모델 지원** — OpenAI, DeepMind 등 다양한 LLM을 다운로드 및 실행
- **원격 서버 실행** — RTX 5090, Spark 등 원격 GPU 서버에서 모델 추론 수행
- **엑셀 파일 관리** — 업로드, 리스트 조회, 삭제 등 파일 관리 기능
- **엑셀 데이터 처리** — 프롬프트로 엑셀 파일을 복제·통합·연산 등 가공
- **결과 저장 및 공유** — 처리 결과를 엑셀/마크다운으로 저장 및 전송

### 활용 예시

> 동일한 표 구조를 가진 엑셀 파일 5개를 업로드한 뒤,
> "5개 파일을 하나로 통합하고, 동일 항목은 평균값으로 입력해줘"
> 와 같은 프롬프트로 일괄 처리 결과 파일을 생성.

## 아키텍처 개요

```
┌──────────────────┐       ┌──────────────────┐       ┌────────────────────┐
│  Streamlit UI    │ ───►  │  Backend / API   │ ───►  │  Remote GPU Server │
│  (대화 / 파일)    │       │  (모델 라우팅)    │       │  (RTX 5090, Spark) │
└──────────────────┘       └──────────────────┘       └────────────────────┘
        │                          │
        ▼                          ▼
   엑셀 업로드/관리            모델 호출 (OpenAI /
                                 DeepMind / 로컬)
```

## 기술 스택

- **UI**: Streamlit
- **언어**: Python 3.10+
- **모델 런타임**: OpenAI API, DeepMind(Gemini) API, 로컬 모델 실행기
- **원격 실행**: RTX 5090 / NVIDIA Spark 기반 GPU 서버
- **데이터 처리**: pandas, openpyxl

## 사전 요구사항

- Python 3.10 이상
- 원격 GPU 서버 접속 정보 (SSH 키 또는 API 엔드포인트)
- 사용할 모델 제공자의 API 키 (OpenAI, Google DeepMind 등)

## 설치 및 실행

```bash
# 저장소 클론
git clone https://github.com/YYeoeun/cowork-llm-lab.git
cd cowork-llm-lab

# 가상환경 생성
python -m venv .venv
source .venv/bin/activate

# 의존성 설치
pip install -r requirements.txt

# 환경 변수 설정 (.env)
# OPENAI_API_KEY=...
# GOOGLE_API_KEY=...
# REMOTE_SERVER_HOST=...

# Streamlit 앱 실행
streamlit run app.py
```

## 사용 방법

1. Streamlit 페이지(`http://localhost:8501`)에 접속
2. 사이드바에서 사용할 모델과 실행 서버(로컬/원격)를 선택
3. 작업에 사용할 엑셀 파일을 업로드 (다중 업로드 지원)
4. 채팅창에 프롬프트 입력 (예: 통합, 평균, 필터링 등)
5. 결과를 미리보기 후 엑셀/마크다운 파일로 다운로드

## 프로젝트 구조

```
cowork-llm-lab/
├── README.md
├── app.py                  # Streamlit 엔트리포인트
├── requirements.txt
├── src/
│   ├── ui/                 # Streamlit 페이지/컴포넌트
│   ├── models/             # 모델 클라이언트(OpenAI, DeepMind 등)
│   ├── remote/             # 원격 서버 실행 어댑터
│   ├── files/              # 엑셀 업로드/관리/처리
│   └── prompts/            # 프롬프트 템플릿/체인
└── data/                   # 업로드 파일 저장소(런타임)
```

## 로드맵

- [ ] Streamlit 채팅 UI 골격
- [ ] 엑셀 업로드/리스트/삭제 기능
- [ ] OpenAI 모델 연동
- [ ] DeepMind(Gemini) 모델 연동
- [ ] 원격 서버 실행 어댑터 (RTX 5090 / Spark)
- [ ] 엑셀 통합/연산 프롬프트 파이프라인
- [ ] 결과 다운로드 (xlsx, md)