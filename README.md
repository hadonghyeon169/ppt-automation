# CultureFi PPT 자동화 시스템

한국어 교육용 PPT를 업로드하면 (1) Claude API로 목표 언어 번역 → (2) AI 자동 검수 + 사람 확인
→ (3) 타입캐스트로 음성 생성 → (4) 음성/최종 PPT AI 자동 비교 + 사람 최종 승인까지
한 화면에서 진행하는 내부용 웹 서비스입니다. `culturefi-ppt-translation` 스킬 문서의 번역
규칙(도형 유형 판단, 흰색 텍스트 보존, 볼드 규칙, 넘침 보정 등)을 코드로 이식해 자동화했습니다.

## 왜 이렇게 만들었나 (설계 요약)

- **판단은 AI, 반영은 코드**: "이 도형을 번역해야 하는가", "번역문은 무엇인가" 같은 판단은
  Claude API가 스킬 규칙을 프롬프트로 받아 수행합니다. 반면 "볼드를 켤지, 흰색을 유지할지,
  폰트를 어디에 적용할지" 같은 PPTX XML 조작은 스킬 문서의 함수를 그대로 이식한 결정론적
  코드(`app/services/ppt_xml_ops.py`)가 담당합니다. 이렇게 나눈 이유는 XML 조작 실수(예:
  흰색 글씨가 검정으로 바뀌는 버그)는 AI가 아니라 코드가 보장해야 재현 가능하기 때문입니다.
- **모든 API 키는 선택 사항으로 시작**: 서버는 키가 하나도 없어도 켜집니다. 각 단계 실행 시
  필요한 키가 없으면 화면에 안내가 뜹니다. `/settings` 페이지에서 언제든 등록/교체할 수 있습니다.
- **"AI 자동 비교 + 사람 최종 확인"**: 3단계(번역 검수)와 5단계(음성·최종 PPT 검수)는 AI가
  먼저 이상 징후를 찾아 화면에 표시하고, 사람이 확인 버튼을 눌러야 다음 단계로 진행됩니다.
  특히 5단계는 "승인된 번역 스냅샷과 현재 PPT 텍스트가 100% 일치하는지"를 결정론적으로
  비교해서, AI의 판단이 아니라 정확한 diff로 최종 확인합니다.
- **1인 사용 → 팀 사용 확장 대비**: 지금은 관리자 계정 1개로 시작하지만, `users` 테이블에
  `role` 컬럼이 이미 있어 나중에 팀원 계정을 추가하는 기능만 얹으면 됩니다 (아래 "다인원
  확장" 참고).

## 아키텍처

```
[브라우저] ──HTTP──> [Flask 앱 (app/)] ──> [SQLite (data/app.db)]
                          │
                          ├─ services/ppt_xml_ops.py    : PPTX XML 저수준 조작 (스킬 함수 이식)
                          ├─ services/slide_extractor.py : PPT → 구조화 JSON 추출
                          ├─ services/translator.py      : Claude API 호출 → 번역 계획(JSON)
                          ├─ services/pptx_pipeline.py    : 번역 계획을 실제 PPTX에 반영 + 넘침 보정
                          ├─ services/qa_reviewer.py      : GPT API 기반 번역/오디오 자동 검수
                          ├─ services/tts_typecast.py     : 타입캐스트 TTS REST 클라이언트
                          ├─ services/render.py           : LibreOffice 기반 PDF/PNG 렌더링 미리보기
                          └─ services/pipeline_runner.py  : 위 서비스들을 단계별로 오케스트레이션
                                                             (백그라운드 스레드에서 실행)
```

- **웹 프레임워크**: Flask (+ 표준 라이브러리 `sqlite3`). 외부 ORM/큐 의존성을 최소화해
  어떤 호스팅 환경에서도 바로 배포되도록 설계했습니다.
- **비동기 처리**: 번역/음성 생성처럼 오래 걸리는 작업은 `threading.Thread`로 백그라운드
  실행하고, 진행 상황은 전부 DB에 기록합니다. 화면은 2.5초 간격으로 `/projects/<id>/status.json`
  을 폴링해 진행률을 보여줍니다. (트래픽이 커지면 Celery+Redis 같은 작업 큐로 교체를
  권장하지만, 소규모 내부 도구에는 이 구조로 충분합니다.)
- **파일 저장**: 로컬 디스크(`DATA_DIR` 환경변수, 기본 `./data`)에 원본/번역본/오디오/렌더링
  미리보기를 저장합니다. 배포 플랫폼의 영구 볼륨을 이 경로에 마운트해야 재배포 시 파일이
  사라지지 않습니다 (아래 "배포" 참고). 여러 사용자가 커지고 파일이 많아지면 S3/R2 같은
  오브젝트 스토리지로 옮기는 것을 권장합니다.

## 처리 흐름 (5단계)

1. **업로드**: PPT 파일 + 목표 언어 선택. 같은 시리즈의 기존 번역 완성본이 있으면 "참고
   파일"로 함께 업로드하면 번역 프롬프트에 참고 자료로 포함됩니다 (스킬 문서의 "완성본
   대조 작업법").
2. **번역 (Claude API)**: 슬라이드를 배치로 묶어 Claude에 보내고, 스킬 규칙이 담긴
   시스템 프롬프트로 "어떤 도형을 어떻게 번역할지" 판단을 받습니다. 받은 계획을 코드가
   PPTX XML에 반영하고(폰트/볼드/색상/넘침 보정), LibreOffice로 미리보기 이미지를 만듭니다.
3. **번역 검수 (GPT 자동 비교 + 사람 확인)**: OpenAI API로 원문·번역문 쌍을 검수해 의심
   항목을 표시합니다. 오류(error)가 있으면 기본적으로 다음 단계로 못 넘어가고, "오류 있어도
   강제 진행"을 눌러야 넘어갈 수 있습니다.
4. **음성 생성 (타입캐스트)**: 슬라이드별 번역 텍스트를 대본으로 타입캐스트 TTS를 호출해
   음성 파일을 만듭니다. WAV로 생성하면 재생 길이를 정확히 측정해 5단계에서 활용합니다.
5. **최종 검수 (AI 자동 비교 + 사람 최종 확인)**: (a) 번역 승인 시점 스냅샷과 현재 PPT
   텍스트를 비교해 100% 일치하는지 확인 (b) 오디오 길이가 텍스트 분량과 크게 다르면 경고
   (c) OpenAI 키가 있으면 Whisper로 오디오를 전사해 빈 음성(생성 실패) 여부를 확인합니다.
   문제 없으면(또는 강제로) 최종 승인하면 완료 처리되고 다운로드할 수 있습니다.

## 로컬 실행

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # 값 채우기 (키는 비워둬도 됨)
python run.py           # http://localhost:8000
```

LibreOffice(`soffice`)와 `pdftoppm`(poppler-utils)이 설치되어 있어야 렌더링 미리보기가
동작합니다. 없어도 나머지 기능은 정상 동작하고, 렌더링만 건너뛰며 안내 메시지를 남깁니다.

기본 로그인은 `.env`의 `ADMIN_USERNAME`/`ADMIN_PASSWORD`로 최초 실행 시 자동 생성됩니다.

## 배포 (호스팅 사이트로 운영하기)

API 키가 아직 없다고 하셔서 우선 **구조만** 완성했습니다. 아래는 실제로 어디서든 브라우저로
접속 가능한 사이트로 올리는 방법입니다. `Dockerfile`이 포함되어 있어 Docker를 지원하는
플랫폼이면 어디든 배포할 수 있습니다. 추천 순서:

1. **Railway** 또는 **Render** (추천): GitHub 저장소만 연결하면 `Dockerfile`을 감지해
   자동 빌드합니다. "영구 볼륨(Persistent Volume/Disk)"을 `/data` 경로에 추가하고,
   환경변수(`.env.example` 참고: `FLASK_SECRET_KEY`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`,
   `DATA_DIR=/data`, API 키들)를 설정하면 끝입니다. 두 서비스 모두 월 몇 달러 수준의
   소규모 요금제로 충분합니다.
2. **Fly.io**: `fly volumes create data --size 1` 로 볼륨을 만들고 `fly deploy`.
3. **직접 서버(VPS/AWS 등)**: `docker build -t ppt-auto . && docker run -d -p 8000:8000
   -v $(pwd)/data:/data --env-file .env ppt-auto` 로 실행 후 Nginx/Caddy로 HTTPS 리버스
   프록시를 붙이면 됩니다.

배포 후에는 `/settings` 페이지에서 Claude / OpenAI / 타입캐스트 API 키를 입력하면 바로
전체 파이프라인이 동작합니다.

## API 키를 나중에 넣는 방법

두 가지 방법이 있습니다 (둘 다 지원됩니다):
- 배포 플랫폼의 환경변수로 `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`, `TYPECAST_API_KEY` 설정
- 로그인 후 `/settings` 페이지에서 직접 입력 (DB에 저장되며, 환경변수보다 우선 적용됨)

## 다인원(팀) 확장 방법

지금은 관리자 계정 1개(`ADMIN_USERNAME`/`ADMIN_PASSWORD`)로만 로그인됩니다. 나중에 동료들과
같이 쓰실 때는 다음을 추가하면 됩니다 (스키마는 이미 준비되어 있어 큰 변경 없이 가능합니다):
- `/settings` 또는 별도 `/users` 페이지에 "팀원 추가" 폼 (아이디/비밀번호 발급)
- `projects` 테이블의 `created_by` 컬럼을 활용해 "내 프로젝트만 보기" 또는 "전체 공유" 정책 결정
- 필요하면 `role` 컬럼으로 admin/member 권한 분리 (예: API 키 설정은 admin만)

## 알려진 한계 / 다음에 다듬으면 좋은 부분

- **번역 프롬프트 튜닝**: 실제 Claude 키로 몇 개 강의를 돌려보면서 스킬 문서의 세부 규칙
  (예: 러시아어 문법 슬라이드 폰트 크기 10pt/9pt, 정의=흰색/구조·예문=검정 색상 규칙)이
  잘 지켜지는지 확인하고 `app/services/skill_prompt.py`를 보강하는 과정이 필요합니다.
  지금은 스킬 문서의 핵심 규칙을 프롬프트로 압축해 넣었지만, 실전에서는 예외 케이스가
  계속 발견될 수 있습니다.
- **넘침(overflow) 보정 정확도**: 실제 배포 폰트가 없는 환경에서는 DejaVu Sans로 근사
  계산합니다. 완벽하지 않으므로 LibreOffice 렌더링 미리보기로 육안 확인하는 습관이
  중요합니다 (화면에 자동으로 보여줍니다).
- **슬라이드별 재번역**: 지금은 문제가 있으면 전체를 처음부터 다시 번역합니다 (스킬
  문서도 "부분 수정 금지, 처음부터 재생성"을 권장). 특정 슬라이드만 다시 번역하는 기능은
  다음 단계로 추가할 수 있습니다.
- **동시 작업 규모**: 현재는 스레드 기반으로 소규모 팀 사용에 적합합니다. 사용량이 크게
  늘면 Celery/RQ + Redis 같은 작업 큐로 교체하는 것을 권장합니다.
- **타입캐스트 텍스트 길이 제한**: 1회 호출당 2,000자 제한이 있어 긴 슬라이드는 문장
  단위로 분할 후 이어붙입니다 (`tts_typecast.split_text_for_tts`). 자연스러운 끊김을
  위해 추후 무음 구간 삽입 등을 보강할 수 있습니다.

## 폴더 구조

```
app/
  __init__.py          앱 팩토리
  config.py            환경설정 (언어별 폰트 규칙 포함)
  db.py                sqlite 스키마 + 커넥션
  repo.py              DB 접근 함수 모음
  auth.py              로그인/세션
  routers/             Flask 블루프린트 (auth, projects, settings)
  services/            핵심 로직 (위 아키텍처 참고)
  templates/           Jinja2 HTML 템플릿
data/                  업로드 원본/번역본/오디오/렌더링 미리보기 (배포 시 영구 볼륨으로 마운트)
requirements.txt
Dockerfile
run.py / wsgi.py
.env.example
```
