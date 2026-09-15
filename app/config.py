import os


class Config:
    def __init__(self):
        base_data_dir = os.environ.get("DATA_DIR", "./data")
        self.DATA_DIR = os.path.abspath(base_data_dir)
        self.UPLOAD_DIR = os.path.join(self.DATA_DIR, "uploads")
        self.PROJECTS_DIR = os.path.join(self.DATA_DIR, "projects")
        self.DB_PATH = os.path.join(self.DATA_DIR, "app.db")

        self.SECRET_KEY = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")
        self.ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
        self.ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin")

        self.MAX_CONTENT_LENGTH = 200 * 1024 * 1024  # 200MB, 대용량 PPT 대비

        # 지원 언어 → (표시명, 폰트, lang 코드, Complex Script 여부, 파일명용 국적 코드)
        # "PPT 강의 제작 인수인계서" 1-6 폰트 규칙 표 기준 (2026-09 갱신).
        # 태국어=Noto Sans TI, 러시아어/베트남어/영어=Noto Sans, 한국어=Noto Sans KR.
        # 인도네시아어는 회사 언어 목록엔 없지만 베트남어와 동일하게 라틴 문자권이라
        # 동일 폰트(Noto Sans)를 적용한다.
        # tts_lang: 타입캐스트 API가 요구하는 ISO 639-3 언어 코드 (auto-detect에 맡기지 않고
        # 명시적으로 지정해야 정확한 언어로 읽는다 — 특히 이 앱은 "받침" 같은 한글 단어를
        # 번역문 안에 그대로 섞어 넣는 규칙이 있어서, 언어를 지정하지 않으면 자동감지가
        # 혼동해 엉뚱한 언어(중국어 등)로 읽는 경우가 있었다).
        self.LANGUAGES = {
            "th-TH": {"label": "태국어", "font": "Noto Sans TI", "complex_script": True, "country_code": "TH", "tts_lang": "tha"},
            "ru-RU": {"label": "러시아어", "font": "Noto Sans", "complex_script": False, "country_code": "RU", "tts_lang": "rus"},
            "vi-VN": {"label": "베트남어", "font": "Noto Sans", "complex_script": False, "country_code": "VI", "tts_lang": "vie"},
            "id-ID": {"label": "인도네시아어", "font": "Noto Sans", "complex_script": False, "country_code": "ID", "tts_lang": "ind"},
        }

    def ensure_dirs(self):
        for d in (self.DATA_DIR, self.UPLOAD_DIR, self.PROJECTS_DIR):
            os.makedirs(d, exist_ok=True)
