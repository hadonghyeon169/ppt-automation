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
        self.LANGUAGES = {
            "th-TH": {"label": "태국어", "font": "Noto Sans TI", "complex_script": True, "country_code": "TH"},
            "ru-RU": {"label": "러시아어", "font": "Noto Sans", "complex_script": False, "country_code": "RU"},
            "vi-VN": {"label": "베트남어", "font": "Noto Sans", "complex_script": False, "country_code": "VI"},
            "id-ID": {"label": "인도네시아어", "font": "Noto Sans", "complex_script": False, "country_code": "ID"},
        }

    def ensure_dirs(self):
        for d in (self.DATA_DIR, self.UPLOAD_DIR, self.PROJECTS_DIR):
            os.makedirs(d, exist_ok=True)
