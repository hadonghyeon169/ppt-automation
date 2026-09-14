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

        # 지원 언어 → (표시명, 폰트, lang 코드, Complex Script 여부)
        # culturefi-ppt-translation 스킬의 "폰트 규칙" 테이블을 그대로 반영
        self.LANGUAGES = {
            "th-TH": {"label": "태국어", "font": "Sarabun", "complex_script": True},
            "ru-RU": {"label": "러시아어", "font": "Noto Sans", "complex_script": False},
            "vi-VN": {"label": "베트남어", "font": "Noto Sans KR", "complex_script": False},
            "id-ID": {"label": "인도네시아어", "font": "Noto Sans KR", "complex_script": False},
        }

    def ensure_dirs(self):
        for d in (self.DATA_DIR, self.UPLOAD_DIR, self.PROJECTS_DIR):
            os.makedirs(d, exist_ok=True)
