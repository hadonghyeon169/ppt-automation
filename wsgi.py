"""gunicorn 등 프로덕션 WSGI 서버용 진입점. 예: gunicorn -w 2 -b 0.0.0.0:8000 wsgi:app"""
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from app import create_app

app = create_app()
