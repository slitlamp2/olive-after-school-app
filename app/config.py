import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'olive-youth-after-school-2026')
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
    OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB

    # OpenAI
    OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY', '')

    # 외부 공유용 단축 URL (bit.ly 등 → 실제 ngrok 주소로 리다이렉트). 비워두면 메인에 표시 안 함.
    PUBLIC_SHORT_URL = (os.environ.get('PUBLIC_SHORT_URL') or '').strip()

    # HWP 템플릿 경로 (프로젝트 루트의 '양식' 폴더 사용)
    PROJECT_ROOT = os.path.dirname(BASE_DIR)
    FORMS_DIR = os.path.join(PROJECT_ROOT, "양식")

    DAILY_LOG_TEMPLATE = os.path.join(FORMS_DIR, "일일활동일지양식.hwp")
    PURCHASE_DOC_TEMPLATE = os.path.join(FORMS_DIR, "품위서양식.hwp")
    PLAN_TEMPLATE = os.path.join(FORMS_DIR, "계획서양식.hwp")

    # HWP MCP 모듈 경로 (현재 Windows 사용자 기준 동적 계산)
    HWP_MCP_PATH = os.environ.get(
        "HWP_MCP_PATH",
        os.path.join(os.environ.get("APPDATA", r"C:\Users\user\AppData\Roaming"), "Cursor", "mcp-servers", "hwp-mcp"),
    )
