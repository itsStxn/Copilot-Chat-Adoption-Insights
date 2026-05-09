from .imports import load_dotenv, os


load_dotenv()
DEFAULT_EMAILS = os.getenv("default_emails") or ""

TEST_EXCLUDED_NAMES = os.getenv("test_excluded_names") or ""
TEST_EXCLUDED = os.getenv("test_excluded") or ""

TEST_TO = os.getenv("test_to") or ""
TEST_CC = os.getenv("test_cc") or ""

URL_SHAREPOINT = os.getenv("url_sharepoint") or ""
URL_POWERBI_SHOW = os.getenv("url_powerbi_show") or ""
URL_POWERBI_EXPORT = os.getenv("url_powerbi_export") or ""
