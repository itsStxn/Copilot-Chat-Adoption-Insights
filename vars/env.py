from .imports import load_dotenv, os


load_dotenv()
DEFAULT_EMAILS = os.getenv("default_emails")

TEST_EXCLUDED_NAMES = os.getenv("test_excluded_names")
TEST_EXCLUDED = os.getenv("test_excluded")

TEST_TO = os.getenv("test_to")
TEST_CC = os.getenv("test_cc")

URL_SHAREPOINT = os.getenv("url_sharepoint")
URL_POWERBI_SHOW = os.getenv("url_powerbi_show")
URL_POWERBI_EXPORT = os.getenv("url_powerbi_export")
