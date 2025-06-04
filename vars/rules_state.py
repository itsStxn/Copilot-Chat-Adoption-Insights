from .custom_types import EXCLUDER
from .imports import pd


MAIL_STRUCTURES: dict[str, pd.DataFrame] = {}
ACTUAL: dict[str, pd.DataFrame] = {}
SKIP: dict[str, pd.DataFrame] = {}
EXCLUDED: EXCLUDER = {}
