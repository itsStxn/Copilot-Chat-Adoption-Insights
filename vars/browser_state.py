from .imports import pd, BrowserContext, Page, TypedDict


class Accounts(TypedDict):
	data: pd.DataFrame | None
	filter: list[str] | None

ACCOUNTS: Accounts = {
	"data": None,
	"filter": None,
}
TEST = {
	"active": True
}
BROWSER: dict[str, BrowserContext] = {}
PAGES: dict[str, Page] = {}
