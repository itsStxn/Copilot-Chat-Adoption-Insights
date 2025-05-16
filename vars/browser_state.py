from .imports import pd, BrowserContext, Page

TEST = {
  "active": True
}
ACCOUNTS: dict[str, pd.DataFrame|list[str]|None] = {
  "data": None,
  "filter": None
}
BROWSER: dict[str, BrowserContext] = {}
PAGES: dict[str, Page] = {}
