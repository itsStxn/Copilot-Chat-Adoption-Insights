from .imports import Playwright
from .logs import log
from vars.exports import (
	EDGE_USER_DATA_DIR,
	BrowserContext,
	BROWSER,
	TEST,
	Page,
	URLS,
	os,
)


#* ---------------------------------------------------------------------------
#* Exceptions
#* ---------------------------------------------------------------------------

class EdgeProfileNotFoundError(Exception):
	"""Raised when no Edge profile with the required site access can be found."""


class SiteAccessError(Exception):
	"""Raised when a browser profile is redirected away from an expected URL."""

	def __init__(self, profile: str, site: str, redirected_to: str) -> None:
		self.profile = profile
		self.site = site
		self.redirected_to = redirected_to
		super().__init__(
				f"Profile {profile!r} was redirected away from {site!r} to {redirected_to!r}."
		)


#* ---------------------------------------------------------------------------
#* Browser actions
#* ---------------------------------------------------------------------------

def list_edge_profiles() -> list[str]:
	"""Return Edge profile folder names found in the user data directory."""
	return [
		folder
		for folder in os.listdir(EDGE_USER_DATA_DIR)
		if "profile" in folder.lower()
	]


def goto(page: Page, url: str) -> Page:
	"""Navigate to a URL and wait until the page has settled on it."""
	page.goto(url)
	page.wait_for_url(url)
	return page


def _launch_edge_context(p: Playwright, profile_path: str) -> BrowserContext:
	"""Launch a persistent Edge browser context for the given profile path."""
	return p.chromium.launch_persistent_context(
		profile_path,
		channel="msedge",
		headless=not TEST["active"],
		viewport={"width": 1920, "height": 1080},
	)


def _assert_site_access(browser: BrowserContext, profile: str, site: str) -> None:
	"""
	Open a new page and verify it lands on the expected URL for the given site.
	Closes the page and raises SiteAccessError if a redirect is detected.

	Raises:
		SiteAccessError: If the page does not land on the expected URL.
	"""
	expected_url = URLS[site]
	page = goto(browser.new_page(), expected_url)

	if expected_url not in page.url:
		page.close()
		raise SiteAccessError(profile, site, page.url)


def find_work_profile(p: Playwright) -> tuple[BrowserContext, Page, Page]:
	"""
	Find the first Edge profile with access to both Power BI and SharePoint.

	Iterates available profiles, launching each in a persistent browser context
	and verifying that navigation to both target sites succeeds without redirects.

	Args:
		p: The Playwright instance used to launch browser contexts.

	Returns:
		A tuple of (browser_context, powerbi_page, sharepoint_page).

	Raises:
		EdgeProfileNotFoundError: If no profile with the required access is found.
	"""
	for profile in list_edge_profiles():
		profile_path = os.path.join(EDGE_USER_DATA_DIR, profile)
		log(f"Checking profile: {profile}...", left_nl=1)

		browser = _launch_edge_context(p, profile_path)

		try:
				for site in ["powerbi export", "sharepoint"]:
					_assert_site_access(browser, profile, site)
		except SiteAccessError as e:
				log(str(e))
				browser.close()
				continue

		log(f"Work profile found: {profile}")
		powerbi, sharepoint = browser.pages[1], browser.pages[2]
		return browser, powerbi, sharepoint

	raise EdgeProfileNotFoundError("No Edge profile with the required site access was found.")


def open_outlook() -> Page:
	"""
	Open Outlook's Sent Items view in a new browser page.

	Returns:
		The Playwright Page object for the Outlook Sent Items view.
	"""
	log("Opening Outlook — Sent Items view...", left_nl=1)
	outlook_sent = goto(BROWSER["edge"].new_page(), URLS["outlook"])
	log("Outlook opened.")
	return outlook_sent
