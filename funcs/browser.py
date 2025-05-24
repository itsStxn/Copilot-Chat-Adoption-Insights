from .imports import Playwright
from .logs import log
from vars.exports import (
    EDGE_USER_DATA_DIR, 
    BrowserContext, 
    BROWSER, 
    TEST,
    Page, 
    URLS, 
    os
)


def list_edge_profiles() -> list[str]:
    """
    Lists all Microsoft Edge user profiles found in the EDGE_USER_DATA_DIR directory.
    Returns:
        list[str]: A list of folder names representing Edge profiles. Only folders containing 'profile' (case-insensitive) in their names are included.
    """

    return [folder for folder in os.listdir(EDGE_USER_DATA_DIR) if "profile" in folder.lower()]

def goto(page:Page, url:str) -> Page:
    """
    Navigates the given Playwright page to the specified URL and waits until the navigation is complete.
    Args:
        page (Page): The Playwright Page object to navigate.
        url (str): The URL to navigate to.
    Returns:
        Page: The same Page object after navigation.
    """

    page.goto(url)
    page.wait_for_url(url)

    return page

def find_work_profile(p:Playwright) -> tuple[BrowserContext, Page, Page]:
    """
    Attempts to find a valid Microsoft Edge browser profile with access to both Power BI and SharePoint.
    Iterates through available Edge profiles, launching each in a persistent browser context.
    For each profile, it checks if the profile can successfully navigate to the Power BI and SharePoint URLs.
    If both sites are accessible (i.e., the navigation does not redirect away from the expected URL), the function returns the browser context and the corresponding pages.
    Args:
        p (Playwright): The Playwright instance used to launch the browser.
    Returns:
        tuple[BrowserContext, Page, Page]: A tuple containing the browser context, the Power BI page, and the SharePoint page.
    Raises:
        Exception: If no valid work profile with the required access is found.
    """

    for profile in list_edge_profiles():
        profile_path = os.path.join(EDGE_USER_DATA_DIR, profile)
        log(f"Checking profile: {profile}...", left_nl=1)
        
        browser = p.chromium.launch_persistent_context(
            profile_path,
            channel="msedge",
            headless=False if TEST["active"] else True,
            viewport={"width": 1920, "height": 1080}
        )

        for site in ["powerbi export", "sharepoint"]:
            access = URLS[site]
            page = goto(browser.new_page(), access)

            if access not in page.url:
                log(f"{profile} requires {site} access: unexpected navigation to {page.url}")
                while len(browser.pages) > 1:
                    browser.pages[1].close()
                break

        if len(browser.pages) != 3:
            browser.close()
            continue
        
        log(f"Work profile found: {profile}")
        powerbi, sharepoint = browser.pages[1], browser.pages[2]

        return browser, powerbi, sharepoint

    raise Exception("No valid work profile found!")

def open_outlook() -> Page:
    """
    Opens the Outlook Sent Items view in a new browser page.
    Returns:
        Page: The Playwright Page object representing the newly opened Outlook page.
    """

    log("Opening Outlook - Sent items view...", left_nl=1)

    outlook_sent = BROWSER["edge"].new_page()
    goto(outlook_sent, URLS["outlook"])

    log("Outlook opened")

    return outlook_sent
