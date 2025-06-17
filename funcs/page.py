from .str_actions import str_normalize, strip_edges
from .imports import Locator
from .logs import log
from vars.exports import (
    SHAREPOINT_DOM, 
    POWERBI_DOM, 
    FINDER,
    PAGES, 
    Page
)


def click_and_wait(element:Locator, page:Page, timeout=500, render_time=30000, clicks=1) -> None:
    """
    Clicks on a given element one or more times, ensuring it is attached, visible, and scrolled into view before each click, with configurable wait times.
    Args:
        element (Locator): The element to interact with.
        page (Page): The Playwright page instance.
        timeout (int, optional): Time in milliseconds to wait before each click and after having clicked n times. Defaults to 500.
        render_time (int, optional): Maximum time in milliseconds to wait for the element to be attached and visible. Defaults to 30000.
        clicks (int, optional): Number of times to click the element. Defaults to 1.
    Returns:
        None
    """

    element.wait_for(state="attached", timeout=render_time)
    element.scroll_into_view_if_needed()
    element.wait_for(state="visible", timeout=render_time)
    page.wait_for_timeout(timeout)
    for _ in range(clicks):
        page.wait_for_timeout(timeout)
        log(f'Clicking element: {element}')
        element.click()
    
    page.wait_for_timeout(timeout)

def wait_for(container:Page|Locator, dom_attr:FINDER, timeout=60000, at:list[str]|None=None, skip:list[str]=[], dynamic:dict[str, str]|None=None, index:dict[str, int]|None=None, strict=True) -> list[Locator]:
    """
    Waits for and locates one or more elements within a given container (Page or Locator) based on provided DOM attributes.
    Args:
        container (Page | Locator): The Playwright Page or Locator to search within.
        dom_attr (FINDER): A mapping of element names to tuples of (selector, filter).
        timeout (int, optional): Maximum time to wait for each element in milliseconds. Defaults to 60000.
        at (list[str], optional): List of element names to locate. If None, all in dom_attr are used.
        skip (list[str], optional): List of element names to skip. Defaults to empty list.
        dynamic (dict[str, str], optional): Mapping of element names to dynamic attribute values to substitute in selectors.
        index (dict[str, int], optional): Mapping of element names to the index of the element to select (0-based, -1 for last).
        strict (bool, optional): Whether to strictly wait for the element to be attached and visible. Defaults to True.
    Returns:
        list[Locator]: List of located Playwright Locator objects corresponding to the requested elements.
    Raises:
        Exception: If an element cannot be located after multiple attempts.
    """

    if isinstance(container, Locator):
        container.scroll_into_view_if_needed()

    entities = {name: dom_attr[name] for name in at} if at else dom_attr

    if dynamic:
        for name, dyn_attr in dynamic.items():
            if name in entities:
                attr, filter = entities[name]
                entities[name] = (attr.replace("*attr*", dyn_attr), filter)
    
    elements = []
    for name, (attr, filter) in entities.items():
        for attempt in range(5):
            if name in skip:
                continue

            element = container.locator(attr)

            try:
                if filter:
                    element = element.filter(has_text=filter)

                selected = element

                if strict:
                    i = 0 if index is None else index[name]
                    selected = element.last if i == -1 else element.nth(i)

                    selected.wait_for(state="attached", timeout=timeout)
                    selected.scroll_into_view_if_needed()
                    selected.wait_for(state="visible", timeout=timeout)

                log(f'Element "{attr}" located')
                elements.append(selected)
                break
            except TimeoutError as e:
                if attempt >= 4:
                    log(f'Element "{attr}" not located')
                    raise Exception(e)
                
                log(f'Failed to locate "{attr}" in attempt {attempt+1}...')
                if isinstance(container, Page):
                    container.wait_for_timeout(1000)

    return elements

def powerbi_url() -> str:
    """
    Retrieves the PowerBI report URL by automating UI interactions.
    This function performs the following steps:
    1. Logs the start of the PowerBI URL retrieval process.
    2. Navigates to the PowerBI page.
    3. Clicks the "share" button and waits for the UI to update.
    4. Clicks the "url" button and waits for the UI to update.
    5. Extracts the URL from the "copy url" element's 'aria-label' attribute.
    6. Closes the URL dialog.
    7. Returns the extracted PowerBI report URL as a string.
    Returns:
        str: The URL of the PowerBI report.
    """

    log("Getting PowerBI URL...", left_nl=1)

    powerbi = PAGES["powerbi"]

    click_and_wait(
        wait_for(
            powerbi, 
            POWERBI_DOM, 
            at=["share"]
        )[0], 
        powerbi
    )

    click_and_wait(
        wait_for(
            powerbi, 
            POWERBI_DOM, 
            at=["url"]
        )[0], 
        powerbi
    )

    copy_url = wait_for(
        powerbi, 
        POWERBI_DOM, 
        at=["copy url"]
    )[0]
    url = copy_url.get_attribute("aria-label")

    click_and_wait(
        wait_for(
            powerbi, 
            POWERBI_DOM, 
            at=["close url"]
        )[0], 
        powerbi
    )

    return url

def search_option(powerbi:Page, dropdown:Locator, at:str, text_filter:str, go_down=True) -> Locator:
    """
    Searches for an option in a dropdown menu by simulating keyboard navigation and returns the corresponding locator.
    Args:
        powerbi (Page): The Playwright Page object representing the Power BI page.
        dropdown (Locator): The Playwright Locator for the dropdown element.
        at (str): The key used to identify the dropdown option in the DOM mapping.
        text_filter (str): The text to search for within the dropdown options.
        go_down (bool, optional): If True, navigates downwards through the dropdown options; if False, navigates upwards. Defaults to True.
    Returns:
        Locator: The locator for the dropdown option matching the text_filter.
    Raises:
        Exception: If the option with the specified text_filter is not found in the dropdown.
    """
    
    powerbi.keyboard.press("ArrowDown")

    attr, _ = POWERBI_DOM["option outline"]
    prev_el = dropdown.locator(attr).last

    while text_filter not in str_normalize(dropdown.inner_text()):
        powerbi.keyboard.press("ArrowDown" if go_down else "ArrowUp")
        current_el = dropdown.locator(attr).last

        if current_el == prev_el:
            raise Exception(f"Option '{text_filter}' not found in dropdown")
        
        prev_el = current_el
        

    return wait_for(
        dropdown, 
        at=[at], 
        dom_attr=POWERBI_DOM, 
        dynamic={at: text_filter}
    )[0]

def sharepoint_close_editor() -> None:
    """
    Closes the editor in the SharePoint page by clicking the 'close' button.
    This function locates the 'close' button within the SharePoint page DOM and performs a click action,
    waiting for the operation to complete. It assumes that the SharePoint page object and DOM selectors
    are defined in the global context.
    Returns:
        None
    """

    sharepoint = PAGES["sharepoint"]
    click_and_wait(
        wait_for(
            sharepoint, 
            SHAREPOINT_DOM, 
            at=["close"]
        )[0], 
        sharepoint
    )

def powerbi_row(row:Locator) -> list[str]:
    """
    Extracts and splits the normalized inner text of a Power BI row element based on a separator.
    Args:
        row (Locator): The locator object representing a row in Power BI.
    Returns:
        list[str]: A list of strings obtained by splitting the processed row text using the separator "Select Row".
    """

    sep = "Select Row"
    return strip_edges(str_normalize(row.inner_text()).strip(), sep).split(sep)

def powerbi_headers(row:Locator) -> list[str]:
    """
    Extracts and returns a list of header titles from a Power BI row locator.
    Args:
        row (Locator): The locator object representing a row in Power BI.
    Returns:
        list[str]: A list of non-empty, stripped header titles extracted from the row.
    """
    
    sep = "Row Selection\n"
    header = strip_edges(str_normalize(row.inner_text()).strip(), sep).split("\n")

    return [title.strip() for title in header if title.strip()]

def is_locator_scrolled_to_bottom(container:Locator) -> bool:
    """
    Checks if the given container element has been scrolled to the bottom.
    Args:
        container (Locator): The Playwright Locator representing the scrollable container element.
    Returns:
        bool: True if the container is scrolled to the bottom, False otherwise.
    """

    result = container.evaluate('''
        el => {
            const scrollTop = el.scrollTop;
            const scrollHeight = el.scrollHeight;
            const clientHeight = el.clientHeight;
            return (scrollTop + clientHeight) >= scrollHeight;
        }
    ''')
    return result

def sharepoint_rows(row:Locator) -> list[str]:
    """
    Extracts and normalizes the inner text from a SharePoint row element, splitting it into a list of strings by line.
    Args:
        row (Locator): The locator object representing a SharePoint row.
    Returns:
        list[str]: A list of normalized strings, each representing a line from the row's inner text.
    """

    return str_normalize(row.inner_text()).split("\n")

def row_counter_info(row_counter:Locator) -> list[str]:
    """
    Extracts the inner text of the last child element for each child of the given row_counter locator.
    Args:
        row_counter (Locator): A Playwright Locator object representing a parent node whose children will be inspected.
    Returns:
        list[str]: A list of strings containing the inner text of the last child element for each child node. If a child has no children, an empty string is returned for that child.
    """

    return row_counter.evaluate_all("""
        nodes => Array.from(nodes[0].children).map(child => {
            const n = child.children.length;
            return n > 0 ? child.children[n - 1].innerText : "";
        })
    """)

def observe_rows(container:Locator, listName:str, initialRows:list[str], selector:str="div") -> None:
    """
    Observes a container element for newly added child nodes matching a selector and updates a global list in the browser context.
    Args:
        container (Locator): The Playwright Locator for the container element to observe.
        listName (str): The name of the global list (window[listName]) to store observed row texts in the browser context.
        initialRows (list[str]): The initial list of row texts to populate the global list.
        selector (str, optional): The CSS selector for the rows to observe. Defaults to "div".
    Returns:
        None
    """

    container.evaluate(
        """(target, args) => {
            const { rowSelector, initialRows, listName } = args;

            window[listName] = [...initialRows];
            const observer = new MutationObserver(mutations => {
                for (const mutation of mutations) {
                    for (const node of mutation.addedNodes) {
                        if (node.matches && node.matches(rowSelector)) {
                            window[listName].push(node.innerText);
                        }
                    }
                }
            });
            observer.observe(target, { childList: true, subtree: true });
        }""",
        {
            "listName": listName,
            "rowSelector": selector,
            "initialRows": initialRows
        }
    )

def get_new_rows(listName:str) -> list[str]:
    """
    Retrieves and removes all rows from a specified list in the SharePoint page context.
    Args:
        listName (str): The name of the list (as a string) to retrieve and clear.
    Returns:
        list[str]: A list of strings representing the rows that were present in the specified list before removal.
    Note:
        This function assumes that the SharePoint page context contains a global JavaScript array with the given list name.
    """

    return PAGES["sharepoint"].evaluate(f"window['{listName}'].splice(0, window['{listName}'].length)")
