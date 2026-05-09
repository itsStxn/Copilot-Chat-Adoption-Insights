from .str_actions import str_normalize, strip_edges
from .imports import Locator
from .logs import log
from vars.exports import (
	SHAREPOINT_DOM,
	POWERBI_DOM,
	FINDER,
	PAGES,
	Page,
)


#* ---------------------------------------------------------------------------
#* Exceptions
#* ---------------------------------------------------------------------------

class ElementNotFoundError(RuntimeError):
	"""Raised when a DOM element cannot be located after exhausting all retry attempts."""

	def __init__(self, selector: str) -> None:
		super().__init__(f"Element {selector!r} could not be located after retries.")


class PowerBIUrlError(RuntimeError):
	"""Raised when the Power BI share dialog yields an empty URL."""

	def __init__(self) -> None:
		super().__init__("Power BI share dialog returned an empty URL.")


class DropdownOptionNotFoundError(LookupError):
	"""Raised when a searched option is not found within a dropdown."""

	def __init__(self, text_filter: str) -> None:
		super().__init__(f"Option {text_filter!r} not found in the dropdown.")


#* ---------------------------------------------------------------------------
#* Click helpers
#* ---------------------------------------------------------------------------

def click_and_wait(
	element: Locator,
	page: Page,
	timeout: int = 500,
	render_time: int = 30000,
	clicks: int = 1,
) -> None:
	"""
	Ensure element is ready, then click it one or more times.

	Waits for the element to be attached and visible before the first click,
	and pauses for timeout milliseconds between every click and after the last.

	Args:
		element:     Locator to interact with.
		page:        Page the element belongs to, used for timeout waits.
		timeout:     Milliseconds to wait before each click and after the last. Defaults to 500.
		render_time: Maximum milliseconds to wait for attach/visible state. Defaults to 30000.
		clicks:      Number of times to click. Defaults to 1.
	"""
	element.wait_for(state="attached", timeout=render_time)
	element.scroll_into_view_if_needed()
	element.wait_for(state="visible", timeout=render_time)
	page.wait_for_timeout(timeout)

	for _ in range(clicks):
		page.wait_for_timeout(timeout)

		#* Build a human-readable CSS selector from the element's tag, id, and
		#* class list so the log line is meaningful without needing DevTools
		selector = page.evaluate(
			"""el =>
					el.tagName.toLowerCase()
					+ (el.id ? '#*' + el.id : '')
					+ (!el.id && el.className
						? '.' + el.className.split(' ').join('.')
						: '')
			""",
			element.element_handle(),
		)
		log(f"Clicking element: {selector}")
		element.click()

	page.wait_for_timeout(timeout)


#* ---------------------------------------------------------------------------
#* Element location
#* ---------------------------------------------------------------------------

def wait_for(
	container: Page | Locator,
	dom_attr: FINDER,
	timeout: int = 60000,
	at: list[str] | None = None,
	skip: list[str] | None = None,
	dynamic: dict[str, str] | None = None,
	index: dict[str, int] | None = None,
	strict: bool = True,
) -> list[Locator]:
	"""
	Locate one or more DOM elements within a page or locator container.

	Elements are identified by name via dom_attr, which maps each name to a
	(selector, text_filter) tuple. When strict=True each element is confirmed
	attached and visible before being returned.

	Args:
		container: Page or Locator to search within.
		dom_attr:  Mapping of element names to (selector, text_filter) tuples.
		timeout:   Max milliseconds to wait for each element. Defaults to 60000.
		at:        Names to locate; all dom_attr entries are used when None.
		skip:      Names to skip entirely. Defaults to [].
		dynamic:   Name → value substitutions applied to "*attr*" placeholders
						in selectors before locating.
		index:     Name → positional index for nth-element selection;
						0-based, -1 selects the last element. Defaults to 0 for all.
		strict:    Confirm attached + visible before returning. Defaults to True.

	Returns:
		Ordered list of Locators matching the requested names.

	Raises:
		ElementNotFoundError: If any element cannot be located after 5 attempts.
	"""
	skip = skip or []

	if isinstance(container, Locator):
		container.scroll_into_view_if_needed()

	#* Build the working set — either a subset defined by `at` or the full map
	entities: dict = {name: dom_attr[name] for name in at} if at else dict(dom_attr)

	#* Substitute dynamic placeholders in selectors before locating
	if dynamic:
		for name, dyn_value in dynamic.items():
			if name in entities:
				attr, text_filter = entities[name]
				entities[name] = (attr.replace("*attr*", dyn_value), text_filter)

	elements: list[Locator] = []

	for name, (attr, text_filter) in entities.items():
		if name in skip:
			continue

		for attempt in range(5):
			try:
				element = container.locator(attr)

				if text_filter:
					element = element.filter(has_text=text_filter)

				if strict:
					#* Resolve to a single element using the provided index (default: first)
					i = 0 if index is None else index.get(name, 0)
					selected = element.last if i == -1 else element.nth(i)

					selected.wait_for(state="attached", timeout=timeout)
					selected.scroll_into_view_if_needed()
					selected.wait_for(state="visible", timeout=timeout)
				else:
					selected = element

				log(f'Element "{attr}" located.')
				elements.append(selected)
				break

			except TimeoutError:
				if attempt == 4:
					raise ElementNotFoundError(attr)

				log(f'Failed to locate "{attr}" — attempt {attempt + 1}/5, retrying...')
				if isinstance(container, Page):
					container.wait_for_timeout(1000)

	return elements


#* ---------------------------------------------------------------------------
#* Power BI page interactions
#* ---------------------------------------------------------------------------

def powerbi_url() -> str:
	"""
	Extract the current report URL from Power BI's share dialog.

	Opens the share panel, navigates to the URL copy field, reads the
	aria-label attribute (which Power BI populates with the full URL),
	then closes the dialog.

	Returns:
		The report URL as a string.

	Raises:
		PowerBIUrlError: If the aria-label attribute is empty or missing.
	"""
	log("Getting Power BI URL...", left_nl=1)

	powerbi = PAGES["powerbi"]

	#* The URL is surfaced through two nested dialogs: share → url copy field
	click_and_wait(wait_for(powerbi, POWERBI_DOM, at=["share"])[0], powerbi)
	click_and_wait(wait_for(powerbi, POWERBI_DOM, at=["url"])[0],   powerbi)

	copy_url = wait_for(powerbi, POWERBI_DOM, at=["copy url"])[0]
	url = copy_url.get_attribute("aria-label")

	#* Close the dialog regardless of whether the URL was found
	click_and_wait(wait_for(powerbi, POWERBI_DOM, at=["close url"])[0], powerbi)

	if not url:
		raise PowerBIUrlError()

	return url


def search_option(
	powerbi: Page,
	dropdown: Locator,
	at: str,
	text_filter: str,
	go_down: bool = True,
) -> Locator:
	"""
	Keyboard-navigate a Power BI dropdown until the target option is visible.

	Presses ArrowDown (or ArrowUp when go_down=False) repeatedly, comparing
	the last focused element after each press to detect when the list wraps
	without finding the target.

	Args:
		powerbi:     Page driving the keyboard events.
		dropdown:    Locator for the open dropdown container.
		at:          dom_attr key used to locate the matched option element.
		text_filter: Text the target option must contain.
		go_down:     Navigate downward through the list. Defaults to True.

	Returns:
		Locator for the matching option element.

	Raises:
		DropdownOptionNotFoundError: If the list wraps without a match.
	"""
	direction = "ArrowDown" if go_down else "ArrowUp"

	#* Prime the dropdown so it registers keyboard focus
	powerbi.keyboard.press("ArrowDown")

	outline_attr, _ = POWERBI_DOM["option outline"]
	prev_el = dropdown.locator(outline_attr).last

	while text_filter not in str_normalize(dropdown.inner_text()):
		powerbi.keyboard.press(direction)
		current_el = dropdown.locator(outline_attr).last

		#* If the focused element did not change the list has wrapped — give up
		if current_el == prev_el:
			raise DropdownOptionNotFoundError(text_filter)

		prev_el = current_el

	return wait_for(
		dropdown,
		at=[at],
		dom_attr=POWERBI_DOM,
		dynamic={at: text_filter},
	)[0]


def powerbi_row(row: Locator) -> list[str]:
	"""
	Split a Power BI data row's text into individual cell values.

	Power BI renders each row with "Select Row" as a separator between cells;
	strip_edges removes any leading/trailing occurrence before splitting.

	Args:
		row: Locator for the row element.

	Returns:
		List of cell text strings.
	"""
	sep = "Select Row"
	return strip_edges(str_normalize(row.inner_text()).strip(), sep).split(sep)


def powerbi_headers(row: Locator) -> list[str]:
	"""
	Extract column header titles from a Power BI header row element.

	Headers are newline-separated after normalisation; "Row Selection" is
	stripped from the edges since Power BI prepends it to the header text.

	Args:
		row: Locator for the header row element.

	Returns:
		Non-empty header title strings in column order.
	"""
	sep = "Row Selection\n"
	header = strip_edges(str_normalize(row.inner_text()).strip(), sep).split("\n")
	return [title.strip() for title in header if title.strip()]


def is_locator_scrolled_to_bottom(container: Locator) -> bool:
	"""
	Return True if the container's scroll position has reached the bottom.

	Uses scrollTop + clientHeight >= scrollHeight, which matches the browser's
	own definition of "scrolled to bottom" (accounting for sub-pixel rounding).

	Args:
		container: Locator for the scrollable element.
	"""
	return container.evaluate("""el => {
		return (el.scrollTop + el.clientHeight) >= el.scrollHeight;
	}""")


#* ---------------------------------------------------------------------------
#* SharePoint page interactions
#* ---------------------------------------------------------------------------

def sharepoint_close_editor() -> None:
	"""Click the SharePoint editor's close button and wait for the UI to settle."""
	sharepoint = PAGES["sharepoint"]
	click_and_wait(wait_for(sharepoint, SHAREPOINT_DOM, at=["close"])[0], sharepoint)


def sharepoint_rows(row: Locator) -> list[str]:
	"""
	Split a SharePoint row element's normalised text into individual lines.

	Args:
		row: Locator for the row element.

	Returns:
		List of line strings from the row's inner text.
	"""
	return str_normalize(row.inner_text()).split("\n")


def row_counter_info(row_counter: Locator) -> list[str]:
	"""
	Extract the last child's text from each direct child of the row counter element.

	The SharePoint row counter renders each row's number inside nested elements;
	reading the last child gives the most specific (innermost) value.

	Args:
		row_counter: Locator for the row counter container.

	Returns:
		List of text strings, one per direct child of the container.
		Empty string for children that have no sub-children.
	"""
	return row_counter.evaluate_all("""nodes =>
		Array.from(nodes[0].children).map(child => {
			const n = child.children.length;
			return n > 0 ? child.children[n - 1].innerText : "";
		})
	""")


def observe_rows(
	container: Locator,
	list_name: str,
	initial_rows: list[str],
	selector: str = "div",
) -> None:
	"""
	Attach a MutationObserver that appends newly added matching nodes to a browser-global list.

	The list is initialised with initial_rows so that rows already present in
	the DOM at observation time are not missed by get_new_rows().

	Args:
		container:    Locator for the element to observe.
		list_name:    Name of the window-level JS array to populate.
		initial_rows: Rows already present before observation begins.
		selector:     CSS selector used to filter added nodes. Defaults to "div".
	"""
	container.evaluate(
		"""(target, { listName, rowSelector, initialRows }) => {
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
		{"listName": list_name, "rowSelector": selector, "initialRows": initial_rows},
	)


def get_new_rows(list_name: str) -> list[str]:
	"""
	Drain and return all rows accumulated in a browser-global list since the last call.

	Uses splice(0, length) so the list is cleared atomically — rows appended
	by the MutationObserver between the read and the clear are not lost.

	Args:
		list_name: Name of the window-level JS array to drain.

	Returns:
		All strings that were in the list at the moment of the call.
	"""
	return PAGES["sharepoint"].evaluate(
		f"window['{list_name}'].splice(0, window['{list_name}'].length)"
	)
