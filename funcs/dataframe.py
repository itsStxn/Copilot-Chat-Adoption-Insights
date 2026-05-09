from .str_actions import str_normalize
from .imports import dt, td, pytz
from .logs import log
from vars.exports import (
	EMPTY_CELLS_REGEX,
	MAIL_STRUCTURES,
	SHAREPOINT_DOM,
	DATE_FORMAT,
	POWERBI_DOM,
	COLTYPES,
	ACTUAL,
	PAGES,
	COLS,
	TEST,
	pd,
)
from .page import (
	is_locator_scrolled_to_bottom,
	sharepoint_close_editor,
	row_counter_info,
	powerbi_headers,
	sharepoint_rows,
	get_new_rows,
	observe_rows,
	powerbi_row,
	wait_for,
)


_BRUSSELS_TZ = pytz.timezone("Europe/Brussels")
_ROW_MISMATCH_RETRIES = 5


#* ---------------------------------------------------------------------------
#* Exceptions
#* ---------------------------------------------------------------------------

class SharePointFileNotFoundError(KeyError):
	"""Raised when a requested SharePoint file descriptor is not in SHAREPOINT_DOM."""

	def __init__(self, desc: str) -> None:
		super().__init__(f"SharePoint file not found: {desc!r}")


class SharePointReadError(RuntimeError):
	"""Raised when SharePoint row data cannot be read after exhausting retries."""


class SharePointEmptyDataError(ValueError):
	"""Raised when the SharePoint editor yields no parseable data."""


class TAMMismatchError(ValueError):
	"""Raised when the accounts and TAM counts for an AE ID do not match."""

	def __init__(self, ae_id: str, n_accounts: int, n_tams: int) -> None:
		super().__init__(
			f"Accounts/TAM length mismatch for AE ID {ae_id!r}: "
			f"{n_accounts} account(s), {n_tams} TAM(s)."
		)


class RowAttributeError(ValueError):
	"""Raised when an expected DOM attribute is missing on a Power BI row element."""

	def __init__(self, attribute: str) -> None:
		super().__init__(f"Attribute {attribute!r} not found on Power BI row element.")


#* ---------------------------------------------------------------------------
#* Date helpers
#* ---------------------------------------------------------------------------

def _now_brussels() -> dt:
	"""Return the current datetime in the Europe/Brussels timezone."""
	return dt.now(_BRUSSELS_TZ)


def _localize_to_brussels(timestamp: dt) -> dt:
	"""Attach or convert a datetime to the Europe/Brussels timezone."""
	if timestamp.tzinfo is None:
		return _BRUSSELS_TZ.localize(timestamp)
	return timestamp.astimezone(_BRUSSELS_TZ)


def is_week_old(data: pd.DataFrame) -> bool:
	"""Return True if the most recent date in the DataFrame is at least one week before yesterday."""
	dates = pd.to_datetime(data[COLS["date"]])
	last_date = _localize_to_brussels(dates.iloc[-1])
	yesterday = _now_brussels() - td(days=1)
	return (yesterday - last_date) >= td(weeks=1)


#* ---------------------------------------------------------------------------
#* DataFrame cleaning
#* ---------------------------------------------------------------------------

def empty_cells_to_num(df: pd.DataFrame) -> pd.DataFrame:
	"""Replace empty-cell matches in object columns with zero."""
	for col in df.select_dtypes(include="object").columns:
		df[col] = df[col].replace(EMPTY_CELLS_REGEX, 0, regex=True)
	return df


def drop_empty_cells(df: pd.DataFrame) -> pd.DataFrame:
	"""Drop rows that contain NaN values or empty-cell matches in text columns."""
	no_nulls = df.dropna()
	has_empty = no_nulls[COLTYPES["text"]].apply(
		lambda row: row.astype(str).str.match(EMPTY_CELLS_REGEX).any(), axis=1
	)
	return no_nulls.loc[~has_empty]


def adjust_powerbi_excel_data(df: pd.DataFrame) -> pd.DataFrame:
	"""
	Clean and normalise a raw Power BI Excel export.

	- Strips commas from non-account string values.
	- Sorts rows by account name (case-insensitive).
	- Appends a Date column set to yesterday (Brussels time).
	- Drops empty rows and converts empty string cells to zero.
	"""
	top_parent = df[[COLS["account"]]]

	#* Remove commas from numeric-like string columns so they can be cast later;
	#* the account column is excluded because it may legitimately contain commas
	body = df.drop(columns=COLS["account"]).map(
		lambda x: x.replace(",", "") if isinstance(x, str) else x
	)

	df = pd.concat([top_parent, body], axis=1)
	df = df.sort_values(by=COLS["account"], key=lambda col: col.str.lower())

	#* Tag every row with yesterday's date — this export always represents
	#* the previous day's snapshot in the Brussels timezone
	yesterday = _now_brussels() - td(days=1)
	df[COLS["date"]] = yesterday.strftime(DATE_FORMAT)

	return empty_cells_to_num(drop_empty_cells(df))


#* ---------------------------------------------------------------------------
#* Power BI extraction
#* ---------------------------------------------------------------------------

def powerbi_excel_data() -> pd.DataFrame:
	"""
	Scrape and return all Power BI accounts data as a cleaned DataFrame.

	Scrolls through the Power BI table incrementally, deduplicating rows,
	until the bottom is reached.

	Raises:
		RowAttributeError: If a row element is missing its index attribute.
	"""
	log("Exporting PowerBI accounts data...", left_nl=1)

	powerbi = PAGES["powerbi"]
	powerbi.bring_to_front()

	#* The snapshot container holds the table; top and mid are the header
	#* and scrollable body sub-panels respectively
	table = wait_for(powerbi, POWERBI_DOM, at=["snapshot"])[0]
	table.scroll_into_view_if_needed()

	top = wait_for(table, POWERBI_DOM, at=["top"])[0]
	mid = wait_for(table, POWERBI_DOM, at=["mid"])[0]

	data: list[list[str]] = []
	seen_rows: set[str] = set()
	row_num = 0

	row_attr,   _ = POWERBI_DOM["row"]
	index_attr, _ = POWERBI_DOM["index"]

	while True:
		#* Block until the expected row index is present in the DOM,
		#* ensuring the virtual list has rendered before we read it
		wait_for(mid, {"next_row": (f".row[row-index='{row_num}']", None)})

		for row in (field.strip() for field in powerbi_row(mid)):
			if row not in seen_rows:
					data.append(row.split("\n"))
					seen_rows.add(row)

		if is_locator_scrolled_to_bottom(mid):
			break

		#* Advance by a third of the viewport so the virtual list renders
		#* the next batch without skipping any rows at the boundary
		mid.evaluate("el => el.scrollBy(0, el.clientHeight / 3)")

		last_row = mid.locator(row_attr).last
		raw_attr = last_row.get_attribute(index_attr)

		if raw_attr is None:
			raise RowAttributeError(index_attr)

		#* Track the index of the last visible row so the next wait_for
		#* can block on the correct row appearing after the scroll
		row_num = int(raw_attr)
		powerbi.wait_for_timeout(500)

	df = pd.DataFrame(data, columns=powerbi_headers(top))
	log("PowerBI accounts data exported.")
	return adjust_powerbi_excel_data(df)


#* ---------------------------------------------------------------------------
#* SharePoint helpers
#* ---------------------------------------------------------------------------

def read_sharepoint_txt_data() -> pd.DataFrame:
	"""
	Parse tabular data from the SharePoint text editor and return it as a DataFrame.

	Scrolls through the editor, accumulating semicolon-delimited rows. Retries
	up to _ROW_MISMATCH_RETRIES times when row and counter counts diverge.

	Raises:
		SharePointReadError: If row counts remain mismatched after all retries.
		SharePointEmptyDataError: If no data is found after reading completes.
	"""
	sharepoint = PAGES["sharepoint"]

	#* Wait for both the row counter and the editor to be present in the DOM
	row_counter, sheet = wait_for(sharepoint, SHAREPOINT_DOM, at=["row count", "editor"])

	#* Ensure the editor has rendered at least one semicolon-delimited row before proceeding
	sheet.filter(has_text=";").wait_for(state="visible")
	row_counter.scroll_into_view_if_needed()
	sheet.scroll_into_view_if_needed()

	#* Begin observing DOM mutations on both the sheet and row counter
	#* so get_new_rows() can diff what has changed between scroll steps
	attr, _ = SHAREPOINT_DOM["row"]
	observe_rows(sheet, "rows", sharepoint_rows(sheet), attr)
	observe_rows(row_counter, "counter", row_counter_info(row_counter))

	data: list[list[str]] = []

	#* prev_row buffers the last raw text fragment across scroll boundaries,
	#* since a logical row can be split across two scroll windows
	prev_row = ""

	log("Reading SharePoint txt data...")

	while True:
		new_rows, counter = _read_rows_with_retry(sharepoint)

		#* No new rows means we have reached the end of the document;
		#* flush the buffered fragment and exit
		if not new_rows:
			data.append(prev_row.split(";"))
			break

		prev_row = _accumulate_rows(new_rows, counter, data, prev_row)

		#* Scroll the sheet and counter forward, then wait for the DOM to settle
		sheet.evaluate("el => el.scrollBy(0, el.clientHeight * 0.85)")
		row_counter.wait_for(state="visible")
		sharepoint.wait_for_timeout(500)

	if not prev_row.strip():
		raise SharePointEmptyDataError("No data found in the SharePoint editor.")

	#* Row 0 is the header; remaining rows are data
	return pd.DataFrame(data[1:], columns=data[0])


def _read_rows_with_retry(sharepoint) -> tuple[list, list]:
	"""
	Fetch new rows and their counters, retrying if the counts diverge.

	The SharePoint editor can transiently report mismatched row/counter lengths
	while the DOM is still settling after a scroll. This function retries until
	both sequences align or the retry budget is exhausted.

	Args:
		sharepoint: The Playwright Page for SharePoint, used for wait timeouts.

	Raises:
		SharePointReadError: If the mismatch persists after _ROW_MISMATCH_RETRIES attempts.

	Returns:
		A (new_rows, counter) tuple where both lists have equal length.
	"""
	for attempt in range(_ROW_MISMATCH_RETRIES):
		new_rows = get_new_rows("rows")
		counter  = get_new_rows("counter")

		if len(new_rows) == len(counter):
			return new_rows, counter

		if attempt < _ROW_MISMATCH_RETRIES - 1:
			log(f"Row count mismatch — retrying ({attempt + 1}/{_ROW_MISMATCH_RETRIES})...")
			sharepoint.wait_for_timeout(1000)

	raise SharePointReadError(
		f"Row/counter mismatch persisted after {_ROW_MISMATCH_RETRIES} retries."
	)


def _accumulate_rows(
	new_rows: list[str],
	counter: list,
	data: list[list[str]],
	prev_row: str,
) -> str:
	"""
	Merge newly observed row fragments into the data buffer.

	The SharePoint editor does not guarantee that each observed DOM node maps
	to exactly one logical row; long rows can be split across scroll windows.
	counter[i] is truthy when the i-th fragment starts a new logical row.

	Algorithm:
		- If prev_row is empty, the first fragment bootstraps the buffer.
		- A truthy counter value means the current fragment opens a new logical
			row, so the buffered fragment is committed to data first.
		- A falsy counter value means the fragment is a continuation of the
			current logical row and is appended directly to prev_row.

	Args:
		new_rows: Raw text fragments observed since the last scroll.
		counter:  Parallel list of row-boundary signals (truthy = new row).
		data:     Accumulator list; committed rows are appended here in-place.
		prev_row: The carry-over fragment from the previous scroll window.

	Returns:
		The updated prev_row buffer to carry into the next scroll window.
	"""
	for i, row in enumerate(new_rows):
		row = str_normalize(row)

		#* Bootstrap: nothing buffered yet, so just start the buffer
		if not prev_row:
			prev_row = row
			continue

		if counter[i]:
			#* This fragment opens a new logical row — commit the previous one
			data.append(prev_row.split(";"))
			prev_row = row
		else:
			#* This fragment continues the current logical row
			prev_row += row

	return prev_row


def sharepoint_txt_data(desc: str, close_editor: bool = True, attempts: int = 5) -> pd.DataFrame:
	"""
	Load a SharePoint text file into a DataFrame, retrying on transient failures.

	Args:
		desc:         Key identifying the SharePoint file in SHAREPOINT_DOM.
		close_editor: Whether to close the editor after reading. Defaults to True.
		attempts:     Maximum number of read attempts before re-raising. Defaults to 5.

	Raises:
		SharePointFileNotFoundError: If desc is not present in SHAREPOINT_DOM.
		SharePointReadError: If all read attempts fail.
	"""
	log(f"Loading SharePoint txt data ({desc})...", left_nl=1)

	if desc not in SHAREPOINT_DOM:
		raise SharePointFileNotFoundError(desc)

	sharepoint = PAGES["sharepoint"]
	sharepoint.bring_to_front()

	wait_for(sharepoint, SHAREPOINT_DOM, at=[desc])[0].click()

	last_error: Exception | None = None

	#* Retry the read in case the editor is still rendering after the click
	for attempt in range(attempts):
		try:
			sharepoint.wait_for_timeout(1000)
			df = read_sharepoint_txt_data()
			break
		except Exception as exc:
			last_error = exc
			if attempt < attempts - 1:
					log(f"Read failed ({desc}), attempt {attempt + 1}/{attempts} — retrying...")
	else:
		raise SharePointReadError(
			f"All {attempts} read attempts failed for {desc!r}."
		) from last_error

	if close_editor:
		sharepoint_close_editor()

	log(f"SharePoint txt data loaded ({desc}).")
	return df


#* ---------------------------------------------------------------------------
#* TAM correction
#* ---------------------------------------------------------------------------

def fix_tam(progress: pd.DataFrame) -> pd.DataFrame:
	"""
	Overwrite TAM values from ACTUAL data and recalculate adoption percentages.

	For each AE ID present in the actual TAM lookup, replaces the stored TAM
	figures per account, then recomputes adoption % as (adoption / TAM) * 100.

	Raises:
		TAMMismatchError: If the number of accounts and TAMs for an AE ID differ.
	"""
	t, a, a_pct = COLS["tam"], COLS["adoption"], COLS["adoption %"]
	actual_tam = ACTUAL["tam"]

	for ae_id in progress[COLS["ae"]].unique():
		matches = actual_tam.loc[actual_tam["ID"] == ae_id]

		if matches.empty:
			continue

		accounts = str(matches["Accounts"].iat[0]).split("|")
		tams     = str(matches["TAM"].iat[0]).split(",")

		if len(accounts) != len(tams):
			raise TAMMismatchError(ae_id, len(accounts), len(tams))

		for account, tam in zip(accounts, tams):
			mask = (progress[COLS["ae"]] == ae_id) & (progress[COLS["account"]] == account)
			progress.loc[mask, t] = tam

		#* Exclude zero-TAM rows from the percentage calculation to avoid
		#* division-by-zero; those rows keep whatever adoption % they had before
		nonzero  = progress[t] != 0
		adoption = progress.loc[nonzero, a].astype(int)
		tam_vals = progress.loc[nonzero, t].astype(int)

		progress.loc[nonzero, a_pct] = (adoption / tam_vals * 100).map(lambda x: f"{x:.2f}%")

	return progress


#* ---------------------------------------------------------------------------
#* Snapshot update
#* ---------------------------------------------------------------------------

def try_update_accounts_data() -> pd.DataFrame:
	"""
	Return the latest accounts snapshot, updating SharePoint first if data is stale.

	Fetches the stored snapshot; if it is not yet a week old (or a test run is
	active) returns it unchanged. Otherwise, pulls fresh Power BI data, appends
	it to the SharePoint file, saves, and returns the combined DataFrame.
	"""
	sharepoint  = PAGES["sharepoint"]
	stored_data = sharepoint_txt_data("snapshots", close_editor=False)

	if TEST["active"] or not is_week_old(stored_data):
		motive = "running test" if TEST["active"] else "data not yet a week old"
		log(f"Returning stored data ({motive}).")
		sharepoint_close_editor()
		return stored_data

	log("Updating SharePoint Copilot MAU snapshots...", left_nl=1)
	new_data = fix_tam(powerbi_excel_data())

	sharepoint.bring_to_front()

	#* Position the cursor at the last existing row so the new CSV is
	#* appended rather than inserted in the middle of the file
	last_line = wait_for(sharepoint, SHAREPOINT_DOM, at=["rows"], index={"rows": -1})[0]
	last_line.scroll_into_view_if_needed()
	last_line.click()

	log("Appending new snapshot rows...")
	csv = new_data.to_csv(index=False, header=False, sep=";").strip()
	sharepoint.keyboard.press("Enter")
	sharepoint.keyboard.insert_text(csv)
	sharepoint.wait_for_timeout(500)

	#* Click Save and wait for the confirmation popup before closing,
	#* ensuring the write is committed before we navigate away
	wait_for(sharepoint, SHAREPOINT_DOM, at=["save"])[0].click()
	wait_for(sharepoint, SHAREPOINT_DOM, at=["popup"])[0]
	sharepoint_close_editor()

	log("Returning updated data.")
	return pd.concat([stored_data, new_data], ignore_index=True)


#* ---------------------------------------------------------------------------
#* Mail structure helpers
#* ---------------------------------------------------------------------------

def set_structure_variables(role: str, names: str, url: str = "") -> pd.DataFrame:
	"""
	Build a mail structure DataFrame for the given role with name and URL substitutions.

	Args:
		role:  Key into MAIL_STRUCTURES selecting the template.
		names: Comma-separated full names; only first names are substituted.
		url:   Replaces the *URL* placeholder. Defaults to an empty string.

	Raises:
		KeyError: If role is not present in MAIL_STRUCTURES.
	"""
	#* Extract only the first name from each "First Last" entry so the greeting
	#* reads naturally even when multiple recipients are addressed together
	first_names = ", ".join(name.split()[0] for name in names.split(","))

	structure = MAIL_STRUCTURES[role].copy()
	structure["Text"] = (
		structure["Text"]
		.str.replace("*NAME*", first_names, regex=False)
		.str.replace("*URL*", url, regex=False)
	)

	return structure
