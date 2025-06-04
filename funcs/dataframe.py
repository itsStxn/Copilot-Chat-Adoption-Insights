from .str_actions import str_normalize
from .imports import dt, td, pytz
from .logs import log
from vars.exports import (
  EMPTY_CELLS_REGEX, 
  MAIL_STRUCTURES, 
  SHAREPOINT_DOM, 
  SHAREPOINT_DOM,
  DATE_FORMAT, 
  POWERBI_DOM, 
  COLTYPES, 
  ACTUAL,
  PAGES, 
  COLS,
  TEST,
  pd
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
  wait_for
)


def is_week_old(data:pd.DataFrame) -> bool:
    """
    Checks if the most recent date in the DataFrame is at least one week older than yesterday (in the Europe/Brussels timezone).
    Args:
        data (pd.DataFrame): A pandas DataFrame containing a "Date" column with date values.
    Returns:
        bool: True if the last date in the "Date" column is at least one week older than yesterday, False otherwise.
    """

    dates = data[[COLS["date"]]]
    dates.loc[:, COLS["date"]] = pd.to_datetime(dates[COLS["date"]])
    last_date: dt = dates[COLS["date"]].iloc[-1]

    bxl_tz = pytz.timezone("Europe/Brussels")
    yesterday = dt.now(bxl_tz) - td(days=1)
    
    if last_date.tzinfo is None:
        last_date = bxl_tz.localize(last_date)
    else:
        last_date = last_date.astimezone(bxl_tz)

    return (yesterday - last_date) >= td(weeks=1)

def empty_cells_to_num(df:pd.DataFrame) -> pd.DataFrame:
    """
    Replaces empty cells in object-type columns of a DataFrame with zeros.
    This function iterates over all columns in the given DataFrame. For columns with dtype 'object',
    it replaces values matching the pattern specified by EMPTY_CELLS_REGEX with the integer 0.
    Args:
        df (pd.DataFrame): The input DataFrame to process.
    Returns:
        pd.DataFrame: The DataFrame with empty cells replaced by zeros in object-type columns.
    Note:
        - The function assumes that EMPTY_CELLS_REGEX is defined elsewhere in the code.
        - Only columns with dtype 'object' are processed.
    """

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].replace(EMPTY_CELLS_REGEX, 0, regex=True)
    return df

def drop_empty_cells(df:pd.DataFrame) -> pd.DataFrame:
    """
    Removes rows from the DataFrame that contain any empty cells or cells matching a specific empty cell regex pattern.
    Parameters:
        df (pd.DataFrame): The input DataFrame to process.
    Returns:
        pd.DataFrame: A DataFrame with rows containing empty cells or cells matching the EMPTY_CELLS_REGEX in the column specified by COLTYPES["text"] removed.
    Notes:
        - The function first drops rows with any NaN values.
        - Then, it filters out rows where any value in the column specified by COLTYPES["text"] matches the EMPTY_CELLS_REGEX pattern.
        - Requires the global variables COLTYPES and EMPTY_CELLS_REGEX to be defined.
    """

    return df.dropna().loc[~df[COLTYPES["text"]].apply(lambda row: row.astype(str).str.match(EMPTY_CELLS_REGEX).any(), axis=1)]

def adjust_powerbi_excel_data(df:pd.DataFrame) -> pd.DataFrame:
    """
    Adjusts and cleans a Power BI Excel DataFrame.
    This function performs the following operations on the input DataFrame:
    1. Separates the 'TopParent' column and temporarily removes it from the DataFrame.
    2. Removes commas from string values in all columns except 'TopParent'.
    3. Reattaches the 'TopParent' column to the DataFrame.
    4. Sorts the DataFrame by the 'TopParent' column in a case-insensitive manner.
    5. Adds a 'Date' column with the previous day's date in the 'Europe/Brussels' timezone, formatted according to DATE_FORMAT.
    6. Cleans the DataFrame by dropping empty cells and converting empty cells to numeric values.
    Args:
        df (pd.DataFrame): The input DataFrame to be adjusted and cleaned.
    Returns:
        pd.DataFrame: The cleaned and adjusted DataFrame ready for Power BI Excel processing.
    """

    top_parent = df[[COLS["account"]]]

    df = df.drop(COLS["account"], axis=1)
    
    df = df.map(lambda x: x.replace(",", "") if isinstance(x, str) else x)
    df = pd.concat([top_parent, df], axis=1)

    df = df.sort_values(
        by=[COLS["account"]],
        key=lambda col: col.str.lower()
    )

    bxl_tz = pytz.timezone("Europe/Brussels")
    time = dt.now(bxl_tz) - td(days=1)
    df[COLS["date"]] = time.strftime(DATE_FORMAT)
    
    return empty_cells_to_num(drop_empty_cells(df))

def powerbi_excel_data() -> pd.DataFrame:
    """
    Extracts and returns PowerBI accounts data as a pandas DataFrame.
    This function automates the process of exporting data from a PowerBI web page.
    It brings the PowerBI page to the front, locates the relevant table, and iteratively
    scrolls through the table to collect all visible rows, ensuring no duplicates are added.
    Once all data is collected, it constructs a DataFrame with appropriate headers and applies
    any necessary adjustments to the data.
    Returns:
        pd.DataFrame: A DataFrame containing the exported PowerBI accounts data.
    """

    log("Exporting PowerBI accounts data...", left_nl=1)

    powerbi = PAGES["powerbi"]
    powerbi.bring_to_front()

    table = wait_for(powerbi, POWERBI_DOM, at=["snapshot"])[0]
    table.scroll_into_view_if_needed()

    top = wait_for(table, POWERBI_DOM, at=["top"])[0]
    mid = wait_for(table, POWERBI_DOM, at=["mid"])[0]

    data = []
    row_num = 0
    seen_rows = set()

    while True:
        wait_for(mid,  {
            "next_row": (f".row[row-index='{row_num}']", None)
        })

        for row in [field.strip() for field in powerbi_row(mid)]:
            if row not in seen_rows:
                data.append(row.split("\n"))
                seen_rows.add(row)

        if is_locator_scrolled_to_bottom(mid):
            break

        mid.evaluate("el => el.scrollBy(0, el.clientHeight/3)")

        row_attr, _ = POWERBI_DOM["row"]
        index, _ = POWERBI_DOM["index"]

        last_row = mid.locator(row_attr).last
        row_num = int(last_row.get_attribute(index))
        powerbi.wait_for_timeout(500)
    
    header = powerbi_headers(top)
    df = adjust_powerbi_excel_data(
        pd.DataFrame(data, columns=header)
    )

    log("PowerBI accounts data exported")

    return df
    
def sharepoint_txt_data(desc:str, close_editor=True) -> pd.DataFrame:
    """
    Loads a text data file from SharePoint, reads it into a pandas DataFrame, and optionally closes the editor.
    Args:
        desc (str): The description or identifier of the SharePoint text file to load.
        close_editor (bool, optional): Whether to close the SharePoint editor after loading the data. Defaults to True.
    Returns:
        pd.DataFrame: The data from the SharePoint text file as a pandas DataFrame.
    Raises:
        Exception: If the specified SharePoint file is not found.
    Side Effects:
        Brings the SharePoint page to the front, interacts with the UI to open the file, and logs progress messages.
    """

    log(f"Loading SharePoint txt data ({desc})...", left_nl=1)

    sharepoint = PAGES["sharepoint"]
    sharepoint.bring_to_front()

    if desc not in SHAREPOINT_DOM:
        raise Exception("SharePoint file not found")

    txt_file = wait_for(sharepoint, SHAREPOINT_DOM, at=[desc])[0]
    txt_file.click()

    df = read_sharepoint_txt_data()
    
    if close_editor:
        sharepoint_close_editor()

    log(f"SharePoint txt data loaded ({desc})")

    return df

def fix_tam(progress: pd.DataFrame) -> pd.DataFrame:
    """
    Updates the TAM (Total Addressable Market) values in the given DataFrame based on actual TAM data,
    and recalculates the adoption percentage for each unique AE (Account Executive) ID.
    Args:
        progress (pd.DataFrame): The DataFrame containing progress data, including AE IDs, TAM, and adoption columns.
    Returns:
        pd.DataFrame: The updated DataFrame with corrected TAM values and recalculated adoption percentages.
    Notes:
        - Assumes the existence of global variables COLS (column mappings) and ACTUAL (actual TAM data).
        - The adoption percentage is recalculated as (adoption / TAM) * 100 for non-zero TAM values,
          and formatted as a percentage string with two decimal places.
    """

    t, a = COLS["tam"], COLS["adoption"]
    a_pct = COLS["adoption %"]
    actual_tam = ACTUAL["tam"]

    for ID in progress[COLS["ae"]].unique():
        if ID not in actual_tam["ID"].values:
            continue

        accounts = str(actual_tam.loc[actual_tam["ID"] == ID, "Accounts"].values[0]).split("|")
        tams = str(actual_tam.loc[actual_tam["ID"] == ID, "TAM"].values[0]).split(",")

        if len(accounts) != len(tams):
            raise Exception(
                f"Accounts and TAM lengths mismatch for AE ID {ID}: "
                f"{len(accounts)} accounts, {len(tams)} TAMs"
            )
        
        for account, tam in zip(accounts, tams):
            progress.loc[
                (progress[COLS["ae"]] == ID) & 
                (progress[COLS["account"]] == account), t
            ] = tam

        a_nozero = (progress.loc[progress[t] != 0, a]).astype(int)
        t_nozero = (progress.loc[progress[t] != 0, t]).astype(int)

        progress[a_pct] = a_nozero / t_nozero
        progress[a_pct] = progress[a_pct] * 100
        progress[a_pct] = progress[a_pct].map(lambda x: f"{x:.2f}%")
        
    return progress

def try_update_accounts_data() -> pd.DataFrame:
    """
    Updates the accounts data from SharePoint if the stored data is older than a week and not in test mode.
    Returns:
        pd.DataFrame: The combined DataFrame containing both the stored and newly updated data if an update occurs,
                      or just the stored data if no update is needed.
    Workflow:
        - Retrieves the current stored data from SharePoint.
        - If running in test mode or the stored data is not a week old, returns the stored data.
        - Otherwise, fetches new data from Power BI Excel, updates the SharePoint snapshot, and saves the changes.
        - Returns the concatenated DataFrame of stored and new data.
    Side Effects:
        - Interacts with the SharePoint UI to update and save data.
        - Logs progress and actions throughout the process.
    """

    sharepoint = PAGES["sharepoint"]
    stored_data = sharepoint_txt_data("snapshots", close_editor=False)

    if TEST["active"] or not is_week_old(stored_data):
        motive = "running test" if TEST["active"] else "not a week old"
        log(f"Returning stored data ({motive})")
        sharepoint_close_editor()
        return stored_data

    log("Updating SharePoint Copilot MAU snapshots...", left_nl=1)
    new_data = fix_tam(powerbi_excel_data())

    sharepoint.bring_to_front()

    id = "rows"
    line = wait_for(sharepoint, SHAREPOINT_DOM, at=[id], index={id: -1})[0]
    line.scroll_into_view_if_needed()
    line.click()

    log("Adding new snapshots data...")
    csv = new_data.to_csv(index=False, header=False, sep=";").strip()
    sharepoint.keyboard.press("Enter")
    sharepoint.keyboard.insert_text(csv)
    sharepoint.wait_for_timeout(500)

    save_btn = wait_for(sharepoint, SHAREPOINT_DOM, at=["save"])[0]
    save_btn.click()

    wait_for(sharepoint, SHAREPOINT_DOM, at=["popup"])[0]
    sharepoint_close_editor()

    log("Returning updated data")
    
    return pd.concat([stored_data, new_data], ignore_index=True)

def set_structure_variables(role:str, names:str, url:str="") -> pd.DataFrame:
    """
    Generates a structured DataFrame for a given role by replacing placeholders with provided names and URL.
    Args:
        role (str): The key representing the role to select the appropriate mail structure from MAIL_STRUCTURES.
        names (str): A comma-separated string of full names. Only the first names will be used for replacement.
        url (str, optional): The URL to replace the '*URL*' placeholder in the structure. Defaults to an empty string.
    Returns:
        pd.DataFrame: A DataFrame with the structure for the specified role, where '*NAME*' is replaced by the first names and '*URL*' by the provided URL.
    Raises:
        KeyError: If the specified role does not exist in MAIL_STRUCTURES.
    """

    first_names = [name.split(" ")[0] for name in names.split(",")]

    new_structure = MAIL_STRUCTURES[role].copy()
    new_structure["Text"] = new_structure["Text"].str.replace(
        "*NAME*",
        ", ".join(first_names)
    ).replace(
        "*URL*", 
        url
    )

    return new_structure

def read_sharepoint_txt_data() -> pd.DataFrame:
    """
    Reads tabular data from a SharePoint text editor interface and returns it as a pandas DataFrame.
    This function interacts with a SharePoint page, waits for the relevant DOM elements to load,
    and extracts row data separated by semicolons (";"). It handles potential row count mismatches
    by retrying the read operation up to five times. The function normalizes and accumulates the
    row data, handling cases where rows may be split across multiple reads, and finally constructs
    a DataFrame using the first row as column headers.
    Returns:
        pd.DataFrame: A DataFrame containing the parsed data from the SharePoint text editor,
        with columns inferred from the first row of the data.
    Raises:
        Exception: If the number of rows read does not match the expected row count after retries.
    """

    sharepoint = PAGES["sharepoint"]

    row_counter, sheet = wait_for(
        sharepoint, 
        SHAREPOINT_DOM, 
        at=["row count", "editor"]
    )

    sheet.filter(has_text=";").wait_for(state="visible")
    row_counter.scroll_into_view_if_needed()
    sheet.scroll_into_view_if_needed()

    r, c = "rows", "counter"
    attr, _ = SHAREPOINT_DOM["row"]
    observe_rows(sheet, r, sharepoint_rows(sheet), attr)
    observe_rows(row_counter, c, row_counter_info(row_counter))

    prev_row = ""
    data = []

    log(f"Reading SharePoint txt data...")
    while True:
        new_rows, counter = [], []

        for i in range(5):
            new_rows, counter = get_new_rows(r), get_new_rows(c)
            
            if len(new_rows) == len(counter):
                break
            if i == 4:
                raise Exception("Row count mismatch in the SharePoint editor")
            
            log(f"Row count mismatch in the SharePoint editor. Trying to read again ({i+1}/5)")
            sharepoint.wait_for_timeout(1000)
        
        if len(new_rows) == 0:
            data.append(prev_row.split(";"))
            break
        
        for i, row in enumerate(new_rows):
            row = str_normalize(row)

            if not prev_row:
                prev_row = row
                continue

            if counter[i]:
                data.append(prev_row.split(";"))
                prev_row = row
            else:
                prev_row += row

        sheet.evaluate("el => el.scrollBy(0, el.clientHeight*0.85)")
        row_counter.wait_for(state="visible")
        sharepoint.wait_for_timeout(500)
    
    return pd.DataFrame(data[1:], columns=data[0])
