from .dataframe import set_structure_variables
from .str_actions import str_normalize
from .imports import Callable, Any
from .logs import log
from vars.exports import (
    TEST_EXCLUDED_NAMES,
    TEST_EXCLUDED,
    MEDIA_LABELS,
    EXCLUDED,
    ACCOUNTS,
    TEST,
    SKIP,
    IMG,
    pd,
)
from .media import (
    capture_powerbi_elements, 
    plot_ae_progress, 
    img_to_b64
)


def prepare_mail(team:pd.Series, settings:dict[str,None|Any], execute:Callable[..., Any], params:dict[str,Any]={}) -> None:
    """
    Prepares email content by generating and storing necessary elements, then executes a provided function with updated parameters.
    Args:
        team (pd.Series): A pandas Series containing team information, including roles and names.
        settings (dict[str, None | Any]): A dictionary to store generated elements such as screenshots and structure variables.
        execute (Callable[..., Any]): A function to be called with the prepared parameters.
        params (dict[str, Any], optional): Additional parameters to pass to the execute function. Defaults to an empty dictionary.
    Returns:
        None
    Side Effects:
        - Modifies the `settings` dictionary in-place by adding or updating the "pics" and "structure" keys.
        - Modifies the `params` dictionary in-place by adding or updating the "pics" and "structure" keys.
        - Calls the `execute` function with the updated parameters.
    Notes:
        - If both "pics" and "structure" are already present in `settings`, the function will immediately call `execute` and return.
        - Relies on external functions: `capture_powerbi_elements`, `plot_ae_progress`, and `set_structure_variables`.
    """
    
    k1, k2 = ["pics", "structure"] 
    if all(settings[key] is not None for key in [k1, k2]):
        execute(**params)
        return

    screenshots, powerbi_ae_url = capture_powerbi_elements(team)
    pics = plot_ae_progress(team) + screenshots

    settings[k1] = pics
    settings[k2] = set_structure_variables(
        team["Role"],
        url=powerbi_ae_url,
        names=team["Names"].replace(",", ", ")
    )

    params[k1] = settings[k1]
    params[k2] = settings[k2]

    execute(**params)

def split_cc(emails:str, id:str, role:str) -> tuple[str, list[tuple[str, str]]]:
    """
    Splits a string of email addresses into a default CC list and a list of excluded emails based on provided rules.
    Args:
        emails (str): A space-separated string of email addresses to process.
        id (str): An identifier used to determine which emails to skip.
        role (str): The role used to determine which emails to exclude.
    Returns:
        tuple[str, list[tuple[str, str]]]: 
            - A string containing the default CC email addresses, separated by spaces.
            - A list of tuples, each containing an excluded email address and its associated name.
    Notes:
        - Emails present in the SKIP["emails"] DataFrame for the given id (or empty id) are skipped.
        - Emails present in the EXCLUDED[role] dictionary are excluded and added to the exclusions list.
        - Logging is performed for skipped and excluded emails.
    """

    emails = emails.split(" ")
    exclusions: list[tuple[str, str]] = []
    default_cc= ""

    to_skip = SKIP["emails"].copy()
    to_skip = set(to_skip.loc[to_skip["ID"].isin([id, ""]), "Email"].values)

    for email in emails:
        email = email.strip()

        if email in to_skip:
            log(f"Skipping {email} from being mailed...")
            continue

        if email in EXCLUDED[role]:
            name = EXCLUDED[role][email]["name"]
            log(f"Excluding {name} ({email}) from CC...")

            exclusions.append((email, name))
        else:
            if default_cc:
                default_cc += " "
            default_cc += email

    return default_cc, exclusions

def adjust_to(to:str, id:str) -> str:
    """
    Adjusts the 'to' email address string by removing any emails that should be skipped for the given ID.
    Args:
        to (str): A string containing one or more email addresses separated by spaces.
        id (str): The identifier used to determine which emails should be skipped.
    Returns:
        str: A string of email addresses, separated by spaces, with skipped emails removed.
    Notes:
        - Emails to be skipped are determined from the SKIP["emails"] DataFrame, where the "ID" matches the provided id or is empty.
        - Skipped emails are logged and not included in the returned string.
    """

    checked_to = ""
    to_skip = SKIP["emails"].copy()
    skips = set(to_skip.loc[to_skip["ID"].isin([id, ""]), "Email"].values)

    for email in to.split(" "):
        email = email.strip()

        if email in skips:
            log(f"Skipping {email} from being mailed...")
            continue
        
        if checked_to:
            checked_to += " "
        checked_to += email

    return checked_to.strip()

def handle_exclusions(exclusions:list[tuple[str, str]], structure:pd.DataFrame, pics:list[IMG], team:pd.Series) -> None:
    """
    Handles the exclusion of specific visuals for given email addresses by saving their corresponding images in a designated exclusion structure.
    Args:
        exclusions (list[tuple[str, str]]): A list of tuples containing email addresses and associated data to be excluded.
        structure (pd.DataFrame): A DataFrame containing command information, including which visuals to save.
        pics (list[IMG]): A list of image objects corresponding to available visuals.
        team (pd.Series): A Series containing team information such as Names, ID, and Role.
    Returns:
        None
    """

    saved_visuals = structure[
        structure["Command"].str.contains("visual") &
        structure["Command"].str.contains("save")
    ]

    for email, _ in exclusions:
        for _, row in saved_visuals.iterrows():
            text = f"{team["Names"]} ({team["ID"]}) - {row["Text"]}"
            
            _, visual_id = row["Command"].split(": ")
            visual = pics[MEDIA_LABELS[visual_id]]

            excluded = EXCLUDED[team["Role"]][email]
            excluded["pics"][text] = img_to_b64(visual)

    return

def define_exclusions(exclusions:pd.DataFrame) -> None:
    """
    Populates the global EXCLUDED dictionary with email exclusions based on V-team roles.
    Args:
        exclusions (pd.DataFrame): A DataFrame containing exclusion information with at least
            the columns "V-team role", "Email", and "FullName".
    Behavior:
        - For each unique V-team role in the exclusions DataFrame, normalizes the role name.
        - If TEST["active"] is True, uses test emails and names from TEST_EXCLUDED and TEST_EXCLUDED_NAMES
            to populate the EXCLUDED dictionary for each role.
        - If TEST["active"] is False, iterates through the exclusions DataFrame and adds each email and name
            (normalized) to the EXCLUDED dictionary under the corresponding role.
        - Raises an Exception if TEST["active"] is True and the number of test emails does not match
            the number of test names.
    Side Effects:
        - Modifies the global EXCLUDED dictionary in place.
    """

    roles = exclusions["V-team role"].unique()

    test_emails = TEST_EXCLUDED.split(" ")
    test_names = TEST_EXCLUDED_NAMES.split(", ")

    if TEST["active"] and len(test_emails) != len(test_names):
        raise Exception("Test emails and names mismatch")

    for role in roles:
        role = str_normalize(role)

        if role not in EXCLUDED:
            EXCLUDED[role] = {}

        if TEST["active"]:
            for i, email in enumerate(test_emails):
                EXCLUDED[role][email] = {"name": test_names[i], "pics": {}}
        else:
            for _, row in exclusions.loc[exclusions["V-team role"] == role].iterrows():
                email = str_normalize(row["Email"])
                name = str_normalize(row["FullName"])
                EXCLUDED[role][email] = {"name": name, "pics": {}}

def define_filter(id:str, filter:pd.DataFrame) -> None:
    """
    Assigns a filter to the global ACCOUNTS dictionary based on the provided ID and filter DataFrame.
    Parameters:
        id (str): The identifier used to filter the DataFrame.
        filter (pd.DataFrame): A DataFrame containing at least 'ID' and 'Accounts' columns.
    Side Effects:
        - Updates the global ACCOUNTS dictionary by setting the "filter" key to a list of accounts associated with the given ID.
        - If no accounts are found for the given ID, sets ACCOUNTS["filter"] to None and logs a message.
    Notes:
        - Assumes the existence of a global ACCOUNTS dictionary and a log function.
        - The 'Accounts' column should contain comma-separated account strings.
    """

    account_list = filter.loc[filter["ID"] == id, "Accounts"].values

    if len(account_list) == 0:
        log(f"V-team with {id} has no filters", left_nl=1)
        ACCOUNTS["filter"] = None
        return

    ACCOUNTS["filter"] = account_list[0].split("|")
