from .imports import sync_playwright, tb
from .outlook import send_new_mail
from vars.exports import (
    MAIL_STRUCTURES,
    DEFAULT_EMAILS,
    MAIL_LOGS,
    EXCLUDED,
    ACCOUNTS,
    BROWSER,
    TEST_TO,
    TEST_CC,
    PAGES,
    ACTUAL,
    TEST,
    LOGS,
    SKIP,
    URLS,
    pd
)
from .mail_rules import (
    handle_exclusions, 
    define_exclusions, 
    define_filter, 
    prepare_mail, 
    adjust_to, 
    split_cc
  )
from .logs import (
    choose_keep_logs,
    mail_log_exists, 
    clear_console, 
    adjust_logs, 
    delete_file, 
    log
  )
from .dataframe import (
    set_structure_variables, 
    sharepoint_txt_data, 
    try_update_accounts_data
  )
from .browser import (
    find_work_profile,
    open_outlook,
    goto
)


def notify_v_team(team:pd.Series) -> bool:
    """
    Notifies a v-team by preparing and sending emails based on the provided team information.
    Args:
        team (pd.Series): A pandas Series containing team information, including 'ID', 'Role', 'Names', 'To', and 'CC'.
    Returns:
        bool: True if the v-team was processed (emails prepared/sent or already sent), False if skipped due to missing recipients.
    Workflow:
        - Logs the start of v-team processing.
        - Determines the recipient ('to') and CC list, adjusting for test mode if active.
        - Skips processing if no recipient is found.
        - Handles exclusions by preparing mails for excluded emails if not already logged.
        - Prepares and sends a new mail to the v-team if not already sent, otherwise logs that the mail was already sent.
        - Logs completion of v-team processing.
    """

    id = team["ID"]
    role = team["Role"]
    names = team["Names"]

    log(f'Processing v-team of the {role}/s {names}...', left_nl=1)

    to = adjust_to(
        TEST_TO if TEST["active"] else team["To"].strip(),
        id
    )

    if not to:
        log(f"Skipping {team['Names']} ({team['ID']})...")
        return False
    
    cc, exclusions = split_cc(
        TEST_CC if TEST["active"] else f"{team['CC']} {DEFAULT_EMAILS}".strip(),
        id,
        role
    )

    settings = {
        "pics": None,
        "structure": None
    }

    for email, _ in exclusions:
        if not mail_log_exists(email, "", role):
            prepare_mail(
                team,
                settings,
                execute=handle_exclusions,
                params={
                    "exclusions": exclusions,
                    "team": team
                }
            )

    if not mail_log_exists(
            adjust_to(team["To"], id) if TEST["active"] else to,
            cc,
            role
        ):
        prepare_mail(
            team,
            settings,
            execute=send_new_mail,
            params={
                "to": to,
                "cc": cc,
                "role": role,
                "names": names,
                "pics": settings["pics"],
                "structure": settings["structure"]
            }
        )
    else:
        log(f"A mail already sent to the v-team of {names} ({to})")

    log("V-team processed")

    return True

def notify_v_teams(v_teams:pd.DataFrame, filter:pd.DataFrame, dialog:str) -> None:
    """
    Notifies a list of V-teams by iterating over the provided DataFrame, applying a filter, and sending notifications.
    Args:
        v_teams (pd.DataFrame): DataFrame containing V-team information. Each row represents a V-team.
        filter (pd.DataFrame): DataFrame used to filter or define criteria for each V-team.
        dialog (str): Message to display or log when notifying each V-team.
    Returns:
        None
    Side Effects:
        - Calls define_filter for each V-team.
        - Clears the console for each iteration.
        - Notifies each V-team using notify_v_team.
        - Logs the notification dialog if TEST["active"] is True.
        - Prompts user input and breaks the loop if the user enters "q".
    """
    
    for _, v_team in v_teams.iterrows():
        define_filter(v_team["ID"], filter)
        clear_console()

        dialog = f"V-teams notified. {dialog}"
        if notify_v_team(v_team) and TEST["active"]:
            log(dialog, left_nl=1, silent=True)
            if input(dialog).lower() == "q":
                break

def notify_exclusions(dialog:str) -> None:
    """
    Notifies excluded users by sending them emails and logging the notification.
    Args:
        dialog (str): The message to include in the notification dialog.
    Side Effects:
        - Sends emails to users listed in the EXCLUDED dictionary who have associated pictures.
        - Logs the notification dialog if TEST["active"] is True.
        - Prompts for user input if TEST["active"] is True; breaks the loop if input is 'q'.
    Dependencies:
        - EXCLUDED: A dictionary containing roles, emails, and associated user data.
        - URLS: A dictionary containing URLs for various services.
        - TEST: A dictionary indicating test mode status.
        - set_structure_variables: Function to prepare email structure variables.
        - send_new_mail: Function to send emails.
        - log: Function to log messages.
    """
    
    dialog = f"Exclusions notified. {dialog}"
    for role in EXCLUDED:
        for email in EXCLUDED[role]:
            data = EXCLUDED[role][email]
            if data["pics"]:
                structure = set_structure_variables(
                    url=URLS["powerbi show"],
                    names=data["name"],
                    role="excluded"
                )
                send_new_mail(
                    to=email,
                    cc="",
                    role=role,
                    names=data["name"],
                    structure=structure,
                    collected_pics=data["pics"]
                )
            if TEST["active"]:
                log(dialog, left_nl=1, silent=True)
                if input(dialog).lower() == "q":
                    break

def notification_flow() -> None:
    """
    Executes the notification workflow by retrieving configuration data from SharePoint,
    setting up exclusions and mail structures, navigating between application pages,
    and triggering notification functions for V-teams and exclusions.
    Steps performed:
    1. Loads filter, skip, V-teams, exclusions, and mail structure data from SharePoint.
    2. Defines exclusions based on the retrieved data.
    3. Closes the SharePoint page and navigates to the Power BI page.
    4. Opens Outlook and prepares a dialog prompt.
    5. Notifies V-teams and exclusions using the prepared dialog.
    Returns:
        None
    """

    filter = sharepoint_txt_data("filter accounts")
    SKIP["emails"] = sharepoint_txt_data("skip")
    v_teams = sharepoint_txt_data("V-teams")

    define_exclusions(sharepoint_txt_data("excluded"))
    
    MAIL_STRUCTURES["AE"] = sharepoint_txt_data("AE mail")
    MAIL_STRUCTURES["excluded"] = sharepoint_txt_data("excluded mail")

    PAGES["sharepoint"].close()
    
    goto(PAGES["powerbi"], URLS["powerbi show"])
    PAGES["outlook"] = open_outlook()

    dialog = "Press Enter to continue or 'Q' to quit: "
    notify_v_teams(v_teams, filter, dialog)
    notify_exclusions(dialog)

def choose_mode() -> str:
    """
    Prompts the user to select an operation mode from three options: Test mode, Update mode, or Rollout mode.
    Displays the available modes, accepts user input, and validates the selection.
    Returns the selected mode as a string ("1", "2", or "3").
    Logs the selected mode and prompts again if the input is invalid.
    Returns:
        str: The user's selected mode ("1", "2", or "3").
    """
    
    log("Choose mode:")
    log("1. Test mode")
    log("2. Update mode")
    log("3. Rollout mode")

    modes = {
        "1": "Test mode",
        "2": "Update mode",
        "3": "Rollout mode"
    }

    while True:
        choice = input("Enter your choice (1, 2 or 3): ").strip()
        if choice in modes:
            log(f"{modes[choice]} selected", left_nl=1)
            return choice
        else:
            print("Invalid choice. Please enter 1, 2 or 3.")

def run_flow() -> None:
    """
    Executes the main automation flow for the application.
    This function initializes logging, selects the operation mode, manages browser sessions,
    updates account data, and optionally triggers notificatKion flows. It also handles log
    retention based on user input and ensures proper cleanup of resources. Any exceptions
    encountered during execution are logged with detailed tracebacks.
    Raises:
        Exception: Logs any exception that occurs during the flow execution.
    """

    start, end = "START", "END"
    log(start, silent=True, right_nl=2)
    log(start, silent=True, type=MAIL_LOGS)
    keep_logs = True
    
    try:
        with sync_playwright() as p:
            choice = choose_mode()
            TEST["active"] = choice == "1"
            adjust_logs()

            BROWSER["edge"], PAGES["powerbi"], PAGES["sharepoint"] = find_work_profile(p)
            ACTUAL["tam"] = sharepoint_txt_data("actual tam")
            ACCOUNTS["data"] = try_update_accounts_data()

            if choice != "2":
                notification_flow()
                
            BROWSER["edge"].close()
            keep_logs = choose_keep_logs()

            if not keep_logs:
                delete_file(LOGS["path"])
                delete_file(MAIL_LOGS["path"])
    except Exception as e:
        keep_logs = True
        log(f"Error: {str(e)}", left_nl=1, right_nl=2)

        for error_msg in tb.format_exc().split("\n"):
            if error_msg.strip():
                log(error_msg)

    if keep_logs:
        log(end, silent=True, write=True, left_nl=1, right_nl=3)
        log(end, silent=True, write=True, type=MAIL_LOGS, right_nl=3)
