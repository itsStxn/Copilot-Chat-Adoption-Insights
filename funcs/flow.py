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
	pd,
)
from .mail_rules import (
	handle_exclusions,
	define_exclusions,
	define_filter,
	prepare_mail,
	adjust_to,
	split_cc,
)
from .logs import (
	choose_keep_logs,
	mail_log_exists,
	clear_console,
	adjust_logs,
	delete_file,
	log,
)
from .dataframe import (
	set_structure_variables,
	sharepoint_txt_data,
	try_update_accounts_data,
)
from .browser import (
	find_work_profile,
	open_outlook,
	goto,
)


#* ---------------------------------------------------------------------------
#* Mode selection
#* ---------------------------------------------------------------------------

_MODES: dict[str, str] = {
	"1": "Test mode",
	"2": "Update mode",
	"3": "Rollout mode",
}


def choose_mode() -> str:
	"""
	Prompt the user to select an operation mode.

	Returns:
		The selected mode key: "1" (Test), "2" (Update), or "3" (Rollout).
	"""
	log("Choose mode:")
	for key, label in _MODES.items():
		log(f"{key}. {label}")

	while True:
		choice = input("Enter your choice (1, 2 or 3): ").strip()

		if choice in _MODES:
			log(f"{_MODES[choice]} selected", left_nl=1)
			return choice

		print("Invalid choice. Please enter 1, 2 or 3.")


#* ---------------------------------------------------------------------------
#* V-team notification
#* ---------------------------------------------------------------------------

def notify_v_team(team: pd.Series) -> bool:
	"""
	Prepare and send notification emails for a single v-team.

	Handles both exclusion emails and the main v-team mail. Returns False
	early if no valid recipient can be determined.

	Args:
		team: A Series row from the v-teams DataFrame, expected to contain
				at least "ID", "Role", "Names", "To", and "CC".

	Returns:
		True if the v-team was processed (or already sent), False if skipped.
	"""
	ae_id: str = team["ID"]
	role: str  = team["Role"]
	names: str = team["Names"]

	log(f"Processing v-team of the {role}/s {names}...", left_nl=1)

	#* In test mode, route all mail to the test address instead of the real recipient
	raw_to: str = TEST_TO if TEST["active"] else team["To"].strip()
	to: str = adjust_to(raw_to, ae_id)

	if not to:
		log(f"Skipping {names} ({ae_id}) — no valid recipient.")
		return False

	#* In test mode, use the test CC list; otherwise merge the team CC with the default addresses
	raw_cc: str = TEST_CC if TEST["active"] else f"{team['CC']} {DEFAULT_EMAILS}".strip()
	cc, exclusions = split_cc(raw_cc, ae_id, role)

	#* settings is shared across prepare_mail calls so that the pics/structure
	#* resolved during exclusion handling can be reused for the main mail
	settings: dict[str, None] = {"pics": None, "structure": None}

	#* Send exclusion emails for any excluded address that hasn't been mailed yet
	for email, _ in exclusions:
		if not mail_log_exists(email, "", role):
			prepare_mail(
				team,
				settings,
				execute=handle_exclusions,
				params={"exclusions": exclusions, "team": team},
			)

	#* Determine the address to check in the mail log:
	#* in test mode the actual recipient is the real address, not the test redirect
	log_address: str = adjust_to(team["To"], ae_id) if TEST["active"] else to

	if not mail_log_exists(log_address, cc, role):
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
				"structure": settings["structure"],
			},
		)
	else:
		log(f"Mail already sent to the v-team of {names} ({to}).")

	log("V-team processed.")
	return True


def notify_v_teams(v_teams: pd.DataFrame, accounts_filter: pd.DataFrame, dialog: str) -> None:
	"""
	Iterate over all v-teams, applying per-team filters and sending notifications.

	Clears the console between teams. In test mode, pauses after each team
	and exits early if the user enters "q".

	Args:
		v_teams: DataFrame where each row represents one v-team.
		accounts_filter: Filter DataFrame passed to define_filter for each team.
		dialog: Prompt string shown to the user between iterations in test mode.
	"""
	progress_dialog = f"V-teams notified. {dialog}"

	for _, v_team in v_teams.iterrows():
		#* Narrow the global account filter to only the rows relevant to this team
		define_filter(v_team["ID"], accounts_filter)
		clear_console()

		if notify_v_team(v_team) and TEST["active"]:
			log(progress_dialog, left_nl=1, silent=True)
			if input(progress_dialog).lower() == "q":
				break


def notify_exclusions(dialog: str) -> None:
	"""
	Send notification emails to all excluded users who have associated pictures.

	Iterates EXCLUDED by role then by address. Skips addresses whose pics
	collection is empty. In test mode, pauses after each send and exits early
	if the user enters "q".

	Args:
		dialog: Prompt string shown to the user between sends in test mode.
	"""
	progress_dialog = f"Exclusions notified. {dialog}"

	for role, addresses in EXCLUDED.items():
		for email, data in addresses.items():
			#* Only mail excluded users for whom pictures have been collected
			if not data["pics"]:
				continue

			structure = set_structure_variables(
				url=URLS["powerbi show"],
				names=data["name"],
				role="excluded",
			)
			send_new_mail(
				to=email,
				cc=DEFAULT_EMAILS.strip(),
				role=role,
				names=data["name"],
				structure=structure,
				collected_pics=data["pics"],
			)

			if TEST["active"]:
				log(progress_dialog, left_nl=1, silent=True)
				if input(progress_dialog).lower() == "q":
					return


#* ---------------------------------------------------------------------------
#* Notification flow
#* ---------------------------------------------------------------------------

def notification_flow() -> None:
	"""
	Orchestrate the full notification workflow.

	Loads all required configuration from SharePoint, sets up exclusions and
	mail structures, navigates the browser to the correct pages, then fires
	notifications for v-teams and excluded users.
	"""
	#* Load all SharePoint configuration data up front before closing the page
	accounts_filter = sharepoint_txt_data("filter accounts")
	SKIP["emails"]  = sharepoint_txt_data("skip")
	v_teams         = sharepoint_txt_data("V-teams")

	define_exclusions(sharepoint_txt_data("excluded"))

	MAIL_STRUCTURES["AE"]       = sharepoint_txt_data("AE mail")
	MAIL_STRUCTURES["excluded"] = sharepoint_txt_data("excluded mail")

	#* SharePoint is no longer needed; switch the browser to the Power BI report view
	PAGES["sharepoint"].close()
	goto(PAGES["powerbi"], URLS["powerbi show"])
	PAGES["outlook"] = open_outlook()

	dialog = "Press Enter to continue or 'Q' to quit: "
	notify_v_teams(v_teams, accounts_filter, dialog)
	notify_exclusions(dialog)


#* ---------------------------------------------------------------------------
#* Entry point
#* ---------------------------------------------------------------------------

def run_flow() -> None:
	"""
	Execute the end-to-end automation flow.

	Selects the operating mode, opens a persistent Edge session, refreshes
	account data, and — unless Update mode was chosen — runs the full
	notification flow. Handles log retention and cleans up the browser on exit.
	Any unhandled exception is caught, logged with a full traceback, and causes
	logs to be preserved regardless of the user's preference.
	"""
	log("START", silent=True, right_nl=2)
	log("START", silent=True, type=MAIL_LOGS)

	#* keep_logs is flipped to False only if the user explicitly opts out after a clean run
	keep_logs = True

	try:
		with sync_playwright() as p:
			choice = choose_mode()
			TEST["active"] = choice == "1"
			adjust_logs()

			#* Launch the browser and hydrate shared state before any other work
			BROWSER["edge"], PAGES["powerbi"], PAGES["sharepoint"] = find_work_profile(p)
			ACTUAL["tam"]    = sharepoint_txt_data("actual tam")
			ACCOUNTS["data"] = try_update_accounts_data()

			#* Update mode (choice "2") only refreshes data; no notifications are sent
			if choice != "2":
					notification_flow()

			BROWSER["edge"].close()
			keep_logs = choose_keep_logs()

			if not keep_logs:
					delete_file(LOGS["path"])
					delete_file(MAIL_LOGS["path"])

	except Exception as e:
		#* Always preserve logs when something goes wrong so the run can be diagnosed
		keep_logs = True
		log(f"Error: {e}", left_nl=1, right_nl=2)

		for line in tb.format_exc().splitlines():
			if line.strip():
					log(line)

	if keep_logs:
		log("END", silent=True, write=True, left_nl=1, right_nl=3)
		log("END", silent=True, write=True, type=MAIL_LOGS, right_nl=3)
