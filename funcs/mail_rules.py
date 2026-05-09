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
	img_to_b64,
)


#* ---------------------------------------------------------------------------
#* Exceptions
#* ---------------------------------------------------------------------------

class TestExclusionMismatchError(ValueError):
	"""Raised when the number of test exclusion emails and names differ."""

	def __init__(self, n_emails: int, n_names: int) -> None:
		super().__init__(
			f"TEST_EXCLUDED has {n_emails} email(s) but TEST_EXCLUDED_NAMES has {n_names} name(s)."
		)


#* ---------------------------------------------------------------------------
#* Mail preparation
#* ---------------------------------------------------------------------------

_SETTINGS_KEYS = ("pics", "structure")


def prepare_mail(
	team: pd.Series,
	settings: dict[str, Any | None],
	execute: Callable[..., Any],
	params: dict[str, Any] | None = None,
) -> None:
	"""
	Populate settings with screenshots and mail structure, then call execute.

	If both "pics" and "structure" are already present in settings (i.e. a
	previous call already generated them), execute is called immediately with
	the existing params to avoid redundant captures.

	Args:
		team:     Series row for the current v-team (must contain "Role" and "Names").
		settings: Shared dict that caches generated pics and structure across
					multiple calls for the same team.
		execute:  Callable invoked with the final merged params.
		params:   Extra keyword arguments forwarded to execute. Defaults to {}.

	Side effects:
		Mutates settings and params in-place; calls execute.
	"""
	#* Default to an empty dict rather than a mutable default argument
	if params is None:
		params = {}

	#* Short-circuit if a previous prepare_mail call already generated the assets
	if all(settings[key] is not None for key in _SETTINGS_KEYS):
		execute(**params)
		return

	screenshots, powerbi_ae_url = capture_powerbi_elements(team)
	pics = plot_ae_progress(team) + screenshots

	settings["pics"] = pics
	settings["structure"] = set_structure_variables(
		team["Role"],
		url=powerbi_ae_url,
		#* Normalise "Last, First" formatting to "Last, First" with consistent spacing
		names=team["Names"].replace(",", ", "),
	)

	#* Propagate the freshly generated assets into params before dispatching
	params.update({key: settings[key] for key in _SETTINGS_KEYS})
	execute(**params)


#* ---------------------------------------------------------------------------
#* Email address filtering
#* ---------------------------------------------------------------------------

def _build_skip_set(id: str) -> set[str]:
	"""
	Return the set of email addresses that should be skipped for the given AE ID.

	Includes both ID-specific skips and global skips (where ID is an empty string).
	"""
	to_skip = SKIP["emails"]
	return set(to_skip.loc[to_skip["ID"].isin([id, ""]), "Email"].values)


def adjust_to(to: str, ae_id: str) -> str:
	"""
	Remove skipped addresses from a space-separated recipient string.

	Args:
		to:    Space-separated email addresses.
		ae_id: AE identifier used to look up applicable skips.

	Returns:
		Filtered space-separated email string, or an empty string if all
		addresses were skipped.
	"""
	skips = _build_skip_set(ae_id)
	kept: list[str] = []

	for email in to.split():
		if email in skips:
			log(f"Skipping {email} from being mailed...")
			continue
		kept.append(email)

	return " ".join(kept)


def split_cc(
	emails: str, ae_id: str, role: str
) -> tuple[str, list[tuple[str, str]]]:
	"""
	Partition a space-separated CC string into default recipients and exclusions.

	An address is excluded when it appears in EXCLUDED[role]; it is skipped
	entirely when it appears in the SKIP list for the given AE ID.

	Args:
		emails: Space-separated email addresses to partition.
		ae_id:  AE identifier used to look up applicable skips.
		role:   V-team role used to look up exclusion rules.

	Returns:
		A tuple of:
			- default_cc: space-separated string of non-excluded, non-skipped addresses.
			- exclusions: list of (email, name) pairs for excluded addresses.
	"""
	skips = _build_skip_set(ae_id)
	exclusions: list[tuple[str, str]] = []
	kept: list[str] = []

	for email in emails.split():
		if email in skips:
			log(f"Skipping {email} from being mailed...")
			continue

		if email in EXCLUDED[role]:
			name = EXCLUDED[role][email]["name"]
			log(f"Excluding {name} ({email}) from CC...")
			exclusions.append((email, name))
		else:
			kept.append(email)

	return " ".join(kept), exclusions


#* ---------------------------------------------------------------------------
#* Exclusion handling
#* ---------------------------------------------------------------------------

def handle_exclusions(
	exclusions: list[tuple[str, str]],
	structure: pd.DataFrame,
	pics: list[IMG],
	team: pd.Series,
) -> None:
	"""
	Store the relevant visual captures for each excluded email address.

	Iterates the "save visual" commands in structure and maps each to its
	corresponding image, storing a base-64 encoded copy in EXCLUDED so it
	can be attached when the exclusion mail is sent later.

	Args:
		exclusions: List of (email, name) pairs to process.
		structure:  Mail structure DataFrame; rows with a "save visual: <id>"
						command identify which pics to capture.
		pics:       Ordered list of captured images (indexed via MEDIA_LABELS).
		team:       Series row for the current v-team.
	"""
	#* Identify rows that instruct us to save a specific visual for excluded recipients
	saved_visuals = structure[
		structure["Command"].str.contains("visual", na=False) &
		structure["Command"].str.contains("save", na=False)
	]

	role = team["Role"]

	for email, _ in exclusions:
		excluded_entry = EXCLUDED[role][email]

		for _, row in saved_visuals.iterrows():
			#* Command format: "save visual: <visual_id>"
			_, visual_id = row["Command"].split(": ", maxsplit=1)
			visual = pics[MEDIA_LABELS[visual_id]]

			caption = f"{team['Names']} ({team['ID']}) - {row['Text']}"
			excluded_entry["pics"][caption] = img_to_b64(visual)


def define_exclusions(exclusions: pd.DataFrame) -> None:
	"""
	Populate the global EXCLUDED dict from a SharePoint exclusions DataFrame.

	In test mode, TEST_EXCLUDED and TEST_EXCLUDED_NAMES are used instead of
	the real data, and every role receives the same test addresses.

	Args:
		exclusions: DataFrame with at least "V-team role", "Email", and
						"FullName" columns.

	Raises:
		TestExclusionMismatchError: If test mode is active and the number of
											test emails and names differ.
	"""
	roles = exclusions["V-team role"].unique()

	test_emails = []
	test_names  = []
	if TEST["active"]:
		test_emails = TEST_EXCLUDED.split()
		test_names  = TEST_EXCLUDED_NAMES.split(", ")

		if len(test_emails) != len(test_names):
			raise TestExclusionMismatchError(len(test_emails), len(test_names))

	for role in roles:
		role = str_normalize(role)
		EXCLUDED.setdefault(role, {})

		if TEST["active"]:
			for email, name in zip(test_emails, test_names):
					EXCLUDED[role][email] = {"name": name, "pics": {}}
		else:
			role_rows = exclusions.loc[exclusions["V-team role"] == role]
			for _, row in role_rows.iterrows():
				email = str_normalize(row["Email"])
				name  = str_normalize(row["FullName"])
				EXCLUDED[role][email] = {"name": name, "pics": {}}


#* ---------------------------------------------------------------------------
#* Account filter
#* ---------------------------------------------------------------------------

def define_filter(ae_id: str, accounts_filter: pd.DataFrame) -> None:
	"""
	Set ACCOUNTS["filter"] to the account list for the given AE ID.

	Args:
		ae_id:           AE identifier to look up in the filter DataFrame.
		accounts_filter: DataFrame with "ID" and "Accounts" (pipe-separated) columns.

	Side effects:
		Sets ACCOUNTS["filter"] to a list of account strings, or None if the
		AE ID has no filter entry.
	"""
	matches = accounts_filter.loc[accounts_filter["ID"] == ae_id, "Accounts"].values

	if len(matches) == 0:
		log(f"V-team {ae_id!r} has no account filters.", left_nl=1)
		ACCOUNTS["filter"] = None
		return

	ACCOUNTS["filter"] = matches[0].split("|")
