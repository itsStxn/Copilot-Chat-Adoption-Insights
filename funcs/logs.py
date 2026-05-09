from vars.exports import LOGS, MAIL_LOGS, TEST, os
from .imports import dt


#* ---------------------------------------------------------------------------
#* Console
#* ---------------------------------------------------------------------------

def clear_console() -> None:
	"""Clear the terminal screen on both Windows and Unix-like systems."""
	os.system("cls" if os.name == "nt" else "clear")


#* ---------------------------------------------------------------------------
#* File I/O
#* ---------------------------------------------------------------------------

def read_file_as_str(path: str) -> str:
	"""
	Read a file and return its contents as a string.

	Raises:
		FileNotFoundError: If no file exists at path.
		IOError: If the file cannot be read.
	"""
	with open(path, "r", encoding="utf-8") as f:
		return f.read()


def write_to_file(path: str, content: str) -> None:
	"""Write content to a file, appending a trailing newline."""
	with open(path, "w", encoding="utf-8") as f:
		f.write(content + "\n")


def delete_file(path: str) -> None:
	"""
	Delete a file if it exists, refusing paths outside the project root.

	Raises:
		ValueError: If path resolves to a location outside the project root.
	"""
	root = os.path.abspath(os.getcwd())
	abs_path = os.path.abspath(path)

	if not abs_path.startswith(root):
		raise ValueError(f"File {path!r} is outside the project root and cannot be deleted.")

	try:
		os.remove(abs_path)
		print(f"File {abs_path!r} deleted.")
	except FileNotFoundError:
		print(f"File {abs_path!r} was never created — nothing to delete.")


#* ---------------------------------------------------------------------------
#* Log parsing
#* ---------------------------------------------------------------------------

def _parse_log_lines(lines: list[str]) -> set[str]:
	"""
	Extract the message portion from a sequence of log lines.

	Each valid line is expected to match the pattern "<timestamp> - <message>".
	Lines that are blank or lack the separator are silently skipped.

	Returns:
		A set of unique message strings.
	"""
	return {
		line.split(" - ", maxsplit=1)[1].strip()
		for line in lines
		if line.strip() and " - " in line
	}


def logs_to_set(text: str) -> set[str]:
	"""
	Parse a multiline log string into a set of unique message strings.

	Raises:
		ValueError: If text is empty or contains only whitespace.
	"""
	if not text.strip():
		raise ValueError("Cannot parse an empty log string.")

	return _parse_log_lines(text.splitlines())


def logs_file_to_set(path: str) -> set[str]:
	"""
	Read a log file and return its messages as a set of unique strings.

	Raises:
		FileNotFoundError: If no file exists at path.
	"""
	try:
		with open(path, "r", encoding="utf-8") as f:
			return _parse_log_lines(f.readlines())
	except FileNotFoundError:
		raise FileNotFoundError(f"Log file not found: {path!r}")


#* ---------------------------------------------------------------------------
#* Logging
#* ---------------------------------------------------------------------------

def log(
	message: str = "",
	type: dict = LOGS,
	write: bool = False,
	silent: bool = False,
	left_nl: int = 0,
	right_nl: int = 1,
	from_last: bool = False,
) -> None:
	"""
	Append a timestamped message to an in-memory log buffer, optionally flushing to disk.

	Each non-empty line in message is stored as "<timestamp> - <line>".
	When write=True the buffer is persisted only if it contains at least one
	meaningful entry (START and END markers are excluded from the count).

	Args:
		message:   Text to log. Multi-line strings are split and stored line by line.
		type:      Log target dict with "path" and "content" keys. Defaults to LOGS.
		write:     Flush the buffer to type["path"] after appending. Defaults to False.
		silent:    Suppress stdout output. Defaults to False.
		left_nl:   Blank lines to prepend before each stored entry. Defaults to 0.
		right_nl:  Blank lines to append after each stored entry. Defaults to 1.
		from_last: Prepend the existing log file's content to the buffer before
						appending. Useful for resuming across sessions. Defaults to False.
	"""
	prefix = "\n" * left_nl
	suffix = "\n" * right_nl
	timestamp = dt.now().strftime("%Y-%m-%d %H:%M:%S")

	#* Seed the in-memory buffer with whatever was written in a previous session
	if from_last and os.path.exists(type["path"]):
		type["content"] = read_file_as_str(type["path"]) + type["content"]

	if message.strip():
		for line in message.splitlines():
			if line.strip():
				type["content"] += f"{prefix}{timestamp} - {line}{suffix}"

		if not silent:
			print(f"{prefix}{message}{suffix}", end="")

	if write:
		#* Exclude bookkeeping markers before deciding whether to persist
		entries = logs_to_set(type["content"]) - {"START", "END"}
		if entries:
			write_to_file(type["path"], type["content"])


#* ---------------------------------------------------------------------------
#* Mail log helpers
#* ---------------------------------------------------------------------------

def mail_log(to: str, cc: str, role: str) -> str:
	"""Return the canonical log string for a sent v-team email."""
	return f"Sent => V-team's role: {role}, To: {to}, CC: {cc}".strip()


def mail_log_exists(to: str, cc: str, role: str) -> bool:
	"""
	Return True if a matching mail log entry already exists in the buffer.

	Avoids redundant sends by checking the in-memory MAIL_LOGS buffer before
	attempting to deliver a message.
	"""
	content = MAIL_LOGS["content"]

	#* An empty buffer means nothing has been sent yet in this session
	if not content:
		return False

	return mail_log(to, cc, role) in logs_to_set(content)


#* ---------------------------------------------------------------------------
#* Interactive prompts
#* ---------------------------------------------------------------------------

_LOG_ACTIONS: dict[str, str] = {
	"1": "Keep logs",
	"2": "Discard logs",
}

_MODES: dict[str, str] = {
	"1": "Test mode",
	"2": "Update mode",
	"3": "Rollout mode",
}


def choose_keep_logs() -> bool:
	"""
	Prompt the user to keep or discard logs after a run.

	Returns:
		True if the user chooses to keep logs, False to discard.
	"""
	log("Choose action:")
	for key, label in _LOG_ACTIONS.items():
		log(f"{key}. {label}")

	while True:
		choice = input("Enter your choice (1 or 2): ").strip()

		if choice in _LOG_ACTIONS:
			log(f"{_LOG_ACTIONS[choice]} selected", left_nl=1)
			return choice == "1"

		print("Invalid choice. Please enter 1 or 2.")


#* ---------------------------------------------------------------------------
#* Log path adjustment
#* ---------------------------------------------------------------------------

def adjust_logs() -> None:
	"""
	Prefix log file paths with "test_" when running in test mode, then
	seed both in-memory buffers from their respective files on disk.

	Seeding via from_last=True ensures that log entries from previous
	sessions are not lost when the same log file is reused.
	"""
	if TEST["active"]:
		#* Redirect all output to test-specific files so real logs are not polluted
		LOGS["path"]      = f"test_{LOGS['path']}"
		MAIL_LOGS["path"] = f"test_{MAIL_LOGS['path']}"

	log(from_last=True)
	log(type=MAIL_LOGS, from_last=True)
