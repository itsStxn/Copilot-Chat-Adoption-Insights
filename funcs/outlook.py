from .page import wait_for, click_and_wait
from .media import img_to_b64, imgs_to_b64
from .logs import log, mail_log
from .imports import Locator
from vars.exports import (
	INCENTIVES_SLIDE_PATH,
	MEDIA_LABELS,
	OUTLOOK_DOM,
	MAIL_LOGS,
	PAGES,
	MEDIA,
	IMG,
	pd,
)


#* ---------------------------------------------------------------------------
#* Exceptions
#* ---------------------------------------------------------------------------

class NoRecipientsError(ValueError):
	"""Raised when a 'To' field is submitted with no email addresses."""

	def __init__(self) -> None:
		super().__init__("At least one recipient address is required in the 'To' field.")


class MissingVisualError(KeyError):
	"""Raised when a visual ID referenced in the mail structure cannot be resolved."""

	def __init__(self, visual_id: str) -> None:
		super().__init__(f"Visual ID {visual_id!r} not found in MEDIA, MEDIA_LABELS, or known aliases.")


#* ---------------------------------------------------------------------------
#* Low-level Outlook interactions
#* ---------------------------------------------------------------------------

def _outlook():
	"""Return the shared Outlook page, accessed via the module-level PAGES dict."""
	return PAGES["outlook"]


def mail_attach_image(body: Locator, pic: str, wait_time: int = 500) -> None:
	"""
	Paste a base64-encoded PNG into the Outlook email body via a synthetic ClipboardEvent.

	Moves to a new line before and after the image so subsequent content is
	not joined to the image inline.

	Args:
		body:      Locator for the compose body element.
		pic:       Base64-encoded PNG string.
		wait_time: Milliseconds to wait after the paste event settles. Defaults to 500.
	"""
	log("Attaching image...")

	outlook = _outlook()
	outlook.keyboard.press(" ")
	outlook.keyboard.press("Enter")

	#* Inject the image by constructing a DataTransfer and firing a paste event —
	#* Outlook's compose window does not expose a native file-input for inline images
	body.evaluate(
		"""(el, base64) => {
			fetch(`data:image/png;base64,${base64}`)
					.then(res => res.blob())
					.then(blob => {
						const file = new File([blob], "pic.png", { type: "image/png" });
						const dt = new DataTransfer();
						dt.items.add(file);

						const pasteEvent = new ClipboardEvent("paste", {
							clipboardData: dt,
							bubbles: true
						});

						el.focus();
						el.dispatchEvent(pasteEvent);
					});
		}""",
		pic,
	)

	outlook.wait_for_timeout(wait_time)
	outlook.keyboard.press("Enter")


def bold(text: str, wait_time: int = 0) -> None:
	"""
	Type text in bold inside the active Outlook compose window.

	Wraps the text in Ctrl+B toggles and appends a trailing space so the
	cursor exits bold mode before the next character is typed.

	Args:
		text:      Text to type in bold.
		wait_time: Milliseconds to wait after typing. Defaults to 0.
	"""
	outlook = _outlook()
	outlook.keyboard.press("Control+B")
	outlook.keyboard.type(text)
	outlook.keyboard.press(" ")
	outlook.keyboard.press("Control+B")
	outlook.keyboard.press(" ")
	outlook.wait_for_timeout(wait_time)


def add_emails(field: Locator, emails: list[str], is_to: bool = True) -> None:
	"""
	Type and confirm each address into an Outlook recipient field.

	Each address is terminated with a semicolon so Outlook resolves it
	before the next one is typed.

	Args:
		field:   Locator for the To/CC input element.
		emails:  List of email address strings.
		is_to:   When True, raises NoRecipientsError if emails is empty.

	Raises:
		NoRecipientsError: If is_to is True and emails is empty.
	"""
	if not emails and is_to:
		raise NoRecipientsError()

	outlook = _outlook()
	field.click()

	for email in emails:
		outlook.keyboard.type(email)
		outlook.wait_for_timeout(500)
		outlook.keyboard.press(";")
		outlook.wait_for_timeout(500)


def check_bullets(cmd: str, is_bulleted: bool) -> bool:
	"""
	Synchronise Outlook's bullet state with whether the current command needs bullets.

	Pressing Ctrl+. toggles bullets; Ctrl+Space clears inline formatting first.
	Both transitions repeat the shortcut three times to account for Outlook's
	occasional input lag.

	Args:
		cmd:         Lowercased command string from the mail structure row.
		is_bulleted: Current bullet state.

	Returns:
		Updated bullet state.
	"""
	outlook = _outlook()
	_REPEAT = 3

	if "bullet" in cmd and not is_bulleted:
		for _ in range(_REPEAT):
			outlook.keyboard.press("Control+Space")
			outlook.wait_for_timeout(500)
		outlook.keyboard.press("Control+.")
		return True

	if "bullet" not in cmd and is_bulleted:
		for _ in range(_REPEAT):
			outlook.keyboard.press("Control+.")
			outlook.wait_for_timeout(500)
		return False

	return is_bulleted


def insert_url(label: str, url: str) -> None:
	"""
	Open Outlook's Insert Hyperlink dialog and fill in url and label.

	The dialog is keyboard-driven: Ctrl+K opens it, Tab/Shift+Tab move
	between the URL and display-text fields.

	Args:
		label: Display text shown in the email body.
		url:   Destination URL.
	"""
	outlook = _outlook()

	outlook.keyboard.press("Control+K")
	outlook.wait_for_timeout(500)

	#* The dialog focuses the URL field first
	outlook.keyboard.insert_text(url)

	#* Navigate to the display-text field, clear it, and type the label
	outlook.keyboard.press("Shift+Tab")
	outlook.keyboard.press("Control+A")
	outlook.keyboard.press("Backspace")
	outlook.keyboard.insert_text(label)

	outlook.keyboard.press("Enter")


#* ---------------------------------------------------------------------------
#* Visual resolution
#* ---------------------------------------------------------------------------

def _resolve_visual(visual_id: str, pics: list[str]) -> str:
	"""
	Resolve a visual ID from the mail structure to a base64-encoded PNG string.

	Resolution order:
	1. MEDIA dict  — pre-captured named visuals (e.g. "we mau").
	2. MEDIA_LABELS — index into the pics list captured for this team.
	3. "incentives" alias — the shared incentives slide image.

	Args:
		visual_id: Raw ID string parsed from the command column.
		pics:      Ordered list of base64 PNG strings for this team's visuals.

	Raises:
		MissingVisualError: If the ID cannot be resolved by any of the above.
	"""
	if visual_id in MEDIA:
		return img_to_b64(MEDIA[visual_id])

	if visual_id in MEDIA_LABELS:
		return pics[MEDIA_LABELS[visual_id]]

	if visual_id == "incentives":
		return img_to_b64(INCENTIVES_SLIDE_PATH)

	raise MissingVisualError(visual_id)


#* ---------------------------------------------------------------------------
#* Mail composition
#* ---------------------------------------------------------------------------

def compose_mail(
	mail_structure: pd.DataFrame,
	pics: list[str] | None = None,
	collected_pics: dict[str, str] | None = None,
	to: list[str] | None = None,
	cc: list[str] | None = None,
) -> Locator:
	"""
	Drive the Outlook compose window according to a mail structure DataFrame.

	Each row in mail_structure carries a "Command" and a "Text" value.
	Supported commands (case-insensitive):

	- "subject"         — fills the subject field.
	- "break"           — inserts a blank line and exits bullet mode.
	- "line"            — types plain text; if "end" is also in the command,
								scrolls to the end first and waits for layout.
	- "title"           — types text in bold.
	- "bullet line"     — types text as a bulleted list item.
	- "visual: <id>"    — resolves and embeds a single image.
	- "visuals"         — embeds all collected_pics with bold captions.
	- "url: <label>"    — inserts a hyperlink using the text as the URL.

	Processing stops as soon as a row whose command contains "end" is handled.

	Args:
		mail_structure: DataFrame with at least "Command" and "Text" columns.
		pics:           Base64 PNG strings for this team's named visuals.
		collected_pics: Caption → base64 PNG mapping for batch visual insertion.
		to:             Recipient addresses for the To field.
		cc:             Recipient addresses for the CC field.

	Returns:
		Locator for the Send button, ready to be clicked.
	"""
	pics           = pics           or []
	collected_pics = collected_pics or {}
	to             = to             or []
	cc             = cc             or []

	log("Composing mail...", left_nl=1)

	outlook = _outlook()

	reader = wait_for(outlook, OUTLOOK_DOM, at=["reader"])[0]
	send   = wait_for(reader,  OUTLOOK_DOM, at=["send"])[0]

	#* Locate all interactive fields up front to avoid repeated DOM queries
	field_key = "field"
	fields: dict[str, Locator] = {
		"subject": wait_for(reader, OUTLOOK_DOM, at=["subject"])[0],
		"to":      wait_for(reader, OUTLOOK_DOM, at=[field_key], index={field_key: 0})[0],
		"cc":      wait_for(reader, OUTLOOK_DOM, at=[field_key], index={field_key: 1})[0],
		"body":    wait_for(reader, OUTLOOK_DOM, at=[field_key], index={field_key: 2})[0],
	}

	add_emails(fields["to"], to)
	add_emails(fields["cc"], cc, is_to=False)

	fields["body"].click()
	is_bulleted = False

	for _, row in mail_structure.iterrows():
		cmd_original: str = row["Command"]
		cmd:          str = cmd_original.lower()
		text:         str = row["Text"]
		is_end:       bool = "end" in cmd

		#* Subject is filled once and control returns to the body immediately
		if "subject" in cmd:
			fields["subject"].click()
			outlook.keyboard.insert_text(text)
			fields["body"].click()
			continue

		outlook.keyboard.press("Control+End")
		outlook.wait_for_timeout(500)

		is_bulleted = check_bullets(cmd, is_bulleted)

		if "break" in cmd:
			outlook.keyboard.press("Enter")
			is_bulleted = False

		if "line" in cmd:
			if is_end:
				#* Scroll to the very end and give Outlook time to reflow before typing
				outlook.wait_for_timeout(1000)
				outlook.keyboard.press("Control+End")
			outlook.keyboard.insert_text(text)

		elif "title" in cmd:
			if is_bulleted:
				#* Exit bullet mode before typing a heading
				outlook.keyboard.press("Control+.")
				is_bulleted = False
			bold(text)

		elif "visuals" in cmd:
			#* Batch-insert all collected visuals with bold captions;
			#* separate groups that share a common ID prefix with extra line breaks
			prev_group = ""
			for title, pic in collected_pics.items():
				group, _ = title.split(" - ", maxsplit=1)

				outlook.wait_for_timeout(500)
				if prev_group and group != prev_group:
					for _ in range(3):
						outlook.keyboard.press("Enter")

				bold(title, wait_time=1000)
				mail_attach_image(fields["body"], pic, wait_time=1000)
				outlook.keyboard.press("Enter")
				prev_group = group

		elif "visual" in cmd:
			#* Single visual — parse the ID after the colon, resolve, and embed
			_, visual_id = cmd.split(": ", maxsplit=1)
			pic = _resolve_visual(visual_id.strip(), pics)
			bold(text)
			mail_attach_image(fields["body"], pic)

		elif "url" in cmd:
			#* The label comes from the original (un-lowercased) command
			_, label = cmd_original.split(": ", maxsplit=1)
			insert_url(label, text)

		if is_end:
			break

		outlook.keyboard.press("Control+End")
		outlook.wait_for_timeout(500)
		outlook.keyboard.press("Enter")

	return send


#* ---------------------------------------------------------------------------
#* Mail dispatch
#* ---------------------------------------------------------------------------

def send_new_mail(
	to: str,
	cc: str,
	names: str,
	role: str,
	structure: pd.DataFrame,
	pics: list[IMG] | None = None,
	collected_pics: dict[str, str] | None = None,
) -> None:
	"""
	Open a new Outlook compose window, fill it, and send it.

	Splits space-separated to/cc strings into lists, converts IMG objects to
	base64, composes the mail, and clicks Send.

	Args:
		to:             Space-separated recipient addresses.
		cc:             Space-separated CC addresses.
		names:          Display name(s) used in the sent-confirmation log line.
		role:           V-team role, recorded in the mail log for dedup checks.
		structure:      Mail structure DataFrame passed to compose_mail.
		pics:           IMG objects to embed as named visuals. Defaults to [].
		collected_pics: Caption → base64 map for batch visuals. Defaults to {}.
	"""
	pics           = pics           or []
	collected_pics = collected_pics or {}

	log("Sending mail...", left_nl=1)
	log(f"To: {to}")
	log(f"CC: {cc}")

	outlook = _outlook()
	outlook.bring_to_front()

	wait_for(outlook, OUTLOOK_DOM, timeout=10000, at=["new mail"])[0].click()

	send = compose_mail(
		to=to.split(),
		cc=cc.split(),
		pics=imgs_to_b64(pics),
		mail_structure=structure,
		collected_pics=collected_pics,
	)

	click_and_wait(send, outlook, timeout=2000)

	log(f"Mail sent to {names}.")
	log(mail_log(to, cc, role), type=MAIL_LOGS, silent=True)
