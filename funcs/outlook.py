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
    pd
)


def mail_attach_image(body:Locator, pic:str) -> None:
    """
    Attaches an image to an Outlook email body by simulating a paste event with a base64-encoded image.
    Args:
        body (Locator): The locator for the email body element where the image should be attached.
        pic (str): The base64-encoded string of the image to be attached.
    Returns:
        None
    Side Effects:
        - Simulates keyboard and clipboard events to paste the image into the email body.
        - Waits for a short timeout after attaching the image.
    """

    log("Attaching image...")

    outlook = PAGES["outlook"]
    outlook.keyboard.press(" ")
    outlook.keyboard.press("Enter")

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
        pic
    )
    outlook.wait_for_timeout(500)
    outlook.keyboard.press("Enter")

def bold(text:str) -> None:
    """
    Types the given text in bold in the Outlook application using keyboard shortcuts.
    Args:
        text (str): The text to be typed in bold.
    Returns:
        None
    """

    outlook = PAGES["outlook"]

    outlook.keyboard.press("Control+B")
    outlook.keyboard.type(f"{text}")
    outlook.keyboard.press(" ")
    outlook.keyboard.press("Control+B")
    outlook.keyboard.press(" ")

def add_emails(input:Locator, emails:list[str], is_to=True) -> None:
    """
    Adds a list of email addresses to an Outlook input field.
    Args:
        input (Locator): The locator for the input field where emails will be entered.
        emails (list[str]): A list of email addresses to add.
        is_to (bool, optional): If True, raises an exception when no emails are provided. Defaults to True.
    Raises:
        Exception: If no emails are provided and is_to is True.
    Returns:
        None
    """

    if not emails and is_to:
        raise Exception("No emails provided")

    outlook = PAGES["outlook"]

    input.click()
    for email in emails:
        outlook.keyboard.type(email)
        outlook.wait_for_timeout(500)
        outlook.keyboard.press(";")
        outlook.wait_for_timeout(500)

def check_bullets(cmd:str, is_bulleted:bool) -> bool:
    """
    Toggles bullet formatting in Outlook based on the command and current bullet state.
    Args:
        cmd (str): The command string to check for bullet-related instructions.
        is_bulleted (bool): The current bullet state (True if bulleted, False otherwise).
    Returns:
        bool: The updated bullet state after processing the command.
    Behavior:
        - If "bullet" is in the command and bullets are not currently enabled, enables bullets by simulating keyboard shortcuts.
        - If "bullet" is not in the command and bullets are currently enabled, disables bullets by simulating keyboard shortcuts.
        - Uses a fixed number of keyboard shortcut repetitions to ensure the desired bullet state.
    """

    outlook = PAGES["outlook"]

    n = 3
    if "bullet" in cmd and not is_bulleted:
        is_bulleted = True
        for i in range(n):
            outlook.keyboard.press("Control+Space")
            outlook.wait_for_timeout(500)
        outlook.keyboard.press("Control+.")
    elif "bullet" not in cmd and is_bulleted:
        is_bulleted = False
        for _ in range(n):
            outlook.keyboard.press("Control+.")
            outlook.wait_for_timeout(500)

    return is_bulleted

def insert_url(label:str, url:str) -> None:
    """
    Inserts a hyperlink into the Outlook compose window using keyboard automation.
    Args:
        label (str): The display text for the hyperlink.
        url (str): The URL to be inserted as the hyperlink.
    Returns:
        None
    """

    outlook = PAGES["outlook"]

    outlook.keyboard.press("Control+K")
    outlook.wait_for_timeout(500)

    outlook.keyboard.insert_text(url)

    outlook.keyboard.press("Shift+Tab")
    outlook.keyboard.press("Control+A")
    outlook.keyboard.press("Backspace")
    outlook.keyboard.insert_text(label)
    
    outlook.keyboard.press("Enter")

def compose_mail(mail_structure:pd.DataFrame, pics:list[str]=[], collected_pics:dict[str, str]={}, to:list[str]=[], cc:list[str]=[]) -> Locator:
    """
    Composes an email in Outlook using the provided mail structure and attachments.
    Args:
        mail_structure (pd.DataFrame): A DataFrame containing the structure of the email. Each row should specify a command and associated text.
        pics (list[str], optional): A list of image paths or base64 strings to be used for visual attachments. Defaults to an empty list.
        collected_pics (dict[str, str], optional): A dictionary mapping visual titles to image paths or base64 strings for batch visual insertion. Defaults to an empty dict.
        to (list[str], optional): List of recipient email addresses for the "To" field. Defaults to an empty list.
        cc (list[str], optional): List of recipient email addresses for the "CC" field. Defaults to an empty list.
    Returns:
        Locator: The locator for the "Send" button in the Outlook compose window.
    Notes:
        - The function interacts with the Outlook UI to fill in the subject, recipients, body, and attachments as specified in the mail_structure.
        - Supports commands for inserting subject, bulleted lists, titles, images, URLs, and line breaks.
        - Handles both individual and batch visual attachments.
        - Assumes the presence of global variables and helper functions such as PAGES, OUTLOOK_DOM, wait_for, add_emails, check_bullets, bold, mail_attach_image, img_to_b64, MEDIA, MEDIA_LABELS, INCENTIVES_SLIDE_PATH, and insert_url.
    """

    log("Composing mail...", left_nl=1)

    outlook = PAGES["outlook"]

    reader = wait_for(
        outlook, 
        OUTLOOK_DOM,
        at=["reader"],
    )[0]

    send = wait_for(
        reader, 
        OUTLOOK_DOM, 
        at=["send"]
    )[0]

    id = "field"
    fields: dict[str, Locator|None] = {
        "subject": wait_for(reader, OUTLOOK_DOM, at=["subject"])[0],
        "to": wait_for(reader, OUTLOOK_DOM, at=[id], index={id: 0})[0],
        "cc": wait_for(reader, OUTLOOK_DOM, at=[id], index={id: 1})[0],
        "body": wait_for(reader, OUTLOOK_DOM, at=[id], index={id: 2})[0]
    }

    add_emails(fields["to"], to)
    add_emails(fields["cc"], cc, is_to=False)

    fields["body"].click()
    is_bulleted = False

    for _, row in mail_structure.iterrows():
        cmd_original = row["Command"]
        cmd = cmd_original.lower()
        is_end = "end" in cmd
        text = row["Text"]

        if "subject" in cmd:
            fields["subject"].click()
            outlook.keyboard.insert_text(f"{text}")
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
                outlook.wait_for_timeout(1000)
                outlook.keyboard.press("Control+End")
            outlook.keyboard.insert_text(f"{text}")
        elif "title" in cmd:
            if is_bulleted:
                outlook.keyboard.press("Control+.")
                is_bulleted = False
            bold(f"{text}")
        elif "visuals" in cmd:
            prev_id = ""
            for title, pic in collected_pics.items():
                id, _ = title.split(" - ")
                if prev_id and id != prev_id:
                    outlook.keyboard.press("Enter")
                    outlook.keyboard.press("Enter")
                    outlook.keyboard.press("Enter")

                bold(f"{title}")
                mail_attach_image(fields["body"], pic)
                outlook.keyboard.press("Enter")
                outlook.wait_for_timeout(500)
                prev_id = id
        elif "visual" in cmd:
            _, id = cmd.split(": ")
            pic = None

            if id in MEDIA:
                pic = img_to_b64(MEDIA[id])
            elif id in MEDIA_LABELS:
                pic = pics[MEDIA_LABELS[id]]
            elif id == "incentives":
                pic = img_to_b64(INCENTIVES_SLIDE_PATH)

            bold(f"{text}")
            mail_attach_image(fields["body"], pic)
        elif "url" in cmd:
            _, label = cmd_original.split(": ")
            insert_url(label, text)

        if is_end: break

        outlook.keyboard.press("Control+End")
        outlook.wait_for_timeout(500)
        outlook.keyboard.press("Enter")
    
    return send

def send_new_mail(to:str, cc:str, names:str, role:str, structure:pd.DataFrame, pics:list[IMG]=[], collected_pics:dict[str, str]={}) -> None:
    """
    Sends a new email using the Outlook application interface.
    Args:
        to (str): Space-separated string of recipient email addresses.
        cc (str): Space-separated string of CC email addresses.
        names (str): Names of the primary recipients for logging purposes.
        role (str): Role or context of the email for logging.
        structure (pd.DataFrame): DataFrame representing the structure or content of the email.
        pics (list[IMG], optional): List of image objects to attach or embed in the email. Defaults to an empty list.
        collected_pics (dict[str, str], optional): Dictionary mapping image identifiers to base64-encoded image strings. Defaults to an empty dictionary.
    Returns:
        None
    """

    log(f"Sending mail...", left_nl=1)
    log(f"To: {to}")
    log(f"CC: {cc}")

    outlook = PAGES["outlook"]
    outlook.bring_to_front()
    
    wait_for(
        outlook, 
        OUTLOOK_DOM,
        timeout=10000,
        at=["new mail"],
    )[0].click()

    send = compose_mail(
        to=to.split(" "),
        cc=cc.split(" "),
        pics=imgs_to_b64(pics),
        mail_structure=structure,
        collected_pics=collected_pics
    )

    click_and_wait(send, outlook, timeout=2000)

    log(f"Mail sent to {names}")
    log(
        mail_log(to, cc, role),
        type=MAIL_LOGS,
        silent=True
    )
