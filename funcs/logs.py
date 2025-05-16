from vars.exports import LOGS, MAIL_LOGS, TEST, os
from .imports import dt


def clear_console() -> None:
    """
    Clears the terminal or command prompt screen.
    This function detects the operating system and executes the appropriate command
    to clear the console window. On Windows, it uses 'cls'; on Unix-like systems, it uses 'clear'.
    """

    os.system('cls' if os.name == 'nt' else 'clear')

def read_file_as_str(path: str) -> str:
    """
    Reads the contents of a file and returns it as a string.
    Args:
        path (str): The path to the file to be read.
    Returns:
        str: The contents of the file as a string.
    Raises:
        FileNotFoundError: If the file does not exist at the specified path.
        IOError: If an I/O error occurs while reading the file.
    """
   
    with open(path, 'r', encoding='utf-8') as file:
        return file.read()
    
def write_to_file(path:str, content:str) -> None:
    """
    Writes the given content to a file at the specified path.
    Parameters:
        path (str): The file path where the content will be written.
        content (str): The content to write to the file.
    Returns:
        None
    """

    with open(path, 'w') as file:
        file.write(content + "\n")

def delete_file(path:str) -> None:
    """
    Deletes the specified file if it exists and is within the project's root folder.
    Args:
        path (str): The relative or absolute path to the file to be deleted.
    Raises:
        ValueError: If the specified file is outside the project's root folder.
    Notes:
        - Prints a confirmation message if the file is deleted.
        - Prints a message if the file does not exist.
    """

    root_folder = os.path.abspath(os.getcwd())
    file_path = os.path.abspath(path)

    if not file_path.startswith(root_folder):
        raise ValueError(f"Error: The file '{path}' is outside the project's root folder")

    try:
        os.remove(file_path)
        print(f"File '{file_path}' deleted")
    except FileNotFoundError:
        print(f"Successful run, file '{file_path}' was not created")

def logs_to_set(text:str) -> set:
    """
    Converts a multiline log string into a set of log messages.
    Each line in the input text is expected to contain a delimiter " - ".
    The function extracts the part after the first occurrence of " - " from each non-empty line,
    strips leading and trailing whitespace, and adds it to a set to ensure uniqueness.
    Args:
        text (str): Multiline string containing log entries.
    Returns:
        set: A set of unique log message strings extracted from the input.
    Raises:
        ValueError: If the input string is empty or contains only whitespace.
    """

    if not text.strip():
        raise ValueError("Error: The string is empty")

    content_set = {line.split(" - ")[1].strip() for line in text.splitlines() if line.strip() and " - " in line}
    
    return content_set

def logs_file_to_set(path:str) -> set:
    """
    Reads a log file and extracts the second field from each line (split by ' - ') into a set.
    Args:
        path (str): The path to the log file.
    Returns:
        set: A set containing the second field from each valid log line.
    Raises:
        FileNotFoundError: If the specified file does not exist.
    """

    try:
        with open(path, 'r') as file:
            content_set = {line.split(" - ")[1].strip() for line in file.readlines() if line.strip() and " - " in line}
        
        return content_set
    except FileNotFoundError:
        raise FileNotFoundError(f"Error: The file {path} does not exist")

def log(message:str="", type=LOGS, write=False, silent=False, left_nl=0, right_nl=1, from_last=False) -> None:
    """
    Logs a message with timestamp to a specified log type, optionally writing to file and controlling output formatting.
    Args:
        message (str): The message to log. Defaults to an empty string.
        type (dict): The log type dictionary, typically containing 'path' and 'content' keys. Defaults to LOGS.
        write (bool): If True, writes the log content to the file specified in type['path']. Defaults to False.
        silent (bool): If True, suppresses printing the message to stdout. Defaults to False.
        left_nl (int): Number of newlines to prepend before each log entry. Defaults to 0.
        right_nl (int): Number of newlines to append after each log entry. Defaults to 1.
        from_last (bool): If True, prepends existing log file content to the current log content. Defaults to False.
    Returns:
        None
    """

    left_nl = "\n" * left_nl
    right_nl = "\n" * right_nl
    time = dt.now().strftime("%Y-%m-%d %H:%M:%S")

    if from_last and os.path.exists(type["path"]):
        type["content"] = read_file_as_str(type["path"]) + type["content"]

    if message.strip():
        for line in message.splitlines():
            if line.strip():
                type["content"] += f"{left_nl}{time} - {line}{right_nl}"
        
        if not silent:
            print(f"{left_nl}{message}{right_nl}", end="")

    if write:
        logs = logs_to_set(type["content"])
        logs.discard("START")
        logs.discard("END")

        if len(logs) > 0:
            write_to_file(type["path"], type["content"])

def mail_log(to:str, cc:str, role:str) -> str:
    """
    Generates a formatted log message for an email sent to a V-team member.
    Args:
        to (str): The recipient's email address.
        cc (str): The CC'd email addresses.
        role (str): The role of the V-team member.
    Returns:
        str: A formatted string describing the sent email log.
    """

    return f"Sent => V-team's role: {role}, To: {to}, CC: {cc}".strip()

def mail_log_exists(to:str, cc:str, role:str, mail_details="", is_draft=False) -> bool:
    """
    Checks if a mail log entry exists in the mail logs.
    Args:
        to (str): The recipient email address.
        cc (str): The CC email address.
        role (str): The role associated with the mail.
        mail_details (str, optional): Additional details about the mail. Defaults to "".
        is_draft (bool, optional): Whether the mail is a draft. Defaults to False.
    Returns:
        bool: True if the mail log entry exists, False otherwise.
    """

    logs = MAIL_LOGS["content"]
    log = mail_log(
        to, 
        cc, 
        role, 
        mail_details, 
        is_draft
    ) if is_draft else mail_log(
        to, 
        cc, 
        role 
    )

    return logs and log in logs_to_set(logs)

def choose_keep_logs() -> bool:
    """
    Prompts the user to choose whether to keep or discard logs.
    Displays a menu with options to keep or discard logs, and repeatedly asks for user input until a valid choice is made.
    Logs the selected action and returns True if "Keep logs" is chosen, or False if "Discard logs" is chosen.
    Returns:
        bool: True if the user chooses to keep logs, False if the user chooses to discard logs.
    """
    
    log("Choose action:")
    log("1. Keep logs")
    log("2. Discard logs")

    modes = {
        "1": "Keep logs",
        "2": "Discard logs",
    }

    while True:
        choice = input("Enter your choice (1, 2 or 3): ").strip()
        if choice in modes:
            log(f"{modes[choice]} selected", left_nl=1)
            return choice == "1"
        else:
            print("Invalid choice. Please enter 1, 2 or 3.")

def adjust_logs() -> None:
    """
    Adjusts the log file paths for testing and logs recent entries.
    If the global TEST dictionary has "active" set to True, this function prefixes the log file paths
    in the MAIL_LOGS and LOGS dictionaries with "test_". It then calls the `log` function twice:
    first to log recent entries from the default log, and second to log recent entries from the mail logs.
    Returns:
        None
    """

    if TEST["active"]:
        MAIL_LOGS["path"] = f"test_{MAIL_LOGS['path']}"
        LOGS["path"] = f"test_{LOGS['path']}"

    log(from_last=True)
    log(type=MAIL_LOGS, from_last=True)
