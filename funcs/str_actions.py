from .imports import uni


def str_normalize(text:str) -> str:
    """
    Normalizes a Unicode string to the NFKC (Normalization Form KC) form.
    Args:
        text (str): The input string to normalize.
    Returns:
        str: The normalized Unicode string in NFKC form.
    Raises:
        TypeError: If the input is not a string.
    Example:
        >>> str_normalize("ｅ́")
        'é'
    """

    return uni.normalize("NFKC", text)

def strip_edges(text:str, strip:str) -> str:
    """
    Removes the specified substring from the start and end of the given text, if present.
    Args:
        text (str): The input string to process.
        strip (str): The substring to remove from the start and end of the text.
    Returns:
        str: The resulting string with the specified substring removed from both edges, if present.
    """

    if text.startswith(strip):
        text = text[len(strip):]
    if text.endswith(strip):
        text = text[:-len(strip)]
        
    return text
