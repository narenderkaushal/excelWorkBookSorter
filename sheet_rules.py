"""Sheet sorting and renaming helpers."""
import re
from typing import Callable, List, Tuple

# Mapping of common month representations to calendar index (1–12).
MONTH_NAME_MAP = {
    "jan": 1, "january": 1, "JANUARY": 1,
    "feb": 2, "february": 2, "FEBRUARY": 2,
    "mar": 3, "march": 3, "MARCH": 3,
    "apr": 4, "april": 4, "APRIL": 4,
    "may": 5, "MAY": 5,
    "jun": 6, "june": 6, "JUNE": 6,
    "jul": 7, "july": 7, "JULY": 7,
    "aug": 8,"august": 8, "AUGUST": 8,
    "sep": 9, "sept": 9, "september": 9, "SEPTEMBER": 9,
    "oct": 10, "october": 10, "OCTOBER": 10,
    "nov": 11, "november": 11, "NOVEMBER": 11,
    "dec": 12, "december": 12, "DECEMBER": 12,
}

def alpha_key(ws) -> str:
    """Alphabetical key (case-insensitive)."""
    return ws.title.lower()

def numeric_suffix_key(ws) -> tuple:
    """Sort by numeric suffix if present, else alpha."""
    title = ws.title
    match = re.match(r"(.*?)(\d+)$", title)
    if match:
        prefix, num = match.groups()
        print(f"[DEBUG] Numeric sort key for {title}: ({prefix.lower()}, {int(num)})")
        return (prefix.lower(), int(num))
    print(f"[DEBUG] No numeric match for {title}")
    return (title.lower(), float("inf"))

def _normalize_month_token(title: str) -> str:
    """Extract a candidate month token from a sheet title.
    The function normalizes the title to lowercase and then takes the
    first token split by space, underscore, or hyphen. For example:
    - "Jan" -> "jan", - "January_Sales" -> "january"
    - "03-Mar" -> "03-mar" (no direct month match, handled by caller)"""
    token = title.strip().lower()
    for separator in (" ", "_", "-"):
        if separator in token:
            token = token.split(separator, maxsplit=1)[0]
            break
    return token

def _month_rank(title: str) -> Tuple[int, str]:
    """Compute a stable sort rank for month-like sheet names.
    Returns a tuple (rank, normalized_title):
    - rank 1–12 for recognized months (Jan–Dec).
    - rank 13 for non-month sheets so they appear after month sheets."""
    normalized_title = title.lower()
    token = _normalize_month_token(title)
    month_index = MONTH_NAME_MAP.get(token)
    if month_index is None:
        # Non-month sheets follow month sheets, ordered alphabetically.
        return 13, normalized_title
    return month_index, normalized_title

def contains_month_sheets(sheet_names: List[str]) -> bool:
    """Determine whether any sheet title corresponds to a recognized month.
    Normalizes each sheet title and checks whether the leading token 
    matches any key in MONTH_NAME_MAP.
    Returns True if at least one month sheet is detected."""
    for title in sheet_names:
        token = _normalize_month_token(title)
        if token in MONTH_NAME_MAP:
            return True
    return False

def month_order_key(ws) -> Tuple[int, str]:
    """Sort key for calendar month order (Jan–Dec).
    Sheet names that represent months (e.g. 'Jan', 'March', 'sep')
    are ordered by calendar index. Non-month sheets are ordered
    after month sheets, alphabetically."""
    return _month_rank(ws.title)


def month_order_desc_key(ws) -> Tuple[int, str]:
    """Sort key for reverse calendar month order (Dec–Jan).
    Month sheets are reversed (Dec first, then Nov, ... Jan).
    Non-month sheets remain after month sheets and are kept in
    alphabetical order relative to each other."""
    rank, normalized_title = _month_rank(ws.title)
    if rank == 13:
        # Non-month sheets keep their relative (alphabetical) ordering
        # and remain after all month sheets.
        return rank, normalized_title

    # Invert 1–12 so that December (12) becomes smallest rank.
    reversed_rank = 13 - rank
    return reversed_rank, normalized_title

def regex_order_key(pattern: str) -> Callable:
    """Return a key function that matches regex groups for ordering."""
    prog = re.compile(pattern)

    def key(ws):
        m = prog.search(ws.title)
        if not m:
            return ("", ws.title.lower())
        # use first group then title
        return (m.group(1), ws.title.lower())
    return key

def apply_template(title: str, template: str, index: int = None) -> str:
    """
    Apply a simple template for renaming.
    Supported tokens: {title}, {i}, {index}
    """
    out = template.replace("{title}", title)
    if index is not None:
        out = out.replace("{i}", str(index)).replace("{index}", str(index))
    return out
