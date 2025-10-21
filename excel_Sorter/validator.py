"""Sheet name validation helpers."""
import re
from typing import List

INVALID_CHARS_RE = re.compile(r'[:\\/?*\[\]]')  # Excel forbids these

def has_invalid_chars(name: str) -> bool:
    """Return True if name contains Excel-invalid characters."""
    return bool(INVALID_CHARS_RE.search(name))

def is_too_long(name: str, limit: int = 31) -> bool:
    """Excel sheet name max length is 31 characters."""
    return len(name) > limit

def find_duplicates(names: List[str]) -> List[str]:
    """Return list of duplicate names (unique in result)."""
    seen = {}
    dup = []
    for n in names:
        seen[n] = seen.get(n, 0) + 1
    for n, c in seen.items():
        if c > 1:
            dup.append(n)
    return dup
