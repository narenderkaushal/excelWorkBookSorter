"""Sheet sorting and renaming helpers."""
import re
from typing import Callable, List

def alpha_key(ws) -> str:
    """Alphabetical key (case-insensitive)."""
    return ws.title.lower()

def numeric_suffix_key(ws) -> tuple:
    """Sort by numeric suffix if present, else alpha.
    m = re.search(r"(\d+)$", ws.title)
    if m:
        return (0, int(m.group(1)), ws.title.lower())
    return (1, ws.title.lower())"""
    import re
    title = ws.title
    match = re.match(r"(.*?)(\d+)$", title)
    if match:
        prefix, num = match.groups()
        return (prefix.lower(), int(num))
    # fallback for names without digits
    return (title.lower(), float("inf"))

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
