"""Simple backup utility for workbooks."""
import os
import shutil
from datetime import datetime

def make_backup(path: str) -> str:
    """Make a timestamped backup copy of 'path'.
    Returns the backup path or empty string on failure."""
    try:
        if not os.path.exists(path):
            return ""
        base = os.path.basename(path)
        dirn = os.path.dirname(path) or "."
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{base}.backup.{stamp}"
        backup_path = os.path.join(dirn, backup_name)
        # Use shutil.copy2 to preserve metadata
        shutil.copy2(path, backup_path)
        return backup_path
    except (OSError, shutil.Error):
        return ""
