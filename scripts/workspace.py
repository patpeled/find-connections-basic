"""
Ephemeral work directory for the Find Connections (Basic) skill.

All transient run artifacts — the ``{url: page_text}`` response cache and the
pre-save workbook backups — live here, *outside* the user's data folder, so a
successful run leaves the spreadsheet folder clean. Everything written here is
disposable: cache is deleted once a run applies cleanly, and each backup is
deleted the moment its save is confirmed.
"""

import os
import tempfile

# A single stable subfolder under the OS temp dir keeps run artifacts together
# and out of the data folder. Regenerated on demand; safe to wipe anytime.
_WORK_SUBDIR = "find-connections-basic"


def work_dir() -> str:
    """Return the skill's temp work directory, creating it if needed."""
    path = os.path.join(tempfile.gettempdir(), _WORK_SUBDIR)
    os.makedirs(path, exist_ok=True)
    return path


def default_cache_path() -> str:
    """Default location for the response cache JSON (in the work dir)."""
    return os.path.join(work_dir(), "cache.json")
