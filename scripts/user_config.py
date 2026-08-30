"""
Per-user configuration for the Find Connections (Basic) skill.

The config lives **inside** the skill package, at ``<skill-root>/config.json``,
and is git-ignored (the ``config.json`` rule in ``.gitignore``). Keeping it in the
package folder means it stays self-contained — uninstall by deleting the skill
directory, with nothing left behind in ``~/.claude`` — while cloning/pushing the
repo still never ships anyone's personal data (git honours the ignore rule). Each
user creates their own config at first run (via the ``--setup`` wizard); the
package ships only ``config.example.json`` as a template.

Schema (JSON):
    {
      "excel_dir":              "<folder holding the spreadsheets>",
      "excel_filename":         "<default .xlsx to update>",
      "relevant_roles":         ["Product Manager", "VP Product", ...],
      "relevant_industries":    ["Fintech", "Analytics", ...],
      "relevant_role_patterns": ["\\bcpo\\b", ...]   # optional, advanced
    }

``relevant_roles`` are plain job titles — matched case-insensitively with word
boundaries. ``relevant_role_patterns`` is optional and holds raw regex patterns
for power users (and for the one-time migration of the original hardcoded
pattern library); most users never touch it.
"""

import json
import os
import re
from pathlib import Path
from typing import List, Optional


class ConfigNotFoundError(FileNotFoundError):
    """Raised when no per-user config exists yet (first run)."""


def title_to_pattern(title: str) -> str:
    """Turn a plain job title into a word-boundary-anchored, case-insensitive regex."""
    return r"\b" + re.escape(title.strip()) + r"\b"


class UserConfig:
    """Loads/saves the per-user config and derives role-matching patterns."""

    def __init__(
        self,
        excel_dir: str,
        excel_filename: str,
        relevant_roles: Optional[List[str]] = None,
        relevant_industries: Optional[List[str]] = None,
        relevant_role_patterns: Optional[List[str]] = None,
    ):
        self._excel_dir = excel_dir
        self.excel_filename = excel_filename
        self.relevant_roles = relevant_roles or []
        self.relevant_industries = relevant_industries or []
        self.relevant_role_patterns = relevant_role_patterns or []

    @property
    def excel_dir(self) -> str:
        """
        ``excel_dir`` with ``~`` and environment variables resolved.

        Held unexpanded in ``_excel_dir`` so it survives a ``save()`` round-trip,
        and expanded only on read. That lets the config say ``~/admin/07_PM``
        and resolve correctly under any username — the same portability rule the
        skill already applies to its own location via ``Path(__file__).resolve()``.
        Plain absolute paths pass through untouched.
        """
        return os.path.expandvars(os.path.expanduser(self._excel_dir))

    @excel_dir.setter
    def excel_dir(self, value: str) -> None:
        """Store the raw value; expansion happens on read, never on write."""
        self._excel_dir = value

    # ------------------------------------------------------------------ #
    # Locations                                                           #
    # ------------------------------------------------------------------ #

    @staticmethod
    def config_dir() -> Path:
        """
        The skill package root — this file is at ``<root>/scripts/user_config.py``.

        The per-user config lives **inside** the package (git-ignored via the
        ``config.json`` rule in ``.gitignore``), so it stays contained in the
        skill folder: uninstalling is a single delete of the skill directory with
        nothing left behind in ``~/.claude``, and cloning/pushing the repo never
        carries anyone's config (git honours the ignore rule).
        """
        return Path(__file__).resolve().parent.parent

    @classmethod
    def config_path(cls) -> Path:
        return cls.config_dir() / "config.json"

    @classmethod
    def exists(cls) -> bool:
        return cls.config_path().exists()

    # ------------------------------------------------------------------ #
    # Load / save                                                         #
    # ------------------------------------------------------------------ #

    @classmethod
    def load(cls) -> "UserConfig":
        """Load the config, or raise ConfigNotFoundError if it doesn't exist yet."""
        path = cls.config_path()
        if not path.exists():
            raise ConfigNotFoundError(
                f"No config found at {path}. Run first-run setup to create it."
            )
        data = json.loads(path.read_text(encoding="utf-8"))
        missing = [k for k in ("excel_dir", "excel_filename") if not data.get(k)]
        if missing:
            raise ValueError(f"Config {path} is missing required keys: {missing}")
        return cls(
            excel_dir=data["excel_dir"],
            excel_filename=data["excel_filename"],
            relevant_roles=data.get("relevant_roles", []),
            relevant_industries=data.get("relevant_industries", []),
            relevant_role_patterns=data.get("relevant_role_patterns", []),
        )

    def save(self) -> Path:
        """Persist the config, creating the directory if needed. Returns its path."""
        path = self.config_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            # Unexpanded, so re-saving a portable '~/...' config keeps it portable.
            "excel_dir": self._excel_dir,
            "excel_filename": self.excel_filename,
            "relevant_roles": self.relevant_roles,
            "relevant_industries": self.relevant_industries,
            "relevant_role_patterns": self.relevant_role_patterns,
        }
        path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        return path

    # ------------------------------------------------------------------ #
    # Derived values                                                      #
    # ------------------------------------------------------------------ #

    def role_patterns(self) -> List[str]:
        """
        The full regex-pattern list the RoleMatcher compiles: each plain title
        anchored with word boundaries, plus any raw advanced patterns verbatim.
        """
        return [title_to_pattern(t) for t in self.relevant_roles] + list(
            self.relevant_role_patterns
        )

    def get_relevant_industries(self) -> List[str]:
        return list(self.relevant_industries)
