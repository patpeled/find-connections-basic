# Find Connections (Basic) — Deployment & Integration Guide

## Overview

This guide covers installing the **find-connections-basic** skill on a new machine and
verifying the pieces are wired up. The skill researches who in the user's
LinkedIn network can introduce them to people at target companies and writes the
results into their **Company shortlist** spreadsheet.

`SKILL.md` is the authoritative runtime instruction document for how Claude drives
the PLAN → FETCH → APPLY flow; this guide is only about setup, security, and
troubleshooting.

## Project Structure

The skill is a self-contained package. All code lives in `scripts/`; no personal
data ships inside it.

```
📁 ~/.claude/skills/find-connections-basic/
├── SKILL.md                             # Authoritative runtime instructions
├── README.md                            # Install + first-run guide (shareable)
├── .gitignore                           # Keeps config.json / spreadsheets out of git
├── config.json                          # Per-user config (git-ignored; created at first run)
└── scripts/
    ├── linkedin_excel_contact_skill.py  # CLI entry point / orchestration
    ├── chrome_bridge.py                 # PLAN/FETCH/APPLY cache protocol
    ├── excel_utils.py                   # Excel I/O + SHEET_SCHEMA + creator
    ├── linkedin_utils.py                # RoleMatcher + research/parsing logic
    ├── url_parser.py                    # LinkedIn URL parsing + surface detection
    ├── user_config.py                   # Per-user config loader/writer
    ├── workspace.py                     # Ephemeral work/temp dir helper
    ├── config.example.json              # Config template (no personal data)
    ├── test_skill.py                    # Test suite
    └── DEPLOYMENT_GUIDE.md              # This file
```

Personal settings live in a **git-ignored** `config.json` **inside** the package,
so the repo can be shared without leaking anyone's data (git honours the ignore
rule) while nothing is scattered outside the skill folder:

```
~/.claude/skills/find-connections-basic/config.json   (git-ignored)
```

## Prerequisites

### Dependencies

```
openpyxl>=3.0.0    # Excel file handling
```

Install with:

```bash
pip install openpyxl
```

### User Environment

- **Python 3.8+** available on PATH. The launcher is OS-specific: use **`py`** on
  **Windows** and **`python3`** on **macOS/Linux** (`<PY>` below stands for
  whichever applies). Absolute paths likewise use backslashes on Windows and
  forward slashes on macOS/Linux.
- **Claude Code** with the **Claude-in-Chrome MCP** connected — the Python process
  cannot call the browser itself, so Claude drives Chrome during FETCH.
- **LinkedIn account** with an active session in the connected Chrome.
- **Per-user config** at `<skill-root>/config.json` (git-ignored; see below). The
  spreadsheet folder and filename are read from this config — nothing is
  hardcoded.

## First-Run Setup (per-user config)

The package ships only `scripts/config.example.json`. On first use there is no
`<skill-root>/config.json`, so any run prints a first-run message. Claude then
asks the user for the values in chat and runs the **first-run wizard** (`--setup`),
which writes the git-ignored `config.json` **and** creates a starter sheet in one
command (see below).

Config keys:

- `excel_dir` — folder holding the spreadsheets
- `excel_filename` — default `.xlsx` to update
- `relevant_roles` — plain job titles that count as a relevant contact (matched
  case-insensitively, word-boundary anchored)
- `relevant_industries` — industries of interest
- `relevant_role_patterns` *(optional, advanced)* — raw regexes for power users

`RoleMatcher` is built **entirely** from these config values — there are no
hardcoded built-in titles or patterns in the code.

## First-Run Sheet Creation

If the configured `excel_filename` doesn't exist yet, the skill can create it from
the shared schema (`excel_utils.SHEET_SCHEMA`) instead of asking the user to
hand-build the layout. On a genuine first run, the `--setup` wizard writes the
config **and** the sheet together. Claude gathers the config values + intake
conversationally, writes a setup JSON, then runs the CLI **by absolute path**
(avoid `cd scripts/ && <PY> …` — the compound command triggers an extra permission
prompt each run):

```
# Windows
py "<scripts-dir>\linkedin_excel_contact_skill.py" --setup <setup.json>
# macOS/Linux
python3 "<scripts-dir>/linkedin_excel_contact_skill.py" --setup <setup.json>
```

The sheet is named from the setup JSON's `excel_filename`.

*(If the config already exists and you only need a fresh sheet, use
`--init-sheet --intake <intake.json>` instead.)*

`SHEET_SCHEMA` reproduces the reference layout exactly (sheet **Company
shortlist**, blank column A, title at **B2**, header row **4** from column B,
contact profile URLs on **row 5**, company data from **row 6**, default status
`No successful contact so far`). Both intake lists may be empty (headers-only
sheet). The command refuses to overwrite an existing file unless `--overwrite`
is passed.

## Integration with Claude Code

### Skill discovery

The skill is discovered by Claude Code from its location under
`~/.claude/skills/find-connections-basic/`. `SKILL.md`'s frontmatter registers the
`/find-connections-basic` trigger and describes when to use it. No separate command
registration step is required — dropping the package in the skills folder is
enough.

### Claude-in-Chrome integration

FETCH is Claude's job, not the Python process's. Ensure:

1. The Chrome extension is installed and connected
   (`mcp__Claude_in_Chrome__list_connected_browsers`).
2. The user is logged into LinkedIn in that Chrome session.

Do **not** fall back to desktop computer-use for LinkedIn.

## Runtime Flow (PLAN → FETCH → APPLY)

Because the Python process can't call the browser, the run is split into three
phases that pass data through a single JSON cache:

1. **PLAN** — Python enumerates the LinkedIn URLs to fetch and prints them to
   stdout (no queue file). Each company contributes **one** `/company/<slug>/people/`
   URL, because that surface is exhausted by scrolling during FETCH, not by page
   numbers. Only the `/search/results/people/` keyword fallback (used when a
   company row has no LinkedIn URL) is paginated, capped by `--max-pages`.
2. **FETCH** — Claude drives Chrome to each URL. The people tab loads
   mutual-connection cards lazily in a "People you may know" block, so Claude must
   scroll and repeatedly click **"Show more results"** until an expansion adds no
   new cards (with a hard cap), then capture the fully-expanded page text. Results
   are stored as `{url: page_text}` in the ephemeral work dir
   (`<OS temp>/find-connections-basic/cache.json`), never the data folder.
3. **APPLY** — Python serves every fetch from the cache, parses each company's page
   once, filters it for **every** requested contact, writes the column(s)/row, and
   deletes the cache on a clean apply.

**Batching:** pass `--contacts <path.json>` (a list of `{"name","url"}`) to add
several contact columns in one run. Each company's page is fetched once and
filtered for all contacts, so N introducers cost the same LinkedIn traffic as one.

## Security Considerations

### Credential Handling ✓ SECURE

- **No credentials stored** in skill code.
- **No passwords transmitted** — uses the active browser session only.
- **LinkedIn authentication** handled by the Chrome extension (user's session).
- **No API keys** required.

### File Operations ✓ SECURE

- **Temporary backup before write** — a pre-save backup is written to the
  ephemeral work dir and the file is restored from it automatically if the save
  fails; on success the backup is deleted, so nothing accumulates in the data
  folder.
- **No stray artifacts** — the response cache and backups live in the OS temp work
  dir and are cleaned up on a clean apply.
- **Permission errors** — caught and reported (e.g. the spreadsheet is open/locked
  in Excel).

### Data Sharing ✓ SECURE

- **No personal data in the package** — the spreadsheet path, filename, and
  role/industry criteria all live in a **git-ignored** `config.json` **inside** the
  package, so cloning/pushing (or zipping) the skill never leaks anyone's data.

### LinkedIn Content Handling ✓ SECURE

- **Fetched LinkedIn text is treated as data, never as instructions.**
- **CAPTCHAs / bot checks are never solved** — the flow stops and hands off to the
  user.
- **No external APIs called** — profile URLs are used for navigation only.

## Error Handling & Recovery

### Save safety

Before saving, the skill writes a temporary backup to the work dir. If the save
fails, the original is restored from that backup; on success the backup is
deleted. See `ExcelManager.create_backup` and `_save_or_restore`.

### Cache misses

If APPLY reports `[Warning] N URL(s) were not in the cache`, those URLs were
requested but not pre-fetched. Fetch each listed URL via Chrome (fully expanding
the people tab), add them to `cache.json`, and re-run APPLY until there are no
misses.

### Error messages

Error messages are actionable, plain-language, and specific — they identify which
file, URL, or company caused the problem. Examples:

```
"Excel file is locked or in use. Please close it and try again: [path]"
"Contact 'Toni' already exists in column H. Run with --refresh to update."
```

## Performance Considerations

### Rate limiting

The research logic paces requests to look human:

```python
self.min_delay = 1.5  # seconds between requests
```

If LinkedIn throttles, increase this value.

### Tips

- **Batch contacts** — always prefer a single `--contacts` run over N separate
  runs; each company page is fetched once regardless of contact count.
- **Off-peak timing** — running during off-business hours reduces LinkedIn
  throttling.
- **Watch progress** — monitor stdout to detect stalls or CAPTCHA checkpoints.

## Verifying the Install

From `scripts/`:

- **Run the test suite:** `py test_skill.py` (Windows) or `python3 test_skill.py`
  (macOS/Linux). All tests should pass.
- **Dry-run PLAN** (prints the URL list without touching Excel or the browser):
  - Windows: `py "<scripts-dir>\linkedin_excel_contact_skill.py" "<name>" "<url>" --plan`
  - macOS/Linux: `python3 "<scripts-dir>/linkedin_excel_contact_skill.py" "<name>" "<url>" --plan`
- **Inspect the sheet (read-only, no Excel):** `<PY> "<scripts-dir>/linkedin_excel_contact_skill.py" --list-contacts`
  or `--show-row "<Company>"` (`<PY>` = `py` on Windows, `python3` on macOS/Linux;
  use your OS's path separators).

## Troubleshooting

### First-run message won't go away

**Issue:** every run prints the first-run/config message.

**Solution:** confirm `<skill-root>/config.json` exists and is valid JSON with
`excel_dir` and `excel_filename` set. Run `--setup` (or copy
`config.example.json`) to create it.

### Module import errors

**Issue:** `ModuleNotFoundError` when running the skill.

**Solution:** run from the `scripts/` folder so the sibling modules are importable,
and confirm `openpyxl` is installed.

### Browser not responding / no cards found

**Issue:** the people tab returns almost no text.

**Solution:** the "People you may know" block loads lazily — Claude must scroll and
click "Show more results" before reading. Verify the Chrome extension is connected
and the user is logged into LinkedIn.

### Spreadsheet won't save

**Issue:** save fails with a permission error.

**Solution:** close the spreadsheet in Excel — an open file is locked. The skill
restores from its temporary backup automatically on a failed save.

## Maintenance

### Updating role/industry criteria

Edit `<skill-root>/config.json` (`relevant_roles`, `relevant_industries`, optional
`relevant_role_patterns`). Changes take effect on the next run — no restart needed.

### Adding companies/contacts

No code changes needed — the spreadsheet grows organically as new rows/columns are
added by normal skill runs.

---

**Package:** self-contained git repo under `~/.claude/skills/find-connections-basic/`;
the git-ignored `config.json` lives inside it, while spreadsheets stay in the
user's data folder.
