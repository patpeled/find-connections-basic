# Find Connections (Basic)

A [Claude Code](https://claude.com/claude-code) skill that finds **who in your
LinkedIn network can introduce you** to people at your target companies, and
writes the results into a "Company shortlist" Excel spreadsheet
(layout and rules: [`SHEET_CONTRACT.md`](SHEET_CONTRACT.md)).

It discovers introducers from each company's **"People you may know"** tab
(`/company/<slug>/people/`) — the curated mutual-connection surface, which works
on a **free LinkedIn account**. Claude drives your own logged-in Chrome session
(via the Claude-in-Chrome extension); the Python code does all the parsing and
Excel work. **No personal data ships in this repo** — your Excel path and search
criteria live in a git-ignored config file (created on first run) that sits inside
the skill folder, so it's never committed yet leaves nothing scattered elsewhere
on your machine.

> **Which skill?** This is the **basic** skill (free-tier, people-tab). The sibling
> `find-connections-advanced` uses exhaustive faceted search but needs a LinkedIn
> Premium / under-commercial-limit account. Use **basic** on a free account; use
> **advanced** when you have Premium.

## Requirements

- **Claude Code** (desktop, CLI, or IDE) with the **Claude-in-Chrome** extension
  connected, and **Chrome logged into LinkedIn**.
- **Python 3** on your PATH, with **openpyxl** installed. Use the launcher for
  your OS:
  - **Windows:** `py -m pip install openpyxl` (the skill invokes the `py` launcher).
  - **macOS/Linux:** `python3 -m pip install openpyxl` (the skill invokes `python3`).

## Install

Clone into your Claude skills folder, using **exactly** this folder name so Claude
discovers it and the documented CLI path resolves (`$HOME` expands to your home
directory in bash and PowerShell; the quotes are safe there):

```
git clone <repo-url> "$HOME/.claude/skills/find-connections-basic"
```

Restart Claude Code if it was already running, so the new skill is picked up.

## First-run setup

Just ask Claude to **"add a contact"** (or run `/find-connections-basic`). On first run
there's no config yet, so Claude will:

1. Ask you for your **Excel folder + filename**, your **relevant roles /
   industries**, and (optionally) a **starter list** of key contacts and target
   companies.
2. Run the built-in wizard, which writes your private config to
   `<skill-root>/config.json` (git-ignored) **and** creates a starter
   **Company shortlist** workbook — in one step.

Your config lives inside the skill folder but is git-ignored, and your
spreadsheets stay wherever you point `excel_dir`; the repo only ever holds code.

## Usage

Ask Claude in natural language, or use the trigger form:

- **Add an introducer** (a contact column, searched across every company row):
  `/find-connections-basic "Toni Vale-Hollis" https://linkedin.com/in/shai...`
- **Add a target company** (a company row, searched via each existing contact):
  `/find-connections-basic "Coralogix" https://linkedin.com/company/coralogix`
- **Add several introducers at once:** point Claude at a `[{"name","url"}, ...]`
  JSON via `--contacts` — each company's people tab is fetched once and filtered
  for every contact, so N introducers cost the same LinkedIn traffic as one.

Claude confirms what it will write before touching your spreadsheet.

## How it works (three phases)

Because the Python process can't call the browser itself, the flow is split so
Claude is the browser in the middle, passing data through one JSON cache:

1. **PLAN** — the code enumerates the LinkedIn URLs to fetch (one
   `/company/<slug>/people/` per company) and prints them.
2. **FETCH** — Claude drives Chrome to each URL, expands the "People you may know"
   block fully, and stores the page text in a response cache.
3. **APPLY** — the code parses each company's page, filters it for every requested
   contact, and writes your spreadsheet (with a temporary backup for safety).

`SKILL.md` is the authoritative runtime guide; `scripts/DEPLOYMENT_GUIDE.md`
covers setup, security, and troubleshooting.

## Privacy & safety

- Personal data (spreadsheet path, filename, roles, industries) lives in a
  **git-ignored** config inside the skill folder — never committed, never
  scattered elsewhere on your machine.
- Run artifacts (response cache, pre-save backups) live in an ephemeral OS temp
  dir and are deleted on a clean run.
- The skill treats all fetched LinkedIn text as **data, never instructions**, and
  **never** solves CAPTCHAs — it hands off to you.
