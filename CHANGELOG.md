# Changelog

All notable changes to the **find-connections-basic** skill (formerly `add-excel-contact`).

## [2026-08-31] — Priority-aware writes, autofilter refresh, synthetic fixtures (PLAN-036)

### Added
- `excel_utils` — `priority_col` and `signal_col` on `SheetSchema` / `SheetStructure`,
  resolved by header text like the other fixed columns and `Optional` like them, so a
  sheet without those columns still loads and the writes are simply skipped. The
  `Momentum signal` match is a prefix match, so a dated variant of the header
  (`Momentum signal (Aug 2026)`) still resolves — a date belongs in the cell, not in a
  column key that code depends on.
- `excel_utils.refresh_autofilter()` — re-spans the sheet's autofilter over the full
  used range. Called before every save via `_refresh_filters()`, which re-detects the
  structure so a contact column added during the run is included.
- `--show-row` now reports `Priority` and `Momentum signal`, in sheet order.
- Test coverage for all of the above, including the older layout without the two
  columns, and the empty-sheet case.

### Changed
- New companies are written with `Priority = P2` and a deliberately blank
  `Momentum signal`. Never `P1`: nothing should read as a top target before a human
  has reviewed it, and the blank cell is the visible cue that the rating's
  justification is still owed.
- **All person names and LinkedIn slugs in tests, docs and examples are now
  synthetic.** The fixtures previously used real people's names and profile slugs,
  including strings asserting a connection between two named individuals. The
  replacements are shape-preserving — the suite depends on a four-character first
  name, a hyphenated surname, a single-character-deletion typo pair, two distinct
  slug formats, and a pair sharing a first name that must not collide — and every
  affected assertion was verified to behave identically before and after.
- `test_sheet_structure` no longer asserts literal column indices against the user's
  own workbook, nor names any column header from it. Both couplings were real bugs:
  the indices break whenever the fixed block is legitimately extended, and the header
  assertion read a real person's name out of private data.

### Fixed
- A workbook filename passed into the `name` or `url` positional (the workbook is the
  *third*) used to be ignored silently, and the run would fall back to the configured
  default — reporting one workbook's layout while the user believed they were
  inspecting another. It is now refused with the corrected invocation.

### Fixed (review pass)
- `create_sheet_from_schema` now emits `Priority` and `Momentum signal` too. It
  previously did not, so a **freshly created** sheet lacked both columns and the `P2`
  write on the next company added was silently skipped — the same defect this release
  fixes elsewhere, still present on the one path that builds a sheet from scratch.
- Seeded company rows are written **by header** instead of by `first_col + 1/2/3`.
  The offset form was correct only while the fixed block was four columns wide;
  widening it would have shifted every seeded value one column right of its label,
  with no exception raised and no test failing.
- `fixed_col_widths` is validated against the fixed-column list at construction, so
  adding a column without a width now fails loudly instead of silently under-applying.
- The fixed block's order is declared once, in `SheetSchema.fixed_headers()`.
- The test suite sets its own UTF-8 output encoding. It previously died with a
  `UnicodeEncodeError` on a legacy-codepage console unless the caller set
  `PYTHONIOENCODING=utf-8` first, which read like a code failure rather than a
  console limitation.
- `main()` error paths — including the misplaced-workbook guard — exit non-zero.
  The guard exists to stop a run that would otherwise read the wrong workbook, so a
  caller must not see it as a successful no-op.
- `DEFAULT_STATUS` / `DEFAULT_PRIORITY` derive from the schema rather than repeating
  its literals.
- Writer/reader round-trip tests resolve columns through the reader instead of
  asserting literal positions, and a new test seeds a company *through*
  `create_sheet_from_schema` and asserts each value landed under its own header —
  the case no existing test covered, because they all build sheets by hand.

### Fixed (peer-review pass)
- `refresh_autofilter` resolves **both** range edges from the sheet. The left edge
  previously came from `schema.first_col` — assuming a layout instead of reading one —
  so the filter started in the wrong place on any sheet whose `Company` column had moved.
- A contact column headed exactly `Priority` or `Momentum signal` is no longer absorbed
  as the fixed column of that name. Fixed columns live left of the `Status` anchor;
  anything right of it is a person. Previously such a column vanished from the contact
  block with no error.
- `import sys` is module-level rather than function-local, so the `main()` error paths no
  longer depend on the import happening to appear above them.
- The remaining literal column numbers in the new column tests resolve through the reader,
  completing the de-coupling the review pass began.


### Added (structure)
- `SHEET_CONTRACT.md` at the skill root — the layout and the eight rules for writing to
  the sheet, in one place. Previously this lived outside the skill entirely, so the
  description and the code that implements it could drift apart unnoticed; twice they did.
- `scripts/check_no_private_names.py` — pre-publish check that no real person from your
  own workbooks appears anywhere in this repository. Takes your file paths as arguments
  and hardcodes none of them. Refuses to report a clean result if it could not load a
  plausible number of names, or if a named sheet is missing, rather than reporting a pass
  it cannot justify.

### Changed (structure)
- `SKILL.md` no longer restates the sheet layout; it points at `SHEET_CONTRACT.md`.
  The `SheetSchema` docstring and `README.md` cite it too, so the rules are stated once
  per skill rather than in four places.

## [2026-07-28] — Portable `excel_dir` (PLAN-034)

Machine-independence pass, so this copy restores onto a machine with a different
username without editing anything.

### Changed
- `config.json` — `excel_dir` is now `~/admin/07_PM`, replacing an absolute
  path that named one specific user profile.
- `scripts/user_config.py` — `excel_dir` is now a property that expands `~` and
  environment variables **on read**. The raw string is kept in `_excel_dir`, so
  `save()` writes the portable form back rather than flattening it to an absolute
  path. Extends the `Path(__file__).resolve()` rule the skill already uses to find
  itself.

### Fixed
- A `~` in `excel_dir` used to be treated as a **literal directory name** —
  `os.path.join` does not expand it. Together with
  `os.makedirs(excel_dir, exist_ok=True)`, that meant a wrong path was silently
  *created* and a blank spreadsheet written into it: the run reported success while
  the real spreadsheet went untouched. Paths are now expanded before use.

## [2026-07-15] — Cross-platform docs (Windows + macOS/Linux)

Documentation-only pass (PLAN-S-030) so the repo runs on **both** Windows and
macOS/Linux from a single shared copy — the Python code was already OS-portable.

### Changed
- `SKILL.md` runtime guidance is now **OS-branching, not Windows-only**. Introduced
  a `<PY>` launcher placeholder — `py` on Windows, `python3` on macOS/Linux — and
  swept every `py "<SKILL_PY>" …` / `py -c` example to `<PY>`. `<SKILL_PY>` now
  documents both backslash (Windows) and forward-slash (macOS/Linux) absolute-path
  forms. Removed the machine-specific "use `py`, not `python`, not on PATH on this
  machine" assertion.
- `README.md` Requirements now shows both `py -m pip install openpyxl` (Windows)
  and `python3 -m pip install openpyxl` (macOS/Linux).
- `scripts/DEPLOYMENT_GUIDE.md` env line reworded to per-OS `<PY>`; added
  macOS/Linux (`python3` + forward-slash) command variants alongside the Windows
  ones for setup and install-verification.

### Fixed
- Genericized residual personal spreadsheet filenames to a neutral placeholder
  (`Shortlisted companies.xlsx`) in `scripts/test_skill.py` (`REAL_XLSX` fallback,
  only used behind `.exists()` guards) **and** the `scripts/linkedin_excel_contact_skill.py`
  `Usage:` docstring examples. Cosmetic — no code path uses either value.
- Corrected the macOS absolute-path example in `SKILL.md` to use `/Users/<you>/`
  (macOS home root) instead of the Linux-only `/home/<you>/`, with a note that
  Linux uses `/home/<you>/`.

## [2026-07-14] — Renamed to find-connections-basic

Renamed the skill `add-excel-contact` → `find-connections-basic` for naming
symmetry with its sibling `find-connections-advanced` (the Premium faceted-search
skill), following PLAN-S-029. Pure rename — slug, `/trigger`, human/display titles,
`.py` docstrings, and the ephemeral temp-work namespace
(`<temp>/find-connections-basic/`). No logic, config-format, or Python-filename
changes. Historical entries below keep the former name where it accurately
describes past state.

## [Unreleased]

Portable, shareable, structurally clean rework (PLAN-026, PLAN-032). The skill is
now a self-contained, git-shareable repository that ships no personal data and
leaves no artifacts in the spreadsheet folder.

### Added
- Per-user config at `<skill-root>/config.json` (`excel_dir`, `excel_filename`,
  `relevant_roles`, `relevant_industries`, optional `relevant_role_patterns`). It
  lives **inside** the skill folder but is **git-ignored**, so a clone/push never
  carries it and nothing is scattered outside the package. Ships
  `scripts/config.example.json` as a template.
- Shareable repository: `.gitignore` (keeps `config.json`, spreadsheets,
  `__pycache__`, and run artifacts out of git) and a `README.md` install /
  first-run guide, so the skill can be cloned into `~/.claude/skills/add-excel-contact`
  and used by anyone with no source edits.
- First-run wizard: `--setup <setup.json>` validates the setup JSON and, in one
  command, writes the per-user config **and** creates a starter **Company
  shortlist** sheet. Config and sheet outcomes are reported separately — if the
  sheet already exists, the config still saves and the wizard says so, pointing at
  `--overwrite`. Malformed input (non-object JSON, missing `excel_dir`/
  `excel_filename`) errors cleanly instead of crashing. A missing config prints a
  setup message pointing at `--setup`; no hardcoded defaults.
- First-run sheet creation: `--init-sheet --intake <json>` builds a fresh
  **Company shortlist** workbook from the shared `SHEET_SCHEMA` (refuses to
  clobber an existing file without `--overwrite`) — for when the config already
  exists and only a fresh sheet is needed.
- Batch contacts: `--contacts <json>` adds several contact columns in one run;
  each company's people page is fetched once and filtered for every contact, so N
  introducers cost the LinkedIn traffic of one.
- `SHEET_SCHEMA` single source of truth shared by the sheet reader
  (`detect_sheet_structure`) and writer (`create_sheet_from_schema`).

### Changed
- All code moved into `scripts/` so the skill is zip-and-send portable; data
  folder no longer holds `.py` or `__pycache__`.
- `SKILL.md` now documents the CLI path as home-relative
  (`~/.claude/skills/add-excel-contact/scripts/…`) with one pinned canonical
  invocation form — no username is baked in, so the skill is portable across users.
- `RoleMatcher` is built entirely from the per-user config — no hidden built-ins.
- Fetch model reworked for the `/company/<slug>/people/` surface: fetched once and
  expanded via scroll + "Show more results" during FETCH, never `?page=N`. Only
  the `/search/results/people/` keyword fallback still paginates. Removed the
  `len(text) < 200` early-stop that aborted on a freshly-navigated page.
- Contact column headers are now the contact's bare full name (was "Potential
  contacts via <name>"). Contact columns are detected positionally — any header to
  the right of Status — so no header prefix is required.

### Fixed
- New/appended contact columns now write the contact's own LinkedIn profile URL to
  row 5, matching the init-sheet path (previously only the create path did this).
- Contact-column matching bridges first-name-only headers and full-name lookups
  (an "Wren Rivera" lookup finds an existing "Wren" column), so a refresh/append
  updates the column instead of creating a silent duplicate; two distinct full
  names sharing a first name still stay separate.
- Repeated names within one `--contacts` batch dedupe to a single column.
- Industry parser handles LinkedIn's inline no-city header form
  ("Software Development  37M followers  10K+ employees").
- Company duplicate detection no longer false-positives on short names (e.g.
  "Meta" matching "Metabase"); substring match now requires ≥ 5 chars, shorter
  names fall through to fuzzy-ratio matching.
- `--apply` cache cleanup guards `os.remove` against a locked/vanished file.
- Malformed `--contacts`/`--intake` JSON now reports a friendly error instead of
  raising `KeyError` mid-run.
- Pre-save backup filenames use microsecond timestamps to avoid same-second
  collisions in the shared work dir.

### Removed
- `fetch_queue.json` and dead `read_queue`/`write_queue` — `--plan` prints URLs to
  stdout only.
- Hardcoded `DEFAULT_BASE_PATH`, `RoleMatcher.BUILT_IN_PATTERNS`, `DEFAULT_TITLES`,
  and the `coaching-state.md` dependency.
- Dead `contact_header_prefix` schema field — contact columns are now detected
  positionally rather than by header prefix.

### Security
- No personal data ships in the package — spreadsheet path, filename, and
  role/industry criteria all live in the external per-user config.
- Response cache and pre-save backups route to an ephemeral OS temp work dir
  (`<temp>/add-excel-contact/`); cache deleted on a clean apply, backups deleted
  on a confirmed save — nothing accumulates in the data folder.

---

Permission-prompt reduction (PLAN-S-023). Repeat runs no longer prompt per
command: the CLI is invoked by absolute path (no `cd … &&` compound), and
routine inspection goes through the script instead of ad-hoc `py -c` one-liners.

### Added
- Read-only helper flags (all print-and-exit): `--print-cache-path` (emit the
  response-cache path, honoring `--cache`), `--show-row "<Company>"` and
  `--list-contacts` (inspect the sheet without opening Excel), and
  `--build-cache <map.json>` (assemble the response cache from a
  `{ "<url>": "<page_text_file>" }` map). Only `--build-cache` writes — to the
  work dir or `--cache`.

### Changed
- SKILL.md invokes the CLI by absolute path
  (`py "<abs>\…\linkedin_excel_contact_skill.py" …`) instead of
  `cd scripts/ && py …`, and standardizes on the `py` launcher. A single
  allow-list rule then covers every invocation, so repeat runs stop triggering a
  permission prompt per command. FETCH/APPLY now use the helper flags rather than
  ad-hoc `py -c`.

### Fixed
- `--build-cache` hardening: a bare `--cache` filename no longer crashes on
  `os.makedirs("")`; the cache write is wrapped in `try/except OSError`; and a
  non-string map value now reports a clear error instead of raising an uncaught
  `TypeError`.

---

People-tab connection reliability (PLAN-031). More accurate connection discovery
from the `/company/<slug>/people/` tab, with guards against mixing in another
company's people.

### Added
- Current-company filter: connection results now drop people whose current
  employer clearly isn't the target company (the people tab leaks unrelated people
  in). Conservative — keeps substring, short-abbreviation ("AWS" under "Amazon"),
  fuzzy, and unparseable-company cases; drops only a clear mismatch.
- Org-verification guard on the people tab: a company URL that resolves to the
  wrong company, or to an unavailable/removed page, is skipped with a distinct
  warning (backstop to the FETCH-time title check) so another company's
  connections aren't attributed to the target. Accepts a match on the sheet's
  company name **or** the company URL slug, so a correct URL with a
  differently-worded label ("Alphabet" vs slug `google`) still matches.

### Changed
- FETCH (SKILL.md): the people tab is now exhausted fully — keep clicking
  "Show more results" until it disappears or the pager reads "Page N of N" (was a
  ~10-expansion cap), reading incrementally after each expansion and retrying on a
  frozen renderer (CDP timeout) instead of accumulating the whole list then reading
  once. FETCH also verifies the tab resolves to the expected company before
  harvesting, skipping bad/`unavailable` slugs.
