# Company shortlist — sheet contract

The layout this skill reads and writes, and the rules any code touching it must follow.
`scripts/excel_utils.py` implements all of it; this document explains *why* the rules
are what they are.

The workbook is whichever file `config.json` names (`excel_dir` + `excel_filename`), or
whatever is passed as the third positional argument on the command line. The sheet
inside it is `Company shortlist`.

## Layout

| Row | Contents |
|-----|----------|
| 4 | header row |
| 5 | each introducer's own LinkedIn profile URL (blank across the fixed block) |
| 6+ | company data |

Column A is intentionally blank; the title sits at B2. Two regions, with very
different stability:

```
  B        C          D                  E              F          G        H ......
+--------+----------+------------------+--------------+----------+--------+-----------------+
| Company| Priority | Momentum signal  | Linkedin URL | Industry | Status | introducers ... |
+--------+----------+------------------+--------------+----------+--------+-----------------+
         <---------------- FIXED LEFT BLOCK ---------------------> <-- GROWS RIGHTWARD -->
```

- **Fixed left block.** A stable set of per-company fields. Its width is *not* a
  constant — `Priority` and `Momentum signal` were added to it after the first version,
  which shifted everything to their right.
- **Introducer block.** One column per person who can make an introduction, headed by
  that person's name. **It grows every time someone is added.** It must stay the
  rightmost region so that growth never collides with anything else.

That asymmetry is the reason for rule 1.

## Rules

1. **Resolve every column by header text — never by letter or index.** Column positions
   are not stable and have already moved once. Code that says "Industry is column D" is
   correct until the day it silently isn't.
2. **A new introducer is appended at the right end** (`max(column) + 1`), with their name
   on row 4 and their LinkedIn URL on row 5.
3. **A new fixed field is inserted inside the left block**, never appended at the right —
   the next introducer added would overwrite it. This is *enforced*, not merely advised:
   everything right of `Status` is read as an introducer, so a column headed `Priority`
   parked out there is treated as a person of that name, not as the fixed field.
4. **`Priority` accepts only `P1`, `P2`, `P3`.** A newly added company defaults to **`P2`**,
   never `P1`, so nothing reads as a top target before a human has reviewed it.
5. **`Momentum signal` is the dated one-line justification for `Priority`** (funding round,
   growth rate, layoffs, guidance change). Write the two together; never set a rating
   without its evidence. A new company gets `P2` plus a *blank* signal — the empty cell is
   the visible cue that the justification is still owed.
6. **Match company names case-insensitively** and reject duplicates when adding.
7. **Copy cell styles from the previous data row** when appending, so formatting holds.
8. **Refresh the autofilter across the full used range after any structural change**
   (a row appended, a column added). Both edges are resolved from the sheet, so the
   filter follows the data rather than an assumed starting column.

## Older sheets

A sheet predating `Priority` / `Momentum signal` still loads. Both resolve to `None` and
their writes are skipped. Only a missing `Company` column is fatal.

That is deliberate: a missing optional column means "this sheet does not have that
field", not "this sheet is broken" — so schema drift surfaces as a skipped write, never
as a write into the wrong column.

## Using it from code

```python
from excel_utils import detect_sheet_structure, refresh_autofilter

ws = load_workbook(path)["Company shortlist"]
s = detect_sheet_structure(ws)      # every column resolved BY HEADER TEXT

s.company_col, s.priority_col, s.signal_col    # priority/signal are None on an
s.url_col, s.industry_col, s.status_col        # older sheet that lacks them
s.contact_cols                                  # {introducer name: column index}
s.next_contact_col_idx(ws)                      # where the next introducer goes

refresh_autofilter(ws, s)           # rule 8, after any structural change
```

Every fixed column except `Company` is `Optional[int]`, so guard before writing:

```python
if s.priority_col:
    ws.cell(row=r, column=s.priority_col).value = "P2"
```

`SheetSchema.fixed_headers()` is the single declaration of the left block's order. Any
code writing a fixed column must resolve its position through that order, or through the
reader — never by offset arithmetic from the first column. Offsets were the cause of a
real bug: they were correct only while the block happened to be four columns wide, and
widening it shifted every seeded value one column right of its own label, silently.

Rules 4 and 5 are applied together in
`LinkedInExcelContactSkill._insert_company_row`.
