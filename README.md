# random-work

## ILSFILE2 debit-decline counter

Counts debit-decline events from `ILSFILE2` files in the TNB Velocity archive
and writes per-day and per-month totals to an Excel workbook.

### Source layout

```
TNB_VelocityArchive_Jan-Mar_2026\
  2026-01-01\
    ILSFILE2          <- comma-delimited, no header
  2026-01-02\
  ...
```

`ILSFILE2` columns (no header in file):
`Date, Acct#, TranType, Amt, Last4Card, Description, TerminalID`

Each row = one debit decline event.

### Run

```
pip install pandas openpyxl
python count_ilsfile2.py
```

The base path is hard-coded for the one-time run against
`C:\Users\james.gilmore\Downloads\TNB_VelocityArchive_Jan-Mar_2026`.
Edit the `BASE` constant at the top of the script if needed.

### Output

`ILSFILE2_decline_counts.xlsx` next to the archive folder, with two sheets:

- **Daily** — `date, count, status, malformed, note`
  - `status`: `ok`, `empty` (file present but no rows), `missing_file`
- **Monthly** — per-month totals plus a `TOTAL` row

### Validation

- Blank lines skipped.
- Rows must have exactly 6 commas (7 fields); anything else counts as
  `malformed` so it doesn't silently inflate the decline count.
- Missing daily folders are skipped silently.
- Missing `ILSFILE2` is noted, not an error.
