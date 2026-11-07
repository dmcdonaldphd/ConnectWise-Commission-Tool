# PHDTech Commission Tool

Internal Windows tool for reconciling ConnectWise PSA invoice payments against employee commission Excel workbooks.

## Current working workflow (November 2026)

1. Open `Run Commission Tool.cmd` from the project directory.
2. On first launch, select the folder containing your working commission workbooks. The tool remembers it. Choose one workbook or **Process all workbooks** from that folder.
3. Sign in to ConnectWise in the Playwright browser if needed, then press Enter in the console to continue.
4. Review the proposed reconciliation changes. Before a write, the tool creates a timestamped file in the selected workbook folder's `Backups/` directory.
5. The writer edits XLSX files directly (no Microsoft Excel COM automation), and records reconciliation CSVs, run summaries and writer logs in that folder's `Logs/` directory.
6. Open changed workbooks in Microsoft Excel and allow formula recalculation before using commission totals for payroll. Review any workbook-only invoices flagged for manual review.

## Workbook folder and settings

The menu always shows the active workbook folder. **Change workbook folder** opens
a folder picker; canceling keeps the current selection. **Open Logs folder** and
**Open Backups folder** launch visible Windows Explorer windows for the current
selection and print the requested folder path. Explorer runs independently of the
console menu; launch errors include the Windows error and full folder path. **0. Exit**
closes the tool before connecting to ConnectWise. If a folder is empty or no longer
available, choose another folder from the menu. The scan is not recursive and
excludes Excel lock files, working copies, and the tool's timestamped backups.
Select the folder holding the working workbooks, not a backup folder.

The application does not move workbooks or existing backups/logs when a folder
is selected. New backups and logs go inside the selected workbook folder. Settings,
the browser profile and temporary write plans are stored per Windows user under:

```text
%LOCALAPPDATA%\PHDTech\Commission Tool\
├── settings.json
├── connectwise-profile\
└── Temp\
```

The first run of this version may require signing into ConnectWise again because
the browser profile now lives in the per-user location. Changing workbook folders
does not clear that profile. Old project-folder profiles are retained for rollback;
they are not copied into or deleted by the new settings workflow.

Copies of the application run by the same Windows user share these settings and
the browser profile. When testing on copied workbooks, choose the test folder and
later use **Change workbook folder** to restore the regular workbook folder.
An unavailable or malformed saved selection prompts for a new selection; it never
silently falls back to processing the application directory. Settings are replaced
atomically only after an explicit valid folder selection.

The Windows folder picker uses the built-in Windows Forms dialog through
PowerShell; it performs no workbook operations or Excel automation. If company
policy blocks it, the tool offers a full-folder-path prompt instead (pasted quoted
paths are accepted). For isolated development/testing, `PHD_COMMISSION_DATA_DIR`
can point to a separate private data directory. Normal users need no environment
variables or configuration-file edits.

## Invoice amount and date

On the `Commission` worksheet, reconciliation fills **Inv Amt (Pretax)** from
ConnectWise's invoice subtotal and **Invoice Date** from its invoice date. It
replaces differing existing values, fills project invoice slots, and populates
newly inserted invoices. These fields are populated for unpaid and partially paid
invoices as well as fully paid invoices; the existing payment and commission
eligibility rules still apply. Already-paid and duplicate-workbook invoices are
excluded. New zero-dollar invoices are ignored; an invoice already on Commission
is refreshed if its ConnectWise subtotal has changed to zero.

Both column headers are required and may be reordered. Existing-row changes are
highlighted yellow; newly populated invoice fields and inserted rows are green.
The reconciliation CSV includes old/new invoice values. ConnectWise is the source
of truth for all four managed amount/date fields: invoice amount, invoice date,
paid amount and paid date. Differing nonblank values are corrected too. An
explicitly empty source date clears an old date. A missing response field or
invalid subtotal, balance, or date holds that project's changes for review rather
than inferring a value. Historical `Paid to Emp*` worksheets are never updated.

## Payment corrections and run results

The existing full-payment eligibility rule is preserved: **Paid Amt (Pretax)** is
the invoice subtotal only when ConnectWise reports the invoice fully paid. If it
now reports unpaid or partially paid, old paid amounts and paid dates are cleared;
the commission formula remains in place for Excel to recalculate. Fully paid
invoices receive the current ConnectWise paid date, including corrections or an
explicitly blank date. Partial gross payments are not written into the pretax
commission basis.

Every selected workbook run saves a reconciliation CSV and a readable **Run
Summary** text file in `Logs`, including no-change runs and failed runs. The CSV's
`Result` column distinguishes `APPLIED`, `NOT_APPLIED`, `REVIEW` and other status
entries. A writer error with an uncertain file outcome is labelled `UNCONFIRMED`
for comparison with the backup. Final row numbers account for inserted rows, with original addresses
retained in `Original Workbook Row` for comparison with the backup. Proposed
changes are saved before writing, then finalized after the writer returns.

Project changes are held until all invoice retrieval for that project succeeds.
If one project fails, its proposed changes are discarded and labelled as not
applied; completed projects can still be written. The menu lists each workbook's
result, continues sequentially after an individual workbook failure, and returns
a nonzero exit code when projects or workbooks fail.

Invoice checks span the Commission sheet, including rows with missing/invalid
project IDs. Identical repeated source records are processed once. Conflicting
records or invoices returned under multiple source projects hold the affected
projects for review. Existing duplicates and active/paid-history overlaps are
flagged for manual review. Paid-history scans include records after blank rows
and read cached formula values; an uncached invoice formula stops safely.
ConnectWise invoice lists are paginated past 250 results and repeated pages are
rejected as incomplete retrievals.

## Data safety

- Already-paid invoices are identified from `Paid to Emp*` worksheets; they must not be reintroduced as new payable commissions.
- Relevant workbook fields are resolved by **header name**, rather than a fixed Excel column number.
- `app/xlsx-writer.js` preserves historical sheets, edits the Commission sheet, updates styles, clears prior automation highlights, adds new rows (expanding the table when needed), and forces formula recalculation on the next Excel open.
- The writer rejects unsupported/inconsistent workbook structures instead of silently overwriting unrelated data.
- Keep sensitive workbooks, browser sessions, backups, logs, and credentials out of GitHub. The `.gitignore` excludes local `.xlsx` workbooks and normal runtime data.

## Developer setup

Use Node.js on Windows, then from the repository root:

```cmd
npm ci
npx playwright install chromium
```

To launch without the bundled `runtime/node.exe`, run `npm start` or `node app\commission-menu.js`. The included `.cmd` launcher expects the portable runtime under `runtime/` and supports application paths containing spaces.

**There is not yet a validated zero-dependency Windows installer.** Deployment currently requires the Node dependencies and Playwright Chromium browser to be present; this will be addressed separately.

## Automated regression tests

With the dependencies installed, run `npm test` from the repository root using
Node.js 24. The tests generate synthetic workbooks in temporary directories and
mock ConnectWise responses. They exercise the actual offline writer without
launching a browser, connecting to ConnectWise, or using Microsoft Excel.

Coverage includes source corrections and reversals, sparse paid history, duplicate
prevention, failed-project rollback, saved run outcomes, pagination, header mapping,
blank-row fills, insertion with and without table growth, highlights, repeated
runs, backups, formulas, historical sheets, rejected writes, remembered folders,
selection cancellation/recovery, and selected-folder reconciliation. GitHub Actions
runs the suite on Linux and Windows in UTC and America/Chicago time zones.
Desktop Excel recalculation and a live run against copies of the six production
workbooks remain separate Windows acceptance checks.

## Key source files

- `app/commission-menu.js`: workbook discovery and ConnectWise browser session
- `app/settings.js`: folder selection, per-user settings, and opening logs/backups
- `app/paths.js`: application root and configured workbook/per-user paths
- `app/connectwise.js`: ConnectWise retrieval
- `app/app.js`: reconciliation and write-plan generation
- `app/workbook.js`: workbook reading, header mapping and paid-history support
- `app/xlsx-writer.js`: current offline XLSX writer, including row growth and highlight handling
- `app/apply-workbook-changes.ps1`: obsolete Excel-COM writer, retained only for historical reference; **not used** by the current reconciliation path

## Known scope

The reconciliation baseline was validated against the six production workbooks in
November 2026. Remembered workbook-folder selection and per-user storage prepare
the application for installation; the native folder dialog and live workflow need
their own Windows acceptance run. A bundled Windows installer remains separate
follow-up work and is not yet implemented or validated.
