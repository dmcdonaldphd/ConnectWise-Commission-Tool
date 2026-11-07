# PHDTech Commission Tool

Internal Windows tool for reconciling ConnectWise PSA invoice payments against employee commission Excel workbooks.

## Windows installer preview

The installer packages the portable Node runtime, production dependencies, and the
Chromium revision matching the locked Playwright version. On a workstation it
requires no npm commands or separate browser download. Installation is per user,
with a Start Menu shortcut and an optional desktop shortcut.

Preview builds are available from the **Windows installer preview** GitHub Actions
workflow. Open a successful run, download its **PHDTech-Commission-Tool-Windows-x64**
artifact, extract the ZIP, and run `PHDTech-Commission-Tool-Setup-1.2.0.exe`.
Launch the installed application from the Start Menu. For acceptance testing,
choose a folder holding copies of the commission workbooks before processing.
The installed application shares the remembered folder and browser profile with
other copies run by the same Windows user; confirm the displayed folder.

The default application location is `%LOCALAPPDATA%\Programs\PHDTech Commission Tool`.
Workbooks, backups, logs, settings and browser sign-in remain separate and are
preserved across updates and removal. The installer does not change system PATH
or install development tools. Preview installers are unsigned and still require
testing on a standard company workstation before general deployment.

## Desktop workflow (November 2026)

1. Open **PHDTech Commission Tool** from the Start Menu or desktop shortcut. The installed executable opens the HTML desktop interface without a console window.
2. Choose the folder containing your workbooks, or confirm the remembered folder. Check individual workbooks or **Select all**. Close selected files in Excel.
3. Click **Open ConnectWise**, sign in in its separate browser window, then return to the app and click **Continue reconciliation**. Keep both windows open while processing finishes.
4. Review progress and each workbook's result. Before a write, the tool creates a timestamped file in the selected workbook folder's `Backups/` directory. **Run summary** opens the readable saved summary inside the app; failed and partially failed runs are explicitly identified.
5. The writer edits XLSX files directly (no Microsoft Excel COM automation), and records reconciliation CSVs, run summaries and writer logs in that folder's `Logs/` directory.
6. Open changed workbooks in Microsoft Excel and allow formula recalculation before using commission totals for payroll. Review any workbook-only invoices flagged for manual review.

## Workbook folder and settings

The interface always shows the active workbook folder. **Change** opens a folder
picker; canceling keeps the current selection. Results include **Logs folder** and
**Backups folder** buttons that open Windows Explorer. The activity log contains
technical details. If a folder is empty or no longer available, choose another
folder. The scan is not recursive and
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
An unavailable or malformed saved selection requires a new selection; it never
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

The source checkout starts the desktop interface with `npm start` or
`node app/gui.js`. `npm run start:console` retains the validated console workflow
for troubleshooting. `Run Commission Tool.cmd` opens the installed executable when
present and otherwise starts the source desktop interface.

The HTML/CSS/JavaScript interface is served only on `127.0.0.1` at a temporary port
and displayed by the bundled Chromium in an app window. It uses an unpredictable
session token, verifies request host/origin, and loads no remote scripts or fonts.
It is not a hosted website. ConnectWise still opens separately for authentication;
the financial engine and direct XLSX writer are shared by both interfaces.
Selecting another folder or starting another run is blocked during processing.
Closing the desktop window lets the selected batch finish before closing the
ConnectWise session; it never forcibly terminates the writer.

Source checkouts still need the setup above. The installer preview includes these
components; end users of the installer do not perform developer setup.

## Building the installer

Use a Windows x64 build machine with Node.js 24 and Inno Setup 6, then run
`npm run build:installer`. Downloads happen on the build machine. Distribution
files are staged under `dist/package`, and the setup EXE is written to `dist`.
The build copies an explicit list of application files and the bundled Node runtime,
installs production dependencies using the lockfile, downloads compatible Chromium,
rejects workbook/profile data, compiles a small Windows GUI launcher using the
built-in .NET Framework compiler, and smoke-tests the packaged writer and browser.
No dependency versions were upgraded for packaging.

The Windows workflow also installs the EXE into a path with spaces, tests its
bundled runtime, executable launcher, actual XLSX writer and desktop UI, reinstalls it, and removes
it. Synthetic settings, browser-profile and workbook markers must survive each
step, including an untracked workbook inside the application directory. These
checks are separate from real ConnectWise/desktop Excel and company-PC testing.
`installer/test-install.ps1` is a CI smoke test and refuses an account with existing
application data or an installed copy. It is not an end-user setup command.

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

- `app/gui.js`, `gui-service.js`, `gui-server.js`, `ui/`: desktop window, workflow, local server and HTML interface
- `app/commission-menu.js`: shared workbook discovery and console fallback
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
the application for installation and have been checked on Windows. Installer
packaging is a preview with automated Windows install/update/removal checks;
company-PC installation and live reconciliation remain its acceptance gate.
