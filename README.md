# PHDTech Commission Tool

Internal Windows tool for reconciling ConnectWise PSA invoice payments against employee commission Excel workbooks.

## Current working workflow (November 2026)

1. Open `Run Commission Tool.cmd` from the project directory.
2. Choose one workbook or **Process all workbooks**. The current launcher discovers eligible `*.xlsx` files in the same folder as the project.
3. Sign in to ConnectWise in the Playwright browser if needed, then press Enter in the console to continue.
4. Review the proposed reconciliation changes. Before a write, the tool creates a timestamped file in `Backups/`.
5. The writer edits XLSX files directly (no Microsoft Excel COM automation), and records a reconciliation CSV plus writer log in `Logs/`.
6. Open changed workbooks in Microsoft Excel and allow formula recalculation before using commission totals for payroll. Review any workbook-only invoices flagged for manual review.

## Invoice amount and date

On the `Commission` worksheet, reconciliation fills **Inv Amt (Pretax)** from
ConnectWise's invoice subtotal and **Invoice Date** from its invoice date. It
replaces differing existing values, fills project invoice slots, and populates
newly inserted invoices. These fields are populated for unpaid and partially paid
invoices as well as fully paid invoices; the existing payment and commission
eligibility rules still apply. Already-paid, duplicate-workbook, and zero-dollar
invoice exclusions are unchanged.

Both column headers are required and may be reordered. Existing-row changes are
highlighted yellow; newly populated invoice fields and inserted rows are green.
The reconciliation CSV includes old/new invoice values. Missing or invalid source
dates are flagged for review and do not erase existing dates; missing or invalid
subtotals produce a project error rather than an inferred amount. Historical
`Paid to Emp*` worksheets are never updated by this enhancement.

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

To launch without the bundled `runtime/node.exe`, run `node app\commission-menu.js`. The included `.cmd` launcher expects the portable runtime under `runtime/`.

**There is not yet a validated zero-dependency Windows installer.** Deployment currently requires the Node dependencies and Playwright Chromium browser to be present; this will be addressed separately.

## Automated regression tests

With the dependencies installed, run `npm test` from the repository root using
Node.js 24. The tests generate synthetic workbooks in temporary directories and
mock ConnectWise responses. They exercise the actual offline writer without
launching a browser, connecting to ConnectWise, or using Microsoft Excel.

Coverage includes invoice-field replacement, reordered/missing/duplicate headers,
blank-row fills, insertion with and without table growth, highlights, repeated
runs, backups, formulas, historical sheets, and rejected writes. GitHub Actions
runs the suite on Linux and Windows in UTC and America/Chicago time zones.
Desktop Excel recalculation and a live run against copies of the six production
workbooks remain separate Windows acceptance checks.

## Key source files

- `app/commission-menu.js`: workbook discovery and ConnectWise browser session
- `app/connectwise.js`: ConnectWise retrieval
- `app/app.js`: reconciliation and write-plan generation
- `app/workbook.js`: workbook reading, header mapping and paid-history support
- `app/xlsx-writer.js`: current offline XLSX writer, including row growth and highlight handling
- `app/apply-workbook-changes.ps1`: obsolete Excel-COM writer, retained only for historical reference; **not used** by the current reconciliation path

## Known scope

This repository sync reflects the Windows version validated in November 2026, which still uses workbooks next to the launcher. Earlier experiments with first-run workbook-folder selection and an Inno Setup installer are **not included** in this baseline; they should be integrated and tested in a separate change.
