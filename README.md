# PHDTech Commission Tool

Internal Windows tool for reconciling ConnectWise PSA invoice payments against employee commission Excel workbooks.

## Current working workflow (November 2026)

1. Open `Run Commission Tool.cmd` from the project directory.
2. Choose one workbook or **Process all workbooks**. The current launcher discovers eligible `*.xlsx` files in the same folder as the project.
3. Sign in to ConnectWise in the Playwright browser if needed, then press Enter in the console to continue.
4. Review the proposed reconciliation changes. Before a write, the tool creates a timestamped file in `Backups/`.
5. The writer edits XLSX files directly (no Microsoft Excel COM automation), and records a reconciliation CSV plus writer log in `Logs/`.
6. Open changed workbooks in Microsoft Excel and allow formula recalculation before using commission totals for payroll. Review any workbook-only invoices flagged for manual review.

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

## Key source files

- `app/commission-menu.js`: workbook discovery and ConnectWise browser session
- `app/connectwise.js`: ConnectWise retrieval
- `app/app.js`: reconciliation and write-plan generation
- `app/workbook.js`: workbook reading, header mapping and paid-history support
- `app/xlsx-writer.js`: current offline XLSX writer, including row growth and highlight handling
- `app/apply-workbook-changes.ps1`: obsolete Excel-COM writer, retained only for historical reference; **not used** by the current reconciliation path

## Known scope

This repository sync reflects the Windows version validated in November 2026, which still uses workbooks next to the launcher. Earlier experiments with first-run workbook-folder selection and an Inno Setup installer are **not included** in this baseline; they should be integrated and tested in a separate change.
