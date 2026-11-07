# PHDTech Commission Tool

This application handles employee compensation. Preserve the validated Windows
workflow, header-based mapping, authenticated Playwright session, sequential
workbook processing, and direct XLSX writer. Prefer focused changes with regression
tests; avoid major rewrites or additional dependencies without a demonstrated need.

## Data integrity

- Resolve workbook columns by header name. Missing or duplicate required headers
  stop safely. Never assume physical column numbers.
- ConnectWise is authoritative for every amount/date field populated from it.
  Correct differing nonblank values. Explicit blank dates may clear stale dates;
  absent response fields or malformed values must not be interpreted as zero/blank.
- Preserve the full-payment commission eligibility rule. A fully paid invoice's
  pretax commission basis is its subtotal; unpaid/partial invoices have blank paid
  fields. Never substitute a partial gross payment for a pretax commission basis.
- Never change percentage values, commission formulas or employee-maintained
  paycheck information to resolve source discrepancies.
- Never reintroduce paid-history invoices or create duplicate active invoices.
  Historical `Paid to Emp*` sheets are read-only. Flag overlaps and conflicts.
- Stage a project's changes until its retrieval succeeds. Keep discarded proposals
  clearly distinct from applied changes in the audit.
- Create an exclusive, unchanged original backup before any write. Preserve
  formulas, formatting, tables and unrelated XLSX parts. Unsupported writes fail
  before replacement. Do not reintroduce Excel COM/PowerShell workbook writing.
- Save run audits and summaries on no-change and failed runs too. Report partial
  failures clearly; do not label a partial run as wholly successful.
- Never commit employee workbooks, payment records, credentials, browser profiles,
  real backups or sensitive logs. Tests generate synthetic fixtures in temp folders.
- Keep workbooks/backups/logs in the explicitly selected workbook directory and
  settings/browser profiles/temp plans in the per-user data directory. Display the
  selected folder, preserve it on cancellation/failure, and never silently switch
  to the application directory when it becomes unavailable. Do not move user data.
- Folder-picker PowerShell is limited to the native Windows Forms selection dialog,
  with a typed-path fallback; it must not automate Excel or perform workbook writes.

## Development and validation

Use Node.js 24 and `npm ci`. Run `npm test`; tests must not require live ConnectWise,
a browser, or desktop Excel. Cover meaningful failure and repeat-run cases.
Run the suite in UTC and America/Chicago when changing date handling. CI also runs
on Windows. Test first to reproduce a defect, then make a focused correction.

Keep Windows/live ConnectWise/desktop Excel acceptance distinct from synthetic CI
results. The six real production workbooks are the acceptance baseline. Do not
claim validation that did not occur. Use branches and pull requests for meaningful
changes, explain behavior in plain language, and include concrete test results.
