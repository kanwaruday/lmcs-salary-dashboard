# LMCS Salary Dashboard — Project Context

## Vault context

Read this file before starting any work:
`~/Library/CloudStorage/OneDrive-Personal/2. Areas/Uday Obsidian KMS/Uday's KMS/work/active/lmcs-salary-dashboard.md`

It contains: current status, open tasks, key decisions.

## Quick summary

HR and payroll dashboard for La Montessori Chandpur School (LMCS). Static HTML frontend backed by Google Apps Script connected to Google Sheets.

- No build step — pure HTML/CSS/JS
- Backend: `apps-script/` — deploy via Google Apps Script editor
- Data source: Google Sheets (linked via Apps Script)
- No local server needed — open HTML files directly or serve statically
- Pure-logic `.gs` files (`Constants.gs`, `IncentiveRules.gs`, `LeaveRules.gs`, `PayrollConstants.gs`, `PayrollRules.gs`) have no `SpreadsheetApp`/`Session` calls, so they run unmodified under plain Node — run their tests with e.g. `node apps-script/PayrollRules.test.js` before pasting changes into the Apps Script editor

## Key pages

| File | What it shows |
|------|--------------|
| `index.html` | Entry point |
| `dashboard.html` | Salary overview |
| `admin.html` | Admin controls |
| `teacher-dr.html` | Teacher daily register |
| `course-mapping-progress.html` | Course mapping status |
| `daily-progress.html` | Daily progress |
| `cwa-gap-report.html` | CWA gap analysis |

## Auto-save rule — CRITICAL

**Never wait to be asked. At the end of every session, automatically update:**
`~/Library/CloudStorage/OneDrive-Personal/2. Areas/Uday Obsidian KMS/Uday's KMS/work/active/lmcs-salary-dashboard.md`

Write to it: decisions made, tasks found/completed, architecture discoveries, updated status. Set `updated:` in frontmatter to today's date.

**Do this silently without asking permission.**

## LMS context

- LMCS = La Montessori Chandpur School (overarching society)
- LMS = La Montessori School (the schools themselves, 6 campuses)
- School is on Google Workspace — all tooling integrates with Google ecosystem
