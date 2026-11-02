# HR Attendance — Time-Clock Report Analyzer

A browser app for HR managers: upload the Excel export from a fingerprint / badge time clock and instantly get each employee's worked hours, absences, overtime and missing hours for the period — then export a clean PDF report.

**Live demo:** [hr-attendance-nu.vercel.app](https://hr-attendance-nu.vercel.app)

## Features

- **Excel import** — reads `.xlsx` / `.xls` time-clock exports (employee ID, name, department, date, punch times). Everything is processed locally in the browser; no data is uploaded.
- **Punch cleanup** — merges duplicate punches recorded within 5 minutes of each other and flags days with an odd number of punches as errors.
- **Daily analysis** — each day is classified as *Worked*, *Absent*, *Day off*, *Day off + worked* or *Error*, with worked time rounded to the quarter-hour.
- **Configurable schedule** — set weekly days off, extra days off (holidays) and exception periods with different required hours (e.g. Ramadan).
- **Manual overrides** — mark any day as a full or half day to correct the calculation.
- **Summary dashboard** — worked hours, days worked, absences, overtime, missing hours and net balance (credit / debit).
- **PDF export** — one-click printable attendance report per employee.

## Tech stack

- React 19 + TypeScript
- Vite
- [SheetJS (xlsx)](https://sheetjs.com/) for Excel parsing
- jsPDF for PDF generation

## Run locally

```bash
npm install
npm run dev
```

Then open the URL shown in the terminal (usually http://localhost:5173).

```bash
npm run build   # production build in dist/
```
