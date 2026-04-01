# MEG Software Cost Allocation Tool

A web-based tool for the MEG Finance team to allocate software costs across cost centers. Supports any software product (Adobe, Dropbox, Zoom, Slack, etc.) with two allocation methods used in Workday.

## Allocation Methods

### Fixed Dollar Allocation
Charge each cost center based on per-user license costs. Used when users have different license tiers with known monthly rates (e.g., Adobe Creative Cloud vs. Adobe Stock).

- Upload or manually define a **rate card** with monthly cost per license type
- Supports multiple licenses per user (delimiter-configurable, e.g., semicolon-separated)
- Outputs a cost-per-cost-center summary for accounting entry

### Fixed % Allocation
Allocate a percentage of monthly amortized invoice cost to each cost center based on active user count. Used when all users share the same subscription (e.g., Dropbox).

- Supports **multiple invoices** with different amounts and service periods
- Calculates combined monthly amortized cost automatically
- Outputs percentage allocation per cost center for quarterly Workday updates

## How It Works

1. **Upload** - Upload the HR Employee List (.xlsx with emails and cost centers) and the Software User List (.xlsx from IT)
2. **Configure** - Select the allocation method, map columns, and set rates or invoice details
3. **Review** - See match statistics and assign unmatched users to cost centers via a searchable dropdown
4. **Results** - View the allocation summary and download an Excel workbook for accounting

### Handling Unmatched Users

Users not found in the HR file are flagged for manual assignment. A **searchable cost center picker** lets you quickly find and select from existing HR cost centers instead of typing IDs manually. IT-provided cost centers are shown as reference but not used automatically (they tend to be stale).

## Excel Export

**Fixed Dollar** produces:
- AllocEntry for Accounting (pivot by cost center)
- User Detail (per-user license breakdown)
- Rate Sheet (reference)
- Flagged - Unmatched (if any)

**Fixed %** produces:
- Monthly Allocation (with invoice breakdown and per-CC percentages)
- Active Users (roster with cost center assignments)
- Flagged - Unmatched (if any)

## Getting Started

```bash
npm install
npm run dev
```

Open [http://localhost:5173](http://localhost:5173) in your browser.

### Build for Production

```bash
npm run build
npm run preview
```

## Tech Stack

- **React** with Vite
- **xlsx** for Excel parsing and generation
- **lodash** for data grouping
- Single-page app, no backend required - all processing happens in the browser
