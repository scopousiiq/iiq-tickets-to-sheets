# IIQ Tickets to Sheets

Extract IncidentIQ ticket and SLA data into Google Sheets for analytics and Power BI dashboards. Built for K-12 school districts.

## Quick Start

### 1. Copy the Template

[**Make a copy of the Google Sheets template**](https://docs.google.com/spreadsheets/d/1azxyYEefeAJ-oSmTKb199yzdULaaPvi_pjUydiC-fuA/edit?usp=sharing)

### 2. Configure API Access

In the `Config` sheet, enter your IncidentIQ credentials:

| Setting | Value |
|---------|-------|
| `API_BASE_URL` | `https://yourdistrict.incidentiq.com/api` |
| `BEARER_TOKEN` | Your API token |
| `SITE_ID` | Your site UUID (if required) |

### 3. Load Your Data

1. Reload the spreadsheet to activate the **IIQ Data** menu
2. Click **IIQ Data > Refresh Teams** and authorize when prompted
3. Click **IIQ Data > Ticket Data > Continue Loading** to start importing tickets
4. Click **IIQ Data > SLA Data > Continue Loading** to import SLA data

> **Note:** Large datasets may require multiple runs. The loader saves progress and resumes automatically.

### 4. Set Up Automated Refresh (Optional)

In Apps Script (**Extensions > Apps Script > Triggers**), add these daily triggers:

| Function | Time | Purpose |
|----------|------|---------|
| `triggerTicketDataUpdate` | 5:00 AM | Sync new tickets |
| `triggerSlaDataUpdate` | 5:30 AM | Sync SLA data |
| `triggerDailySnapshot` | 7:00 PM | Capture backlog metrics |

## What You Get

| Sheet | Description |
|-------|-------------|
| `TicketData` | Raw ticket data (28 columns) for analysis |
| `TicketSlaData` | SLA timing metrics (response/resolution times) |
| `Teams` | Team directory with Functional Area mapping |
| `MonthlyVolume` | Created/Closed/Open tickets by month |
| `BacklogAging` | Open tickets by age bucket (0-15, 16-30, 31-60, 61-90, 90+ days) |
| `TeamWorkload` | Per-team metrics with sorting |
| `SLACompliance` | Breach rates and average response/resolution times |
| `DailySnapshot` | Historical backlog metrics for trending |

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "API configuration missing" | Fill in `API_BASE_URL` and `BEARER_TOKEN` in Config sheet |
| HTTP 401 | Refresh your Bearer token |
| HTTP 429 / Rate limited | Increase `THROTTLE_MS` in Config (try 2000+) |
| Script timeout | Normal for large datasets - run "Continue Loading" again |

## Documentation

- [**Implementation Guide**](GUIDE.md) - Detailed sheet setup, formulas, and architecture
- [**CLAUDE.md**](CLAUDE.md) - Developer reference for AI-assisted development

## License

MIT
