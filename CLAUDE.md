# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script project for extracting IncidentIQ (IIQ) ticket and SLA data into Google Sheets. The data is used for Power BI consumption and analytics dashboards. This is not a Node.js project - it's Google Apps Script (.gs files) that runs within Google Sheets.

## Architecture

```
IIQ API  →  Google Apps Script  →  Google Sheets  →  Power BI
           (scripts/*.gs)         (formula-based     (consumption)
                                   analytics)
```

**Data Flow:**
1. Scripts fetch data from IIQ API using Bearer token authentication
2. Raw data lands in `TicketData` (28 columns) and `TicketSlaData` (11 columns) sheets
3. Analytics sheets (`MonthlyVolume`, `BacklogAging`, `TeamWorkload`, etc.) calculate via Google Sheets formulas - no scripts needed
4. Power BI connects to Google Sheets for dashboards

**Year Discovery:**
- Years are auto-discovered from Config sheet keys (no code changes needed)
- Historical years: `TICKET_{YEAR}_LAST_PAGE` rows → pagination-based loading
- Current year: `TICKET_{YEAR}_LAST_FETCH` row → date windowing for incremental updates

## Code Structure (scripts/)

| File | Purpose |
|------|---------|
| `Config.gs` | Reads settings from Config sheet, year discovery via regex, logging utilities |
| `ApiClient.gs` | HTTP client with retry/exponential backoff (429, 503, network errors) |
| `TicketData.gs` | Bulk ticket loader - 28 columns, year-based pagination, 5.5min timeout with resume |
| `TicketSlaData.gs` | SLA timing loader - matches TicketData pattern, API capped at 100/page |
| `Teams.gs` | Team directory loader, preserves Functional Area mappings |
| `DailySnapshot.gs` | Captures daily backlog metrics (cannot be calculated retroactively) |
| `Menu.gs` | Creates "IIQ Data" menu in Google Sheets |
| `Triggers.gs` | Time-driven trigger functions (no UI dialogs) |

**Key Dependencies:**
- `ApiClient.gs` → `Config.gs`
- All data loaders → `ApiClient.gs`
- `DailySnapshot.gs` reads from `TicketData` sheet (not API)

## Key Patterns

**Pagination with Timeout Handling:**
```javascript
const MAX_RUNTIME_MS = 5.5 * 60 * 1000; // Stay under 6min Apps Script limit
while (Date.now() - startTime < MAX_RUNTIME_MS) {
  // Process one batch, save progress to Config sheet
  // User runs "Continue Loading" to resume
}
```

**Year-Based Loading:**
- Historical years: Standard pagination (`TICKET_{YEAR}_LAST_PAGE`)
- Current year: Date windowing (`TICKET_{YEAR}_LAST_FETCH`) for incremental updates

**Rate Limiting:**
- Exponential backoff: 2s base, doubles on retry (up to 3 retries)
- Configurable throttle via `THROTTLE_MS` in Config sheet

## Google Apps Script Notes

- No `package.json` or npm - this is Apps Script, not Node.js
- Deploy via Extensions > Apps Script in Google Sheets
- API reference: https://developers.google.com/apps-script/reference
- 6-minute execution limit per invocation
- `UrlFetchApp.fetch()` for HTTP, `SpreadsheetApp` for sheet operations
- `Utilities.sleep()` for delays (required for rate limiting)

## Testing Changes

1. Open the Google Sheet linked to this project
2. Go to Extensions > Apps Script
3. Copy updated `.gs` file content
4. Save and run the function manually or via the IIQ Data menu
5. Check the `Logs` sheet for operation results

## Trigger Setup (after initial load completes)

| Function | Schedule | Purpose |
|----------|----------|---------|
| `triggerTicketDataUpdate` | Daily 5:00 AM | Incremental ticket sync |
| `triggerSlaDataUpdate` | Daily 5:30 AM | Incremental SLA sync |
| `triggerDailySnapshot` | Daily 11:00 PM | Capture backlog metrics for trending |

For bulk loading: Use `triggerBulkLoadContinue` every 10 minutes until historical data is complete.

## Config Sheet Keys

Required:
- `API_BASE_URL`: IIQ instance URL (e.g., `https://district.incidentiq.com/api`)
- `BEARER_TOKEN`: JWT authentication token
- `SITE_ID`: Site UUID (if required)

Optional:
- `PAGE_SIZE`: Records per API call (default 100)
- `THROTTLE_MS`: Delay between requests (default 1000)
- `TICKET_BATCH_SIZE`: Tickets per page for bulk load (default 2000)
- `SLA_BATCH_SIZE`: SLA records per page (capped at 100 by API)
