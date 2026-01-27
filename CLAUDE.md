# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script project for extracting Incident IQ (iiQ) ticket and SLA data into Google Sheets. The data is used for Power BI consumption and analytics dashboards. This is not a Node.js project - it's Google Apps Script (.gs files) that runs within Google Sheets.

## Architecture

```
iiQ API  →  Google Apps Script  →  Google Sheets  →  Power BI
           (scripts/*.gs)         (formula-based     (consumption)
                                   analytics)
```

**Data Flow:**
1. Scripts fetch data from iiQ API using Bearer token authentication
2. Raw data lands in `TicketData` sheet (36 columns including consolidated SLA metrics)
3. Analytics sheets (`MonthlyVolume`, `BacklogAging`, `TeamWorkload`, etc.) calculate via Google Sheets formulas - no scripts needed
4. Power BI connects to Google Sheets for dashboards

**Year Discovery (Tickets):**
- Years are auto-discovered from Config sheet keys (no code changes needed)
- Historical years: `TICKET_{YEAR}_LAST_PAGE` rows → pagination-based loading
- Current year: `TICKET_{YEAR}_LAST_FETCH` row → date windowing for incremental updates

**Consolidated SLA Data:**
- SLA metrics are fetched per-batch during ticket loading (single API call per batch)
- No separate SLA loading phase - SLA is always in sync with ticket data
- Columns 29-35: ResponseThreshold, ResponseActual, ResponseBreach, ResolutionThreshold, ResolutionActual, ResolutionBreach, IsRunning

## Code Structure (scripts/)

| File | Purpose |
|------|---------|
| `Setup.gs` | Initial spreadsheet setup - creates all sheets, headers, and formulas |
| `Config.gs` | Reads settings from Config sheet, year discovery via regex, logging utilities |
| `ApiClient.gs` | HTTP client with retry/exponential backoff (429, 503, network errors) |
| `TicketData.gs` | Bulk ticket loader - 36 columns (28 ticket + 7 SLA), fetches SLA per-batch, year-based pagination, 5.5min timeout with resume |
| `Teams.gs` | Team directory loader, preserves Functional Area mappings |
| `DailySnapshot.gs` | Captures daily backlog metrics (cannot be calculated retroactively). Skips if loading incomplete. |
| `Menu.gs` | Creates "iiQ Data" menu in Google Sheets |
| `Triggers.gs` | Time-driven trigger functions (no UI dialogs) |
| `OptionalMetrics.gs` | Additional analytics sheets added via menu (14 optional KPI sheets) |

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

**Year-Based Loading (Tickets with SLA):**
- Historical years: Standard pagination (`TICKET_{YEAR}_LAST_PAGE`)
- Current year: Date windowing (`TICKET_{YEAR}_LAST_FETCH`) for incremental updates
- Each batch: Fetch tickets → Fetch SLA for those ticket IDs → Merge → Write to sheet

**Per-Batch SLA Fetching:**
```javascript
// After fetching ticket batch:
const ticketIds = tickets.map(t => t.TicketId);
const slaMap = fetchSlaForTicketIds(ticketIds); // Single API call
const rows = tickets.map(t => extractTicketRow(t, now, year, slaMap));
```

**Rate Limiting:**
- Exponential backoff: 2s base, doubles on retry (up to 3 retries)
- Configurable throttle via `THROTTLE_MS` in Config sheet

## Sheets Overview

### Data Sheets (Required)
| Sheet | Type | Purpose |
|-------|------|---------|
| Instructions | Static | Setup and usage guide |
| Config | Manual | API settings, progress tracking |
| TicketData | Data | Main ticket data (36 columns with SLA) |
| Teams | Data | Team directory with Functional Area mapping |
| DailySnapshot | Data | Daily backlog metrics for trending |
| Logs | Data | Operation logs |

### Default Analytics Sheets (Created by Setup)
| Sheet | Type | Purpose |
|-------|------|---------|
| MonthlyVolume | Formula | Ticket creation/closure by month |
| BacklogAging | Formula | Current open ticket distribution by age |
| TeamWorkload | Formula | Open tickets and aging by team |
| SLACompliance | Formula | Monthly SLA breach rates and avg response/resolution times |
| PerformanceTrends | Formula | "Are we getting better?" trending metrics |
| AtRiskResponse | Formula | Tickets approaching Response SLA breach |
| AtRiskResolution | Formula | Tickets approaching Resolution SLA breach |

## Analytics Sheets (All Available)

All analytics sheets can be added/recreated via **iiQ Data > Add Analytics Sheet** menu. Each is formula-based and auto-calculates from TicketData. Default sheets marked with ★.

### Volume & Trends
| Sheet | Question Answered | Key Metrics |
|-------|-------------------|-------------|
| MonthlyVolume ★ | "Is our backlog growing or shrinking?" | Created/Closed by month, net change, closure rate |
| PerformanceTrends ★ | "Are we getting better over time?" | Monthly trending of resolution time, breach rate, backlog |
| SeasonalComparison | "How does this year compare to last year?" | YoY change by month, identifies seasonal patterns |
| TemporalPatterns | "When do tickets come in?" | Day-of-week and hour-of-day distributions |

### Backlog & Quality
| Sheet | Question Answered | Key Metrics |
|-------|-------------------|-------------|
| BacklogAging ★ | "How old are our open tickets?" | Distribution by age bucket (0-15, 16-30, 31-60, 61-90, 90+) |
| StaleTickets | "Which tickets have no recent activity?" | Open tickets with no update in X days |
| ReopenRate | "Are we truly resolving issues?" | Reopened ticket detection and quality indicators |

### SLA & Response
| Sheet | Question Answered | Key Metrics |
|-------|-------------------|-------------|
| SLACompliance ★ | "Are we meeting our SLAs?" | Monthly breach rates, avg response/resolution times |
| AtRiskResponse ★ | "Which tickets are about to breach Response SLA?" | Tickets at 75%+ of Response threshold |
| AtRiskResolution ★ | "Which tickets are about to breach Resolution SLA?" | Tickets at 75%+ of Resolution threshold |
| FirstContactResolution | "How many tickets resolved same-day?" | Same-day %, 4-hour resolution %, monthly trend |
| ResponseDistribution | "How consistent is our response time?" | Percentiles, distribution buckets, consistency metrics |
| ResponseTrends | "Are response times improving over time?" | Monthly avg/median/90th%, % meeting targets, consistency |
| QueueTimeAnalysis | "How long do tickets wait before being picked up?" | Queue time by team, distribution buckets, waiting tickets |

### Team & Staff
| Sheet | Question Answered | Key Metrics |
|-------|-------------------|-------------|
| TeamWorkload ★ | "How is work distributed across teams?" | Per-team: open, created, closed, aged 30+ |
| TechnicianPerformance | "How is workload distributed among staff?" | Per-technician: open, closed, aged, breach rate |
| FunctionalAreaSummary | "How are functional areas performing?" | Aggregated metrics by functional area |

### Location
| Sheet | Question Answered | Key Metrics |
|-------|-------------------|-------------|
| LocationBreakdown | "Which locations generate the most tickets?" | Per-location: open, created, closed |
| LocationTypeComparison | "Which school types generate most tickets?" | Metrics by location type (Elementary/Middle/High) |

### Issue & Requester
| Sheet | Question Answered | Key Metrics |
|-------|-------------------|-------------|
| IssueCategoryVolume | "What types of problems are we handling?" | Open/Closed by category, breach rate per category |
| PriorityAnalysis | "Are high-priority tickets handled faster?" | Metrics by priority level, response/resolution times |
| FrequentRequesters | "Who generates the most tickets?" | Top 50 requesters with category and resolution data |

## TicketData Column Layout (36 columns)

| Columns | Description |
|---------|-------------|
| A-D | Core: TicketId, TicketNumber, Subject, Year |
| E-I | Dates: CreatedDate, StartedDate, ModifiedDate, ClosedDate, IsClosed |
| J | Status (WorkflowStep) |
| K-L | Team: TeamId, TeamName |
| M-O | Location: LocationId, LocationName, LocationType |
| P-Q | Owner: OwnerId, OwnerName |
| R | AgeDays |
| S-U | Priority: Priority, IsPastDue, DueDate |
| V-W | SLA (basic): SlaId, SlaName |
| X-AA | Issue: IssueCategoryId, IssueCategoryName, IssueTypeId, IssueTypeName |
| AB-AC | Requester: RequesterId, RequesterName |
| AD-AJ | SLA Metrics: ResponseThreshold, ResponseActual, ResponseBreach, ResolutionThreshold, ResolutionActual, ResolutionBreach, IsRunning |

## Google Apps Script Notes

- No `package.json` or npm - this is Apps Script, not Node.js
- Deploy via Extensions > Apps Script in Google Sheets
- API reference: https://developers.google.com/apps-script/reference
- 6-minute execution limit per invocation
- `UrlFetchApp.fetch()` for HTTP, `SpreadsheetApp` for sheet operations
- `Utilities.sleep()` for delays (required for rate limiting)

## Initial Setup (New Spreadsheet)

1. Create a new Google Spreadsheet
2. Go to Extensions > Apps Script
3. Copy all `.gs` files from the `scripts/` directory
4. Save and refresh the spreadsheet
5. Run **iiQ Data > Setup > Setup Spreadsheet** to create all sheets
6. Fill in **Config** sheet with API credentials:
   - `API_BASE_URL`: Your iiQ instance URL
   - `BEARER_TOKEN`: JWT authentication token
   - `SITE_ID`: Optional site UUID
7. Configure year tracking rows (e.g., `TICKET_2024_LAST_PAGE`, `TICKET_2025_LAST_FETCH`)
8. Run **iiQ Data > Setup > Verify Configuration** to check settings
9. Run **iiQ Data > Refresh Teams** to load team directory
10. Run **iiQ Data > Ticket Data > Continue Loading** to start loading data
11. Run **iiQ Data > Setup > Setup Automated Triggers** to enable automated refresh

## Testing Changes

1. Open the Google Sheet linked to this project
2. Go to Extensions > Apps Script
3. Copy updated `.gs` file content
4. Save and run the function manually or via the iiQ Data menu
5. Check the `Logs` sheet for operation results

## Trigger Setup (Optimized for Freshness + Efficiency)

**Easy Setup:** Run **iiQ Data > Setup > Setup Automated Triggers** to create all triggers automatically.

**Manual Setup:** Use Extensions > Apps Script > Triggers to add these:

| Function | Schedule | Purpose |
|----------|----------|---------|
| `triggerDataContinue` | Every 10 min | Continue any in-progress loading (initial OR open refresh) |
| `triggerOpenTicketRefresh` | Every 2 hours | Start open ticket + SLA refresh |
| `triggerNewTickets` | Every 30 min | Fetch newly created tickets (fast incremental check) |
| `triggerDailySnapshot` | Daily 7:00 PM | Capture backlog metrics for trending |
| `triggerWeeklyFullRefresh` | Weekly Sun 2 AM | Full reload to catch deletions/corrections |

**Data Freshness:**
- Open ticket SLA data: max 2 hours stale
- New tickets: appear within 30 minutes
- Status changes: captured within 2 hours
- Deletions/corrections: captured weekly

**The `triggerDataContinue` Trigger:**
This is the "keep things moving" trigger. It serves two purposes:
1. **Initial load not complete**: Continues loading historical ticket data
2. **Open refresh in progress**: Continues open ticket refresh if it timed out
3. **Both complete**: Does nothing (safe to leave enabled permanently)

**Ongoing Operations:**
1. **Open Ticket Refresh** (every 2 hours): Primary refresh mechanism
   - Fetches all open tickets with fresh SLA metrics
   - Fetches tickets closed in last 7 days (STALE_DAYS config)
   - Updates existing rows IN PLACE (efficient, no delete/recreate)
   - If timeout occurs, `triggerDataContinue` continues it

2. **New Tickets** (every 30 min): Fast incremental check
   - Uses date windowing from last fetch timestamp
   - Very fast - typically 0-50 new tickets

3. **Weekly Full Refresh** (Sunday 2 AM): Catch edge cases
   - Preserves historical data (2+ years old)
   - Clears and reloads recent years
   - Catches deleted tickets, data corrections

**Historical Data Preservation:**
- Weekly refresh preserves data from (currentYear - 2) and older
- Example: In 2026, data from 2024 is preserved; only 2025-2026 reload
- Uses efficient filter-and-rewrite (not row deletion)

## Config Sheet Keys

Required:
- `API_BASE_URL`: iiQ instance URL (e.g., `https://district.incidentiq.com`) — `/api` is added automatically
- `BEARER_TOKEN`: JWT authentication token
- `SITE_ID`: Site UUID (if required)

Optional:
- `PAGE_SIZE`: Records per API call (default 100)
- `THROTTLE_MS`: Delay between requests (default 1000)
- `TICKET_BATCH_SIZE`: Tickets per page for bulk load (default 2000)
- `STALE_DAYS`: Days to look back for recently closed tickets (default 7)
- `SLA_RISK_PERCENT`: Percentage threshold for SLA risk warnings (default 75)

Progress Tracking (managed automatically):
- `TICKET_{YEAR}_LAST_PAGE`: Last completed page index (0-indexed, -1 = not started)
- `TICKET_{YEAR}_TOTAL_PAGES`: Last page index (0-indexed, -1 = unknown). Matches `LAST_PAGE` when complete.
- `TICKET_{YEAR}_COMPLETE`: TRUE when year is fully loaded
- `TICKET_{YEAR}_LAST_FETCH`: Current year incremental sync timestamp
- `OPEN_REFRESH_DATE`: Date of current open ticket refresh
- `OPEN_REFRESH_OPEN_PAGE`, `OPEN_REFRESH_OPEN_COMPLETE`: Open ticket refresh progress
- `OPEN_REFRESH_CLOSED_PAGE`, `OPEN_REFRESH_CLOSED_COMPLETE`: Recently closed refresh progress

**Pagination Note:** All page tracking uses 0-indexed values. For a year with 6 pages of data, after completion both `LAST_PAGE` and `TOTAL_PAGES` will be `5`. The UI displays 1-indexed values ("Page 6 of 6").
