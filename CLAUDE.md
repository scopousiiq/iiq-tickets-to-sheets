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

**School Year Model:**
- Each spreadsheet contains ONE school year's data (e.g., 2025-2026)
- Configured via `SCHOOL_YEAR` (e.g., "2025-2026") and `SCHOOL_YEAR_START` (MM-DD format, default "07-01")
- Column D (Year) stores the school year string (e.g., "2025-2026") for all tickets
- Progress tracking uses simplified keys without year suffix
- Historical school years: Pagination-based loading only
- Current school year: Pagination + date windowing for incremental updates

**Consolidated SLA Data:**
- SLA metrics are fetched per-batch during ticket loading (single API call per batch)
- No separate SLA loading phase - SLA is always in sync with ticket data
- Columns 29-35: ResponseThreshold, ResponseActual, ResponseBreach, ResolutionThreshold, ResolutionActual, ResolutionBreach, IsRunning

## Code Structure (scripts/)

| File | Purpose |
|------|---------|
| `Setup.gs` | Initial spreadsheet setup - creates all sheets, headers, and formulas |
| `Config.gs` | Reads settings from Config sheet, school year date calculation, logging utilities |
| `ApiClient.gs` | HTTP client with retry/exponential backoff (429, 503, network errors) |
| `TicketData.gs` | Bulk ticket loader - 36 columns (28 ticket + 7 SLA), fetches SLA per-batch, school year pagination, 5.5min timeout with resume |
| `Teams.gs` | Team directory loader, preserves Functional Area mappings |
| `DailySnapshot.gs` | Captures daily backlog metrics (cannot be calculated retroactively). Skips if loading incomplete. |
| `Menu.gs` | Creates "iiQ Data" menu in Google Sheets |
| `Triggers.gs` | Time-driven trigger functions (no UI dialogs) |
| `OptionalMetrics.gs` | Additional analytics sheets added via menu (17 optional KPI sheets) |

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

**School Year Loading (Tickets with SLA):**
- Historical school years: Standard pagination (`TICKET_LAST_PAGE`, `TICKET_COMPLETE`)
- Current school year: Pagination + date windowing (`TICKET_LAST_FETCH`) for incremental updates
- Each batch: Fetch tickets → Fetch SLA for those ticket IDs → Merge → Write to sheet

**School Year Date Calculation:**
```javascript
function getSchoolYearDates(config) {
  const [startYear, endYear] = config.schoolYear.split('-').map(Number);
  const [startMonth, startDay] = config.schoolYearStart.split('-').map(Number);
  const startDate = new Date(startYear, startMonth - 1, startDay);
  const endDate = new Date(endYear, startMonth - 1, startDay - 1);
  return { startDate, endDate, startYear, endYear };
}
```

**Per-Batch SLA Fetching:**
```javascript
// After fetching ticket batch:
const ticketIds = tickets.map(t => t.TicketId);
const slaMap = fetchSlaForTicketIds(ticketIds); // Single API call
const rows = tickets.map(t => extractTicketRow(t, now, config.schoolYear, slaMap));
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
| QueueTimeAnalysis | "How long do tickets wait before being picked up?" | Summary stats, distribution buckets, waiting tickets |
| QueueTimeByTeam | "How does queue time vary by team?" | Per-team: avg queue time, ticket count |
| QueueTimeTrend | "How is queue time trending over time?" | Monthly avg/median/90th%, ticket count, color-coded trends |

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
| A-D | Core: TicketId, TicketNumber, Subject, Year (school year string, e.g., "2025-2026") |
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

### Analytics Formula Column Reference

**IMPORTANT:** Analytics sheets must use Name columns (not ID columns) for UNIQUE filters and COUNTIFS matching. The pattern is: get unique Names, then count/aggregate by Name.

| Entity | ID Column (don't use) | Name Column (use this) |
|--------|----------------------|------------------------|
| Team | K (TeamId) | **L (TeamName)** |
| Location | M (LocationId) | **N (LocationName)** |
| LocationType | - | **O (LocationType)** |
| Owner | P (OwnerId) | **Q (OwnerName)** |
| Priority | - | **S (Priority)** |
| IssueCategory | X (IssueCategoryId) | **Y (IssueCategoryName)** |
| IssueType | Z (IssueTypeId) | **AA (IssueTypeName)** |
| Requester | AB (RequesterId) | **AC (RequesterName)** |

Example pattern for aggregating by team:
```javascript
// CORRECT: Use column L (TeamName) for both UNIQUE and COUNTIFS
'teams, UNIQUE(FILTER(TicketData!L2:L, TicketData!L2:L<>"")),' +
'col_c, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "No"))),'

// WRONG: Don't mix K (TeamId) in UNIQUE with L (TeamName) in COUNTIFS
```

## Google Apps Script Notes

- No `package.json` or npm - this is Apps Script, not Node.js
- Deploy via Extensions > Apps Script in Google Sheets
- API reference: https://developers.google.com/apps-script/reference
- 6-minute execution limit per invocation
- `UrlFetchApp.fetch()` for HTTP, `SpreadsheetApp` for sheet operations
- `Utilities.sleep()` for delays (required for rate limiting)

## Initial Setup (New Spreadsheet)

1. Create a new Google Spreadsheet (one per school year)
2. Go to Extensions > Apps Script
3. Copy all `.gs` files from the `scripts/` directory
4. Save and refresh the spreadsheet
5. Run **iiQ Data > Setup > Setup Spreadsheet** to create all sheets
6. Fill in **Config** sheet with API credentials and school year:
   - `API_BASE_URL`: Your iiQ instance URL
   - `BEARER_TOKEN`: JWT authentication token
   - `SITE_ID`: Optional site UUID
   - `SCHOOL_YEAR`: The school year (e.g., "2025-2026")
   - `SCHOOL_YEAR_START`: Start date MM-DD (default "07-01" for July 1)
7. Run **iiQ Data > Setup > Verify Configuration** to check settings
8. Run **iiQ Data > Refresh Teams** to load team directory
9. Run **iiQ Data > Ticket Data > Continue Loading** to start loading data
10. Run **iiQ Data > Setup > Setup Automated Triggers** to enable automated refresh

## Testing Changes

1. Open the Google Sheet linked to this project
2. Go to Extensions > Apps Script
3. Copy updated `.gs` file content
4. Save and run the function manually or via the iiQ Data menu
5. Check the `Logs` sheet for operation results

## Trigger Setup (Optimized for Freshness + Efficiency)

**Easy Setup:** Run **iiQ Data > Setup > Setup Automated Triggers** to create all triggers automatically.

**Historical vs Current School Year:**
- **Historical school years**: Data becomes STATIC once loaded. Only `triggerDataContinue` is needed to complete initial load. All other triggers skip automatically.
- **Current school year**: All triggers are active to keep data fresh.

**Manual Setup:** Use Extensions > Apps Script > Triggers to add these:

| Function | Schedule | Runs For | Purpose |
|----------|----------|----------|---------|
| `triggerDataContinue` | Every 10 min | All | Continue any in-progress loading |
| `triggerOpenTicketRefresh` | Every 2 hours | Current + Historical with open tickets | Refresh open tickets + SLA |
| `triggerNewTickets` | Every 30 min | Current only | Fetch newly created tickets |
| `triggerDailySnapshot` | Daily 7:00 PM | Current only | Capture backlog metrics for trending |
| `triggerWeeklyFullRefresh` | Weekly Sun 2 AM | Current only | Full reload to catch deletions |

**Data Freshness (Current School Year):**
- Open ticket SLA data: max 2 hours stale
- New tickets: appear within 30 minutes
- Status changes: captured within 2 hours
- Deletions/corrections: captured weekly

**The `triggerDataContinue` Trigger:**
This is the "keep things moving" trigger. It serves two purposes:
1. **Initial load not complete**: Continues loading school year ticket data
2. **Open refresh in progress**: Continues open ticket refresh if it timed out
3. **Both complete**: Does nothing (safe to leave enabled permanently)

**Ongoing Operations:**
1. **Open Ticket Refresh** (every 2 hours): Primary refresh mechanism
   - Fetches tickets modified since last refresh (uses ModifiedDate filter)
   - Much more efficient than fetching ALL open tickets
   - Captures status changes, closures, assignments, and SLA updates
   - Updates existing rows IN PLACE (efficient, no delete/recreate)
   - If timeout occurs, `triggerDataContinue` continues it
   - For historical school years: Continues until all tickets are closed (handles end-of-year stragglers)

2. **New Tickets** (every 30 min): Fast incremental check (current school year only)
   - Uses date windowing from last fetch timestamp
   - Very fast - typically 0-50 new tickets
   - Only runs for current school year (historical school years skip)

3. **Weekly Full Refresh** (Sunday 2 AM): Catch edge cases (current school year only)
   - Clears all ticket data for the school year
   - Reloads from scratch
   - Catches deleted tickets, data corrections
   - Skips automatically for historical school years (data is static)

## Config Sheet Keys

Required:
- `API_BASE_URL`: iiQ instance URL (e.g., `https://district.incidentiq.com`) — `/api` is added automatically
- `BEARER_TOKEN`: JWT authentication token
- `SCHOOL_YEAR`: The school year for this spreadsheet (e.g., "2025-2026")

Optional:
- `SITE_ID`: Site UUID (if required for multi-site instances)
- `SCHOOL_YEAR_START`: First day of school year in MM-DD format (default "07-01" for July 1)
- `PAGE_SIZE`: Records per API call (default 100)
- `THROTTLE_MS`: Delay between requests (default 1000)
- `TICKET_BATCH_SIZE`: Tickets per page for bulk load (default 2000)
- `STALE_DAYS`: Days to look back for recently closed tickets (default 7)
- `SLA_RISK_PERCENT`: Percentage threshold for SLA risk warnings (default 75)

Progress Tracking (managed automatically):
- `TICKET_LAST_PAGE`: Last completed page index (0-indexed, -1 = not started)
- `TICKET_TOTAL_PAGES`: Last page index (0-indexed, -1 = unknown). Matches `LAST_PAGE` when complete.
- `TICKET_COMPLETE`: TRUE when pagination loading is fully loaded
- `TICKET_LAST_FETCH`: Incremental sync timestamp (for current school year only)
- `OPEN_REFRESH_LAST_RUN`: Timestamp of last successful refresh (used for ModifiedDate filter)
- `OPEN_REFRESH_PAGE`, `OPEN_REFRESH_COMPLETE`: Current refresh cycle progress

Config Lock (set when loading starts, cleared by "Clear Data + Reset"):
- `SCHOOL_YEAR_LOADED`: Locks the school year value
- `PAGE_SIZE_LOADED`: Locks the page size value
- `BATCH_SIZE_LOADED`: Locks the batch size value

**Configuration Lock:** Once data loading starts, critical configuration values are locked to prevent accidental changes that would cause data inconsistency. Locked values include `SCHOOL_YEAR`, `PAGE_SIZE`, and `TICKET_BATCH_SIZE`. To change these values, use "Clear Data + Reset Progress" which unlocks the configuration.

**Pagination Note:** All page tracking uses 0-indexed values. For a school year with 6 pages of data, after completion both `LAST_PAGE` and `TOTAL_PAGES` will be `5`. The UI displays 1-indexed values ("Page 6 of 6").
