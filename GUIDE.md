# Implementation Guide

This guide explains how the iiQ Tickets to Sheets system works — the sheet structure, formulas, and technical details. It's written for IT staff who want to understand or customize the system.

**Just want to get started?** See the [README](README.md) for quick setup instructions.

### Architecture

```
┌─────────────┐     ┌──────────────────┐     ┌─────────────┐
│  iiQ API    │────▶│  Google Sheets   │────▶│  Power BI   │
│             │     │  + Apps Script   │     │             │
│ /tickets    │     │                  │     │ Dashboards  │
│ /teams      │     │ • Config sheet   │     │             │
└─────────────┘     │ • Volume data    │     └─────────────┘
                    │ • Aging data     │
                    │ • Team data      │
                    └──────────────────┘
```

---

## Part 1: Google Sheets Setup

### Step 1: Create the Spreadsheet

Create a new Google Sheet. The **Setup Spreadsheet** function will create all required sheets automatically, or you can create them manually.

**Data Sheets (required):**

| Sheet Name | Purpose |
|------------|---------|
| `Config` | API credentials and settings |
| `TicketData` | Raw ticket dump with SLA metrics (36 columns) |
| `Teams` | Team directory with FA mapping |
| `DailySnapshot` | Daily backlog metrics for trending |
| `Logs` | API call logs and errors |

**Default Analytics Sheets (created by Setup):**

| Sheet Name | Purpose |
|------------|---------|
| `MonthlyVolume` | Created/Closed/Open by month |
| `BacklogAging` | Current open tickets by age bucket |
| `TeamWorkload` | Per-team metrics for FA rollup |
| `SLACompliance` | SLA breaches, response & resolution times |
| `PerformanceTrends` | "Are we getting better?" trending metrics |
| `AtRiskResponse` | Tickets approaching Response SLA threshold |
| `AtRiskResolution` | Tickets approaching Resolution SLA threshold |

**Optional Analytics Sheets (add via menu):**

Additional analytics sheets can be added via **iiQ Data > Add Analytics Sheet**. See the [Optional Analytics Sheets](#optional-analytics-sheets) section below for the full list.

### Step 2: Configure the Sheets

#### Sheet: `Config`

**Required Settings (you must fill these in):**

| Setting | Value | Description |
|---------|-------|-------------|
| `API_BASE_URL` | `https://your-district.incidentiq.com` | Your iiQ URL (the `/api` is added automatically) |
| `BEARER_TOKEN` | (your JWT token) | Get this from iiQ Admin > Integrations > API |
| `SITE_ID` | (your site UUID) | Only needed for multi-site districts |

**Optional Settings (defaults work for most districts):**

| Setting | Default | Description |
|---------|---------|-------------|
| `PAGE_SIZE` | 100 | Records per API call |
| `THROTTLE_MS` | 1000 | Milliseconds between API calls (increase if rate limited) |
| `STALE_DAYS` | 7 | Days without update before a ticket shows in StaleTickets |
| `SLA_RISK_PERCENT` | 75 | % of SLA used before ticket shows in AtRiskResponse/AtRiskResolution |
| `TICKET_BATCH_SIZE` | 2000 | Tickets per batch during bulk load |

**Progress Tracking (auto-managed, don't edit):**

| Setting | Purpose |
|---------|---------|
| `TICKET_2024_LAST_PAGE` | Tracks loading progress for 2024 tickets |
| `TICKET_2024_COMPLETE` | TRUE when 2024 is fully loaded |
| `TICKET_2025_LAST_PAGE` | Tracks loading progress for 2025 tickets |
| `TICKET_2025_COMPLETE` | TRUE when 2025 is fully loaded |
| `TICKET_2026_LAST_FETCH` | Timestamp for incremental 2026 loading |
| `LAST_REFRESH` | When data was last refreshed (shown on analytics sheets) |
| `OPEN_REFRESH_*` | Progress tracking for open ticket refresh |

> **Year Configuration:** Years are auto-discovered from these rows — no code changes needed.
> - **Historical years** (pagination-based): Detected from `TICKET_{YEAR}_LAST_PAGE` rows
> - **Current year** (date windowing): Detected from `TICKET_{YEAR}_LAST_FETCH` row
>
> **Consolidated SLA Data:**
> - SLA metrics are fetched per-batch during ticket loading (single API call per batch)
> - No separate SLA loading phase - SLA data is always in sync with ticket data
> - TicketData columns 30-36 (AD-AJ) contain: ResponseThreshold, ResponseActual, ResponseBreach, ResolutionThreshold, ResolutionActual, ResolutionBreach, IsRunning
> - Tickets without SLA policies will have blank values in these columns
>
> **To add a ticket year (e.g., 2023):**
> ```
> TICKET_2023_TOTAL_PAGES | (auto-populated)
> TICKET_2023_LAST_PAGE   | -1
> TICKET_2023_COMPLETE    | FALSE
> ```
>
> **To remove a ticket year:** Delete the corresponding `TICKET_{YEAR}_*` rows from Config, then optionally delete data rows from TicketData sheet.
>
> **To change current ticket year:** Replace `TICKET_2026_LAST_FETCH` with `TICKET_2027_LAST_FETCH`.
>
> **Important:** Format column B as **Plain text** (Format → Number → Plain text) to prevent Sheets from auto-formatting page counts as dates.

#### Sheet: `MonthlyVolume`

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| **Month** | **Year** | **Created** | **Closed** | **Net Change** | **Closure Rate** |
| January | 2026 | (formula) | (formula) | (formula) | (formula) |
| February | 2026 | (formula) | (formula) | (formula) | (formula) |
| ... | ... | | | | |

> **Note:** This sheet uses formulas to calculate from TicketData - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Column A & B:** Manually enter the months and years you want to track
3. **Cell C2:** Count tickets created in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!E:E, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!E:E, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")))
   ```
4. **Cell D2:** Count tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")))
   ```
5. **Cell E2:** Calculate net change (drag down)
   ```
   =C2-D2
   ```
6. **Cell F2:** Calculate closure rate (drag down, format as percentage)
   ```
   =IF(C2>0, D2/C2, "N/A")
   ```

> **Tip:** The LET function extracts the month number from the month name. The formula compares ISO date strings lexicographically, which works correctly for YYYY-MM-DD format.
>
> **Current open ticket count** is available in the BacklogAging sheet's TOTAL row.

#### Sheet: `BacklogAging`

| A | B | C | D | E |
|---|---|---|---|---|
| **Age Bucket** | **Count** | **% of Total** | **Sample Ticket** | **Last Refreshed** |
| 0-15 days | (formula) | (formula) | (formula) | (formula) |
| 16-30 days | (formula) | (formula) | (formula) | (formula) |
| 31-60 days | (formula) | (formula) | (formula) | (formula) |
| 61-90 days | (formula) | (formula) | (formula) | (formula) |
| 90+ days | (formula) | (formula) | (formula) | (formula) |
| **TOTAL** | (formula) | 100% | | |

> **Note:** This sheet uses formulas to calculate from TicketData - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Column A (Age Buckets):** Enter the bucket labels manually: `0-15 days`, `16-30 days`, `31-60 days`, `61-90 days`, `90+ days`, `TOTAL`

3. **Column B (Count):** Enter these formulas for each row:
   - **B2 (0-15 days):** `=COUNTIFS(TicketData!I:I, "No", TicketData!R:R, ">=0", TicketData!R:R, "<=15")`
   - **B3 (16-30 days):** `=COUNTIFS(TicketData!I:I, "No", TicketData!R:R, ">=16", TicketData!R:R, "<=30")`
   - **B4 (31-60 days):** `=COUNTIFS(TicketData!I:I, "No", TicketData!R:R, ">=31", TicketData!R:R, "<=60")`
   - **B5 (61-90 days):** `=COUNTIFS(TicketData!I:I, "No", TicketData!R:R, ">=61", TicketData!R:R, "<=90")`
   - **B6 (90+ days):** `=COUNTIFS(TicketData!I:I, "No", TicketData!R:R, ">90")`
   - **B7 (TOTAL):** `=SUM(B2:B6)`

4. **Column C (% of Total):** Enter for each row (format as percentage):
   - **C2-C6:** `=IF($B$7>0, B2/$B$7, 0)` (drag down for each bucket row)
   - **C7:** Enter `100%` or `=1`

5. **Column D (Sample Ticket):** Oldest ticket in each bucket:
   - **D2:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="No")*(TicketData!R2:R>=0)*(TicketData!R2:R<=15)),2,FALSE),1,1),"")`
   - **D3:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="No")*(TicketData!R2:R>=16)*(TicketData!R2:R<=30)),2,FALSE),1,1),"")`
   - **D4:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="No")*(TicketData!R2:R>=31)*(TicketData!R2:R<=60)),2,FALSE),1,1),"")`
   - **D5:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="No")*(TicketData!R2:R>=61)*(TicketData!R2:R<=90)),2,FALSE),1,1),"")`
   - **D6:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="No")*(TicketData!R2:R>90)),2,FALSE),1,1),"")`

6. **Column E (Last Refreshed):**
   - **E2:** `=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")` (drag down)

> **How it works:** Uses TicketData column I (IsClosed) and R (AgeDays) to count open tickets by age bucket. Sample Ticket formula filters, sorts by age descending, and returns the oldest ticket number in each bucket.

#### Sheet: `Teams`

| A | B | C | D |
|---|---|---|---|
| **Team ID** | **Team Name** | **Functional Area** | **Is Active** |
| (auto-populated) | (auto-populated) | (you fill this in) | (auto-populated) |

> **What is Functional Area?** This is your own label for grouping teams. For example, you might label Help Desk Tier 1, Help Desk Tier 2, and Field Services as "IT Support Services" — then the FunctionalAreaSummary sheet will show combined metrics for all those teams.
>
> After running **iiQ Data > Refresh Teams**, fill in column C with your Functional Area labels. You can paste values directly — there's no dropdown restriction.

#### Sheet: `TeamWorkload`

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| **Team Name** | **Functional Area** | **Open** | **Created (MTD)** | **Closed (MTD)** | **Aged 30+** | **Last Refreshed** | **Sort Col#** | **Desc?** |
| (formula - spills down) | | | | | | (formula) | 3 | FALSE |

> **Note:** This sheet uses a single array formula to calculate all data from TicketData. The entire table is sortable via H2 (column number 1-6) and I2 (FALSE=descending, TRUE=ascending).

**Setup Instructions:**

1. **Row 1 (Headers):** Enter column headers: Team Name, Functional Area, Open, Created (MTD), Closed (MTD), Aged 30+, Last Refreshed, Sort Col#, Desc?

2. **Cell H2:** Enter the sort column number (default: `3` for Open)
   - Options: `1` (Team), `2` (FA), `3` (Open), `4` (Created), `5` (Closed), `6` (Aged)

3. **Cell I2:** Enter sort order (default: `FALSE` for descending)
   - `FALSE` = Descending (highest first), `TRUE` = Ascending (lowest first)

4. **Cell A2:** Paste this formula — it generates the entire table automatically:
   ```
   =LET(
     teams, UNIQUE(FILTER(TicketData!L2:L, TicketData!L2:L<>"", TicketData!L2:L<>"TeamName")),
     mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),
     mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),
     col_a, teams,
     col_b, BYROW(teams, LAMBDA(t, IFERROR(VLOOKUP(t, Teams!B:C, 2, FALSE), ""))),
     col_c, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "No"))),
     col_d, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),
     col_e, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),
     col_f, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "No", TicketData!R:R, ">=30"))),
     data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),
     SORT(data, $H$2, $I$2)
   )
   ```

5. **Cell G2:** Show when data was last refreshed:
   ```
   =IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")
   ```

> **Prerequisite:** Populate the Teams sheet first (via script or manually) with TeamName in column B and Functional Area in column C.
>
> **How sorting works:** Set H2 to the column number (1=Team, 2=FA, 3=Open, 4=Created, 5=Closed, 6=Aged). Set I2 to FALSE for descending or TRUE for ascending. The table re-sorts automatically.

#### Sheet: `SLACompliance`

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| **Month** | **Year** | **Closed** | **Breaches** | **Breach Rate** | **Avg Response (hrs)** | **Avg Resolution (hrs)** |
| January | 2026 | (formula) | (formula) | (formula) | (formula) | (formula) |
| February | 2026 | (formula) | (formula) | (formula) | (formula) | (formula) |
| ... | ... | | | | | |

> **Note:** This sheet uses formulas to calculate from TicketData - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Column A & B:** Manually enter the months and years you want to track (or copy from MonthlyVolume)
3. **Cell C2:** Count tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!I:I, "Yes", TicketData!H:H, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")))
   ```
4. **Cell D2:** Count breaches (Response OR Resolution) for closed tickets that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), startDate, TEXT(DATE(B2,m,1), "YYYY-MM-DD"), endDate, TEXT(DATE(B2,m+1,1), "YYYY-MM-DD"), COUNTIFS(TicketData!I:I, "Yes", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AF:AF, TRUE) + COUNTIFS(TicketData!I:I, "Yes", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AI:AI, TRUE) - COUNTIFS(TicketData!I:I, "Yes", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AF:AF, TRUE, TicketData!AI:AI, TRUE))
   ```
5. **Cell E2:** Breach rate (drag down, format as percentage)
   ```
   =IF(C2>0, D2/C2, "N/A")
   ```
6. **Cell F2:** Avg Response time in hours (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), startDate, TEXT(DATE(B2,m,1), "YYYY-MM-DD"), endDate, TEXT(DATE(B2,m+1,1), "YYYY-MM-DD"), IFERROR(AVERAGEIFS(TicketData!AE:AE, TicketData!I:I, "Yes", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AE:AE, ">0")/60, "N/A"))
   ```
7. **Cell G2:** Avg Resolution time in hours (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), startDate, TEXT(DATE(B2,m,1), "YYYY-MM-DD"), endDate, TEXT(DATE(B2,m+1,1), "YYYY-MM-DD"), IFERROR(AVERAGEIFS(TicketData!AH:AH, TicketData!I:I, "Yes", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AH:AH, ">0")/60, "N/A"))
   ```

**Columns Explained:**
- **Closed**: Count of tickets closed in that month
- **Breaches**: Count where ResponseBreach=TRUE OR ResolutionBreach=TRUE (avoiding double-count)
- **Breach Rate**: Breaches / Closed
- **Avg Response (hrs)**: Average ResponseActual (column AE, converted from minutes to hours)
- **Avg Resolution (hrs)**: Average ResolutionActual (column AH, converted from minutes to hours)

> **Prerequisites:** TicketData must be loaded with SLA metrics (columns AD-AJ).
>
> **Tip:** Add more months by copying a row and changing the Month/Year values.

#### Sheet: `PerformanceTrends`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| **Month** | **Year** | **Closed** | **Avg Resolution (days)** | **Closure Rate** | **Breach Rate** | **Backlog EOM** | **% Aged 30+ EOM** |
| January | 2026 | (formula) | (formula) | (formula) | (formula) | (from snapshot) | (from snapshot) |
| February | 2026 | (formula) | (formula) | (formula) | (formula) | (from snapshot) | (from snapshot) |

> **Purpose:** This is your "executive dashboard" sheet — it answers the question every IT director gets asked: "Are we getting better?"
>
> **What the metrics mean:**
> - **Closure Rate** over 100% = closing more tickets than you're receiving (shrinking backlog)
> - **Avg Resolution** going down = solving problems faster
> - **Breach Rate** going down = meeting SLA commitments more often
> - **% Aged 30+** going down = not letting tickets sit and get stale

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Column A & B:** Manually enter the months and years you want to track

3. **Cell C2:** Count tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")))
   ```

4. **Cell D2:** Average resolution time in DAYS for tickets closed that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!I:I, "Yes", TicketData!H:H, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")), "N/A"))
   ```

5. **Cell E2:** Closure rate (Closed ÷ Created) - drag down, format as percentage
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), created, COUNTIFS(TicketData!E:E, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!E:E, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")), IF(created>0, C2/created, "N/A"))
   ```

6. **Cell F2:** Breach rate - reference SLACompliance or calculate (drag down, format as percentage)
   ```
   =IFERROR(VLOOKUP(A2&B2, ARRAYFORMULA(SLACompliance!A:A&SLACompliance!B:B&"|"&SLACompliance!E:E), 1, FALSE), "N/A")
   ```
   *Or simpler:* `=IFERROR(INDEX(SLACompliance!E:E, MATCH(A2&B2, SLACompliance!A:A&SLACompliance!B:B, 0)), "N/A")`

7. **Cell G2:** Backlog at end of month - lookup from DailySnapshot (drag down)
   ```
   =IFERROR(INDEX(DailySnapshot!B:B, MATCH(EOMONTH(DATE(B2, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), 1), 0), DailySnapshot!A:A, 0)), "No snapshot")
   ```

8. **Cell H2:** % Aged 30+ at end of month - lookup from DailySnapshot (drag down)
   ```
   =IFERROR(INDEX(DailySnapshot!D:D, MATCH(EOMONTH(DATE(B2, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), 1), 0), DailySnapshot!A:A, 0)), "No snapshot")
   ```

> **Interpreting Trends - "Getting Better" means:**
> - **Avg Resolution (days):** Trending DOWN = faster ticket resolution
> - **Closure Rate:** Trending UP toward or above 100% = keeping up with demand
> - **Breach Rate:** Trending DOWN = meeting SLA commitments
> - **Backlog EOM:** Stable or DOWN = not accumulating debt
> - **% Aged 30+ EOM:** Trending DOWN = addressing older tickets
>
> **Note:** Columns G and H require the DailySnapshot sheet to have data. Start the daily snapshot trigger to begin collecting historical data.

#### Sheet: `DailySnapshot`

| A | B | C | D |
|---|---|---|---|
| **Date** | **Open Tickets** | **Aged 30+ Count** | **% Aged 30+** |
| (auto-populated) | (auto-populated) | (auto-populated) | (auto-populated) |

> **Purpose:** Captures daily backlog metrics that cannot be calculated retroactively. Used by PerformanceTrends for end-of-month values.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Data rows:** Populated automatically by the `appendDailySnapshot` trigger function

> **Important:** Start the daily snapshot trigger as soon as possible to begin building historical data. You cannot calculate past backlog sizes from current data.

#### Sheet: `Logs`

| A | B | C | D |
|---|---|---|---|
| **Timestamp** | **Operation** | **Status** | **Details** |

#### Sheet: `LocationBreakdown`

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| **Location Name** | **Location Type** | **Open** | **Created (MTD)** | **Closed (MTD)** | **Last Refreshed** |
| (formula) | (formula) | (formula) | (formula) | (formula) | (formula) |

> **Note:** This sheet uses formulas to calculate from TicketData - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Cell A2:** Get unique locations from TicketData
   ```
   =UNIQUE(FILTER(TicketData!M2:M, TicketData!M2:M<>"", TicketData!M2:M<>"LocationName"))
   ```
3. **Cell B2:** Look up location type (drag down to match A column)
   ```
   =IFERROR(INDEX(TicketData!N:N, MATCH(A2, TicketData!M:M, 0)), "")
   ```
4. **Cell C2:** Count open tickets for location (drag down)
   ```
   =COUNTIFS(TicketData!M:M, A2, TicketData!I:I, "No")
   ```
5. **Cell D2:** Count tickets created this month (drag down)
   ```
   =COUNTIFS(TicketData!M:M, A2, TicketData!E:E, ">="&EOMONTH(TODAY(),-1)+1, TicketData!E:E, "<="&TODAY())
   ```
6. **Cell E2:** Count tickets closed this month (drag down)
   ```
   =COUNTIFS(TicketData!M:M, A2, TicketData!H:H, ">="&EOMONTH(TODAY(),-1)+1, TicketData!H:H, "<="&TODAY())
   ```
7. **Cell F2:** Reference last refresh from Config (drag down)
   ```
   =IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "Not yet refreshed")
   ```

> **Tip:** After entering A2, wait for the UNIQUE formula to populate, then drag B2:F2 down to cover all location rows.

#### Sheet: `FunctionalAreaSummary`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| **Functional Area** | **Teams** | **Open** | **Created (MTD)** | **Closed (MTD)** | **Aged 30+** | **% Aged 30+** | **Last Refreshed** |
| (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) |

> **Note:** This sheet uses formulas to aggregate from TeamWorkload - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Cell A2:** Get unique Functional Areas from TeamWorkload
   ```
   =UNIQUE(FILTER(TeamWorkload!B2:B, TeamWorkload!B2:B<>"", TeamWorkload!A2:A<>""))
   ```
3. **Cell B2:** Count teams in this FA (drag down)
   ```
   =COUNTIF(TeamWorkload!$B:$B, A2)
   ```
4. **Cell C2:** Sum open tickets for this FA (drag down)
   ```
   =SUMIF(TeamWorkload!$B:$B, A2, TeamWorkload!$C:$C)
   ```
5. **Cell D2:** Sum created MTD for this FA (drag down)
   ```
   =SUMIF(TeamWorkload!$B:$B, A2, TeamWorkload!$D:$D)
   ```
6. **Cell E2:** Sum closed MTD for this FA (drag down)
   ```
   =SUMIF(TeamWorkload!$B:$B, A2, TeamWorkload!$E:$E)
   ```
7. **Cell F2:** Sum aged 30+ for this FA (drag down)
   ```
   =SUMIF(TeamWorkload!$B:$B, A2, TeamWorkload!$F:$F)
   ```
8. **Cell G2:** Calculate % aged 30+ (drag down, format as percentage)
   ```
   =IF(C2>0, F2/C2, 0)
   ```
9. **Cell H2:** Reference last refresh from Config (drag down)
   ```
   =IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")
   ```

> **Handling Unassigned Teams:** If some teams have blank Functional Areas, manually add an "Unassigned" row and use these formulas:
> - **B (Teams):** `=COUNTBLANK(FILTER(TeamWorkload!B2:B, TeamWorkload!A2:A<>""))`
> - **C-F (Sums):** `=SUMPRODUCT((TeamWorkload!$B$2:$B="")*TeamWorkload!C$2:C)` (adjust column letter)
>
> **Prerequisite:** TeamWorkload sheet must be populated first (formula-based from TicketData and Teams).

#### Sheet: `AtRiskResponse`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| **Ticket Number** | **Subject** | **Team** | **SLA Type** | **Threshold (hrs)** | **Elapsed (hrs)** | **% of SLA** | **Time Remaining** |
| (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) |

> **Note:** This sheet uses formulas reading from consolidated TicketData (SLA metrics in columns AD-AJ) - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually

2. **Cell A2:** Response SLA at-risk tickets (single formula populates A-H)
   ```
   =LET(riskPct, IFERROR(VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100, 0.75),
   data, FILTER({TicketData!B:B, TicketData!C:C, TicketData!L:L, TicketData!AD:AD, TicketData!AE:AE},
   (TicketData!I:I="No")*(TicketData!AJ:AJ=TRUE)*(TicketData!AD:AD>0)*(TicketData!AE:AE>0)*
   (TicketData!AE:AE/TicketData!AD:AD>=riskPct)*(TicketData!AE:AE/TicketData!AD:AD<1)*(TicketData!AF:AF<>TRUE)),
   IFERROR(SORT({INDEX(data,,1), LEFT(INDEX(data,,2),60), INDEX(data,,3),
   IF(ROWS(data)>0,"Response",""), INDEX(data,,4)/60, INDEX(data,,5)/60,
   INDEX(data,,5)/INDEX(data,,4), (INDEX(data,,4)-INDEX(data,,5))/60}, 7, FALSE),
   "No at-risk Response tickets"))
   ```

> **How it works:**
> - Filters TicketData where IsClosed="No" (col I), IsRunning=TRUE (col AJ), threshold > 0, and % used is between SLA_RISK_PERCENT and 100%
> - SLA columns: AD=ResponseThreshold, AE=ResponseActual, AF=ResponseBreach
> - Threshold/Elapsed/Remaining are converted from minutes to hours
> - Sorted by % of SLA descending (most urgent first)

#### Sheet: `AtRiskResolution`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| **Ticket Number** | **Subject** | **Team** | **SLA Type** | **Threshold (hrs)** | **Elapsed (hrs)** | **% of SLA** | **Time Remaining** |
| (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) |

> **Note:** This sheet uses formulas reading from consolidated TicketData (SLA metrics in columns AD-AJ) - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually

2. **Cell A2:** Resolution SLA at-risk tickets (single formula populates A-H)
   ```
   =LET(riskPct, IFERROR(VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100, 0.75),
   data, FILTER({TicketData!B:B, TicketData!C:C, TicketData!L:L, TicketData!AG:AG, TicketData!AH:AH},
   (TicketData!I:I="No")*(TicketData!AJ:AJ=TRUE)*(TicketData!AG:AG>0)*(TicketData!AH:AH>0)*
   (TicketData!AH:AH/TicketData!AG:AG>=riskPct)*(TicketData!AH:AH/TicketData!AG:AG<1)*(TicketData!AI:AI<>TRUE)),
   IFERROR(SORT({INDEX(data,,1), LEFT(INDEX(data,,2),60), INDEX(data,,3),
   IF(ROWS(data)>0,"Resolution",""), INDEX(data,,4)/60, INDEX(data,,5)/60,
   INDEX(data,,5)/INDEX(data,,4), (INDEX(data,,4)-INDEX(data,,5))/60}, 7, FALSE),
   "No at-risk Resolution tickets"))
   ```

> **How it works:**
> - Filters TicketData where IsClosed="No" (col I), IsRunning=TRUE (col AJ), threshold > 0, and % used is between SLA_RISK_PERCENT and 100%
> - SLA columns: AG=ResolutionThreshold, AH=ResolutionActual, AI=ResolutionBreach, AJ=IsRunning
> - Threshold/Elapsed/Remaining are converted from minutes to hours
> - Sorted by % of SLA descending (most urgent first)
>
> **Formatting Tips:**
> - Format column G as percentage
> - Format column H as number with 1 decimal place (shows hours remaining)
>
> **Prerequisites:** TicketData must be loaded with SLA metrics. SLA_RISK_PERCENT must be set in Config (default 75).

#### Sheet: `StaleTickets`

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| **Ticket Number** | **Subject** | **Team** | **Days Since Update** | **Last Update** | **Created Date** | **Status** |
| (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) |

> **Note:** This sheet uses a single formula to filter and sort from TicketData - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Cell A2:** Single formula that returns all stale tickets, sorted by days since update
   ```
   =SORT(FILTER({TicketData!B2:B, LEFT(TicketData!C2:C,80), TicketData!K2:K, INT(TODAY()-DATEVALUE(LEFT(TicketData!F2:F,10))), LEFT(TicketData!F2:F,10), LEFT(TicketData!E2:E,10), TicketData!I2:I}, (TicketData!H2:H="No")*(INT(TODAY()-DATEVALUE(LEFT(TicketData!F2:F,10)))>=VLOOKUP("STALE_DAYS",Config!A:B,2,FALSE))), 4, FALSE)
   ```

> **How it works:**
> - Filters TicketData where IsClosed="No" AND days since ModifiedDate >= STALE_DAYS from Config
> - Returns 7 columns: TicketNumber, Subject (truncated), TeamName, DaysSinceUpdate, LastUpdate, CreatedDate, Status
> - Sorts by DaysSinceUpdate (column 4) in descending order (oldest first)
> - Uses `LEFT(...,10)` to extract date portion from ISO timestamps
>
> **Tip:** Adjust STALE_DAYS in the Config sheet to change the threshold (default 7 days).

---

### Optional Analytics Sheets

Additional analytics sheets can be added via **iiQ Data > Add Analytics Sheet** menu. Each district can customize which metrics they track. All sheets can be deleted and recreated as needed.

**Menu Structure:**

```
iiQ Data > Add Analytics Sheet >
├── Volume & Trends
│   ├── Monthly Volume ★
│   ├── Performance Trends ★
│   ├── Seasonal Comparison (YoY)
│   └── Temporal Patterns
├── Backlog & Quality
│   ├── Backlog Aging ★
│   ├── Stale Tickets
│   └── Reopen Rate
├── SLA & Response
│   ├── SLA Compliance ★
│   ├── At-Risk Queue ★
│   ├── First Contact Resolution
│   └── Response Distribution
├── Team & Staff
│   ├── Team Workload ★
│   ├── Technician Performance
│   └── Functional Area Summary
├── Location
│   ├── Location Breakdown
│   └── Location Type Comparison
└── Issue & Requester
    ├── Issue Category Volume
    ├── Priority Analysis
    └── Frequent Requesters
```

★ = Default sheet (created by Setup Spreadsheet, can be recreated if deleted)

**Available Optional Sheets:**

| Sheet | Question Answered | Key Metrics |
|-------|-------------------|-------------|
| **Seasonal Comparison** | "How does this year compare to last year?" | YoY change by month for volume planning |
| **Temporal Patterns** | "When do tickets come in?" | Day-of-week and hour-of-day distributions |
| **Stale Tickets** | "Which tickets have no recent activity?" | Open tickets with no update in X days |
| **Reopen Rate** | "Are we truly resolving issues?" | Reopened ticket detection and quality metrics |
| **First Contact Resolution** | "How many tickets resolved same-day?" | Same-day %, 4-hour resolution %, monthly trend |
| **Response Distribution** | "How consistent is our response time?" | Percentiles, distribution buckets, consistency |
| **Technician Performance** | "How is workload distributed among staff?" | Per-technician: open, closed, aged, breach rate |
| **Functional Area Summary** | "How are functional areas performing?" | Aggregated metrics by functional area |
| **Location Breakdown** | "Which locations generate the most tickets?" | Per-location: open, created, closed |
| **Location Type Comparison** | "Which school types generate most tickets?" | Metrics by Elementary/Middle/High/Admin |
| **Issue Category Volume** | "What types of problems are we handling?" | Open/Closed by category, breach rate per category |
| **Priority Analysis** | "Are high-priority tickets handled faster?" | Metrics by priority level, response times |
| **Frequent Requesters** | "Who generates the most tickets?" | Top 50 requesters with category data |

> **How to Add a Sheet:**
> 1. Go to **iiQ Data > Add Analytics Sheet**
> 2. Select the category (e.g., "Volume & Trends")
> 3. Click the sheet you want to add
> 4. The sheet will be created with all formulas pre-configured
>
> **How to Remove a Sheet:**
> - Right-click the sheet tab and select "Delete"
> - You can recreate it anytime via the menu
>
> **Customization:**
> - All formulas are visible and editable
> - Modify age buckets, thresholds, or calculations as needed
> - Create your own analytics sheets using TicketData as the source

---

#### Sheet: `TicketData`

| Col | Header | Description |
|-----|--------|-------------|
| A | **TicketId** | UUID primary key |
| B | **TicketNumber** | Human-readable ticket number |
| C | **Subject** | Ticket subject line (truncated to 200 chars) |
| D | **Year** | Extracted from CreatedDate (2024, 2025, 2026) |
| E | **CreatedDate** | ISO timestamp when ticket was created |
| F | **StartedDate** | ISO timestamp when work began (for queue time) |
| G | **ModifiedDate** | ISO timestamp of last modification |
| H | **ClosedDate** | ISO timestamp when closed (blank if open) |
| I | **IsClosed** | "Yes" or "No" |
| J | **Status** | Current workflow step name |
| K | **TeamId** | Assigned team UUID |
| L | **TeamName** | Assigned team name (for FA mapping) |
| M | **LocationId** | Location UUID |
| N | **LocationName** | Location name |
| O | **LocationType** | Location type (School, Department, etc.) |
| P | **OwnerId** | Ticket owner UUID |
| Q | **OwnerName** | Ticket owner name |
| R | **AgeDays** | Days open (or days between created and closed) |
| S | **Priority** | Priority weight (integer) |
| T | **IsPastDue** | "Yes" or "No" |
| U | **DueDate** | ISO timestamp of due date (if set) |
| V | **SlaId** | Applied SLA UUID (if any) |
| W | **SlaName** | Applied SLA name |
| X | **IssueCategoryId** | Issue category UUID |
| Y | **IssueCategoryName** | Issue category name (e.g., "Hardware", "Software") |
| Z | **IssueTypeId** | Issue type UUID |
| AA | **IssueTypeName** | Issue type name (e.g., "Display", "Battery") |
| AB | **ForId** | Requester user UUID |
| AC | **ForName** | Requester user name |
| AD | **ResponseThreshold** | Required first response time in minutes (from SLA policy) |
| AE | **ResponseActual** | Actual first response time in minutes |
| AF | **ResponseBreach** | TRUE if actual > threshold, FALSE otherwise |
| AG | **ResolutionThreshold** | Required resolution time in minutes (from SLA policy) |
| AH | **ResolutionActual** | Actual resolution time in minutes |
| AI | **ResolutionBreach** | TRUE if actual > threshold, FALSE otherwise |
| AJ | **IsRunning** | TRUE if SLA timer is still active (ticket not yet resolved) |

> **Note:** Raw ticket data dump with consolidated SLA metrics for Power BI analysis. Data is loaded by year with automatic resume capability. 36 columns total.
>
> **Loading Strategy:**
> - **Historical years (2024, 2025)**: Standard pagination with page tracking. Once complete, these don't change.
> - **Current year (2026)**: Date windowing for incremental updates. Use "Open Ticket Refresh" every 2 hours for open ticket SLA updates.
>
> **SLA Columns Explained (AD-AJ):**
> - **ResponseThreshold/ResolutionThreshold**: SLA policy limits in minutes
> - **ResponseActual/ResolutionActual**: Actual elapsed time in minutes
> - **ResponseBreach/ResolutionBreach**: TRUE if actual exceeded threshold
> - **IsRunning**: TRUE if SLA timer is still active (ticket not resolved)
> - Tickets without SLA policies will have blank values in these columns
>
> **Formula-Based Analytics:** With these columns, you can build all common IT metrics using sheet formulas:
> - **Volume/Throughput**: COUNTIFS on CreatedDate, ClosedDate by month/year
> - **Backlog Aging**: COUNTIFS on AgeDays ranges where IsClosed="No"
> - **Priority Distribution**: COUNTIFS on Priority or pivot table
> - **Category/IssueType Breakdown**: COUNTIFS or pivot on IssueCategoryName, IssueTypeName
> - **Team Workload**: COUNTIFS on TeamName where IsClosed="No"
> - **SLA Compliance**: Use columns AD-AJ for response/resolution times and breach status
> - **At-Risk Queue**: Filter on IsRunning=TRUE and % of SLA threshold
> - **Location Analysis**: COUNTIFS on LocationName, LocationType

---

## Part 2: Google Apps Script Code

### Step 1: Open Script Editor

1. In your Google Sheet, go to **Extensions > Apps Script**
2. Delete any default code
3. Create the files listed below (copy from the source files)

### Source Code Files

The Apps Script source code is in the `scripts/` folder of this repository.

| File | Description |
|------|-------------|
| [`Config.gs`](scripts/Config.gs) | Configuration reading from Config sheet, logging utilities |
| [`ApiClient.gs`](scripts/ApiClient.gs) | HTTP requests with retry/exponential backoff for rate limiting |
| [`Teams.gs`](scripts/Teams.gs) | Team data loading from API |
| [`TicketData.gs`](scripts/TicketData.gs) | Bulk ticket data loader with consolidated SLA (36 columns, year-based pagination) |
| [`DailySnapshot.gs`](scripts/DailySnapshot.gs) | Daily backlog metrics capture for trending |
| [`Menu.gs`](scripts/Menu.gs) | iiQ Data menu for data loader and analytics functions |
| [`Triggers.gs`](scripts/Triggers.gs) | Time-driven trigger functions for automated updates |
| [`Setup.gs`](scripts/Setup.gs) | Spreadsheet setup and default sheet creation |
| [`OptionalMetrics.gs`](scripts/OptionalMetrics.gs) | Optional analytics sheets (24 total, added via menu) |

> **Note:** Analytics sheets use formulas (see Part 1). Scripts handle data loading (TicketData with SLA, Teams), daily snapshots (DailySnapshot), and orchestration (Menu, Triggers). OptionalMetrics.gs provides menu functions to add/recreate any analytics sheet.

### File Dependencies

```
Script-based data loaders:
    ├── Config.gs (Configuration)
    ├── ApiClient.gs (HTTP Client)
    │       └── Config.gs
    ├── Teams.gs
    │       └── ApiClient.gs
    ├── TicketData.gs (includes consolidated SLA metrics)
    │       └── ApiClient.gs
    └── DailySnapshot.gs (reads from TicketData sheet)
            └── Config.gs

Menu.gs (iiQ Data Menu)
    └── References data loader and OptionalMetrics functions

Triggers.gs (Automated Updates)
    └── Data loader and analytics scripts above

Setup.gs (Spreadsheet Setup)
    └── Creates data sheets and 7 default analytics sheets

OptionalMetrics.gs (Optional Analytics)
    └── Creates any of 24 analytics sheets via menu
    └── Uses setup functions from Setup.gs for default sheets

Formula-based analytics sheets (24 total, no scripts needed):
    DEFAULT (created by Setup):
    ├── MonthlyVolume           → reads from TicketData
    ├── BacklogAging            → reads from TicketData and Config
    ├── TeamWorkload            → reads from TicketData and Teams
    ├── SLACompliance           → reads from TicketData (SLA columns AD-AJ)
    ├── PerformanceTrends       → reads from TicketData, SLACompliance, DailySnapshot
    ├── AtRiskResponse          → reads from TicketData (SLA columns AD-AJ) and Config
    └── AtRiskResolution        → reads from TicketData (SLA columns AD-AJ) and Config

    OPTIONAL (add via menu):
    ├── StaleTickets            → reads from TicketData and Config
    ├── LocationBreakdown       → reads from TicketData and Config
    ├── FunctionalAreaSummary   → reads from TeamWorkload and Config
    ├── SeasonalComparison      → reads from TicketData (YoY comparison)
    ├── TemporalPatterns        → reads from TicketData (day/hour analysis)
    ├── ReopenRate              → reads from TicketData (quality metric)
    ├── FirstContactResolution  → reads from TicketData (FCR metric)
    ├── ResponseDistribution    → reads from TicketData (percentiles)
    ├── ResponseTrends          → reads from TicketData (monthly response trends)
    ├── TechnicianPerformance   → reads from TicketData (per-owner)
    ├── LocationTypeComparison  → reads from TicketData (by school type)
    ├── IssueCategoryVolume     → reads from TicketData (by category)
    ├── PriorityAnalysis        → reads from TicketData (by priority)
    ├── FrequentRequesters      → reads from TicketData (top requesters)
    ├── QueueTimeAnalysis       → reads from TicketData (queue time stats)
    ├── QueueTimeByTeam         → reads from TicketData (queue time per team)
    └── QueueTimeTrend          → reads from TicketData (monthly queue time trend)
```

---

## Part 3: Setup Instructions

### Step 1: Create and Configure the Spreadsheet

1. Create a new Google Spreadsheet
2. **Recommended:** Run **iiQ Data > Setup > Setup Spreadsheet** to auto-create all required sheets
   - Creates data sheets: Config, TicketData, Teams, DailySnapshot, Logs
   - Creates 7 default analytics sheets: MonthlyVolume, BacklogAging, TeamWorkload, SLACompliance, PerformanceTrends, AtRiskResponse, AtRiskResolution
3. **Or create manually:** Set up each sheet with headers as shown in Part 1
4. **Add more analytics later:** Use **iiQ Data > Add Analytics Sheet** to add any of the 17 optional analytics sheets

### Step 2: Add the Apps Script Code

1. Go to **Extensions > Apps Script**
2. Create a new `.gs` file for each script in the `scripts/` folder (10 files total)
3. Copy the code from each file and save the project

### Step 3: Configure API Access

1. In the `Config` sheet, enter:
   - `API_BASE_URL`: Your iiQ instance URL (e.g., `https://yourdistrict.incidentiq.com`)
   - `BEARER_TOKEN`: Your API authentication token
   - `SITE_ID`: Your site UUID (if required)
   - `PAGE_SIZE`: `100` (recommended)

### Step 4: Authorize and Run

1. Reload the spreadsheet
2. You should see a new menu: **iiQ Data**
3. Click **iiQ Data > Refresh Teams** first
4. Authorize the script when prompted
5. Once teams load, manually fill in the **Functional Area** column in the Teams sheet

### Step 5: Initial Data Load

Before setting up automated triggers, complete the initial bulk load:

1. Click **iiQ Data > Ticket Data > Continue Loading (Initial)**
2. Wait for it to finish (~5 minutes), then check **iiQ Data > Ticket Data > Show Status**
3. If any year shows "Page X of Y" instead of "Complete", run Continue Loading again
4. Repeat until all years show "Complete"

> **Why multiple runs?** Google Apps Script has a 6-minute timeout. Large districts may have 50,000+ tickets spanning multiple years. The script saves progress after each batch, so you just keep running it until done. SLA data is included automatically.

**Alternative: Let it run automatically**

If you don't want to babysit the initial load:

1. Click **iiQ Data > Setup > Setup Automated Triggers** to create all triggers automatically
2. Check the Logs sheet periodically — look for "COMPLETE" entries
3. **You can leave these triggers enabled** — they also handle ongoing refreshes for large districts

### Step 6: Set Up Scheduled Refresh

After initial data load is complete, set up triggers to keep data current:

**Easy Setup (Recommended):**
Click **iiQ Data > Setup > Setup Automated Triggers** — this creates all recommended triggers automatically.

**Manual Setup:**
In Apps Script, go to **Triggers** (clock icon) and add these triggers:

| Function | Event Source | Type | Time | Purpose |
|----------|--------------|------|------|---------|
| `triggerDataContinue` | Time-driven | Minutes timer | Every 10 minutes | Continue any in-progress loading |
| `triggerOpenTicketRefresh` | Time-driven | Hours timer | Every 2 hours | Start open ticket + SLA refresh |
| `triggerNewTickets` | Time-driven | Minutes timer | Every 30 minutes | Fetch newly created tickets |
| `triggerDailySnapshot` | Time-driven | Day timer | 7:00 PM - 8:00 PM | Capture backlog metrics |
| `triggerWeeklyFullRefresh` | Time-driven | Week timer | Sunday 2:00 AM | Full reload (catch deletions) |

> **Data Freshness with this schedule:**
> - **Open ticket SLA data**: max 2 hours stale
> - **New tickets**: appear within 30 minutes
> - **Status changes**: captured within 2 hours
> - **Deletions/corrections**: captured weekly
>
> **Trigger Details:**
> - `triggerDataContinue`: Your "keep things moving" trigger — continues initial load if not done, OR continues open refresh if in progress. Does nothing when both are complete. Safe to leave enabled permanently.
> - `triggerOpenTicketRefresh`: Starts a fresh open ticket refresh — fetches all open tickets and recently closed (last 7 days), updates rows in-place
> - `triggerNewTickets`: Fetches tickets created since last fetch timestamp, appends new rows
> - `triggerDailySnapshot`: Captures current backlog metrics for PerformanceTrends (cannot be calculated retroactively)
> - `triggerWeeklyFullRefresh`: Clears and reloads all ticket data to catch deletions and corrections
>
> **Note:** Analytics sheets are formula-based and auto-update when data changes. All triggers log activity to the Logs sheet.

---

## Part 4: Power BI Connection

### Option A: Direct Google Sheets Connector

1. In Power BI Desktop, go to **Get Data > Online Services > Google Sheets**
2. Authenticate with your Google account
3. Select your spreadsheet and import the sheets you need
4. Set up scheduled refresh in Power BI Service

### Option B: Export to CSV/Excel

1. In Google Sheets, go to **File > Download > Microsoft Excel (.xlsx)**
2. Import the Excel file into Power BI
3. Repeat periodically or automate with Power Automate

### Option C: Google Sheets API (Advanced)

1. Use Power BI's Web connector with Google Sheets API
2. Requires API key configuration
3. More control but more complex setup

### Recommended Power BI Data Model

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│ MonthlyVolume   │     │ BacklogAging    │     │ TeamWorkload    │
├─────────────────┤     ├─────────────────┤     ├─────────────────┤
│ Month           │     │ AgeBucket       │     │ TeamName        │
│ Year            │     │ Count           │     │ FunctionalArea  │◀── FA Rollup
│ Created         │     │ Percentage      │     │ Open            │
│ Closed          │     │ LastRefreshed   │     │ CreatedMTD      │
│ NetChange       │     └─────────────────┘     │ ClosedMTD       │
│ ClosureRate     │                             │ Aged30Plus      │
└─────────────────┘     ┌─────────────────┐     │ LastRefreshed   │
                        │ Teams           │     └─────────────────┘
┌─────────────────┐     ├─────────────────┤
│ SLACompliance   │     │ TeamId          │
├─────────────────┤     │ TeamName        │
│ Month           │     │ FunctionalArea  │◀── Manual mapping
│ Year            │     │ IsActive        │
│ Closed          │     └─────────────────┘
│ Breaches        │
│ BreachRate      │
│ AvgResponseHrs  │
│ AvgResolutionHrs│
└─────────────────┘
```

---

## Part 5: Sample Outputs

### MonthlyVolume Sheet (Formula-Calculated)

| Month | Year | Created | Closed | Net Change | Closure Rate |
|-------|------|---------|--------|------------|--------------|
| January | 2026 | 487 | 452 | 35 | 92.8% |
| February | 2026 | 423 | 445 | -22 | 105.2% |
| March | 2026 | 512 | 498 | 14 | 97.3% |

> **Formula-Based:** This sheet calculates automatically from TicketData - no refresh needed. Add new rows with Month/Year values and drag formulas down to extend.

### BacklogAging Sheet (Formula-Calculated)

| Age Bucket | Count | % of Total | Sample Ticket | Last Refreshed |
|------------|-------|------------|---------------|----------------|
| 0-15 days | 145 | 46.5% | | 2026-01-21T14:30:00Z |
| 16-30 days | 78 | 25.0% | TKT-45892 | 2026-01-21T14:30:00Z |
| 31-60 days | 52 | 16.7% | TKT-44521 | 2026-01-21T14:30:00Z |
| 61-90 days | 24 | 7.7% | TKT-43102 | 2026-01-21T14:30:00Z |
| 90+ days | 13 | 4.2% | TKT-41856 | 2026-01-21T14:30:00Z |
| **TOTAL** | **312** | **100%** | | |

> **Formula-Based:** This sheet calculates automatically from TicketData using COUNTIFS on the AgeDays column. Updates when TicketData is refreshed.

### TeamWorkload Sheet (Formula-Calculated, Sortable)

| Team Name | Functional Area | Open | Created (MTD) | Closed (MTD) | Aged 30+ | | Sort Col# | Desc? |
|-----------|-----------------|------|---------------|--------------|----------|---|-----------|-------|
| Help Desk Tier 1 | IT Support Services | 89 | 156 | 142 | 12 | | 3 | FALSE |
| Help Desk Tier 2 | IT Support Services | 45 | 67 | 58 | 8 | | | |
| Network Operations | Network | 34 | 28 | 31 | 5 | | | |
| App Development | Applications and Data | 28 | 15 | 12 | 6 | | | |
| Field Services | IT Support Services | 23 | 45 | 48 | 3 | | | |
| Database Admin | Applications and Data | 18 | 8 | 10 | 4 | | | |

> **Formula-Based & Sortable:** This sheet uses a single array formula (LET + BYROW + HSTACK + SORT) that outputs all data at once. Set H2 to column number (1-6) and I2 to FALSE (descending) or TRUE (ascending) to re-sort.
>
> **Power BI Tip:** Use this sheet to create a pivot table or chart grouped by Functional Area to see aggregate metrics across all teams in each area.

### SLACompliance Sheet (Formula-Calculated)

| Month | Year | Closed | Breaches | Breach Rate | Avg Response (hrs) | Avg Resolution (hrs) |
|-------|------|--------|----------|-------------|-------------------|---------------------|
| January | 2026 | 452 | 34 | 7.5% | 2.4 | 18.6 |
| February | 2026 | 445 | 28 | 6.3% | 2.1 | 16.2 |
| March | 2026 | 498 | 41 | 8.2% | 2.8 | 22.1 |

> **Formula-Based:** This sheet calculates automatically from TicketData - no refresh needed. Data updates when TicketData is refreshed.
>
> **Metrics Explained:**
> - **Breaches**: Count of tickets where response OR resolution time exceeded SLA threshold
> - **Breach Rate**: Breaches ÷ Closed tickets
> - **Avg Response**: Average time to first response (in hours)
> - **Avg Resolution**: Average time to close ticket (in hours)

### PerformanceTrends Sheet (Formula-Calculated)

| Month | Year | Closed | Avg Resolution (days) | Closure Rate | Breach Rate | Backlog EOM | % Aged 30+ EOM |
|-------|------|--------|----------------------|--------------|-------------|-------------|----------------|
| October | 2025 | 445 | 4.2 | 98% | 6.8% | 298 | 12.4% |
| November | 2025 | 467 | 3.8 | 102% | 5.9% | 285 | 11.2% |
| December | 2025 | 412 | 4.1 | 95% | 7.2% | 301 | 13.1% |
| January | 2026 | 452 | 3.5 | 104% | 5.5% | 278 | 10.8% |

> **Formula-Based:** Combines data from TicketData, SLACompliance, and DailySnapshot to show performance trends over time.
>
> **"Are We Getting Better?" Interpretation:**
> - **Avg Resolution ↓** = Faster ticket resolution
> - **Closure Rate ↑** (≥100%) = Keeping up with or reducing backlog
> - **Breach Rate ↓** = Better SLA compliance
> - **Backlog EOM ↓** = Shrinking ticket debt
> - **% Aged 30+ ↓** = Addressing older tickets
>
> **Note:** Backlog EOM and % Aged 30+ EOM require DailySnapshot data. Start the daily trigger to begin capturing this data.

### DailySnapshot Sheet (Script-Populated)

| Date | Open Tickets | Aged 30+ Count | % Aged 30+ |
|------|--------------|----------------|------------|
| 2026-01-20 | 285 | 32 | 11.2% |
| 2026-01-21 | 282 | 31 | 11.0% |
| 2026-01-22 | 278 | 30 | 10.8% |

> **Script-Based:** The `triggerDailySnapshot` function captures these metrics daily. Cannot be calculated retroactively.
>
> **Getting Started:** Run `populateHistoricalSnapshots()` once to create estimated historical data based on current ticket dates (estimates only).

### LocationBreakdown Sheet (Formula-Calculated)

| Location Name | Location Type | Open | Created (MTD) | Closed (MTD) | Last Refreshed |
|---------------|---------------|------|---------------|--------------|----------------|
| Central High School | School | 45 | 67 | 58 | 2026-01-21T14:30:00Z |
| Technology Center | District Office | 38 | 42 | 45 | 2026-01-21T14:30:00Z |
| East Elementary | School | 32 | 51 | 48 | 2026-01-21T14:30:00Z |
| West Middle School | School | 28 | 34 | 31 | 2026-01-21T14:30:00Z |
| Admin Building | District Office | 24 | 28 | 30 | 2026-01-21T14:30:00Z |

> **Formula-Based:** This sheet calculates automatically from TicketData - no refresh needed. Data updates when TicketData is refreshed.
>
> **Power BI Tip:** Use this sheet to identify schools or departments with high ticket volume or growing backlogs.

### FunctionalAreaSummary Sheet (Formula-Calculated)

| Functional Area | Teams | Open | Created (MTD) | Closed (MTD) | Aged 30+ | % Aged 30+ | Last Refreshed |
|-----------------|-------|------|---------------|--------------|----------|------------|----------------|
| IT Support Services | 5 | 157 | 268 | 248 | 23 | 14.6% | 2026-01-21T14:30:00Z |
| Applications and Data | 3 | 46 | 23 | 22 | 10 | 21.7% | 2026-01-21T14:30:00Z |
| Network | 2 | 41 | 35 | 38 | 7 | 17.1% | 2026-01-21T14:30:00Z |
| IT Security | 2 | 18 | 12 | 14 | 3 | 16.7% | 2026-01-21T14:30:00Z |
| Unassigned | 4 | 12 | 8 | 6 | 5 | 41.7% | 2026-01-21T14:30:00Z |

> **Formula-Based:** This sheet aggregates from TeamWorkload using UNIQUE and SUMIF formulas. Updates automatically when TeamWorkload data changes.
>
> **Executive Dashboard:** This sheet provides a high-level Functional Area view for leadership reporting. Monitor "% Aged 30+" to identify areas with growing backlog risk.

### AtRiskResponse / AtRiskResolution Sheets (Formula-Calculated)

These two sheets filter tickets approaching their respective SLA thresholds. Same layout, different SLA metrics.

| Ticket Number | Subject | Team | SLA Type | Threshold (hrs) | Elapsed (hrs) | % of SLA | Time Remaining |
|---------------|---------|------|----------|-----------------|---------------|----------|----------------|
| TKT-47123 | Network printer offline | Help Desk Tier 2 | Response | 4.0 | 3.8 | 95% | 0.2 |
| TKT-47045 | Software installation | Help Desk Tier 1 | Response | 2.0 | 1.6 | 80% | 0.4 |
| TKT-46982 | Account lockout | Help Desk Tier 1 | Response | 1.0 | 0.8 | 80% | 0.2 |
| TKT-47098 | Laptop replacement | Field Services | Resolution | 48.0 | 44.2 | 92% | 3.8 |

> **Formula-Based:** This sheet filters TicketData (SLA columns AD-AJ) to find tickets where SLA % is between the risk threshold and 100%. Time Remaining shows hours.
>
> **Operational Alert:** Tickets approaching SLA threshold (default 75%+) sorted by urgency. Use this queue for daily triage to prevent breaches.

### StaleTickets Sheet (Formula-Calculated)

| Ticket Number | Subject | Team | Days Since Update | Last Update | Created Date | Status |
|---------------|---------|------|-------------------|-------------|--------------|--------|
| TKT-42156 | Server migration planning | App Development | 45 | 2025-12-07 | 2025-11-15 | In Progress |
| TKT-43201 | Network cabling request | Network Operations | 32 | 2025-12-20 | 2025-12-01 | Pending |
| TKT-44567 | Equipment disposal | Field Services | 21 | 2026-01-01 | 2025-12-15 | On Hold |
| TKT-45012 | Training request | Help Desk Tier 1 | 14 | 2026-01-07 | 2026-01-02 | New |

> **Formula-Based:** This sheet calculates automatically from TicketData using a single FILTER/SORT formula. Updates when TicketData is refreshed.
>
> **Risk Indicator:** Tickets with no activity in the configured STALE_DAYS threshold (default 7 days). Review weekly to prevent tickets from falling through the cracks.

### TicketData Sheet (After Refresh)

The TicketData sheet includes 36 columns (A-AJ) with consolidated SLA metrics. Sample rows:

| Column | Row 1 | Row 2 | Row 3 |
|--------|-------|-------|-------|
| A: TicketId | abc-123... | mno-345... | yza-567... |
| B: TicketNumber | TKT-47123 | TKT-47098 | TKT-46982 |
| C: Subject | Network printer offline | Laptop replacement | Password reset |
| D: Year | 2026 | 2026 | 2026 |
| E: CreatedDate | 2026-01-15T09:30:00 | 2026-01-10T11:15:00 | 2026-01-18T08:00:00 |
| F: StartedDate | 2026-01-15T09:35:00 | 2026-01-10T11:30:00 | 2026-01-18T08:02:00 |
| G: ModifiedDate | 2026-01-21T14:20:00 | 2026-01-20T16:45:00 | 2026-01-18T08:15:00 |
| H: ClosedDate | | | 2026-01-18T08:15:00 |
| I: IsClosed | No | No | Yes |
| J: Status | In Progress | Pending Parts | Resolved |
| K: TeamId | def-456... | pqr-678... | bcd-890... |
| L: TeamName | Help Desk Tier 2 | Field Services | Help Desk Tier 1 |
| M: LocationId | ghi-789... | stu-901... | efg-123... |
| N: LocationName | Central High School | East Elementary | Admin Building |
| O: LocationType | High School | Elementary | Administrative |
| P: OwnerId | jkl-012... | vwx-234... | hij-456... |
| Q: OwnerName | John Smith | Jane Doe | Bob Jones |
| R: AgeDays | 6 | 11 | 0 |
| S: Priority | 50 | 75 | 25 |
| T: IsPastDue | No | Yes | No |
| U: DueDate | 2026-01-22T17:00:00 | 2026-01-18T17:00:00 | |
| V: SlaId | sla-111... | sla-222... | sla-333... |
| W: SlaName | Standard Response | Device Repair | Quick Fix |
| X: IssueCategoryId | cat-aaa... | cat-bbb... | cat-ccc... |
| Y: IssueCategoryName | Hardware | Hardware | Software |
| Z: IssueTypeId | iss-xxx... | iss-yyy... | iss-zzz... |
| AA: IssueTypeName | Printer | Laptop | Password |
| AB: RequesterId | usr-111... | usr-222... | usr-333... |
| AC: RequesterName | Sarah Wilson | Mike Brown | Emily Davis |
| AD: ResponseThreshold | 240 | 480 | 60 |
| AE: ResponseActual | 45 | 120 | 12 |
| AF: ResponseBreach | FALSE | FALSE | FALSE |
| AG: ResolutionThreshold | 2880 | 4320 | 480 |
| AH: ResolutionActual | | | 15 |
| AI: ResolutionBreach | | | FALSE |
| AJ: IsRunning | TRUE | TRUE | FALSE |

> **Raw Data Export:** This sheet contains all 36 columns including consolidated SLA metrics for custom analysis, pivot tables, or Power BI integration.
>
> **SLA Columns (AD-AJ):**
> - Threshold values are in minutes (240 = 4 hours)
> - Actual values are in minutes (blank if not yet measured)
> - Breach columns are TRUE/FALSE (blank if not applicable)
> - IsRunning is TRUE while SLA timer is active
>
> **Formula-Based Analytics:** With this data, build metrics using COUNTIFS, SUMIFS, and pivot tables without additional API calls:
> - Volume: `=COUNTIFS(E:E, ">=2026-01-01", E:E, "<2026-02-01")` for monthly created
> - Aging: `=COUNTIFS(I:I, "No", R:R, ">=31", R:R, "<=60")` for 31-60 day bucket
> - SLA Breach Rate: `=COUNTIFS(AF:AF, TRUE)/COUNTIFS(AF:AF, "<>"&"")` for response breach rate
> - At-Risk: Filter on AJ:AJ=TRUE and AE:AE/AD:AD >= 0.75 for tickets approaching response SLA
> - By Category: `=COUNTIFS(Y:Y, "Hardware", I:I, "No")` for open hardware tickets

---

## Troubleshooting

### Common Problems

| What You See | What's Wrong | How to Fix It |
|--------------|--------------|---------------|
| "API configuration missing" | Config sheet isn't set up | Fill in `API_BASE_URL` and `BEARER_TOKEN` in the Config sheet |
| "HTTP 401" error | Your API token expired | Get a new Bearer token from iiQ Admin > Integrations > API |
| "HTTP 403" error | API user doesn't have permission | Check that your API user has read access to tickets |
| "HTTP 429" or "RATE_LIMITED" in Logs | iiQ is throttling your requests | Increase `THROTTLE_MS` in Config to 2000 or 3000 |
| Script timeout after ~6 minutes | Normal — Google's limit | Just run "Continue Loading" again — progress is saved |
| Some tickets have blank SLA columns | Those tickets don't have SLA policies | Expected behavior — not all tickets have SLA assigned in iiQ |
| Formulas show #REF! errors | TicketData sheet is empty | Run the initial data load first |
| TeamWorkload shows blank Functional Areas | Teams sheet column C is empty | Fill in Functional Area labels in the Teams sheet |

### Rate Limiting (HTTP 429)

If you're getting rate limited by Incident IQ:

1. **Increase the delay between API calls**: Set `THROTTLE_MS` to `2000` (2 seconds) or `3000` (3 seconds) in the Config sheet
2. **Check your Logs sheet**: Look for "RATE_LIMITED" or "RETRY" entries to see how often it's happening
3. **Be patient during initial load**: Large districts may take several hours to fully load — that's normal
4. **Reduce batch size if needed**: Lower `TICKET_BATCH_SIZE` from 2000 to 1000 if problems persist

---

## Next Steps

Once your data is flowing, here are some ideas for getting more value:

1. **Connect Power BI** — Create visual dashboards for your leadership team. The Google Sheets connector works well, or export to Excel periodically.

2. **Set up email alerts** — Add Apps Script triggers to email you when:
   - AtRiskResponse or AtRiskResolution has more than 10 tickets (SLA breaches coming)
   - StaleTickets has more than 20 tickets (things are falling through cracks)
   - A specific team's backlog exceeds a threshold

3. **Share with your team** — Give your team leads read access to see their TeamWorkload metrics. Consider creating filtered views by team.

4. **Track month-over-month** — The DailySnapshot trigger captures backlog metrics daily. After a few months, you'll have great data for showing improvement trends to administration.

5. **Customize for your district** — The formulas are all visible and editable. Add columns, change age buckets, or create new analytics sheets for your specific needs.

---

## Features & Capabilities

### Default Analytics (created by Setup)

| Feature | Sheet(s) | Status |
|---------|----------|--------|
| Tickets Created/Closed by month | MonthlyVolume | Default |
| Net Backlog Change | MonthlyVolume | Default |
| Closure Rate | MonthlyVolume, PerformanceTrends | Default |
| Backlog Aging (0-15, 16-30, 31-60, 61-90, 90+) | BacklogAging | Default |
| SLA Breaches and Breach Rate | SLACompliance, PerformanceTrends | Default |
| Avg Response/Resolution Time | SLACompliance, PerformanceTrends | Default |
| Team-level Metrics | TeamWorkload | Default |
| At-Risk Response (approaching Response SLA) | AtRiskResponse | Default |
| At-Risk Resolution (approaching Resolution SLA) | AtRiskResolution | Default |
| **"Are we getting better?" Trending** | PerformanceTrends, DailySnapshot | **Default** |
| Historical Backlog Tracking | DailySnapshot | Default |

### Optional Analytics (add via menu)

| Feature | Sheet(s) | Category |
|---------|----------|----------|
| Stale Tickets (no update X days) | StaleTickets | Backlog & Quality |
| Reopen Rate (quality metric) | ReopenRate | Backlog & Quality |
| Location Breakdown | LocationBreakdown | Location |
| Location Type Comparison (Elem/Middle/High) | LocationTypeComparison | Location |
| Functional Area Reporting | FunctionalAreaSummary | Team & Staff |
| Technician Performance | TechnicianPerformance | Team & Staff |
| Seasonal/Year-over-Year Comparison | SeasonalComparison | Volume & Trends |
| Temporal Patterns (day/hour) | TemporalPatterns | Volume & Trends |
| First Contact Resolution Rate | FirstContactResolution | SLA & Response |
| Response Time Distribution | ResponseDistribution | SLA & Response |
| Response Time Trends | ResponseTrends | SLA & Response |
| Queue Time Analysis | QueueTimeAnalysis | SLA & Response |
| Queue Time by Team | QueueTimeByTeam | SLA & Response |
| Queue Time Trend | QueueTimeTrend | SLA & Response |
| Issue Category Volume | IssueCategoryVolume | Issue & Requester |
| Priority Analysis | PriorityAnalysis | Issue & Requester |
| Frequent Requesters | FrequentRequesters | Issue & Requester |

> **24 Total Analytics Sheets:** 7 default + 17 optional. All can be deleted and recreated via **iiQ Data > Add Analytics Sheet** menu.

---

*Questions? Open an issue on this repository, or check the Logs sheet for error details.*
