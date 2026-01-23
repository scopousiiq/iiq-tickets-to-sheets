# IncidentIQ Data Extraction for Google Sheets

**Purpose:** Extract IIQ ticket and SLA data into Google Sheets for Power BI consumption and analytics
**Use Cases:** Monthly Volume Tracking, Backlog Aging Analysis, SLA Compliance, Team Workload Metrics

---

## Overview

This guide provides ready-to-use Google Sheets templates and Apps Script code to extract ticket data from IncidentIQ's API. The extracted data can be connected to Power BI for visualization.

### Architecture

```
┌─────────────┐     ┌──────────────────┐     ┌─────────────┐
│  IIQ API    │────▶│  Google Sheets   │────▶│  Power BI   │
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

Create a new Google Sheet with the following tabs (sheets):

| Sheet Name | Purpose |
|------------|---------|
| `Config` | API credentials and settings |
| `MonthlyVolume` | Created/Closed/Open by month |
| `BacklogAging` | Current open tickets by age bucket |
| `Teams` | Team directory with FA mapping |
| `TeamWorkload` | Per-team metrics for FA rollup |
| `SLACompliance` | SLA breaches, response & resolution times |
| `PerformanceTrends` | "Are we getting better?" trending metrics |
| `LocationBreakdown` | Ticket distribution by location |
| `FunctionalAreaSummary` | Aggregated metrics by Functional Area |
| `AtRiskQueue` | Tickets approaching SLA thresholds |
| `StaleTickets` | Tickets with no update in X days |
| `DailySnapshot` | Daily backlog metrics for trending |
| `TicketData` | Raw ticket dump for analysis |
| `TicketSlaData` | SLA timing data for compliance analysis |
| `Logs` | API call logs and errors |

### Step 2: Configure the Sheets

#### Sheet: `Config`

| Row | A (Setting) | B (Value) |
|-----|-------------|-----------|
| 1 | **Setting** | **Value** |
| 2 | API_BASE_URL | https://your-district.incidentiq.com/api |
| 3 | BEARER_TOKEN | (your JWT token) |
| 4 | SITE_ID | (your site UUID) |
| 5 | PAGE_SIZE | 100 |
| 6 | THROTTLE_MS | 1000 |
| 7 | STALE_DAYS | 7 |
| 8 | SLA_RISK_PERCENT | 75 |
| 9 | TICKET_BATCH_SIZE | 2000 |
| 10 | TICKET_2024_TOTAL_PAGES | (auto-populated) |
| 11 | TICKET_2024_LAST_PAGE | -1 |
| 12 | TICKET_2024_COMPLETE | FALSE |
| 13 | TICKET_2025_TOTAL_PAGES | (auto-populated) |
| 14 | TICKET_2025_LAST_PAGE | -1 |
| 15 | TICKET_2025_COMPLETE | FALSE |
| 16 | TICKET_2026_LAST_FETCH | (auto-populated ISO timestamp) |
| 17 | LAST_REFRESH | (auto-populated) |
| 18 | SLA_BATCH_SIZE | 100 |
| 19 | SLA_2024_LAST_PAGE | -1 |
| 20 | SLA_2024_COMPLETE | FALSE |
| 21 | SLA_2025_LAST_PAGE | -1 |
| 22 | SLA_2025_COMPLETE | FALSE |
| 23 | SLA_2026_LAST_FETCH | (auto-populated ISO timestamp) |

> **Year Configuration:** Years are auto-discovered from Config rows - no code changes needed.
> - **Historical years** (pagination-based): Detected from `TICKET_{YEAR}_LAST_PAGE` rows
> - **Current year** (date windowing): Detected from `TICKET_{YEAR}_LAST_FETCH` row
>
> **To add a year (e.g., 2023):** Add these rows to Config:
> ```
> TICKET_2023_TOTAL_PAGES | (auto-populated)
> TICKET_2023_LAST_PAGE   | -1
> TICKET_2023_COMPLETE    | FALSE
> SLA_2023_LAST_PAGE      | -1
> SLA_2023_COMPLETE       | FALSE
> ```
>
> **To remove a year (e.g., 2024):** Delete the `TICKET_2024_*` and `SLA_2024_*` rows from Config, then optionally delete data rows from TicketData/TicketSlaData sheets.
>
> **To change current year:** Replace `TICKET_2026_LAST_FETCH` with `TICKET_2027_LAST_FETCH` (and same for SLA).
>
> **SLA Data Settings:** The `/tickets/slas` endpoint is capped at 100 records per page and does NOT return paging metadata. The loader iterates until fewer than 100 records are returned (indicating the last page).
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
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!G:G, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!G:G, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")))
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
   - **B2 (0-15 days):** `=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=0", TicketData!Q:Q, "<=15")`
   - **B3 (16-30 days):** `=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=16", TicketData!Q:Q, "<=30")`
   - **B4 (31-60 days):** `=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=31", TicketData!Q:Q, "<=60")`
   - **B5 (61-90 days):** `=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=61", TicketData!Q:Q, "<=90")`
   - **B6 (90+ days):** `=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">90")`
   - **B7 (TOTAL):** `=SUM(B2:B6)`

4. **Column C (% of Total):** Enter for each row (format as percentage):
   - **C2-C6:** `=IF($B$7>0, B2/$B$7, 0)` (drag down for each bucket row)
   - **C7:** Enter `100%` or `=1`

5. **Column D (Sample Ticket):** Oldest ticket in each bucket:
   - **D2:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=0)*(TicketData!Q2:Q<=15)),2,FALSE),1,1),"")`
   - **D3:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=16)*(TicketData!Q2:Q<=30)),2,FALSE),1,1),"")`
   - **D4:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=31)*(TicketData!Q2:Q<=60)),2,FALSE),1,1),"")`
   - **D5:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=61)*(TicketData!Q2:Q<=90)),2,FALSE),1,1),"")`
   - **D6:** `=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>90)),2,FALSE),1,1),"")`

6. **Column E (Last Refreshed):**
   - **E2:** `=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")` (drag down)

> **How it works:** Uses TicketData column H (IsClosed) and Q (AgeDays) to count open tickets by age bucket. Sample Ticket formula filters, sorts by age descending, and returns the oldest ticket number in each bucket.

#### Sheet: `Teams`

| A | B | C | D |
|---|---|---|---|
| **Team ID** | **Team Name** | **Functional Area** | **Is Active** |
| (auto-populated) | (auto-populated) | (manual entry) | (auto-populated) |

#### Sheet: `TeamWorkload`

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| **Team Name** | **Functional Area** | **Open** | **Created (MTD)** | **Closed (MTD)** | **Aged 30+** | **Last Refreshed** | **Sort Col#** | **Desc?** |
| (formula - spills down) | | | | | | (formula) | 3 | FALSE |

> **Note:** This sheet uses a single array formula to calculate all data from TicketData. The entire table is sortable via H2 (column number 1-6) and I2 (FALSE=descending, TRUE=ascending).

**Setup Instructions:**

1. **Row 1 (Headers):** Enter column headers A1:F1 manually, plus "Sort By" in H1 and "Order" in I1

2. **Cell H1:** Label "Sort Col#", **Cell I1:** Label "Desc?"

4. **Cell H2:** Enter the sort column number (use Data Validation dropdown)
   - Values: `3` (Open - default), `1` (Team), `2` (FA), `4` (Created), `5` (Closed), `6` (Aged)

5. **Cell I2:** Enter sort order (use Data Validation dropdown)
   - Values: `FALSE` (Descending - default), `TRUE` (Ascending)

6. **Cell A2:** Single formula that outputs the entire sortable table
   ```
   =LET(
     teams, UNIQUE(FILTER(TicketData!K2:K, TicketData!K2:K<>"", TicketData!K2:K<>"TeamName")),
     mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),
     mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),
     col_a, teams,
     col_b, BYROW(teams, LAMBDA(t, IFERROR(VLOOKUP(t, Teams!B:C, 2, FALSE), ""))),
     col_c, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!H:H, "No"))),
     col_d, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),
     col_e, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!G:G, ">="&mtdStart, TicketData!G:G, "<"&mtdEnd))),
     col_f, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!H:H, "No", TicketData!Q:Q, ">=30"))),
     data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),
     SORT(data, $H$2, $I$2)
   )
   ```

5. **Cell G1:** Add "Last Refreshed" header, **Cell G2:** Show last refresh timestamp
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

> **Note:** This sheet uses formulas joining TicketData and TicketSlaData - no script required.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Column A & B:** Manually enter the months and years you want to track

3. **Cell C2:** Count tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!G:G, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!G:G, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")))
   ```

4. **Cell D2:** Count SLA breaches for tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), startD, TEXT(DATE(B2,m,1), "YYYY-MM-DD"), endD, TEXT(DATE(B2,m+1,1), "YYYY-MM-DD"), closedIds, FILTER(TicketData!A:A, (TicketData!G:G>=startD)*(TicketData!G:G<endD)), SUMPRODUCT((COUNTIF(closedIds, TicketSlaData!A2:A)>0)*((TicketSlaData!G2:G="TRUE")+(TicketSlaData!J2:J="TRUE")>0)))
   ```

5. **Cell E2:** Calculate breach rate (drag down, format as percentage)
   ```
   =IF(C2>0, D2/C2, 0)
   ```

6. **Cell F2:** Average response time in hours for tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), startD, TEXT(DATE(B2,m,1), "YYYY-MM-DD"), endD, TEXT(DATE(B2,m+1,1), "YYYY-MM-DD"), closedIds, FILTER(TicketData!A:A, (TicketData!G:G>=startD)*(TicketData!G:G<endD)), times, FILTER(TicketSlaData!F:F, (COUNTIF(closedIds, TicketSlaData!A:A)>0)*(TicketSlaData!F:F>0)), IFERROR(AVERAGE(times)/60, "N/A"))
   ```

7. **Cell G2:** Average resolution time in hours for tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), startD, TEXT(DATE(B2,m,1), "YYYY-MM-DD"), endD, TEXT(DATE(B2,m+1,1), "YYYY-MM-DD"), closedIds, FILTER(TicketData!A:A, (TicketData!G:G>=startD)*(TicketData!G:G<endD)), times, FILTER(TicketSlaData!I:I, (COUNTIF(closedIds, TicketSlaData!A:A)>0)*(TicketSlaData!I:I>0)), IFERROR(AVERAGE(times)/60, "N/A"))
   ```

> **How it works:** These formulas join TicketSlaData to TicketData via TicketId (column A in both sheets) to filter SLA metrics by ticket ClosedDate.
>
> **Performance Note:** The join formulas (D2, F2, G2) may be slow with large datasets (50k+ rows). If performance is an issue, consider using pivot tables or keeping this sheet script-based.
>
> **Prerequisites:** Both TicketData and TicketSlaData must be loaded first.

#### Sheet: `PerformanceTrends`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| **Month** | **Year** | **Closed** | **Avg Resolution (days)** | **Closure Rate** | **Breach Rate** | **Backlog EOM** | **% Aged 30+ EOM** |
| January | 2026 | (formula) | (formula) | (formula) | (formula) | (from snapshot) | (from snapshot) |
| February | 2026 | (formula) | (formula) | (formula) | (formula) | (from snapshot) | (from snapshot) |

> **Purpose:** Answer "Are we getting better?" by tracking key performance metrics over time.

**Setup Instructions:**

1. **Row 1 (Headers):** Enter the column headers manually
2. **Column A & B:** Manually enter the months and years you want to track

3. **Cell C2:** Count tickets closed in that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!G:G, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!G:G, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")))
   ```

4. **Cell D2:** Average resolution time in DAYS for tickets closed that month (drag down)
   ```
   =LET(m, MATCH(A2, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), IFERROR(AVERAGEIFS(TicketData!Q:Q, TicketData!H:H, "Yes", TicketData!G:G, ">="&TEXT(DATE(B2,m,1), "YYYY-MM-DD"), TicketData!G:G, "<"&TEXT(DATE(B2,m+1,1), "YYYY-MM-DD")), "N/A"))
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
   =COUNTIFS(TicketData!M:M, A2, TicketData!H:H, "No")
   ```
5. **Cell D2:** Count tickets created this month (drag down)
   ```
   =COUNTIFS(TicketData!M:M, A2, TicketData!E:E, ">="&EOMONTH(TODAY(),-1)+1, TicketData!E:E, "<="&TODAY())
   ```
6. **Cell E2:** Count tickets closed this month (drag down)
   ```
   =COUNTIFS(TicketData!M:M, A2, TicketData!G:G, ">="&EOMONTH(TODAY(),-1)+1, TicketData!G:G, "<="&TODAY())
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

#### Sheet: `AtRiskQueue`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| **Ticket Number** | **Subject** | **Team** | **SLA Type** | **Threshold (hrs)** | **Elapsed (hrs)** | **% of SLA** | **Time Remaining** |
| (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) | (formula) |

> **Note:** This sheet uses formulas joining TicketSlaData and TicketData - no script required.

**Setup Instructions:**

This sheet requires two formula blocks - one for Response SLA at-risk tickets, one for Resolution. Each ticket can appear in both if both metrics are at risk.

1. **Row 1 (Headers):** Enter the column headers manually

2. **Cell A2:** Response SLA at-risk tickets (single formula populates A-H)
   ```
   =LET(riskPct, VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100, atRisk, FILTER({TicketSlaData!A2:A, TicketSlaData!E2:E, TicketSlaData!F2:F}, (TicketSlaData!K2:K="TRUE")*(TicketSlaData!E2:E>0)*(TicketSlaData!F2:F/TicketSlaData!E2:E>=riskPct)*(TicketSlaData!F2:F/TicketSlaData!E2:E<1)*(TicketSlaData!G2:G<>"TRUE")), IFERROR(SORT({VLOOKUP(INDEX(atRisk,,1),TicketData!A:B,2,FALSE), VLOOKUP(INDEX(atRisk,,1),TicketData!A:C,3,FALSE), VLOOKUP(INDEX(atRisk,,1),TicketData!A:K,11,FALSE), IF(ROW(atRisk),"Response"), INDEX(atRisk,,2)/60, INDEX(atRisk,,3)/60, INDEX(atRisk,,3)/INDEX(atRisk,,2), (INDEX(atRisk,,2)-INDEX(atRisk,,3))/60}, 7, FALSE), "No at-risk Response tickets"))
   ```

3. **Resolution at-risk:** Add below the Response results (find next empty row after Response data)
   ```
   =LET(riskPct, VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100, atRisk, FILTER({TicketSlaData!A2:A, TicketSlaData!H2:H, TicketSlaData!I2:I}, (TicketSlaData!K2:K="TRUE")*(TicketSlaData!H2:H>0)*(TicketSlaData!I2:I/TicketSlaData!H2:H>=riskPct)*(TicketSlaData!I2:I/TicketSlaData!H2:H<1)*(TicketSlaData!J2:J<>"TRUE")), IFERROR(SORT({VLOOKUP(INDEX(atRisk,,1),TicketData!A:B,2,FALSE), VLOOKUP(INDEX(atRisk,,1),TicketData!A:C,3,FALSE), VLOOKUP(INDEX(atRisk,,1),TicketData!A:K,11,FALSE), IF(ROW(atRisk),"Resolution"), INDEX(atRisk,,2)/60, INDEX(atRisk,,3)/60, INDEX(atRisk,,3)/INDEX(atRisk,,2), (INDEX(atRisk,,2)-INDEX(atRisk,,3))/60}, 7, FALSE), "No at-risk Resolution tickets"))
   ```

> **How it works:**
> - Filters TicketSlaData where IsRunning=TRUE, threshold > 0, and % used is between SLA_RISK_PERCENT and 100%
> - Joins to TicketData via TicketId to get TicketNumber, Subject, TeamName
> - Threshold/Elapsed/Remaining are converted from minutes to hours
> - Sorted by % of SLA descending (most urgent first)
>
> **Formatting Tips:**
> - Format column G as percentage
> - Format column H as number with 1 decimal place (shows hours remaining)
>
> **Performance Note:** These join formulas may be slow with large datasets. If performance is an issue, consider filtering TicketSlaData first or using pivot tables.
>
> **Prerequisites:** Both TicketData and TicketSlaData must be loaded. SLA_RISK_PERCENT must be set in Config (default 75).

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

#### Sheet: `TicketData`

| Col | Header | Description |
|-----|--------|-------------|
| A | **TicketId** | UUID primary key |
| B | **TicketNumber** | Human-readable ticket number |
| C | **Subject** | Ticket subject line (truncated to 200 chars) |
| D | **Year** | Extracted from CreatedDate (2024, 2025, 2026) |
| E | **CreatedDate** | ISO timestamp when ticket was created |
| F | **ModifiedDate** | ISO timestamp of last modification |
| G | **ClosedDate** | ISO timestamp when closed (blank if open) |
| H | **IsClosed** | "Yes" or "No" |
| I | **Status** | Current workflow step name |
| J | **TeamId** | Assigned team UUID |
| K | **TeamName** | Assigned team name (for FA mapping) |
| L | **LocationId** | Location UUID |
| M | **LocationName** | Location name |
| N | **LocationType** | Location type (School, Department, etc.) |
| O | **OwnerId** | Ticket owner UUID |
| P | **OwnerName** | Ticket owner name |
| Q | **AgeDays** | Days open (or days between created and closed) |
| R | **Priority** | Priority weight (integer) |
| S | **IsPastDue** | "Yes" or "No" |
| T | **DueDate** | ISO timestamp of due date (if set) |
| U | **SlaId** | Applied SLA UUID (if any) |
| V | **SlaName** | Applied SLA name |
| W | **IssueCategoryId** | Issue category UUID |
| X | **IssueCategoryName** | Issue category name (e.g., "Hardware", "Software") |
| Y | **IssueTypeId** | Issue type UUID |
| Z | **IssueTypeName** | Issue type name (e.g., "Display", "Battery") |
| AA | **ForId** | Requester user UUID |
| AB | **ForName** | Requester user name |

> **Note:** Raw ticket data dump for Power BI analysis. Data is loaded by year with automatic resume capability. 28 columns total.
>
> **Loading Strategy:**
> - **Historical years (2024, 2025)**: Standard pagination with page tracking. Once complete, these don't change.
> - **Current year (2026)**: Date windowing for incremental updates. Use "Refresh Current Year" periodically to catch deletions.
>
> **Formula-Based Analytics:** With these columns, you can build all common IT metrics using sheet formulas:
> - **Volume/Throughput**: COUNTIFS on CreatedDate, ClosedDate by month/year
> - **Backlog Aging**: COUNTIFS on AgeDays ranges where IsClosed="No"
> - **Priority Distribution**: COUNTIFS on Priority or pivot table
> - **Category/IssueType Breakdown**: COUNTIFS or pivot on IssueCategoryName, IssueTypeName
> - **Team Workload**: COUNTIFS on TeamName where IsClosed="No"
> - **SLA Assignment**: COUNTIFS on SlaName, IsPastDue
> - **Location Analysis**: COUNTIFS on LocationName, LocationType
>
> **Power BI Joins:** Join to TicketSlaData on TicketId for detailed SLA timing (response/resolution times, breach status).

#### Sheet: `TicketSlaData`

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| **TicketId** | **Year** | **SlaId** | **SlaName** | **ResponseThreshold** | **ResponseActual** | **ResponseBreached** | **ResolutionThreshold** | **ResolutionActual** | **ResolutionBreached** | **IsRunning** |

> **Note:** SLA timing data for joining with TicketData via TicketId. Uses the `/tickets/slas` endpoint which returns detailed SLA metrics.
>
> **Loading Strategy:**
> - Same year-based approach as TicketData (historical pagination + current year date windowing)
> - API is capped at 100 records per page (slower than TicketData)
> - Estimated ~1,700 API calls for full historical load (~170k tickets)
>
> **Columns Explained:**
> - **TicketId**: Foreign key to join with TicketData
> - **Year**: Extracted from ticket creation for filtering (2024, 2025, 2026)
> - **SlaId/SlaName**: The SLA policy applied to this ticket
> - **ResponseThreshold**: Required first response time in minutes (from Sla.Metrics)
> - **ResponseActual**: Actual first response time in minutes (from SlaTimes, MetricType=1)
> - **ResponseBreached**: TRUE if actual > threshold, FALSE otherwise
> - **ResolutionThreshold**: Required resolution time in minutes (from Sla.Metrics)
> - **ResolutionActual**: Actual resolution time in minutes (from SlaTimes, MetricType=2)
> - **ResolutionBreached**: TRUE if actual > threshold, FALSE otherwise
> - **IsRunning**: TRUE if SLA is still active (ticket not yet resolved)
>
> **Power BI Usage:** Join TicketSlaData to TicketData on TicketId for SLA compliance analysis by team, location, or time period.

---

## Part 2: Google Apps Script Code

### Step 1: Open Script Editor

1. In your Google Sheet, go to **Extensions > Apps Script**
2. Delete any default code
3. Create the files listed below (copy from the source files)

### Source Code Files

The Apps Script source code is available in the [`gAppScript/gSheetDashboardSample/`](../gAppScript/gSheetDashboardSample/) directory.

| File | Description |
|------|-------------|
| [`Config.gs`](../gAppScript/gSheetDashboardSample/Config.gs) | Configuration reading from Config sheet, logging utilities |
| [`ApiClient.gs`](../gAppScript/gSheetDashboardSample/ApiClient.gs) | HTTP requests with retry/exponential backoff for rate limiting |
| [`Teams.gs`](../gAppScript/gSheetDashboardSample/Teams.gs) | Team data loading from API |
| [`TicketData.gs`](../gAppScript/gSheetDashboardSample/TicketData.gs) | Bulk ticket data loader (28 columns, year-based pagination) |
| [`TicketSlaData.gs`](../gAppScript/gSheetDashboardSample/TicketSlaData.gs) | SLA timing data loader (response/resolution metrics) |
| [`DailySnapshot.gs`](../gAppScript/gSheetDashboardSample/DailySnapshot.gs) | Daily backlog metrics capture for trending |
| [`Menu.gs`](../gAppScript/gSheetDashboardSample/Menu.gs) | IIQ Data menu for data loader functions |
| [`Triggers.gs`](../gAppScript/gSheetDashboardSample/Triggers.gs) | Time-driven trigger functions for automated updates |

> **Note:** Analytics sheets use formulas (see Part 1). Scripts handle data loading (TicketData, TicketSlaData, Teams), daily snapshots (DailySnapshot), and orchestration (Menu, Triggers).

### File Dependencies

```
Script-based data loaders:
    ├── Config.gs (Configuration)
    ├── ApiClient.gs (HTTP Client)
    │       └── Config.gs
    ├── Teams.gs
    │       └── ApiClient.gs
    ├── TicketData.gs
    │       └── ApiClient.gs
    ├── TicketSlaData.gs
    │       └── ApiClient.gs
    └── DailySnapshot.gs (reads from TicketData sheet)
            └── Config.gs

Menu.gs (IIQ Data Menu)
    └── References data loader functions

Triggers.gs (Automated Updates)
    └── Data loader scripts above

Formula-based analytics sheets (no scripts needed):
    ├── MonthlyVolume           → reads from TicketData
    ├── BacklogAging            → reads from TicketData and Config
    ├── TeamWorkload            → reads from TicketData and Teams
    ├── FunctionalAreaSummary   → reads from TeamWorkload and Config
    ├── SLACompliance           → reads from TicketData and TicketSlaData
    ├── PerformanceTrends       → reads from TicketData, SLACompliance, DailySnapshot
    ├── AtRiskQueue             → reads from TicketSlaData, TicketData, and Config
    ├── LocationBreakdown       → reads from TicketData and Config
    └── StaleTickets            → reads from TicketData and Config
```

---

## Part 3: Setup Instructions

### Step 1: Create and Configure the Spreadsheet

1. Create a new Google Spreadsheet
2. Create all sheets: `Config`, `MonthlyVolume`, `BacklogAging`, `Teams`, `TeamWorkload`, `SLACompliance`, `LocationBreakdown`, `FunctionalAreaSummary`, `AtRiskQueue`, `StaleTickets`, `TicketData`, `TicketSlaData`, `Logs`
3. Set up headers as shown in Part 1

### Step 2: Add the Apps Script Code

1. Go to **Extensions > Apps Script**
2. Create the eight `.gs` files with the code from Part 2
3. Save the project

### Step 3: Configure API Access

1. In the `Config` sheet, enter:
   - `API_BASE_URL`: Your IIQ instance URL (e.g., `https://yourdistrict.incidentiq.com/api`)
   - `BEARER_TOKEN`: Your API authentication token
   - `SITE_ID`: Your site UUID (if required)
   - `PAGE_SIZE`: `100` (recommended)

### Step 4: Authorize and Run

1. Reload the spreadsheet
2. You should see a new menu: **IIQ Data**
3. Click **IIQ Data > Refresh Teams** first
4. Authorize the script when prompted
5. Once teams load, manually fill in the **Functional Area** column in the Teams sheet

### Step 5: Initial Data Load

Before setting up automated triggers, complete the initial bulk load:

1. Run **IIQ Data > Ticket Data > Continue Loading** repeatedly until all historical years show "Complete"
2. Run **IIQ Data > SLA Data > Continue Loading** repeatedly until all historical years show "Complete"
3. Analytics sheets will auto-populate via formulas once TicketData and TicketSlaData have data

**Alternative: Automated Bulk Load**
1. In Apps Script, go to **Triggers** (clock icon)
2. Add a trigger for `triggerBulkLoadContinue` every 10 minutes
3. Monitor the Logs sheet until historical data is complete
4. Delete the bulk load trigger once complete

### Step 6: Set Up Scheduled Refresh

After initial data load is complete, set up daily triggers to keep data current:

1. In Apps Script, go to **Triggers** (clock icon)
2. Add these triggers:

| Function | Event Source | Type | Time |
|----------|--------------|------|------|
| `triggerTicketDataUpdate` | Time-driven | Day timer | 5:00 AM - 6:00 AM |
| `triggerSlaDataUpdate` | Time-driven | Day timer | 5:30 AM - 6:30 AM |
| `triggerDailySnapshot` | Time-driven | Day timer | 7:00 PM - 8:00 PM |

> **Note:** Analytics sheets are formula-based and auto-update when data changes. The daily snapshot captures backlog metrics for PerformanceTrends that cannot be calculated retroactively. All triggers log activity to the Logs sheet.

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

> **Formula-Based:** This sheet joins TicketData and TicketSlaData via TicketId to calculate SLA metrics for tickets closed in each month.
>
> **Metrics Explained:**
> - **Breaches**: Count of tickets where response or resolution time exceeded SLA threshold
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

### AtRiskQueue Sheet (Formula-Calculated)

| Ticket Number | Subject | Team | SLA Type | Threshold (hrs) | Elapsed (hrs) | % of SLA | Time Remaining |
|---------------|---------|------|----------|-----------------|---------------|----------|----------------|
| TKT-47123 | Network printer offline | Help Desk Tier 2 | Response | 4.0 | 3.8 | 95% | 0.2 |
| TKT-47098 | Laptop replacement | Field Services | Resolution | 48.0 | 44.2 | 92% | 3.8 |
| TKT-47045 | Software installation | Help Desk Tier 1 | Response | 2.0 | 1.6 | 80% | 0.4 |
| TKT-46982 | Account lockout | Help Desk Tier 1 | Response | 1.0 | 0.8 | 80% | 0.2 |

> **Formula-Based:** This sheet joins TicketSlaData and TicketData to find tickets where SLA % is between the risk threshold and 100%. Time Remaining shows hours.
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

The expanded TicketData sheet now includes 28 columns (A-AB) for comprehensive analytics. Sample rows:

| Column | Row 1 | Row 2 | Row 3 |
|--------|-------|-------|-------|
| A: TicketId | abc-123... | mno-345... | yza-567... |
| B: TicketNumber | TKT-47123 | TKT-47098 | TKT-46982 |
| C: Subject | Network printer offline | Laptop replacement | Password reset |
| D: Year | 2026 | 2026 | 2026 |
| E: CreatedDate | 2026-01-15T09:30:00 | 2026-01-10T11:15:00 | 2026-01-18T08:00:00 |
| F: ModifiedDate | 2026-01-21T14:20:00 | 2026-01-20T16:45:00 | 2026-01-18T08:15:00 |
| G: ClosedDate | | | 2026-01-18T08:15:00 |
| H: IsClosed | No | No | Yes |
| I: Status | In Progress | Pending Parts | Resolved |
| J: TeamId | def-456... | pqr-678... | bcd-890... |
| K: TeamName | Help Desk Tier 2 | Field Services | Help Desk Tier 1 |
| L: LocationId | ghi-789... | stu-901... | efg-123... |
| M: LocationName | Central High School | East Elementary | Admin Building |
| N: LocationType | High School | Elementary | Administrative |
| O: OwnerId | jkl-012... | vwx-234... | hij-456... |
| P: OwnerName | John Smith | Jane Doe | Bob Jones |
| Q: AgeDays | 6 | 11 | 0 |
| R: Priority | 50 | 75 | 25 |
| S: IsPastDue | No | Yes | No |
| T: DueDate | 2026-01-22T17:00:00 | 2026-01-18T17:00:00 | |
| U: SlaId | sla-111... | sla-222... | sla-333... |
| V: SlaName | Standard Response | Device Repair | Quick Fix |
| W: IssueCategoryId | cat-aaa... | cat-bbb... | cat-ccc... |
| X: IssueCategoryName | Hardware | Hardware | Software |
| Y: IssueTypeId | iss-xxx... | iss-yyy... | iss-zzz... |
| Z: IssueTypeName | Printer | Laptop | Password |
| AA: ForId | usr-111... | usr-222... | usr-333... |
| AB: ForName | Sarah Wilson | Mike Brown | Emily Davis |

> **Raw Data Export:** This sheet contains all 28 ticket fields needed for custom analysis, pivot tables, or Power BI integration.
>
> **Formula-Based Analytics:** With this data, build metrics using COUNTIFS, SUMIFS, and pivot tables without additional API calls:
> - Volume: `=COUNTIFS(E:E, ">=2026-01-01", E:E, "<2026-02-01")` for monthly created
> - Aging: `=COUNTIFS(H:H, "No", Q:Q, ">=31", Q:Q, "<=60")` for 31-60 day bucket
> - SLA At-Risk: `=COUNTIFS(H:H, "No", S:S, "Yes")` for past-due open tickets
> - By Category: `=COUNTIFS(X:X, "Hardware", H:H, "No")` for open hardware tickets
>
> **Important:** After expanding the sheet, run "Full Reload" to populate the new columns for all historical data.

---

## Troubleshooting

| Issue | Likely Cause | Solution |
|-------|--------------|----------|
| "API configuration missing" | Config sheet not set up | Fill in API_BASE_URL and BEARER_TOKEN |
| "HTTP 401" | Invalid or expired token | Refresh your Bearer token |
| "HTTP 403" | Insufficient permissions | Check API user permissions |
| "HTTP 429" or "RATE_LIMITED" in logs | Rate limited | Increase THROTTLE_MS in Config (try 2000 or higher) |
| "Bandwidth quota exceeded" | Too many requests | Increase THROTTLE_MS to 3000+, or wait and retry |
| No data returned | Wrong filters or empty result | Check filter syntax and date ranges |
| Slow performance | Too many tickets | Increase PAGE_SIZE, use date filters |
| Script timeout (6 min limit) | Too much data | Use Continue Loading - it saves progress and resumes |
| Analytics formulas slow | Large datasets | Consider using pivot tables or reduce data range |

### Rate Limiting Tips

1. **Start with higher throttle**: Set `THROTTLE_MS` to `2000` (2 seconds) initially
2. **Check the Logs sheet**: Look for RATE_LIMITED or RETRY entries to diagnose issues
3. **Reduce PAGE_SIZE**: Lower from 100 to 50 if bandwidth is constrained
4. **Use incremental loading**: The data loaders save progress and can resume after timeout

---

## Next Steps

After validating the data extraction:

1. **Build Power BI dashboards** - Connect to Google Sheets and create executive visualizations
2. **Set up automated alerts** - Use Google Apps Script triggers to send email alerts when:
   - At-Risk Queue exceeds threshold count
   - Stale Tickets exceed threshold count
   - SLA Breach Rate exceeds target percentage
3. **Add historical trending** - Store daily snapshots of FA Summary for month-over-month comparison
4. **Enhance drill-down** - Add ticket-level detail exports for each sheet

---

## Features & Capabilities

| Feature | Sheet(s) | Status |
|---------|----------|--------|
| Tickets Created/Closed by month | MonthlyVolume | Implemented |
| Net Backlog Change | MonthlyVolume | Implemented |
| Closure Rate | MonthlyVolume, PerformanceTrends | Implemented |
| Backlog Aging (0-15, 16-30, 31-60, 61-90, 90+) | BacklogAging | Implemented |
| SLA Breaches and Breach Rate | SLACompliance, PerformanceTrends | Implemented |
| Avg Response/Resolution Time | SLACompliance, PerformanceTrends | Implemented |
| Functional Area Reporting | FunctionalAreaSummary | Implemented |
| Team-level Metrics | TeamWorkload | Implemented |
| Location Breakdown | LocationBreakdown | Implemented |
| At-Risk Queue (approaching SLA) | AtRiskQueue | Implemented |
| Stale Tickets (no update X days) | StaleTickets | Implemented |
| % Tickets Aged 30+ | FunctionalAreaSummary, PerformanceTrends | Implemented |
| Drill-down to Assigned Groups | TeamWorkload | Implemented |
| **"Are we getting better?" Trending** | PerformanceTrends, DailySnapshot | **Implemented** |
| Historical Backlog Tracking | DailySnapshot | Implemented |

---

*For questions or contributions, please open an issue on this repository.*
