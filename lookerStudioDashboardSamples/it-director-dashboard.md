# IT Director Operations Dashboard — Looker Studio Build Guide

Single-page operations dashboard for a K-12 district IT Director focused on daily execution, backlog control, and measurable performance improvement.

## Management Purpose

This dashboard should answer:
- Are we keeping up with incoming demand?
- Where is backlog aging or SLA risk building?
- Which teams/technicians need rebalancing today?
- Are response and resolution times improving month-over-month?

---

## Data Source Setup

### Connect Three Sheets

1. Go to [lookerstudio.google.com](https://lookerstudio.google.com)
2. Create > Report > Google Sheets connector
3. Select your spreadsheet, choose the **TicketData** sheet, check "Use first row as headers"
4. Add a second data source: Add Data > Google Sheets > same spreadsheet > **DailySnapshot** sheet
5. Add a third data source: Add Data > Google Sheets > same spreadsheet > **MonthlyVolume** sheet

### Field Type Fixes (TicketData)

| Field | Set Type To | Notes |
|-------|-------------|-------|
| `CreatedDate` | Date & Time | ISO 8601 datetime from API |
| `StartedDate` | Date & Time | ISO 8601 datetime from API |
| `ModifiedDate` | Date & Time | ISO 8601 datetime from API |
| `ClosedDate` | Date & Time | ISO 8601 datetime from API |
| `DueDate` | Date & Time | ISO 8601 datetime from API |
| `AgeDays` | Number | |
| `ResponseThreshold` | Number | |
| `ResponseActual` | Number | |
| `ResolutionThreshold` | Number | |
| `ResolutionActual` | Number | |
| `IsClosed` | Text | |
| `IsPastDue` | Text | |
| `ResponseBreach` | Number | |
| `ResolutionBreach` | Number | |
| `IsRunning` | Number | |
| `AssetTag` | Text | |
| `ModelName` | Text | |
| `SerialNumber` | Text | |

> **Date handling:** The date fields arrive as ISO 8601 strings from the iiQ API (e.g., `2025-02-05T14:30:00Z`). When you set these to **Date & Time** in Looker Studio, the Google Sheets connector should parse them automatically. If a date field isn't recognized after changing the type, create a calculated field using:
> ```
> PARSE_DATETIME("%Y-%m-%dT%H:%M:%S", REGEXP_REPLACE(CreatedDate, "Z$", ""))
> ```
>
> **Granularity tip:** You can change a Date & Time field's granularity per-chart by clicking the dimension field in the chart's DATA panel and changing its type (e.g., to **Year Month**). However, this can fail silently — especially in blends, but also in regular charts — causing "too many rows" errors. If that happens, create a calculated field using `DATETIME_TRUNC(YourDateField, MONTH)` and use that as the dimension instead. Set the granularity on the calculated field to **Year Month** in the chart.

### Field Type Fixes (DailySnapshot)

| Field | Set Type To | Notes |
|-------|-------------|-------|
| `Date` | Date | Date-only values (YYYY-MM-DD) |
| `Open Tickets` | Number | |
| `Aged 30+ Count` | Number | |
| `% Aged 30+` | Number (Percent) | |

> **Data values:**
> - `IsClosed`: `"Closed"` / `"Open"` (text strings)
> - `IsPastDue`: `"Overdue"` / `"On Track"` (text strings)
> - `ResponseBreach`, `ResolutionBreach`, `IsRunning`: `1` / `0` (numeric — use directly in AVG/SUM)

---

## Calculated Fields (TicketData)

Create these in the **TicketData** data source.

### Is Open

```text
CASE WHEN IsClosed = "Open" THEN 1 ELSE 0 END
```

- Type: Number
- Usage: `SUM` for open counts

### Is Closed

```text
CASE WHEN IsClosed = "Closed" THEN 1 ELSE 0 END
```

- Type: Number
- Usage: `SUM` for closed counts

### SLA Breached

```text
CASE
  WHEN ResponseBreach = 1 OR ResolutionBreach = 1 THEN 1
  WHEN ResponseBreach = 0 OR ResolutionBreach = 0 THEN 0
  ELSE NULL
END
```

- Type: Number
- Usage: `AVG`, format as Percent
- **Important:** Add a "Has SLA" filter (see below) to any chart using this field. Looker Studio does not reliably exclude NULLs from AVG calculations, which will dilute the result to near 0%.

**"Has SLA" filter** — Create this as a reusable filter and apply it to any chart that uses `SLA Breached`:
- Name: `Has SLA`
- Clause 1: **Include** `ResponseBreach` **Equal to (=)** `1`
- **OR**
- Clause 2: **Include** `ResponseBreach` **Equal to (=)** `0`

Looker Studio only allows one value per filter condition, so the OR with two clauses is required.

### ClosedMonth

Used as the dimension for time-based charts grouped by month. Changing date granularity directly on `ClosedDate` in a chart's DATA panel can fail silently, causing "too many rows" errors. This calculated field avoids that.

```text
DATETIME_TRUNC(ClosedDate, MONTH)
```

- Type: Date & Time

### CreatedWeek

Used as the dimension for the Intake vs Throughput chart. Blended data sources do not support changing date granularity via the dropdown.

```text
DATETIME_TRUNC(CreatedDate, ISOWEEK)
```

- Type: Date & Time

### ClosedWeek

Used as the join key in the Intake vs Throughput blend.

```text
DATETIME_TRUNC(ClosedDate, ISOWEEK)
```

- Type: Date & Time

### Open Aged 30+

```text
CASE
  WHEN IsClosed = "Open" AND AgeDays >= 30 THEN 1
  ELSE 0
END
```

- Type: Number
- Usage: `SUM`

### Open Past Due

```text
CASE
  WHEN IsClosed = "Open" AND IsPastDue = "Overdue" THEN 1
  ELSE 0
END
```

- Type: Number
- Usage: `SUM`

### SLA At Risk (75%+)

```text
CASE
  WHEN IsClosed = "Open" AND (
    (ResponseThreshold > 0 AND ResponseActual >= ResponseThreshold * 0.75 AND ResponseBreach <> 1)
    OR
    (ResolutionThreshold > 0 AND ResolutionActual >= ResolutionThreshold * 0.75 AND ResolutionBreach <> 1)
  )
  THEN 1
  ELSE 0
END
```

- Type: Number
- Usage: `SUM`

### Response Hours

```text
ResponseActual / 60
```

- Type: Number
- Usage: `AVG`

### Resolution Days

```text
ResolutionActual / 1440
```

- Type: Number
- Usage: `AVG`

### Age Bucket

```text
CASE
  WHEN AgeDays <= 7 THEN "0-7"
  WHEN AgeDays <= 14 THEN "8-14"
  WHEN AgeDays <= 30 THEN "15-30"
  WHEN AgeDays <= 60 THEN "31-60"
  WHEN AgeDays <= 90 THEN "61-90"
  ELSE "90+"
END
```

- Type: Text
- Usage: backlog aging distribution

### Age Bucket Order

Text dimensions sort alphabetically, which breaks bucket ordering (e.g., "0-7" before "15-30" but "8-14" after "90+"). Use this for sort order.

```text
CASE
  WHEN AgeDays <= 7 THEN 1
  WHEN AgeDays <= 14 THEN 2
  WHEN AgeDays <= 30 THEN 3
  WHEN AgeDays <= 60 THEN 4
  WHEN AgeDays <= 90 THEN 5
  ELSE 6
END
```

- Type: Number
- Usage: Sort the Age Bucket chart by this field ascending. Keep `Age Bucket` as the visible dimension.

---

## Page Setup

- Page size: 1366 x 768 (landscape operations view)
- Background: White (`#FFFFFF`)
- Font: Roboto or district standard
- Chart titles: Use title case ("Backlog Trend", "Team Performance"), not ALL CAPS
- Card spacing: Keep gaps between cards to 12-16px
- Snap to grid: Enable View > Snap to grid for consistent alignment

## Brand Colors & Chart Styling

Colors sourced from the Incident IQ brand style guide.

### Color Palette

| Color | Hex | Role |
|-------|-----|------|
| Dark Blue | `#365c96` | Primary — headers, titles, default chart color |
| Teal | `#22b2a3` | Secondary — second series, success states |
| Gold | `#febb12` | Accent — third series, highlights |
| Orange | `#f1663c` | Warning — breach/alert indicators |
| Purple | `#715091` | Tertiary — fifth series if needed |
| Dark Gray | `#1F2937` | Text on dark backgrounds |
| Light Gray | `#F9FAFB` | Card/panel backgrounds |

### Styling Guidelines

- **Report-level theme:** Go to **Theme and Layout > Theme > Customize** to set the default chart color sequence to: `#365c96`, `#22b2a3`, `#febb12`, `#f1663c`, `#715091`
- **Header:** Insert a Rectangle shape across the full width, fill `#365c96` (Dark Blue). Place title text in white on top.
- **Filter bar:** Insert a Rectangle below the header, fill `#F9FAFB` (Light Gray), border 1px `#E5E7EB`. Place dropdown filters on it.
- **Scorecard banner:** Insert a Rectangle behind all scorecards, fill `#365c96`, border radius 8px. Set metric font color to white.
- **Scorecard comparison colors:** Default green/red gets drowned out by the dark blue banner. Override in STYLE tab: positive `#86EFAC` (light green), negative `#FCA5A5` (light red).
- **Card containers:** For each chart section, insert a Rectangle behind it: fill `#FFFFFF`, border 1px `#E5E7EB`, border radius 8px.
- **Chart styling (all charts):** Grid lines `#F3F4F6`, axis labels `#6B7280` at 10-11px, chart title left-aligned bold `#1F2937` at 14px.
- **Time series area charts:** Use **Time Series** chart type (not Area chart, which forces a breakdown dimension). Enable "Show area" and "Smooth line" in STYLE tab. Disable "Show points".

---

## Layout

```text
+─────────────────────────────────────────────────────────────────────────────+
│ IT Operations — Director Dashboard        [Date Range] [Team] [Priority]  │
├────────────┬────────────┬────────────┬────────────┬────────────┬──────────┤
│ Open       │ Open 30+   │ Open Past  │ SLA Breach │ Avg Resp   │ SLA At   │
│ Backlog    │ Days       │ Due        │ Rate (30d) │ Hrs (30d)  │ Risk     │
├───────────────────────────────────┬─────────────────────────────────────────┤
│ Backlog + Aging Trend             │ Intake vs Throughput                   │
│ (DailySnapshot combo chart)       │ (Created vs Closed by month)           │
├───────────────────────────────────┴─────────────────────────────────────────┤
│ Team Performance Table: Open | 30+ | Past Due | SLA Breach % | Avg Res Days │
├───────────────────────────────────┬─────────────────────────────────────────┤
│ Open Backlog by Priority          │ Open Backlog by Age Bucket             │
│ (stacked bar by team)             │ (horizontal bar)                       │
├───────────────────────────────────┴─────────────────────────────────────────┤
│ Action Queue: Top open tickets at risk (table with owner/team/age/SLA)    │
+─────────────────────────────────────────────────────────────────────────────+
```

---

## Component Build Instructions

### Header Controls

- Date range control: default to last 90 days
- Drop-down filter: `TeamName`
- Drop-down filter: `Priority`
- Drop-down filter: `LocationType`

### Row 1: KPI Scorecards

**Comparison date ranges:** Only scorecards with a chart-level date range override (like the 30-day cards) have a meaningful "previous period" for comparison. The other scorecards operate on the full dataset — there is no previous period within the same school year to compare against, so skip comparison on those.

#### Open Backlog
- Metric: `SUM(Is Open)`
- Filter: none
- Comparison: none (full dataset, no meaningful previous period)

#### Open 30+ Days
- Metric: `SUM(Open Aged 30+)`
- Filter: none

#### Open Past Due
- Metric: `SUM(Open Past Due)`
- Filter: none

#### SLA Breach Rate (30d)
- Metric: `AVG(SLA Breached)`
- Filters: `IsClosed = "Closed"`, `Has SLA` filter (see Calculated Fields section)
- Chart date range: last 30 days
- Comparison: previous period (compares to prior 30 days)

#### Avg Response Hrs (30d)
- Metric: `AVG(Response Hours)`
- Filter: `IsClosed = "Closed"`
- Chart date range: last 30 days
- Comparison: previous period (compares to prior 30 days)

#### SLA At Risk
- Metric: `SUM(SLA At Risk (75%+))`
- Filter: none

### Row 2 Left: Backlog + Aging Trend

- Data source: `DailySnapshot`
- Chart type: Combo chart
- Dimension: `Date`
- Metric 1 (bars): `Open Tickets`
- Metric 2 (line): `% Aged 30+`
- Purpose: watch both total backlog and quality of backlog age

### Row 2 Right: Intake vs Throughput

**Recommended approach — MonthlyVolume sheet (no blends):**

The Superintendent Dashboard originally tried blending two TicketData queries (Created vs Closed by date) and ran into persistent issues: date granularity dropdowns fail silently in blends, join misalignment causes missing data, and "too many rows" errors are hard to debug. The solution was to use the pre-calculated **MonthlyVolume** sheet instead, which avoids blends entirely.

- Data source: **MonthlyVolume** (add as a separate data source from the same spreadsheet)
- Chart type: Bar chart (vertical, grouped)
- Dimension: `Month` — if full month names overlap, create a calculated field `LEFT(Month, 3)` and use that
- Metrics: `Created` (SUM), `Closed` (SUM)
- Sort: `MonthOrder` ascending (calculated field — see below)
- Purpose: quickly see if closures are keeping pace with intake

**MonthlyVolume field types:**

| Field | Type |
|-------|------|
| `Month` | Text |
| `Year` | Number |
| `Created` | Number |
| `Closed` | Number |
| `Net Change` | Number |
| `Closure Rate` | Number (Percent) |

**Required calculated field** — `MonthOrder` (in the MonthlyVolume data source) to sort months in school year order:

```text
CASE
  WHEN Month = "July" THEN 1
  WHEN Month = "August" THEN 2
  WHEN Month = "September" THEN 3
  WHEN Month = "October" THEN 4
  WHEN Month = "November" THEN 5
  WHEN Month = "December" THEN 6
  WHEN Month = "January" THEN 7
  WHEN Month = "February" THEN 8
  WHEN Month = "March" THEN 9
  WHEN Month = "April" THEN 10
  WHEN Month = "May" THEN 11
  WHEN Month = "June" THEN 12
  ELSE 13
END
```

- Type: Number
- Usage: Sort the chart by this field ascending. Keep `Month` as the visible dimension.

**Alternative — Blend approach (weekly granularity):**

If you need weekly granularity instead of monthly, use blended data. Be aware that blends are fragile in Looker Studio — date granularity dropdowns don't work, and join alignment issues can silently drop data.

- Chart type: Time series or grouped bar using blended data
- **Important:** You must use the `CreatedWeek` and `ClosedWeek` calculated fields (see Calculated Fields section). Changing date granularity via the dropdown does not work in blends.
- Blend setup:
  1. Left source: `TicketData`, dimension `CreatedWeek` (calculated field), metric `Record Count`, rename to "Created"
  2. Right source: `TicketData`, dimension `ClosedWeek` (calculated field), metric `Record Count`, rename to "Closed", filter `IsClosed = "Closed"`
  3. Join type: **Full outer join**, join key: `CreatedWeek` = `ClosedWeek`
- Sort: dimension ascending

### Row 3: Team Performance Table

- Chart type: Table
- Dimension: `TeamName`
- Metrics:
  - `SUM(Is Open)` as Open
  - `SUM(Open Aged 30+)` as Aged 30+
  - `SUM(Open Past Due)` as Past Due
  - `AVG(SLA Breached)` as SLA Breach % — apply `Has SLA` filter to avoid NULL dilution
  - `AVG(Resolution Days)` as Avg Resolution Days
- Table filter: `TeamName` is not null
- Sort: Open descending

> **Note:** The `Has SLA` filter applies to the entire chart, which may reduce row counts for other metrics. If this is a problem, consider splitting SLA Breach % into a separate scorecard per team instead.

### Row 4 Left: Open Backlog by Priority and Team

- Chart type: Stacked horizontal bar
- Dimension: `TeamName`
- Breakdown dimension: `Priority`
- Metric: `SUM(Is Open)`
- Sort: Open descending
- Purpose: identify team capacity pressure and high-priority concentration

### Row 4 Right: Open Backlog by Age Bucket

- Chart type: Horizontal bar
- Dimension: `Age Bucket`
- Metric: `SUM(Is Open)`
- Filter: `IsClosed = "Open"`
- Sort: `Age Bucket Order` ascending (calculated field — text dimension sorts alphabetically, not logically)
- Style: Disable **Color by dimension** if enabled (bars should be a single solid color, not a rainbow of colors per bucket)

### Row 5: Action Queue (Table)

- Chart type: Table
- Dimensions: `TicketNumber`, `Subject`, `TeamName`, `OwnerName`, `Priority`
- Metrics: `AgeDays`, `ResponseActual`, `ResolutionActual`
- Filter: `SLA At Risk (75%+) = 1`
- Sort: `AgeDays` descending
- Rows: top 25
- Purpose: daily operational huddle list

### Optional: Device Reliability (Table or Bar Chart)

- Chart type: Table or horizontal bar
- Dimension: `ModelName`
- Metrics: Record Count, `AVG(AgeDays)`, `AVG(SLA Breached)`
- Filter: `ModelName` is not null and `ModelName` != ""
- Sort: Record Count descending
- Bars shown: Top 15
- Purpose: identify unreliable device models for replacement budget justification

---

## Why These Metrics Work for IT Directors

| Metric Group | Operational Question |
|--------------|----------------------|
| Backlog, 30+, Past Due | Where is work accumulating and becoming risky? |
| Intake vs Throughput | Are we reducing or growing demand pressure each week? |
| SLA Breach %, At Risk | Are we missing service commitments, and what will breach next? |
| Team-level table | Which teams need staffing or escalation support today? |
| Priority and age views | Are urgent tickets and old tickets getting stuck? |
| Action queue | What specific tickets need intervention now? |

---

## Director Cadence (Recommended)

- Daily standup: Row 1 KPIs + Action Queue
- Twice weekly: Team Performance Table + Priority stack
- Weekly service review: Intake vs Throughput + SLA trend
- Monthly leadership sync: summarize improvements from 30+ aging and breach rate trends

---

## Troubleshooting Notes

### Breach/Past Due fields

Since breach and running fields are now numeric (`1`/`0`), compare directly:

```text
ResponseBreach = 1
IsPastDue = "Overdue"
```

### Intake vs Throughput chart looks misaligned or shows "too many rows"

Changing date granularity via the dropdown does **not** work in blended data sources — this is a known Looker Studio bug. Use the `CreatedWeek` and `ClosedWeek` calculated fields (which use `DATETIME_TRUNC`) instead of raw date fields. Both sides of the blend must use the same truncation (ISOWEEK or MONTH) for the join to align. If blends continue to cause issues, use the MonthlyVolume sheet as a standalone data source instead.

### Too many blank teams or owners

Add chart filters:
- `TeamName` is not null and `TeamName` != ""
- `OwnerName` is not null and `OwnerName` != ""

### Date range silently filtering data

When you connect a Google Sheets data source, Looker Studio assigns a **default date dimension** (likely `CreatedDate`). Every component is automatically filtered by the report's date range — even without a date range control. If unfiltered record counts are lower than expected, set the default date range to cover your full school year in File > Report settings.

### Date fields not recognized

If Looker Studio doesn't auto-detect date columns, manually set the type to **Date & Time** in the data source config (Resource > Manage added data sources > Edit). The TicketData dates are ISO format (`2025-02-05T14:30:00Z`). If setting the type alone doesn't work, create a calculated field using:

```text
PARSE_DATETIME("%Y-%m-%dT%H:%M:%S", REGEXP_REPLACE(CreatedDate, "Z$", ""))
```
