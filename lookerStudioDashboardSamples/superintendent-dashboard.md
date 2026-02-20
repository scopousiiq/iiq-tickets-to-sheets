# Superintendent IT Dashboard — Looker Studio Build Guide

Single A4 page, at-a-glance dashboard for a school district superintendent to understand IT service desk health.

---

## Data Source Setup

Complete all data source configuration before building dashboard components. This includes connecting sheets, setting field types, creating calculated fields, and defining reusable filters.

### Connect Three Sheets

1. Go to [lookerstudio.google.com](https://lookerstudio.google.com)
2. Create > Report > Google Sheets connector
3. Select your spreadsheet, choose the **TicketData** sheet, check "Use first row as headers"
4. Add a second data source: Add Data > Google Sheets > same spreadsheet > **DailySnapshot** sheet
5. Add a third data source: Add Data > Google Sheets > same spreadsheet > **MonthlyVolume** sheet

### TicketData

#### Field Types

Apply these in the data source config (Resource > Manage added data sources > Edit):

| Field | Set Type To | Notes |
|-------|-------------|-------|
| `CreatedDate` | Date & Time | ISO 8601 datetime from API |
| `ClosedDate` | Date & Time | ISO 8601 datetime from API |
| `ModifiedDate` | Date & Time | ISO 8601 datetime from API |
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

> **Data values:**
> - `IsClosed`: `"Closed"` / `"Open"` (text strings)
> - `IsPastDue`: `"Overdue"` / `"On Track"` (text strings)
> - `ResponseBreach`, `ResolutionBreach`, `IsRunning`: `1` / `0` (numeric — use directly in AVG/SUM)

#### Calculated Fields

Create these in the TicketData data source (Resource > Manage added data sources > Edit > Add a field):

**SLA Compliant** — Used for the SLA Compliance % scorecard.

```text
CASE
  WHEN ResponseBreach = 1 OR ResolutionBreach = 1 THEN 0
  WHEN ResponseBreach = 0 OR ResolutionBreach = 0 THEN 1
  ELSE NULL
END
```

- Type: Number
- Usage: AVG aggregation, formatted as Percent
- **Important:** Apply the `Has SLA` filter (see Reusable Filters below) to any chart using this field. The `ELSE NULL` is intended to exclude tickets without SLA data, but Looker Studio does not reliably exclude NULLs from AVG calculations, which will dilute the result to near 0%.

**SLA Breached** — Used for the breach rate trend line chart.

```text
CASE
  WHEN ResponseBreach = 1 OR ResolutionBreach = 1 THEN 1
  WHEN ResponseBreach = 0 OR ResolutionBreach = 0 THEN 0
  ELSE NULL
END
```

- Type: Number
- Usage: AVG aggregation, formatted as Percent

**ClosedMonth** — Used as the dimension for time-based charts grouped by month. Changing date granularity directly on `ClosedDate` in a chart's DATA panel can fail silently, causing "too many rows" errors. This calculated field avoids that.

```text
DATETIME_TRUNC(ClosedDate, MONTH)
```

- Type: Date & Time

**Response Hours**

```text
ResponseActual / 60
```

- Type: Number

**Resolution Hours**

```text
ResolutionActual / 60
```

- Type: Number

### DailySnapshot

#### Field Types

| Field | Set Type To | Notes |
|-------|-------------|-------|
| `Date` | Date | Date-only values (YYYY-MM-DD) |
| `Open Tickets` | Number | |
| `Aged 30+ Count` | Number | |
| `% Aged 30+` | Number (Percent) | |

### MonthlyVolume

#### Field Types

| Field | Set Type To | Notes |
|-------|-------------|-------|
| `Month` | Text | |
| `Year` | Number | |
| `Created` | Number | |
| `Closed` | Number | |
| `Net Change` | Number | |
| `Closure Rate` | Number (Percent) | |

#### Calculated Fields

**MonthOrder** — Sorts months in school year order (July–June) since the `Month` field is text and would otherwise sort alphabetically.

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
- Usage: Sort charts by this field ascending. Keep `Month` as the visible dimension so axis labels show month names.

### Reusable Filters

**Has SLA** — Apply to any chart that uses `SLA Compliant` or `SLA Breached`. Excludes tickets without SLA data, which would otherwise dilute AVG calculations to near 0%.

- Name: `Has SLA`
- Clause 1: **Include** `ResponseBreach` **Equal to (=)** `1`
- **OR**
- Clause 2: **Include** `ResponseBreach` **Equal to (=)** `0`

Looker Studio only allows one value per filter condition, so the OR with two clauses is required.

---

## Page Setup

- **Page size:** 595 x 842 px (A4 portrait) — set in File > Report settings
- **Background:** White (#FFFFFF)
- **Font:** Roboto or your district standard
- **Chart titles:** Use title case ("Backlog Trend", "Monthly Volume"), not ALL CAPS. Title case looks more polished.
- **Card spacing:** Keep gaps between cards to 12-16px. Larger gaps make the page feel stretched.
- **Snap to grid:** Enable View > Snap to grid for consistent alignment

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

### Per-Component Color Assignments

| Component | Color(s) | Hex |
|-----------|----------|-----|
| Scorecard metric values | Dark Blue | `#365c96` |
| Backlog Trend line | Dark Blue | `#365c96` |
| Monthly Volume — Created bars | Dark Blue | `#365c96` |
| Monthly Volume — Closed bars | Teal | `#22b2a3` |
| SLA Breach Rate line | Orange | `#f1663c` |
| Tickets by Team bars | Dark Blue | `#365c96` |
| Top Issue Categories bars | Teal | `#22b2a3` |
| Location Type donut — Elementary | Dark Blue | `#365c96` |
| Location Type donut — Middle | Teal | `#22b2a3` |
| Location Type donut — High | Gold | `#febb12` |
| Location Type donut — Other | Orange | `#f1663c` |
| Chart titles/headers | Dark Blue | `#365c96` |
| Card backgrounds | Light Gray | `#F9FAFB` |
| Page background | White | `#FFFFFF` |

### Applying Colors in Looker Studio

1. **Report-level theme:** Go to **Theme and Layout > Theme > Customize** to set the default chart color sequence to: `#365c96`, `#22b2a3`, `#febb12`, `#f1663c`, `#715091`
2. **Per-chart overrides:** In each chart's **STYLE** tab, override individual series colors as needed per the table above
3. **Scorecards:** In the STYLE tab, set the metric font color to `#365c96`

---

## Layout

```
+─────────────────────────────────────────────────+
│  IT Service Desk — Superintendent Dashboard     │
│  [Date range control]              School Year  │
├────────┬────────┬────────┬────────┬─────────────┤
│ TOTAL  │ OPEN   │ CLOSED │  SLA   │  AVG DAYS   │
│TICKETS │BACKLOG │THIS MO │COMPLY% │ TO RESOLVE  │
│  4,218 │  312   │  287   │ 91.2%  │    3.4      │
├────────┴────────┴────────┴────────┴─────────────┤
│                                                 │
│  BACKLOG TREND (Line — DailySnapshot source)    │
│  ───────────────────────────────────            │
│  Shows open ticket count over time              │
│                                                 │
├─────────────────────┬───────────────────────────┤
│                     │                           │
│  MONTHLY VOLUME     │  SLA BREACH RATE          │
│  (Grouped Bar)      │  (Line chart)             │
│  Created vs Closed  │  % breaching over time    │
│  by month           │                           │
│                     │                           │
├─────────────────────┴───────────────────────────┤
│                     │                           │
│  TICKETS BY TEAM    │  TICKETS BY LOCATION TYPE │
│  (Horiz. Bar)       │  (Donut chart)            │
│  Open backlog       │  Elementary / Middle /    │
│  per team           │  High / Other             │
│                     │                           │
├─────────────────────┴───────────────────────────┤
│  TOP ISSUE CATEGORIES (Horizontal bar, top 8)   │
│  ████████████  Hardware     (142)               │
│  █████████     Software     (98)                │
│  ███████       Network      (71)                │
│  ...                                            │
+─────────────────────────────────────────────────+
```

---

## Component Build Instructions

### Header

- Insert a **Rectangle** shape across the full page width
  - Fill: `#365c96` (Dark Blue), border radius: 0 (flush to edges)
- Place a **Text** element on top: "IT Service Desk — Superintendent Dashboard"
  - Font: white, bold, 18-20px
- Add a **Date range control** (Insert > Date range control)
  - Default: Auto (matches data range)
  - Applies to TicketData source
  - Style: set text color to white so it's visible on the dark background
- Add a text label showing the school year (white text, right-aligned)

**Filter bar (below header):**
- Insert a **Rectangle** shape: fill `#F9FAFB` (Light Gray), border: 1px `#E5E7EB`, full page width
- Place the date range control and any dropdown filters on this bar

---

### Row 1: Scorecards (5 KPI cards)

**Banner background:**
- Insert a **Rectangle** shape behind all 5 scorecards
  - Fill: `#365c96` (Dark Blue), border radius: 8px
- Place scorecards on top with transparent backgrounds

**All scorecards:**
- Metric font color: **white** (to contrast against the dark blue banner)
- Font size: metric value 24-28px bold, comparison text 11-12px
- **Comparison colors** (STYLE tab) — default green/red gets drowned out by the dark blue banner. Override to brighter variants:
  - Positive change: `#86EFAC` (light green)
  - Negative change: `#FCA5A5` (light red/pink)
- **Comparison date range:** Only applicable to date-scoped scorecards. Full school year scorecards (Total Tickets, Open Backlog, SLA Compliance, Avg Days to Resolve) have no meaningful previous period within the same data — skip comparison on those.

#### Total Tickets
- **Chart type:** Scorecard
- **Metric:** Record Count
- **Filter:** None

#### Open Backlog
- **Chart type:** Scorecard
- **Metric:** Record Count
- **Filter:** `IsClosed` = "Open"

#### Closed This Month
- **Chart type:** Scorecard
- **Metric:** Record Count
- **Comparison date range:** Previous period (compares to last month's closures)
- **Filters:**
  - `IsClosed` = "Closed"
  - `ClosedDate` is in the current month (use a date range filter or relative date filter)

#### SLA Compliance %
- **Chart type:** Scorecard
- **Metric:** AVG of `SLA Compliant`
- **Format:** Percent (0 decimal places)
- **Filters:**
  - `IsClosed` = "Closed"
  - `Has SLA` filter

#### Avg Days to Resolve
- **Chart type:** Scorecard
- **Metric:** AVG of `AgeDays`
- **Format:** Number (1 decimal place)
- **Filter:** `IsClosed` = "Closed"

---

### Row 2: Backlog Trend (area chart)

**Card container:**
- Insert a **Rectangle** behind the chart: fill `#FFFFFF`, border: 1px `#E5E7EB`, border radius: 8px

**Chart setup:**
- **Data source:** DailySnapshot
- **Chart type:** Time Series (not Area chart — Area chart forces a breakdown dimension)
- **Dimension:** `Date` (keep as Date type, not Year Month)
- **Metric:** `Open Tickets` (SUM or AVG — only one row per date so either works)
- **Purpose:** Shows whether backlog is growing or shrinking over weeks/months

**Style tab:**
- Enable **Show area** and **Smooth line**
- Line color: `#365c96` (Dark Blue), area fill auto-derives a lighter shade at ~20-30% opacity
- Disable **Show points** (removes data point markers for a cleaner look)
- Grid lines: `#F3F4F6` (very light gray)
- Axis labels: `#6B7280` (muted gray), 10-11px
- Chart title: left-aligned, bold, `#1F2937`, 14px

---

### Row 3 Left: Monthly Volume (bar chart)

**Card container:** Rectangle behind the chart: fill `#FFFFFF`, border: 1px `#E5E7EB`, border radius: 8px

- **Data source:** MonthlyVolume
- **Chart type:** Bar chart (vertical, grouped)
- **Dimension:** `Month` — if full month names overlap on the x-axis, create a calculated field `LEFT(Month, 3)` (gives "Jul", "Aug", etc.) and use that as the dimension instead
- **Metrics:** `Created` (SUM), `Closed` (SUM)
- **Sort:** `MonthOrder` ascending

**Style tab:**
- Created bars: `#365c96` (Dark Blue), Closed bars: `#22b2a3` (Teal)
- Grid lines: `#F3F4F6`, axis labels: `#6B7280`, 10-11px
- Rotate x-axis labels to 45 degrees if they still overlap after abbreviating
- Chart title: left-aligned, bold, `#1F2937`, 14px

---

### Row 3 Right: SLA Breach Rate (line chart)

**Card container:** Rectangle behind the chart: fill `#FFFFFF`, border: 1px `#E5E7EB`, border radius: 8px

- **Data source:** TicketData
- **Chart type:** Time series (line)
- **Dimension:** `ClosedMonth` — set its granularity to **Year Month** in the chart's DATA panel
- **Metric:** AVG of `SLA Breached`
- **Format:** Percent
- **Filters:**
  - `IsClosed` = "Closed"
  - `Has SLA` filter

**Style tab:**
- Line color: `#f1663c` (Orange — breach = warning)
- Add a reference line at your SLA target (e.g., 10% breach rate)
- Grid lines: `#F3F4F6`, axis labels: `#6B7280`, 10-11px
- Chart title: left-aligned, bold, `#1F2937`, 14px

---

### Row 4 Left: Tickets by Team (horizontal bar)

**Card container:** Rectangle behind the chart: fill `#FFFFFF`, border: 1px `#E5E7EB`, border radius: 8px

- **Data source:** TicketData
- **Chart type:** Bar chart (horizontal)
- **Dimension:** `TeamName`
- **Metric:** Record Count
- **Filter:** `IsClosed` = "Open"
- **Sort:** Record Count descending
- **Bars shown:** Top 10
- **Purpose:** Which teams carry the most open backlog

**Style tab:**
- Bar color: `#365c96` (Dark Blue) — disable **Color by dimension** if enabled (bars should be a single solid color, not a gradient of colors per team)
- Show data labels (count on bars)
- Grid lines: `#F3F4F6`, axis labels: `#6B7280`, 10-11px
- Chart title: left-aligned, bold, `#1F2937`, 14px

---

### Row 4 Right: Tickets by Location Type (donut chart)

**Card container:** Rectangle behind the chart: fill `#FFFFFF`, border: 1px `#E5E7EB`, border radius: 8px

- **Data source:** TicketData
- **Chart type:** Donut chart
- **Dimension:** `LocationType`
- **Metric:** Record Count
- **Filter:** `IsClosed` = "Open"
- **Purpose:** Distribution across Elementary / Middle / High / Other

**Style tab:**
- Segment colors: Elementary `#365c96`, Middle `#22b2a3`, High `#febb12`, Other `#f1663c`
- Show percentage labels on segments
- Chart title: left-aligned, bold, `#1F2937`, 14px

---

### Row 5: Top Issue Categories (horizontal bar)

**Card container:** Rectangle behind the chart: fill `#FFFFFF`, border: 1px `#E5E7EB`, border radius: 8px

- **Data source:** TicketData
- **Chart type:** Bar chart (horizontal)
- **Dimension:** `IssueCategoryName`
- **Metric:** Record Count
- **Sort:** Record Count descending
- **Bars shown:** Top 8
- **Filter:** None (shows overall volume by category)

**Style tab:**
- Bar color: `#22b2a3` (Teal)
- Show data labels (count values on bars)
- Grid lines: `#F3F4F6`, axis labels: `#6B7280`, 10-11px
- Chart title: left-aligned, bold, `#1F2937`, 14px

> **Data quality note:** Issue categories with inconsistent capitalization in iiQ (e.g., "Issue not listed" vs "Issue Not Listed") will appear as separate bars. This is a source data issue, not a Looker Studio issue.

---

## Why These Metrics

| Metric | What it tells leadership |
|--------|--------------------------|
| Open Backlog | "How much work is piling up?" |
| Closed This Month | "Is the team productive?" |
| SLA Compliance | "Are we meeting our commitments?" |
| Avg Resolution | "How long do staff wait for help?" |
| Backlog Trend | "Are things getting better or worse?" |
| Monthly Volume | "Is demand increasing?" |
| By Team | "Where are bottlenecks?" |
| By Location Type | "Which school levels need the most support?" |
| By Category | "What types of problems dominate?" |

---

## Troubleshooting

### Date range silently filtering data

When you connect a Google Sheets data source, Looker Studio assigns a **default date dimension** (likely `CreatedDate`). Every component is automatically filtered by the report's date range — even without a date range control. If unfiltered record counts are lower than expected, set the default date range to cover your full school year in File > Report settings.

### Date fields not recognized

If Looker Studio doesn't auto-detect date columns, manually set the type to **Date & Time** in the data source config. The TicketData dates are ISO format (`2025-02-05T14:30:00Z`). If setting the type alone doesn't work, create a calculated field using `PARSE_DATETIME` (see the TicketData Field Types section above for the formula).

### Blank/null values in charts

Many fields can be empty (no SLA, no team assigned, no location). Add a filter to exclude blanks if they clutter a chart:
- Filter: `FieldName` is not null AND `FieldName` != ""

### Monthly Volume uses a separate data source

The Monthly Volume chart uses the **MonthlyVolume** sheet (not TicketData) because charting Created and Closed in one bar chart requires two different date columns (`CreatedDate` and `ClosedDate`), which Looker Studio blends handle poorly. The MonthlyVolume sheet pre-calculates both counts per month, avoiding blend issues entirely.
