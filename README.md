# iiQ Tickets to Sheets

Extract Incident IQ ticket and SLA data into Google Sheets for analytics and Power BI dashboards. Built for K-12 school district IT departments.

## Quick Start

### 1. Create and Set Up Your Spreadsheet

**Option A: Copy the Template (Fastest)**

[**Make a copy of the Google Sheets template**](https://docs.google.com/spreadsheets/d/1azxyYEefeAJ-oSmTKb199yzdULaaPvi_pjUydiC-fuA/edit?usp=sharing)

This template includes all sheets and formulas pre-configured. Skip to step 2.

**Option B: Build from Scratch**

1. Create a new Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Copy all `.gs` files from the `scripts/` folder in this repository
4. Save and reload the spreadsheet
5. Click **iiQ Data > Setup > Setup Spreadsheet**
6. When prompted, enter the school year (e.g., `2023-2024` for historical data, or leave blank for the current year)

> Setup creates all sheets and configures analytics formulas for the correct date range automatically. Each spreadsheet holds one school year's data.

### 2. Configure API Access

In the `Config` sheet, enter your Incident IQ credentials:

| Setting | Value | Where to Find It |
|---------|-------|------------------|
| `API_BASE_URL` | `https://yourdistrict.incidentiq.com` | Your iiQ URL (the `/api` is added automatically) |
| `BEARER_TOKEN` | Your API token | iiQ Admin > Integrations > API |
| `SITE_ID` | Your site UUID (optional) | Only needed for multi-site districts |
| `MODULE` | `Ticketing` or `Facilities` (optional) | Dropdown in Config sheet — defaults to Ticketing |

Then click **iiQ Data > Setup > Verify Configuration** to confirm everything looks good.

### 3. Load Your Data

1. Click **iiQ Data > Refresh Teams** — authorize when prompted
2. Click **iiQ Data > Ticket Data > Continue Loading (Initial)** to start importing
3. The script runs for ~5.5 minutes then pauses — run "Continue Loading" again to resume

> **Tip for large districts:** Instead of manually running "Continue Loading" repeatedly, set up a trigger:
> 1. Go to **Extensions > Apps Script > Triggers** (clock icon)
> 2. Add `triggerDataContinue` to run every 10 minutes
> 3. Check **iiQ Data > Ticket Data > Show Status** periodically until complete
>
> SLA data is included automatically — no separate step needed.

### 4. Set Up Automated Refresh (Recommended)

Once your initial load is complete, set up triggers to keep data fresh:

**Easy Setup (Recommended):**
Click **iiQ Data > Setup > Setup Automated Triggers** to create all triggers automatically.

**Manual Setup:**
Go to Extensions > Apps Script > Triggers (clock icon) and add these triggers:

| Function | Schedule | What It Does |
|----------|----------|--------------|
| `triggerDataContinue` | Every 10 minutes | Continues any in-progress loading (initial or open refresh) |
| `triggerOpenTicketRefresh` | Every 2 hours | Starts open ticket + SLA refresh |
| `triggerNewTickets` | Every 30 minutes | Catches newly created tickets |
| `triggerDailySnapshot` | Daily at 7:00 PM | Saves backlog metrics for trending |
| `triggerWeeklyFullRefresh` | Weekly, Sunday 2 AM | Full reload to catch deleted tickets |

> **About `triggerDataContinue`:** This is your "keep things moving" trigger. It automatically continues initial data loading OR open ticket refresh — whichever needs work. Once both are complete, it does nothing. You can leave it enabled permanently.

**How fresh is the data?**
- Open ticket status and SLA timers: Updated every 2 hours
- New tickets: Appear within 30 minutes
- Deleted or corrected tickets: Caught on weekly refresh

## What You Get

### Data Sheets (always created)

| Sheet | What It Shows |
|-------|---------------|
| `TicketData` | All ticket data with SLA metrics and device info (39 columns) — your raw data source |
| `Teams` | Team directory — add your Functional Area labels here |
| `DailySnapshot` | Historical backlog counts for trend analysis |

### Default Analytics (created by Setup)

| Sheet | What It Shows |
|-------|---------------|
| `MonthlyVolume` | Tickets created vs. closed by month, with closure rate |
| `BacklogAging` | Open tickets grouped by age (0-15, 16-30, 31-60, 61-90, 90+ days) |
| `TeamWorkload` | Open tickets, created/closed this month, and aging by team |
| `SLACompliance` | Monthly SLA breach rates and average response/resolution times |
| `PerformanceTrends` | "Are we getting better?" — key metrics over time |
| `AtRiskResponse` | Open tickets approaching Response SLA breach |
| `AtRiskResolution` | Open tickets approaching Resolution SLA breach |

### Additional Analytics (add via menu)

Use **iiQ Data > Add Analytics Sheet** to add any of these 18 optional sheets:

| Category | Available Sheets |
|----------|------------------|
| Volume & Trends | Seasonal Comparison (YoY), Temporal Patterns |
| Backlog & Quality | Stale Tickets, Reopen Rate |
| SLA & Response | First Contact Resolution, Response Distribution, Response Trends, Queue Time Analysis, Queue Time by Team, Queue Time Trend |
| Team & Staff | Technician Performance, Functional Area Summary |
| Location | Location Breakdown, Location Type Comparison |
| Issue & Requester | Issue Category Volume, Priority Analysis, Frequent Requesters |
| Device | Device Reliability |

> **Flexible & Customizable:** Districts can delete any analytics sheet and recreate it later via the menu. Default sheets (marked with ★ in the menu) can also be recreated if accidentally deleted.

## Common Issues

| Problem | Solution |
|---------|----------|
| "API configuration missing" | Fill in `API_BASE_URL` and `BEARER_TOKEN` in the Config sheet |
| HTTP 401 error | Your Bearer token expired — get a new one from iiQ |
| HTTP 429 / Rate limited | iiQ is throttling you — increase `THROTTLE_MS` in Config to 2000 or higher |
| Script timeout | Normal for large districts — just run "Continue Loading" again |
| Missing SLA data for some tickets | Expected — not all tickets have SLA policies assigned in iiQ |

## Documentation

- [**Implementation Guide**](GUIDE.md) — Detailed setup, formulas, and how everything works
- [**CLAUDE.md**](CLAUDE.md) — Technical reference for developers

### Looker Studio Dashboard Guides

Step-by-step build guides for creating Looker Studio dashboards from your data:

- [**Superintendent Dashboard**](lookerStudioDashboardSamples/superintendent-dashboard.md) — Strategic overview for district leadership (single-page, 6 KPIs, trend charts)
- [**IT Director Dashboard**](lookerStudioDashboardSamples/it-director-dashboard.md) — Operational dashboard for daily execution and team management

## License

MIT — Free to use and modify for your district.
