# IIQ Tickets to Sheets

Extract IncidentIQ ticket and SLA data into Google Sheets for analytics and Power BI dashboards. Built for K-12 school district IT departments.

## Quick Start

### 1. Create Your Spreadsheet

**Option A: Copy the Template (Fastest)**

[**Make a copy of the Google Sheets template**](https://docs.google.com/spreadsheets/d/1azxyYEefeAJ-oSmTKb199yzdULaaPvi_pjUydiC-fuA/edit?usp=sharing)

This template includes all sheets and formulas pre-configured. You just need to add your API credentials.

**Option B: Build from Scratch**

1. Create a new Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Copy all `.gs` files from the `scripts/` folder in this repository
4. Save and reload the spreadsheet
5. Click **IIQ Data > Setup > Setup Spreadsheet** to create all sheets automatically

### 2. Configure API Access

In the `Config` sheet, enter your IncidentIQ credentials:

| Setting | Value | Where to Find It |
|---------|-------|------------------|
| `API_BASE_URL` | `https://yourdistrict.incidentiq.com` | Your IIQ URL (the `/api` is added automatically) |
| `BEARER_TOKEN` | Your API token | IIQ Admin > Integrations > API |
| `SITE_ID` | Your site UUID (optional) | Only needed for multi-site districts |

### 3. Load Your Data

1. Reload the spreadsheet to see the **IIQ Data** menu
2. Click **IIQ Data > Setup > Verify Configuration** to check your settings
3. Click **IIQ Data > Refresh Teams** — authorize when prompted
4. Click **IIQ Data > Ticket Data > Continue Loading (Initial)** to start importing

> **Tip for large districts:** Instead of manually running "Continue Loading" repeatedly, set up a trigger:
> 1. Go to **Extensions > Apps Script > Triggers** (clock icon)
> 2. Add `triggerDataContinue` to run every 10 minutes
> 3. Check **IIQ Data > Ticket Data > Show Status** periodically until all years show "Complete"
> 4. You can leave this trigger enabled — it also helps complete open ticket refreshes
>
> SLA data is included automatically — no separate step needed.

### 4. Set Up Automated Refresh (Recommended)

Once your initial load is complete, set up triggers to keep data fresh:

1. Go to **Extensions > Apps Script**
2. Click the clock icon (Triggers) in the left sidebar
3. Add these triggers:

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
| `TicketData` | All ticket data with SLA metrics (35 columns) — your raw data source |
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
| `AtRiskQueue` | Open tickets approaching SLA breach (action needed!) |

### Additional Analytics (add via menu)

Use **IIQ Data > Add Analytics Sheet** to add any of these 13 optional sheets:

| Category | Available Sheets |
|----------|------------------|
| Volume & Trends | Seasonal Comparison (YoY), Temporal Patterns |
| Backlog & Quality | Stale Tickets, Reopen Rate |
| SLA & Response | First Contact Resolution, Response Distribution |
| Team & Staff | Technician Performance, Functional Area Summary |
| Location | Location Breakdown, Location Type Comparison |
| Issue & Requester | Issue Category Volume, Priority Analysis, Frequent Requesters |

> **Flexible & Customizable:** Districts can delete any analytics sheet and recreate it later via the menu. Default sheets (marked with ★ in the menu) can also be recreated if accidentally deleted.

## Common Issues

| Problem | Solution |
|---------|----------|
| "API configuration missing" | Fill in `API_BASE_URL` and `BEARER_TOKEN` in the Config sheet |
| HTTP 401 error | Your Bearer token expired — get a new one from IIQ |
| HTTP 429 / Rate limited | IIQ is throttling you — increase `THROTTLE_MS` in Config to 2000 or higher |
| Script timeout | Normal for large districts — just run "Continue Loading" again |
| Missing SLA data for some tickets | Expected — not all tickets have SLA policies assigned in IIQ |

## Documentation

- [**Implementation Guide**](GUIDE.md) — Detailed setup, formulas, and how everything works
- [**CLAUDE.md**](CLAUDE.md) — Technical reference for developers

## License

MIT — Free to use and modify for your district.
