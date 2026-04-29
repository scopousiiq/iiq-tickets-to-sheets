# Changelog — iiQ Tickets to Sheets

All notable changes to this project are documented here.

---

## v1.6.1 — Documentation: changelog backfill + dashboard chart map (2026-04-29)

### Added
- **Backfilled changelog entries** for v1.3.7, v1.5.1, and v1.6.0.
- **Dashboard chart-coverage tables** in `README.md` and `GUIDE.md` enumerating which analytics sheets render as Chart.js cards (28 sheets across 7 tabs) vs KPI badges (5 list-style sheets) vs the fixed KPI row computed directly from `TicketData`.

No code or behavior changes — `SCRIPT_VERSION` bumped to 1.6.1 so districts pick up the docs in the next version-check cycle.

---

## v1.6.0 — iiQ Sheets Telemetry rollout (2026-04-27)

### Added
- **`Telemetry.gs`** — client module that pings the iiQ-owned Telemetry Master once per successful trigger refresh. The payload contains a stable anonymous install UUID, project name, script version, iiQ instance hostname (e.g. `demo.incidentiq.com`), TicketData row count, and the names of the canonical analytics sheets present in the spreadsheet. No ticket data, no API tokens, no custom field values, no user-added sheet names.
- **`TELEMETRY_ENABLED` Config row** — stamped `TRUE` for new installs by `setupConfigSheet`. Districts can flip it to `FALSE` to opt out.
- **`TELEMETRY_URL` constant in Config.gs** — maintainer-managed `/exec` URL for the Telemetry Master Web App. Districts never see or edit this.
- **iiQ TELEMETRY section in the Instructions sheet** — documents what is sent, what is not sent, the opt-out flow, and how to re-enable.

### Changed
- **Automated polling now requires telemetry opt-in.** Every CLOCK trigger handler (`triggerOpenTicketRefresh`, `triggerNewTickets`, `triggerWeeklyFullRefresh`, `triggerDataContinue`, `triggerDailySnapshot`) now calls `enforceTelemetryGate()` as its first line. If `TELEMETRY_ENABLED` is not `TRUE`, the gate uninstalls every time-based trigger and the trigger returns without doing work. Manual menu items continue to work.
- `setupAutomatedTriggers` now calls `assertTelemetryEnabledForTriggers()` before creating any triggers — installs with telemetry off see a clear "Telemetry Required" dialog instead of silently creating triggers that would self-uninstall on first fire.

### Upgrade Notes
1. Update all scripts to v1.6.0
2. **Existing v1.5.x installs:** the `TELEMETRY_ENABLED` row does not appear in your Config sheet automatically. Add it manually with value `TRUE`, then re-run **iiQ Data > Setup > Setup Automated Triggers**. Without this step, your next trigger fire will uninstall all polling triggers.
3. New installs (post-1.6.0 Setup Spreadsheet) ship with `TELEMETRY_ENABLED=TRUE` by default — no action needed.

---

## v1.5.1 — Native Web-App Dashboard + ChartRegistry (2026-04-25)

### Added
- **Native Apps Script Web App dashboard** — `doGet()` entry point in `Dashboard.gs` and a full-page, brand-styled `DashboardApp.html` render a tabbed analytics view at a shareable `/exec` URL. Districts deploy once, paste the URL into the new `DASHBOARD_URL` Config row, and share the link with leadership without granting spreadsheet access.
- **`ChartRegistry.gs`** — declarative map of all 33 analytics sheets to chart specs (type, label/series columns, row mode). The dashboard iterates this registry, keeps entries whose sheet exists, reads the data, and ships a Chart.js-ready payload to the client. Adding a registered analytics sheet via the menu makes its chart appear automatically; deleting the sheet removes it on next refresh.
- **Fixed KPI row** computed directly from TicketData (Open, Closed, Avg Resolution, SLA Breach %) — always present.
- **Gold KPI badges** for list-style sheets (`AtRiskResponse`, `AtRiskResolution`, `StaleTickets`, `ReopenRate`, `FrequentFlyers`) — only the ones that exist are shown.
- **"Show Dashboard URL" menu item** — reads `DASHBOARD_URL` from Config and copies the link to the user.

### Fixed
- **`QueueTimeAnalysis` chart coverage gap** — sheet existed but was not registered, so it never appeared on the dashboard. Now registered as a horizontal-bar distribution chart in the SLA category.

### Changed
- Documentation: corrected sheet counts in `CLAUDE.md` and `README.md` (25 optional + 8 default = 33 total) and added `DevicesByRole` to the Device category table.

---

## v1.3.7 — Resolve IiqLocation custom field UUIDs to campus names (2026-04-20)

### Fixed
- **Custom field values stored as raw UUIDs for `IiqLocation`-type fields (EditorType 22).** The iiQ API returns the `Value` for these fields as a JSON-encoded array of LocationIds with no display names attached. The previous JSON-parse path returned primitive UUIDs as-is, so cells showed strings like `b8a72c…-0010` instead of the campus name. `Select`/`MultiSelect` fields had the same shape problem.
- New `getAllLocations()` helper in `ApiClient.gs` performs a site-scoped paginated fetch (`/v1.0/locations/all/{siteId}`) to bypass the user-scoped endpoint's permission cap.
- `buildCustomFieldLookupMaps()` in `TicketData.gs` inspects each configured field's `EditorType` and builds a UUID → name lookup (locations API for type 22, `Options` JSON for types 9/10).
- `enrichLocationMapsFromTickets()` fills any gaps from each batch's `ticket.Location` data as a fallback.
- `extractCustomFieldValue` and `formatCustomFieldEntry` accept a `lookupMap` so primitive IDs inside parsed JSON arrays are resolved instead of returned raw.

### Upgrade Notes
1. Update all scripts to v1.3.7
2. Run **iiQ Data > Ticket Data > Full Reload** to re-resolve custom field values on existing rows (requires removing triggers first)

---

## v1.3.4 — RequesterRole column + Devices by Role analytics sheet (2026-04-17)

### Added
- **RequesterRole column (AU)** — TicketData now captures the role of the ticket's "For" user (`ticket.For.Role.Name`). Possible values: `Student`, `Staff`, `Agent`, `Guest`, `No Access`, `iiQ Administrator`. 47 columns total. The value is already present in every `searchTickets` API response — no extra API calls needed.
- **DevicesByRole analytics sheet** — New sheet under iiQ Data > Add Analytics Sheet > Device > Devices by Role. Shows device model metrics (Total, Open, Closed, Avg Resolution, Breach Rate) with a role dropdown (All / Student / Staff / Agent / Guest) so faculty/staff device issues can be viewed separately from student device issues.

### Changed
- `updateCustomFieldHeaders` migration function extended to cover the 46→47 column upgrade (adds `RequesterRole` header to existing sheets automatically on next data load).

### Upgrade Notes
1. Update all scripts to v1.3.4
2. Run **iiQ Data > Ticket Data > Full Reload** to populate `RequesterRole` on existing rows (requires removing triggers first — the column is blank for any rows loaded before this version)
3. Add the new sheet via **iiQ Data > Add Analytics Sheet > Device > Devices by Role**

---

## v1.3.3 — Duplicate ticket fix + MTD formula reliability (2026-04-17)

### Fixed
- **Duplicate ticket rows from triggerNewTickets** — `runNewTicketsCheck` was calling `updateConfigValue('TICKET_LAST_FETCH', ...)` without a subsequent `SpreadsheetApp.flush()`. The write buffered and never persisted before the trigger exited, so every 30-minute run re-read the stale timestamp and appended the same recent tickets again. Fixed by adding `SpreadsheetApp.flush()` after both `updateConfigValue` call sites in the function. Also added a `buildTicketIdMap` deduplication check before appending, so existing tickets are never double-written even if the timestamp stalls for any reason.
- **MTD "Created/Closed" columns showing wrong counts** — All analytics sheets (LocationBreakdown, IssueCategoryVolume, FunctionalAreaSummary, PriorityAnalysis, and others) used `TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD")` as COUNTIFS criteria. When TicketData date columns store actual date-time values (which GAS `setValues` with ISO 8601 strings auto-converts to), comparing with a text string is unreliable in Sheets. Replaced with `DATE(YEAR(TODAY()),MONTH(TODAY()),1)` — a numeric date value — for reliable date-to-date comparison.

### Upgrade Notes
1. Update all scripts to v1.3.3
2. Run **iiQ Data > Ticket Data > Full Reload** to clear the ~220 duplicate rows that accumulated before this fix (requires removing triggers first)
3. Regenerate affected analytics sheets via **iiQ Data > Add Analytics Sheet** to pick up the MTD formula fix

---

## v1.3.2 — Custom field value parsing (2026-04-17)

### Fixed
- **Custom field values pulling raw JSON with GUIDs** — For complex custom field types (dropdowns, user pickers, multi-selects, asset/location selectors), the iiQ API returns `Value` as a JSON-encoded string (e.g. `[{"Id":"...","Name":"..."}]` or `["guid","guid"]`). The extractor was writing that JSON directly into the cell. It now parses the JSON and emits the human-readable display text (Name/Text/Label/DisplayValue), falling back to the ID only when no display field is present. Multiple selections are joined with a comma.

---

## v1.3.1 — Small fixes and tweaks (2026-04-16)

### Added
- **Asset Tag column in FrequentFlyers requester detail** — Requester ticket details now include the asset tag alongside Model, making it easier to identify the specific device tied to each ticket.

### Fixed
- **DeviceReliability % of Model calculation** — Percentages were inconsistently calculating due to array/scalar broadcasting behavior in Google Sheets. Replaced `counts/totalForModel` with `MAP(counts, LAMBDA(c, c/totalForModel))` for reliable per-row division.

### Changed
- **DeviceReliability model dropdown** — K2 "Selected Model" dropdown now sorts by ticket count descending (highest-ticket models first), matching the left-side data table's default sort.

---

## v1.3.0 — Device Analytics + Improved Custom Fields Workflow (2026-04-16)

### Added
- **AssetId and AssetCategory columns (AP-AQ)** — TicketData now captures the iiQ AssetId (stable device identifier) and AssetCategory (e.g., "Chromebooks") for reliable device aggregation. 46 columns total.
- **FrequentFlyers sheet** — New analytics sheet with two sections: user frequent flyers (left, with min tickets + date range filters) and a requester ticket detail drill-down (right, dropdown-driven).
- **DeviceReliability model detail view** — Right side of sheet now has a model dropdown (sorted by ticket count) that reveals issue category breakdown with percentages and top 3 issue types per category.
- **CustomFields sheet + dropdown workflow** — Replaces the old "List Available Custom Fields" popup. A sheet lists all available custom fields from the iiQ API, and CUSTOM_FIELD_1/2/3 in Config get dropdown validation populated from that sheet. Use **iiQ Data > Setup > Refresh Custom Fields** to update the list.

### Changed
- TicketData sheet expanded from 44 to 46 columns
- `searchTickets` API payload now includes `Includes: ["Assets"]` to get full asset data on each ticket
- Menu: "List Available Custom Fields" → "Refresh Custom Fields"
- Menu: "Frequent Flyers (Users & Devices)" → "User Frequent Flyers" (device analysis is now part of DeviceReliability)

### Upgrade Notes
After updating scripts:
1. Run **Clear Data + Reset Progress**, then **Continue Loading** to populate AssetId/AssetCategory on existing rows
2. Run **iiQ Data > Setup > Refresh Custom Fields** to populate the new CustomFields sheet
3. Recreate FrequentFlyers and DeviceReliability sheets via **iiQ Data > Add Analytics Sheet** to pick up the new layouts

---

## v1.2.0 — Custom Field Extraction (2026-04-14)

### Added
- **Configurable custom field columns (AP-AR)** — Extract up to 3 custom field values from tickets. Configure field names in Config sheet (`CUSTOM_FIELD_1`, `CUSTOM_FIELD_2`, `CUSTOM_FIELD_3`), and the system auto-resolves them to iiQ CustomFieldTypeIds via the discovery API on the next data load.
- **`TICKET_COLUMN_COUNT` constant** — All hardcoded column count references replaced with a single constant in Config.gs. Safer for future column additions.
- **"List Available Custom Fields" menu item** — Discover what custom fields exist in your district directly from the iiQ Data > Setup menu.
- **Non-destructive Config migration** — Existing spreadsheets automatically gain the new custom field Config rows without re-running Setup.
- **Warnings in Verify Configuration** — Custom field resolution status (resolved, pending, not found) now displayed alongside blocking issues.

### Changed
- TicketData sheet expanded from 41 to 44 columns
- Config lock now includes custom field names (prevents mid-load changes)
- Unlock/Reset now clears resolved custom field IDs (forces re-resolution)
- `extractTicketRow()` accepts optional `customFieldIds` parameter (backward-compatible default)
- `verifyConfiguration()` now shows warnings separately from blocking errors

### Upgrade Notes
After updating scripts:
1. Configure custom field names in Config (`CUSTOM_FIELD_1`, `CUSTOM_FIELD_2`, `CUSTOM_FIELD_3`) — use **iiQ Data > Setup > List Available Custom Fields** to find valid names
2. Run **Clear Data + Reset Progress**, then **Continue Loading** to populate the new columns
3. Alternatively, new columns auto-extend on the next data load, but existing rows will have empty custom field values until reloaded

---

## v1.1.0 — Assigned Technician + Formula Fixes

- **New AssignedToUser columns (AN-AO)** — TicketData now includes the assigned technician/agent (`AssignedToUserId`, `AssignedToUserName`) separately from the ticket Owner. Owner (P-Q) is the staff member responsible for tracking the ticket; AssignedToUser is the person actually working it. 41 columns total.
- **Fixed TechnicianPerformance** — was aggregating by Owner (column Q), which included non-tech staff. Now correctly uses AssignedToUserName (column AO). Unassigned tickets are excluded.
- **Fixed StaleTickets formula** — "Days Since Update" and "Last Update" were calculated from StartedDate (column F) instead of ModifiedDate (column G), creating an inconsistency with the staleness filter threshold.
- **Fixed BacklogAging Sample Ticket formulas** — all five age bucket formulas were filtering/sorting by OwnerName (column Q) instead of AgeDays (column R).
- **Fixed ReopenRate "Days Open" column** — same column reference bug as BacklogAging (Q instead of R).

**Upgrade notes:** After updating scripts, run a full data reload to populate the new AN-AO columns. Recreate TechnicianPerformance, StaleTickets, BacklogAging, and ReopenRate sheets via the menu to pick up formula fixes.

---

## v1.0.2 — Resolution Aging (new default sheet)

- **New ResolutionAging sheet** — shows how long it takes to close tickets, bucketed by 0-15, 16-30, 31-60, 61-90, and 90+ days. Same layout as BacklogAging but for closed tickets. This is the "days to resolve" dashboard leaders want for at-a-glance performance visibility.
- Now a **default sheet** — created automatically by Setup Spreadsheet.

---

## v1.0.1 — Backlog Aging cross-tabs + version check

- **New Backlog Aging by Team / Location Type / Priority / Functional Area sheets** — see where backlog is accumulating across different dimensions.
- **HRSD module support** — pull data from HR Service Delivery in addition to Ticketing and Facilities.
- **Automatic version checking** — the sheet now checks GitHub daily and shows update status in the Config sheet (green = current, yellow = update available).
- Fixed MonthlyVolumeByFA date comparison bug.

---

## v1.0.0 — Device tracking, Looker Studio guides, concurrency

- **Device/Asset columns (AssetTag, ModelName, SerialNumber)** added to TicketData — 39 columns total.
- **New DeviceReliability analytics sheet** — which device models generate the most tickets?
- **Looker Studio dashboard guides** — step-by-step build guides for Superintendent and IT Director dashboards.
- **Concurrency control** — LockService prevents overlapping operations from corrupting data.
- **Looker Studio-safe values** — IsClosed uses "Open"/"Closed", breaches use 1/0 (no booleans).

---

All analytics sheets (31 total: 8 default + 23 optional) can be added or recreated via **iiQ Data > Add Analytics Sheet**.
