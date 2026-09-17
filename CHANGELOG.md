# Changelog — iiQ Tickets to Sheets

All notable changes to this project are documented here.

---

## v1.8.4 — Analytics sheets no longer split a category across two rows (2026-09-17)

### Fixed
- **Name-keyed analytics sheets showed duplicate rows carrying identical numbers.** `UNIQUE` compares text case-sensitively but `COUNTIFS` does not. Every rollup grouped with `UNIQUE` and aggregated with `COUNTIFS` therefore gave a district holding two capitalizations of the same catalog entry — `Issue Not Listed` and `Issue not listed` — two rows in which *every* figure was the combined total of both spellings, inflating the column total by a full copy of that group.
  - Reported on `IssueCategoryVolume`, but the same defect was latent in **20 formulas across 15 sheets**: `IssueCategoryVolume`, `IssueTypeVolume`, `TeamWorkload`, `LocationBreakdown`, `FunctionalAreaSummary`, `PriorityAnalysis`, `TechnicianPerformance`, `LocationTypeComparison`, `FrequentRequesters`, `FrequentFlyers`, `DevicesByRole`, `DeviceReliability`, and the three `BacklogAgingBy*` sheets.
  - Grouping now runs through a new `distinctIgnoringCase()` helper, which groups on the lowercased value and labels each group with the first spelling encountered — matching how the counts already behaved. Dropdown sources were fixed too; previously a dropdown offered both spellings and either one filtered to the merged total.
  - Device models were affected as often as issue categories (`MacBook Air` / `Macbook Air`, `iMac` / `Imac`), since neither catalog forces consistent capitalization.
  - Verified against synthetic data: all 20 rewritten formulas produce output identical to the originals wherever no spelling collision exists, and collapse to one correct row where one does.

### Upgrade Notes
- **Formula-only change — no API calls and no reload.** Re-create the affected analytics sheets from **iiQ Data > Add Analytics Sheet** to pick up the corrected formulas, or copy the current template.
- Row counts on these sheets will drop where duplicates were present. The remaining row carries the total that was previously printed twice; it was never a per-row undercount.

---

## v1.8.3 — SLA batch size lowered to stay under the API's filter cap (2026-09-14)

### Fixed
- **`TICKET_BATCH_SIZE` now defaults to 1250 instead of 2000, so SLA columns populate during bulk load.** Each batch's ticket IDs are sent to `POST /tickets/slas` as one filter per ticket. That endpoint returns **HTTP 500** once a request carries more than roughly 1300 filters, so at the old 2000 default every SLA request in a bulk load failed. The failure was silent: `fetchSlaForTicketIds` catches the error, logs `SLA_ERROR`, and returns an empty map, so tickets still wrote to the sheet with every SLA column blank.
  - Verified against a non-production demo site: 1250 filters returned complete data in ~7s, while 1300, 1500, and 2000 each returned HTTP 500 after ~22s.
  - `SLACompliance`, `AtRiskResponse`, and `AtRiskResolution` all read from those columns, so they were empty too.

### Upgrade Notes
- **New copies of the template pick this up automatically.** The 1250 default is written into the Config sheet at setup.
- **Existing spreadsheets keep whatever `TICKET_BATCH_SIZE` their Config sheet already holds** — the default only applies to a blank cell, and the value is locked once loading starts. Districts seeing blank SLA columns should move to the current template.
- Do not raise `TICKET_BATCH_SIZE` above 1250. Values past ~1300 reintroduce the HTTP 500.

---

## v1.8.2 — Issue Type Volume sheet + dashboard chart legibility (2026-08-27)

### Added
- **`IssueTypeVolume` analytics sheet** — a drill-down partner to `IssueCategoryVolume`, grouping by issue type rather than category. Same metrics (Open, Created/Closed MTD, Avg Resolution, Breach Rate), plus a **Category** column showing which category each type rolls up to, so a row is readable without cross-referencing the category sheet.
  - Add it via **iiQ Data > Add Analytics Sheet > Issue & Requester > Issue Type Volume**.
  - Capped at the top 50 rows with a **Category Filter** control. Districts commonly define hundreds of issue types against a couple dozen categories; without a cap and a filter, a full-catalog sweep would run thousands of full-column `COUNTIFS` on every recalculation. Narrowing to one category is both the natural drill-down and the fast path.
  - No API change — `IssueTypeName` was already extracted into `TicketData` column AA. This is a formula-only sheet.

### Changed
- **Dashboard chart cards are now sized from their data instead of a fixed 280px.** Horizontal bar charts grow with their category count (300–1600px) and claim the full grid row once they exceed eight bars; vertical charts get a modest height increase and go full width past fourteen categories. The dashboard content area widened from 1400px to 1600px.
- **Each bar now carries its value**, drawn beside the bar so the number is readable without hovering. Implemented as a small inline Chart.js plugin rather than adding a third-party datalabels package. Suppressed on stacked charts, where the labels would overlap.
- Long category labels are truncated on the axis with the full text still shown in the tooltip.
- Tooltips format per series rather than per chart, so a chart holding both a rate and a count no longer renders the count as a percentage.
- Stacked bar segments carry a hairline separator, and the dynamic-series palette expanded from six colors to eight.

### Fixed
- **Dense charts were silently dropping category labels.** Cards were a fixed height in a fixed two-column grid, so Chart.js quietly skipped axis labels to make them fit — the chart looked complete while showing only a fraction of its labels. Horizontal charts are now sized to fit every label and no longer skip any.
- **Four charts flattened one of their series into an unreadable line.** `PerformanceTrends` (both charts), `PriorityAnalysis`, and `DeviceReliability` each plotted a 0–1 rate or a single-digit day average on the *same* axis as raw ticket counts, so the smaller series collapsed onto the baseline. On `PerformanceTrends` the shared percent formatter also labeled a days axis in percent, reading up to "650%". These series now get their own axis via a new `axis: 'right'` field in `ChartRegistry`.
- **Percentage axes were collapsed to just 0% and 100%.** Integer-precision ticks were applied to every value axis, including rate axes spanning 0–1, leaving two gridlines. Percent axes no longer force integer ticks.
- **A single "N/A" cell discarded its entire chart row.** Any non-numeric value in one series dropped the whole row, so an issue category with no closed tickets (and therefore no average resolution) also lost its Open count. Rows now survive as long as one series has a real number, with the missing value charted as a gap.
- Stacked charts with more than six series repeated colors within a single bar, making adjacent segments indistinguishable.

### Upgrade Notes
- Existing dashboards pick up the chart changes on reload — no action needed, no data affected.
- `IssueTypeVolume` is added from the menu; it is not created automatically by `Setup Spreadsheet`.
- On a large `TicketData`, the unfiltered `All` view takes noticeably longer to settle on first calculation than a single-category view. Selecting a category is the fast path.

---

## v1.8.1 — Instructions sheet table formatting (2026-08-27)

### Changed
- **Tabular rows on the Instructions sheet now render in Roboto Mono**, so their column separators actually line up instead of drifting under a proportional font. Applied to any line containing two or more pipe characters, which covers the automated-triggers table and the custom field config table.
- The custom field setup step is now a table rather than space-aligned text, so it benefits from the same treatment.
- Section headers and dividers are matched by content rather than by hardcoded row numbers, and all three style classes apply as batched `getRangeList` calls (four formatting calls instead of one per row).

This matches the Instructions formatting already used across the sibling `iiq-*-to-sheets` projects.

### Upgrade Notes
Re-run **iiQ Data > Setup > Setup Spreadsheet** to regenerate the Instructions tab with the new formatting. No data is affected — the Instructions sheet is documentation only.

---

## v1.8.0 — Location custom fields in TicketData (2026-08-27)

### Added
- **Location custom fields as TicketData columns (AV-AZ)** — up to five custom fields defined on your district's *locations* can now be pulled into TicketData, alongside the three ticket custom fields already supported. Set them via the new `LOCATION_CUSTOM_FIELD_1`–`LOCATION_CUSTOM_FIELD_5` keys in Config.
  - This covers fields that live on the school/building rather than the ticket: district site codes, regions, building numbers, hardware support areas. Previously these were unreachable — the ticket custom field extraction only ever looked at values set on the ticket itself.
  - Values are joined onto each ticket by its `LocationId`. The grain is unchanged: still exactly one row per ticket.
  - New columns are appended at the end of the header row, so existing column references, analytics formulas, and Google Data Studio / Domo / Power BI field mappings do not shift.
- **`Refresh Custom Fields` now lists both entities** — the CustomFields sheet has a new leading `Entity` column marking each field as `Ticket` or `Location`. The Config dropdowns are scoped per entity, so a location field can't be selected into a ticket slot (where it would never resolve) or vice versa.
- **Five location slots rather than three** — location values are resolved once per run from a single paginated call, so extra slots cost columns rather than API calls. Districts commonly define more location fields than ticket fields.

### Changed
- `TicketData` is now 52 columns (was 47).
- **Custom field slots now take a `CustomFieldTypeId`, pasted from the `CustomFields` sheet, instead of a name picked from a dropdown.** The id is what the loader actually stores and uses, and it is unambiguous — a district can define two different fields sharing the same display name, which a name-based selection cannot distinguish. Field names are still accepted, so existing configurations keep working. Applies to both `CUSTOM_FIELD_1/2/3` and `LOCATION_CUSTOM_FIELD_1-5`.
- A pasted id needs no lookup, so configuring fields by id makes **no** `/custom-fields/for/*` call during resolution.
- `Verify Configuration` reports location custom field resolution status, and warns specifically that Ticket and Location fields are separate namespaces when a name doesn't resolve.

### Fixed
- **`Refresh Custom Fields` failed outright on districts with many custom fields**, reporting `The data validation rule has more items than the limit of 500`. Config cells no longer carry list validation at all, so the cap does not apply.
- **The `CustomFields` sheet listed the same field many times over.** The `/custom-fields/for/*` endpoints return one row per field-to-filter-set mapping rather than one per field, so a single field could appear dozens of times with an identical name and `CustomFieldTypeId`. Rows are now deduplicated per field type, which in testing more than halved the row count on districts with large custom field catalogs.
- **`Clear Data + Reset Progress` failed with `Sorry, it is not possible to delete all non-frozen rows`.** Because each spreadsheet holds a single school year, clearing that year matches every data row, and a completed load leaves the grid sized exactly to the data — so the delete covered every non-frozen row, which Sheets rejects. A spare row is now added before the delete.
- **A pasted `CustomFieldTypeId` was silently rejected**, resolving to `NOT_FOUND` because the value was only ever looked up as a field name. Ids are now recognized and used directly.
- **Custom field errors blamed API credentials regardless of cause.** A failure while writing the sheet reported "Make sure API_BASE_URL and BEARER_TOKEN are configured" even though the API call had already succeeded. Fetch failures and sheet failures now report separately.
- The `EditorType` labels on the `CustomFields` sheet were misaligned from type 4 onward against iiQ's actual editor type values (type 4 showed as "Dropdown" when it is "Number"; Select and MultiSelect were also wrong).

### Performance
- Cost per run is **one paginated locations sweep** (one API call for any district under 500 locations), and only when at least one location field is configured. Nothing scales with ticket volume — a district with hundreds of thousands of tickets pays exactly the same as one with a few thousand.
- Editor types are cached in Config (`LOCATION_CF_TYPE_CACHE`) for 30 days, so a steady-state load makes **no** `/custom-fields/for/location` call. The cache is filled for free whenever definitions are already being fetched (field name resolution, `Refresh Custom Fields`), keyed by field UUID so pointing a slot at a different field invalidates only that entry.
- `IiqLocation`-typed location fields build their name map from the locations payload already in hand rather than issuing a second locations call.
- Measured overhead on a 5.5-minute execution budget: well under a second for a few hundred locations, and around a second for districts approaching a thousand — under 0.5% of the run, and less than a third of a single ticket batch.

### Upgrade Notes
- Existing sheets keep working untouched. With no location fields configured, the five new columns stay blank and no extra API calls are made.
- Config rows are added automatically on next load or on `Verify Configuration`. Sheets upgraded from an earlier version pick up the new rows even though they already have the ticket custom field rows.
- The `CustomFields` sheet is rebuilt on the next `Refresh Custom Fields` run to add the `Entity` column. That run also strips the old list validation from the Config custom field cells and adds a note explaining the id paste.
- If an admin changes a location field's *type* in iiQ (rare), run `Refresh Custom Fields` to correct the cached editor type immediately rather than waiting out the 30-day window. `Verify Configuration` shows the cache age.
- To start using location fields on a sheet that already has data loaded: run `Refresh Custom Fields`, pick your fields in Config, then `Clear Data + Reset Progress` and reload. The fields are part of the config lock, so they can't be changed mid-load.

### Known Limitations
- Location custom fields of type `IiqUser` (e.g. a "Principal" field) emit the raw user UUID rather than a name. Resolving them would require a bulk user fetch; the ticket custom fields have the same limitation.

---

## v1.7.1 — TechnicianPerformance: split breach into Response/Resolution (2026-07-14)

### Changed
- **The single `Breach Rate` column on the TechnicianPerformance sheet is now two columns: `Response Breach %` (col J, from ResponseBreach/AF) and `Resolution Breach %` (col K, from ResolutionBreach/AI).** Each is computed over the same denominator — tickets closed in the selected window — so a ticket that breaches both SLAs is no longer double-counted (the old combined rate could exceed 100%).
  - The data table grew from 10 to 11 columns (A–K). Sort controls moved to M/N and window controls to P/Q. Open (col C) and Aged 30+ (col F) positions are unchanged, so the web-app dashboard chart is unaffected.
  - Sort Col# now accepts 1–11 (10 = Response Breach, 11 = Resolution Breach).

### Upgrade Notes
Recreate the TechnicianPerformance sheet via **iiQ Data > Add Analytics Sheet > Team & Staff > Technician Performance** to pick up the split breach columns.

---

## v1.7.0 — TechnicianPerformance time window + SLA times (2026-07-14)

### Added
- **Window Start (O2) and Window End (P2) controls on the TechnicianPerformance sheet** — slice technician metrics to any period (a single week, a sprint, a quarter) for week-by-week or period-over-period comparison. Defaults to school-year-start → today (blank End = today, blank Start = all time).
  - **Created / Closed** count tickets dated within the window (previously fixed to month-to-date).
  - **Avg Resolution / SLA times / Breach Rate** cover tickets *closed* within the window.
  - **Open / Aged 30+** are reconstructed *as of the window end* from CreatedDate (E) and ClosedDate (H) — a ticket is open as-of the window end if it was created on/before that date and either never closed or closed after it. Age is measured at the window end. These no longer read the live "as of now" AgeDays column, so past periods reflect the state at that time.
- **Avg SLA Response (hrs) and Avg SLA Resolution (hrs) columns** on the TechnicianPerformance sheet — per-technician SLA-clock times from ResponseActual (AE) and ResolutionActual (AH), stored in minutes and shown in hours, averaged over tickets closed in the window and filtered to tickets that actually had an SLA clock (`>0`).
  - Distinct from the existing **Avg Resolution (days)**, which is *calendar* time-to-close (from AgeDays). The new SLA columns are the SLA-policy clock (business hours, paused-time aware). Both are kept.
  - The data table grew from 8 to 10 columns (A–J); sort controls moved to L/M and window controls to O/P. Open (col C) and Aged 30+ (col F) positions are unchanged, so the web-app dashboard chart reads them without modification.
- The `Created`/`Closed` headers dropped their `(MTD)` suffix.

### Upgrade Notes
Recreate the TechnicianPerformance sheet via **iiQ Data > Add Analytics Sheet > Team & Staff > Technician Performance** to pick up the Window controls and SLA-time columns.

---

## v1.6.3 — FrequentFlyers Location filter (2026-04-29)

### Added
- **Location dropdown (E2) on the FrequentFlyers sheet** — populated dynamically from your district's unique location names (column Q, hidden). Select a specific campus to see only its frequent flyers, or leave it at `All` to see everyone.
  - The `Selected Requester` dropdown (H2) also respects the location filter — it only shows requesters who appear in the filtered result.
  - Combines with the Role filter (D2) from v1.6.2: both criteria are always applied, using `"*"` as a COUNTIFS wildcard when set to "All" so the branch count stays at 2 regardless of filter state.
  - The `allReqs` requester pool uses exact FILTER logic (4-branch: applyRole × applyLocation) for accuracy; all COUNTIFS carry both `roleCriterion` and `locationCriterion` at all times.

### Upgrade Notes
Recreate the FrequentFlyers sheet via **iiQ Data > Add Analytics Sheet > Issue & Requester > Frequent Flyers** to pick up the new Location control.

---

## v1.6.2 — FrequentFlyers Role filter (2026-04-29)

### Added
- **Role dropdown (D2) on the FrequentFlyers sheet** — filters the user frequent-flyer list by requester role. Options: `All`, `Exclude Agent`, `Student`, `Staff`, `Agent`, `Guest`. Defaults to `Exclude Agent` so internal agent self-tickets don't inflate frequent-flyer results.
  - The dropdown also narrows the `Selected Requester` picker (H2) to only show requesters who match the active role filter, so the drilldown stays coherent.
  - All four ticket-count metrics (Total, Open, Closed, Avg Resolution) apply the role filter consistently using a `roleCriterion` LET variable. `Exclude Agent` maps to the COUNTIFS criterion `<>Agent`.

### Upgrade Notes
Recreate the FrequentFlyers sheet via **iiQ Data > Add Analytics Sheet > Issue & Requester > Frequent Flyers** to pick up the new Role control.

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

## v1.0.0 — Device tracking, Google Data Studio guides, concurrency

- **Device/Asset columns (AssetTag, ModelName, SerialNumber)** added to TicketData — 39 columns total.
- **New DeviceReliability analytics sheet** — which device models generate the most tickets?
- **Google Data Studio dashboard guides** — step-by-step build guides for Superintendent and IT Director dashboards.
- **Concurrency control** — LockService prevents overlapping operations from corrupting data.
- **Google Data Studio-safe values** — IsClosed uses "Open"/"Closed", breaches use 1/0 (no booleans).

---

All analytics sheets (31 total: 8 default + 23 optional) can be added or recreated via **iiQ Data > Add Analytics Sheet**.
