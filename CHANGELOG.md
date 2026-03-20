# Changelog — iiQ Tickets to Sheets

All notable changes to this project are documented here.

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
