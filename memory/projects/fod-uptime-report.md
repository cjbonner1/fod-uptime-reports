# FoD Uptime Report 2026

**Status:** Active, updated monthly
**File:** FoD_Uptime_Report_2026.xlsx
**Build Script:** build_update_feb.py (Python/openpyxl, generates entire workbook from scratch)
**Owner:** Chance Bonner (cbonner@opentext.com)
**Created:** February 19, 2026
**Last Updated:** September 16, 2026 (scope change: the official number covers core services only, Aug 2026 forward)

## What It Is
Monthly uptime report for Fortify on Demand (FoD). Tracks outages and maintenance and calculates uptime percentage. Delivered as an Excel workbook. The report states measured availability only and does not assert a service level target.

## Data Source
- URL: https://status.fortify.com/history (Hund.io status page)
- Scrape via browser automation (Claude in Chrome MCP); direct WebFetch is blocked by egress proxy
- Events listed in reverse chronological order with: event type, date/time range (GMT), affected service(s)

## Key Design Decisions
1. **2026-only scope**: A prior report covered Sep 2024 - Sep 2025. This file starts fresh at Jan 2026 to avoid contradicting that report.
2. **Outage-only uptime**: Only "Outage" events count toward uptime. Maintenance, Service Degradation, and DR Exercises are tracked but excluded. Rationale: during degradation the portal is still available.
3. **De-duplication**: When Portal + API go down simultaneously (start times within 60 seconds), they count as ONE incident with MAX duration (not sum).
4. **Formatting**: Matches a sample report provided by Chance. Aptos Narrow font, 24pt title, dark blue (#2F5496) table headers, 0.0000% uptime format.
5. **Core-services scope, Aug 2026 forward** (ruled by Chance 2026-09-16): the official uptime figure covers ONLY the customer-facing Tenant Portal and API across AMS, EMEA, APAC, SGP, EU and FedRAMP. Add-on services (Vulncat, Debricked, SAST Aviator) are still reported in full on the Regional Breakout tab, below a tracking divider, but do not roll into the official number. The CALCULATION METHOD IS UNCHANGED, same formula, Outage-only, same 60-second de-duplication. Only the set of components feeding the number changed. Months before August 2026 are deliberately left as previously issued so a client comparing an older filed copy sees the same figures.
6. **No SLA verdict** (same ruling): the workbook previously graded itself against a 99.9% target that had no source outside these notes. The target and the MEETING/BELOW TARGET status were removed from every tab. Report the number, the customer judges it against their own agreement.

## Workbook Structure (14 tabs)
1. **2026 Executive Summary** - Yearly rollup table (12 months), YTD totals, Availability Summary, scope statement, Key Observations in column G.
2. **Regional Breakout** - Per-component uptime percentages by month/YTD.
3. **Jan 2026** - Completed month detail: summary block, daily outage table (worst-first), services affected, detailed incident write-ups.
4. **Feb 2026** - Completed month detail (same structure as Jan).
5. **Mar 2026** - Completed month detail.
6. **Apr 2026** - Completed month detail.
7. **May 2026** - Completed month detail.
8. **Jun 2026** - Completed month detail.
9. **Jul 2026** - Completed month detail.
10. **Aug 2026** - Completed month detail.
11. **Sep 2026** - Stub tab (in progress, partial through Sep 16; 2 outages, 8 min to date, core services only).
12. **2026 Incident Data** - Raw fact table. Every event (outage + maintenance) in reverse chronological order.
13. **Bucket Mapping** - Reference table: 13 buckets mapping component strings to region/service categories.
14. **Notes & Methodology** - Calculation docs, update instructions, formatting conventions. Also serves as a handoff playbook.

## Current Numbers
| Month | Outages | Minutes | Uptime |
|-------|---------|---------|--------|
| Jan 2026 | 4 | 5 | 99.9888% |
| Feb 2026 | 3 | 9 | 99.9777% |
| Mar 2026 | 7 | 57 | 99.8723% |
| Apr 2026 | 4 | 25 | 99.9421% |
| May 2026 | 2 | 4 | 99.9910% |
| Jun 2026 | 2 | 4 | 99.9907% |
| Jul 2026 | 8 | 39 | 99.9126% |
| Aug 2026 | 1 | 105 | 99.7648% |
| Sep 2026 (partial thru 9/16) | 2 | 8 | n/a (in progress) |
| **YTD (Jan-Aug)** | **31** | **248** | **99.9291%** |

Scope note: from Aug 2026 these figures cover core services only. August's single incident was the FedRAMP Portal and API outage on Aug 29 (105 min, de-duplicated from 103 + 105). The Aug 2026 Vulncat outage cluster (4580 component-minutes) is still reported on the Regional Breakout tab and no longer rolls into the official number. Jan through Jul are unchanged from the versions previously issued and still include all services.

## Uptime Formula
- Monthly: 1 - (Outage Minutes / Total Minutes in Month)
- Daily: 1 - (Outage Minutes / 1440)
- YTD: 1 - (Sum Outage Min for Completed Months / Sum Total Min for Completed Months)
- Minutes per month: Jan=44640, Feb=40320, Mar=44640, Apr=43200, May=44640, Jun=43200, Jul=44640, Aug=44640, Sep=43200, Oct=44640, Nov=43200, Dec=44640

## Monthly Update Process
1. Scrape https://status.fortify.com/history for the completed month
2. Append raw events to the raw_2026 list in the build script
3. Finalize the month's detail tab (summary, daily table, services, incident write-ups)
4. Create next month's stub tab
5. Update Executive Summary (fill month row, update YTD, refresh Key Observations, update the Availability Summary block)
6. Apply de-dup logic consistently (60-second grouping window, max duration)
7. Run the script to regenerate the workbook

## Build Script Usage
The build script (build_update_feb.py or most recent build_*.py) is self-contained Python/openpyxl. To update:
- Add new raw events to the raw_2026 list at the top
- Adjust months_complete list
- Update Key Observations text
- Update "Last updated" in Notes & Methodology section
- Run: python3 build_update_feb.py

## Regions / Buckets
FedRAMP (Portal, API), AMS (Portal, API), EMEA (Portal, API), APAC (Portal, API), SGP (Portal, API), Vulncat (Global), EU SAST Aviator, AMS SAST Aviator

## Notes
- The Notes & Methodology tab in the workbook itself contains full update instructions
- Browser scraping required: status.fortify.com is blocked by direct fetch; use Claude in Chrome MCP (navigate + get_page_text)
- The build script is the fastest path to regeneration; single Python file produces the entire workbook
