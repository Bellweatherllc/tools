# CORE Tool Versions

Current version of every deployed file in this repo. Update the row whenever a delivery bumps a version, and add a line to the log below.

Versions are whole numbers. The version lives in the file's badge, never in the filename or the tab title. Version badges are always red, `#ef4444`.

Last updated: 2026-10-20

## Live tools

| Tool | File | Version |
|---|---|---|
| **Pipeline & Projects** | | |
| Project Pipeline | `project-pipeline.html` | v1135 |
| Projects Manager | `core-projects-manager.html` | v61 (`MGR_VERSION`) |
| Project Basis | `project-basis.html` | v39 |
| Scout Board | `scout-board.html` | v13 |
| **PM Workbook & Budget** | | |
| PM Workbook | `pm-workbook.html` | v82 |
| PM Budget Planner | `pm-budget-planner.html` | v66 |
| **Finance** | | |
| Revenue Tracker | `revenue-tracker.html` | v11 |
| WIP Tracker | `wip-tracker.html` | v30 |
| **Permits** | | |
| Permit & Zoning Tracker | `permit-zoning-tracker.html` | v91 |
| **Analyzers** | | |
| Project Hours Analyzer | `project-hours-analyzer.html` | v57 |
| People Analyzer | `people-analyzer.html` | v1 |
| **Process** | | |
| CP Process Scoring | `cp-process-scoring.html` | v6 |
| FBA Matrix | `fba-matrix.html` | v4 |
| **SharePoint & Admin** | | |
| SharePoint Sherpa | `sharepoint-sherpa.html` | v13 |
| Sherpa Lookup Tester | `sherpa-lookup-tester.html` | v1 |
| Core Admin | `core-admin.html` | v6 |
| Team Members | `team-members.html` | v23 |
| Core Roadmap | `core_roadmap.html` | v57 |
| **Reference** | | |
| BW Tool Index | `bw-tool-index.html` | v3 |
| CORE System Map | `core-system-map.html` | v10 |
| Brand Style Guide | `bellweather-style-guide.html` | v12 |

## Manuals

Each HTML manual is built from the `.md` file of the same name. The manuals have no version badge.

| Manual | Files |
|---|---|
| Pipeline Manual | `pipeline-manual.html` / `.md` |
| PM Tools Manual (Workbook & Budget Planner) | `pm-tools-manual.html` / `.md` |
| Revenue Tracker Manual | `revenue-tracker-manual.html` / `.md` |
| WIP Tracker Manual | `wip-tracker-manual.html` / `.md` |
| Team Members Manual | `team-members-manual.html` / `.md` |

## Retired

Removed from the repo. These files are gone from `main`; they are still in git history.

| File | Removed in |
|---|---|
| `Bellweather-Scope-Tool.html`, `design-deliverables.html` | System Map v6 |
| `design-intake.html`, `timecard-analyzer.html` | System Map v4 |
| CA writing tools: `ca-setup.html`, `CA_DataSources.html`, `da_extractor.html`, `estimate_extractor.html`, `plan_note_extractor.html`, `reference_ca_extractor.html`, `scope_outline_extractor.html`, `selections_extractor.html`, `lib/ca-session.js` | System Map v2 |

## Log

Newest first. One line per delivery.

- **2026-10-20** — Project Basis v38 → v39: guide-box 1/2/3 badges now colorize legacy "LOW/MED/HIGH" wording too, not just literal digits — fixes stale-looking plain-bold guide text on block configs saved before the ternary redesign, without needing to hand-edit the SharePoint config. Also: selecting the sample project in Assess now prefills ratings/notes from a mock assessment instead of staying blank, matching what Raw Results/Analysis already show for it.
- **2026-10-20** — Project Basis v37 → v38: reverted v37's demo-only restriction — Assess lists Lead-stage projects plus the sample again. The three-tab mismatch (Assess shows all submittable projects, Raw Results/Analysis show only assessed ones) is understood behavior, not a bug; it'll resolve naturally as real assessments get submitted.
- **2026-10-20** — Project Basis v36 → v37: Assess project list narrowed to the sample project only, matching Raw Results/Analysis (which already only show projects with submitted assessments — currently just the sample). Will widen back to real Lead-stage projects once real assessments are in use.
- **2026-10-20** — Project Basis v35 → v36: fixed the sample project disappearing once SharePoint sign-in succeeded — `fetchProjects()` was overwriting the project list wholesale, dropping the sample. It's now always present alongside the real list.
- **2026-10-20** — Project Basis v23 → v35: ternary rating control redesigned as a colored circular 1/2/3 popup; sample "Hartwell Residence" project wired into Assess/Raw Results/Analysis for demoing without SharePoint sign-in; Results renamed to Raw Results and stripped to unprocessed data only; Summary renamed to Analysis with an inline-pie layout, phase timeline bars, and a tagged Key Conclusions library; Assess panel gained Scope-of-Work bullets.
- **2026-09-29** — Team Members v23: Access screen renames PM Updater → PM Workbook and PM Budget Worksheet → PM Budget Planner (manual updated to match); sign-in footer badge was stuck on v21 in the old red, now stamped current and `#ef4444`.
- **2026-09-28** — Core Roadmap v57: moved from decimal versions (v0.56) to whole numbers; the build count carries on.
- **2026-09-28** — All version badges set to red `#ef4444`; every tool above bumped one version. New badges on People Analyzer (v1) and Sherpa Lookup Tester (v1). Revenue Tracker's tab title no longer carries the version. Permit & Zoning Tracker, Core Roadmap and FBA Matrix badges no longer show a stale hard-coded number before the script stamps them.
- **2026-09-28** — `versions.md` started.
- **2026-09-28** — BW Tool Index v2: added Finance (Revenue Tracker, WIP Tracker), PM Workbook & Budget, and Permits sections; added Scout Board, Project Hours Analyzer, FBA Matrix and the five manuals; removed the dead `fba_tool_JW` tile.
- **2026-09-28** — CORE System Map v9: Tool Index entry updated; `fba_tool_JW` gap node removed.
