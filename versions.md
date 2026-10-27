# CORE Tool Versions

Current version of every deployed file in this repo. Update the row whenever a delivery bumps a version, and add a line to the log below.

Versions are whole numbers. The version lives in the file's badge, never in the filename or the tab title. Version badges are always red, `#ef4444`.

Last updated: 2026-10-27

## Live tools

| Tool | File | Version |
|---|---|---|
| **Pipeline & Projects** | | |
| Project Pipeline | `project-pipeline.html` | v1135 |
| Projects Manager | `core-projects-manager.html` | v61 (`MGR_VERSION`) |
| Project Basis | `project-basis.html` | v42 |
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
| People Analyzer | `people-analyzer.html` | v11 |
| **Process** | | |
| CP Process Scoring | `cp-process-scoring.html` | v6 |
| FBA Matrix | `fba-matrix.html` | v4 |
| **SharePoint & Admin** | | |
| SharePoint Sherpa | `sharepoint-sherpa.html` | v13 |
| Sherpa Lookup Tester | `sherpa-lookup-tester.html` | v1 |
| Core Admin | `core-admin.html` | v6 |
| Team Members | `team-members.html` | v28 |
| Core Roadmap | `core_roadmap.html` | v57 |
| **Reference** | | |
| BW Tool Index | `bw-tool-index.html` | v4 |
| CORE System Map | `core-system-map.html` | v15 |
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
| People Analyzer Manual | `people-analyzer-manual.html` / `.md` |

## Retired

Removed from the repo. These files are gone from `main`; they are still in git history.

| File | Removed in |
|---|---|
| `Bellweather-Scope-Tool.html`, `design-deliverables.html` | System Map v6 |
| `design-intake.html`, `timecard-analyzer.html` | System Map v4 |
| CA writing tools: `ca-setup.html`, `CA_DataSources.html`, `da_extractor.html`, `estimate_extractor.html`, `plan_note_extractor.html`, `reference_ca_extractor.html`, `scope_outline_extractor.html`, `selections_extractor.html`, `lib/ca-session.js` | System Map v2 |

## Log

Newest first. One line per delivery.

- **2026-10-20** — Project Basis v41 → v42: sample data now self-heals against dimensions added via Admin mode (Presentation expectations, Sales Strategy, Quality, Timeline, Contract Scrutiny, Neighbor Relationship were showing blank since they didn't exist in the fixed default set). Added hand-written sample content for those six plus a generic fallback for anything added later, applied everywhere the sample project's data is read (Assess prefill, Raw Results, Analysis) — real-project data paths are untouched.
- **2026-10-27** — People Analyzer v10 → v11: Core Values descriptors replaced with the revised sheet (NEW_Core_Values_v2, 08-03-26); "Cares & Connects with People" is now "Cares & Connects". The Evaluate form shows "Revised Aug 3, 2026" beside the Core Values heading (one `CORE_VALUES_REVISED` constant), and the text export carries it. The Excel export's EOS-form label is unchanged.
- **2026-10-27** — People Analyzer v9 → v10: "HOD" relabeled back to **Admin** throughout the tool and manual (header switch, Admin panel, messages, name chip), matching Team Members and every other tool. Team Members v27 → v28: People Analyzer legend description no longer says "HODs". CORE System Map v14 → v15: People Analyzer entry says Admin.
- **2026-10-27** — Team Members v26 → v27: People Analyzer column back to **Member / Admin**, consistent with every other tool. People Analyzer v8 → v9: its messages and manual now say to set **Admin** in Team Members; inside People Analyzer that role is still called HOD.
- **2026-10-27** — People Analyzer v7 → v8: says why someone is or isn't an HOD — the Evaluate notice and the name-chip tooltip show the sign-in address, the roster entry matched (and whether by email or name), and the raw People Analyzer role. Each list now loads independently; a roster load failure is shown on screen and defaults to self-evaluation only instead of silently passing as "not an HOD" (or, with the column probe also failing, as everyone-is-HOD). Role value `HOD` is accepted as well as `admin`.
- **2026-10-27** — People Analyzer v6 → v7: "Admin" relabeled **HOD** throughout (header switch, panel, messages, name-chip role); the stored role value is still `admin`. New **SharePoint Access Check** in the HOD panel reads the lock-down straight from SharePoint — unique permissions and item-level read/edit on `PA_Evaluations` and `PA_Sessions`, the signed-in HOD's own submit / see-everyone / delete / session rights, and (for site Owners) who has access to each list, including whether staff can submit and see sessions. HODs whom SharePoint limits to their own evaluations get a warning at sign-in. Manual relabeled and its lock-down steps updated (Owner-only; HODs get Design).
- **2026-10-27** — Team Members v25 → v26: the People Analyzer column's top level shows as **HOD** in the Access dropdown and legend (stored value unchanged).
- **2026-10-27** — CORE System Map v13 → v14: People Analyzer entry says HOD.
- **2026-10-23** — People Analyzer v5 → v6: self-evaluations kept separate from the consensus. They're excluded from consensus, status and "N of M"; shown in a shaded column to the right of Consensus; tracked in a new **Self** column on the summary tables; listed as "Self-rating" (not counted) in text and Excel exports. Manual updated.
- **2026-10-23** — People Analyzer v3 → v5. v4: header rebuilt on the Pipeline's app-header pattern — Linen monogram, Swear title, Admin switch and gear, red badge, signed-in name chip, uppercase tabs with the gold underline, **?** manual link. v5: admin gating. A new `PeopleAnalyzerRole` column on `CORE_TeamMembers` (set on the Team Members Access tab) decides who is an admin. Non-admins get a self-evaluation: Evaluate locked to themselves with their name fixed, one per session, plus a **My Self-Evaluations** tab; Summaries, Admin and Manage are hidden and their actions refuse. The signed-in user is matched to the roster by email (any address in the Email column), then by exact name. Admins see a **self** tag on self-evaluations. Until the column exists, everyone is treated as admin and Admin says so. Manual updated, including item-level permissions for `PA_Evaluations`.
- **2026-10-23** — Team Members v24 → v25: Access tab lists **People Analyzer** (Member: own self-evaluations; Admin: everything) once the `PeopleAnalyzerRole` column exists.
- **2026-10-23** — CORE System Map v12 → v13: People Analyzer entry notes the `PeopleAnalyzerRole` gate.
- **2026-10-23** — People Analyzer Manual added (`people-analyzer-manual.html` / `.md`): using, admin, repair, and a data reference. People Analyzer v2 → v3: **?** link in the header opens it. BW Tool Index v3 → v4 and CORE System Map v11 → v12 list the manual.
- **2026-10-23** — People Analyzer v1 → v2: Admin panel no longer throws on open. Dead roster code removed — the tool no longer has any path that writes to or deletes from `CORE_TeamMembers`. Session on/off and drag order now save to SharePoint (new `PA_Sessions` columns `IsActive` Yes/No and `SortOrder` Number). Evaluations tie to people by team member Id (new `PA_Evaluations` column `EmployeeId` Number); older rows still match by name, and Admin has a one-click "Link" to write the Id onto them. Columns are probed on load, so the tool runs unchanged until they're added, and Admin lists what's missing. Evaluate lists active team members only; inactive people keep their history in Summaries. "Evals" shows "N of M" against the session's actual evaluators instead of a hard-coded 5. Tied votes resolve to the lower rating and show a dashed "split" chip with the vote breakdown (also noted in text and Excel exports). Names, comments and session names are HTML-escaped. Department filters come from the roster. `title=` tooltips replaced with `data-tip`. Unused detail modal and legacy export removed.
- **2026-10-23** — CORE System Map v10 → v11: People Analyzer no longer listed as writing `CORE_TeamMembers`; `PA_Sessions` / `PA_Evaluations` notes name the new columns.
- **2026-10-20** — Project Basis v40 → v41: all four sample assessments (Alex Rivera/Design, Sam Chen/PM, Jordan Blake/Estimating, Priya Nair/Structural Eng.) now have a score and comment on all twelve dimensions — Priya's previously covered only Engineering, with the other eleven silently missing rather than blank. Confirmed Raw Results and Analysis both show all four respondents fully on every dimension, and the Assess-tab prefill has no blank notes.
- **2026-10-20** — Project Basis v39 → v40: project selection now persists across Assess/Raw Results/Analysis — pick a project in one tab and it stays selected (with content auto-loaded) when you switch to the others, instead of resetting to "Select a project..." every time.
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
