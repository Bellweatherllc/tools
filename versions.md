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
| Levels of Development | `project-basis.html` | v46 |
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
| People Analyzer | `people-analyzer.html` | v25 |
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
| BW Tool Index | `bw-tool-index.html` | v5 |
| CORE System Map | `core-system-map.html` | v17 |
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

- **2026-11-10** — Levels of Development v45 → v46: fixed "Scope saved failed... Graph 400" — the field it was reading/writing (`ScopeSummary`) doesn't exist on CORE_Projects; switched to the real field, `DesignScope` (confirmed against core-projects-manager.html's own FIELDS schema). Added Target Investment editing (Admin-gated, same as Scope of Work) — Idea Low/High and Target Investment Amount now writable here, not just viewable.
- **2026-11-10** — Levels of Development v44 → v45: pasted DA scope text no longer explodes into one bullet per wrapped line — Scope of Work now splits on blank lines into distinct items and merges single line breaks within an item back into flowing prose. Added a Target Investment field (reads Projects Manager's existing TargetInvestmentAmount/IdeaLow/IdeaHigh fields — single amount once committed, Idea Low–High range before that) next to Address/CA Value; read-only here since Projects Manager owns that data.
- **2026-11-10** — Levels of Development v43 → v44: Scope of Work is now editable (Admin-gated, same toggle as block/dimension config) instead of read-only. Saving PATCHes the same CORE_Projects `ScopeSummary` field Projects Manager reads/writes — no separate copy — so a tool admin (Byron/Joey) can set up a project's scope before everyone else assesses it, without a trip to another tool.
- **2026-11-10** — Project Basis renamed "Levels of Development" in the header (file stays `project-basis.html` — it already has a live URL), v42 → v43: subtitle updated from "Pre-DA Assessment Tool" to "Project Assessment Tool" since Assess's project filter now includes DA Signed projects alongside Lead, not just pre-DA ones — a project signed the same morning was previously invisible once it left Lead stage.
- **2026-10-20** — Project Basis v41 → v42: sample data now self-heals against dimensions added via Admin mode (Presentation expectations, Sales Strategy, Quality, Timeline, Contract Scrutiny, Neighbor Relationship were showing blank since they didn't exist in the fixed default set). Added hand-written sample content for those six plus a generic fallback for anything added later, applied everywhere the sample project's data is read (Assess prefill, Raw Results, Analysis) — real-project data paths are untouched.
- **2026-11-03** — People Analyzer v24 → v25: Summaries renamed **Session Summary** and narrowed to one session at a time — it starts on "Pick a session…" (no "All Sessions", no "All" button; department buttons filter and clear on a second click), and a person shows only that session with a **Full review history ›** link. The cross-session **Trajectory** (and Download All Sessions) moves to the top of the person's **Review History**; the imported-reviews section leaves Session Summary.
- **2026-11-03** — People Analyzer v23 → v24: tabs reordered to Evaluate · Summaries · Review History · Compensation. **Review History** is now each person's complete record — reviews done in People Analyzer (tagged In tool / Annual, with each scorer's ratings, status and an Open in Summaries link) alongside imported spreadsheets (tagged Imported), newest first. Review History sits above Add History on the tab. The section inside a person's Summaries view is renamed "Imported reviews".
- **2026-11-03** — People Analyzer v22 → v23: every menu starts empty — Evaluate session (no longer auto-picked), Compensation Show, Review History Show and the Annual Review Scorers session each open on a "Pick…" prompt. Quick guide: "Two kinds of review" rewritten as two short lists instead of a table; the manual page now styles its tables.
- **2026-11-03** — People Analyzer v21 → v22: History tab split into **Add History** (import) and **Review History**, which shows one person at a time, picked from a Show list grouped by department with review counts.
- **2026-11-03** — People Analyzer v20 → v21: department compensation gains a **By year** layout — years down the side (newest open, click to expand), people under each year, and pay lines across (salary, total, bonus available/earned, PTO, then each benefit line) with a total per year, so people can be compared within a year. Both layouts group people by their Team Members **Role** (job title) with a subtotal per role. `Role` is read from `CORE_TeamMembers` only after checking the column exists.
- **2026-11-03** — People Analyzer v19 → v20: department compensation table — click a name (or **Expand all**) to open that person's breakdown under their row: salary, each benefit line, total, bonus earned and PTO, aligned to the department's years. "open ›" goes to the person's own page.
- **2026-11-03** — People Analyzer v18 → v19: one-line instruction on the Compensation tab — pick anyone on the roster (even with no review yet), click + Add a year, enter their salary.
- **2026-11-03** — People Analyzer v17 → v18: evaluations saved under a name that matches no one on the roster (e.g. a misspelling) can be linked to the right person from the Admin panel — a "This is…" picker and **Link** per name, writing only `EmployeeId`; the stored name is unchanged.
- **2026-11-03** — People Analyzer v16 → v17: SharePoint connection indicator moved left, beside the signed-in name chip; reads "Not connected" instead of "Offline" now that offline mode is gone.
- **2026-11-03** — People Analyzer v15 → v16: offline mode removed from the sign-in screen — the tool only runs signed in. New admin **Compensation** tab: per person, salary and total comp by year (chart + side-by-side table with year-over-year change, bonus, PTO); per department, everyone on one chart (total comp or salary) with a department total that flags partial years. **+ Add a year** / **Edit** save one compensation record per person per year in `PA_ReviewHistory`; imported years are never changed (edits override them; Undo restores). The annual summary page now reads and saves that same yearly record (its Salary line is the year's salary; the separate "New salary" box is gone). Manual rewritten as a short quick guide built around periodic vs annual reviews. CORE System Map v16 → v17.
- **2026-10-28** — People Analyzer v14 → v15: annual review, part 2 — the **summary page** (admins). Per staffer per annual session: HOD / second seat / self side by side with a Final (agreed) column (auto-fill where HOD and second seat agree); wins, opportunities, goals; compensation & benefits pre-filled from the last imported review, editable, totalled in integer cents; new salary; PTO; bonus available/earned (entered, not calculated); meeting close with transcript paste or .txt/.vtt/.srt drop, a prepared Claude prompt to copy (no eNPS, no comp figures), and the close text. Save, Save & mark complete / Reopen, Print / PDF. Stored on the annual record in `PA_ReviewHistory`; transcript as a `transcript.txt` attachment. Entry points: Summary column in the annual session table and a button on each person's annual block. `apiCall()` / `AI_PROXY_URL` scaffold ready for the Azure Function proxy (empty — not deployed).
- **2026-10-27** — People Analyzer v13 → v14: annual review, part 1. Sessions can be created as **Annual review** (`SessionType` on `PA_Sessions`; until that column exists, a name containing "Annual" counts). Admin panel → **Annual Review Scorers** assigns each staffer's HOD and second seat (admins), stored as admin-only records in `PA_ReviewHistory`. In an annual session, admins see only the staffers they score and the form shows their role; everyone rates V/TO (Core Focus, 10-Year Target, 3-Year Picture & 1-Year Plan); the staffer answers eNPS (0–10, required). V/TO, eNPS and role are saved in a new `AnnualData` column on `PA_Evaluations`. Summaries tag HOD / 2nd seat / self and add a V/TO grid; the company eNPS shows for the annual session, never individual answers. One score per scorer per staffer per annual review.
- **2026-10-27** — People Analyzer v12 → v13: new admin-only **History** tab — drag-and-drop import of past review workbooks into a new `PA_ReviewHistory` list. Each sheet is one review; the parser finds sections by their grey headings rather than fixed cells, handles +/− and 1–5 layouts, the pre-2023 value set, legacy and threaded comments, wins/opportunities/outcomes/notes, bonus and compensation panels, and keeps anything else with its cell reference. Preview lets you fix person/month before saving; re-imports are skipped. History shows grouped by person and inside each person's Summaries view. The Access Check covers `PA_ReviewHistory` (must not inherit; no member/staff access — it holds compensation). Tested on the Sean, Hayley and Stephanie workbooks: 21 reviews. CORE System Map v15 → v16: `PA_ReviewHistory` added.
- **2026-10-27** — People Analyzer v11 → v12: Evaluate top section rebuilt as one compact card — Evaluator, Session and Team Member in a row, one-line hint. The evaluator is set from the sign-in (roster name, else Microsoft account name) and can't be edited; offline mode still asks. Members see their own name as both Evaluator and Team Member, with no dropdown. A lone open session is picked automatically. The large guidance box is gone, and the member note no longer shows sign-in diagnostics (the name-chip tooltip still does).
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
