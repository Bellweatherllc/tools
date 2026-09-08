# Bellweather CORE Tools

Single-file HTML tools for Bellweather Design-Build's internal operations system (CORE). Byron builds and maintains these.

**This repo is PUBLIC.** GitHub Pages serves it, and view-source exposes every file. Never put an API key, token, secret, or connection string in any file here. The Azure `clientId` and `tenantId` are public identifiers and are fine. Reference documentation lives in the private `Bellweatherllc/core-docs` repo, not here.

---

## What these tools are

- Standalone single-file HTML. All JS and CSS inline. CDN dependencies only. **No build step.**
- Data lives in SharePoint Lists, reached through the Microsoft Graph API.
- Auth is Azure AD OAuth via MSAL. `lib/sherpa-auth.js` is the shared sign-in helper every tool loads — it wraps the app registration config and exposes `SherpaAuth.init/isSignedIn/signIn/signOut/getToken/getAccount/graph()`. It caches to `sessionStorage` and requires msal-browser to load first.
- Deployed via GitHub Pages at `bellweatherllc.github.io/tools/`.

Three SharePoint locations — don't confuse them:

| Site | Access | Holds |
|---|---|---|
| `BWCore` | Open | Operational lists — projects, tools, schedule, permits, capacity |
| `OperationsSecure` | Restricted | Sensitive — HR/People Analyzer, OPS cash/payroll/bank |
| `/sites/BWC` → `Customers` drive | Staff only | Project document library |

---

## Delivery discipline — follow every time

- **Bump the version on every delivery.** Whole number.
- **The version lives in the badge, never the filename.** In the pipeline it's a single `CORE_VERSION` constant; badge elements in the markup are empty and stamped by an IIFE right after the constant, so there's no stale flash.
- **Run `node --check` on the extracted script block before delivering.** Every time.
- **Use single-occurrence-assert string replacement for edits** — fail loudly rather than silently editing the wrong match.
- **Verify the delivered bytes**, not a `VALID` result that may have checked an unmodified file.
- After deploying, the browser still caches the Pages file. Bump `?v=N` on the URL to see the new build.
- **Merge feature branches into `main` and push automatically once work is committed there — don't wait for an explicit "merge" or "deploy" instruction each time.**

---

## Hard rules

**No SharePoint column may be named `Status`** — it's reserved. Build state is `BuildStatus`.

**Graph `$select` / `$expand` must name every field explicitly, including lookups.** An omitted field returns empty with no error. This has cost real debugging time.

**Tooltips: never `title="…"` or `el.title =`.** Use `data-tip`, plus `data-tip-d` for a second detail line, routed through the shared floating-tooltip engine. Icon-only controls keep an `aria-label`.

**SharePoint is authoritative from first load. `localStorage` is a read cache only.**

**Clocks and derived risk signals are computed live, never stored.** A stored snapshot cries wolf and destroys trust.

**Doc links are GUID-anchored** (`Doc.aspx?sourcedoc={GUID}`), and projects tie to their folder by **driveItem id**. Both survive phase-folder moves. Path-based URLs break — don't use them.

**Phase folders are discovered at runtime, never hard-coded.** The `TEMPLATE` folder in SharePoint is the source of truth. If SharePoint diverges from a doc, SharePoint wins. Legacy folder aliases exist (`1. DAs` ← `1. IDEA & FDDA`) and both names must be recognized — never rename an existing project's folder to match current convention.

**Project folder names are `LASTNAME - Brief description`, but humans stray.** Use runtime folder lists and fuzzy matching, not filename parsing.

**CORE never writes to Buildertrend.** The pipeline decides invoice dates; a human carries them into BT. That division of labor is deliberate, not a sync bug to fix.

**Sherpa and Doc Linker never overwrite, delete, rename, or auto-archive.** Ambiguity surfaces for human confirmation. `SAMPLE Scope Outlines` is reference-only. `FORM`-prefixed files are templates — flag, don't file.

**API keys never live in client code in production.** The paste-your-key settings modal is a removable scaffold. Wrap Claude calls behind a swappable `apiCall()` so switching to the Azure Function proxy is a one-line change per tool.

**All money arithmetic in integer cents.**

---

## Conventions

**SharePoint list naming: `CORE_` prefix plus the list name.** `CORE_Projects`, `CORE_TeamMembers`, `CORE_Tools`, `CORE_ToolAssignments`. New lists follow the same shape — `CORE_Permits`, `CORE_Milestones`.

The older `{DomainPrefix}_{ToolType}_{ListName}` pattern in `CORE_Glossary.md` (`PRJ_A_Sessions`, `PRM_T_Permits`) is **retired**. Don't use it for anything new. If you hit a list already named that way, leave it alone — renaming a live SharePoint list breaks every tool pointing at it.

**Brand.** Logos `.svg` only (colorways `-B`, `-W`, `-Linen`), fonts `.woff2` only, both served from this repo. Palette: Limestone `#282828`, Cotton `#F6F2EE`, Navy `#263860`, Linen `#E5E1D8`. Brand type is TT Norms Pro + Swear.

The pipeline and Gantt still run their own Navy `#1f2e3e` / Gold `#d8b64f` with Fraunces + DM Sans. That migration is a separate future track — **don't "fix" it mid-task.**

---

## Repo layout

**Deployed tools stay at the repo root. Flat, on purpose.**

The URL is the contract. `bellweatherllc.github.io/tools/<name>.html` is referenced by bookmarks, by cross-tool links, by `pipeline-manual.md`, and — critically — by the **MSAL redirect URI registered in the Azure app registration**. Moving or renaming a deployed file breaks all of them, and the Azure side can lock users out.

**Never move or rename a file that already has a live URL.** Not to tidy up, not to match a newer convention. If a name is wrong, it stays wrong.

**Folders that are safe**, because they aren't page URLs: `lib/`, `fonts/`, `logos/`, `azure-function-proxy/`.

**Retired files go in `archive/`.** A tool that's been superseded, an experiment that didn't land, an old version kept for reference. Nothing links to these, so moving them breaks nothing — and it separates "live" from "was live once," which is the thing a flat root can't show. Never archive a file that something still links to; check first.

**New tools may be born in a folder** (`permits/tracker.html` serves fine from Pages). The cost only exists for files that already have a URL in the wild.

**Naming — new files only:** all lowercase, hyphens not underscores, domain prefix where it groups things (`permit-`, `project-`, `ca-`). The repo currently mixes both styles (`core_roadmap.html` beside `core-projects-manager.html`); that's historical and stays. Don't retrofit.

**Keep the index current.** `bw-tool-index.html` and `versions.md` are what a flat repo has instead of a folder tree. Update them when a tool is added, renamed at birth, or archived.

---

## The pipeline's deploy convention

`project-pipeline-plus.html` is the testing copy. When it's ready, it's renamed to `project-pipeline.html` and becomes the live one.

---

## Working style

Byron wants direct implementation guidance — column names, types, exact code — not an explanation of the reasoning behind it. Be concrete. Skip re-derivation.

When something is genuinely ambiguous (see the list-naming note above), ask rather than pick silently.

---

## Where to look for more

The private `Bellweatherllc/core-docs` repo holds the reference documentation: glossary, roadmap, decision log, tool registry schema, SharePoint routing, budget workbook reference, permit clock design, cost codes. If a question here needs business context rather than code context, that's where it lives.
