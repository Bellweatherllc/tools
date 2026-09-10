# Revenue Tracker — Help & Repair Manual

*What this page shows, where its numbers come from, and what to do when something looks wrong — written for people, not programmers.*

*September 2026 edition*

## What this page is

Revenue Tracker joins two independent views of the same revenue, side by side: what the Pipeline's schedule *says should* happen, and what the monthly WIP review says *actually* happened. It lives at [a web address](https://bellweatherllc.github.io/tools/revenue-tracker.html); sign in with your Bellweather Microsoft account and it loads.

> Its close relative is [WIP Tracker](https://bellweatherllc.github.io/tools/wip-tracker.html), which answers a narrower, always-live question — is a signed job currently billed ahead of or behind its work — using a simpler formula and no stored history. Revenue Tracker is the wider, historical view; WIP Tracker is the current-moment one. Use the link in either tool's top bar to jump to the other.

### The forecast

Computed live, on every page load, straight from the Pipeline's schedule: each project's `EstimatedProjectValue` divided by its construction-phase duration gives a weekly revenue rate, spread across wherever that phase actually sits on the calendar. Nothing is stored for this half of the page — move a bar in the Pipeline and the forecast moves with it, so the two can never disagree.

Solid gold in the monthly bars means CA-signed work; the striped portion is projected (DA-signed, unsigned, or paused) — work that's expected but not yet contractually certain.

### The actuals

These come from the monthly WIP workbook that Joey and Ryan maintain on the Operations site — the percent-complete review that turns QuickBooks costs into earned revenue per job.

**Sync from WIP** (the button in the top bar) reads that workbook directly from SharePoint, matches its project names against the Pipeline's (the same fuzzy name-matching used elsewhere in CORE), and writes one row per job per month into the `CORE_Revenue` list. Ryan's (RS) sheets are treated as the record of truth; Joey's (JW) percent-complete is stored alongside for comparison, not blended in.

The page always displays from `CORE_Revenue`, never from the workbook itself — so it stays fast, and the history survives the workbook being reorganized or renamed later.

### Where each piece lives

| | Owns | Written by |
|---|---|---|
| `CORE_Projects` | What **will** happen — schedule, contract value | Keyed in the Pipeline only |
| `CORE_Revenue` | What **did** happen — actual costs, percent complete, earned revenue | Written only by Sync from WIP, frozen after that |

This page is where the two meet. Nothing it displays gets written back into either list beyond what Sync from WIP explicitly writes.

### The Δ drift flag

A **Δ** next to a project in the table means the WIP workbook's contract value has drifted from the Pipeline's `EstimatedProjectValue` for that job — usually a change order that was keyed into one system but not the other. It's a flag to go reconcile, not an error in the page.

## Running a sync

1. Click **Sync from WIP** in the top bar.
2. The dialog lists the workbook's monthly tabs — pick which ones to include.
3. It matches each sheet's project names against the Pipeline automatically; anything it can't match confidently is left for you to resolve by hand before continuing.
4. A preview shows exactly what will be written — new rows and changed rows — before anything touches SharePoint.
5. Confirm, and it writes to `CORE_Revenue`. Existing rows for a job/period/category are updated in place rather than duplicated.

Re-running a sync is safe — it never re-writes a month that hasn't changed, and it never touches `CORE_Projects`.

## Open question

`CORE_Revenue` currently lives on the main **BWCore** SharePoint site — the same open-access site as the Pipeline — where anyone with CORE access can read it and derive job-level margin from costs vs. contract. Whether that visibility is correct, or whether the list should move to **OperationsSecure** with the rest of the financial data, hasn't been decided yet. The move is trivial on the tooling side; it just needs a decision from Byron, Joey, and Ryan. Once it's settled, this note goes away.

## When something looks wrong

**The whole page says it failed to load.**
Usually a sign-in issue — reload and sign in again. If other CORE tools are also having trouble reaching `CORE_Projects` or `CORE_Revenue`, it's likely a wider SharePoint/Graph problem, not specific to this page.

**A project's forecast bar looks wrong.**
The forecast is entirely schedule-driven — check the project's construction phase dates directly in the Pipeline. If they're right there, they're right here; this page doesn't have its own copy to go stale.

**A project isn't picking up its WIP actuals.**
The workbook's project name probably didn't match closely enough during sync. Check the project's name in the Pipeline against its name in the WIP workbook — small spelling differences are usually the cause. Re-running Sync from WIP lets you resolve the match by hand.

**A Δ flag won't go away.**
It clears once the two values agree — either update `EstimatedProjectValue` in the Pipeline to match the true contract value, or correct the WIP workbook, whichever is wrong.
