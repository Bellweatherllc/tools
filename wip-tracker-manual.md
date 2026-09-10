# WIP Tracker — Help & Repair Manual

*What it does, where its numbers come from, and what to do when something looks wrong — written for people, not programmers.*

*September 2026 edition*

## What this page is

WIP Tracker answers one question, for every signed construction job: **are we billed ahead of the work, or behind it?**

It replaces the old monthly spreadsheet review. There's no percent-complete judgment call and no profit-margin math here — just a fixed formula, calculated fresh every time you open the page. Nothing is reviewed, nothing is typed in, and nothing goes stale between visits.

It lives at [a web address](https://bellweatherllc.github.io/tools/wip-tracker.html), like every CORE tool. Sign in with your Bellweather Microsoft account and it loads.

> **This tool only covers signed construction contracts** (`CA Signed` in the Pipeline). Projects still in design aren't part of this calculation — see *Why design-phase projects aren't shown*, below.

### How earned revenue is calculated

Two steps, applied to every signed job:

- **20% is counted earned the moment the Construction Agreement is signed** — it covers the design work that led up to it.
- **The remaining 80% spreads evenly across the construction schedule** already tracked in the Pipeline's Gantt. A job with a 40-week build earns 2% of that 80% for every week that passes (80 ÷ 40).

Add those two together, compare it to what's actually been invoiced, and the difference is the number that matters:

- **Invoiced more than earned** (shown in red, "over") — the client's been billed ahead of the work. Worth watching if costs come due before the job catches up.
- **Invoiced less than earned** (shown in green, "under") — the work is ahead of the billing. Money earned that hasn't been asked for yet.

### Why design-phase projects aren't shown

Earlier drafts of this idea considered a percent-complete formula for projects still in design. It was deliberately dropped — the 20% signing-day figure already accounts for design work, and running a second formula before a contract even exists added complexity without adding anything useful. If a project hasn't reached `CA Signed` in the Pipeline, it won't appear here.

## Where this data lives

Everything on this page is read live and recalculated on the spot — nothing is written back to SharePoint, and nothing is cached beyond your browser tab.

### Contract value & construction schedule

Both come from the **`CORE_Projects`** list on the BWCore SharePoint site — the same list the Pipeline and Projects Manager read and write. Specifically:

- **Contract value** is the `EstimatedProjectValue` field — the actual sale price once a job is signed.
- **Construction schedule** is the `Construction` phase inside `GanttData` — its start week and duration, exactly as drawn on the Pipeline's Gantt.

If either is missing for a project, the page flags it rather than guessing — see *Reading the flags*, below.

### Invoiced to date

This comes from a BuilderTrend export, dropped by hand into:

> `Operations → FINANCIAL → 1. WIP Reports & Job Costs → BT_InvoicingReports_forWIPTool`

The page always reads whichever file in that folder was **most recently saved**. It sums every invoice line marked `Paid` or `Pending/Sent` — invoices that have actually gone out — and ignores anything still marked `Draft`. Invoiced totals are matched to a project by name; a name that doesn't match closely enough is left out of the table and listed separately (see *"No BT match"*, below) rather than silently guessed at.

### Keeping the export current

To refresh the invoiced-to-date numbers:

1. In BuilderTrend: **Financial → Invoice → All Jobs → Export**.
2. Save the file as `Invoices_MM_DD_YYYY.xls` (today's date).
3. Drop it in the folder above — either loose, or filed into a year subfolder once one exists, the same way the AR Two-Week Reports and QuickBooks weekly drops are already filed.
4. Reload WIP Tracker, or click **Refresh** in the top bar.

There's no fixed schedule for this yet — do it as often as the numbers need to stay current. The page always shows the file's save date and, where available, the date the report itself says it was exported, so it's always clear how fresh the numbers are.

> This page does **not** keep its own history. For a record of what things looked like on a given date, the dated export file itself *is* the record — there's no separate snapshot system to maintain.

## Reading the table

| Column | What it means |
|---|---|
| Contract Value | `EstimatedProjectValue` from the Pipeline. |
| Build Progress | Weeks elapsed ÷ total construction weeks, from the Gantt. |
| Earned | 20% at signing, plus 80% × Build Progress. |
| Invoiced | Sum of `Paid` + `Pending/Sent` invoices matched to this job. |
| Gap | Invoiced − Earned. Red = billed ahead; green = work ahead of billing. |

### Reading the flags

A small **!** badge next to a value means the page couldn't compute something and is telling you rather than guessing:

- **"No schedule"** — the project has no `Construction` phase in its Gantt data, so earned revenue is showing the 20%-at-signing figure only, with nothing added for progress. Fix: add the construction phase in the Pipeline.
- **"No BT match"** — nothing in the latest BuilderTrend export matched this project closely enough by name. Check the *BuilderTrend jobs not matched* list further down the page — the job may be there under a different name, or genuinely hasn't been invoiced yet.
- **"No contract value"** — `EstimatedProjectValue` is empty in the Pipeline for this project. Nothing can be calculated until it's set.

## When something looks wrong

**The whole page says "Failed to load."**
Usually a sign-in or permissions issue. Reload the page and sign in again. If it keeps happening, the `CORE_Projects` list may be unreachable — check whether other CORE tools (the Pipeline, Projects Manager) are also having trouble.

**The source panel says it couldn't read the invoicing export.**
The page looks for a folder named `1. WIP Reports & Job Costs`, and inside it a folder matching `BT_InvoicingReports_forWIPTool`, on the Operations SharePoint site. If either has been renamed or moved, the page won't find it — check the folder path in the error message against what's actually in SharePoint.

**A project I expect to see isn't there.**
It's either not `CA Signed` yet in the Pipeline, or it is and something's off with the Pipeline data — check its stage in the Pipeline directly.

**The BuilderTrend job total looks wrong.**
Remember the invoiced figure only counts `Paid` and `Pending/Sent` rows — `Draft` invoices are deliberately excluded because they haven't gone out to the client yet. If a job has multiple BuilderTrend job codes (a separate demo contract, say), each one is matched independently, and one of them may be landing in the *unmatched* list instead of adding into the project's total.

## What this page deliberately doesn't do

- **No profit margin or GPM.** This is a billing-pace tool, not a profitability one. Job-by-job margin lives elsewhere.
- **No liability recognition.** It doesn't try to model what's owed on a job beyond the invoiced-vs-earned gap.
- **No design-phase calculation.** See *Why design-phase projects aren't shown*, above.
- **No stored history.** Every number is live. For history, see *Keeping the export current*, above.

These are deliberate simplifications, not gaps waiting to be filled in — the point of this tool is to answer one question clearly, not to become a second financial dashboard.
