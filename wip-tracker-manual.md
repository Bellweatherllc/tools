# WIP Tracker — Help & Repair Manual

*What it does, where its numbers come from, and what to do when something looks wrong — written for people, not programmers.*

*September 2026 edition*

## What this page is

WIP Tracker answers one question, for every job that's collected money: **are we billed ahead of the work, or behind it?**

It replaces the old monthly spreadsheet review. There's no profit-margin math here — just a formula (or a manual figure, where the formula doesn't apply yet), calculated fresh every time you open the page.

It lives at [a web address](https://bellweatherllc.github.io/tools/wip-tracker.html), like every CORE tool. Sign in with your Bellweather Microsoft account and it loads.

> **Scope:** only **DA Signed** and **CA Signed** jobs show up here. Leads (no agreement signed yet) are excluded outright — nothing's been collected on a job that isn't under contract in some form.

### How earned revenue is calculated

- **CA Signed jobs** use a formula: **20% is counted earned the moment the Construction Agreement is signed**, covering the design work that led up to it. **The remaining 80% spreads evenly across the construction schedule** already tracked in the Pipeline's Gantt — a 40-week build earns 2% of that 80% for every week that passes.
- **DA Signed jobs** have no schedule to hook a formula onto, so they start with **no earned figure at all** until someone sets one — see *Manual % overrides*, below. This is deliberate: better an honest blank than a guessed number nobody signed off on.
- **A manual override, when set, always wins** — see below. It replaces whichever of the above would otherwise apply.

Compare earned to what's actually been invoiced, and the difference is the number that matters:

- **Invoiced more than earned** (red, "over") — the client's been billed ahead of the work.
- **Invoiced less than earned** (green, "under") — the work is ahead of the billing.

## Manual % overrides

Click the **✎** next to any project's % Complete to set it by hand. This is for exactly the cases the formula can't handle:

- **A pre-CA job** — there's no formula yet, so this is the only way to give it an earned figure.
- **A CA-Signed job where the schedule says something the work doesn't.** Fike is the standing example: the Pipeline's construction phase reads 100% complete, but a large chunk of the work (windows) is still on backorder. Rather than drag the project's end date out to today — which would throw off its $/week average by inflating the apparent project length — set the override to what's actually true, say 90%, with a note explaining why.

An override replaces the formula entirely: earned becomes contract value × the percentage you set, full stop. It's marked **manual** in the table so it's never confused with a calculated figure, and its note (visible on hover) is the record of why a human overrode the math. Overrides persist until changed or cleared — they carry forward month to month rather than resetting, so update Fike's number as the real picture changes rather than re-entering it from scratch.

Overrides are stored in **`CORE_Config`** (key `wip_overrides`) — the same shared list the Pipeline already uses for things like its OPS cash-flow scenarios. No SharePoint schema change was needed to add this.

## Locking the month — Ryan's button, Joey's button

There's no single generic "lock" button. Instead there are two named chips near the top of the page — **Ryan** and **Joey** — each with its own status and its own button. A chip's button only works for that person: it checks who's actually signed in, not a name anyone could type, so Ryan can't lock Joey's slot and vice versa. If you're signed in as neither, both buttons are visibly greyed out — hover one to see why.

1. **Ryan** clicks his own **Lock as Ryan** button. His chip shows a ✓ and a timestamp. Joey's still shows "not locked." Nothing is final yet.
2. **Joey**, signed in as himself, clicks **Lock as Joey**. That fills the second slot, **finalizes the record**, and immediately opens the browser's print dialog so it can be **saved as a PDF**.

Order doesn't matter — whoever locks first, the record only finalizes once *both* slots are filled.

Clicking your own button again before the other person has locked just **updates your slot** with today's figures — it's still only one signature. The live page keeps recalculating after that — invoiced totals and schedules don't freeze — but the finalized lock is untouched by that. Once finalized, a banner appears with **View locked figures**, which switches the page to show exactly what was recorded, and **Back to live**, which returns to the current numbers.

**Locking again after it's already finalized** starts a brand-new two-slot cycle — both chips reset to "not locked," and both Ryan and Joey need to lock again before a new PDF comes out. That's expected for revising a month after something changes, not an error.

A new calendar month always starts with both slots empty — nothing carries over.

Locks are stored in `CORE_Config`, one row per month (key `wip_lock_YYYY-MM`, holding both reviewers' snapshots by name), so they don't compete for space with anything else and there's no limit on how many months of history accumulate.

> Reviewer matching is by first name on the signed-in Microsoft account (Ryan, Joey) — if either of their accounts doesn't display that first name for some reason, their button would never enable. Worth confirming once, then it's a non-issue going forward.

## Where the rest of this data lives

Everything else is read live and recalculated on the spot — nothing else is written back to SharePoint, and nothing is cached beyond your browser tab.

### Contract value & construction schedule

Both come from the **`CORE_Projects`** list on the BWCore SharePoint site — the same list the Pipeline and Projects Manager read and write. Specifically:

- **Contract value** is the `EstimatedProjectValue` field.
- **Construction schedule** is the `Construction` phase inside `GanttData` — its start week and duration, exactly as drawn on the Pipeline's Gantt. Only used for CA-Signed jobs without an override.

### Invoiced to date

This comes from a BuilderTrend export, dropped by hand into:

> `Operations → FINANCIAL → 1. WIP Reports & Job Costs → BT_InvoicingReports_forWIPTool`

The page always reads whichever file in that folder was **most recently saved**. It sums every invoice line marked `Paid` or `Pending/Sent` — invoices that have actually gone out — and ignores anything still marked `Draft`.

### Keeping the export current

1. In BuilderTrend: **Financial → Invoice → All Jobs → Export**.
2. Save the file as `Invoices_MM_DD_YYYY.xls` (today's date).
3. Drop it in the folder above.
4. Reload WIP Tracker, or click **Refresh** in the top bar.

## Reading the table

| Column | What it means |
|---|---|
| Stage | The project's current Pipeline stage. Gold badge = CA Signed. |
| Contract Value | `EstimatedProjectValue` from the Pipeline. |
| % Complete | The construction-schedule formula (CA-Signed only), or a manual override, or "not set." |
| Earned | Contract Value × % Complete. Blank until a % exists. |
| Invoiced | Sum of `Paid` + `Pending/Sent` invoices matched to this job. |
| Gap | Invoiced − Earned. Red = billed ahead; green = work ahead of billing. |

### Reading the flags

- **"No schedule"** (CA-Signed jobs) — no `Construction` phase in the Gantt, so earned is showing the 20%-at-signing figure only. Fix: add the construction phase in the Pipeline, or set a manual override.
- **"Not set"** (earlier-stage jobs) — no formula applies yet and no override has been set. Click **✎** to give it one.
- **"~" (weak match)** next to an invoiced figure — this job matched a BuilderTrend job code by a shared name fragment rather than a close full-name match. Worth a second look; it's shown, not hidden, so it stays checkable rather than silently guessed.
- **"No BT match"** — nothing in the latest export matched this project by name at all. Check the *BuilderTrend jobs not matched* list further down the page.
- **"No contract value"** — `EstimatedProjectValue` is empty in the Pipeline.

## When something looks wrong

**The whole page says "Failed to load."**
Usually a sign-in or permissions issue. Reload and sign in again. If it keeps happening, check whether other CORE tools are also having trouble.

**The source panel says it couldn't read the invoicing export.**
The page looks for `1. WIP Reports & Job Costs`, then `BT_InvoicingReports_forWIPTool`, inside FINANCIAL on the Operations SharePoint site. If either's been renamed or moved, check the error message's folder path against what's actually there.

**A BuilderTrend job total isn't showing up on a project.**
First check the *BuilderTrend jobs not matched* list at the bottom of the page — the job may be sitting there under a name that didn't match closely enough. Matching tries a close full-name comparison first, then falls back to a shared name-fragment ("Lee" or "Paulino" matching even if the two systems order or format the names differently) flagged with **~**. If a job still isn't matching, the two systems' names for it have drifted further than either check can bridge — correct the name in one system to match the other, or note the mismatch to whoever maintains the BuilderTrend job codes.

**Remember the invoiced figure only counts `Paid` and `Pending/Sent` rows** — `Draft` invoices are deliberately excluded because they haven't gone out to the client yet.

## What this page deliberately doesn't do

- **No profit margin or GPM.** This is a billing-pace tool, not a profitability one.
- **No liability recognition.** It doesn't model what's owed on a job beyond the invoiced-vs-earned gap.
- **No automatic formula before CA Signed.** See *How earned revenue is calculated*, above — that's what manual overrides are for.

These are deliberate simplifications, not gaps waiting to be filled in.
