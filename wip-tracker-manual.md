# WIP Tracker — Help & Repair Manual

*What it does, where its numbers come from, and what to do when something looks wrong — written for people, not programmers.*

*September 2026 edition*

## What this page is

WIP Tracker answers one question, for every job that's collected money: **are we billed ahead of the work, or behind it?**

It replaces the old monthly spreadsheet review. There's no profit-margin math here — just a formula (or a manual figure, where the formula doesn't apply yet), calculated fresh every time you open the page.

It lives at [a web address](https://bellweatherllc.github.io/tools/wip-tracker.html), like every CORE tool. Sign in with your Bellweather Microsoft account and it loads.

> **Scope:** only **DA Signed** and **CA Signed** jobs show up here. Leads (no agreement signed yet) are excluded outright — nothing's been collected on a job that isn't under contract in some form.

Everything from **Viewing as of** through the **Projects** heading — the date picker, the lock/PDF chips, the Portfolio totals, and **Data sources and calculations** — stays pinned at the top of the window as you scroll the Projects table below it, the way a spreadsheet's frozen header row would. The Portfolio totals themselves are laid out the same way: one row of labels, one row of figures, one row of context, like a small spreadsheet rather than a row of separate cards.

### How earned revenue is calculated

The formula depends on which of the three phases a project is in — see *Reading the Stage column*, just below, for exactly how those phases are defined. The same walkthrough is also built into the page itself: click **Data sources and calculations** near the top of the page to expand it without leaving the page.

- **Design** (DA signed, before CA signing) — earned is **50% of whatever's actually been invoiced so far**. Design-phase billing runs roughly 25% of the *projected* construction amount (the estimate set at DA signing) over the course of Design, paid in installments; only half of each of those payments is booked as earned, the other half is collected ahead of the work. This is new as of this revision — Design-phase jobs used to have no formula at all and needed a manual % just to show anything.
- **Pre-Con** (CA signed, before construction starts) — earned is a flat **20% of the current contract value**. A larger deposit — **30%** — is actually collected at CA signing (sized so cumulative billing reaches 30% of the *exact* construction amount, which commonly comes in higher than the Design-phase estimate), but only 20 of those percentage points are recognized as earned; the remaining 10 are collected ahead of the work, the same idea as the Design-phase split above. Progress stays at 0% for the whole Pre-Con stretch; it doesn't move until construction's first week.
- **Construction** (the Pipeline's construction phase, start to last week) — earned is the **20% already earned at signing, plus 80% of contract value × schedule progress**, where schedule progress is elapsed weeks ÷ construction duration.
- **A manual override, when set, always wins** — see *Manual % overrides*, below. It replaces whichever of the above would otherwise apply, and works the same in every phase: earned = contract value × the percentage set by hand.
- **Missing schedule data** (no DA/CA signing week yet, or a Construction phase that hasn't been built into the Gantt) leaves a project blank — better an honest gap than a guessed number nobody signed off on. See *Reading the flags*, below.

Compare earned to what's actually been invoiced, and the difference is the number that matters:

- **Invoiced more than earned** (red, "over") — the client's been billed ahead of the work. This is *expected* during Design, by design — half of every design payment is deliberately booked ahead of the work, so Design-phase jobs will almost always show "over."
- **Invoiced less than earned** (green, "under") — the work is ahead of the billing.

### Reading the Stage column

The Stage column shows one of three lifecycle phases — plain text, not a badge, but still the label — and it's separate from a project's DA Signed/CA Signed contract status underneath:

| Phase | Runs from | Runs through |
|---|---|---|
| **Design** | The week the DA is signed | The week before the CA is signed |
| **Pre-Con** | The week the CA is signed | The last week of the Pipeline's Pre-Con phase |
| **Construction** | The first week of the Pipeline's Construction phase | Its last week |

This moves with **Viewing as of** the same way earned and invoiced do — pick an earlier date and the Stage column shows whichever phase the project was in as of that date, not necessarily where it sits today.

A blank (**—**) means the Pipeline doesn't have enough schedule data yet — no DA/CA signing week, or no Pre-Con/Construction phase — to place the project in one of the three.

## Viewing a different date

**Viewing as of**, near the top of the page, defaults to today. Pick an earlier date and the whole page recalculates as if that date *were* today:

- Build progress (and therefore earned revenue) is measured through that date instead of today.
- The invoiced-to-date total only counts invoices dated on or before that exact date, not the end of that month.

This is a genuinely different calculation, not just a narrower live view — pick August 31st and you see the WIP picture exactly as it would read if August 31st were the day you opened the page.

> **This is not a historical replay.** It recalculates using **today's** project data (contract values, schedules) and the **latest** BuilderTrend export, just with the clock turned back for the purposes of the math. If a project's schedule or contract value has changed since that date, the recalculation uses the current version, not what it looked like back then. For a true frozen record of a specific month, see *Locking the month*, below — that's the one figure this page actually preserves unchanged.

**Manual overrides work from any date, including a historical one** — the ✎ pencils stay available no matter what **Viewing as of** is set to. Overrides aren't scoped to a date either way: whatever's currently set applies to the calculation regardless of which date you're looking at, which is worth keeping in mind — Fike's override, say, reflects today's understanding of the windows delay, not necessarily what was known as of the earlier date you're viewing.

**Locking follows the date you're viewing, not the real calendar.** See *Locking the month*, below, for how that actually works — the short version is that the Ryan/Joey chips stay visible everywhere, but only turn active for today or the last day of a month.

Click **Today** to snap back to the live view.

## Manual overrides

Three different **✎** buttons let a human override the math — one for % Complete (and therefore Earned), one for Invoiced, one for the Gap ("over/under") column itself. They're separate because they fix separate problems: % Complete is for when the *formula* doesn't apply or doesn't match reality; Invoiced is for when the *BuilderTrend match* is wrong; Gap is for adjusting the bottom-line over/under figure directly, without implying anything about Invoiced or Earned individually. Overriding one never touches the others.

**None of the three are scoped to "Viewing as of."** They can be set, changed, or cleared no matter which date is currently being viewed, and once set they apply to every date the same way — a manual figure isn't "as of" anything, it's just the number until someone changes it.

### % Complete

Click the **✎** next to any project's % Complete to set it by hand. This is for exactly the cases the formula can't handle:

- **A project with no phase data yet** — no DA/CA signing week in the Gantt, so there's no way to place it in Design, Pre-Con, or Construction, let alone calculate a formula for it.
- **A Design-phase job with no invoicing yet** — the 50%-of-invoiced formula has nothing to work from until something posts.
- **A Construction-phase job where the schedule says something the work doesn't.** Fike is the standing example: the Pipeline's construction phase reads 100% complete, but a large chunk of the work (windows) is still on backorder. Rather than drag the project's end date out to today — which would throw off its $/week average by inflating the apparent project length — set the override to what's actually true, say 90%, with a note explaining why.

An override replaces the formula entirely: earned becomes contract value × the percentage you set, full stop. It's marked **manual** in the table so it's never confused with a calculated figure, and its note (visible on hover) is the record of why a human overrode the math. Overrides persist until changed or cleared — they carry forward month to month rather than resetting, so update Fike's number as the real picture changes rather than re-entering it from scratch.

Overrides are stored in **`CORE_Config`** (key `wip_overrides`) — the same shared list the Pipeline already uses for things like its OPS cash-flow scenarios. No SharePoint schema change was needed to add this.

### Invoiced amount

Click the **✎** next to any project's Invoiced figure to replace it with a number you enter by hand. This is the fix for a bad BuilderTrend match — most often the exact problem the **invoiced-exceeds-contract** flag catches (see *Reading the flags*, below): a short or common job name absorbs another project's invoices and inflates the total. Rather than let a wrong number sit in the portfolio total, correct it at the project where it's actually wrong.

A manual invoiced amount is marked **manual** in the table, replaces the BuilderTrend-matched total everywhere that figure is used (Gap, and — during Design — the 50%-of-invoiced earned calculation), and flows straight into the **Net Position** figure in the Portfolio grid, which is the sum of every row's own Gap. There's no separate override for Net Position itself; fixing the project-level number is what fixes the portfolio total.

Like % Complete overrides, this persists until changed or cleared and is stored in **`CORE_Config`** (key `wip_invoiced_overrides`).

### Over/under (Gap)

Click the **✎** next to any project's Gap figure to set the over/under amount directly, with an **over**/**under** choice next to it. This is for when you want the bottom-line figure to say something specific — reconciled against a statement, agreed with the client, whatever the reason — without changing what Invoiced or Earned show above it. Both of those keep displaying their own real (or separately overridden) values; only the Gap cell itself reflects the manual figure.

A manual Gap is marked **manual** in the table, with its note visible on hover. Like the other two overrides, it persists until changed or cleared and applies regardless of which date is being viewed — stored in **`CORE_Config`** (key `wip_gap_overrides`). It also flows into the **Net Position** figure at the top, since that's just the sum of every row's own Gap — zero out one project's over/under and the portfolio total moves with it.

## Sorting the table

Click any column header — **Project, Stage, Contract Value, % Complete, Earned, Invoiced, Gap** — to sort the table by that column. Click it again to flip between ascending and descending; an arrow (▲/▼) on the header shows which one is active.

Before you click anything, the table opens in its original order: worst gap first, regardless of over or under. That's still the most useful default for a quick scan, so it's not a "sort," it's just how the page starts — there's no arrow on any header until you pick one.

Sorting is view-only. It doesn't change what's calculated, what's saved, or what a lock records — it just changes the order rows are listed in on your screen.

## Locking the month — Ryan's button, Joey's button

There's no single generic "lock" button. Instead there are two named chips near the top of the page — **Ryan** and **Joey** — each with its own status and its own button. A chip's button only works for that person: it checks who's actually signed in, not a name anyone could type, so Ryan can't lock Joey's slot and vice versa.

**The chips are always visible, but only active for today or the last day of a month.** Locking targets whatever month is currently being viewed (see *Viewing a different date*, above) — not a fixed "real now." That means:

- **Viewing today** — both buttons work as usual, locking the current, in-progress month.
- **Viewing the last day of a past month** (e.g. set *Viewing as of* to August 31st) — both buttons work too, letting Ryan and Joey retroactively approve a month that was never locked at the time, or re-approve one with corrected figures.
- **Viewing any other day** (a mid-month date) — both buttons are greyed out, with a note explaining why and a suggestion to jump to that month's last day instead. This is deliberate: a mid-month figure was never meant to be "the" number for a month, so it's shown as unavailable rather than hidden — the feature hasn't gone anywhere, it's just not the right moment to use it.

Whichever date it targets, the flow is the same:

1. **Ryan** clicks his own **Lock as Ryan** button. His chip shows a ✓, a timestamp, and the date it was locked as of. Joey's still shows "not locked." Nothing is final yet.
2. **Joey**, signed in as himself, clicks **Lock as Joey**. That fills the second slot, **finalizes the record**, and immediately opens the browser's print dialog so it can be **saved as a PDF**.

Order doesn't matter — whoever locks first, the record only finalizes once *both* slots are filled. Ryan and Joey should agree beforehand on which date they're locking (typically the month's last day) — the tool doesn't force them to have picked the identical date before each clicks, so coordinate the same way you would for any other joint sign-off.

Clicking your own button again before the other person has locked just **updates your slot** with the current figures — it's still only one signature. The live page keeps recalculating after that — invoiced totals and schedules don't freeze — but the finalized lock is untouched by that. Once finalized, a banner appears with **View locked figures**, which switches the page to show exactly what was recorded, and **Back to live**, which returns to the current numbers.

**Locking again after it's already finalized** starts a brand-new two-slot cycle — both chips reset to "not locked," and both Ryan and Joey need to lock again before a new PDF comes out. That's expected for revising a month after something changes, not an error.

A month that's never been locked always starts with both slots empty — nothing carries over from month to month.

### Saving a PDF to SharePoint

A third chip — **PDF** — sits next to Ryan's and Joey's, always active. There's no need to wait on a lock: a dated snapshot is useful whether or not the month has been formally approved yet, so the button always works.

- **Month is finalized** — the chip reads **"finalized"** and clicking **Save PDF to SharePoint** saves the locked snapshot — the same figures Ryan and Joey approved.
- **Month isn't finalized** — the chip reads **"live snapshot"** and clicking it saves the current live figures as of whatever date **Viewing as of** is set to. The PDF itself is labeled **"Live snapshot"** in its header (not "Finalized"), so anyone who opens it later can tell at a glance it wasn't an approved record — just a saved moment in time.

Either way it saves to:

> `Operations → FINANCIAL → 1. WIP Reports & Job Costs → WIP Reports`

Named to match the files already sitting in that folder: **`BWC WIP Report MM-DD-YY.pdf`**, dated to whichever date the PDF is actually for — the locked-as-of date for a finalized month, or the **Viewing as of** date for a live snapshot. Saving again for the same date overwrites that same filename rather than creating a duplicate — so if you want a record of a specific day, that's the date to set **Viewing as of** to before saving.

This replaces printing to PDF by hand — there's no print dialog involved, and nothing is generated until you click the button. A toast confirms the save, or explains what went wrong (usually the same causes as the invoicing-export folder errors below: the folder's been renamed or moved, or the folder path under FINANCIAL doesn't match what the page expects).

Locks are stored in `CORE_Config`, one row per month (key `wip_lock_YYYY-MM`, holding both reviewers' snapshots by name and the date each was locked as of), so they don't compete for space with anything else and there's no limit on how many months of history accumulate.

> Reviewer matching is by first name on the signed-in Microsoft account (Ryan, Joey) — if either of their accounts doesn't display that first name for some reason, their button would never enable. Worth confirming once, then it's a non-issue going forward.

## Where the rest of this data lives

Everything else is read live and recalculated on the spot — nothing else is written back to SharePoint, and nothing is cached beyond your browser tab.

### Contract value & construction schedule

Both come from the **`CORE_Projects`** list on the BWCore SharePoint site — the same list the Pipeline and Projects Manager read and write. Specifically:

- **Contract value** is the `EstimatedProjectValue` field.
- **Construction schedule** is the `Construction` phase inside `GanttData` — its start week and duration, exactly as drawn on the Pipeline's Gantt. Only used for CA-Signed jobs without an override.
- **Matching to BuilderTrend** uses the `JobCode` and `ProjectName` fields — never `ClientName`. See *A BuilderTrend job total isn't showing up on a project*, below.

### Invoiced to date

The same status shown here — which file was read, when, and how to refresh it — is also live on the page itself, inside **Data sources and calculations** near the top (see *What this page is*, above, for where that panel sits). This comes from a BuilderTrend export, dropped by hand into:

> `Operations → FINANCIAL → 1. WIP Reports & Job Costs → BT_InvoicingReports_forWIPTool`

The page always reads whichever file in that folder was **most recently saved**. It sums every invoice line that isn't still a `Draft` (a Draft hasn't actually gone out to the client, whatever date it carries) **and is dated on or before the date being viewed** — today by default, or whichever date is set in *Viewing as of* (above). There's no dedicated "invoice date" column in this export, so `Deadline` stands in for it — checked against `Date Paid` across a real export, invoices here come due the same day they're paid more often than not, so `Deadline` tracks the real invoice date closely. (If a row has no `Deadline`, `Date Paid` is used instead; if neither exists, it's counted regardless of date rather than silently dropped.) Note this is the exact date, not the end of its month — an invoice dated for later this month doesn't count yet just because it's still technically "this month."

### Keeping the export current

1. In BuilderTrend: **Financial → Invoice → All Jobs → Export**.
2. Save the file as `Invoices_MM_DD_YYYY.xls` (today's date).
3. Drop it in the folder above.
4. Reload WIP Tracker, or click **Refresh** in the top bar.

## Reading the table

| Column | What it means |
|---|---|
| Stage | The project's current phase — Design, Pre-Con, or Construction. See *Reading the Stage column*, below. |
| Contract Value | `EstimatedProjectValue` from the Pipeline. |
| % Complete | Construction schedule progress — blank during Design (no physical schedule yet), 0% through Pre-Con, ramping during Construction — or a manual override. |
| Earned | The phase formula's result (see *How earned revenue is calculated*, above), or a manual override. Blank if the phase can't be determined or its formula is missing an input. |
| Invoiced | Sum of non-Draft invoices matched to this job, dated on or before the date being viewed. |
| Gap | Invoiced − Earned, or a manual override. Red = billed ahead; green = work ahead of billing. |

Click any column header to sort by it — see *Sorting the table*, above.

### Reading the flags

- **"Not set"** — the project has no DA/CA signing week in the Gantt yet, so it can't be placed in Design, Pre-Con, or Construction at all. Fix in the Pipeline, or set a manual override.
- **"No invoicing yet"** (Design-phase jobs) — Design-phase earned comes from actual invoicing, and nothing's posted for this job yet. Earned stays blank until something does, or until a manual override is set.
- **"Pre-con"** — not a problem flag, just a label confirming the 0%/deposit-only stage: construction hasn't started, so earned reflects the deposit only until it does.
- **"No construction schedule"** (Construction-phase jobs) — no `Construction` phase in the Gantt, so earned is showing the 20% signing share only, without the schedule-driven remainder. Fix: add the construction phase in the Pipeline, or set a manual override.
- **"~" (weak match)** next to an invoiced figure — this job matched a BuilderTrend job code by a shared name fragment against Job Code or Project Name, rather than a close match on either. Worth a second look, but often just means the project's Job Code is blank or doesn't follow BuilderTrend's naming convention — it's shown, not hidden, so it stays checkable rather than silently guessed.
- **"No BT match"** — nothing in the latest export matched this project's Job Code or Project Name at all. Check the *BuilderTrend jobs not matched* list further down the page.
- **"Invoiced exceeds contract value"** — the invoiced total is higher than the project's contract value. This can mean a bad BuilderTrend match pulled in another job's invoices (see *A project's invoiced total looks way too high*, below), but just as often it means the contract value itself is stale — change orders are common after CA signing, and if `EstimatedProjectValue` was never updated to include them, real invoicing can legitimately outpace it. Check both before assuming the match is wrong; set a manual invoiced amount with the ✎ next to the figure only if the match itself is actually broken.
- **"No contract value"** — `EstimatedProjectValue` is empty in the Pipeline.

## When something looks wrong

**The whole page says "Failed to load."**
Usually a sign-in or permissions issue. Reload and sign in again. If it keeps happening, check whether other CORE tools are also having trouble.

**The source panel says it couldn't read the invoicing export.**
The page looks for `1. WIP Reports & Job Costs`, then `BT_InvoicingReports_forWIPTool`, inside FINANCIAL on the Operations SharePoint site. If either's been renamed or moved, check the error message's folder path against what's actually there.

**A BuilderTrend job total isn't showing up on a project.**
First check the *BuilderTrend jobs not matched* list at the bottom of the page — the job may be sitting there under a name that didn't match closely enough. Matching compares against a project's **Job Code** and **Project Name** in the Pipeline — never Client Name, which often carries both clients' full names ("Zoe Odenwalder & Evan Skalski") and reliably defeats a close match even when the job is obvious to a person. It tries a close comparison first, then falls back to a shared name-fragment ("Odenwalder" matching even if the rest of the text around it differs) flagged with **~**. If a job still isn't matching, check that the project's **Job Code** is actually filled in and matches BuilderTrend's own naming convention (surname plus the client's first initial, e.g. `ODENWALDERZ`) — that field exists specifically to keep this reliable, and a blank or mistyped one is the most common cause of a job landing in *unmatched* or matching only weakly.

**The lock chips are greyed out.**
That's by design too, but for a narrower reason than the override pencil: locking only works for today or the last day of a month. Check *Viewing as of* — if it's set to a mid-month date, jump to that month's last day (or click **Today**) to make the buttons active again.

**A project's invoiced total looks way too high.**
Check whether it's absorbing a job that isn't really its. Matching only ever attaches a BuilderTrend job to the single closest project — it should never invent a match out of nothing, so if a job's real project isn't showing up (often because that project isn't `DA Signed`/`CA Signed` right now, or its name has drifted), that job's total belongs in the *unmatched* list, not silently parked on whichever open project happened to look closest. If a project's number seems inflated, it's worth cross-checking against the raw BuilderTrend export directly for that job's actual code. If invoiced has actually gone past the contract value, the page will already be flagging it with **"Invoiced exceeds contract value"** — once you've confirmed the real number from the export, correct it with the ✎ next to the Invoiced figure rather than leaving the bad match in place.

**Remember the invoiced figure excludes `Draft` rows** (not actually sent yet) **and anything dated after the date being viewed** — a milestone invoice pre-staged for later this month doesn't count yet just because BuilderTrend already shows it as sent.

## What this page deliberately doesn't do

- **No profit margin or GPM.** This is a billing-pace tool, not a profitability one.
- **No liability recognition.** It doesn't model what's owed on a job beyond the invoiced-vs-earned gap.
- **No formula for a project with no phase data.** See *How earned revenue is calculated*, above — that's what manual overrides are for.

These are deliberate simplifications, not gaps waiting to be filled in.
