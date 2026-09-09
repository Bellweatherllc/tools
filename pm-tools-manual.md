# PM Workbook &amp; Budget Planner — Help Manual

*How the two tools work, how they talk to each other, and what to do when something misbehaves — written for people, not programmers.*

*early-September 2026 edition*

## Part I — The PM Updater

For Project Managers keeping their projects current, and anyone reviewing that record.

### What it is

Project Updates (usually just called "the Updater") is where a Project Manager makes simple, tracked edits to their projects — end dates and costs. It's deliberately narrow: it isn't a place to redesign a schedule or rewrite a scope, just to log the small changes that happen as a project runs, in a way Ryan can see without asking.

It lives at [a web address](https://bellweatherllc.github.io/tools/project-updates.html). Sign in with your Bellweather Microsoft account, and you land on your own projects.

### Signing in and finding your projects

Across the top, a row of tabs — one per Project Manager — lets you jump to any PM's projects. You'll normally live on your own tab, but nothing stops you from looking at someone else's if you need to.

Below the tabs, **or jump directly to a project** is a second way in: a picker listing every project regardless of PM (plain text, no color-coding — that lives in the Budget Planner's picker instead, see Part II). Pick one and it becomes the only card on screen, replacing whatever the PM tabs were showing — useful when you know the project you want but not who's PM on it. Click any PM tab, or the **✕ Show all** button that appears next to the picker, to drop back to the normal tab view.

The header itself reads **PM Workbook** — the tool is still what this manual calls the PM Updater throughout; that's just the label in its header bar.

### Making an update

Open a project card and you'll find its editable fields: end date and cost figures. Type the change and save it. Two things are deliberate:

- **Every save is tracked and shared with Ryan automatically.** There's no separate step to "submit" or "report" a change — saving it is reporting it.
- **A saved entry can't be re-edited.** If you got something wrong, don't hunt for an edit button — there isn't one. Just log a new entry with the correction. The log is a history, not a form you fill in once; showing what changed and when is the point.

### Admin / delete mode

A gear-and-switch control in the header — only visible to admins in the first place — toggles delete mode. Hover it and it says exactly what it does: it only controls whether you can remove entries from a project's Change History outright. It has nothing to do with submitting updates — filling in an end date or cost is already open to anyone with ProjectUpdatesRole access, Member or Admin, whether this switch is on or off.

### The Budget Planner button

Next to a project's name is a small **Budget Planner** button. It opens that exact project directly in the PM Budget Planner — see Part III for how that handoff works. Beside it, a dot and a label always show that project's budget status:

- **red dot — Not set up yet** — no budget config, or one exists but nothing's been filled in.
- **muted gold dot — Budget partially set up** — some of the three key financial fields are filled in, some aren't.
- **green dot — Budget set up** — all three are filled in.

See **Setting up a budget**, below, for exactly which three fields decide this.

## Part II — The PM Budget Planner

For Project Managers tracking a budget against its estimate, and for Ryan and other admins setting projects up.

### What it is

The Budget Planner takes a project's original estimate (the COMP Estimating workbook, built during Sales) and turns it into a living budget: the same divisions and cost codes, but now with editable PM figures tracked against the original numbers, division by division, line by line.

It lives at [a web address](https://bellweatherllc.github.io/tools/pm-budget-workbook.html). Sign in the same way as every other CORE tool.

### Admin mode

A gear-and-switch control in the header, matching the Updater's, toggles admin mode. It controls two things:

- **The financials block** (sale price, GPM figures, document links) — visible only in admin mode.
- **The Setup button** — only admins can set up or re-import a project's budget.

Admin mode does **not** control which projects you can pick — see the picker, next.

### The project picker

The dropdown at the top lists projects to choose from, each with a small flat-colored dot in front of its name — a plain CSS circle, not an emoji, matching the status dots used elsewhere in CORE. A legend right beside the picker spells out what the colors mean, so you never have to remember it:

- **red — Not set up** — none of the three key financial fields are filled in yet (see **Setting up a budget**, below, for which three).
- **gold — Partial** — some are filled in, some aren't.
- **green — Set up** — all three are filled in.

The dot is about those three fields specifically, not about whether a budget config row exists at all. A project can have a config row — meaning picking it opens the real budget grid, not the Setup card — and still show red, if nobody's filled in Selections, Unresolved Allowance, or DESIGN COGS yet.

If the list looks incomplete, scroll within the panel — it holds every project that matches your access (see below), just capped in height so it doesn't run off the bottom of the screen.

Who sees which projects:

- **Budget admins** (BudgetRole = Admin on Team Members, or Pipeline admin) see every active project, whether admin mode is toggled on or off.
- **Everyone else** sees only the projects where they're the listed PM.

Toggling admin mode while a not-yet-set-up project is selected immediately swaps between the plain heads-up message and the Setup card — you don't need to reselect the project for that to catch up.

### Setting up a budget

An admin picks a project with no budget yet, and either sees the estimate link pulled automatically from the Pipeline or pastes one in. Loading it shows the tabs in that workbook — pick the one that matches the project, and **Create budget** copies its divisions, cost codes, and every estimating comment into the Budget Planner. From then on the project's budget lives independently: PM edits to budget lines never touch the estimate file, and re-loading the estimate later (see **Change estimate / tab**, below) never overwrites a PM's budget entries.

Creating the budget alone doesn't turn its dot green — that only happens once someone fills in the three fields the financials strip asks for: **Selections**, **Unresolved Allowance**, and **DESIGN COGS**. Fill in none and it's red, some and it's gold, all three and it's green.

**Back to Estimate Spreadsheet and Tab** (labeled **Change estimate / tab** in earlier versions), available to admins from an open budget, re-runs setup against a (possibly different) file or tab — useful if the wrong tab was picked originally. It moves the estimate figures and comments to match the new tab and cleans up anything left over from the old one; PM budget lines and budgeting comments are never touched by this.

**Reset Budget**, next to it, is the more drastic option: it permanently deletes the project's budget config, every budget line (estimate-seeded and PM-added alike), and every comment (estimating and budgeting), putting the project straight back to a red dot — no leftover data. Use it when a setup needs to start over from nothing, not just correct which tab was used. There's no undo, and it asks for confirmation before doing anything.

### Working the grid

Once a budget exists, it opens as a spreadsheet-style grid: gray division headers, a scope line per row, an Estimate column (read-only, from the snapshot) beside an editable Budget column. A few things worth knowing:

- **FIXED checkbox** — a per-line marker a PM ticks once a figure is confirmed rather than still being worked. Ticked lines turn from red to black. It's just a visual "I'm done with this one," not a lock.
- **Comments** — the speech-bubble icon on any line or division opens a comment thread. **Estimating** comments (amber) came from the original estimate file and are read-only here. **Budgeting** comments (blue) are written in the Planner itself, and a PM can edit or delete their own (an admin can delete anyone's).
- **+Other** — the last line in most divisions; it adds a new row rather than a comment, which is why the comment icon is a speech bubble and not a "+" (a "+" right next to "Add Other" was too easy to misread).

### GPM (Estimate) and GPM (Budget)

At the top of the financials, two stacked figures show gross profit margin two ways — once against the original estimate, once against the live budget — so the gap between the two is visible at a glance. Both use the same formula, Ryan's:

> (Selections + Unresolved Allowance + DESIGN COGS + Indirect + Payroll) ÷ Sale Price, using either the Estimate total or the Budget total as the denominator's cost base.

Sale Price itself always comes from the Pipeline — the Budget Planner only displays it, never edits it. To change a sale price, change it in the Pipeline.

## Part III — How the two tools connect

Neither tool duplicates the other's data — the Updater never touches a budget figure, and the Planner never touches an end date. What connects them is a shared list on SharePoint and one deep link.

### The link

The Updater's **Budget Planner** button opens `pm-budget-workbook.html?project=<id>` — `<id>` being that project's SharePoint item ID from CORE_Projects. On load, the Planner checks for that `?project=` parameter and, if present, opens straight to that project — bypassing the normal picker (and its PM-only scoping) entirely, so the link works even for a project outside your own picker filter. This is the same mechanism an admin relies on to jump straight to any project.

### The shared status logic

Both tools independently compute the same thing — red / partial / green, from the same three CORE_Budget_Config fields (Selections, Unresolved Allowance, DESIGN COGS) — and show it two different ways:

- The Updater shows a dot and a label (**Not set up yet** / **Budget partially set up** / **Budget set up**) next to the Budget Planner button.
- The Planner shows the same-colored dot next to the project's name in its own picker.

They're reading the same underlying fields, just at different moments (the Updater checks once per project list load; the Planner keeps its own copy current as fields get filled in and as budgets get created or reset), so the two should never disagree by more than a page refresh.

### Getting Ryan's attention

Nothing pages Ryan directly when a project needs a budget. Instead, the Project Pipeline's Weekly Review has a **Pipeline Housekeeping** section that flags any active project with no budget set up — that's what the Planner's "hasn't been set up" message means when it says Ryan's been alerted.

### One sign-in for everything

All of CORE — the Pipeline, the Updater, the Budget Planner, Team Members — shares the same Azure AD sign-in. Signing in once (or out once) affects every tool, which is why neither the Updater nor the Planner has its own separate Sign Out button; there used to be one on the Planner, left over from troubleshooting, and it was removed because it silently signed people out of everything else too.

### Who can see what, in one place

Both tools read the same two things off a person's row in **Team Members** (CORE_TeamMembers):

- **BudgetRole** (Member / Admin) — governs the Budget Planner specifically.
- **ProjectUpdatesRole** (Member / Admin) — governs the Updater specifically.

A person can be an admin in one and a plain member (or nothing at all) in the other — they're independent settings. See the Team Members manual for how those get assigned.

## Part IV — Everyday questions

**I'm a budget admin, but off admin mode the picker is empty.** It shouldn't be — budget admins see every project in the picker regardless of admin mode; only *real* PMs are scoped to their own assigned projects. If you're seeing an empty picker as an admin, check that BudgetRole (or PipelineRole) is actually set to Admin for you on Team Members.

**A project I expect to see in the picker isn't there.** The picker only lists projects whose Pipeline stage is CA Signed, DA Signed, or Paused. A lead-stage or archived project won't appear in either tool.

**The Updater's status doesn't match what I just did in the Planner.** Reload the Updater — its status check runs once when the page loads, not continuously, so it won't notice a field filled in or a budget reset in another tab until you refresh.

**I edited a Budget line and it disappeared after Back to Estimate Spreadsheet and Tab.** It shouldn't have — re-import is built to leave PM budget lines and budgeting comments alone, only refreshing the read-only estimate side. If a PM entry is genuinely gone, that's worth reporting rather than re-entering, since it points at a bug in the sync.

**Why don't Estimating comments let me reply or edit?** They're a straight read from the original estimate file, refreshed whenever setup re-runs against that file. Anything you want to say about a line goes in a Budgeting comment instead — that one's yours.

**Where do I actually change a PM's access to either tool?** Not in either tool — that's set on the Access tab in Team Members. See the Team Members manual for exactly where.
