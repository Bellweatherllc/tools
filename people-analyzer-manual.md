# People Analyzer — Quick Guide

*How to fill out a review, run one, and fix the usual problems.*

*November 2026 · People Analyzer v23*

## The basics

### What it is

People Analyzer rates each team member on Bellweather's **five Core Values** and **GWC** (Gets it, Wants it, Capacity to do it). Several people score each person; the tool pulls the scores together.

Open it at [bellweatherllc.github.io/tools/people-analyzer.html](https://bellweatherllc.github.io/tools/people-analyzer.html) and sign in with your Bellweather Microsoft account.

### Two kinds of review

**Periodic review** — mid-year (June), or any check-in

- Admins score staff; staff score themselves.
- Rates the Core Values and GWC.
- Ends with a consensus for each person in **Summaries**.

**Annual review** — January

- Each staffer has three scorers: their **HOD**, a **second seat**, and **themselves**.
- Rates the Core Values, GWC and the **V/TO**. The staffer also answers **eNPS**.
- Ends with a **summary page** for each person: final agreed ratings, compensation, bonus, goals and the meeting close.

An admin picks the kind when creating the session — tick **Annual review** for the January one.

### Who sees what

- **Staff** see only their own self-evaluation (the Evaluate tab and **My Self-Evaluations**).
- **Admins** see everything: Summaries, Compensation, History and the Admin panel.

Admin is set per person in **Team Members → Access → People Analyzer**.

## Filling out an evaluation (everyone)

1. Open People Analyzer. Your name is already filled in.
2. Pick the **Session**.
3. Staff: you're already picked — this is your self-evaluation. Admins: pick the person you're scoring.
4. Rate each item, add a comment where it helps, and click **Submit Evaluation**.

In an annual review you'll also rate three parts of the V/TO, and staff answer one eNPS question (0–10). Only the company-wide eNPS is ever shown — never yours by name.

## Running a review (admins)

### A periodic review

1. Click the **Admin** switch in the header, type a session name (e.g. *June 2027*), and click **Create**.
2. Tell people it's open. They score; you watch **Summaries** fill in.
3. When it's done, switch the session off in the Admin panel. Nothing is deleted.

### The annual review

1. In the Admin panel, type a name (e.g. *2027 Annual Review*), tick **Annual review**, and click **Create**.
2. Under **Annual Review Scorers**, pick each staffer's **HOD** and **second seat**.
3. Everyone scores. Each admin only sees the people they're assigned to.
4. Before each review meeting, open **Summaries**, pick the annual session, and click the person's **Summary** button.
5. In the meeting, fill in the summary page:
   - **Final (agreed)** ratings — **Fill final where HOD & second seat agree** does the easy ones.
   - Wins, opportunities and goals.
   - Compensation, PTO and bonus for the year.
6. Afterwards, record the close: paste the meeting transcript (from your phone or any recording app), click **Copy prompt for Claude**, paste it into Claude, paste Claude's draft back, and edit it.
7. Click **Save & mark complete**. **Print / PDF** gives you a clean copy.

## Compensation (admins)

The **Compensation** tab shows each person's pay year by year, with a chart.

- **Pick a person** to see their salary and total compensation side by side for every year, with the change from the year before.
- **Pick a department** to compare everyone in it on one chart, by total comp or salary. Under the chart:
  - **By person** — each person's figures year by year. Click a name (or **Expand all**) for their full breakdown.
  - **By year** — click a year to see everyone that year side by side: salary, total, bonuses, PTO and each benefit. Read down a column to compare, e.g. who earned what bonus in 2025.
  - People are grouped by their **Role** in Team Members (e.g. Project Manager, Logistics Coordinator, Field Manager), with a subtotal per role. Set or fix a role in Team Members.
- **+ Add a year** to enter this year or next year. It starts from last year's figures — change what changed.
- **Edit** any year. Editing an imported year keeps the import safe; **Undo edits** brings it back.

The annual summary page saves its compensation here too, so there's one set of figures per person per year.

## History (admins)

The **History** tab has two parts. **Add History** brings in past review spreadsheets: drag the Excel files onto it, check the preview (fix the person or month if they're blank), and click **Import**. **Review History** shows one person's past reviews at a time — pick them in **Show**. Each sheet becomes one review, and its compensation shows on the Compensation tab. Importing the same file twice does nothing.

## If something's wrong

**"I'm an admin but only see Evaluate."** Hover over your name in the header — it shows your role. Set **People Analyzer = Admin** for you in Team Members, then reload.

**"We couldn't find you on the roster."** Your sign-in email doesn't match your Team Members row. Add it there (separate several addresses with semicolons), then reload.

**"I already submitted for this session."** Each person gets one self-evaluation per session. An admin can delete it: Summaries → the person → turn on **🗑 Manage** → **✕ remove**.

**Old evaluations "don't match anyone on the roster."** Usually a misspelled name. In the Admin panel under SharePoint Links, pick who each one is and click **Link**.

**Someone's missing from the list.** They're marked inactive in Team Members, or were added after you opened the page — reload.

**A yellow box in the Admin panel.** A SharePoint column or list is missing; the box names it. See *Setup* below.

**Seeing an old version.** Add `?v=` and the version number (shown in red in the header) to the address.

## Setup (for whoever maintains the tool)

The data lives on the **OperationsSecure** SharePoint site. If a list or column is missing, the Admin panel and History tab say which.

| List | Columns the tool needs |
|---|---|
| `PA_Sessions` | `SessionName`, `CreatedDate`, `IsActive` (Yes/No), `SortOrder` (Number), `SessionType` (text) |
| `PA_Evaluations` | the ratings, `EmployeeId` (Number), `AnnualData` (multiple lines, plain text) |
| `PA_ReviewHistory` | `EmployeeId` (Number), `EmployeeName`, `ReviewDate` (Date), `ReviewLabel`, `ReviewType`, `Reviewer`, `SourceFile`, `ReviewData` (multiple lines, plain text) |

**Permissions.** Each list has its own permissions:

- `PA_Evaluations` — staff Contribute, admins Design, with item-level permissions set so people read and edit only their own.
- `PA_Sessions` — staff Read, admins Contribute.
- `PA_ReviewHistory` holds compensation — admins only, no staff or group members.

**Admin panel → SharePoint Access Check** confirms all of this in one place.

The roster and the Admin role come from `CORE_TeamMembers` on BWCore (Team Members tool).
