# Team Members — Help Manual

*What it's for, how to use it, and where to look when a tool isn't showing what someone expects — written for people, not programmers.*

*early-September 2026 edition*

## Part I — Using Team Members

For everyone who touches the roster or hands out access to a CORE tool. No prior knowledge assumed.

### What Team Members is

Team Members is the roster of everyone at Bellweather, and the single place that controls who can get into every other CORE tool — the Project Pipeline, the PM Updater, the PM Budget Worksheet, and the Operations cash-flow lens. There's no separate sign-up process for those tools: a person's access comes entirely from what's set for them here.

It lives at [a web address](https://bellweatherllc.github.io/tools/team-members.html), like the rest of CORE. Open the link, sign in, and you're looking at the roster.

### Signing in

Team Members uses your regular Bellweather Microsoft account — the same one you use for email. Click **Sign in with Microsoft**, pick your Bellweather account if asked, and the roster loads.

Anyone with a Bellweather account can sign in and look at the roster. Actually changing it — adding people, editing their details, or touching the Access matrix or Tool Settings — depends on permissions set up ahead of time on the underlying SharePoint list. If you sign in and can't save a change, that's a permissions question — ask Byron.

> **Where editing access is actually controlled.** Unlike the Pipeline, PM Updater, and Budget Worksheet, Team Members has no admin toggle or admin role built into the page itself — there's no on/off switch in here that decides who can edit. Whether you can save a change comes entirely from your permission level on the **CORE_TeamMembers list**, in the **BWCore** SharePoint site: **Edit/Contribute** lets you save changes on any tab; **Read** lets you sign in and see the roster, but every save (adding, editing, or removing a person; changing an Access-matrix cell; flipping a Tool Settings toggle) will fail with a permissions error. To grant or remove someone's editing access, go to the BWCore site, open the CORE_TeamMembers list's permissions, and add or remove them from the group (or individual permission) that has Edit/Contribute there — that's the one and only place this is controlled.

### The three tabs

Team Members has three tabs across the top:

- **Roster** — the list of people: name, email, department, role, and whether they're active.
- **Access** — a grid of every person against every CORE tool, controlling what each person can see and do in each one.
- **Tool Settings** — a handful of extra on/off flags that individual tools read, beyond simple access.

## Part II — The Roster tab

### Adding and editing people

The add-row at the top of the Roster tab takes a name, an email, and a department, and adds them to the bottom of the list. Click any existing row to edit it in place — name, email, department, role, and active/inactive — then save or cancel.

**Department** is one of Operations, Sales, Design, or Production. **Role** is a free-form job title (Project Manager, Logistics Coordinator, Designer, and so on) — it's what the Tool Settings tab uses to decide who a setting applies to.

### Active and inactive

Turning someone inactive doesn't delete them — it dims their row and excludes them from lists elsewhere in CORE (like the LC capacity panel), while keeping their history and access settings intact in case they come back. Use inactive for someone on leave or between roles; use delete only for a person who was added by mistake or has genuinely left and won't be back.

### Removing someone

The delete button on a row asks for confirmation, then removes them from the SharePoint list entirely — including whatever access and settings they had. There's no undo. If there's any chance the person returns, mark them inactive instead.

## Part III — The Access tab

### Reading the matrix

The Access tab is a grid: one row per person, one column per CORE tool. Above the grid, a **Tool Reference** legend spells out, for each tool, what a **Member** can do versus what an **Admin** can do — read that first if you're not sure what a given level actually unlocks.

This matrix controls access to the *other* CORE tools — Pipeline, PM Updater, PM Budget Worksheet, and OPS/Cash Flow. It has no effect on who can edit Team Members itself; that's a separate SharePoint permission, covered under **Signing in** above.

The four tools, and roughly what each level means:

- **Pipeline** — *Member* can view projects, phases, and milestones on the board. *Admin* can also edit project data, dates, phases, and statuses.
- **PM Updater** — *Member* can submit project updates. *Admin* can submit and delete them.
- **PM Budget Worksheet** — *Member* can view the budget worksheet for the projects they're assigned to as PM. *Admin* can view, edit, and set up the worksheet for any project.
- **OPS/Cash Flow** — this is the Operations lens inside the Pipeline, not a separate tool. Either level opens it — there's no separate write tier, so pick whichever is convenient.

### Setting a person's level

Click the dropdown in the cell where a person's row meets a tool's column, and choose **—** (no access), **Member**, or **Admin**. It saves as soon as you pick it — no separate save button — and takes effect the next time that person loads (or reloads) the tool in question.

> If a tool a person should have doesn't appear as a column at all, it means the underlying SharePoint field for it doesn't exist yet on the Team Members list. That's a setup step for Byron, not something fixable from this page.

## Part IV — The Tool Settings tab

This tab holds small on/off flags that don't fit the Member/Admin model — settings a specific tool checks for a specific person. Right now there's one:

### Show name for LC Capacity Tracking in Pipeline?

When turned on for a Logistics Coordinator, that person's name appears in the LC capacity panel inside the Project Pipeline and counts toward LC capacity targets there. It only applies to people whose Role is Logistics Coordinator — everyone else won't appear on this list at all.

Like the Access matrix, each toggle saves immediately.

## Part V — Everyday questions

**I gave someone Member access and they still can't see the tool.** Have them fully sign out and back in to that tool, not just refresh — access is read when the tool loads, not continuously. If that doesn't fix it, double-check you saved the right person's row and the right column.

**Someone left the company. What do I do here?** Mark them inactive on the Roster tab. Leave the Access tab alone unless you also want to strip their tool access immediately — inactive alone doesn't revoke it.

**A tool I expect to see as a column on the Access tab isn't there.** The column only appears once the matching SharePoint field exists on the CORE_TeamMembers list. Ask Byron to add it.

**I don't see a Sign Out button.** Team Members doesn't have one — it isn't needed day to day. Signing out of your Bellweather Microsoft account (or closing the browser) signs you out everywhere, including here.

**What's the difference between "inactive" and "delete"?** Inactive hides someone from active lists but keeps every setting intact, ready to switch back on. Delete removes the person and everything tied to them, permanently. When in doubt, use inactive.

**Where do I go to give someone edit access to Team Members itself?** Not on this page — there's no admin switch here. Go to the CORE_TeamMembers list in the BWCore SharePoint site and give them Edit/Contribute permission on that list. See **Where editing access is actually controlled** under Signing in.
