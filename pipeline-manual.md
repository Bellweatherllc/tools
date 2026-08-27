# The Project Pipeline — Help & Repair Manual

*How to use it, how it works, and what to do when it misbehaves — written for people, not programmers.*

*late-August 2026 edition*

## Part I — Using the Pipeline

For everyone — especially if it's your first week. No prior knowledge assumed.

### What the Pipeline is

The Pipeline is Bellweather's shared project tracking board. It shows every project the company is working on — from first lead to finished construction — laid out on a timeline, week by week, grouped by Department.

The most important thing to understand is that **everyone is looking at the same board**. There is one set of information, and the Pipeline shows it to you through four different views (more on those below). When someone changes something — moves a date, adds a note, confirms a project — that change appears on everyone else's screen within a few seconds. You never need to ask "is this the latest version?" It always is.

The Pipeline lives at a web address, like any website. There's nothing to install. Open the link, sign in, and you're looking at the board.

### Signing in

The Pipeline uses your regular Bellweather Microsoft account — the same one you use for email. When you open the page you'll see a **Sign in with Microsoft** button. Click it, pick your Bellweather account if asked, and the board loads.

What you can see and do on the board depends on your role, which is set up ahead of time. If you sign in and something you expect to see is missing, that's usually a role question — ask Byron.

### Reading the board

The board is a timeline. Weeks run left to right, and the column outlined in **gold** is the current week — "now." Each person gets a section, and each project appears as a row of colored bands within it.

- **Colored bands** are phases of work — a stretch of design, permitting, or construction. Their length shows how many weeks that phase takes.
- **Small markers on a row** are milestones — single moments like an agreement signing. Hover over anything to see what it is.
- **Badges on a project** tell you its status at a glance: whether it's still prospective, has a signed Design Agreement (DA), a signed Construction Agreement (CA), is paused, cancelled, or archived.
- **Collapsed projects** appear as a single slim bar to save space. Click to expand them and edit contents.
- **Clicking on a project name** opens a card with its details — who's on it, the address, the budget, links to almost every document that's stored on SharePoint. When multiple versions of documents exist, they're provided in a list for you to determine which is most relevant.

Color is meaningful everywhere on the board — nothing is decorated. If two things are different colors, they're telling you something different.

### The four screens

Across the top are four tabs: **Sales**, **Design**, **Production**, and **Operations**. These are four views of the same board, each arranged for a different kind of work. (Keyboard shortcut: press 1, 2, 3, or 4.)

- **Sales** is where future work gets planned. It shows leads, open production slots (reserved capacity for projects that haven't been matched yet), and the matching process that pairs a promising lead with a slot. This is where the CONFIRM button lives.
- **Design** shows the design team's workload — every project in design, by designer, with their capacity over time. Important milestones are noted (selections, crit meetings, etc.)
- **Production** shows construction — every project by Project Manager or Logistics Coordinator.
- **Operations** shows cash flow.

You'll naturally live in the screen that matches your job, but you can always look at the others — it's all one board.

### The Weekly Review

The Weekly Review is the Pipeline's running list of **what changed and what needs attention**. It gathers everything into sections — Pipeline Planning, Sales Updates, Design Updates, Production Updates, Pipeline Housekeeping, Bugs & Requests, and a look-ahead of what's coming in the next few weeks. Each section shows a count of open items.

It works on a simple rule: **every change stays in the review until someone marks it reviewed.** Date moves, status changes, new notes, assignments, costs — they all appear, and they stay there week after week until a person clicks **Mark Reviewed**. Reviewing an item resolves it for everyone, not just you. If that same thing changes again later, it comes back. Nothing expires on its own, and nothing needs resetting.

This means the review is trustworthy: if it's empty, the team has genuinely seen everything.

#### Reporting bugs and making requests

The review is also the front door for problems and ideas. In the **Bugs & Requests** section, anyone can log a **Bug** (something isn't working right) or a **Request** (something you wish the Pipeline did). Type a short description, and attach a screenshot if it helps — you can even just paste one in. Each entry shows who logged it and when, and it stays on the list, visible to everyone, until it's marked **Resolved** and goes away. If you spot a problem or have an idea, log it there rather than in an email or a hallway conversation — the list is where improvements actually get picked up from.

### The CONFIRM button

On the Sales screen, when a lead has been matched to an open slot and has a proposed designer, a green **CONFIRM** button appears. Clicking it is the biggest single action in the Pipeline. In one click:

- The project becomes real — its status changes to DA Signed.
- The open slot is consumed. Its reserved schedule transfers to the project, and the empty slot disappears from the board.
- The project goes live on the Design and Production screens, with its schedule in place.

Confirming is not casually reversible, so it's done deliberately, usually after conversation among Sales, Design, and Production.

### Document pills

Projects carry small labeled chips — **EST**, **PB**, **DA**, **TRL**, **MP**, **SO**, **SS**, **CA** — each one a shortcut to a key document (the estimate, Project Basis, Design Agreement, Trello card, master plan, Scope Outline, selections sheet, Construction Agreement). Click a pill to open that document directly.

The Pipeline finds most of these documents by itself — automatically — by looking in the project's folder. A pill's appearance tells you how it got its link — found automatically, pinned by a person, or added by hand. If a pill is missing or points at the wrong file, it can be corrected right on the project; you don't need to go hunting through folders.

### Permit & zoning information (currently in development)

For projects in permitting, the hover card shows live permit and zoning status. When you open a project's card, the Pipeline checks the permit board at that moment — you may see a brief "Checking Trello…" while it looks. What you see is current. If the check can't be completed (a connection problem, usually), the card shows the most recent information it has, clearly labeled with when it was last checked — so you always know whether you're looking at fresh or older news.

### Designer Time off

Each Designer's section has a **Time Off** button. Clicking it (or clicking any existing OFF strip on the timeline) opens a small calendar where individual days can be toggled off and on — weekends are left out automatically. Days off appear as strips on the timeline, and consecutive days merge into one band, even across a weekend. The calendar stays open while you work through several days; close it with its × button or the Escape key.

### Admin mode

Some changes are everyday moves anyone can make — like reordering projects within a person's group. Bigger changes — moving a project to a different person, editing key dates, adjusting hours — require **Admin mode**, a switch available to people with the admin role. If you try something and see "Admin mode required," that's the Pipeline protecting the schedule from accidental edits, not a malfunction.

### Everyday questions

#### Do I need to save?

No. Everything saves automatically, within moments of a change — including when you close the tab.

#### Will I see my colleagues' changes?

Yes, within a few seconds, without reloading. New projects, stage changes, schedule moves, even a project someone else confirmed into a slot — they all arrive on your open board on their own.

#### Something looks odd. What's the first thing to try?

Reload the page. It's safe — nothing is lost by reloading, because nothing important is stored on your computer. A reload fetches everything fresh and cures most oddities.

#### Where do I see what's new?

The Pipeline keeps a changelog written in plain language — every release, described in normal sentences. The version number in the top corner (and on the sign-in screen) tells you which release you're on. Byron is constantly refining, so the number will change with regularity.

## Part II — Care & Repair

For whoever maintains the Pipeline — today that's Byron. Written so that someone who is *not* a programmer could keep it healthy, diagnose trouble, and get it fixed.

### How it's built, in plain terms

The entire Pipeline is **one file** — a single web page that contains everything: the layout, the logic, the changelog. That file is hosted on **GitHub Pages**, a free service that serves web pages from an address the company controls. Think of GitHub as the shelf the page sits on; it doesn't hold any project data.

The project data itself lives in **SharePoint**, part of the company's Microsoft 365. SharePoint holds several **lists** — think of a list as a shared spreadsheet that lives in the cloud, one row per project or per person. SharePoint is the single source of truth: the Pipeline page reads from it when it loads, writes back to it when anything changes, and never stores anything important anywhere else. Your computer only keeps a temporary reading copy, which is why reloading is always safe.

Sign-in works through the company's Microsoft system. There's a registration (in a Microsoft service called Azure) that tells Microsoft "this page is allowed to ask users to sign in and read these lists on their behalf." **Byron manages that registration, and only Byron** — it's the one piece nobody else should touch, because a wrong change there locks everyone out.

Three ideas summarize the whole architecture: **one page file** (on GitHub), **one source of truth** (SharePoint), **one sign-in gate** (the Microsoft registration).

### The moving parts

- **CORE_Projects** — the main list. One row per project (and per open slot). Its rows carry the name, stage, schedule, assignments, document links — everything a project is.
- **CORE_TeamMembers** — who appears on the board, and what each person is allowed to see and do. Access is managed by editing this list directly in SharePoint — roles are just columns on a person's row.
- **CORE_Config** — settings: capacity targets, thresholds, and the keys that let the Pipeline read the Trello permit board.
- **CORE_Pipeline_Snapshots** — a library of daily board backups. See *Backups and restore*.
- **Trello** — the [PRE-CON] DESIGN ABSOLUTES board. The Pipeline reads it live when you hover a permitting project; it never writes to it. It feeds information into the PERMIT AND ZONING TRACKER BOARD and the ABSOLUTES (both in development).


> **A naming rule with a reason** — Project stages in the list are called **PipelineStage**, with exactly these values: Lead, DA Signed, CA Signed, Paused, Cancelled, Archive. Don't invent new stage names in SharePoint, and don't rename the column — the page matches these words exactly.

### Where projects come from — the Projects Manager

Projects don't originate on the Pipeline. They're created in a companion tool, the **Projects Manager** — a separate page with its own address and the same Microsoft sign-in, writing to the very same project list the Pipeline reads. Most people never open it; it's the working home of whoever handles intake and keeps project records straight. It does two jobs.

The first is **intake** — getting a new project into CORE. There are three routes in, and the everyday one runs by itself. When an estimator gives a Trello card the **EST REVIEW** label, an automation reads that card, lifts the details out of its description, and writes a new project record to SharePoint on its own. Within a few minutes the project exists in CORE, with nothing typed twice. Because the card's description is doing the work, it pays to write it cleanly: the labeled lines the automation looks for — **NAME:**, **ADDRESS:**, **HOME VALUE:**, **PROJECT DESCRIPTION:** — become the matching fields, and spacing or capitalization on those labels doesn't matter. The other two routes handle the exceptions: a **manual Trello pull** inside the Projects Manager, for catching up on older cards or ones that never got the label, and plain **manual entry** (with admin turned on) for a project that has no Trello card at all.

The second job is being the **record book**. Once a project exists, the Projects Manager is where its facts are kept true over time — who's assigned as designer, PM, and logistics coordinator; the address, neighborhood, and scope; the values, fees, and key dates; and free-form notes. When a project needs a field filled in or a wrong value corrected, this is the place to do it, and the change lands in the shared list immediately.

> **Two tools, one list** — The Projects Manager and the Pipeline are two windows onto the same CORE_Projects list. The Projects Manager is the front door and the filing cabinet; the Pipeline is the schedule and the shared timeline. A change in one is a change to the underlying list, so it appears in the other — neither holds its own copy.

### Adding and editing people

Everyone who appears on the board — every designer, Project Manager, and Logistics Coordinator — is a row in the **CORE_TeamMembers** list on SharePoint. There's no separate screen for managing them, by design: people are added and changed by editing that list directly, the way you'd edit any shared spreadsheet.

To **add** someone, open CORE_TeamMembers and create a new item. Give them their name as it should read on the board and their Bellweather work email, then set their **Role** — that field is what decides where they show up. The Role has to match one of the names the tools watch for, spelled out in full: **Designer**, **Project Manager**, or **Logistics Coordinator**. Capitalization doesn't matter, but the whole phrase does — "PM" or "Logistics" on its own won't be recognized, and the person just won't appear where you expect. Save, and they turn up the next time the board is reloaded.

To **change** someone — move a coordinator into a different role, or fix a spelling — edit their row. Changing the Role moves them to the matching section; changing the name changes how they read everywhere at once.

To **remove** someone who’s left, delete their row in CORE_TeamMembers. One caution (the same one below): a project stores a person’s name as plain text, so deleting them here doesn’t scrub the name from projects that already reference it — they drop off the board, but their name lingers on those projects until you update them.

> **Names are the connection** — Projects point at their people by *name*: the designer, PM, and logistics fields on a project row hold the person's name as text, not a hidden ID. So renaming someone in CORE_TeamMembers doesn't automatically follow through to projects already carrying the old spelling. If you rename a person, plan to update the projects that still reference the old name — or, easier, get the spelling right the first time.

What each person is *allowed to see and do* — admin rights, which screens they can touch — is governed by other columns on the same row, set once when they're added and rarely revisited afterward.

### How changes get made

The Pipeline is developed in conversation with Claude (an AI assistant), in a project workspace that holds the reference documents and history. New work usually starts from the **Bugs & Requests** list in the Weekly Review — that's the intake. From there, the working method matters more than the tool:

1. **Always start from the live file.** Download the currently deployed page and upload it to the conversation. Never let a fix be built from an old copy.
2. **Describe the change or the symptom in plain words.** Screenshots help.
3. **Every delivery is a complete new file with a higher version number.** The version appears in the top corner and on the sign-in screen. If a delivered file doesn't show a new version, something went wrong — don't deploy it.
4. **The filename never changes.** It is always `project-pipeline.html`. The version lives in the badge, not the name. (Renaming the file would break sign-in and everyone's bookmarks.)
5. **Deploy by drag-and-drop.** On the GitHub website, in the tools repository, drag the new file in to replace the old one. Within a minute or two the live address serves the new version. Ask everyone to reload.
6. **The changelog is written for humans.** Every release adds plain-language entries inside the page itself. If you ever wonder why something behaves the way it does, the changelog is the first place to look — years of decisions are recorded there in normal sentences.

### If something breaks

The Pipeline protects itself in several ways, and most trouble announces itself clearly. Work by symptom:

#### Sign-in keeps bouncing to Microsoft, or the page won't load past the sign-in screen

**What's happening:** The page couldn't get a valid sign-in token. It will retry at most three times, then stop and display the real underlying error on the sign-in screen instead of looping forever.

**What to do:** Read the error message on the sign-in screen and write it down word for word — that message is the diagnosis. Check whether Microsoft 365 itself is having an outage (if email is also down, that's your answer — wait it out). Otherwise, send the exact message to whoever is fixing. If the message mentions permissions or the application, the Microsoft registration may have been changed — that's Byron's territory.

#### Changes aren't saving

**What's happening:** A write to SharePoint failed. The Pipeline shows a save-error banner when this happens — it does not fail silently.

**What to do:** Stop editing, note the banner's wording, and reload. If the reload shows your change survived, it saved after all. If not, check your internet connection and whether SharePoint is reachable (open any company SharePoint page). Persistent save failures with a good connection mean something changed in the list itself — a renamed column is the classic culprit — and that needs a fix.

#### Projects disappeared — or a banner says the Pipeline refused to remove them

**What's happening:** Projects leave the board legitimately when a slot is consumed or a row is deleted. But if an implausible number seem to vanish at once, the Pipeline assumes the problem is the connection, not the data — it refuses to remove any of them and says so.

**What to do:** Reload first. If projects are genuinely missing after a reload, open the CORE_Projects list in SharePoint and look for the rows directly — if the rows exist, the page has a reading problem; if the rows are gone, someone deleted them, and the snapshot system can bring them back (see below).

#### A banner names a row that was "renamed from a slot into a project"

**What's happening:** Someone edited an open-slot row in SharePoint and turned it into a project by typing over it. It now confuses the board — it looks like a slot but claims to be a project.

**What to do:** The fix is always the same: delete that row and create the lead fresh through the proper flow. Never convert a slot by renaming it.

#### Permit information looks stale

**What's happening:** The hover checks Trello live every time. If it shows older information, it says so, with a "last checked" time — meaning the live check failed, usually a connection issue or expired Trello access keys.

**What to do:** If it persists across reloads and devices, the Trello keys stored in CORE_Config likely need renewing.

> **The universal repair procedure** — Whatever the problem, the repair path is the same three steps: **(1)** describe the symptom in plain words, with the exact wording of any error and a screenshot; **(2)** download the live deployed file and upload it to the Claude project alongside the description; **(3)** deploy the returned file — checking that its version number went up — by drag-and-drop on GitHub. No programming knowledge is required at any step. Any competent web developer could also work from this same package if Claude weren't available; the whole application is that one readable file.

### Backups and restore

The Pipeline takes snapshots of the whole board and keeps them in a SharePoint library. A snapshot is a complete picture of a day: every project, the planning board, and time off.

To recover, open the snapshot browser in the Pipeline, pick a date, and review the comparison it shows against today. The **Restore board to this snapshot** button sets projects, planning, and time off back to that day and brings back anything removed since. Restore is designed to be safe: **it recovers and corrects, it never deletes newer work** — projects created after the snapshot are left alone.

What's *not* covered: the documents behind the pills (those live in project folders, which have their own SharePoint version history), the Trello permit board, and Buildertrend. Each of those has its own recovery path in its own system.

### Keeping this manual current

This manual describes the Pipeline and its companion Projects Manager as deployed in late-August 2026. It should be re-edited whenever a feature changes how a person works — not for every release, but for every change a first-time reader would notice. The edition line at the top is the manual's own date; update it on every edit.

The words of this manual live in a plain-text file (`pipeline-manual.md` stored on GitHub) that anyone can read and the maintainer can edit — no web-page code involved. The page that displays it (`pipeline-manual.html`) handles all the styling automatically and never needs touching. Edit the words, save, and the manual updates.
