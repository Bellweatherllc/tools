# People Analyzer — Help & Repair Manual

*How to run an evaluation round, how to read the results, and what to check when something looks wrong. Written for people, not programmers.*

*late-October 2026 edition · People Analyzer v12*

## Part I — Using People Analyzer

For anyone who evaluates team members or runs an evaluation round. No prior knowledge assumed.

### What People Analyzer is

People Analyzer is where Bellweather rates each team member against the **five Core Values** and **GWC**:

- **The five Core Values:** Disciplined, Humbly Confident, Always Improving, Forward-Looking, and Cares & Connects.
- **GWC** means *Gets it*, *Wants it*, and has the *Capacity* to do it.

Several evaluators rate each person independently. The tool combines their answers into a **consensus** for each person and gives an overall status.

It lives at [a web address](https://bellweatherllc.github.io/tools/people-analyzer.html), like the rest of CORE. The **?** in the header, next to the ⚙️ switch, opens this manual.

### Signing in

Click **Sign in with Microsoft** and pick your Bellweather account. The tool finds you on the CORE roster by your email address, and your name fills in automatically as the evaluator. Your first name shows in a small chip in the header; hover over it to see which account you're signed in with and your People Analyzer role.

Evaluations are sensitive, so they're stored in **OperationsSecure**, the restricted SharePoint site. If you can sign in but the page says *Error loading data from SharePoint*, you don't have access to that site. Ask Byron.

### Admins and everyone else

What you see depends on your **People Analyzer** role, set on the **Access** tab of Team Members, where the People Analyzer column offers **Member** or **Admin**, as it does for every CORE tool.

| | Admin | Everyone else |
|---|---|---|
| **Evaluate** | Anyone on the roster; evaluator name can be edited | Themselves only — a self-evaluation; name locked |
| **Summaries** | Yes | No |
| **My Self-Evaluations** | — | Yes, their own |
| **⚙ Admin panel and 🗑 Manage** | Yes | Hidden |

> **The screen is one layer; SharePoint is the other.** Hiding Summaries and the Admin panel stops people *seeing* them in the tool. It doesn't stop a determined person reading the lists directly, because anyone who can submit an evaluation has some access to `PA_Evaluations`. The real protection is the list's **item-level permissions** in OperationsSecure — see *Locking down PA_Evaluations* in Part VI.

### Offline mode

**Use offline (no SharePoint)** on the sign-in screen lets you try the tool without saving anything:

- You can create sessions, submit evaluations, view Summaries, and download Excel.
- Everything is gone when you close or reload the tab.

Don't run a real evaluation round offline.

### The tabs

The tabs sit at the right of the header:

- **Evaluate** — rate one person at a time.
- **Summaries** (admins) — see the consensus for everyone, or drill into one person.
- **My Self-Evaluations** (everyone else) — what you've submitted about yourself.

The **Admin** switch and ⚙ gear in the header (admins only) open the Admin panel. That's where sessions are created and where the SharePoint connection is checked.

## Part II — Evaluating someone

### Self-evaluations

If you're not an admin, Evaluate is set up for your self-evaluation: your name is filled in as both the Evaluator and the Team Member, with nothing to choose. If only one session is open it's already picked; rate yourself and submit.

- You can submit **one self-evaluation per session**. If you need to change it, ask an admin to delete it, then submit again.
- Afterwards, **My Self-Evaluations** shows what you submitted for each session, with your comments. It doesn't show how anyone else rated you.

Admins see self-evaluations in Summaries, kept **separate**: they never count toward the consensus, the status, or the "N of M" count. In a person's view, the self-evaluation sits in a shaded column to the right of **Consensus**, with a **self** tag, so the two are easy to compare. The **Self** column in the summary tables shows ✓ when someone has submitted one.

### Step by step

1. **Evaluator** is filled in from your sign-in — your name as it appears on the CORE roster. It can't be changed.
2. Pick the **Session**, for example *Q1 2026 Review*. Only sessions that are switched on in the Admin panel appear here; if only one is open, it's picked for you.
3. Pick the **Team Member**. The list is grouped by department and shows **active** team members only. (Everyone else sees only themselves.)
4. Rate each Core Value:
   - **+** — lives it.
   - **+/−** — sometimes.
   - **−** — doesn't.

   Under each value is a short list of what it looks like in practice. The wording comes from the Core Values sheet revised **Aug 3, 2026**; the revision date shows next to the Core Values heading.
5. Rate **Gets It**, **Wants It**, and **Capacity to Do It** as **Y** or **N**.
6. Optionally, add a comment to any Core Value. Comments appear in Summaries and in the Excel export.
7. Click **Submit Evaluation**. Every rating must be filled in first.

The form clears, and you can pick the next person.

### Evaluator names matter

The tool counts evaluators **by name**. Signed in, the name always comes from the CORE roster, so it's consistent. Older evaluations were typed in by hand, so *Joey Smith* and *Joey* there count as two different people. (In offline mode you type your name; nothing is saved.)

### Changing an evaluation

You can't edit an evaluation after you submit it. Instead, delete it and submit a new one (see **Deleting evaluations** in Part III).

## Part III — Reading the Summaries

Summaries is for admins only.

### Filters

The bar along the top of Summaries filters what you see:

- **All**, or a **department** button, shows a table per department. The department buttons come from the roster, so a new department appears on its own.
- **Individual…** opens one person's full view.
- **All Sessions / a session name** limits everything to one evaluation round.

Click any row in a table to open that person's individual view.

### The consensus

For each value, the consensus is **the answer most evaluators gave**.

When there's a **tie**, the consensus is the **lower** rating, and it gets a **dashed outline**. Hover over it to see the split, for example *1 × +, 1 × −*. A tie is a sign the evaluators see the person differently, so it's worth discussing.

### The status

| Status | When it appears |
|---|---|
| **Right Person / Right Seat** | Every Core Value consensus is **+** and every GWC is **Y** |
| **Watch** | At least one Core Value is **+/−**, and nothing is **−** or **N** |
| **Below the Bar** | Any Core Value is **−**, or any GWC is **N** (including from a tie) |

The status is worked out fresh every time the page loads. It's never stored, so it always matches the current evaluations.

### The Evals column

- **With a session picked**, it reads **N of M**. *M* is everyone who has evaluated *anyone* in that session; *N* is how many of them have rated this person. "2 of 4" means two evaluators haven't rated this person yet.
- **With All Sessions picked**, it's just the total number of evaluations across every round.

Self-evaluations aren't counted in either number. The **Self** column next to it shows ✓ when the person has submitted one (with All Sessions picked, how many sessions).

### One person's view

The top of the individual view is the **Trajectory**: a grid with one column per session, showing the consensus and status each round. Sessions appear in the order set in the Admin panel.

Below that, each session has its own block. The block shows:

- every evaluator's ratings, with their comments;
- the consensus;
- the date range of the evaluations.

### Exporting

Each session block has two export buttons:

- **📋 Copy as Text** copies a plain summary to paste into an email or document. Split votes are spelled out, and the self-rating is listed under each value, marked as not counted.
- **📥 Excel** downloads the Core Values consensus laid out to match the Bellweather EOS form. Evaluator comments, a note on any split vote, and the self-rating are attached as cell comments on the rating. The rating itself never includes the self-evaluation.

**📥 Download All Sessions** in the Trajectory header gives one Excel file with a sheet per session. It only appears once the person has been evaluated in more than one session.

### Deleting evaluations

Turn on **🗑 Manage**, under the filter bar. Delete buttons then appear:

- **🗑 at the end of a table row** deletes all of that person's evaluations. If a session is picked, it only deletes the ones in that session.
- **✕ remove under an evaluator's name**, in the individual view, deletes that one evaluation.

The **🗑 on a session block header** deletes that person's evaluations for that session. It's visible even when Manage is off.

Every delete asks first. Deletes are permanent, and there's no undo.

## Part IV — The Admin panel

The Admin panel is for admins only. The **Admin** switch in the header opens and closes it.

### Evaluation sessions

- **Create** — type a name, like *Q1 2026 Review*, and click **Create**. Names must be unique.
- **Switch on or off** — the green switch opens a session for evaluation. Switching it off only hides it from the Evaluate tab. Its evaluations stay, and it still shows in Summaries. Close a round this way once it's done.
- **Reorder** — drag by the ⠿ handle. The order controls the Trajectory columns and the session lists.
- **Delete** — 🗑 deletes the session **and every evaluation in it**, after confirming. Switch it off instead unless you really mean to remove the data.

Switching on or off and reordering save to SharePoint straight away.

### SharePoint Access Check

This section reads the lock-down settings straight from SharePoint each time the panel opens, so you can confirm them without going into SharePoint. Every line should show a green ✓.

- **PA_Evaluations** — it has its own permissions; people can read and edit only their own evaluations; and **you** can submit, see everyone's evaluations, and delete.
- **PA_Sessions** — it has its own permissions, and you can create and change sessions.
- **Who has access** — if you're an Owner of OperationsSecure, it also lists everyone with access to each list, and checks that staff ("Everyone except external users") can submit and can see sessions. If you're not an Owner, that list is hidden; ask one non-admin to sign in instead and check they see only Evaluate and My Self-Evaluations.

A red ✗ comes with the setting to change. If an admin can't see everyone's evaluations, the tool also warns them when they sign in, because their Summaries would otherwise look complete when it isn't.

### SharePoint Links

This section checks the connection between the tool and SharePoint. When everything is healthy it reads *All evaluations are linked to a team member.* Otherwise it can show three kinds of message:

- **A yellow box listing columns to add.** A column the tool needs is missing. See *Repair* below. If it lists `PeopleAnalyzerRole`, **everyone who signs in is treated as an admin** until that column exists — don't send out self-evaluations until it's gone.
- **"N older evaluations are matched by name only"** with a **Link** button. Click it. It records each evaluation's team member Id, so renaming that person later won't separate them from their history. Ratings and names don't change.
- **"N evaluations don't match anyone on the roster."** These are hidden from Summaries. See *Someone's evaluations disappeared* below.

## Part V — Repair: when something looks wrong

### I should be an admin but only see Evaluate and My Self-Evaluations

1. Hover over your name chip in the header. It shows the address you signed in with, which roster entry the tool matched you to (and whether by email or by name), and the People Analyzer role it read. The same line appears in the note on the Evaluate tab.
   - *not found on the CORE roster* — see the next item.
   - *People Analyzer role: blank* — set it to Admin (step 2).
   - *couldn't load the CORE roster* — a SharePoint problem reading BWCore. Reload; if it persists, the message names the error.
2. In Team Members → Access, set your **People Analyzer** column to **Admin**.
3. Reload People Analyzer.

If the role is already Admin, the tool probably matched the wrong roster entry, or none — see the next item.

### "We couldn't find you on the CORE roster"

The tool matches the email address you sign in with against the **Email** on your Team Members row. Some people sign in with one address and receive mail at another. Add every address, separated by semicolons, to that person's Email in Team Members, then reload. If no email matches, the tool tries an exact match on the full name.

### "You already submitted a self-evaluation for this session"

Everyone except admins gets one self-evaluation per session. An admin can delete the old one (Summaries → the person → 🗑 Manage → ✕ remove); then the person can submit again.

### Someone isn't in the Team Member dropdown

Work through these in order:

1. **They're inactive.** Check Team Members. Inactive people can't be evaluated, but their old results still show in Summaries, marked *(inactive)*.
2. **Their department is blank.** They're still listed, under an **Unassigned** group. Set a department in Team Members.
3. **They were added after you opened the page.** Reload.
4. **They aren't on the roster at all.** Add them in Team Members. People Analyzer only reads the roster and never changes it.

### Someone's evaluations disappeared from Summaries

- **Check the session filter first.** It's often still set to an earlier round.
- **If they still don't appear,** open the Admin panel → SharePoint Links. If their old name is listed under *don't match anyone on the roster*, their evaluations were saved under a name that no longer matches the roster and were never linked to their Id. To fix it:
  1. In Team Members, temporarily rename the person back to the exact name shown.
  2. Reload People Analyzer, then click **Link** in the Admin panel.
  3. Rename them back in Team Members.

  Linked evaluations follow the person from then on.

### A session switch or drag "won't stick"

- **A message says *Not saved — PA_Sessions needs a … column*:** a column is missing. See the next item.
- **A message says *Error saving … to SharePoint*:** you probably only have Read access on `PA_Sessions`.

### The yellow "add these columns" box appears

The tool looks for three columns on OperationsSecure each time it loads:

| List | Column | Type |
|---|---|---|
| `PA_Sessions` | `IsActive` | Yes/No, default Yes |
| `PA_Sessions` | `SortOrder` | Number |
| `PA_Evaluations` | `EmployeeId` | Number |
| `CORE_TeamMembers` (BWCore) | `PeopleAnalyzerRole` | Same type as `PipelineRole` |

The tool keeps working without them, but session on/off, session order, and person links aren't saved — and without `PeopleAnalyzerRole`, everyone is an admin.

The name must be typed **exactly** as shown, with no spaces, when the column is *first* created. SharePoint fixes a column's internal name at the moment of creation. A column first made as *Is Active* and renamed later is still invisible to the tool. In that case, delete it and re-create it with the exact name.

### "N of M" looks wrong

Almost always, the same evaluator was entered under two spellings. Open the person's individual view and compare the evaluator names. Delete the stray evaluation and re-submit it under the right name.

### One evaluator rated the same person twice in a session

Both evaluations count toward the consensus. Turn on 🗑 Manage and remove the one you don't want.

### Nothing happens when downloading Excel

The Excel feature loads from a public library site (jsdelivr). If a network filter blocks that site, the download fails silently. Try another network, or use **Copy as Text**.

### "Error loading data from SharePoint"

1. **Check access.** You need access to the OperationsSecure site and to the `PA_Evaluations` and `PA_Sessions` lists in it.
2. **Check the list names.** The lists must still be named `PA_Evaluations` and `PA_Sessions`. Renaming a list breaks the tool.
3. **Check the roster list.** It must be `CORE_TeamMembers` on BWCore, and you need at least Read access there.

### Sign-in keeps coming back to the sign-in screen

Close the tab and open the tool again. Sign-in is remembered only for the current browser tab. If it still loops, sign out of Microsoft in that browser and sign in again.

### I'm seeing an old version

The version badge in the header shows which build you have. After an update, add `?v=` and the version number to the address to skip the browser's cached copy, for example `people-analyzer.html?v=3`.

## Part VI — Reference for repairs

### Where the data lives

| List | Site | Holds |
|---|---|---|
| `PA_Sessions` | OperationsSecure | One row per evaluation round. Columns: `SessionName`, `CreatedDate`, `IsActive`, `SortOrder` |
| `PA_Evaluations` | OperationsSecure | One row per evaluator per person per session. Columns: `SessionId`, `EmployeeId`, `EmployeeName`, `EvaluatorName`, `EvaluationDate`, the five Core Values, `GetsIt`, `WantsIt`, `Capacity`, `Comments` |
| `CORE_TeamMembers` | BWCore | The roster. **Read only** from this tool. Columns used: `EmployeeName`, `Department`, `Active`, `email`, `PeopleAnalyzerRole` |

### How the pieces connect

- **Sessions.** `SessionId` on an evaluation is the Id of its `PA_Sessions` row. Deleting a session deletes its evaluations too.
- **People.** `EmployeeId` on an evaluation is the person's Id in `CORE_TeamMembers`. Evaluations without one fall back to an exact name match, ignoring capitals and extra spaces. **Link** in the Admin panel fills it in.
- **Names stored on evaluations.** `EmployeeName` holds the name as it was at the time of the evaluation. It's kept for the record. The tool shows the person's current roster name.
- **Comments.** `Comments` holds the per-value comments as a small JSON record, keyed by value.
- **Ratings.** They're stored as `+`, `+/-`, `-`, `Y`, and `N`.

### Locking down PA_Evaluations

The screen-level roles keep everyone except admins out of Summaries and the Admin panel, but SharePoint has to do the real work. Only an **Owner** of OperationsSecure can make these changes; group Members can't.

1. **PA_Evaluations** → ⚙ → List settings → **Permissions for this list** → **Stop Inheriting Permissions**. Then grant:
   - **Everyone except external users** — **Contribute**
   - each admin — **Design** (it includes *Override List Behaviors*, which lets them see everyone's evaluations)
2. **PA_Evaluations** → List settings → Advanced settings → **Item-level Permissions**:
   - Read access: **Read items that were created by the user**
   - Create and Edit access: **Create items and edit items that were created by the user**
3. **PA_Sessions** → List settings → Permissions for this list → **Stop Inheriting Permissions**. Then grant **Everyone except external users** — **Read**, and each admin — **Contribute**.

Owners of OperationsSecure see everything regardless. Group *Members* keep Edit, which with item-level permissions on means they see only their own evaluations, unless they're also granted Design as an admin.

Keep the people set to Admin for People Analyzer in Team Members and the people granted Design in step 1 the same. Then check it all with **Admin panel → SharePoint Access Check**.

With that in place, a non-admin's browser only ever receives their own evaluations, whatever they try.

### What People Analyzer never does

- It never adds, edits, or removes anyone in `CORE_TeamMembers`. Team Members owns the roster.
- It never stores a consensus or status. Both are worked out live from the evaluations every time.
- It never deletes anything without asking first.
