# People Analyzer — Help & Repair Manual

*How to run an evaluation round, how to read the results, and what to check when something looks wrong. Written for people, not programmers.*

*late-October 2026 edition · People Analyzer v3*

## Part I — Using People Analyzer

For anyone who evaluates team members or runs an evaluation round. No prior knowledge assumed.

### What People Analyzer is

People Analyzer is where Bellweather rates each team member against the **five Core Values** and **GWC**:

- **The five Core Values:** Disciplined, Humbly Confident, Always Improving, Forward-Looking, and Cares & Connects with People.
- **GWC** means *Gets it*, *Wants it*, and has the *Capacity* to do it.

Several evaluators rate each person independently. The tool combines their answers into a **consensus** for each person and gives an overall status.

It lives at [a web address](https://bellweatherllc.github.io/tools/people-analyzer.html), like the rest of CORE. The **?** in the header, next to the ⚙️ switch, opens this manual.

### Signing in

Click **Sign in with Microsoft** and pick your Bellweather account. Your name fills in automatically as the evaluator.

Evaluations are sensitive, so they're stored in **OperationsSecure**, the restricted SharePoint site. If you can sign in but the page says *Error loading data from SharePoint*, you don't have access to that site. Ask Byron.

> **Who can do what.** The page itself has no admin role. Anyone who can open it can see the ⚙️ Admin panel and the 🗑 Manage switch. What actually protects the data is SharePoint permissions on the two OperationsSecure lists, `PA_Evaluations` and `PA_Sessions`:
>
> - **Edit/Contribute** lets someone submit, delete, and manage sessions.
> - **Read** lets them look, but every save fails.
> - **No access** means the tool won't load data at all.
>
> Grant or remove access on those lists in OperationsSecure. That's the only place it's controlled.

### Offline mode

**Use offline (no SharePoint)** on the sign-in screen lets you try the tool without saving anything:

- You can create sessions, submit evaluations, view Summaries, and download Excel.
- Everything is gone when you close or reload the tab.

Don't run a real evaluation round offline.

### The two tabs

- **Evaluate** — rate one person at a time.
- **Summaries** — see the consensus for everyone, or drill into one person.

The ⚙️ switch in the header opens the **Admin** panel. That's where sessions are created and where the SharePoint connection is checked.

## Part II — Evaluating someone

### Step by step

1. **Your Name** is filled in from your Microsoft account. Leave it as it is. (See *Evaluator names matter* below.)
2. Pick the **Session**, for example *Q1 2026 Review*. Only sessions that are switched on in Admin appear here.
3. Pick the **Team Member**. The list is grouped by department and shows **active** team members only.
4. Rate each Core Value:
   - **+** — lives it.
   - **+/−** — sometimes.
   - **−** — doesn't.

   Under each value is a short list of what it looks like in practice.
5. Rate **Gets It**, **Wants It**, and **Capacity to Do It** as **Y** or **N**.
6. Optionally, add a comment to any Core Value. Comments appear in Summaries and in the Excel export.
7. Click **Submit Evaluation**. Every rating must be filled in first.

The form clears, and you can pick the next person.

### Evaluator names matter

The tool counts evaluators **by the name in the Your Name box**. *Joey Smith* and *Joey* count as two different people. Keep the name as your account fills it in, so that the "N of M" counts in Summaries come out right.

### Changing an evaluation

You can't edit an evaluation after you submit it. Instead, delete it and submit a new one (see **Deleting evaluations** in Part III).

## Part III — Reading the Summaries

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

### One person's view

The top of the individual view is the **Trajectory**: a grid with one column per session, showing the consensus and status each round. Sessions appear in the order set in Admin.

Below that, each session has its own block. The block shows:

- every evaluator's ratings, with their comments;
- the consensus;
- the date range of the evaluations.

### Exporting

Each session block has two export buttons:

- **📋 Copy as Text** copies a plain summary to paste into an email or document. Split votes are spelled out.
- **📥 Excel** downloads the Core Values consensus laid out to match the Bellweather EOS form. Evaluator comments, and a note on any split vote, are attached as cell comments on the rating.

**📥 Download All Sessions** in the Trajectory header gives one Excel file with a sheet per session. It only appears once the person has been evaluated in more than one session.

### Deleting evaluations

Turn on **🗑 Manage**, under the filter bar. Delete buttons then appear:

- **🗑 at the end of a table row** deletes all of that person's evaluations. If a session is picked, it only deletes the ones in that session.
- **✕ remove under an evaluator's name**, in the individual view, deletes that one evaluation.

The **🗑 on a session block header** deletes that person's evaluations for that session. It's visible even when Manage is off.

Every delete asks first. Deletes are permanent, and there's no undo.

## Part IV — Admin

### Evaluation sessions

- **Create** — type a name, like *Q1 2026 Review*, and click **Create**. Names must be unique.
- **Switch on or off** — the green switch opens a session for evaluation. Switching it off only hides it from the Evaluate tab. Its evaluations stay, and it still shows in Summaries. Close a round this way once it's done.
- **Reorder** — drag by the ⠿ handle. The order controls the Trajectory columns and the session lists.
- **Delete** — 🗑 deletes the session **and every evaluation in it**, after confirming. Switch it off instead unless you really mean to remove the data.

Switching on or off and reordering save to SharePoint straight away.

### SharePoint Links

This section checks the connection between the tool and SharePoint. When everything is healthy it reads *All evaluations are linked to a team member.* Otherwise it can show three kinds of message:

- **A yellow box listing columns to add.** A column the tool needs is missing. See *Repair* below.
- **"N older evaluations are matched by name only"** with a **Link** button. Click it. It records each evaluation's team member Id, so renaming that person later won't separate them from their history. Ratings and names don't change.
- **"N evaluations don't match anyone on the roster."** These are hidden from Summaries. See *Someone's evaluations disappeared* below.

## Part V — Repair: when something looks wrong

### Someone isn't in the Team Member dropdown

Work through these in order:

1. **They're inactive.** Check Team Members. Inactive people can't be evaluated, but their old results still show in Summaries, marked *(inactive)*.
2. **Their department is blank.** They're still listed, under an **Unassigned** group. Set a department in Team Members.
3. **They were added after you opened the page.** Reload.
4. **They aren't on the roster at all.** Add them in Team Members. People Analyzer only reads the roster and never changes it.

### Someone's evaluations disappeared from Summaries

- **Check the session filter first.** It's often still set to an earlier round.
- **If they still don't appear,** open Admin → SharePoint Links. If their old name is listed under *don't match anyone on the roster*, their evaluations were saved under a name that no longer matches the roster and were never linked to their Id. To fix it:
  1. In Team Members, temporarily rename the person back to the exact name shown.
  2. Reload People Analyzer, then click **Link** in Admin.
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

The tool keeps working without them, but session on/off, session order, and person links aren't saved.

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
| `CORE_TeamMembers` | BWCore | The roster. **Read only** from this tool. Columns used: `EmployeeName`, `Department`, `Active` |

### How the pieces connect

- **Sessions.** `SessionId` on an evaluation is the Id of its `PA_Sessions` row. Deleting a session deletes its evaluations too.
- **People.** `EmployeeId` on an evaluation is the person's Id in `CORE_TeamMembers`. Evaluations without one fall back to an exact name match, ignoring capitals and extra spaces. **Link** in Admin fills it in.
- **Names stored on evaluations.** `EmployeeName` holds the name as it was at the time of the evaluation. It's kept for the record. The tool shows the person's current roster name.
- **Comments.** `Comments` holds the per-value comments as a small JSON record, keyed by value.
- **Ratings.** They're stored as `+`, `+/-`, `-`, `Y`, and `N`.

### What People Analyzer never does

- It never adds, edits, or removes anyone in `CORE_TeamMembers`. Team Members owns the roster.
- It never stores a consensus or status. Both are worked out live from the evaluations every time.
- It never deletes anything without asking first.
