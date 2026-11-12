# Levels of Development — Help Manual

*What it's for, how to use it, and where to look when something isn't showing what you expect — written for people, not programmers.*

*mid-November 2026 edition*

## Part I — Using Levels of Development

For anyone assessing a project, setting one up, or reading the results. No prior knowledge assumed.

### What Levels of Development is

Levels of Development rates how complex a project is likely to be, across about a dozen dimensions — drawings, site conditions, engineering, selections, client relationship, permit risk, and so on — each on a simple 1 / 2 / 3 scale. Several people weigh in independently on the same project, and the tool rolls their answers up into a composite picture: where the project is genuinely straightforward, where it's going to need extra attention, and what that should mean for timeline, staffing, and the Design Agreement.

It's filled out early — once a project is a real Lead, or once the DA is signed — so the people doing the work have a heads-up before they're already in the middle of it.

It lives at [bellweatherllc.github.io/tools/levels-of-development.html](https://bellweatherllc.github.io/tools/levels-of-development.html), like every other CORE tool.

Anyone with a Bellweather account can sign in, see the project list, and submit an assessment. Setting up a project — writing the Scope of Work, entering Target Investment — is restricted to the **Admin** toggle in the header, covered in Part II.

### The three tabs

- **Assess** — where you pick a project and submit your own rating of it.
- **Raw Results** — every assessment submitted for a project, side by side, unaveraged. See exactly who said what.
- **Analysis** — the rolled-up picture: a composite score, which dimensions need discussion, timeline guidance, and a library of conclusions pulled from the pattern of ratings.

A project only shows up in Raw Results and Analysis once at least one assessment has been submitted for it. Brand-new projects will only appear in Assess until then.

## Part II — Setting up a project

For whoever sets a project up before the team assesses it — today, that's Byron or Joey.

### Turning on Admin mode

Flip the **Admin** toggle in the header (next to the gear icon). This unlocks two things specific to project setup — Scope of Work and Target Investment editing — plus the ability to add or edit the assessment dimensions themselves (the questions everyone rates against).

Turn it back off when you're done. Admin mode doesn't require any special permission beyond being signed in — treat it as "I'm setting this up right now," not a permanent mode to leave on.

### Writing the Scope of Work

With Admin on and a project selected in Assess, the **Scope of Work** box becomes an editable rich-text area with a small toolbar — **B** / *I* / U / bullet list — above it.

- **Paste directly from the DA.** Bold, italic, underline, and bullet lists all carry over from a Word document or PDF. Messy source formatting (fonts, colors, odd spacing) gets cleaned up automatically — only the structure survives, not the clutter.
- **Type from scratch** if there's nothing to paste yet. Use the toolbar buttons the same way you would in Word.
- Click **Save Scope** when you're done. It saves straight back to the project's record — the same place Projects Manager reads scope from — so this isn't a separate copy that can drift out of sync.

Do this *before* asking people to assess the project. Scope of Work is the main context assessors read before they rate anything; an empty box means they're guessing.

### Setting Target Investment

Same idea, smaller field. With Admin on, **Target Investment** becomes a single text box next to Address and CA Value. Type whatever's accurate:

- A specific number: `$950,000`
- A range: `$900,000 – $950,000`

There's no required format — write it the way you'd say it out loud. Click **Save** next to the field.

### What assessors see instead

With Admin off, both of these show as plain, read-only text — no edit boxes, no Save buttons. That's the normal view for everyone doing an assessment.

## Part III — Getting several people to assess the same project

This is the actual day-to-day question: *"I need five people to rate this project — how do I make that happen?"*

### The short answer

1. Go to **Assess**, select the project from the dropdown.
2. Click **🔗 Copy Link to This Project** (top right of Assessment Setup).
3. Paste that link into an email, Teams message, or wherever you're reaching the five people — one message, one link, to all of them.
4. Each person opens the link, signs in with *their own* Microsoft account, and lands directly on that project — no searching the dropdown themselves.
5. Each fills out their own ratings and notes, and clicks **Submit Assessment**.

Because each person signs in with their own account, their name and department are captured automatically — you don't tell them "select Peskin and also put your name as so-and-so." The link does the project selection; their sign-in does the rest.

### If someone doesn't use the link

They can still get there manually: sign in, go to Assess, and pick the project from the dropdown by name. The dropdown lists every project in **Lead** or **DA Signed** stage. If a project was signed that same morning and doesn't show up, it's almost certainly a caching issue — tell them to hard-refresh (Ctrl+Shift+R / Cmd+Shift+R) or add `?v=<current version number>` to the end of the URL.

### Submitting an assessment

For each dimension, click the empty circle to open a small popup with **1 / 2 / 3** — pick the one that fits, using the guide text above each section as a reference. Add a note alongside any rating where there's something worth flagging ("client is undecided on selections," "setback variance likely," etc.) — notes are what make Raw Results and Analysis actually useful instead of just a grid of numbers.

**Open Notes** at the bottom is for anything that doesn't fit a specific dimension — a general concern, a pattern from a similar past project, something that needs to be named before the project moves forward.

Click **Submit Assessment** when done.

**Clear Form** wipes everything you've entered without submitting — use it if you want to start over.

### Autosave and picking up where you left off

While you're rating dimensions, your in-progress work saves itself automatically — a few seconds after each rating click or note, and again periodically in the background either way. A small floating badge in the bottom-left corner of the screen confirms this:

- **Autosave on** — a project is selected and ready; nothing to save yet.
- **Saving draft…** (amber, pulsing) — a save is in progress.
- **Draft saved · [time]** (green) — the last save succeeded.
- **Draft not saved — will retry** (red) — the last save failed; it'll try again on the next change or the next periodic save.

Because this saves to the same place everything else does, you can close the tab, sign out, switch devices, or just come back tomorrow — sign back in, pick the same project, and your ratings and notes are right where you left them, picked up automatically. (The badge will say **Draft restored from [time]** when this happens.) Nothing here is ever saved to your own computer only — it's the same SharePoint record whoever else is reviewing the project would see, which is also why it doesn't show up in Raw Results until you actually submit it.

### Editing a submission after you've already submitted

If you already submitted an assessment for a project and need to change something, go back to **Assess** and select that project again. Under **Already assessed**, your own submission shows an **Edit your submission** link — click it to load your previous ratings and notes back into the form. The Submit button becomes **Update Assessment**; click it to save your changes in place (it doesn't create a second entry — Raw Results and Analysis simply show your latest answers). Autosave is intentionally off while editing a past submission, so a half-finished edit never temporarily overwrites your real, already-submitted answers — nothing changes until you click **Update Assessment**. If you change your mind partway through, **Cancel Edit** discards the in-progress changes and leaves your original submission untouched.

## Part IV — Reading the results

### Raw Results

Pick a project to see every assessment submitted for it, one column per dimension, each respondent's score and note shown individually — nothing is averaged or hidden here. This is the place to go when you want to see *who* thinks what, not just the summary number.

### Analysis

The same data, interpreted:

- A **composite score** and plain-language read of what it means for oversight (standard process, HODs should be informed, or active HOD oversight required).
- Every dimension shown as a small pie of who-rated-what, with each person's note alongside.
- **Needs Discussion** — dimensions where at least one person rated High, surfaced so they don't get missed before the DA is signed.
- **Constraints & Shared Expectations** — anything the team flagged as a hard boundary everyone needs to work from.
- **Design Timeline** and **Production Timeline** guidance — Fast-Track / Standard / Extended ranges, with the one matching the composite score highlighted.
- **Key Conclusions** — a short library of standing implications (duration, deadline risk, cost exposure, scope-deviation risk) tagged by which departments they're relevant to.

## Part V — Everyday questions

**I changed something and it still looks old.** This is almost always the browser serving a cached copy of the page. Hard-refresh (Ctrl+Shift+R / Cmd+Shift+R), or add `?v=<version number shown in the header>` to the end of the URL to force a fresh load.

**A project I know is signed doesn't show up in Assess.** The dropdown only lists projects in **Lead** or **DA Signed** stage. If it was just signed, give it a minute and refresh — if it still doesn't appear, the project's stage in SharePoint may not have updated yet; check Projects Manager.

**A project doesn't show up in Raw Results or Analysis.** Those two tabs only list projects that already have at least one submitted assessment. Submit one from Assess first.

**I turned on Admin mode by accident and don't see anything different.** Admin mode only changes what you see in the Assess tab — the Scope of Work and Target Investment fields become editable, and dimension-editing controls appear. It has no effect on Raw Results or Analysis.

**I tried to save something and got a red error message.** The message after the dash is the actual reason, not just "it failed" — read it, and if it's not self-explanatory, send Byron the exact text. Recent versions of this tool will auto-create a missing SharePoint field the first time it's needed, so most setup-related errors should resolve themselves on a second try.

**I pasted from Word and the formatting didn't come through.** This should work for Scope of Work as of November 2026 — bold, italic, underline, and bullet lists should all survive a paste. If something specific still doesn't carry over, tell Byron exactly what you pasted and what was lost.

**Can two people rate the same project differently?** Yes — that's the entire point. Raw Results shows every individual answer side by side precisely so disagreement is visible, not averaged away before anyone notices it.

**Where do Scope of Work and Target Investment actually live?** On the project's own record in SharePoint — the same place Projects Manager reads and writes them. Editing them here updates the one shared copy; there's no separate version to keep in sync.

**Who can set up Scope of Work and Target Investment?** Anyone who turns on Admin mode — there's no separate permission check beyond being signed in. In practice, that's been Byron and Joey. If editing isn't saving for you, the underlying SharePoint field may need a permissions check — ask Byron.

**How do I know autosave is actually working?** Watch the floating badge in the bottom-left corner while you're rating dimensions — it should move through "Autosave on" → "Saving draft…" → "Draft saved · [time]" as you go. If it ever shows "Draft not saved — will retry," something's wrong with the connection; it'll keep trying on its own, but if it stays red, send Byron a screenshot.

**The Risk Assessment (or Levels of Development) guide text looks out of date compared to what's in the manual.** A tool update can ship improved guide copy, but if that block's guide text was ever customized in Admin mode, your SharePoint copy keeps the custom wording rather than silently being overwritten. In Admin mode, each block's guide text box has a **Reset guide text to current default** button above it — click it to pull in the tool's current built-in wording (it'll ask you to confirm, since this overwrites whatever is there).
