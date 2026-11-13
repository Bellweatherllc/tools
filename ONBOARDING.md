# Onboarding: Working in This Repo With Claude Code

This guide is for teammates who've never used GitHub or written code before. It
walks through getting set up and making your first change, using Claude Code
to do the actual git/coding work for you. You talk to Claude in plain English;
it handles the technical parts.

If you get stuck anywhere in this guide, just ask Claude — paste the step
you're on and what happened, and it'll walk you through it.

---

## What you're working with

This repo holds Bellweather's internal CORE tools — the Pipeline, Budget
Planner, and similar. It's **public** (anyone on the internet can view the
code, though not your company's actual project data, which lives separately
in SharePoint). You don't need to understand how the tools work internally —
Claude does — but keep that public-repo fact in mind: never put a password,
API key, or other secret in any file here. If you're ever asked to, or about
to, stop and ask Byron first.

The full technical ground rules Claude follows in this repo live in
`CLAUDE.md` at the repo root. You don't need to read it — Claude already
has — but if something Claude does here seems off, that file is the first
place to check.

---

## Step 1 — Get access

1. **A GitHub account.** If you don't have one, go to
   [github.com](https://github.com) and sign up — it's free.
2. **Repo access.** Send your GitHub username to Byron and ask him to add you
   as a collaborator on `Bellweatherllc/tools`.
3. **Claude Code access.** Confirm with Byron that your claude.ai account has
   Claude Code enabled. You don't need to install anything — you'll use it
   from your browser at [claude.ai/code](https://claude.ai/code).

---

## Step 2 — Start a session

1. Go to [claude.ai/code](https://claude.ai/code) and start a new session.
2. When asked which repository to work in, choose **Bellweatherllc/tools**.
3. Wait a moment while Claude sets up a private workspace with the repo's
   files. This workspace is yours alone — nothing you do here affects the
   real repo until Claude explicitly pushes changes (see Step 4).

---

## Step 3 — Tell Claude what you want, in plain English

Just describe the change. For example:

> "On project-pipeline.html, the 'DA Signed' column header should say 'Design
> Agreement Signed' instead."

> "Add a note to CLAUDE.md that the Smith project folder uses a legacy name
> and should never be renamed."

You don't need to know which file to open or how to edit it — that's Claude's
job. Treat it like asking a colleague who already knows the codebase. A few
tips that help Claude help you:

- **Be specific about which tool/page you mean** if you know it (e.g. "the
  Budget Planner" or "pm-budget-planner.html") — there are several similar
  tools in this repo.
- **Say what you want to see, not just what's wrong.** "This number should be
  red when negative" is more useful than "this looks wrong."
- **Ask it to explain before you approve**, especially the first few times:
  "before you make this change, explain what you're about to do and why."

---

## Step 4 — Review, then let it ship

By default, Claude will make the change, double-check it, and then commit,
push, and merge it into the live repo automatically — that's a standing rule
for this repo (see `CLAUDE.md`), not something you need to ask for each time.
GitHub Pages (the live site) picks up the change within a few minutes after
that.

If you want to look before it ships:

- Ask **"show me exactly what you changed before merging it"** — Claude will
  show you the before/after and pause.
- Ask **"hold off on merging until I say go"** at the start of your request if
  you want to review everything first.

Either way, nothing is destroyed if you change your mind partway through —
ask Claude to undo or adjust, and it will.

---

## Best practices

- **One topic per conversation.** Keep each session focused on one change or
  one closely related group of changes. It's easier for you to review, and
  easier for Claude to keep track of what it's doing.
- **Ask before you're sure, not after.** If you're not sure whether something
  is reversible, or whether a change might affect other tools, ask first —
  Claude will tell you plainly.
- **Never add secrets.** No passwords, API keys, tokens, or connection
  strings in any file — this repo is public. Claude knows this rule already
  and will refuse, but don't ask it to work around that refusal.
- **Flag shared/critical tools explicitly.** If your change touches something
  a lot of people depend on day-to-day (the Pipeline, Budget Planner), say so
  up front — e.g. "this is a change to the live Pipeline, please be extra
  careful and check with me before shipping." Claude will slow down and
  confirm more before merging.
- **Adding a brand-new tool?** Just ask Claude to do it and mention it's new —
  it knows the repo's conventions (where new files go, how to register them
  in the tool index) and will follow them without you needing to know the
  details.
- **If something looks wrong after a change ships**, don't try to fix it
  yourself in GitHub's web editor — tell Claude what you're seeing instead.
  It can investigate and fix the actual cause rather than just patching the
  symptom.

---

## Where to go for help

- **Claude itself** — ask it directly, including "how do I push my changes?"
  or "what's the convention for X in this repo?" It has access to `CLAUDE.md`
  and the full history of how this repo is maintained.
- **Byron** (byron@bellweatherllc.com) — for repo access, SharePoint/Azure
  access questions, or anything Claude says it can't resolve on its own.
