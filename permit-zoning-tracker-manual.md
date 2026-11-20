# Permit & Zoning Tracker — Help Manual

*What it's for, and ideas for where it could go next. Written for people, not programmers.*

*November 2026 edition*

The Permit & Zoning Tracker shows every project's zoning and permitting progress as a live timeline, synced from Trello — Design Phase I/II and Pre-Con, with up to four tracks per project (Zoning A/B, Permit A/B) — so it's clear at a glance who's on track, at risk, or off track.

It lives at [a web address](https://bellweatherllc.github.io/tools/permit-zoning-tracker.html), like the rest of CORE.

This manual is a stub — it starts with just the one topic below and will grow as the tool does.

## Future Development Possibilities

Ideas that came up and were set aside, kept here so they don't have to be re-discovered from scratch later.

### Cross-checking against Philadelphia's own permit records

Philadelphia publishes its Licenses & Inspections building and zoning permit data as free, public data — no account, no login, no key needed:

- Dataset page: [opendataphilly.org — Licenses and Inspections Building and Zoning Permits](https://opendataphilly.org/datasets/licenses-and-inspections-building-and-zoning-permits/)
- The actual API (Carto's SQL API — callable straight from a web page, nothing in between): [cityofphiladelphia.github.io/carto-api-explorer/#permits](https://cityofphiladelphia.github.io/carto-api-explorer/#permits)
- Example: `https://phl.carto.com/api/v2/sql?q=SELECT * FROM permits`

The live table includes `permitnumber`, `status`, `permitissuedate`, `permitcompleteddate`, `denialdate` / `denialdocumentlink`, and `address` — fields that line up with what this tool already tracks per project (the `Permit Number` field, and each project's address).

**Considered and set aside (November 2026).** Matching by permit number only pays off late — the city doesn't assign one until intake, and by the time it's recorded in Trello the outcome is usually already known from L&I's own email notifications. Matching by **address** instead could catch a filed application earlier, before anyone's typed the permit number in at all, but addresses are fuzzier to match (a reused address returns years of unrelated permits; formats differ) and would need date-range filtering to stay precise.

The real value, if this is ever built: **ground-truth dates** for stage timing (the city's own `permitissuedate`/`denialdate`, more accurate than "whenever someone remembered to move the Trello card"), a **backstop** for a missed or buried email notification, and a **one-glance check across every active project** instead of searching an inbox project by project. If the team's email monitoring stays reliable, none of that is worth the added complexity of address matching — worth reconsidering only if that stops being true.
