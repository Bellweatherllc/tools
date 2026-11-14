# Fleet Tracker — Help Manual

*What it's for, how to use it, and what to check when something looks wrong. Written for people, not programmers.*

*November 2026 edition*

## Part I — Using the Fleet Tracker

For anyone who looks after Bellweather's vehicles. No prior knowledge assumed.

### What it is

The Fleet Tracker keeps every company vehicle, and every employee vehicle that Bellweather pays a parking placard for, in one place: registrations, insurance, inspections, oil changes, parking placards and permits, EZ Pass tags, Compass GPS units, decals and the people to call.

It works out on its own what is overdue or coming due in the next 30 days, every time you open it. Nobody has to keep a due list up to date by hand.

It lives at [a web address](https://bellweatherllc.github.io/tools/fleet-tracker.html), like the rest of CORE.

### Signing in and access

Sign in with your regular Bellweather Microsoft account. What you can do depends on your **Fleet** setting on the Team Members → Access tab:

- **Admin** — add, edit and delete vehicles, EZ Pass tags and authorized drivers, and open **Setup**.
- **Member** — see everything, change nothing.
- **Blank** — no access. You'll see a message telling you to ask an admin.

Your name and level show under the title, with a **Sign out** link.

### The tabs

- **Needs Attention** — open issues, missing spare keys, missing placard dates, and anything overdue or due within 30 days. Admins can click **Mark resolved** (or **Spare key now in office**); this clears the item and notes it in the vehicle's service history.
- **Monthly Reminders** — everything due in a month, grouped by who needs to act. The office handles placards, permits and insurance; drivers handle the rest. **Copy message** gives you a ready-to-send note for each person. Use ‹ › to look at other months.
- **Vehicles** — every vehicle, grouped as Bellweather-owned (office, field, personal) and employee-owned. Search by plate, driver, model or VIN.
- **Parking** — contractor placards and residential permits, with the cost per period.
- **EZ Passes** — tags and plates on each New Jersey and New York account. Admins can add or remove tags.
- **Insurance** — the agent, policy and each vehicle's insurance date and ID card, plus the list of **authorized drivers**.
- **Compass** — the GPS unit serial number for each vehicle, and a link to the Compass live map.
- **Decals** — where the decals go on vans, trucks and cars.
- **Contacts** — mechanics, the car broker, and the branding shop.

### Working with a vehicle

Click any vehicle, anywhere in the tool, to open it.

- Admins edit the fields and click **Save changes**. Members see the same window read-only.
- **Log an oil change today** fills in today's date, sets the next one six months out, and asks for the mileage. Changing the last oil change date also resets the next one.
- **Service History** — add an entry for any repair or service: date, what was done, where, what it cost.
- **Open Issue** — anything typed here shows on Needs Attention until someone marks it resolved.
- **Find in SharePoint** (under Registration) searches the Auto Registrations folders and links the registration file to the vehicle.
- **Delete vehicle** asks you to click twice. It removes the vehicle and its history for everyone.

### Exporting

**Export Spreadsheet** downloads an Excel workbook of every tab, for printing or sending to the insurance agent. The export is a copy; changing it doesn't change the tracker.

## Part II — Repair

When the tracker isn't doing what you expect. Most problems are access problems.

### "You don't have Fleet access yet" or "isn't on the CORE roster"

You need to be on the Team Members roster, marked active, with **Fleet** set to Admin or Member. If the message says your email isn't on the roster, check that your email in Team Members matches the account you signed in with.

### "The Fleet lists don't exist yet" / "haven't been set up"

An admin opens **Setup** and runs step 1, **Create missing lists and columns**. Step 1 only adds what's missing, so it's safe to run again.

### A vehicle or setting is missing after an import

The import never overwrites. Anything already in SharePoint is skipped, and the log says how many. To re-import one vehicle, delete it in the tracker first, then import again.

### A registration or insurance-card link doesn't open

The file was probably moved, renamed, or is somewhere your account can't open. An admin can run **Setup → Anchor file links** (links tied to the file's ID survive moves), or open the vehicle and use **Find in SharePoint** to re-link it.

### Sign-in fails with a "redirect URI" error

The page's address has to be registered on the CORE app in Azure. Ask Byron.

### I made a change and don't see it

Reload the page. If you still see the old version of the tool itself, add `?v=` and any number to the end of the address to skip the browser's cache.

### Where the data lives

Two lists on the **BWCore** SharePoint site: **CORE_FleetVehicles** (one row per vehicle) and **CORE_FleetSettings** (insurance, EZ Pass, Compass, decals, contacts and drivers). SharePoint is the source of truth. Prefer editing in the tracker; hand edits to the list can break the structured fields.
