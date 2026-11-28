# Selections Design — Help Manual

*How designers share a project's selections with the client, and the one-time setup behind it. Written for people, not programmers.*

*November 2026 edition*

## Part I — For designers

Everything a designer needs. No setup and no technical steps.

### What a client link is

A client link is a private web address for one project. When the client opens it, they see their selections the same way **Preview client view** shows them: rooms, options, pictures, full prices, meeting recaps and who to contact. They don't need a Microsoft account or a password.

The client can only look. They can't change anything, and they never see team notes, files or BuilderTrend descriptions.

### Sending a client their link

1. Open the project in the Design screen.
2. Click **Client link** at the top right.
3. Click **Make a link** the first time. After that, the link is already there.
4. Click **Copy link**, then paste it into your email or text to the client.

That's it: about ten seconds, once per project. The link keeps working and always shows the latest selections, so you never need to send a new one after making changes.

> **Check what they'll see first.** **Preview client view** shows exactly what the link shows. Use it before you send the link.

### Turning a link off, or replacing it

The link works for anyone who has it, so send it only to the client.

- **Make a new link** turns off the old address and gives you a new one. Use it if the link went to the wrong person, or the client forwarded it more widely than they should have.
- **Turn off link** (click it twice) stops the link working. The client will see "This link isn't active." You can make a new link at any time.

Each project has one link at a time.

### If the dialog says "Client links aren't set up yet"

This is a company-wide setting, not something to fix per project. Tell Byron. The dialog's grey "What's missing" line says why; include it in your message.

## Part II — For administrators: one-time setup

This setup was done once for the whole company (November 2026). Nothing here is repeated per client or per project. Keep this part for when something needs renewing or fixing.

### How it works

The Design screen and the client's browser never share a Microsoft sign-in. Instead:

1. **Make a link** stores a long random code for that project in SharePoint (`CORE_SelectionDesignDocs`, Job `_links`).
2. The client's link carries that code and the address of a small server program, the **clientView function**, which runs in the company's Azure Function App (`bellweather-core-proxy`).
3. When the client opens the link, the function checks the code, reads that one project from SharePoint with its own read-only permission, removes everything the client shouldn't see, and sends back the rest.

Turning a link off deletes its code, so the function refuses it from then on.

### The pieces

| Piece | Where | What it is |
|---|---|---|
| App registration | Microsoft Entra admin center → App registrations → **CORE Client View (server)** | The function's own identity. Client ID `db07a965-b11f-45fb-b56a-a6399f6d99df`, tenant ID `ff0e3f46-615a-4a22-8a80-4bb83d627785`. |
| Permission | That app → API permissions | Microsoft Graph **Sites.Selected** (application), admin consent granted, plus **read** access granted on the BWCore site only. It can't see any other site. |
| Client secret | That app → Certificates & secrets | The function's password. Its value is stored **only** in the Function App settings. It expires: see below. |
| Function App settings | Azure portal → Function App `bellweather-core-proxy` → Environment variables | `CLIENT_VIEW_TENANT_ID`, `CLIENT_VIEW_CLIENT_ID`, `CLIENT_VIEW_CLIENT_SECRET`. |
| CORS | Same Function App → CORS | Allows `https://bellweatherllc.github.io`. |
| Function address | SharePoint `CORE_Config`, item Title `ClientViewUrl`, in the Value column | `https://bellweather-core-proxy-g3etccfth8cad9fd.westus3-01.azurewebsites.net/api/clientView` |

> **Never paste the client secret's value** into chat, email, Teams or any file. It goes only into the Function App setting. If it ever leaks, delete it in Certificates & secrets and make a new one (see below).

### Renewing the client secret (every 24 months)

The secret was created in November 2026 and expires about **November 2028**. The exact date is under Certificates & secrets. When it expires, every client link stops working at once. Put a reminder in your calendar a month before.

1. Entra admin center → App registrations → **CORE Client View (server)** → **Certificates & secrets** → **New client secret**. Description: the month and year. Expires: 24 months. **Add**.
2. Copy the **Value** right away (it's shown only once).
3. Azure portal → Function App `bellweather-core-proxy` → **Environment variables** → `CLIENT_VIEW_CLIENT_SECRET` → paste the new value → **Apply**, then **Apply** again to confirm. The app restarts.
4. Open any client link in a private window to check it works.
5. Back in Certificates & secrets, delete the old secret.

### Setting or changing the function address

The Design screen reads the address from `CORE_Config`. The easiest way to set it is from the Design screen itself: when links aren't set up, the **Client link** dialog has a **Function address** box. Paste the address, click **Save address**. It saves to the right list automatically.

> **If you add the item by hand in SharePoint,** make sure it actually saves. SharePoint's list screen can keep a new row in your browser (the item panel says "This item is saved locally") without uploading it, and then the tools can't see it. Opening the list in a private window shows what is really saved. Using **Save address** in the dialog avoids this. (The list's web address ends in `CORE_Config1`; that is just its address, not a second list.)

### When something goes wrong

The **Client link** dialog's "What's missing" line names the setup problem:

- **No item named "ClientViewUrl" was found**: the address isn't saved. Use **Save address** in the dialog.
- **isn't an Azure function address**: the saved address is wrong. Paste the address above into the dialog's box and save.
- **CORE_Config couldn't be read**: a SharePoint or sign-in problem. Reload and try again.

If the dialog works but the client's link doesn't open:

- **"This link isn't active"**: the link was turned off or replaced. Make a new link and send it.
- **"Your selections couldn't load right now"** on every link: the function is running but can't read SharePoint. The usual cause is an expired client secret (renew it as above). Otherwise check that its three settings are present.
- **"Couldn't reach Bellweather"**: the client's internet connection, or the Function App is stopped, or the CORS setting is missing.
