# Teams Chat Export (with attachments)

Save your entire Microsoft Teams chat history as a single web page you can keep forever — **with the pictures _and_ the documents from your chats stored right next to it.**

Works on **Windows** and **macOS**. No technical skills needed — you mostly just sign in and wait.

---

## What you get

A folder called **`teams-archive`** containing:

- **`index.html`** — a page that opens in any browser and works completely offline. Search, filter, and read all your conversations.
- **`images/`** — every picture from your chats, shown right inside the messages.
- **`files/`** — attached documents (Word, Excel, PowerPoint, PDF…) saved locally as clickable links.
- **`unrecoverable-attachments.csv`** — an honest list of anything that couldn't be downloaded (for example a file in someone else's private OneDrive, or an old expired link).

**Everything stays on your computer. Nothing is ever uploaded anywhere.**

---

## Before you start (read this first)

You'll get the best result if you have these ready:

1. **Close Microsoft Edge completely.** The tool needs Edge to itself.
2. **Know which account** you want to export (your work account). Edge often signs in with the *wrong* account automatically, so you'll need to double-check.
3. **You do _not_ need your SharePoint address.** For documents, the tool opens a second tab at `https://powerpoint.cloud.microsoft/` automatically — this works for any organisation.

> ⏱️ The whole process takes anywhere from a few minutes to about an hour, depending on how many chats you have. You can leave it running.

---

## Download

1. Click the green **Code** button at the top of this page.
2. Choose **Download ZIP**.
3. Unzip it somewhere easy to find (e.g. **Desktop** or **Downloads**).

---

## How it works (the 6 automatic steps)

When you run the tool it does everything for you, in order:

1. Sets up a small private copy of Python (first run only).
2. Opens a browser window so **you can sign in**. ← *this is the only part that needs you*
3. Exports all your chats.
4. Downloads the images from your chats.
5. Downloads the documents (if you signed in to SharePoint).
6. Builds the `teams-archive` page and lets you save it.

The next sections walk through it for your system.

---

## ▶️ Windows

### Step 1 — Start it

Double-click **`export-windows.bat`**.

> If Windows shows a blue **SmartScreen** box: click **More info** → **Run anyway**. (It's safe — the code is right here in this repository.) No admin rights are needed.

A black command window opens and starts working. **Leave it open the whole time** — it shows progress and is where you'll press Enter later.

### Step 2 — Sign in (the important part)

After a moment, an **Edge browser window opens**. Here's exactly what to do:

1. **Tab 1 — Teams.** Check it's signed in as the **exact** account whose chats you want to export. Edge often picks the **wrong** account automatically — if so, sign out and sign back in with the right one. **Wait until you can see your chats.**
2. **Tab 2 — Office (for documents).** A second tab opens by itself at `https://powerpoint.cloud.microsoft/`. Check it too is signed in as the **same** account. If it's wrong, sign out and sign back in.
3. **Same account in both tabs!** They must match — otherwise you get the wrong person's data, or no documents at all.
4. **Finish.** Go back to the **black command window** and press **Enter**.

> 💡 **Why two tabs?** Teams pictures and SharePoint documents live in two different places. Signing in to both means the tool can fetch *everything*. If you only sign in to Teams, you'll still get all your chats and images — documents will simply be listed as "unrecoverable".

> ⚠️ **Check you're the right person in SharePoint.** SharePoint often signs you in automatically with the *wrong* account (a personal or old work account it remembered). If the SharePoint tab shows the wrong name, or you get **"You don't have access to this right now"**, click your profile picture (top-right) or the **9-dots menu (top-left)** → **SharePoint**, sign out, and sign back in with the **same work account you used for Teams**. It must be the exact same account in both tabs.

> ⚠️ **Important:** You finish the sign-in step by pressing **Enter in the command window** — *not* by closing the browser. The browser closes by itself afterwards.

> **If Teams gets stuck on the loading screen** (just the Teams logo for a long time): you're usually signed in anyway. Open the SharePoint tab as above, then press Enter in the command window. If nothing happens afterwards, run the script again.

### Step 3 — Wait

The tool now exports your chats and downloads images and documents automatically. This is the longest part. The command window shows live progress — just let it run.

### Step 4 — Save your archive

When it's done, a **Choose Folder** window appears. Pick a safe place — for example **OneDrive** or a USB drive. The whole `teams-archive` folder (the page **and** its `images/` and `files/`) is copied there.

The archive then opens in your browser automatically, and Explorer highlights `index.html` so you can see exactly where it landed.

> 💡 **Tip:** Drag `teams-archive/index.html` onto your browser's bookmarks bar for one-click access from now on. Always keep `images/` and `files/` next to `index.html` — moving the page on its own breaks the pictures and document links.

---

## 🍎 macOS

### Step 1 — Unblock the folder (one time only)

macOS blocks downloaded files until you say it's OK. Open **Terminal** (search "Terminal" with Spotlight) and paste this, then press Enter:

```
xattr -cr ~/Downloads/TeamsExportPortableWithAttachments-main
```

> That's the folder name you get after unzipping. If you unzipped somewhere else, change the path, e.g.
> `xattr -cr ~/Desktop/TeamsExportPortableWithAttachments-main`

### Step 2 — Start it

Double-click **`export-mac.command`** (or the **Teams Export** app icon).

> If macOS says it "cannot verify the developer": right-click the file → **Open** → **Open**. (You only need to do this once.)

A Terminal window opens and starts working. **Leave it open the whole time** — it shows progress and is where you'll press Enter later.

### Step 3 — Sign in (the important part)

After a moment, an **Edge browser window opens**. Here's exactly what to do:

1. **Tab 1 — Teams.** Check it's signed in as the **exact** account whose chats you want to export. Edge often picks the **wrong** account automatically — if so, sign out and sign back in with the right one. **Wait until you can see your chats.**
2. **Tab 2 — Office (for documents).** A second tab opens by itself at `https://powerpoint.cloud.microsoft/`. Check it too is signed in as the **same** account. If it's wrong, sign out and sign back in.
3. **Same account in both tabs!** They must match — otherwise you get the wrong person's data, or no documents at all.
4. **Finish.** Go back to the **Terminal window** and press **Enter**.

> 💡 **Why two tabs?** Teams pictures and SharePoint documents live in two different places. Signing in to both means the tool can fetch *everything*. If you only sign in to Teams, you'll still get all your chats and images — documents will simply be listed as "unrecoverable".

> ⚠️ **Check you're the right person in SharePoint.** SharePoint often signs you in automatically with the *wrong* account (a personal or old work account it remembered). If the SharePoint tab shows the wrong name, or you get **"You don't have access to this right now"**, click your profile picture (top-right) or the **9-dots menu (top-left)** → **SharePoint**, sign out, and sign back in with the **same work account you used for Teams**. It must be the exact same account in both tabs.

> ⚠️ **Important:** You finish the sign-in step by pressing **Enter in the Terminal window** — *not* by closing the browser. The browser closes by itself afterwards.

> **If Teams gets stuck on the loading screen** (just the Teams logo for a long time): you're usually signed in anyway. Open the SharePoint tab as above, then press Enter in the Terminal. If nothing happens afterwards, run the script again.

### Step 4 — Wait

The tool now exports your chats and downloads images and documents automatically. This is the longest part. The Terminal shows live progress — just let it run.

### Step 5 — Save your archive

When it's done, a **Choose Folder** window appears. Pick a safe place — for example **OneDrive** or a USB drive. The whole `teams-archive` folder (the page **and** its `images/` and `files/`) is copied there.

The archive then opens in your browser automatically, and Finder reveals `index.html` so you can see exactly where it landed.

> 💡 **Tip:** Drag `teams-archive/index.html` onto your browser's bookmarks bar for one-click access from now on. Always keep `images/` and `files/` next to `index.html` — moving the page on its own breaks the pictures and document links.

---

## ⚡ Advanced: use a specific SharePoint address

The tool already opens the Office sign-in tab (`https://powerpoint.cloud.microsoft/`) for you automatically, which works for most organisations. If your company needs you to sign in at your own SharePoint address instead, you can point the second tab there:

**Windows** — in Command Prompt, in the unzipped folder:
```bat
set "SHAREPOINT_URL=https://YOURCOMPANY.sharepoint.com"
export-windows.bat
```

**macOS** — in Terminal, in the unzipped folder:
```bash
SHAREPOINT_URL="https://YOURCOMPANY.sharepoint.com" ./export-mac.command
```

You'll still sign in once in that tab; after that, just press Enter to continue.

---

## Switching accounts (signing in as a different user)

The export uses whoever is signed in during the sign-in step. To export a **different** person's chats, or if the wrong account is signed in:

1. Run the script again so the browser window opens.
2. In the browser, **sign out** of the current account by going to:
   `https://login.microsoftonline.com/common/oauth2/v2.0/logout`
3. **Sign back in** as the right person:
   - **Teams:** `https://teams.microsoft.com`
   - **SharePoint/OneDrive (for documents):** `https://YOURCOMPANY.sharepoint.com`
4. Wait until both are fully loaded, then press **Enter** in the command/Terminal window.

> To wipe everything and start totally fresh, delete the `.profile` folder next to the scripts, then run the script again.

---

## The archive folder, in detail

`teams-archive/` contains `index.html`, `images/`, `files/`, and `unrecoverable-attachments.csv`. The page:

- Works **offline** — no internet needed.
- Opens in any browser (Chrome, Edge, Firefox, Safari).
- Shows chat **images inline**, and opens **documents** locally.
- Lets you **search** across all chats and **filter** by **People**, **Groups**, and **Meetings**.
- Sorts by name or most recent message, and supports **dark mode** (follows your system).

> Keep `index.html`, `images/`, and `files/` together. To move or share the archive, zip the **whole** `teams-archive` folder.

---

## Running it again later

You can run the export as often as you like. It **skips chats it already exported** and **reuses files it already downloaded**, so later runs are much faster.

> Forgot to sign in to SharePoint the first time, and some documents are listed as unrecoverable? Sign in to SharePoint, then run again — the tool retries the ones that failed.

---

## Requirements

| | Windows | macOS |
|---|---------|-------|
| **Browser** | Microsoft Edge | Microsoft Edge |
| **Python** | Installed automatically | Installed automatically |
| **Internet** | Needed during export | Needed during export |
| **Admin rights** | Not needed | Not needed |

> **Edge must be closed** before you start.

---

## Troubleshooting

| Problem | What to do |
|---------|-----------|
| **Windows SmartScreen warning** | Click **More info** → **Run anyway** |
| **macOS "cannot verify developer"** | Right-click the file → **Open** → **Open** |
| **macOS "Operation not permitted"** | Run `xattr -cr <path-to-folder>` in Terminal |
| **Nothing happens after I sign in** | Make sure you pressed **Enter in the command/Terminal window**, not just closed the browser |
| **Teams stuck on the loading logo** | You're usually signed in — sign in to SharePoint in tab 2, then press Enter. If still stuck, run the script again |
| **Documents show as broken/unrecoverable** | You weren't signed in to SharePoint. Sign in, then run the script again |
| **SharePoint says "You don't have access to this right now"** | SharePoint likely signed you in with the **wrong** account. In the SharePoint tab, click your profile picture (top-right) → sign out, and sign back in with the **same** work account you use for Teams. If it still says you don't meet the requirements, your organisation blocks non-managed browsers (Conditional Access) — only the chats/images can be exported, not SharePoint files |
| **Edge not found** | Install Edge from https://microsoft.com/edge |
| **Export seems frozen** | Large accounts just take a while — check the progress line in the window |
| **Empty archive** | Make sure you signed in to the **correct** Teams account |
| **Windows `[WinError 3]` / "cannot find the path"** | Windows limits paths to 260 characters. Unzip to a **short** location such as `C:\TE` (not a deep, double-nested Downloads folder), then run again |
| **"Edge must be closed" error** | Quit Edge completely, then try again |

---

## Privacy & Security

- All data stays on **your** machine — the archive is fully offline.
- The tool uses your own Edge sign-in. No separate logins, no API keys, no accounts.
- Nothing is sent to any external server. No tracking, no analytics.
- All source code is here in this repository for anyone to review.
