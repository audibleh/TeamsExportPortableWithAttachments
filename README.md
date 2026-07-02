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
2. **Know your Teams sign-in** (your work account).
3. **Know your SharePoint/OneDrive address** if you want documents too — it usually looks like
   `https://YOURCOMPANY.sharepoint.com` (replace `YOURCOMPANY` with your organisation's name).

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

1. **Tab 1 — Teams.** If you're not already signed in, sign in with your work account. **Wait until you can see your chats.**
2. **Tab 2 — SharePoint (for documents).** To get Word/Excel/PowerPoint/PDF files, you must also be signed in to SharePoint in the **same** window:
   - Open a **new tab** (press **Ctrl+T**).
   - Type your SharePoint address, e.g. `https://YOURCOMPANY.sharepoint.com`, and press Enter.
   - Sign in if asked, and wait until the page loads.
3. **Finish.** Go back to the **black command window** and press **Enter**.

> 💡 **Why two tabs?** Teams pictures and SharePoint documents live in two different places. Signing in to both means the tool can fetch *everything*. If you only sign in to Teams, you'll still get all your chats and images — documents will simply be listed as "unrecoverable".

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

1. **Tab 1 — Teams.** If you're not already signed in, sign in with your work account. **Wait until you can see your chats.**
2. **Tab 2 — SharePoint (for documents).** To get Word/Excel/PowerPoint/PDF files, you must also be signed in to SharePoint in the **same** window:
   - Open a **new tab** (press **Cmd+T**).
   - Type your SharePoint address, e.g. `https://YOURCOMPANY.sharepoint.com`, and press Enter.
   - Sign in if asked, and wait until the page loads.
3. **Finish.** Go back to the **Terminal window** and press **Enter**.

> 💡 **Why two tabs?** Teams pictures and SharePoint documents live in two different places. Signing in to both means the tool can fetch *everything*. If you only sign in to Teams, you'll still get all your chats and images — documents will simply be listed as "unrecoverable".

> ⚠️ **Important:** You finish the sign-in step by pressing **Enter in the Terminal window** — *not* by closing the browser. The browser closes by itself afterwards.

> **If Teams gets stuck on the loading screen** (just the Teams logo for a long time): you're usually signed in anyway. Open the SharePoint tab as above, then press Enter in the Terminal. If nothing happens afterwards, run the script again.

### Step 4 — Wait

The tool now exports your chats and downloads images and documents automatically. This is the longest part. The Terminal shows live progress — just let it run.

### Step 5 — Save your archive

When it's done, a **Choose Folder** window appears. Pick a safe place — for example **OneDrive** or a USB drive. The whole `teams-archive` folder (the page **and** its `images/` and `files/`) is copied there.

The archive then opens in your browser automatically, and Finder reveals `index.html` so you can see exactly where it landed.

> 💡 **Tip:** Drag `teams-archive/index.html` onto your browser's bookmarks bar for one-click access from now on. Always keep `images/` and `files/` next to `index.html` — moving the page on its own breaks the pictures and document links.

---

## ⚡ Power-user shortcut: open the SharePoint tab automatically

If you'd rather not open the SharePoint tab by hand each time, tell the tool your SharePoint address up front and it opens that second tab for you:

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
