# Teams Chat Export (with attachments)

Export all your Microsoft Teams chat history into an HTML page you can keep forever — **now with the images _and_ documents from your chats saved locally.**
Works on **Windows** and **macOS**. No technical skills required.

---

## What does it do?

This tool reads your Teams chat history through Microsoft Edge and saves it as a `teams-archive` folder containing an `index.html` page plus:

- an `images/` folder with every picture from your chats (shown inline in the messages), and
- a `files/` folder with attached documents (Office files, PDFs, etc.) saved as local, clickable links.

The page works offline in any browser — you can search, filter, and browse all your conversations, images show up directly inside the messages, and documents open straight from your machine.

**Your data stays 100% on your machine. Nothing is uploaded anywhere.**

> Some attachments can't always be recovered (for example files in another person's private OneDrive, or expired links). Those are listed in `unrecoverable-attachments.csv` next to the archive so you have a complete record.

---

## Download

1. Click the green **Code** button at the top of this page
2. Choose **Download ZIP**
3. Unzip the file to a location you can find (e.g. Desktop or Downloads)

---

## Windows

### Step 1: Run the export

Double-click **`export-windows.bat`**

Windows may show a SmartScreen warning — click **"More info"** → **"Run anyway"**.

The script automatically installs Python if needed. No admin rights required.

### Step 2: Log in to Teams

A browser window opens. If you're not already logged in:
1. Log in with your work account
2. Wait until Teams loads fully
3. **Close the browser window** (important!)

> **NB:** After you log in, Teams sometimes gets stuck on the loading screen showing only the Teams logo.
> If it stays there for a while, you are usually logged in anyway — just close the browser window and let the export continue.
> If the export does not proceed, run the script again.

### Step 3: (Optional) Sign in to SharePoint / OneDrive for documents

Many Teams attachments (Word, Excel, PowerPoint, PDF) are stored in SharePoint or OneDrive. To download those too, sign in to SharePoint **in the same browser window** before closing it:

1. While the Teams login window is still open, **open a new tab** (Ctrl+T)
2. Go to your SharePoint or OneDrive, e.g. `https://<your-company>.sharepoint.com` or `https://<your-company>-my.sharepoint.com`
3. Sign in if prompted, then **close the browser window**

> **Tip:** To open that SharePoint tab automatically, set a `SHAREPOINT_URL` before running the script. In Command Prompt:
> ```bat
> set "SHAREPOINT_URL=https://yourcompany.sharepoint.com"
> export-windows.bat
> ```

### Step 4: Wait

The export runs automatically. This can take **10–60 minutes** depending on how many chats you have, then images and documents are downloaded.
Don't close the command window — it shows progress.

### Step 5: Save the archive

When finished, a **Choose Folder** dialog appears.
Pick a safe location for your archive folder — for example **OneDrive** or a USB drive. The whole `teams-archive` folder (the page **and** its `images/` and `files/`) is copied there.

The archive then opens automatically in your browser, and Explorer opens with `index.html` selected so you can see exactly where it ended up.

> **Tip:** After saving, drag `teams-archive/index.html` from Explorer onto your browser's bookmarks bar.
> That way you have one-click access to your chat history from now on.
> Keep the `images/` and `files/` folders next to `index.html` — moving the page on its own will break the pictures and document links.

---

## macOS

### Step 1: Unblock the app (one-time only)

macOS blocks downloaded files by default. Open **Terminal** (search "Terminal" in Spotlight) and run:

```
xattr -cr ~/Downloads/TeamsExportPortableWithAttachments-main
```

> When you download the ZIP from GitHub, the unzipped folder is named `TeamsExportPortableWithAttachments-main`.
> Adjust the path if you unzipped somewhere else, e.g.
> `xattr -cr ~/Desktop/TeamsExportPortableWithAttachments-main`

### Step 2: Run the export

Double-click **`export-mac.command`** (or the **Teams Export** app icon).

If macOS shows a security warning:
Right-click the file → **Open** → click **Open** in the dialog.

The script automatically downloads Python if needed.

### Step 3: Log in to Teams

Same as Windows — log in, wait for Teams to load, close the browser.

> **NB:** After you log in, Teams sometimes gets stuck on the loading screen showing only the Teams logo.
> If it stays there for a while, you are usually logged in anyway — just close the browser window and let the export continue.
> If the export does not proceed, run the script again.

### Step 4: (Optional) Sign in to SharePoint / OneDrive for documents

To download documents stored in SharePoint or OneDrive, sign in **in the same browser window** before closing it:

1. While the Teams login window is still open, **open a new tab** (Cmd+T)
2. Go to your SharePoint or OneDrive, e.g. `https://<your-company>.sharepoint.com` or `https://<your-company>-my.sharepoint.com`
3. Sign in if prompted, then **close the browser window**

> **Tip:** To open that SharePoint tab automatically, run the script from Terminal with `SHAREPOINT_URL` set:
> ```bash
> SHAREPOINT_URL="https://yourcompany.sharepoint.com" ./export-mac.command
> ```

### Step 5: Wait

The export runs in the Terminal window. Don't close it.

### Step 6: Save the archive

When finished, a **Choose Folder** dialog appears.
Pick a safe location for your archive folder — for example **OneDrive** or a USB drive. The whole `teams-archive` folder (the page **and** its `images/` and `files/`) is copied there.

The archive then opens automatically in your browser, and Finder reveals `index.html` so you can see exactly where it ended up.

> **Tip:** After saving, drag `teams-archive/index.html` from Finder onto your browser's bookmarks bar.
> That way you have one-click access to your chat history from now on.
> Keep the `images/` and `files/` folders next to `index.html` — moving the page on its own will break the pictures and document links.

---

## Switching accounts (logging out and back in)

The export uses whichever Teams/SharePoint account is signed in during the login step. If you want to export a **different** user's chats — or if the wrong account is signed in — sign out first, then sign back in as the right user:

1. Run the export script again so the browser window opens.
2. In that window, sign out of the current account by visiting:
   `https://login.microsoftonline.com/common/oauth2/v2.0/logout`
3. Then sign back in as the desired user at:
   - **Teams:** `https://teams.microsoft.com`
   - **SharePoint / OneDrive (for documents):** `https://<your-company>.sharepoint.com` or `https://<your-company>-my.sharepoint.com`
4. Wait until each page is fully loaded and signed in, then **close the browser window** to continue.

> The login session is stored in the local `.profile` folder next to the scripts. To start completely fresh (clear all stored sign-ins), delete that `.profile` folder and run the script again.

---

## The archive folder

The generated `teams-archive` folder contains:

- `index.html` — the viewer page
- `images/` — every picture from your chats, linked inline
- `files/` — attached documents (Office files, PDFs, etc.) as local links
- `unrecoverable-attachments.csv` — a list of any attachments that couldn't be downloaded (e.g. another user's private files or expired links)

It:

- Works offline — no internet needed
- Opens in any browser (Chrome, Edge, Firefox, Safari)
- Shows chat **images inline** inside the messages
- Opens attached **documents** locally
- Search across all chats
- Filter by **People**, **Groups**, and **Meetings**
- Sort by name or most recent message
- Dark mode support (follows your system setting)

> Keep `index.html`, `images/`, and `files/` together. If you want a single file to move around, zip the whole `teams-archive` folder.

---

## Re-running the export

You can run the export again at any time. It **skips chats already exported** and **reuses attachments already downloaded**, so subsequent runs are much faster.

> If some SharePoint/OneDrive documents failed the first time (before you signed in to SharePoint), run the export again **after** signing in to SharePoint, adding `--retry-failed` to the mirror step, to retry just those.

---

## Requirements

| | Windows | macOS |
|---|---------|-------|
| **Browser** | Microsoft Edge | Microsoft Edge |
| **Python** | Auto-installed if missing | Auto-installed if missing |
| **Internet** | Required during export | Required during export |
| **Admin rights** | Not required | Not required |

> **Important:** Microsoft Edge must be **closed** before starting the export.

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| **Windows SmartScreen warning** | Click "More info" → "Run anyway" |
| **macOS "cannot verify developer"** | Right-click → Open → Open |
| **macOS "Operation not permitted"** | Run `xattr -cr <path-to-folder>` in Terminal |
| **Edge not found** | Install Edge from https://microsoft.com/edge |
| **Export seems stuck** | It's still working — large accounts take a while |
| **Empty archive** | Make sure you're logged into the correct Teams account in Edge |
| **Documents show as broken links** | Sign in to SharePoint/OneDrive during login (see the optional step), then re-run with `--retry-failed` |
| **Wrong account exported** | Sign out via the logout URL above, sign back in as the right user, re-run |
| **"Edge must be closed" error** | Quit Edge completely, then try again |

---

## Privacy & Security

- All data stays on your local machine
- Uses your existing Edge browser session — no separate login or API keys
- No data is sent to any external server
- The HTML archive is fully offline — no tracking, no analytics
- Source code is available for review in this repository
