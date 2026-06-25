# Teams Chat Export (with images)

Export all your Microsoft Teams chat history into an HTML page you can keep forever — **now with the images from your chats shown inline.**  
Works on **Windows** and **macOS**. No technical skills required.

---

## What does it do?

This tool reads your Teams chat history through Microsoft Edge and saves it as a `teams-archive` folder containing an `index.html` page plus an `images/` folder with every picture from your chats. The page works offline in any browser — you can search, filter, and browse all your conversations, and images show up directly inside the messages.

**Your data stays 100% on your machine. Nothing is uploaded anywhere.**

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

### Step 3: Wait

The export runs automatically. This can take **10–60 minutes** depending on how many chats you have.  
Don't close the command window — it shows progress.

### Step 4: Save the archive

When finished, a **Choose Folder** dialog appears.  
Pick a safe location for your archive folder — for example **OneDrive** or a USB drive. The whole `teams-archive` folder (the page **and** its `images/`) is copied there.

The archive then opens automatically in your browser, and Explorer opens with `index.html` selected so you can see exactly where it ended up.

> **Tip:** After saving, drag `teams-archive/index.html` from Explorer onto your browser's bookmarks bar.  
> That way you have one-click access to your chat history from now on.  
> Keep the `images/` folder next to `index.html` — moving the page on its own will break the pictures.

---

## macOS

### Step 1: Unblock the app (one-time only)

macOS blocks downloaded files by default. Open **Terminal** (search "Terminal" in Spotlight) and run:

```
xattr -cr ~/Downloads/TeamsExportPortableWithImages-main
```

> When you download the ZIP from GitHub, the unzipped folder is named `TeamsExportPortableWithImages-main`.  
> Adjust the path if you unzipped somewhere else, e.g.  
> `xattr -cr ~/Desktop/TeamsExportPortableWithImages-main`

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

### Step 4: Wait

The export runs in the Terminal window. Don't close it.

### Step 5: Save the archive

When finished, a **Choose Folder** dialog appears.  
Pick a safe location for your archive folder — for example **OneDrive** or a USB drive. The whole `teams-archive` folder (the page **and** its `images/`) is copied there.

The archive then opens automatically in your browser, and Finder reveals `index.html` so you can see exactly where it ended up.

> **Tip:** After saving, drag `teams-archive/index.html` from Finder onto your browser's bookmarks bar.  
> That way you have one-click access to your chat history from now on.  
> Keep the `images/` folder next to `index.html` — moving the page on its own will break the pictures.

---

## The archive folder

The generated `teams-archive` folder contains:

- `index.html` — the viewer page
- `images/` — every picture from your chats, linked inline

It:

- Works offline — no internet needed
- Opens in any browser (Chrome, Edge, Firefox, Safari)
- Shows chat **images inline** inside the messages
- Search across all chats
- Filter by **People**, **Groups**, and **Meetings**
- Sort by name or most recent message
- Dark mode support (follows your system setting)

> Keep `index.html` and `images/` together. If you want a single file to move around, zip the whole `teams-archive` folder.

---

## Re-running the export

You can run the export again at any time. It **skips chats already exported**, so subsequent runs are much faster.

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
| **"Edge must be closed" error** | Quit Edge completely, then try again |

---

## Privacy & Security

- All data stays on your local machine
- Uses your existing Edge browser session — no separate login or API keys
- No data is sent to any external server
- The HTML archive is fully offline — no tracking, no analytics
- Source code is available for review in this repository
