"""Generate a standalone HTML archive of all exported Teams chats.

Reads the exports/ directory and produces a single self-contained HTML file
that works like Teams: sidebar with chat names, click to view, global search,
per-chat search, date filtering, etc. No server needed.
"""
from __future__ import annotations

import base64
import html
import json
import shutil
from pathlib import Path, PurePosixPath
from typing import Any
from urllib.parse import quote

# Image extensions that browsers can display inline.
_WEB_IMAGE_EXTS = {"png", "jpg", "jpeg", "webp", "bmp", "svg", "gif"}


def _resolve_bundle_relative_path(bundle_root: Path, value: str | None) -> Path | None:
    """Safely resolve a bundle-relative path under bundle_root (no traversal).

    Kept self-contained so this module can also run as a standalone script.
    """
    if value is None:
        return None
    text = str(value).strip().replace("\\", "/")
    if not text:
        return None
    rel = PurePosixPath(text)
    if rel.is_absolute():
        return None
    parts = [p for p in rel.parts if p not in {"", "."}]
    if not parts or any(p == ".." for p in parts):
        return None
    candidate = (bundle_root / Path(*parts)).resolve()
    try:
        candidate.relative_to(bundle_root.resolve())
    except ValueError:
        return None
    return candidate


def _credit_egg_data_uri() -> str:
    """Return a data: URI for the credit easter-egg image, or empty string if missing."""
    img_path = Path(__file__).with_name("credit_egg.png")
    try:
        data = img_path.read_bytes()
    except OSError:
        return ""
    return "data:image/png;base64," + base64.b64encode(data).decode("ascii")


def generate_html_archive(exports_dir: Path, output_path: Path) -> None:
    """Read exports_dir and write a single self-contained HTML archive."""
    all_chats, index_meta = _load_chats(exports_dir, keep_local=False)
    html_content = _build_html(all_chats, index_meta)
    output_path.write_text(html_content, encoding="utf-8")


def generate_html_folder(exports_dir: Path, output_dir: Path) -> dict[str, int]:
    """Write index.html plus an images/ folder into output_dir.

    Mirrored image attachments are copied into ``output_dir/images/`` and linked
    inline so they display directly inside each chat. Returns copy stats.
    """
    all_chats, index_meta = _load_chats(exports_dir, keep_local=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    stats = _copy_images_and_rewrite(exports_dir, all_chats, output_dir / "images")
    html_content = _build_html(all_chats, index_meta)
    (output_dir / "index.html").write_text(html_content, encoding="utf-8")
    return stats


def _load_chats(exports_dir: Path, *, keep_local: bool = False) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Load and normalize all chats from an exports directory."""
    index_path = exports_dir / "index.json"
    if not index_path.exists():
        raise FileNotFoundError(f"No index.json found in {exports_dir}")

    with open(index_path, encoding="utf-8") as f:
        index_data = json.load(f)

    conversations_meta = index_data.get("conversations", [])
    all_chats: list[dict[str, Any]] = []

    for conv in conversations_meta:
        export_path = conv.get("exportPath", "")
        if not export_path:
            continue
        chat_file = exports_dir / export_path
        if not chat_file.exists():
            continue
        try:
            with open(chat_file, encoding="utf-8") as f:
                chat_data = json.load(f)
        except (json.JSONDecodeError, OSError):
            continue

        meta = chat_data.get("meta", {})
        messages = chat_data.get("messages", [])
        if not messages:
            continue

        # Skip system streams (notifications, mentions, calllogs, etc.)
        title_raw = meta.get("title", conv.get("title", ""))
        if "teamsstream_" in title_raw:
            continue

        # Skip chats that have ONLY system messages (no real content)
        has_real_msg = any(
            not m.get("system") and m.get("text")
            for m in messages
        )
        if not has_real_msg:
            continue

        # Derive a better title for ID-only chats (19:xxx...)
        title = title_raw
        if title.startswith("19:") or not title or title == "Untitled":
            authors = sorted(
                {m["author"] for m in messages if m.get("author") and not m.get("system") and m["author"] != "[system]"},
            )
            if authors:
                title = ", ".join(authors[:4])
                if len(authors) > 4:
                    title += f" +{len(authors) - 4}"
            elif title_raw:
                title = title_raw  # keep raw as fallback
            else:
                title = "Ukjent samtale"

        # Determine category: person (1:1), group, meeting
        is_meeting = conv.get("meeting", False)
        non_system_authors = {m["author"] for m in messages if m.get("author") and not m.get("system") and m["author"] != "[system]"}
        if is_meeting:
            category = "meeting"
        elif len(non_system_authors) <= 2:
            category = "person"
        else:
            category = "group"

        # Get last message timestamp for sorting
        last_ts = ""
        for m in reversed(messages):
            if m.get("timestamp"):
                last_ts = m["timestamp"]
                break

        all_chats.append({
            "id": conv.get("id", ""),
            "title": title,
            "messageCount": meta.get("count", len(messages)),
            "timeRange": meta.get("timeRange", ""),
            "startAt": meta.get("startAt", ""),
            "endAt": meta.get("endAt", ""),
            "lastTs": last_ts,
            "hidden": conv.get("hidden", False),
            "meeting": is_meeting,
            "category": category,
            "messages": _clean_messages(messages, keep_local=keep_local),
        })

    # Sort by last activity (most recent first)
    all_chats.sort(key=lambda c: c.get("lastTs", ""), reverse=True)

    # Disambiguate duplicate titles by appending date range or participants
    from collections import Counter
    title_counts = Counter(c["title"] for c in all_chats)
    dupes = {t for t, n in title_counts.items() if n > 1}
    for chat in all_chats:
        if chat["title"] in dupes:
            suffix_parts = []
            if chat.get("timeRange"):
                suffix_parts.append(chat["timeRange"])
            if not suffix_parts:
                # Use first non-system author as disambiguator
                seen = set()
                for m in chat["messages"]:
                    if m.get("author") and not m.get("system") and m["author"] != "[system]":
                        seen.add(m["author"])
                    if len(seen) >= 3:
                        break
                if seen:
                    suffix_parts.append(", ".join(sorted(seen)[:3]))
            if suffix_parts:
                chat["title"] = f'{chat["title"]} ({"; ".join(suffix_parts)})'

    return all_chats, index_data.get("meta", {})


def _is_web_image(path: str) -> bool:
    """True if the path points to an image format browsers can display inline."""
    ext = path.rsplit(".", 1)[-1].lower() if "." in path else ""
    return ext in _WEB_IMAGE_EXTS


def _copy_images_and_rewrite(
    exports_dir: Path,
    chats: list[dict[str, Any]],
    images_dir: Path,
) -> dict[str, int]:
    """Copy mirrored image attachments into images_dir and rewrite localPath.

    Image attachments get a relative ``images/...`` URL so the HTML displays them
    inline. Non-image or missing attachments lose their localPath and fall back
    to the original external link.
    """
    copied = 0
    missing = 0
    for chat in chats:
        for msg in chat["messages"]:
            for att in msg.get("attachments", []):
                local_path = att.get("localPath")
                if not local_path:
                    continue
                if not _is_web_image(local_path):
                    att.pop("localPath", None)
                    continue
                source = _resolve_bundle_relative_path(exports_dir, local_path)
                if source is None or not source.is_file():
                    att.pop("localPath", None)
                    missing += 1
                    continue
                rel = local_path[len("assets/"):] if local_path.startswith("assets/") else local_path
                dest = images_dir / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                if not dest.exists():
                    shutil.copy2(source, dest)
                copied += 1
                # The on-disk segments contain literal percent-encodings, so
                # re-encode for use as a URL inside the HTML.
                att["localPath"] = "images/" + quote(rel, safe="/._-")
    return {"copied": copied, "missing": missing}


def _clean_messages(messages: list[dict[str, Any]], *, keep_local: bool = False) -> list[dict[str, Any]]:
    """Strip large fields we don't need and keep only viewer-relevant data."""
    cleaned = []
    for msg in messages:
        if not isinstance(msg, dict):
            continue
        entry: dict[str, Any] = {
            "author": msg.get("author", ""),
            "timestamp": msg.get("timestamp", ""),
            "text": msg.get("text", ""),
            "system": msg.get("system", False),
            "edited": msg.get("edited", False),
        }
        if msg.get("reactions"):
            entry["reactions"] = msg["reactions"]
        if msg.get("attachments"):
            attachments = []
            for a in msg["attachments"]:
                if not isinstance(a, dict):
                    continue
                item = {"label": a.get("label", ""), "href": a.get("href", ""), "kind": a.get("kind", "")}
                if keep_local and a.get("localPath") and a.get("localStatus") == "mirrored":
                    item["localPath"] = a.get("localPath")
                attachments.append(item)
            entry["attachments"] = attachments
        if msg.get("replyTo") and isinstance(msg["replyTo"], dict):
            entry["replyTo"] = {
                "author": msg["replyTo"].get("author", ""),
                "text": msg["replyTo"].get("text", ""),
            }
        if msg.get("mentions"):
            entry["mentions"] = msg["mentions"]
        if msg.get("importance"):
            entry["importance"] = msg["importance"]
        cleaned.append(entry)
    return cleaned


def _build_html(chats: list[dict[str, Any]], index_meta: dict[str, Any]) -> str:
    # Separate metadata (for sidebar) from messages (lazy-loaded per chat)
    chat_meta = []
    for c in chats:
        chat_meta.append({
            "id": c["id"],
            "title": c["title"],
            "messageCount": c["messageCount"],
            "timeRange": c.get("timeRange", ""),
            "startAt": c.get("startAt", ""),
            "endAt": c.get("endAt", ""),
            "lastTs": c.get("lastTs", ""),
            "hidden": c.get("hidden", False),
            "meeting": c.get("meeting", False),
            "category": c.get("category", ""),
            "preview": _get_preview(c["messages"]),
        })
    meta_json_str = json.dumps(chat_meta, ensure_ascii=False, separators=(",", ":"))
    index_meta_json = json.dumps(index_meta, ensure_ascii=False, separators=(",", ":"))

    # Build per-chat message script blocks
    msg_blocks = []
    for i, c in enumerate(chats):
        msgs_json = json.dumps(c["messages"], ensure_ascii=False, separators=(",", ":"))
        msg_blocks.append(f'<script type="application/json" id="chat-data-{i}">{msgs_json}</script>')
    msg_blocks_str = "\n".join(msg_blocks)
    credit_egg_uri = _credit_egg_data_uri()
    credit_egg_img = (
        f'<img class="credit-img" src="{credit_egg_uri}" alt="">' if credit_egg_uri else ""
    )
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Teams Chat Archive</title>
{_CSS}
</head>
<body>
<div id="app">
  <aside id="sidebar">
    <div class="sidebar-header">
      <div class="sidebar-brand"><div class="brand-icon"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg></div><h1>Teams Archive</h1></div>
      <div class="search-box">
        <input type="text" id="globalSearch" placeholder="Search all chats..." autocomplete="off">
        <button id="clearSearch" class="clear-btn" title="Clear search">&times;</button>
      </div>
      <div class="category-tabs">
        <button class="cat-tab active" data-cat="all">All</button>
        <button class="cat-tab" data-cat="person">People</button>
        <button class="cat-tab" data-cat="group">Groups</button>
        <button class="cat-tab" data-cat="meeting">Meetings</button>
      </div>
      <div class="sidebar-filters">
        <label class="filter-toggle"><input type="checkbox" id="showHidden" checked> Hidden</label>
        <select id="sortOrder" class="sort-select">
          <option value="recent">Most recent</option>
          <option value="alpha">A-Z</option>
          <option value="messages">Most messages</option>
        </select>
      </div>
      <div class="sidebar-stats" id="sidebarStats"></div>
    </div>
    <div id="chatList" class="chat-list"></div>
  </aside>
  <main id="main">
    <div id="welcomeScreen" class="welcome-screen">
      <div class="welcome-content">
        <div class="welcome-icon"><svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg></div>
        <h2>Teams Chat Archive</h2>
        <p>Select a conversation from the list, or use the search bar to find messages across all chats.</p>
        <div id="welcomeStats" class="welcome-stats"></div>
      </div>
    </div>
    <div id="chatView" class="chat-view" style="display:none">
      <div class="chat-header">
        <div class="chat-title-row">
          <button id="backBtn" class="btn-back" style="display:none" title="Back to search results">←</button>
          <h2 id="chatTitle"></h2>
          <span id="chatMeta" class="chat-meta-info"></span>
        </div>
        <div class="chat-toolbar">
          <input type="text" id="chatSearch" placeholder="Search this chat..." autocomplete="off">
          <input type="date" id="dateFrom" title="From date">
          <input type="date" id="dateTo" title="To date">
          <label class="filter-toggle"><input type="checkbox" id="hideSystem"> Hide system</label>
          <select id="authorFilter"><option value="">All authors</option></select>
          <button id="clearChatFilters" class="btn-small">Reset</button>
        </div>
        <div id="searchResultsBar" class="search-results-bar" style="display:none">
          <span id="searchResultCount"></span>
          <button id="prevResult" class="btn-small">&uarr;</button>
          <button id="nextResult" class="btn-small">&darr;</button>
        </div>
      </div>
      <div id="messages" class="messages"></div>
    </div>
    <div id="globalResults" class="global-results" style="display:none">
      <div class="chat-header">
        <h2 id="globalResultsTitle">Search results</h2>
      </div>
      <div id="globalResultsList" class="messages"></div>
    </div>
  </main>
</div>
<div class="credit-egg" aria-hidden="true" title=""><span class="credit-dot">&#9728;</span>{credit_egg_img}<span class="credit-text">Crafted with &#9829; by Jon-Erik Tyvand</span></div>
{msg_blocks_str}
<script>
const CHAT_META={meta_json_str};
const INDEX_META={index_meta_json};
{_JS}
</script>
</body>
</html>"""


def _get_preview(messages: list[dict[str, Any]]) -> str:
    """Return last real message text as preview for sidebar."""
    for m in reversed(messages):
        if not m.get("system") and m.get("text") and m.get("author"):
            author = m["author"]
            text = (m["text"] or "").replace("\n", " ")[:80]
            return f"{author}: {text}"
    return ""


_CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#ffffff;--bg-alt:#f8f9fb;--bg-sidebar:#fafbfc;--bg-msg:#ffffff;
  --text:#1a1d26;--text-muted:#6b7280;--text-light:#9ca3af;
  --border:#e5e7eb;--border-light:#f0f1f3;
  --accent:#4f46e5;--accent-hover:#4338ca;--accent-light:#818cf8;--accent-bg:rgba(79,70,229,0.06);
  --highlight:#fef9c3;--highlight-border:#eab308;
  --reply-bg:#f0f4ff;--reply-border:#818cf8;
  --system-bg:transparent;--system-text:#9ca3af;
  --hover:#f3f4f6;--selected:rgba(79,70,229,0.06);
  --shadow:0 1px 2px rgba(0,0,0,0.04);
  --shadow-lg:0 4px 12px rgba(0,0,0,0.06);
  --radius:8px;--radius-lg:12px;
  --msg-in:#ffffff;--msg-hover:#f9fafb;
  --transition:all .15s ease;
  --font:'Inter',system-ui,-apple-system,'Segoe UI',sans-serif;
}
@media(prefers-color-scheme:dark){
  :root{
    --bg:#111318;--bg-alt:#181a21;--bg-sidebar:#151720;--bg-msg:#181a21;
    --text:#e4e5e9;--text-muted:#8b8fa3;--text-light:#5c5f73;
    --border:#2a2d3a;--border-light:#22242e;
    --accent:#818cf8;--accent-hover:#a5b4fc;--accent-light:#6366f1;--accent-bg:rgba(129,140,248,0.08);
    --highlight:rgba(234,179,8,0.12);--highlight-border:#eab308;
    --reply-bg:rgba(99,102,241,0.08);--reply-border:#6366f1;
    --system-bg:transparent;--system-text:#5c5f73;
    --hover:#1e2030;--selected:rgba(129,140,248,0.08);
    --shadow:0 1px 2px rgba(0,0,0,0.2);
    --shadow-lg:0 4px 12px rgba(0,0,0,0.3);
    --msg-in:#181a21;--msg-hover:#1e2030;
  }
}
html,body{height:100%;font-family:var(--font);background:var(--bg);color:var(--text);-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;font-size:14px;line-height:1.5}
#app{display:flex;height:100vh;overflow:hidden}

/* ── Sidebar ── */
#sidebar{width:340px;min-width:280px;max-width:400px;background:var(--bg-sidebar);border-right:1px solid var(--border);display:flex;flex-direction:column;overflow:hidden}
.sidebar-header{padding:20px 16px 14px;border-bottom:1px solid var(--border)}
.sidebar-brand{display:flex;align-items:center;gap:10px;margin-bottom:16px}
.brand-icon{display:flex;align-items:center;justify-content:center;width:36px;height:36px;border-radius:10px;background:var(--accent);color:#fff;flex-shrink:0}
.sidebar-header h1{font-size:16px;font-weight:700;color:var(--text);letter-spacing:-.02em}
.search-box{position:relative}
.search-box input{width:100%;padding:9px 32px 9px 36px;border:1px solid var(--border);border-radius:var(--radius);font-size:13px;background:var(--bg);color:var(--text);outline:none;transition:var(--transition);font-family:var(--font)}
.search-box input::placeholder{color:var(--text-light)}
.search-box input:focus{border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-bg)}
.search-box::before{content:'';position:absolute;left:11px;top:50%;transform:translateY(-50%);width:16px;height:16px;background:var(--text-light);-webkit-mask:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.3-4.3'/%3E%3C/svg%3E") center/contain no-repeat;mask:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.3-4.3'/%3E%3C/svg%3E") center/contain no-repeat;pointer-events:none}
.clear-btn{position:absolute;right:8px;top:50%;transform:translateY(-50%);background:none;border:none;font-size:16px;color:var(--text-light);cursor:pointer;display:none;width:22px;height:22px;border-radius:4px;transition:var(--transition);line-height:1}
.clear-btn:hover{background:var(--hover);color:var(--text)}
.sidebar-filters{display:flex;gap:12px;margin-top:10px;font-size:12px;color:var(--text-muted);align-items:center;justify-content:space-between}
.filter-toggle{display:flex;align-items:center;gap:5px;cursor:pointer;user-select:none;white-space:nowrap;transition:var(--transition);font-size:12px;color:var(--text-muted)}
.filter-toggle:hover{color:var(--text)}
.filter-toggle input{accent-color:var(--accent);width:14px;height:14px}
.sidebar-stats{margin-top:8px;font-size:11px;color:var(--text-light);font-weight:500;letter-spacing:.01em}

/* ── Category tabs ── */
.category-tabs{display:flex;gap:2px;margin-top:12px;background:var(--bg);border-radius:var(--radius);padding:2px;border:1px solid var(--border)}
.cat-tab{flex:1;padding:6px 4px;border:none;border-radius:6px;background:transparent;color:var(--text-muted);font-size:11px;font-weight:600;cursor:pointer;transition:var(--transition);white-space:nowrap;font-family:var(--font)}
.cat-tab:hover{color:var(--text);background:var(--hover)}
.cat-tab.active{background:var(--accent);color:#fff;box-shadow:var(--shadow)}
.sort-select{padding:4px 8px;border:1px solid var(--border);border-radius:6px;font-size:11px;background:var(--bg);color:var(--text-muted);outline:none;cursor:pointer;transition:var(--transition);font-family:var(--font)}
.sort-select:focus{border-color:var(--accent)}

/* ── Back button ── */
.btn-back{background:none;border:1px solid var(--border);border-radius:6px;font-size:16px;padding:2px 8px;cursor:pointer;color:var(--text-muted);transition:var(--transition);margin-right:4px;line-height:1}
.btn-back:hover{background:var(--hover);border-color:var(--accent);color:var(--accent)}

/* ── Chat list ── */
.chat-list{overflow-y:auto;flex:1;scrollbar-width:thin;scrollbar-color:var(--border) transparent}
.chat-list::-webkit-scrollbar{width:5px}
.chat-list::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
.chat-list::-webkit-scrollbar-thumb:hover{background:var(--text-light)}
.chat-item{padding:12px 16px;border-bottom:1px solid var(--border-light);cursor:pointer;transition:var(--transition);position:relative}
.chat-item:hover{background:var(--hover)}
.chat-item.active{background:var(--selected);border-left:2px solid var(--accent);padding-left:14px}
.chat-item-title{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--text);line-height:1.3}
.chat-item-preview{font-size:12px;color:var(--text-muted);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.4}
.chat-item-meta{font-size:11px;color:var(--text-light);margin-top:4px;display:flex;justify-content:space-between;font-weight:500}
.chat-item-badge{font-size:10px;padding:1px 6px;border-radius:4px;margin-left:6px;font-weight:600;letter-spacing:.02em;text-transform:uppercase}
.chat-item-match{font-size:12px;color:var(--text);margin-top:6px;padding:6px 10px;background:var(--highlight);border-left:2px solid var(--highlight-border);border-radius:0 4px 4px 0;line-height:1.4}
.chat-item-match mark{background:var(--highlight-border);color:#000;padding:0 2px;border-radius:2px;font-weight:600}

/* ── Main content ── */
#main{flex:1;display:flex;flex-direction:column;overflow:hidden;background:var(--bg)}

/* ── Welcome screen ── */
.welcome-screen{flex:1;display:flex;align-items:center;justify-content:center;text-align:center;color:var(--text-muted)}
.welcome-content{max-width:380px;padding:40px}
.welcome-icon{margin-bottom:16px;color:var(--accent);opacity:.7}
.welcome-content h2{font-size:22px;font-weight:700;color:var(--text);margin-bottom:8px;letter-spacing:-.02em}
.welcome-content p{font-size:14px;line-height:1.6;color:var(--text-muted)}
.welcome-stats{margin-top:24px;display:flex;gap:32px;justify-content:center}
.welcome-stats .stat{display:flex;flex-direction:column;align-items:center}
.welcome-stats strong{color:var(--accent);font-size:28px;font-weight:700;line-height:1.2}
.welcome-stats .stat-label{font-size:12px;color:var(--text-light);margin-top:2px;font-weight:500}

/* ── Chat view ── */
.chat-view,.global-results{flex:1;display:flex;flex-direction:column;overflow:hidden}
.chat-header{padding:14px 24px;border-bottom:1px solid var(--border);background:var(--bg-alt)}
.chat-title-row{display:flex;align-items:baseline;gap:12px}
.chat-title-row h2{font-size:16px;font-weight:700;color:var(--text);letter-spacing:-.01em}
.chat-meta-info{font-size:11px;color:var(--text-light);font-weight:500}

/* ── Toolbar ── */
.chat-toolbar{display:flex;gap:6px;margin-top:10px;flex-wrap:wrap;align-items:center}
.chat-toolbar input[type=text],.chat-toolbar input[type=date],.chat-toolbar select{padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px;background:var(--bg);color:var(--text);outline:none;transition:var(--transition);font-family:var(--font)}
.chat-toolbar input[type=text]{flex:1;min-width:180px}
.chat-toolbar input:focus,.chat-toolbar select:focus{border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-bg)}
.btn-small{padding:5px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text-muted);cursor:pointer;font-size:11px;font-weight:500;transition:var(--transition);font-family:var(--font)}
.btn-small:hover{background:var(--hover);border-color:var(--accent);color:var(--text)}
.search-results-bar{display:flex;align-items:center;gap:6px;margin-top:8px;font-size:12px;color:var(--text-muted);font-weight:500}

/* ── Messages ── */
.messages{flex:1;overflow-y:auto;padding:12px 24px;scrollbar-width:thin;scrollbar-color:var(--border) transparent}
.messages::-webkit-scrollbar{width:5px}
.messages::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
.msg{padding:8px 14px;margin-bottom:1px;border-radius:6px;transition:background .1s ease}
.msg:hover{background:var(--msg-hover)}
.msg.highlight{background:var(--highlight);border-left:2px solid var(--highlight-border);padding-left:12px}
.msg.current-match{background:var(--highlight);border-left:2px solid #f59e0b;padding-left:12px;box-shadow:0 0 0 1px rgba(245,158,11,0.2)}
.msg-header{display:flex;align-items:baseline;gap:8px;margin-bottom:2px}
.msg-author{font-weight:600;font-size:13px}
.msg-time{font-size:11px;color:var(--text-light);font-weight:400}
.msg-edited{font-size:10px;color:var(--text-light);font-style:italic;padding:1px 5px;background:var(--hover);border-radius:3px}
.msg-body{font-size:13.5px;line-height:1.65;white-space:pre-wrap;word-break:break-word;color:var(--text)}
.msg-body a{color:var(--accent);text-decoration:none;border-bottom:1px solid transparent;transition:var(--transition)}
.msg-body a:hover{border-bottom-color:var(--accent)}
.msg-body mark{background:var(--highlight-border);color:#000;padding:0 2px;border-radius:2px;font-weight:600}

/* ── Message grouping ── */
.msg.msg-grouped{padding-top:0;margin-top:-1px}
.msg.msg-grouped .msg-header{display:none}
.msg.msg-grouped:hover .msg-time-inline{opacity:1}
.msg-time-inline{font-size:10px;color:var(--text-light);opacity:0;transition:opacity .1s;float:left;width:40px;margin-left:-46px;text-align:right;margin-top:2px}

/* ── Author colors ── */
.msg-author-dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px;vertical-align:middle}
.msg-system{color:var(--system-text);font-size:12px;padding:4px 14px;text-align:center}
.msg-system .msg-body{font-size:12px;font-style:italic}

/* ── Reply ── */
.msg-reply{margin-bottom:6px;padding:6px 10px;background:var(--reply-bg);border-left:2px solid var(--reply-border);border-radius:0 6px 6px 0;font-size:12px}
.msg-reply-author{font-weight:600;font-size:11px;color:var(--accent)}
.msg-reply-text{color:var(--text-muted);margin-top:2px;line-height:1.4}

/* ── Reactions ── */
.msg-reactions{display:flex;gap:4px;margin-top:4px;flex-wrap:wrap}
.reaction{display:inline-flex;align-items:center;gap:3px;padding:2px 8px;background:var(--bg-alt);border:1px solid var(--border);border-radius:12px;font-size:12px;cursor:default;transition:var(--transition)}
.reaction:hover{border-color:var(--accent);background:var(--accent-bg)}
.reaction-count{color:var(--text-muted);font-weight:600;font-size:11px}

/* ── Attachments ── */
.msg-attachments{margin-top:6px;display:flex;flex-wrap:wrap;gap:4px}
.msg-attachments a{color:var(--accent);font-size:12px;text-decoration:none;display:inline-flex;align-items:center;gap:4px;padding:3px 8px;background:var(--accent-bg);border-radius:6px;font-weight:500;transition:var(--transition);border:1px solid transparent}
.msg-attachments a:hover{background:var(--accent);color:#fff;border-color:var(--accent)}
.msg-attachments a::before{content:'📎';font-size:11px}
.msg-attachments .msg-image-link{padding:0;background:none;border:none}
.msg-attachments .msg-image-link::before{content:none}
.msg-attachments .msg-image-link:hover{background:none}
.msg-image{display:block;max-width:min(420px,100%);max-height:420px;width:auto;height:auto;border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);object-fit:contain}
.msg-image-link:hover .msg-image{box-shadow:var(--shadow-lg)}
.msg-mentions{font-size:11px;color:var(--accent);margin-top:3px;font-weight:500}
.msg-importance{font-size:10px;color:#dc2626;font-weight:600;padding:1px 6px;background:rgba(220,38,38,0.06);border-radius:3px}

/* ── Day separator ── */
.day-separator{display:flex;align-items:center;gap:12px;padding:20px 0 8px;font-size:11px;color:var(--text-light);font-weight:600;letter-spacing:.03em;text-transform:uppercase}
.day-separator::before,.day-separator::after{content:'';flex:1;height:1px;background:var(--border)}

/* ── No results ── */
.no-results{text-align:center;padding:60px 20px;color:var(--text-muted);font-size:14px}

/* ── Global results ── */
.global-result-item{padding:12px 14px;margin-bottom:6px;border:1px solid var(--border);border-radius:var(--radius);cursor:pointer;transition:var(--transition);background:var(--bg-alt)}
.global-result-item:hover{background:var(--hover);border-color:var(--accent);box-shadow:var(--shadow)}
.global-result-chat{font-size:12px;font-weight:600;color:var(--accent)}
.global-result-author{font-size:11px;color:var(--text-light);font-weight:500}
.global-result-text{font-size:13px;margin-top:4px;line-height:1.5;color:var(--text)}
.global-result-text mark{background:var(--highlight-border);color:#000;padding:0 2px;border-radius:2px;font-weight:600}

/* ── Responsive ── */
@media(max-width:768px){
  #sidebar{width:100%;max-width:100%;position:absolute;z-index:10;height:100%}
  #sidebar.collapsed{display:none}
  .chat-header{position:relative}
}
.credit-egg{position:fixed;bottom:8px;right:10px;z-index:9999;font-size:11px;color:var(--text-light);opacity:.35;transition:opacity .25s ease;user-select:none;pointer-events:auto;display:flex;align-items:center;gap:6px}
.credit-egg:hover{opacity:1}
.credit-dot{font-size:11px;cursor:default}
.credit-img{width:0;height:64px;object-fit:contain;opacity:0;transition:width .4s ease,opacity .4s ease,margin .4s ease;margin:0;border-radius:6px;pointer-events:none}
.credit-egg:hover .credit-img{width:42px;opacity:1;margin:0 2px}
.credit-text{max-width:0;overflow:hidden;white-space:nowrap;transition:max-width .4s ease,padding .4s ease;padding:0;font-style:italic}
.credit-egg:hover .credit-text{max-width:240px;padding:0 4px}
</style>"""


_JS = r"""
(function(){
'use strict';

const $ = (s,p) => (p||document).querySelector(s);
const $$ = (s,p) => [...(p||document).querySelectorAll(s)];

// State
let selectedChatIdx = -1;
let filteredChats = [];
let currentSearchMatches = [];
let currentMatchIdx = -1;
let activeCategory = 'all';
let sortOrder = 'recent';
let lastGlobalQuery = '';

// Lazy-loaded message cache: idx -> parsed messages array
const messageCache = {};

function getMessages(chatIdx) {
  if (messageCache[chatIdx]) return messageCache[chatIdx];
  const el = document.getElementById('chat-data-' + chatIdx);
  if (!el) return [];
  try {
    messageCache[chatIdx] = JSON.parse(el.textContent);
  } catch { messageCache[chatIdx] = []; }
  return messageCache[chatIdx];
}

// Profile colors: consistent color per author name
const AUTHOR_COLORS = [
  '#4f46e5','#0891b2','#059669','#d97706','#dc2626','#7c3aed',
  '#2563eb','#0d9488','#ca8a04','#e11d48','#6366f1','#0284c7',
  '#16a34a','#ea580c','#9333ea','#0369a1','#15803d','#b91c1c'
];
const authorColorMap = {};
function authorColor(name) {
  if (!name) return AUTHOR_COLORS[0];
  if (authorColorMap[name]) return authorColorMap[name];
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = ((hash << 5) - hash + name.charCodeAt(i)) | 0;
  const color = AUTHOR_COLORS[Math.abs(hash) % AUTHOR_COLORS.length];
  authorColorMap[name] = color;
  return color;
}

// Elements
const globalSearchEl = $('#globalSearch');
const clearSearchEl = $('#clearSearch');
const chatListEl = $('#chatList');
const welcomeEl = $('#welcomeScreen');
const chatViewEl = $('#chatView');
const globalResultsEl = $('#globalResults');
const chatTitleEl = $('#chatTitle');
const chatMetaEl = $('#chatMeta');
const messagesEl = $('#messages');
const chatSearchEl = $('#chatSearch');
const dateFromEl = $('#dateFrom');
const dateToEl = $('#dateTo');
const hideSystemEl = $('#hideSystem');
const authorFilterEl = $('#authorFilter');
const clearChatFiltersEl = $('#clearChatFilters');
const searchResultCountEl = $('#searchResultCount');
const searchResultsBarEl = $('#searchResultsBar');
const prevResultEl = $('#prevResult');
const nextResultEl = $('#nextResult');
const globalResultsListEl = $('#globalResultsList');
const globalResultsTitleEl = $('#globalResultsTitle');
const showHiddenEl = $('#showHidden');
const sortOrderEl = $('#sortOrder');
const catTabs = $$('.cat-tab');
const backBtn = $('#backBtn');

// --- Utility ---
function escapeHtml(s) {
  if (!s) return '';
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function highlightText(text, query) {
  if (!query || !text) return escapeHtml(text);
  const escaped = escapeRegex(query);
  return escapeHtml(text).replace(new RegExp(`(${escapeHtml(escaped)})`, 'gi'), '<mark>$1</mark>');
}

function linkifyText(html) {
  // Convert URLs in already-escaped HTML to clickable links
  return html.replace(/(https?:\/\/[^\s<>&"]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
}

function formatDate(ts) {
  if (!ts) return '';
  try {
    const d = new Date(ts);
    return d.toLocaleDateString('en-GB', {weekday:'long', year:'numeric', month:'long', day:'numeric'});
  } catch { return ts.slice(0,10); }
}

function formatTime(ts) {
  if (!ts) return '';
  try {
    return new Date(ts).toLocaleTimeString('en-GB', {hour:'2-digit', minute:'2-digit'});
  } catch { return ''; }
}

function formatDateTime(ts) {
  if (!ts) return '';
  try {
    const d = new Date(ts);
    return d.toLocaleDateString('en-GB', {year:'numeric', month:'short', day:'numeric'})
      + ' ' + d.toLocaleTimeString('en-GB', {hour:'2-digit', minute:'2-digit'});
  } catch { return ts; }
}

function dayKey(ts) {
  if (!ts) return '';
  return ts.slice(0,10);
}

function formatRelativeDate(ts) {
  if (!ts) return '';
  try {
    const d = new Date(ts);
    const now = new Date();
    const diff = now - d;
    const days = Math.floor(diff / 86400000);
    if (days === 0) return 'Today';
    if (days === 1) return 'Yesterday';
    if (days < 7) return d.toLocaleDateString('en-GB', {weekday:'long'});
    if (days < 365) return d.toLocaleDateString('en-GB', {day:'numeric', month:'short'});
    return d.toLocaleDateString('en-GB', {day:'numeric', month:'short', year:'numeric'});
  } catch { return ''; }
}

// Check if two messages should be grouped (same author, within 2 min)
function shouldGroup(prev, curr) {
  if (!prev || !curr) return false;
  if (prev.system || curr.system) return false;
  if (prev.author !== curr.author) return false;
  if (curr.replyTo) return false;
  if (!prev.timestamp || !curr.timestamp) return false;
  const diff = new Date(curr.timestamp) - new Date(prev.timestamp);
  return diff >= 0 && diff < 120000; // 2 minutes
}

// --- Sidebar ---
function updateSidebar() {
  const query = globalSearchEl.value.trim().toLowerCase();
  const showHidden = showHiddenEl.checked;

  clearSearchEl.style.display = query ? 'block' : 'none';

  filteredChats = [];
  for (let i = 0; i < CHAT_META.length; i++) {
    const c = CHAT_META[i];
    if (!showHidden && c.hidden) continue;
    if (activeCategory !== 'all' && c.category !== activeCategory) continue;
    filteredChats.push({idx: i, chat: c});
  }

  // Sort
  if (sortOrder === 'alpha') {
    filteredChats.sort((a,b) => a.chat.title.localeCompare(b.chat.title, 'nb'));
  } else if (sortOrder === 'messages') {
    filteredChats.sort((a,b) => (b.chat.messageCount||0) - (a.chat.messageCount||0));
  } else {
    filteredChats.sort((a,b) => (b.chat.lastTs||'').localeCompare(a.chat.lastTs||''));
  }

  // Update stats
  const counts = {all:0, person:0, group:0, meeting:0};
  for (let i = 0; i < CHAT_META.length; i++) {
    const c = CHAT_META[i];
    if (!showHidden && c.hidden) continue;
    counts.all++;
    if (c.category) counts[c.category]++;
  }
  const statsEl = $('#sidebarStats');
  if (statsEl) statsEl.textContent = `${filteredChats.length} of ${counts.all} conversations`;

  // Update tab counts
  for (const tab of catTabs) {
    const cat = tab.dataset.cat;
    const count = counts[cat] || 0;
    const labels = {all:'All', person:'People', group:'Groups', meeting:'Meetings'};
    tab.textContent = `${labels[cat]} (${count})`;
  }

  // If global search is active, perform content search
  if (query.length >= 2) {
    showGlobalResults(query);
    return;
  }

  globalResultsEl.style.display = 'none';
  if (selectedChatIdx >= 0) {
    chatViewEl.style.display = 'flex';
    welcomeEl.style.display = 'none';
  } else {
    chatViewEl.style.display = 'none';
    welcomeEl.style.display = 'flex';
  }

  renderChatList('');
}

function renderChatList(query) {
  const items = filteredChats.map(({idx, chat}) => {
    const active = idx === selectedChatIdx ? ' active' : '';
    const badges = [];
    if (chat.hidden) badges.push('<span class="chat-item-badge" style="background:rgba(100,116,139,0.1);color:#64748b">hidden</span>');
    if (chat.category === 'meeting') badges.push('<span class="chat-item-badge" style="background:rgba(124,58,237,0.08);color:#7c3aed">meeting</span>');
    if (chat.category === 'group') badges.push('<span class="chat-item-badge" style="background:rgba(5,150,105,0.08);color:#059669">group</span>');
    const lastDate = formatRelativeDate(chat.lastTs);

    return `<div class="chat-item${active}" data-idx="${idx}">
      <div class="chat-item-title">${escapeHtml(chat.title)}${badges.join('')}</div>
      <div class="chat-item-preview">${escapeHtml(chat.preview || '').slice(0,90)}</div>
      <div class="chat-item-meta">
        <span>${chat.messageCount} messages</span>
        <span>${lastDate}</span>
      </div>
    </div>`;
  });

  chatListEl.innerHTML = items.join('') || '<div class="no-results">No conversations found</div>';

  // Attach click handlers
  for (const el of $$('.chat-item', chatListEl)) {
    el.addEventListener('click', () => {
      const idx = parseInt(el.dataset.idx, 10);
      selectChat(idx);
    });
  }
}

// --- Global search ---
function showGlobalResults(query) {
  welcomeEl.style.display = 'none';
  chatViewEl.style.display = 'none';
  globalResultsEl.style.display = 'flex';
  lastGlobalQuery = query;

  const results = [];
  const lowerQuery = query.toLowerCase();
  const MAX_RESULTS = 200;

  for (let ci = 0; ci < CHAT_META.length && results.length < MAX_RESULTS; ci++) {
    const chat = CHAT_META[ci];
    if (!showHiddenEl.checked && chat.hidden) continue;
    if (activeCategory !== 'all' && chat.category !== activeCategory) continue;

    const titleMatch = chat.title.toLowerCase().includes(lowerQuery);

    // Lazy load messages for search
    const msgs = getMessages(ci);
    for (let mi = 0; mi < msgs.length && results.length < MAX_RESULTS; mi++) {
      const msg = msgs[mi];
      const text = (msg.text || '').toLowerCase();
      const author = (msg.author || '').toLowerCase();
      if (text.includes(lowerQuery) || author.includes(lowerQuery) || titleMatch && mi === 0) {
        results.push({chatIdx: ci, msgIdx: mi, chat, msg});
        if (titleMatch && mi === 0) continue;
      }
    }
  }

  // Render sidebar with matching chats highlighted
  const matchingChatIndices = new Set(results.map(r => r.chatIdx));
  const sidebarItems = filteredChats
    .filter(({idx}) => matchingChatIndices.has(idx))
    .map(({idx, chat}) => {
      const matchCount = results.filter(r => r.chatIdx === idx).length;
      const firstMatch = results.find(r => r.chatIdx === idx);
      const preview = firstMatch ? firstMatch.msg.text || '' : '';
      const active = idx === selectedChatIdx ? ' active' : '';
      return `<div class="chat-item${active}" data-idx="${idx}">
        <div class="chat-item-title">${escapeHtml(chat.title)}</div>
        <div class="chat-item-meta">
          <span>${matchCount} matches</span>
          <span>${chat.messageCount} messages</span>
        </div>
        <div class="chat-item-match">${highlightText(preview.slice(0,120), query)}</div>
      </div>`;
    });

  chatListEl.innerHTML = sidebarItems.join('') || '<div class="no-results">No results</div>';
  for (const el of $$('.chat-item', chatListEl)) {
    el.addEventListener('click', () => {
      const idx = parseInt(el.dataset.idx, 10);
      selectChat(idx, query);
    });
  }

  globalResultsTitleEl.textContent = `Search results: ${results.length} matches for "${query}"`;

  const html = results.slice(0, 100).map(r => {
    const snippet = r.msg.text || '[no text]';
    return `<div class="global-result-item" data-chat="${r.chatIdx}" data-msg="${r.msgIdx}">
      <div class="global-result-chat">${escapeHtml(r.chat.title)}</div>
      <div class="global-result-author">${escapeHtml(r.msg.author)} · ${formatDateTime(r.msg.timestamp)}</div>
      <div class="global-result-text">${highlightText(snippet.slice(0, 200), query)}</div>
    </div>`;
  });

  globalResultsListEl.innerHTML = html.join('') || '<div class="no-results">No results</div>';

  for (const el of $$('.global-result-item', globalResultsListEl)) {
    el.addEventListener('click', () => {
      const ci = parseInt(el.dataset.chat, 10);
      const mi = parseInt(el.dataset.msg, 10);
      selectChat(ci, query, mi);
    });
  }
}

// --- Chat view ---
function selectChat(chatIdx, searchQuery, scrollToMsgIdx) {
  selectedChatIdx = chatIdx;
  const chat = CHAT_META[chatIdx];

  welcomeEl.style.display = 'none';
  globalResultsEl.style.display = 'none';
  chatViewEl.style.display = 'flex';

  chatTitleEl.textContent = chat.title;
  chatMetaEl.textContent = `${chat.messageCount} messages · ${chat.timeRange || ''}`;

  // Show back button if coming from global search
  if (searchQuery && lastGlobalQuery) {
    backBtn.style.display = '';
  } else {
    backBtn.style.display = 'none';
  }

  // Update URL hash
  history.replaceState(null, '', '#chat=' + chatIdx);

  // Update active state in sidebar
  for (const el of $$('.chat-item', chatListEl)) {
    el.classList.toggle('active', parseInt(el.dataset.idx, 10) === chatIdx);
  }

  // Populate author filter (from lazy-loaded messages)
  const messages = getMessages(chatIdx);
  const authorSet = new Set();
  for (const m of messages) if (m.author && !m.system) authorSet.add(m.author);
  const authors = [...authorSet].sort((a,b) => a.localeCompare(b));
  authorFilterEl.innerHTML = '<option value="">All authors</option>' +
    authors.map(a => `<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join('');

  // Set search if coming from global search
  if (searchQuery) {
    chatSearchEl.value = searchQuery;
  }

  renderMessages(scrollToMsgIdx);
}

function renderMessages(scrollToMsgIdx) {
  if (selectedChatIdx < 0) return;
  const chat = CHAT_META[selectedChatIdx];
  const allMessages = getMessages(selectedChatIdx);
  if (!allMessages.length) {
    messagesEl.innerHTML = '<div class="no-results">No messages</div>';
    return;
  }

  const query = chatSearchEl.value.trim();
  const lowerQuery = query.toLowerCase();
  const dateFrom = dateFromEl.value;
  const dateTo = dateToEl.value;
  const hideSystem = hideSystemEl.checked;
  const authorFilter = authorFilterEl.value;

  let messages = allMessages;

  // Apply filters
  messages = messages.filter(m => {
    if (hideSystem && m.system) return false;
    if (authorFilter && m.author !== authorFilter) return false;
    if (dateFrom && m.timestamp && m.timestamp.slice(0,10) < dateFrom) return false;
    if (dateTo && m.timestamp && m.timestamp.slice(0,10) > dateTo) return false;
    return true;
  });

  // Find search matches
  currentSearchMatches = [];
  if (lowerQuery.length >= 2) {
    messages.forEach((m, i) => {
      const text = (m.text || '').toLowerCase();
      const author = (m.author || '').toLowerCase();
      if (text.includes(lowerQuery) || author.includes(lowerQuery)) {
        currentSearchMatches.push(i);
      }
    });
  }

  // Render with grouping
  let lastDay = '';
  const parts = [];

  for (let i = 0; i < messages.length; i++) {
    const m = messages[i];
    const prevMsg = i > 0 ? messages[i-1] : null;
    const day = dayKey(m.timestamp);
    if (day !== lastDay) {
      lastDay = day;
      parts.push(`<div class="day-separator">${formatDate(m.timestamp)}</div>`);
    }

    const isMatch = currentSearchMatches.includes(i);
    const matchClass = isMatch ? ' highlight' : '';
    const grouped = shouldGroup(prevMsg, m) && day === dayKey(prevMsg?.timestamp);

    if (m.system) {
      parts.push(`<div class="msg msg-system${matchClass}" data-idx="${i}">
        <div class="msg-body">${lowerQuery ? highlightText(m.text, query) : escapeHtml(m.text || '')}</div>
      </div>`);
      continue;
    }

    let replyHtml = '';
    if (m.replyTo) {
      replyHtml = `<div class="msg-reply">
        <div class="msg-reply-author">${escapeHtml(m.replyTo.author)}</div>
        <div class="msg-reply-text">${escapeHtml((m.replyTo.text || '').slice(0, 200))}</div>
      </div>`;
    }

    let reactionsHtml = '';
    if (m.reactions && m.reactions.length) {
      reactionsHtml = '<div class="msg-reactions">' +
        m.reactions.map(r =>
          `<span class="reaction" title="${(r.reactors||[]).map(escapeHtml).join(', ')}">${escapeHtml(r.emoji)} <span class="reaction-count">${r.count}</span></span>`
        ).join('') + '</div>';
    }

    let attachmentsHtml = '';
    if (m.attachments && m.attachments.length) {
      attachmentsHtml = '<div class="msg-attachments">' +
        m.attachments.map(a => {
          const label = escapeHtml(a.label || a.href || 'attachment');
          if (a.localPath) {
            const src = escapeHtml(a.localPath);
            return `<a class="msg-image-link" href="${src}" target="_blank" rel="noopener"><img class="msg-image" src="${src}" alt="${label}" loading="lazy"></a>`;
          }
          if (a.href) return `<a href="${escapeHtml(a.href)}" target="_blank" rel="noopener">${label}</a>`;
          return `<span>${label}</span>`;
        }).join(' ') + '</div>';
    }

    let mentionsHtml = '';
    if (m.mentions && m.mentions.length) {
      const names = m.mentions.map(n => typeof n === 'object' ? n.name : n).filter(Boolean);
      if (names.length) mentionsHtml = `<div class="msg-mentions">@${names.map(escapeHtml).join(', @')}</div>`;
    }

    let importanceHtml = '';
    if (m.importance && m.importance !== 'normal') {
      importanceHtml = `<span class="msg-importance">❗ ${escapeHtml(m.importance)}</span>`;
    }

    let bodyText = lowerQuery ? highlightText(m.text, query) : escapeHtml(m.text || '');
    bodyText = linkifyText(bodyText);

    const color = authorColor(m.author);
    const groupedClass = grouped ? ' msg-grouped' : '';
    const timeInline = grouped ? `<span class="msg-time-inline">${formatTime(m.timestamp)}</span>` : '';

    parts.push(`<div class="msg${matchClass}${groupedClass}" data-idx="${i}">
      ${replyHtml}
      <div class="msg-header">
        <span class="msg-author-dot" style="background:${color}"></span><span class="msg-author" style="color:${color}">${escapeHtml(m.author)}</span>
        <span class="msg-time">${formatTime(m.timestamp)}</span>
        ${m.edited ? '<span class="msg-edited">edited</span>' : ''}
        ${importanceHtml}
      </div>
      ${timeInline}
      <div class="msg-body">${bodyText}</div>
      ${reactionsHtml}
      ${attachmentsHtml}
      ${mentionsHtml}
    </div>`);
  }

  messagesEl.innerHTML = parts.join('') || '<div class="no-results">No messages match these filters</div>';

  // Search results navigation
  if (currentSearchMatches.length > 0) {
    searchResultsBarEl.style.display = 'flex';
    searchResultCountEl.textContent = `${currentSearchMatches.length} matches`;
    if (scrollToMsgIdx !== undefined) {
      const targetIdx = messages.indexOf(allMessages[scrollToMsgIdx]);
      if (targetIdx >= 0) {
        currentMatchIdx = currentSearchMatches.indexOf(targetIdx);
        if (currentMatchIdx < 0) currentMatchIdx = 0;
      } else {
        currentMatchIdx = 0;
      }
    } else {
      currentMatchIdx = 0;
    }
    scrollToMatch();
  } else {
    searchResultsBarEl.style.display = 'none';
    currentMatchIdx = -1;
    if (scrollToMsgIdx !== undefined) {
      const el = messagesEl.querySelector(`[data-idx="${scrollToMsgIdx}"]`);
      if (el) {
        el.classList.add('current-match');
        requestAnimationFrame(() => el.scrollIntoView({behavior:'smooth', block:'center'}));
      }
    } else {
      // Scroll to bottom (newest messages) by default
      requestAnimationFrame(() => { messagesEl.scrollTop = messagesEl.scrollHeight; });
    }
  }
}

function scrollToMatch() {
  const prev = messagesEl.querySelector('.current-match');
  if (prev) prev.classList.remove('current-match');

  if (currentMatchIdx < 0 || currentMatchIdx >= currentSearchMatches.length) return;

  const msgIdx = currentSearchMatches[currentMatchIdx];
  const el = messagesEl.querySelector(`[data-idx="${msgIdx}"]`);
  if (el) {
    el.classList.add('current-match');
    el.scrollIntoView({behavior:'smooth', block:'center'});
  }
  searchResultCountEl.textContent = `${currentMatchIdx + 1} of ${currentSearchMatches.length} matches`;
}

// --- Welcome stats ---
function showWelcomeStats() {
  const totalChats = CHAT_META.length;
  const totalMessages = CHAT_META.reduce((s,c) => s + (c.messageCount||0), 0);

  $('#welcomeStats').innerHTML = `
    <div class="stat"><strong>${totalChats}</strong><span class="stat-label">conversations</span></div>
    <div class="stat"><strong>${totalMessages.toLocaleString('en-US')}</strong><span class="stat-label">messages</span></div>
  `;
  $('#sidebarStats').textContent = `${totalChats} chats · ${totalMessages.toLocaleString('en-US')} messages`;
}

// --- Event bindings ---
let searchDebounce = null;
globalSearchEl.addEventListener('input', () => {
  clearTimeout(searchDebounce);
  searchDebounce = setTimeout(() => updateSidebar(), 200);
});
clearSearchEl.addEventListener('click', () => {
  globalSearchEl.value = '';
  clearSearchEl.style.display = 'none';
  updateSidebar();
});
showHiddenEl.addEventListener('change', updateSidebar);
sortOrderEl.addEventListener('change', () => {
  sortOrder = sortOrderEl.value;
  updateSidebar();
});
for (const tab of catTabs) {
  tab.addEventListener('click', () => {
    activeCategory = tab.dataset.cat;
    for (const t of catTabs) t.classList.toggle('active', t === tab);
    updateSidebar();
  });
}

// Back button: return to global search results
backBtn.addEventListener('click', () => {
  if (lastGlobalQuery) {
    globalSearchEl.value = lastGlobalQuery;
    backBtn.style.display = 'none';
    chatSearchEl.value = '';
    selectedChatIdx = -1;
    updateSidebar();
  }
});

chatSearchEl.addEventListener('input', () => {
  clearTimeout(searchDebounce);
  searchDebounce = setTimeout(() => renderMessages(), 200);
});
dateFromEl.addEventListener('change', () => renderMessages());
dateToEl.addEventListener('change', () => renderMessages());
hideSystemEl.addEventListener('change', () => renderMessages());
authorFilterEl.addEventListener('change', () => renderMessages());
clearChatFiltersEl.addEventListener('click', () => {
  chatSearchEl.value = '';
  dateFromEl.value = '';
  dateToEl.value = '';
  hideSystemEl.checked = false;
  authorFilterEl.value = '';
  renderMessages();
});

prevResultEl.addEventListener('click', () => {
  if (currentSearchMatches.length === 0) return;
  currentMatchIdx = (currentMatchIdx - 1 + currentSearchMatches.length) % currentSearchMatches.length;
  scrollToMatch();
});
nextResultEl.addEventListener('click', () => {
  if (currentSearchMatches.length === 0) return;
  currentMatchIdx = (currentMatchIdx + 1) % currentSearchMatches.length;
  scrollToMatch();
});

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if (globalSearchEl.value) {
      globalSearchEl.value = '';
      clearSearchEl.style.display = 'none';
      updateSidebar();
    }
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
    e.preventDefault();
    if (selectedChatIdx >= 0 && chatViewEl.style.display !== 'none') {
      chatSearchEl.focus();
    } else {
      globalSearchEl.focus();
    }
  }
  if (e.key === 'F3' || (e.ctrlKey && e.key === 'g')) {
    e.preventDefault();
    if (e.shiftKey) prevResultEl.click();
    else nextResultEl.click();
  }
});

// --- URL hash navigation ---
function handleHash() {
  const hash = location.hash;
  const m = hash.match(/^#chat=(\d+)$/);
  if (m) {
    const idx = parseInt(m[1], 10);
    if (idx >= 0 && idx < CHAT_META.length) {
      selectChat(idx);
      return;
    }
  }
}

// --- Init ---
showWelcomeStats();
updateSidebar();
handleHash();
window.addEventListener('hashchange', handleHash);

})();
"""


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python generate.py <exports_dir> [output_dir]")
        sys.exit(1)
    # Resolve CLI-provided paths defensively. This is an operator-run tool,
    # but normalising guards against accidental relative traversal in the
    # provided arguments before they are passed downstream.
    exports = Path(sys.argv[1]).expanduser().resolve(strict=False)
    output = (
        Path(sys.argv[2]).expanduser().resolve(strict=False)
        if len(sys.argv) > 2
        else exports.parent / "teams-archive"
    )
    stats = generate_html_folder(exports, output)
    print(f"Archive folder written to {output} (images copied: {stats['copied']})")
