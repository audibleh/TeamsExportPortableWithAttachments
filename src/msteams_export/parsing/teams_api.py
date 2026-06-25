from __future__ import annotations

from html.parser import HTMLParser
import json
import re
from typing import Any
from urllib.parse import urlparse

from msteams_export.attachment_policy import keep_attachment


SYSTEM_TYPES = {
    "Event/Call",
    "RichText/Media_CallRecording",
    "RichText/Media_CallTranscript",
    "ThreadActivity/AddMember",
    "ThreadActivity/DeleteMember",
    "ThreadActivity/MemberJoined",
    "ThreadActivity/MemberLeft",
    "ThreadActivity/PictureUpdate",
    "ThreadActivity/PinnedItemsUpdate",
    "ThreadActivity/SpaceDescriptionUpdated",
    "ThreadActivity/SpaceThreadCreated",
    "ThreadActivity/TabUpdated",
    "ThreadActivity/TopicDeleted",
    "ThreadActivity/TopicThreadCreated",
    "ThreadActivity/TopicUpdate",
    "ThreadActivity/UpdateFavDefault",
}

REACTION_EMOJI: dict[str, str] = {
    "ok": "👌",
    "like": "👍",
    "thumbsup": "👍",
    "thumbs_up": "👍",
    "heart": "❤️",
    "laugh": "😂",
    "haha": "😂",
    "surprised": "😮",
    "wow": "😮",
    "sad": "😢",
    "angry": "😡",
    "crossmark": "❌",
    "no": "🚫",
    "skull": "💀",
    "check": "✔️",
    "checkmark": "✔️",
    "clap": "👏",
    "fire": "🔥",
    "100": "💯",
    "eyes": "👀",
    "pray": "🙏",
    "praying": "🙏",
    "muscle": "💪",
    "tada": "🎉",
    "party": "🎉",
    "rocket": "🚀",
    "wave": "👋",
    "thinking": "🤔",
    "cry": "😢",
    "fistbump": "🤜🤛",
    "worry": "😟",
    "shaking": "🫨",
}

BLOCKQUOTE_RE = re.compile(
    r"<blockquote\b(?P<attrs>[^>]*)>(?P<inner>.*?)</blockquote>",
    re.IGNORECASE | re.DOTALL,
)
AUTHOR_TAG_RE = re.compile(
    r"<(?:span\b[^>]*itemtype=[\"'][^\"']*CreatorName[^\"']*[\"'][^>]*|b\b[^>]*|strong\b[^>]*)>(.*?)</(?:span|b|strong)>",
    re.IGNORECASE | re.DOTALL,
)
TAG_RE = re.compile(r"<[^>]+>")
XML_LEAF_RE = re.compile(r"<(?P<tag>[a-z0-9_-]+)>(?P<value>[^<>]*)</(?P=tag)>", re.IGNORECASE | re.DOTALL)
MRI_RE = re.compile(r"(?:8:orgid:|28:)([0-9a-f-]{8,})", re.IGNORECASE)


class _HtmlToTextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip_stack: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style"}:
            self._skip_stack.append(tag)
            return
        if self._skip_stack:
            return
        attr_map = {key.lower(): value for key, value in attrs}
        if tag == "br":
            self.parts.append("\n")
        elif tag in {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre"}:
            self.parts.append("\n")
        elif tag == "img":
            alt = (attr_map.get("alt") or attr_map.get("title") or "").strip()
            if alt and alt.lower() not in {"image", "media", "shared image", "undefined"}:
                self.parts.append(alt)
        elif tag == "video":
            self.parts.append("[Video]")

    def handle_endtag(self, tag: str) -> None:
        if self._skip_stack:
            if self._skip_stack[-1] == tag:
                self._skip_stack.pop()
            return
        if tag in {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre"}:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip_stack:
            self.parts.append(data)

    def text(self) -> str:
        joined = "".join(self.parts)
        joined = joined.replace("\xa0", " ")
        joined = re.sub(r"\n{3,}", "\n\n", joined)
        joined = re.sub(r"[ \t]+\n", "\n", joined)
        joined = re.sub(r"\n[ \t]+", "\n", joined)
        return joined.strip()


def html_to_text(html: str) -> str:
    if not html:
        return ""
    parser = _HtmlToTextParser()
    parser.feed(html)
    parser.close()
    return parser.text()


class _EmbeddedAttachmentParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.attachments: list[dict[str, Any]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "img":
            return
        attr_map = {key.lower(): value for key, value in attrs}
        href = _optional_str(attr_map.get("src"))
        if not href:
            return
        item_type = (attr_map.get("itemtype") or "").strip().lower()
        media_id = _optional_str(attr_map.get("itemid") or attr_map.get("id"))
        extension = _optional_str(attr_map.get("itemscope")) or _optional_str(attr_map.get("data-type"))
        alt = _optional_str(attr_map.get("alt")) or _optional_str(attr_map.get("title"))
        if not _looks_like_embedded_image(href=href, item_type=item_type, media_id=media_id):
            return
        label = _embedded_attachment_label(alt=alt, extension=extension, href=href, media_id=media_id)
        type_value = _embedded_attachment_type(extension=extension, href=href)
        if not keep_attachment(label=label, href=href, type_value=type_value, kind="image"):
            return
        self.attachments.append(
            {
                "href": href,
                "label": label,
                "type": type_value,
                "size": None,
                "owner": None,
                "metaText": "embedded image",
                "kind": "image",
            }
        )


def extract_reply_from_html(html: str) -> tuple[dict[str, Any] | None, str]:
    if not html or "<blockquote" not in html.lower():
        return None, html

    for match in BLOCKQUOTE_RE.finditer(html):
        attrs = match.group("attrs") or ""
        inner = match.group("inner") or ""
        if "reply" not in attrs.lower() and "itemscope" not in attrs.lower():
            continue
        author = ""
        author_match = AUTHOR_TAG_RE.search(inner)
        inner_without_author = inner
        if author_match:
            author = html_to_text(author_match.group(1))
            inner_without_author = AUTHOR_TAG_RE.sub("", inner, count=1)
        quoted_text = html_to_text(inner_without_author)
        if author and quoted_text.startswith(author):
            quoted_text = quoted_text[len(author):].lstrip()
        id_match = re.search(r'itemid=["\']([^"\']+)["\']', attrs, re.IGNORECASE)
        reply_to = {
            "author": author,
            "timestamp": "",
            "text": quoted_text,
            "id": id_match.group(1) if id_match else None,
        }
        clean_html = html[: match.start()] + html[match.end() :]
        return reply_to, clean_html
    return None, html


def extract_embedded_html_attachments(html: str) -> list[dict[str, Any]]:
    if not html or "<img" not in html.lower():
        return []
    parser = _EmbeddedAttachmentParser()
    parser.feed(html)
    parser.close()
    return parser.attachments


def merge_embedded_html_attachments(existing: Any, html: str | None) -> list[dict[str, Any]]:
    attachments: list[dict[str, Any]] = []
    if isinstance(existing, list):
        attachments.extend(item for item in existing if isinstance(item, dict))
    attachments.extend(extract_embedded_html_attachments(html or ""))
    return _dedupe_attachments(attachments)


def convert_api_messages(api_messages: list[dict[str, Any]], conversation_id: str | None = None) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for raw in api_messages:
        converted = convert_one_message(raw, conversation_id)
        if converted is not None:
            result.append(converted)
    result.reverse()
    return result


def convert_one_message(message: dict[str, Any], conversation_id: str | None = None) -> dict[str, Any] | None:
    properties = _ensure_dict(message.get("properties"))
    message_type = str(message.get("messagetype") or "")
    is_system = is_system_message_type(message_type)
    if properties.get("deletetime") and not message.get("content"):
        return None

    raw_content = str(message.get("content") or "")
    reply_to: dict[str, Any] | None = None
    text: str
    if is_system:
        text = parse_system_content(raw_content, message_type)
        attachment_html = raw_content
    elif message_type == "RichText/Html" or raw_content.lstrip().startswith("<"):
        reply_to, clean_html = extract_reply_from_html(raw_content)
        text = html_to_text(clean_html)
        attachment_html = clean_html
    else:
        text = raw_content
        attachment_html = raw_content

    attachments = merge_embedded_html_attachments(convert_attachments(properties), attachment_html)

    return {
        "id": str(message.get("id") or message.get("clientmessageid") or ""),
        "threadId": str(message.get("conversationid") or conversation_id or ""),
        "author": resolve_author(message, is_system),
        "timestamp": str(message.get("originalarrivaltime") or message.get("composetime") or ""),
        "text": text,
        "contentHtml": raw_content or None,
        "messageType": message_type or None,
        "edited": bool(properties.get("edittime")),
        "system": is_system,
        "importance": _optional_str(properties.get("importance")),
        "subject": _optional_str(properties.get("subject")),
        "reactions": convert_reactions(properties),
        "attachments": attachments,
        "replyTo": reply_to,
        "mentions": convert_mentions(properties),
    }


def resolve_author(message: dict[str, Any], is_system: bool) -> str:
    if is_system:
        return "[system]"
    if message.get("imdisplayname"):
        return str(message["imdisplayname"])
    if message.get("fromDisplayNameInToken"):
        return str(message["fromDisplayNameInToken"])
    parts = [str(value) for value in [message.get("fromGivenNameInToken"), message.get("fromFamilyNameInToken")] if value]
    return " ".join(parts)


def is_system_message_type(message_type: str) -> bool:
    normalized = (message_type or "").strip()
    return normalized in SYSTEM_TYPES or normalized.startswith("ThreadActivity/")


def parse_system_content(content: str, message_type: str) -> str:
    xmlish = parse_xmlish_values(content)
    if xmlish:
        parsed = _parse_structured_system_content(xmlish, message_type)
        if parsed:
            return parsed
    parsed_json = _parse_json_system_content(message_type, content)
    if parsed_json:
        return parsed_json
    text = html_to_text(content)
    if text:
        if _looks_like_machine_blob(text):
            return _generic_system_message(message_type)
        return text
    return TAG_RE.sub(" ", content).strip() or message_type.split("/")[-1] or "system event"


def parse_xmlish_values(content: str) -> dict[str, list[str]]:
    values: dict[str, list[str]] = {}
    for match in XML_LEAF_RE.finditer(content):
        tag = str(match.group("tag") or "").lower()
        value = html_to_text(match.group("value") or "")
        values.setdefault(tag, []).append(value)
    return values


def _parse_structured_system_content(values: dict[str, list[str]], message_type: str) -> str | None:
    if message_type == "ThreadActivity/AddMember":
        targets = _filter_targets(values.get("target", []))
        count = len(targets)
        if count <= 0:
            return "Members were added to the chat."
        if count == 1:
            return "1 member was added to the chat."
        return f"{count} members were added to the chat."
    if message_type in {"ThreadActivity/DeleteMember", "ThreadActivity/MemberLeft"}:
        targets = _filter_targets(values.get("target", []))
        if len(targets) == 1:
            return "1 member left or was removed from the chat."
        if len(targets) > 1:
            return f"{len(targets)} members left or were removed from the chat."
        return "A member left or was removed from the chat."
    if message_type == "ThreadActivity/TopicUpdate":
        candidates = values.get("value", []) + values.get("topic", [])
        new_topic = next((item for item in candidates if item), None)
        if new_topic:
            return f"Chat topic was updated to: {new_topic}"
        return "Chat topic was updated."
    if message_type == "ThreadActivity/SpaceThreadCreated":
        topic = next((item for item in values.get("topic", []) if item), None)
        if topic:
            return f"Team space was created: {topic}"
        return "Team space was created."
    if message_type == "ThreadActivity/TopicThreadCreated":
        topic = next((item for item in values.get("topic", []) if item), None)
        if topic:
            return f"Channel created: {topic}"
        return "Channel was created."
    if message_type == "ThreadActivity/TopicDeleted":
        topic = next((item for item in values.get("topic", []) if item), None)
        if topic:
            return f"Channel deleted: {topic}"
        return "Channel was deleted."
    if message_type == "ThreadActivity/PictureUpdate":
        return "Chat picture was updated."
    if message_type == "ThreadActivity/PinnedItemsUpdate":
        return "Pinned items were updated."
    if message_type == "ThreadActivity/SpaceDescriptionUpdated":
        return "Team description was updated."
    if message_type == "ThreadActivity/TabUpdated":
        return "Channel tab was updated."
    if message_type == "ThreadActivity/UpdateFavDefault":
        return "Default favorites setting was updated."
    if message_type == "Event/Call":
        return "Call event"
    if message_type == "RichText/Media_CallRecording":
        return "Call recording was shared."
    if message_type == "RichText/Media_CallTranscript":
        return "Call transcript was shared."
    return None


def _filter_targets(targets: list[str]) -> list[str]:
    seen: set[str] = set()
    filtered: list[str] = []
    for target in targets:
        text = target.strip()
        if not text:
            continue
        if text in seen:
            continue
        seen.add(text)
        filtered.append(text)
    return filtered


def _looks_like_machine_blob(text: str) -> bool:
    if len(text) < 40:
        return False
    token_like = MRI_RE.findall(text)
    if len(token_like) >= 3:
        return True
    digit_ratio = sum(1 for char in text if char.isdigit()) / max(len(text), 1)
    return digit_ratio > 0.25 and ":" in text


def _generic_system_message(message_type: str) -> str:
    mapping = {
        "ThreadActivity/AddMember": "Members were added to the chat.",
        "ThreadActivity/DeleteMember": "A member was removed from the chat.",
        "ThreadActivity/MemberJoined": "A member joined the chat.",
        "ThreadActivity/MemberLeft": "A member left the chat.",
        "ThreadActivity/PictureUpdate": "Chat picture was updated.",
        "ThreadActivity/PinnedItemsUpdate": "Pinned items were updated.",
        "ThreadActivity/SpaceDescriptionUpdated": "Team description was updated.",
        "ThreadActivity/SpaceThreadCreated": "Team space was created.",
        "ThreadActivity/TabUpdated": "Channel tab was updated.",
        "ThreadActivity/TopicDeleted": "Channel was deleted.",
        "ThreadActivity/TopicThreadCreated": "Channel was created.",
        "ThreadActivity/TopicUpdate": "Chat topic was updated.",
        "ThreadActivity/UpdateFavDefault": "Default favorites setting was updated.",
        "Event/Call": "Call event",
        "RichText/Media_CallRecording": "Call recording was shared.",
        "RichText/Media_CallTranscript": "Call transcript was shared.",
    }
    return mapping.get(message_type, message_type.split("/")[-1] or "system event")


def _parse_json_system_content(message_type: str, content: str) -> str | None:
    payload = _parse_json_object(content)
    if not payload:
        return None

    if message_type == "ThreadActivity/TopicThreadCreated":
        topic = _optional_str(payload.get("topic"))
        return f"Channel created: {topic}" if topic else "Channel was created."

    if message_type == "ThreadActivity/TopicDeleted":
        topic = _optional_str(payload.get("topic"))
        return f"Channel deleted: {topic}" if topic else "Channel was deleted."

    if message_type == "ThreadActivity/SpaceThreadCreated":
        topic = _optional_str(payload.get("topic"))
        description = _optional_str(payload.get("description"))
        if topic:
            return f"Team space was created: {topic}"
        if description:
            return "Team space was created with a description."
        return "Team space was created."

    if message_type == "ThreadActivity/SpaceDescriptionUpdated":
        new_value = _optional_str(payload.get("newValue"))
        old_value = _optional_str(payload.get("oldValue"))
        if new_value and old_value and new_value != old_value:
            return "Team description was updated."
        if new_value and not old_value:
            return "Team description was added."
        if old_value and not new_value:
            return "Team description was cleared."
        return "Team description was updated."

    if message_type == "ThreadActivity/TabUpdated":
        new_tab = _parse_json_object(_optional_str(payload.get("newValue")) or "")
        old_tab = _parse_json_object(_optional_str(payload.get("oldValue")) or "")
        current = new_tab or old_tab
        tab_name = _optional_str(current.get("name")) if current else None
        if new_tab and not old_tab:
            return f"Tab added: {tab_name}" if tab_name else "Tab was added."
        if old_tab and not new_tab:
            return f"Tab removed: {tab_name}" if tab_name else "Tab was removed."
        if new_tab and old_tab:
            return f"Tab updated: {tab_name}" if tab_name else "Tab was updated."
        return "Channel tab was updated."

    if message_type == "ThreadActivity/UpdateFavDefault":
        is_default = _optional_str(payload.get("isDefault"))
        if is_default:
            return f"Default favorites setting changed to: {is_default}"
        return "Default favorites setting was updated."

    return None


def _parse_json_object(content: str) -> dict[str, Any] | None:
    if not content:
        return None
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def convert_reactions(properties: dict[str, Any]) -> list[dict[str, Any]]:
    raw = properties.get("emotions")
    if raw is None:
        return []
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
    except json.JSONDecodeError:
        return []
    if not isinstance(parsed, list):
        return []

    reactions: list[dict[str, Any]] = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        key = str(item.get("key") or "").lower()
        users = item.get("users")
        reactors = []
        if isinstance(users, list):
            reactors = [str(user.get("mri") or "") for user in users if isinstance(user, dict) and user.get("mri")]
        reactions.append(
            {
                "emoji": REACTION_EMOJI.get(key, key or ":reaction:"),
                "count": len(reactors) if reactors else 1,
                "reactors": reactors or None,
            }
        )
    return reactions


def convert_attachments(properties: dict[str, Any]) -> list[dict[str, Any]]:
    raw = properties.get("files")
    if raw is None:
        return []
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
    except json.JSONDecodeError:
        return []
    if not isinstance(parsed, list):
        return []

    attachments: list[dict[str, Any]] = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        label = _optional_str(item.get("fileName") or item.get("title")) or "Attachment"
        href = _optional_str(item.get("objectUrl") or item.get("baseUrl"))
        type_value = _optional_str(item.get("fileType"))
        if not keep_attachment(label=label, href=href, type_value=type_value):
            continue
        attachments.append(
            {
                "href": href,
                "label": label,
                "type": type_value,
                "size": humanize_bytes(item.get("fileSize")),
                "owner": None,
                "metaText": None,
            }
        )
    return attachments


def _dedupe_attachments(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[tuple[str, str, str]] = set()
    deduped: list[dict[str, Any]] = []
    for item in items:
        href = _optional_str(item.get("href")) or ""
        label = _optional_str(item.get("label")) or ""
        type_value = _optional_str(item.get("type")) or ""
        key = (href, label, type_value)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)
    return deduped


def convert_mentions(properties: dict[str, Any]) -> list[dict[str, Any]]:
    raw = properties.get("mentions")
    if raw is None:
        return []
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
    except json.JSONDecodeError:
        return []
    if not isinstance(parsed, list):
        return []
    mentions: list[dict[str, Any]] = []
    for item in parsed:
        if not isinstance(item, dict) or not item.get("displayName"):
            continue
        mentions.append(
            {
                "name": str(item["displayName"]),
                "mri": _optional_str(item.get("mri")),
            }
        )
    return mentions


def humanize_bytes(value: Any) -> str | None:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    if number < 1024:
        return f"{number} B"
    if number < 1024 * 1024:
        return f"{number / 1024:.1f} KB"
    if number < 1024 * 1024 * 1024:
        return f"{number / (1024 * 1024):.1f} MB"
    return f"{number / (1024 * 1024 * 1024):.1f} GB"


def _ensure_dict(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    return {}


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text else None


def _looks_like_embedded_image(*, href: str, item_type: str, media_id: str | None) -> bool:
    if "schema.skype.com/emoji" in item_type:
        return False
    if "schema.skype.com/amsimage" in item_type:
        return True
    host = (urlparse(href).hostname or "").lower()
    if host in {"api.asm.skype.com", "eu-api.asm.skype.com"} or host.endswith(
        (".api.asm.skype.com", ".eu-api.asm.skype.com")
    ):
        return True
    if media_id and media_id.startswith("0-"):
        return True
    return False


def _embedded_attachment_label(
    *,
    alt: str | None,
    extension: str | None,
    href: str,
    media_id: str | None,
) -> str:
    normalized_alt = (alt or "").strip()
    if normalized_alt and normalized_alt.lower() not in {"image", "media", "shared image", "undefined", "obrázek"}:
        return normalized_alt
    normalized_extension = _normalize_extension(extension) or _suffix(href) or "png"
    media_token = (media_id or "embedded-image").strip().strip("x_")
    return f"{media_token}.{normalized_extension}"


def _embedded_attachment_type(*, extension: str | None, href: str) -> str:
    normalized_extension = _normalize_extension(extension) or _suffix(href) or "png"
    if "/" in normalized_extension:
        return normalized_extension
    return f"image/{normalized_extension}"


def _normalize_extension(value: str | None) -> str | None:
    if not value:
        return None
    text = value.strip().lower()
    if not text:
        return None
    if text.startswith("image/"):
        return text
    if text in {"png", "jpg", "jpeg", "webp", "bmp", "heic", "heif"}:
        return text
    return None


def _suffix(value: str | None) -> str:
    if not value:
        return ""
    text = value.strip()
    if not text:
        return ""
    parsed = urlparse(text)
    target = parsed.path or text
    if "." not in target:
        return ""
    return target.rsplit(".", 1)[-1].strip().lower()
