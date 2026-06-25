from __future__ import annotations

from collections import defaultdict
from typing import Any

from msteams_export.attachment_policy import keep_attachment

from msteams_export.parsing.teams_api import html_to_text


def convert_graph_channel_messages(
    raw_messages: list[dict[str, Any]],
    *,
    conversation_id: str | None = None,
) -> list[dict[str, Any]]:
    flattened: list[dict[str, Any]] = []
    for item in raw_messages:
        flattened.extend(_flatten_graph_message(item))

    converted = [
        message
        for message in (
            convert_one_graph_message(item, conversation_id=conversation_id) for item in flattened if isinstance(item, dict)
        )
        if message is not None
    ]
    message_map = {str(item.get("id") or ""): item for item in converted if item.get("id")}
    for item in converted:
        reply_to = item.get("replyTo")
        if not isinstance(reply_to, dict):
            continue
        parent = message_map.get(str(reply_to.get("id") or ""))
        if not parent:
            continue
        reply_to["author"] = str(parent.get("author") or "")
        reply_to["timestamp"] = str(parent.get("timestamp") or "")
        reply_to["text"] = str(parent.get("text") or "")
    converted.sort(key=lambda item: (str(item.get("timestamp") or ""), str(item.get("id") or "")))
    return converted


def convert_one_graph_message(
    message: dict[str, Any],
    *,
    conversation_id: str | None = None,
) -> dict[str, Any] | None:
    body = _ensure_dict(message.get("body"))
    raw_content = str(body.get("content") or "")
    content_type = str(body.get("contentType") or "")
    message_type = _optional_str(message.get("messageType")) or "message"
    is_system = message_type.lower() != "message" or bool(message.get("eventDetail"))
    text = html_to_text(raw_content) if content_type.lower() == "html" or raw_content.lstrip().startswith("<") else raw_content
    if is_system and not text:
        text = _graph_system_text(message)

    reply_to_id = _optional_str(message.get("replyToId"))
    return {
        "id": str(message.get("id") or ""),
        "threadId": str(_channel_id(message) or conversation_id or ""),
        "author": _resolve_graph_author(message, is_system=is_system),
        "timestamp": str(message.get("createdDateTime") or message.get("lastModifiedDateTime") or ""),
        "text": text,
        "contentHtml": raw_content or None,
        "messageType": message_type,
        "edited": bool(message.get("lastEditedDateTime")),
        "system": is_system,
        "importance": _optional_str(message.get("importance")),
        "subject": _optional_str(message.get("subject")),
        "reactions": _convert_graph_reactions(message.get("reactions")),
        "attachments": _convert_graph_attachments(message.get("attachments")),
        "replyTo": (
            {
                "author": "",
                "timestamp": "",
                "text": "",
                "id": reply_to_id,
            }
            if reply_to_id
            else None
        ),
        "mentions": _convert_graph_mentions(message.get("mentions")),
    }


def _flatten_graph_message(message: dict[str, Any]) -> list[dict[str, Any]]:
    flattened = [message]
    replies = message.get("replies")
    if isinstance(replies, list):
        flattened.extend(reply for reply in replies if isinstance(reply, dict))
    return flattened


def _resolve_graph_author(message: dict[str, Any], *, is_system: bool) -> str:
    if is_system:
        return "[system]"
    source = _ensure_dict(message.get("from"))
    user = _ensure_dict(source.get("user"))
    application = _ensure_dict(source.get("application"))
    device = _ensure_dict(source.get("device"))
    conversation = _ensure_dict(source.get("conversation"))
    for candidate in [
        user.get("displayName"),
        application.get("displayName"),
        device.get("displayName"),
        conversation.get("displayName"),
    ]:
        text = _optional_str(candidate)
        if text:
            return text
    return ""


def _convert_graph_reactions(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    grouped: dict[str, list[str]] = defaultdict(list)
    for item in value:
        if not isinstance(item, dict):
            continue
        key = _optional_str(item.get("reactionType")) or ":reaction:"
        user = _ensure_dict(item.get("user"))
        nested_user = _ensure_dict(user.get("user"))
        display_name = _optional_str(nested_user.get("displayName")) or _optional_str(user.get("displayName"))
        if display_name:
            grouped[key].append(display_name)
        else:
            grouped[key]
    return [
        {
            "emoji": reaction,
            "count": len(users) if users else 1,
            "reactors": users or None,
        }
        for reaction, users in grouped.items()
    ]


def _convert_graph_attachments(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    attachments: list[dict[str, Any]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        href = _optional_str(item.get("contentUrl"))
        label = _optional_str(item.get("name")) or "Attachment"
        type_value = _optional_str(item.get("contentType"))
        if not keep_attachment(label=label, href=href, type_value=type_value):
            continue
        attachments.append(
            {
                "href": href,
                "label": label,
                "type": type_value,
                "size": None,
                "owner": None,
                "metaText": _optional_str(item.get("id")),
            }
        )
    return attachments


def _convert_graph_mentions(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    mentions: list[dict[str, Any]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        mentioned = _ensure_dict(item.get("mentioned"))
        user = _ensure_dict(mentioned.get("user"))
        app = _ensure_dict(mentioned.get("application"))
        name = (
            _optional_str(user.get("displayName"))
            or _optional_str(app.get("displayName"))
            or _optional_str(item.get("mentionText"))
        )
        if not name:
            continue
        mentions.append(
            {
                "name": name,
                "mri": _optional_str(user.get("id")) or _optional_str(app.get("id")),
            }
        )
    return mentions


def _graph_system_text(message: dict[str, Any]) -> str:
    event_detail = _ensure_dict(message.get("eventDetail"))
    event_type = _optional_str(event_detail.get("@odata.type")) or _optional_str(message.get("messageType"))
    if event_type:
        return event_type.rsplit(".", 1)[-1]
    return "system event"


def _channel_id(message: dict[str, Any]) -> str | None:
    identity = _ensure_dict(message.get("channelIdentity"))
    return _optional_str(identity.get("channelId"))


def _ensure_dict(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    return {}


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text else None
