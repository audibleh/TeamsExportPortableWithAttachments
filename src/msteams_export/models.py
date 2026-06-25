from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
from typing import Any

from msteams_export.attachment_policy import keep_attachment
from msteams_export.parsing.teams_api import merge_embedded_html_attachments


@dataclass(slots=True)
class Reaction:
    emoji: str
    count: int
    reactors: list[str] = field(default_factory=list)


@dataclass(slots=True)
class Attachment:
    href: str | None = None
    label: str | None = None
    type: str | None = None
    size: str | None = None
    owner: str | None = None
    meta_text: str | None = None
    kind: str | None = None


@dataclass(slots=True)
class ReplyContext:
    author: str = ""
    timestamp: str = ""
    text: str = ""
    message_id: str | None = None


@dataclass(slots=True)
class ForwardContext:
    original_author: str | None = None
    original_timestamp: str | None = None
    original_message_id: str | None = None
    original_thread_id: str | None = None
    original_text: str | None = None


@dataclass(slots=True)
class ExportMessage:
    identifier: str = ""
    thread_id: str | None = None
    author: str = ""
    timestamp: str = ""
    text: str = ""
    content_html: str | None = None
    message_type: str | None = None
    edited: bool = False
    system: bool = False
    importance: str | None = None
    subject: str | None = None
    reactions: list[Reaction] = field(default_factory=list)
    attachments: list[Attachment] = field(default_factory=list)
    reply_to: ReplyContext | None = None
    forwarded: ForwardContext | None = None
    mentions: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ExportMessage":
        reactions = [
            Reaction(
                emoji=str(item.get("emoji", "")),
                count=int(item.get("count", 0)),
                reactors=[str(value) for value in item.get("reactors", []) or []],
            )
            for item in data.get("reactions", []) or []
        ]
        raw_attachments = merge_embedded_html_attachments(data.get("attachments"), _optional_str(data.get("contentHtml")))
        attachments = [
            Attachment(
                href=_optional_str(item.get("href")),
                label=_optional_str(item.get("label")),
                type=_optional_str(item.get("type")),
                size=_optional_str(item.get("size")),
                owner=_optional_str(item.get("owner")),
                meta_text=_optional_str(item.get("metaText")),
                kind=_optional_str(item.get("kind")),
            )
            for item in raw_attachments
            if isinstance(item, dict)
            and keep_attachment(
                label=_optional_str(item.get("label")),
                href=_optional_str(item.get("href")),
                type_value=_optional_str(item.get("type")),
                kind=_optional_str(item.get("kind")),
            )
        ]
        reply_raw = data.get("replyTo")
        forward_raw = data.get("forwarded")
        mentions = [
            str(item.get("name", ""))
            for item in data.get("mentions", []) or []
            if isinstance(item, dict) and item.get("name")
        ]
        return cls(
            identifier=str(data.get("id", "") or ""),
            thread_id=_optional_str(data.get("threadId")),
            author=str(data.get("author", "") or ""),
            timestamp=str(data.get("timestamp", "") or ""),
            text=str(data.get("text", "") or ""),
            content_html=_optional_str(data.get("contentHtml")),
            message_type=_optional_str(data.get("messageType")),
            edited=bool(data.get("edited", False)),
            system=bool(data.get("system", False)),
            importance=_optional_str(data.get("importance")),
            subject=_optional_str(data.get("subject")),
            reactions=reactions,
            attachments=attachments,
            reply_to=_reply_from_dict(reply_raw),
            forwarded=_forward_from_dict(forward_raw),
            mentions=mentions,
        )


@dataclass(slots=True)
class ExportMeta:
    title: str = "Teams Export"
    count: int = 0
    start_at: str | None = None
    end_at: str | None = None
    time_range: str | None = None
    conversation_id: str | None = None
    export_target: str | None = None
    source: str | None = None
    user_region: str | None = None
    hidden: bool = False
    meeting: bool = False
    discovery_sources: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ExportMeta":
        conversation = data.get("conversation")
        conversation_dict = conversation if isinstance(conversation, dict) else {}
        discovery_sources = [
            str(item)
            for item in conversation_dict.get("discoverySources", []) or []
            if item is not None and str(item)
        ]
        return cls(
            title=str(data.get("title", "Teams Export") or "Teams Export"),
            count=int(data.get("count", 0) or 0),
            start_at=_optional_str(data.get("startAt")),
            end_at=_optional_str(data.get("endAt")),
            time_range=_optional_str(data.get("timeRange")),
            conversation_id=_optional_str(data.get("conversationId")),
            export_target=_optional_str(data.get("exportTarget")),
            source=_optional_str(data.get("source")),
            user_region=_optional_str(data.get("userRegion")),
            hidden=bool(conversation_dict.get("hidden", False)),
            meeting=bool(conversation_dict.get("meeting", False)),
            discovery_sources=discovery_sources,
        )


@dataclass(slots=True)
class ExportBundle:
    meta: ExportMeta
    messages: list[ExportMessage]

    @classmethod
    def load(cls, path: Path) -> "ExportBundle":
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            msg = f"Expected top-level JSON object in {path}"
            raise ValueError(msg)
        meta_raw = payload.get("meta", {})
        messages_raw = payload.get("messages", [])
        if not isinstance(meta_raw, dict):
            raise ValueError("Expected 'meta' to be an object")
        if not isinstance(messages_raw, list):
            raise ValueError("Expected 'messages' to be an array")
        messages = [ExportMessage.from_dict(item) for item in messages_raw if isinstance(item, dict)]
        meta = ExportMeta.from_dict(meta_raw)
        if meta.count == 0:
            meta.count = len(messages)
        return cls(meta=meta, messages=messages)

    @property
    def message_count(self) -> int:
        return len(self.messages)

    @property
    def authors(self) -> list[str]:
        values = {message.author for message in self.messages if message.author}
        return sorted(values)

    @property
    def attachment_count(self) -> int:
        return sum(len(message.attachments) for message in self.messages)

    @property
    def reaction_count(self) -> int:
        return sum(len(message.reactions) for message in self.messages)


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text else None


def _reply_from_dict(value: Any) -> ReplyContext | None:
    if not isinstance(value, dict):
        return None
    return ReplyContext(
        author=str(value.get("author", "") or ""),
        timestamp=str(value.get("timestamp", "") or ""),
        text=str(value.get("text", "") or ""),
        message_id=_optional_str(value.get("id")),
    )


def _forward_from_dict(value: Any) -> ForwardContext | None:
    if not isinstance(value, dict):
        return None
    return ForwardContext(
        original_author=_optional_str(value.get("originalAuthor")),
        original_timestamp=_optional_str(value.get("originalTimestamp")),
        original_message_id=_optional_str(value.get("originalMessageId")),
        original_thread_id=_optional_str(value.get("originalThreadId")),
        original_text=_optional_str(value.get("originalText")),
    )
