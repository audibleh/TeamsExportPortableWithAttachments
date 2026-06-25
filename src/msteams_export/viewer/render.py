from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse

from msteams_export.models import Attachment, ExportBundle, ExportMessage


RESET = "\033[0m"
NEON_GREEN = "\033[92m"
NEON_CYAN = "\033[96m"
NEON_MAGENTA = "\033[95m"
NEON_YELLOW = "\033[93m"
NEON_RED = "\033[91m"
DIM = "\033[2m"
IMAGE_TYPES = {"png", "jpg", "jpeg", "gif", "bmp", "webp", "svg", "heic", "heif"}


@dataclass(slots=True)
class ViewOptions:
    limit: int = 10
    author: str | None = None
    query: str | None = None
    hide_system: bool = False


def render_summary(bundle: ExportBundle) -> str:
    image_attachments = sum(
        1 for message in bundle.messages for attachment in message.attachments if _attachment_is_image(attachment)
    )
    lines = [
        _banner("MS Teams Export Inspector"),
        f"{NEON_CYAN}title{RESET}: {bundle.meta.title}",
        f"{NEON_CYAN}messages{RESET}: {bundle.message_count}",
        f"{NEON_CYAN}authors{RESET}: {len(bundle.authors)}",
        f"{NEON_CYAN}attachments{RESET}: {bundle.attachment_count}",
        f"{NEON_CYAN}image attachments{RESET}: {image_attachments}",
        f"{NEON_CYAN}reactions{RESET}: {bundle.reaction_count}",
    ]
    if bundle.meta.export_target:
        lines.append(f"{NEON_CYAN}target{RESET}: {bundle.meta.export_target}")
    if bundle.meta.conversation_id:
        lines.append(f"{NEON_CYAN}conversation{RESET}: {bundle.meta.conversation_id}")
    if bundle.meta.hidden or bundle.meta.meeting:
        flags = []
        if bundle.meta.hidden:
            flags.append("hidden")
        if bundle.meta.meeting:
            flags.append("meeting")
        lines.append(f"{NEON_CYAN}flags{RESET}: {', '.join(flags)}")
    if bundle.meta.discovery_sources:
        lines.append(f"{NEON_CYAN}discovery{RESET}: {', '.join(bundle.meta.discovery_sources)}")
    if bundle.meta.start_at or bundle.meta.end_at:
        lines.append(
            f"{NEON_CYAN}range{RESET}: {bundle.meta.start_at or '?'} -> {bundle.meta.end_at or '?'}"
        )
    if bundle.authors:
        preview = ", ".join(bundle.authors[:6])
        lines.append(f"{NEON_CYAN}author preview{RESET}: {preview}")
    return "\n".join(lines)


def render_messages(bundle: ExportBundle, options: ViewOptions) -> str:
    filtered = [message for message in bundle.messages if _matches(message, options)]
    limited = filtered[: options.limit]
    lines = [_banner("Neon Chat View")]
    lines.append(
        f"{DIM}showing {len(limited)} of {len(filtered)} matching messages{RESET}"
    )
    for index, message in enumerate(limited, start=1):
        lines.extend(_render_message(index, message))
    if not limited:
        lines.append(f"{NEON_YELLOW}No messages matched the current filters.{RESET}")
    return "\n".join(lines)


def _matches(message: ExportMessage, options: ViewOptions) -> bool:
    if options.hide_system and message.system:
        return False
    if options.author and message.author.lower() != options.author.lower():
        return False
    if options.query and options.query.lower() not in message.text.lower():
        return False
    return True


def _render_message(index: int, message: ExportMessage) -> list[str]:
    reaction_preview = ""
    if message.reactions:
        reaction_preview = " ".join(f"{reaction.emoji}x{reaction.count}" for reaction in message.reactions)
    attachment_preview = ""
    if message.attachments:
        attachment_preview = ", ".join(
            attachment.label or attachment.href or "attachment"
            for attachment in message.attachments[:3]
        )
    image_attachments = [attachment for attachment in message.attachments if _attachment_is_image(attachment)]
    file_attachments = [attachment for attachment in message.attachments if not _attachment_is_image(attachment)]
    header = f"{NEON_MAGENTA}[{index:03d}]{RESET} {NEON_GREEN}{message.author or '[unknown]'}{RESET}"
    if message.system:
        header = (
            f"{NEON_MAGENTA}[{index:03d}]{RESET} "
            f"{NEON_YELLOW}[system:{_system_tag(message)}]{RESET}"
        )
    body = _compact_text(message.text or "[no text]")
    lines = [
        f"",
        header,
        f"  {DIM}{message.timestamp or 'unknown time'}{RESET}",
        f"  {body}",
    ]
    if message.edited:
        lines.append(f"  {DIM}(edited){RESET}")
    if message.reply_to and message.reply_to.text:
        lines.append(
            f"  {NEON_YELLOW}reply-to{RESET}: {message.reply_to.author or '[unknown]'} :: {_compact_text(message.reply_to.text, 120)}"
        )
    if reaction_preview:
        lines.append(f"  {NEON_CYAN}reactions{RESET}: {reaction_preview}")
    if image_attachments:
        labels = ", ".join(_attachment_label(attachment) for attachment in image_attachments[:3])
        lines.append(f"  {NEON_CYAN}images{RESET}: {labels}")
    if file_attachments:
        labels = ", ".join(_attachment_label(attachment) for attachment in file_attachments[:3])
        lines.append(f"  {NEON_CYAN}files{RESET}: {labels}")
    elif attachment_preview:
        lines.append(f"  {NEON_CYAN}attachments{RESET}: {attachment_preview}")
    if message.mentions:
        mention_preview = ", ".join(message.mentions[:5])
        lines.append(f"  {NEON_GREEN}mentions{RESET}: {mention_preview}")
    return lines


def _banner(title: str) -> str:
    bar = "=" * len(title)
    return f"{NEON_MAGENTA}{bar}\n{title}\n{bar}{RESET}"


def _system_tag(message: ExportMessage) -> str:
    raw = (message.message_type or "system").split("/")[-1]
    return raw.replace("_", "-").lower()


def _compact_text(value: str, max_length: int = 220) -> str:
    normalized = " ".join(value.split())
    if len(normalized) <= max_length:
        return normalized
    return f"{normalized[: max_length - 3].rstrip()}..."


def _attachment_is_image(attachment: Attachment) -> bool:
    type_value = (attachment.type or "").strip().lower().lstrip(".")
    if type_value in IMAGE_TYPES:
        return True
    for candidate in [attachment.label, attachment.href]:
        suffix = _suffix(candidate)
        if suffix in IMAGE_TYPES:
            return True
    return False


def _attachment_label(attachment: Attachment) -> str:
    label = attachment.label or attachment.href or "attachment"
    if _attachment_is_image(attachment):
        return f"[img] {label}"
    return label


def _suffix(value: str | None) -> str:
    if not value:
        return ""
    parsed = urlparse(value)
    path = parsed.path or value
    suffix = Path(path).suffix.lower().lstrip(".")
    return suffix
