from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
import hashlib
import json
from pathlib import Path
import re
import shutil
import time
from typing import Any, Callable
from urllib.parse import quote
import unicodedata

from msteams_export.attachment_policy import keep_attachment
from msteams_export.browser.session import DEFAULT_TEAMS_URL
from msteams_export.bundle_paths import bundle_relative_path_string, resolve_bundle_relative_path
from msteams_export.export_bundle import resolve_export_root
from msteams_export.parsing.teams_api import merge_embedded_html_attachments
from msteams_export.webapp.attachments import AttachmentUnauthorizedError, open_attachment_fetch_session


DEFAULT_MIN_FREE_BYTES = 30 * 1024 * 1024 * 1024


@dataclass(slots=True)
class MirrorAttachmentsRequest:
    target: Path
    browser_name: str = "auto"
    profile_path: Path | None = None
    teams_url: str = DEFAULT_TEAMS_URL
    timeout_ms: int = 30_000
    max_assets: int | None = None
    min_free_bytes: int = DEFAULT_MIN_FREE_BYTES
    progress: Callable[["AttachmentMirrorProgress"], None] | None = None
    stop_controller: "MirrorStopController | None" = None


@dataclass(slots=True)
class MirrorStopController:
    interrupt_count: int = 0

    def request_interrupt(self) -> int:
        self.interrupt_count += 1
        return self.interrupt_count

    @property
    def stop_after_current_chat(self) -> bool:
        return self.interrupt_count >= 1

    @property
    def stop_after_current_asset(self) -> bool:
        return self.interrupt_count >= 2


@dataclass(slots=True)
class AttachmentMirrorProgress:
    phase: str
    total_chats: int = 0
    total_assets: int = 0
    processed_chats: int = 0
    processed_assets: int = 0
    mirrored_assets: int = 0
    reused_assets: int = 0
    failed_assets: int = 0
    current_chat_title: str | None = None
    current_asset_label: str | None = None
    current_status: str | None = None
    elapsed_seconds: float = 0.0
    eta_seconds: float | None = None
    free_bytes: int | None = None
    min_free_bytes: int = DEFAULT_MIN_FREE_BYTES
    bytes_downloaded: int = 0
    note: str | None = None


@dataclass(slots=True)
class MirrorAttachmentsResult:
    ok: bool
    message: str
    bundle_root: Path | None = None
    assets_dir: Path | None = None
    chat_count: int = 0
    attachment_count: int = 0
    mirrored_count: int = 0
    reused_count: int = 0
    failed_count: int = 0
    interrupted: bool = False
    low_disk: bool = False
    free_bytes: int | None = None
    min_free_bytes: int = DEFAULT_MIN_FREE_BYTES


def mirror_bundle_attachments(request: MirrorAttachmentsRequest) -> MirrorAttachmentsResult:
    try:
        index_path, _ = resolve_export_root(request.target)
    except Exception as exc:
        return MirrorAttachmentsResult(ok=False, message=str(exc))

    bundle_root = index_path.parent
    assets_dir = bundle_root / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    min_free_bytes = max(0, int(request.min_free_bytes))

    try:
        index_payload = json.loads(index_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return MirrorAttachmentsResult(ok=False, message=f"Could not read export index: {exc}", bundle_root=bundle_root)

    if not isinstance(index_payload, dict):
        return MirrorAttachmentsResult(ok=False, message="Expected JSON object in index.json", bundle_root=bundle_root)

    conversations = index_payload.get("conversations", [])
    if not isinstance(conversations, list):
        return MirrorAttachmentsResult(
            ok=False,
            message="Expected 'conversations' array in index.json",
            bundle_root=bundle_root,
        )

    chat_records = [item for item in conversations if isinstance(item, dict) and isinstance(item.get("exportPath"), str)]
    total_assets = 0
    scanned_chats = 0
    for record in chat_records:
        chat_path = _chat_path(bundle_root, record)
        if not chat_path.is_file():
            continue
        scanned_chats += 1
        total_assets += _count_candidate_attachments(chat_path)
        if request.max_assets is not None and total_assets >= request.max_assets:
            total_assets = request.max_assets
            break

    processed_chats = 0
    processed_assets = 0
    mirrored_count = 0
    reused_count = 0
    failed_count = 0
    bytes_downloaded = 0
    interrupted = False
    low_disk = False
    interruption_note: str | None = None
    start_time = time.monotonic()

    initial_free_bytes = _free_disk_bytes(bundle_root)
    if initial_free_bytes < min_free_bytes:
        message = (
            "Not enough free disk space to start attachment mirroring. "
            f"Free { _format_bytes(initial_free_bytes) }, required minimum { _format_bytes(min_free_bytes) }."
        )
        return MirrorAttachmentsResult(
            ok=False,
            message=message,
            bundle_root=bundle_root,
            assets_dir=assets_dir,
            free_bytes=initial_free_bytes,
            min_free_bytes=min_free_bytes,
            low_disk=True,
        )

    def emit(
        *,
        phase: str,
        current_chat_title: str | None = None,
        current_asset_label: str | None = None,
        current_status: str | None = None,
        note: str | None = None,
    ) -> None:
        if request.progress is None:
            return
        elapsed_seconds = max(0.0, time.monotonic() - start_time)
        eta_seconds = _estimate_eta_seconds(
            processed=processed_assets,
            total=total_assets,
            elapsed_seconds=elapsed_seconds,
        )
        request.progress(
            AttachmentMirrorProgress(
                phase=phase,
                total_chats=scanned_chats,
                total_assets=total_assets,
                processed_chats=processed_chats,
                processed_assets=processed_assets,
                mirrored_assets=mirrored_count,
                reused_assets=reused_count,
                failed_assets=failed_count,
                current_chat_title=current_chat_title,
                current_asset_label=current_asset_label,
                current_status=current_status,
                elapsed_seconds=elapsed_seconds,
                eta_seconds=eta_seconds,
                free_bytes=_free_disk_bytes(bundle_root),
                min_free_bytes=min_free_bytes,
                bytes_downloaded=bytes_downloaded,
                note=note,
            )
        )

    emit(
        phase="discovering",
        note=(
            f"Found {scanned_chats} chats and {total_assets} candidate attachment(s). "
            f"Resume is safe: already mirrored assets will be reused."
        ),
    )

    with open_attachment_fetch_session(
        browser_name=request.browser_name,
        profile_path=request.profile_path,
        teams_url=request.teams_url,
        timeout_ms=request.timeout_ms,
    ) as fetch_session:
        for record in chat_records:
            if request.max_assets is not None and processed_assets >= request.max_assets:
                break
            if request.stop_controller is not None and request.stop_controller.stop_after_current_chat:
                interrupted = True
                interruption_note = "Quit requested. Stopped before starting the next chat."
                emit(
                    phase="interrupting",
                    current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                    current_status="interrupting",
                    note=interruption_note,
                )
                break
            chat_path = _chat_path(bundle_root, record)
            if not chat_path.is_file():
                continue
            try:
                payload = json.loads(chat_path.read_text(encoding="utf-8"))
            except Exception:
                processed_chats += 1
                continue
            if not isinstance(payload, dict):
                processed_chats += 1
                continue

            changed = False
            stats = {"assetCount": _count_candidate_attachments_in_payload(payload), "mirrored": 0, "failed": 0, "unauthorized": 0}
            messages = payload.get("messages", [])
            if not isinstance(messages, list):
                messages = []
                payload["messages"] = messages

            for message in messages:
                if request.max_assets is not None and processed_assets >= request.max_assets:
                    break
                if not isinstance(message, dict):
                    continue
                merged_attachments = merge_embedded_html_attachments(
                    message.get("attachments"),
                    _optional_str(message.get("contentHtml")),
                )
                if merged_attachments != message.get("attachments"):
                    message["attachments"] = merged_attachments
                    changed = True
                attachments = message.get("attachments", [])
                if not isinstance(attachments, list):
                    continue
                message_id = _optional_str(message.get("id")) or "message"
                for index, attachment in enumerate(attachments):
                    if request.max_assets is not None and processed_assets >= request.max_assets:
                        break
                    if not isinstance(attachment, dict):
                        continue
                    label = _optional_str(attachment.get("label")) or "attachment"
                    href = _optional_str(attachment.get("href"))
                    type_value = _optional_str(attachment.get("type"))
                    kind = _optional_str(attachment.get("kind"))
                    if not keep_attachment(label=label, href=href, type_value=type_value, kind=kind):
                        continue

                    free_bytes = _free_disk_bytes(bundle_root)
                    if free_bytes < min_free_bytes:
                        interrupted = True
                        low_disk = True
                        interruption_note = (
                            "Stopped attachment mirroring to preserve free disk space. "
                            f"Free { _format_bytes(free_bytes) }, minimum { _format_bytes(min_free_bytes) }."
                        )
                        emit(
                            phase="paused-low-disk",
                            current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                            current_asset_label=label,
                            current_status="paused-low-disk",
                            note=interruption_note,
                        )
                        break

                    processed_assets += 1

                    local_path = _optional_str(attachment.get("localPath"))
                    existing_path = _resolve_local_asset_path(bundle_root, local_path) if local_path else None
                    if existing_path is not None and existing_path.is_file():
                        _apply_local_asset_metadata(
                            attachment,
                            relative_path=_bundle_relative_path(bundle_root, existing_path),
                            size=existing_path.stat().st_size,
                            content_type=_optional_str(attachment.get("localContentType")),
                        )
                        reused_count += 1
                        stats["mirrored"] += 1
                        emit(
                            phase="mirroring",
                            current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                            current_asset_label=label,
                            current_status="reused",
                        )
                        continue

                    if not href:
                        attachment["localStatus"] = "failed"
                        attachment["localError"] = "Attachment does not have a source URL."
                        failed_count += 1
                        stats["failed"] += 1
                        changed = True
                        emit(
                            phase="mirroring",
                            current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                            current_asset_label=label,
                            current_status="failed",
                            note="Missing source URL.",
                        )
                        continue

                    try:
                        downloaded = fetch_session.fetch(url=href, label=label)
                        relative_path = _asset_relative_path(
                            conversation_id=_optional_str(record.get("id")) or "conversation",
                            message_id=message_id,
                            index=index,
                            label=downloaded.filename or label,
                            href=href,
                        )
                        output_path = bundle_root / relative_path
                        output_path.parent.mkdir(parents=True, exist_ok=True)
                        output_path.write_bytes(downloaded.body)
                        _apply_local_asset_metadata(
                            attachment,
                            relative_path=relative_path.as_posix(),
                            size=len(downloaded.body),
                            content_type=downloaded.content_type,
                        )
                        bytes_downloaded += len(downloaded.body)
                        mirrored_count += 1
                        stats["mirrored"] += 1
                        changed = True
                        emit(
                            phase="mirroring",
                            current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                            current_asset_label=label,
                            current_status="mirrored",
                        )
                    except AttachmentUnauthorizedError as exc:
                        attachment["localStatus"] = "unauthorized"
                        attachment["localError"] = str(exc)
                        failed_count += 1
                        stats["failed"] += 1
                        stats["unauthorized"] += 1
                        changed = True
                        emit(
                            phase="mirroring",
                            current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                            current_asset_label=label,
                            current_status="unauthorized",
                            note=str(exc),
                        )
                    except Exception as exc:
                        attachment["localStatus"] = "failed"
                        attachment["localError"] = str(exc)
                        failed_count += 1
                        stats["failed"] += 1
                        changed = True
                        emit(
                            phase="mirroring",
                            current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                            current_asset_label=label,
                            current_status="failed",
                            note=str(exc),
                        )

                    if request.stop_controller is not None and request.stop_controller.stop_after_current_asset:
                        interrupted = True
                        interruption_note = (
                            "Force-quit requested. Stopped after the current attachment and wrote partial mirror state."
                        )
                        emit(
                            phase="interrupting",
                            current_chat_title=_optional_str(record.get("title")) or _optional_str(record.get("id")),
                            current_asset_label=label,
                            current_status="interrupting",
                            note=interruption_note,
                        )
                        break

                if interrupted and (low_disk or (request.stop_controller and request.stop_controller.stop_after_current_asset)):
                    break

            record["assetCount"] = stats["assetCount"]
            record["mirroredAssetCount"] = stats["mirrored"]
            record["assetFailureCount"] = stats["failed"]
            record["unauthorizedAssetCount"] = stats["unauthorized"]
            record["offlineReady"] = stats["assetCount"] == 0 or stats["mirrored"] >= stats["assetCount"]

            if changed:
                chat_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            processed_chats += 1
            if interrupted:
                break

    _update_index_asset_summary(index_payload, conversations)
    index_path.write_text(json.dumps(index_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    final_note = interruption_note or "Attachment mirror metadata written."
    emit(phase="finished", note=final_note)

    free_bytes = _free_disk_bytes(bundle_root)
    status_prefix = "Partially mirrored" if interrupted or low_disk else "Mirrored"
    message = f"{status_prefix} {mirrored_count} attachment(s), reused {reused_count}, failed {failed_count}, across {processed_chats} chat(s)."
    if interrupted or low_disk:
        message += " You can resume later by rerunning the same mirror command."
    if interruption_note:
        message += f" {interruption_note}"
    return MirrorAttachmentsResult(
        ok=not interrupted and not low_disk,
        message=message,
        bundle_root=bundle_root,
        assets_dir=assets_dir,
        chat_count=processed_chats,
        attachment_count=processed_assets,
        mirrored_count=mirrored_count,
        reused_count=reused_count,
        failed_count=failed_count,
        interrupted=interrupted,
        low_disk=low_disk,
        free_bytes=free_bytes,
        min_free_bytes=min_free_bytes,
    )


def _chat_path(bundle_root: Path, record: dict[str, Any]) -> Path:
    candidate = resolve_bundle_relative_path(bundle_root, _optional_str(record.get("exportPath")))
    return candidate.resolve() if candidate is not None else bundle_root / "__missing__"


def _count_candidate_attachments(chat_path: Path) -> int:
    try:
        payload = json.loads(chat_path.read_text(encoding="utf-8"))
    except Exception:
        return 0
    return _count_candidate_attachments_in_payload(payload)


def _count_candidate_attachments_in_payload(payload: dict[str, Any]) -> int:
    if not isinstance(payload, dict):
        return 0
    total = 0
    messages = payload.get("messages", [])
    if not isinstance(messages, list):
        return 0
    for message in messages:
        if not isinstance(message, dict):
            continue
        attachments = merge_embedded_html_attachments(message.get("attachments"), _optional_str(message.get("contentHtml")))
        for attachment in attachments:
            if not isinstance(attachment, dict):
                continue
            if keep_attachment(
                label=_optional_str(attachment.get("label")),
                href=_optional_str(attachment.get("href")),
                type_value=_optional_str(attachment.get("type")),
                kind=_optional_str(attachment.get("kind")),
            ):
                total += 1
    return total


def _asset_relative_path(
    *,
    conversation_id: str,
    message_id: str,
    index: int,
    label: str,
    href: str,
) -> Path:
    conversation_dir = quote(conversation_id, safe="")
    message_dir = quote(message_id, safe="")
    filename = _asset_filename(label=label, href=href, index=index)
    return Path("assets") / conversation_dir / message_dir / filename


def _asset_filename(*, label: str, href: str, index: int) -> str:
    # Non-security digest used only to disambiguate generated asset filenames.
    digest = hashlib.sha1(href.encode("utf-8"), usedforsecurity=False).hexdigest()[:10]
    text = unicodedata.normalize("NFKD", label or "attachment")
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9._-]+", "-", text).strip("-") or "attachment"
    if "." in text:
        stem, ext = text.rsplit(".", 1)
        ext = "." + ext.lower()
    else:
        stem, ext = text, ""
    stem = stem[:60].strip("-._") or "attachment"
    return f"{index:03d}-{stem}-{digest}{ext}"


def _apply_local_asset_metadata(
    attachment: dict[str, Any],
    *,
    relative_path: str,
    size: int,
    content_type: str | None,
) -> None:
    attachment["localPath"] = relative_path
    attachment["localStatus"] = "mirrored"
    attachment["localSize"] = size
    if content_type:
        attachment["localContentType"] = content_type
    attachment["localMirroredAt"] = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    attachment.pop("localError", None)


def _resolve_local_asset_path(bundle_root: Path, local_path: str | None) -> Path | None:
    return resolve_bundle_relative_path(bundle_root, local_path)


def _bundle_relative_path(bundle_root: Path, path: Path) -> str:
    return bundle_relative_path_string(bundle_root, path)


def _update_index_asset_summary(index_payload: dict[str, Any], conversations: list[Any]) -> None:
    meta = index_payload.get("meta")
    if not isinstance(meta, dict):
        meta = {}
        index_payload["meta"] = meta
    records = [item for item in conversations if isinstance(item, dict)]
    meta["attachmentCount"] = sum(int(item.get("assetCount", 0) or 0) for item in records)
    meta["mirroredAttachmentCount"] = sum(int(item.get("mirroredAssetCount", 0) or 0) for item in records)
    meta["attachmentFailureCount"] = sum(int(item.get("assetFailureCount", 0) or 0) for item in records)
    meta["unauthorizedAttachmentCount"] = sum(int(item.get("unauthorizedAssetCount", 0) or 0) for item in records)
    meta["offlineReadyConversationCount"] = sum(1 for item in records if item.get("offlineReady"))
    meta["assetsMirroredAt"] = datetime.now(UTC).isoformat().replace("+00:00", "Z")


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def _free_disk_bytes(path: Path) -> int:
    return int(shutil.disk_usage(path).free)


def _estimate_eta_seconds(*, processed: int, total: int, elapsed_seconds: float) -> float | None:
    if processed <= 0 or total <= processed or elapsed_seconds <= 0:
        return None
    rate = processed / elapsed_seconds
    if rate <= 0:
        return None
    return max(0.0, (total - processed) / rate)


def _format_bytes(value: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(max(0, value))
    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)}{unit}"
            return f"{size:.1f}{unit}"
        size /= 1024
    return f"{size:.1f}TB"
