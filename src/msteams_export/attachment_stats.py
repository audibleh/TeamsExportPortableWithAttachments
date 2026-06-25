from __future__ import annotations

from typing import Any

from msteams_export.attachment_policy import keep_attachment
from msteams_export.parsing.teams_api import merge_embedded_html_attachments


def summarize_payload_attachments(payload: dict[str, Any]) -> dict[str, int | bool]:
    asset_count = 0
    mirrored_asset_count = 0
    failed_asset_count = 0
    unauthorized_asset_count = 0

    messages = payload.get("messages", [])
    if not isinstance(messages, list):
        return {
            "assetCount": 0,
            "mirroredAssetCount": 0,
            "assetFailureCount": 0,
            "unauthorizedAssetCount": 0,
            "offlineReady": True,
        }

    for message in messages:
        if not isinstance(message, dict):
            continue
        attachments = merge_embedded_html_attachments(message.get("attachments"), _optional_str(message.get("contentHtml")))
        for attachment in attachments:
            if not isinstance(attachment, dict):
                continue
            if not keep_attachment(
                label=_optional_str(attachment.get("label")),
                href=_optional_str(attachment.get("href")),
                type_value=_optional_str(attachment.get("type")),
                kind=_optional_str(attachment.get("kind")),
            ):
                continue
            asset_count += 1
            if _optional_str(attachment.get("localPath")):
                mirrored_asset_count += 1
            elif _optional_str(attachment.get("localStatus")) == "unauthorized":
                failed_asset_count += 1
                unauthorized_asset_count += 1
            elif _optional_str(attachment.get("localStatus")) == "failed":
                failed_asset_count += 1

    return {
        "assetCount": asset_count,
        "mirroredAssetCount": mirrored_asset_count,
        "assetFailureCount": failed_asset_count,
        "unauthorizedAssetCount": unauthorized_asset_count,
        "offlineReady": asset_count == 0 or mirrored_asset_count >= asset_count,
    }


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None
