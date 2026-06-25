from __future__ import annotations

from dataclasses import dataclass
import random
import time
from typing import Any


@dataclass(frozen=True, slots=True)
class PoliteModePolicy:
    enabled: bool = True
    request_spacing_ms: int = 650
    page_spacing_ms: int = 450
    conversation_spacing_ms: int = 1_200
    attachment_spacing_ms: int = 900
    retry_limit: int = 8
    retry_base_ms: int = 2_500
    retry_max_ms: int = 60_000
    jitter_ms: int = 350
    retry_statuses: tuple[int, ...] = (429, 503)


DEFAULT_POLITE_MODE = PoliteModePolicy()


def apply_spacing(last_request_at: float | None, spacing_ms: int) -> float:
    now = time.monotonic()
    if last_request_at is None or spacing_ms <= 0:
        return now
    elapsed_ms = (now - last_request_at) * 1000
    remaining_ms = spacing_ms - elapsed_ms
    if remaining_ms > 0:
        time.sleep(remaining_ms / 1000)
        now = time.monotonic()
    return now


def compute_retry_delay_ms(
    *,
    attempt: int,
    retry_after_header: str | None,
    policy: PoliteModePolicy = DEFAULT_POLITE_MODE,
) -> int:
    retry_after_ms = _parse_retry_after_ms(retry_after_header)
    if retry_after_ms is not None:
        return min(max(retry_after_ms, 0), policy.retry_max_ms)
    exponential_ms = min(policy.retry_base_ms * (2**max(0, attempt)), policy.retry_max_ms)
    jitter_ms = random.randint(0, max(0, policy.jitter_ms))
    return min(exponential_ms + jitter_ms, policy.retry_max_ms)


def build_browser_polite_mode_payload(policy: PoliteModePolicy = DEFAULT_POLITE_MODE) -> dict[str, Any]:
    return {
        "enabled": policy.enabled,
        "requestSpacingMs": policy.request_spacing_ms,
        "pageSpacingMs": policy.page_spacing_ms,
        "conversationSpacingMs": policy.conversation_spacing_ms,
        "retryLimit": policy.retry_limit,
        "retryBaseMs": policy.retry_base_ms,
        "retryMaxMs": policy.retry_max_ms,
        "jitterMs": policy.jitter_ms,
        "retryStatuses": list(policy.retry_statuses),
    }


def _parse_retry_after_ms(value: str | None) -> int | None:
    if not value:
        return None
    text = value.strip()
    if not text:
        return None
    try:
        seconds = float(text)
    except ValueError:
        return None
    return int(max(0.0, seconds) * 1000)
