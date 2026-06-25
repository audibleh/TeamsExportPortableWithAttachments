from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class ExportRequest:
    target: str
    output: Path


def plan_export(request: ExportRequest) -> str:
    return (
        f"Export target '{request.target}' is scaffolded, but real browser/API export "
        f"logic is not implemented yet. Planned output: {request.output}"
    )

