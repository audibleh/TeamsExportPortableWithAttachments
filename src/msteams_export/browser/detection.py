from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import platform
from typing import Callable


@dataclass(slots=True, frozen=True)
class BrowserInstall:
    name: str
    executable_path: Path
    source: str


def discover_browsers(
    system_name: str | None = None,
    path_exists: Callable[[Path], bool] | None = None,
) -> list[BrowserInstall]:
    system = (system_name or platform.system()).lower()
    exists = path_exists or _default_path_exists
    installs: list[BrowserInstall] = []
    for install in _candidate_installs(system):
        if exists(install.executable_path):
            installs.append(install)
    return installs


def choose_browser(browser_name: str, installs: list[BrowserInstall]) -> BrowserInstall | None:
    if not installs:
        return None
    normalized = browser_name.lower()
    if normalized == "auto":
        for preferred in ("edge", "chrome", "chromium"):
            selected = next((item for item in installs if item.name == preferred), None)
            if selected is not None:
                return selected
        return installs[0]
    return next((item for item in installs if item.name == normalized), None)


def _default_path_exists(path: Path) -> bool:
    return path.is_file()


def _candidate_installs(system: str) -> list[BrowserInstall]:
    if system == "darwin":
        return [
            BrowserInstall(
                name="edge",
                executable_path=Path("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
                source="system",
            ),
            BrowserInstall(
                name="chrome",
                executable_path=Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
                source="system",
            ),
            BrowserInstall(
                name="chromium",
                executable_path=Path("/Applications/Chromium.app/Contents/MacOS/Chromium"),
                source="system",
            ),
        ]
    if system == "windows":
        return [
            BrowserInstall(
                name="edge",
                executable_path=Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
                source="system",
            ),
            BrowserInstall(
                name="chrome",
                executable_path=Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
                source="system",
            ),
            BrowserInstall(
                name="chromium",
                executable_path=Path(r"C:\Program Files\Chromium\Application\chrome.exe"),
                source="system",
            ),
        ]
    return [
        BrowserInstall(
            name="edge",
            executable_path=Path("/usr/bin/microsoft-edge"),
            source="system",
        ),
        BrowserInstall(
            name="chrome",
            executable_path=Path("/usr/bin/google-chrome"),
            source="system",
        ),
        BrowserInstall(
            name="chromium",
            executable_path=Path("/usr/bin/chromium"),
            source="system",
        ),
    ]

