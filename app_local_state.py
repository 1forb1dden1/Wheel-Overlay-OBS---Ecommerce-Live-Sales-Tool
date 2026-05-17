"""Persist last wheel preset, paths, and winner log between app runs."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

STATE_VERSION = 1
STATE_FILENAME = "energy_break_state.json"


def state_path(base: Path) -> Path:
    return base.resolve() / STATE_FILENAME


def load_state(base: Path) -> dict[str, Any] | None:
    path = state_path(base)
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        return None
    if not isinstance(raw, dict):
        return None
    if int(raw.get("version") or 0) != STATE_VERSION:
        return None
    return raw


def save_state(base: Path, data: dict[str, Any]) -> bool:
    path = state_path(base)
    payload = dict(data)
    payload["version"] = STATE_VERSION
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        tmp.replace(path)
        return True
    except OSError:
        return False


def delete_state(base: Path) -> bool:
    """Remove persisted session file (and any pending temp write)."""
    path = state_path(base)
    ok = True
    for p in (path, path.with_suffix(path.suffix + ".tmp")):
        try:
            if p.is_file():
                p.unlink()
        except OSError:
            ok = False
    return ok
