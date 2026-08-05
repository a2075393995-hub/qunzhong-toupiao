from __future__ import annotations

import copy
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional, Tuple


PROFILE_STORE_VERSION = 1


def default_profile_store_path() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    base = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
    return base / "QunzhongVote" / "template_profiles.json"


def template_profile_key(template_path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(template_path).open("rb") as template_file:
        for chunk in iter(lambda: template_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _empty_store() -> Dict[str, Any]:
    return {"version": PROFILE_STORE_VERSION, "profiles": {}}


def _read_store(store_path: Path) -> Dict[str, Any]:
    try:
        raw = json.loads(store_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return _empty_store()
    if not isinstance(raw, dict) or not isinstance(raw.get("profiles"), dict):
        return _empty_store()
    return raw


def load_template_profile(
    template_path: str | Path,
    store_path: str | Path | None = None,
) -> Optional[Dict[str, Any]]:
    path = Path(template_path)
    try:
        key = template_profile_key(path)
    except OSError:
        return None
    store = _read_store(Path(store_path) if store_path is not None else default_profile_store_path())
    profile = store.get("profiles", {}).get(key)
    if not isinstance(profile, dict) or not isinstance(profile.get("mapping"), dict):
        return None
    return copy.deepcopy(profile["mapping"])


def save_template_profile(
    template_path: str | Path,
    mapping: Dict[str, Any],
    store_path: str | Path | None = None,
    profile_key: str | None = None,
) -> bool:
    path = Path(template_path)
    destination = Path(store_path) if store_path is not None else default_profile_store_path()
    try:
        key = profile_key or template_profile_key(path)
        serialized_mapping = json.loads(json.dumps(mapping, ensure_ascii=False))
        store = _read_store(destination)
        store["version"] = PROFILE_STORE_VERSION
        profiles = store.setdefault("profiles", {})
        profiles[key] = {
            "templateName": path.name,
            "updatedAt": datetime.now(timezone.utc).isoformat(),
            "mapping": serialized_mapping,
        }
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_name(f".{destination.name}.{os.getpid()}.tmp")
        temporary.write_text(json.dumps(store, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporary, destination)
        return True
    except (OSError, TypeError, ValueError):
        try:
            temporary.unlink(missing_ok=True)
        except (OSError, UnboundLocalError):
            pass
        return False


def rebuild_pending_pair_index(mapping: Dict[str, Any]) -> Dict[str, int]:
    pending: Dict[str, int] = {}
    options = mapping.get("options", {})
    if not isinstance(options, dict):
        return pending
    for key, config in options.items():
        if not isinstance(config, dict):
            continue
        pairs = config.get("pairs", [])
        if not isinstance(pairs, list):
            continue
        for index, pair in enumerate(pairs):
            if isinstance(pair, dict) and pair.get("judgment") and not pair.get("mark"):
                pending[str(key)] = index
                break
    return pending


def mapping_area_counts(mapping: Dict[str, Any]) -> Tuple[int, int]:
    judgment_count = 0
    mark_count = 0
    options = mapping.get("options", {})
    if not isinstance(options, dict):
        return judgment_count, mark_count
    for config in options.values():
        if not isinstance(config, dict):
            continue
        pairs = config.get("pairs", [])
        if not isinstance(pairs, list):
            continue
        for pair in pairs:
            if not isinstance(pair, dict):
                continue
            judgment_count += int(bool(pair.get("judgment")))
            mark_count += int(bool(pair.get("mark")))
    return judgment_count, mark_count
