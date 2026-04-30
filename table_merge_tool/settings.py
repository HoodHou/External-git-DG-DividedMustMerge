from __future__ import annotations

import json
import os
from pathlib import Path


APP_DIR_NAME = "FenJiuBiHe"
MAX_RECENT_ROOTS = 12
MAX_QUICK_ROOTS = 16
MAX_RECENT_CONFIGS = 10

DIFF_GRANULARITIES = ("char", "word", "line")


DEFAULT_SETTINGS = {
    "quick_roots": [],
    "recent_roots": [],
    "recent_configs": [],
    "google_auth_mode": "service_account",
    "google_service_account_path": "",
    "google_oauth_client_path": "",
    "google_oauth_token_path": "",
    "left_root": "",
    "right_root": "",
    "base_root": "",
    "left_source": "local",
    "right_source": "local",
    "base_source": "local",
    "left_file": "",
    "right_file": "",
    "base_file": "",
    "left_file_path": "",
    "right_file_path": "",
    "base_file_path": "",
    "left_revision": "HEAD",
    "right_revision": "HEAD",
    "base_revision": "HEAD",
    "template_source": "left",
    "ignore_trim_whitespace_diff": False,
    "manual_key_fields": "",
    "strict_single_id_mode": False,
    "sheet_key_fields": {},
    "three_way_enabled": False,
    "diff_granularity": "char",
    "table_row_height": 24,
    "table_header_height": 32,
    "sheet_panel_collapsed": False,
    "source_panel_collapsed": False,
}


def _settings_dir() -> Path:
    appdata = os.getenv("LOCALAPPDATA") or os.getenv("APPDATA")
    if appdata:
        return Path(appdata) / APP_DIR_NAME
    return Path.home() / f".{APP_DIR_NAME.lower()}"


SETTINGS_FILE = _settings_dir() / "tool_settings.json"


def load_settings() -> dict:
    if not SETTINGS_FILE.exists():
        return dict(DEFAULT_SETTINGS)
    try:
        payload = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return dict(DEFAULT_SETTINGS)
    settings = dict(DEFAULT_SETTINGS)
    settings.update(payload if isinstance(payload, dict) else {})
    settings["quick_roots"] = _normalize_quick_roots(settings.get("quick_roots"))
    settings["recent_roots"] = _normalize_recent_roots(settings.get("recent_roots"))
    settings["recent_configs"] = _normalize_recent_configs(settings.get("recent_configs"))
    settings["google_auth_mode"] = _normalize_google_auth_mode(settings.get("google_auth_mode"))
    settings["google_service_account_path"] = _normalize_google_service_account_path(
        settings.get("google_service_account_path")
    )
    settings["google_oauth_client_path"] = _normalize_google_path_setting(settings.get("google_oauth_client_path"))
    settings["google_oauth_token_path"] = _normalize_google_path_setting(settings.get("google_oauth_token_path"))
    settings["ignore_trim_whitespace_diff"] = _normalize_bool_setting(settings.get("ignore_trim_whitespace_diff"))
    settings["manual_key_fields"] = _normalize_text_setting(settings.get("manual_key_fields"))
    settings["strict_single_id_mode"] = _normalize_bool_setting(settings.get("strict_single_id_mode"))
    settings["sheet_key_fields"] = _normalize_sheet_key_fields(settings.get("sheet_key_fields"))
    settings["three_way_enabled"] = _normalize_bool_setting(settings.get("three_way_enabled"))
    settings["diff_granularity"] = _normalize_diff_granularity(settings.get("diff_granularity"))
    settings["table_row_height"] = _normalize_int_setting(settings.get("table_row_height"), 24, 18, 56)
    settings["table_header_height"] = _normalize_int_setting(settings.get("table_header_height"), 32, 22, 160)
    settings["sheet_panel_collapsed"] = _normalize_bool_setting(settings.get("sheet_panel_collapsed"))
    settings["source_panel_collapsed"] = _normalize_bool_setting(settings.get("source_panel_collapsed"))
    return settings


def save_settings(settings: dict) -> None:
    payload = dict(DEFAULT_SETTINGS)
    payload.update(settings)
    payload["quick_roots"] = _normalize_quick_roots(payload.get("quick_roots"))
    payload["recent_roots"] = _normalize_recent_roots(payload.get("recent_roots"))
    payload["recent_configs"] = _normalize_recent_configs(payload.get("recent_configs"))
    payload["google_auth_mode"] = _normalize_google_auth_mode(payload.get("google_auth_mode"))
    payload["google_service_account_path"] = _normalize_google_service_account_path(
        payload.get("google_service_account_path")
    )
    payload["google_oauth_client_path"] = _normalize_google_path_setting(payload.get("google_oauth_client_path"))
    payload["google_oauth_token_path"] = _normalize_google_path_setting(payload.get("google_oauth_token_path"))
    payload["ignore_trim_whitespace_diff"] = _normalize_bool_setting(payload.get("ignore_trim_whitespace_diff"))
    payload["manual_key_fields"] = _normalize_text_setting(payload.get("manual_key_fields"))
    payload["strict_single_id_mode"] = _normalize_bool_setting(payload.get("strict_single_id_mode"))
    payload["sheet_key_fields"] = _normalize_sheet_key_fields(payload.get("sheet_key_fields"))
    payload["three_way_enabled"] = _normalize_bool_setting(payload.get("three_way_enabled"))
    payload["diff_granularity"] = _normalize_diff_granularity(payload.get("diff_granularity"))
    payload["table_row_height"] = _normalize_int_setting(payload.get("table_row_height"), 24, 18, 56)
    payload["table_header_height"] = _normalize_int_setting(payload.get("table_header_height"), 32, 22, 160)
    payload["sheet_panel_collapsed"] = _normalize_bool_setting(payload.get("sheet_panel_collapsed"))
    payload["source_panel_collapsed"] = _normalize_bool_setting(payload.get("source_panel_collapsed"))
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def remember_quick_root(settings: dict, root: str, name: str = "") -> dict:
    root = (root or "").strip()
    if not root:
        return settings
    quick_roots = _normalize_quick_roots(settings.get("quick_roots"))
    quick_roots = [item for item in quick_roots if item["path"] != root]
    quick_roots.insert(0, {"name": _quick_root_name(name, root), "path": root})
    settings["quick_roots"] = quick_roots[:MAX_QUICK_ROOTS]
    return settings


def forget_quick_root(settings: dict, root: str) -> dict:
    root = (root or "").strip()
    if not root:
        return settings
    quick_roots = _normalize_quick_roots(settings.get("quick_roots"))
    settings["quick_roots"] = [item for item in quick_roots if item["path"] != root]
    return settings


def replace_quick_roots(settings: dict, entries: list[dict]) -> dict:
    settings["quick_roots"] = _normalize_quick_roots(entries)
    return settings


def remember_root(settings: dict, root: str) -> dict:
    root = (root or "").strip()
    if not root:
        return settings
    recent_roots = _normalize_recent_roots(settings.get("recent_roots"))
    recent_roots = [item for item in recent_roots if item != root]
    recent_roots.insert(0, root)
    settings["recent_roots"] = recent_roots[:MAX_RECENT_ROOTS]
    return settings


def remember_config(settings: dict, config: dict) -> dict:
    normalized = _normalize_config(config)
    if not normalized:
        return settings
    recent_configs = _normalize_recent_configs(settings.get("recent_configs"))
    signature = _config_signature(normalized)
    recent_configs = [item for item in recent_configs if _config_signature(item) != signature]
    recent_configs.insert(0, normalized)
    settings["recent_configs"] = recent_configs[:MAX_RECENT_CONFIGS]
    return settings


def format_config_label(config: dict) -> str:
    left_file = config.get("left_file") or "(左侧未选)"
    right_file = config.get("right_file") or "(右侧未选)"
    left_source = str(config.get("left_source") or "local").upper()
    right_source = str(config.get("right_source") or "local").upper()
    left_revision = _display_revision(config.get("left_revision", ""))
    right_revision = _display_revision(config.get("right_revision", ""))
    template_side = "左模板" if config.get("template_source") == "left" else "右模板"
    key_fields = str(config.get("manual_key_fields") or "").strip()
    key_prefix = "强ID:" if config.get("strict_single_id_mode") else "键:"
    key_suffix = f" | {key_prefix}{key_fields}" if key_fields else ""
    sheet_key_count = len(_normalize_sheet_key_fields(config.get("sheet_key_fields")))
    if sheet_key_count:
        key_suffix += f" | 检测ID:{sheet_key_count}"
    return f"{left_source}:{left_file}{left_revision}  <->  {right_source}:{right_file}{right_revision}  |  {template_side}{key_suffix}"


def _normalize_recent_roots(values) -> list[str]:
    return _normalize_root_list(values, MAX_RECENT_ROOTS)


def _normalize_quick_roots(values) -> list[dict]:
    if not isinstance(values, list):
        return []
    result: list[dict] = []
    seen_paths: set[str] = set()
    for value in values:
        if isinstance(value, dict):
            path = str(value.get("path") or "").strip()
            name = str(value.get("name") or "").strip()
        else:
            path = str(value or "").strip()
            name = ""
        if not path or path in seen_paths:
            continue
        seen_paths.add(path)
        result.append({"name": _quick_root_name(name, path), "path": path})
    return result[:MAX_QUICK_ROOTS]


def _normalize_root_list(values, limit: int) -> list[str]:
    if not isinstance(values, list):
        return []
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result[:limit]


def _normalize_google_service_account_path(value) -> str:
    return str(value or "").strip()


def _normalize_google_path_setting(value) -> str:
    return str(value or "").strip()


def _normalize_google_auth_mode(value) -> str:
    text = str(value or "service_account").strip()
    return text if text in {"service_account", "oauth_user"} else "service_account"


def _normalize_bool_setting(value) -> bool:
    return bool(value)


def _normalize_diff_granularity(value) -> str:
    text = str(value or "char").strip().lower()
    return text if text in DIFF_GRANULARITIES else "char"


def _normalize_int_setting(value, default: int, minimum: int, maximum: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = default
    return max(minimum, min(maximum, number))


def _normalize_text_setting(value) -> str:
    return str(value or "").strip()


def _normalize_sheet_key_fields(value) -> dict[str, str]:
    if not isinstance(value, dict):
        return {}
    result: dict[str, str] = {}
    for sheet_name, field_name in value.items():
        sheet = str(sheet_name or "").strip()
        field = _normalize_text_setting(field_name)
        if sheet and field:
            result[sheet] = field
    return result


def _quick_root_name(name: str, path: str) -> str:
    text = (name or "").strip()
    if text:
        return text
    normalized = path.replace("\\", "/").rstrip("/")
    if not normalized:
        return path
    parts = [part for part in normalized.split("/") if part]
    if not parts:
        return path
    if len(parts) >= 2:
        return "/".join(parts[-2:])
    return parts[-1]


def _normalize_recent_configs(values) -> list[dict]:
    if not isinstance(values, list):
        return []
    result: list[dict] = []
    seen: set[tuple] = set()
    for value in values:
        normalized = _normalize_config(value)
        if not normalized:
            continue
        signature = _config_signature(normalized)
        if signature in seen:
            continue
        seen.add(signature)
        result.append(normalized)
    return result[:MAX_RECENT_CONFIGS]


def _normalize_config(value) -> dict:
    if not isinstance(value, dict):
        return {}
    config = {
        "left_root": str(value.get("left_root") or "").strip(),
        "right_root": str(value.get("right_root") or "").strip(),
        "base_root": str(value.get("base_root") or "").strip(),
        "left_source": str(value.get("left_source") or "local").strip() or "local",
        "right_source": str(value.get("right_source") or "local").strip() or "local",
        "base_source": str(value.get("base_source") or "local").strip() or "local",
        "left_file": str(value.get("left_file") or "").strip(),
        "right_file": str(value.get("right_file") or "").strip(),
        "base_file": str(value.get("base_file") or "").strip(),
        "left_file_path": str(value.get("left_file_path") or "").strip(),
        "right_file_path": str(value.get("right_file_path") or "").strip(),
        "base_file_path": str(value.get("base_file_path") or "").strip(),
        "left_revision": str(value.get("left_revision") or "HEAD").strip() or "HEAD",
        "right_revision": str(value.get("right_revision") or "HEAD").strip() or "HEAD",
        "base_revision": str(value.get("base_revision") or "HEAD").strip() or "HEAD",
        "template_source": str(value.get("template_source") or "left").strip() or "left",
        "ignore_trim_whitespace_diff": _normalize_bool_setting(value.get("ignore_trim_whitespace_diff")),
        "manual_key_fields": _normalize_text_setting(value.get("manual_key_fields")),
        "strict_single_id_mode": _normalize_bool_setting(value.get("strict_single_id_mode")),
        "sheet_key_fields": _normalize_sheet_key_fields(value.get("sheet_key_fields")),
        "three_way_enabled": _normalize_bool_setting(value.get("three_way_enabled")),
        "diff_granularity": _normalize_diff_granularity(value.get("diff_granularity")),
    }
    if not any(config[key] for key in ("left_root", "right_root", "left_file", "right_file", "left_file_path", "right_file_path")):
        return {}
    return config


def _config_signature(config: dict) -> tuple:
    return (
        config.get("left_root", ""),
        config.get("right_root", ""),
        config.get("base_root", ""),
        config.get("left_source", ""),
        config.get("right_source", ""),
        config.get("base_source", ""),
        config.get("left_file", ""),
        config.get("right_file", ""),
        config.get("base_file", ""),
        config.get("left_file_path", ""),
        config.get("right_file_path", ""),
        config.get("base_file_path", ""),
        config.get("left_revision", ""),
        config.get("right_revision", ""),
        config.get("base_revision", ""),
        config.get("template_source", ""),
        config.get("ignore_trim_whitespace_diff", False),
        config.get("manual_key_fields", ""),
        config.get("strict_single_id_mode", False),
        tuple(sorted(_normalize_sheet_key_fields(config.get("sheet_key_fields")).items())),
        config.get("three_way_enabled", False),
        config.get("diff_granularity", ""),
    )


def _display_revision(value: str) -> str:
    text = str(value or "").strip()
    if not text or text.upper() == "HEAD":
        return "@HEAD"
    if text.upper() == "WORKING":
        return ""
    if text.startswith("@") or text.startswith("r"):
        return text
    return f"@{text}"
