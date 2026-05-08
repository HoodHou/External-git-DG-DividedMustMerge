from __future__ import annotations

import re
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .sources import run_hidden_process, run_svn_command
from .version import (
    APP_UPDATE_ENABLED,
    APP_UPDATE_PACKAGE_URL,
    APP_UPDATE_SOURCE,
    APP_UPDATE_SVN_URL,
    APP_UPDATE_VERSION_URL,
    APP_VERSION,
)


VERSION_FILE_NAME = "APP_VERSION.txt"


@dataclass(frozen=True, slots=True)
class UpdateInfo:
    current_version: str
    latest_version: str
    release_root: str
    version_target: str
    source: str = "svn"
    package_url: str = ""
    notes: str = ""


@dataclass(frozen=True, slots=True)
class PreparedUpdate:
    info: UpdateInfo
    script_path: Path
    release_dir: Path
    temp_dir: Path
    executable_path: Path


ProgressCallback = Callable[[str, int, int], None]


def check_app_update(current_version: str = APP_VERSION) -> UpdateInfo | None:
    if not APP_UPDATE_ENABLED:
        return None
    source = str(APP_UPDATE_SOURCE or "").strip().lower()
    if source == "svn":
        return check_svn_update(APP_UPDATE_SVN_URL, current_version=current_version)
    if source == "github_zip":
        return check_http_zip_update(
            APP_UPDATE_VERSION_URL,
            APP_UPDATE_PACKAGE_URL,
            current_version=current_version,
        )
    raise RuntimeError(f"未知软件更新源类型: {APP_UPDATE_SOURCE}")


def check_svn_update(update_root: str, current_version: str = APP_VERSION) -> UpdateInfo | None:
    release_root = normalize_update_root(update_root)
    if not release_root:
        return None

    version_target = _version_target(release_root)
    raw_version = run_svn_command(["svn", "cat", version_target]).strip()
    latest_version = _parse_version(raw_version)
    if not latest_version:
        raise RuntimeError(f"SVN 更新源缺少有效版本号: {version_target}")
    if compare_versions(latest_version, current_version) <= 0:
        return None
    return UpdateInfo(
        current_version=current_version,
        latest_version=latest_version,
        release_root=_release_root_from_version_target(release_root),
        version_target=version_target,
        source="svn",
        notes=_read_release_notes(_release_root_from_version_target(release_root)),
    )


def check_http_zip_update(
    version_url: str,
    package_url: str,
    current_version: str = APP_VERSION,
) -> UpdateInfo | None:
    version_target = str(version_url or "").strip()
    release_root = str(package_url or "").strip()
    if not version_target or not release_root:
        return None
    try:
        raw_version = _download_text(version_target).strip()
    except Exception:
        raw_version = _detect_github_latest_version(version_target, release_root)
        if not raw_version:
            raise
    latest_version = _parse_version(raw_version)
    if not latest_version:
        raise RuntimeError(f"更新版本文件缺少有效版本号: {version_target}")
    if compare_versions(latest_version, current_version) <= 0:
        return None
    return UpdateInfo(
        current_version=current_version,
        latest_version=latest_version,
        release_root=release_root,
        version_target=version_target,
        source="github_zip",
        package_url=release_root,
    )


def apply_update(info: UpdateInfo, executable_path: str | None = None) -> Path:
    prepared = prepare_update(info, executable_path=executable_path)
    launch_prepared_update(prepared)
    return prepared.script_path


def prepare_update(
    info: UpdateInfo,
    executable_path: str | None = None,
    progress_callback: ProgressCallback | None = None,
) -> PreparedUpdate:
    source = str(info.source or "svn").strip().lower()
    if source == "svn":
        return prepare_svn_update(info, executable_path=executable_path, progress_callback=progress_callback)
    if source == "github_zip":
        return prepare_zip_update(info, executable_path=executable_path, progress_callback=progress_callback)
    raise RuntimeError(f"未知软件更新源类型: {info.source}")


def apply_svn_update(info: UpdateInfo, executable_path: str | None = None) -> Path:
    prepared = prepare_svn_update(info, executable_path=executable_path)
    launch_prepared_update(prepared)
    return prepared.script_path


def prepare_svn_update(
    info: UpdateInfo,
    executable_path: str | None = None,
    progress_callback: ProgressCallback | None = None,
) -> PreparedUpdate:
    if not getattr(sys, "frozen", False):
        raise RuntimeError("源码运行模式不支持自动覆盖更新。请先打包成 exe 后再使用自动更新。")

    exe_path = Path(executable_path or sys.executable).resolve()
    app_dir = exe_path.parent
    temp_dir = Path(tempfile.mkdtemp(prefix="fenjiubihe_update_"))
    export_dir = temp_dir / "release"
    if progress_callback is not None:
        progress_callback("正在导出更新包", 0, 0)
    command = ["svn", "export", "--force", info.release_root, str(export_dir)]
    process = run_hidden_process(command)
    if process.returncode != 0:
        details = _decode_process_output(process.stderr) or _decode_process_output(process.stdout)
        raise RuntimeError(f"导出更新包失败: {' '.join(command)}\n{details}")

    expected_exe = export_dir / exe_path.name
    if not expected_exe.exists():
        raise RuntimeError(
            f"更新包目录中没有找到 {exe_path.name}。\n"
            f"请确认 SVN 更新源指向打包后的发布目录，而不是源码目录。"
        )

    script_path = temp_dir / "apply_update.bat"
    script_path.write_text(_build_update_script(export_dir, app_dir, exe_path, temp_dir), encoding="utf-8")
    if progress_callback is not None:
        progress_callback("更新包已准备好", 1, 1)
    return PreparedUpdate(info=info, script_path=script_path, release_dir=export_dir, temp_dir=temp_dir, executable_path=exe_path)


def apply_zip_update(info: UpdateInfo, executable_path: str | None = None) -> Path:
    prepared = prepare_zip_update(info, executable_path=executable_path)
    launch_prepared_update(prepared)
    return prepared.script_path


def prepare_zip_update(
    info: UpdateInfo,
    executable_path: str | None = None,
    progress_callback: ProgressCallback | None = None,
) -> PreparedUpdate:
    if not getattr(sys, "frozen", False):
        raise RuntimeError("源码运行模式不支持自动覆盖更新。请先打包成 exe 后再使用自动更新。")

    exe_path = Path(executable_path or sys.executable).resolve()
    app_dir = exe_path.parent
    temp_dir = Path(tempfile.mkdtemp(prefix="fenjiubihe_update_"))
    zip_path = temp_dir / "release.zip"
    export_dir = temp_dir / "release"
    _download_file(info.package_url or info.release_root, zip_path, progress_callback=progress_callback)
    if progress_callback is not None:
        progress_callback("正在解压更新包", 0, 0)
    with zipfile.ZipFile(zip_path) as archive:
        archive.extractall(export_dir)
    release_dir = _find_release_dir(export_dir, exe_path.name)
    if release_dir is None:
        raise RuntimeError(
            f"更新 zip 中没有找到 {exe_path.name}。\n"
            "请确认 GitHub 发布包内包含打包后的完整发布目录。"
    )
    script_path = temp_dir / "apply_update.bat"
    script_path.write_text(_build_update_script(release_dir, app_dir, exe_path, temp_dir), encoding="utf-8")
    if progress_callback is not None:
        progress_callback("更新包已准备好", 1, 1)
    return PreparedUpdate(info=info, script_path=script_path, release_dir=release_dir, temp_dir=temp_dir, executable_path=exe_path)


def launch_prepared_update(prepared: PreparedUpdate) -> Path:
    creationflags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0) if sys.platform.startswith("win") else 0
    subprocess.Popen(["cmd.exe", "/c", str(prepared.script_path)], cwd=str(prepared.temp_dir), creationflags=creationflags)
    return prepared.script_path


def normalize_update_root(value: str) -> str:
    text = str(value or "").strip().replace("\\", "/")
    if text.lower().startswith("svn:") and not text.lower().startswith("svn://"):
        text = f"svn://{text[4:].lstrip('/')}"
    return text.rstrip("/")


def compare_versions(left: str, right: str) -> int:
    left_parts = _numeric_version_parts(left)
    right_parts = _numeric_version_parts(right)
    width = max(len(left_parts), len(right_parts))
    left_parts.extend([0] * (width - len(left_parts)))
    right_parts.extend([0] * (width - len(right_parts)))
    if left_parts > right_parts:
        return 1
    if left_parts < right_parts:
        return -1
    left_text = _parse_version(left)
    right_text = _parse_version(right)
    return (left_text > right_text) - (left_text < right_text)


def _version_target(update_root: str) -> str:
    if update_root.lower().endswith((".txt", ".version")):
        return update_root
    return f"{update_root}/{VERSION_FILE_NAME}"


def _release_root_from_version_target(update_root: str) -> str:
    if update_root.lower().endswith((".txt", ".version")):
        return update_root.rsplit("/", 1)[0]
    return update_root


def _parse_version(value: str) -> str:
    first_line = str(value or "").strip().splitlines()[0].strip() if str(value or "").strip() else ""
    if first_line.lower().startswith("version="):
        first_line = first_line.split("=", 1)[1].strip()
    return first_line.lstrip("vV").strip()


def _numeric_version_parts(value: str) -> list[int]:
    numbers = re.findall(r"\d+", _parse_version(value))
    return [int(item) for item in numbers] if numbers else [0]


def _read_release_notes(release_root: str) -> str:
    for name in ("RELEASE_NOTES.txt", "CHANGELOG.txt"):
        try:
            return run_svn_command(["svn", "cat", f"{release_root}/{name}"]).strip()
        except Exception:
            continue
    return ""


def _download_text(url: str) -> str:
    request = urllib.request.Request(url, headers={"User-Agent": "FenJiuBiHe-Updater"})
    with urllib.request.urlopen(request, timeout=15) as response:
        return response.read().decode("utf-8-sig", errors="replace")


def _detect_github_latest_version(*urls: str) -> str:
    for url in urls:
        repo = _github_repo_from_release_url(url)
        if not repo:
            continue
        latest_url = f"https://github.com/{repo}/releases/latest"
        request = urllib.request.Request(latest_url, headers={"User-Agent": "FenJiuBiHe-Updater"})
        with urllib.request.urlopen(request, timeout=15) as response:
            final_url = response.geturl()
            version = _github_tag_from_url(final_url)
            if version:
                return version
            text = response.read().decode("utf-8", errors="replace")
        match = re.search(r"/releases/tag/([^\"?#/<>\s]+)", text)
        if match:
            return urllib.parse.unquote(match.group(1))
    return ""


def _github_repo_from_release_url(url: str) -> str:
    match = re.search(r"https://github\.com/([^/\s]+/[^/\s]+)/releases/", str(url or ""))
    return match.group(1) if match else ""


def _github_tag_from_url(url: str) -> str:
    match = re.search(r"/releases/tag/([^/?#]+)", str(url or ""))
    return urllib.parse.unquote(match.group(1)) if match else ""


def _download_file(url: str, output_path: Path, progress_callback: ProgressCallback | None = None) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "FenJiuBiHe-Updater"})
    with urllib.request.urlopen(request, timeout=60) as response:
        total = int(response.headers.get("Content-Length") or 0)
        downloaded = 0
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("wb") as handle:
            while True:
                chunk = response.read(1024 * 256)
                if not chunk:
                    break
                handle.write(chunk)
                downloaded += len(chunk)
                if progress_callback is not None:
                    progress_callback("正在下载更新包", downloaded, total)


def _find_release_dir(root: Path, exe_name: str) -> Path | None:
    direct = root / exe_name
    if direct.exists():
        return root
    matches = [path.parent for path in root.rglob(exe_name) if path.is_file()]
    if not matches:
        return None
    matches.sort(key=lambda path: len(path.parts))
    return matches[0]


def _decode_process_output(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "cp936"):
        try:
            return data.decode(encoding, errors="replace").strip()
        except Exception:
            continue
    return ""


def _build_update_script(export_dir: Path, app_dir: Path, exe_path: Path, temp_dir: Path) -> str:
    return "\r\n".join(
        [
            "@echo off",
            "chcp 65001 >nul 2>nul",
            f'set "SRC={export_dir}"',
            f'set "DST={app_dir}"',
            f'set "EXE={exe_path}"',
            f'set "PID={_current_pid()}"',
            ":wait_app",
            'tasklist /FI "PID eq %PID%" 2>nul | find "%PID%" >nul',
            "if not errorlevel 1 (",
            "    timeout /t 1 /nobreak >nul",
            "    goto wait_app",
            ")",
            'xcopy "%SRC%\\*" "%DST%\\" /E /I /Y /Q',
            "if errorlevel 1 (",
            "    echo 更新失败，请检查目录权限。",
            "    pause",
            "    exit /b 1",
            ")",
            'start "" "%EXE%"',
            f'start "" cmd /c "timeout /t 2 /nobreak >nul & rmdir /s /q ""{temp_dir}"""',
            "exit /b 0",
            "",
        ]
    )


def _current_pid() -> int:
    try:
        import os

        return os.getpid()
    except Exception:
        return 0
