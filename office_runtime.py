from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Iterable, Optional


OFFICE_RUNTIME_ENV = "QUNZHONGVOTE_OFFICE_HOME"
FORCE_BUNDLED_OFFICE_ENV = "QUNZHONGVOTE_FORCE_BUNDLED_OFFICE"


def application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _resource_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", application_dir())).resolve()


def _as_soffice_path(value: str | Path) -> Path:
    path = Path(value).expanduser()
    if path.name.lower() in {"soffice.exe", "soffice.com", "soffice"}:
        return path
    direct = path / "program" / "soffice.exe"
    if direct.exists():
        return direct
    nested = path / "LibreOffice" / "program" / "soffice.exe"
    if nested.exists():
        return nested
    return direct


def bundled_soffice_candidates() -> Iterable[Path]:
    configured = os.environ.get(OFFICE_RUNTIME_ENV, "").strip()
    if configured:
        yield _as_soffice_path(configured)

    roots = [application_dir(), _resource_dir(), Path(__file__).resolve().parent]
    seen: set[str] = set()
    for root in roots:
        for relative in (
            Path("runtime/libreoffice/program/soffice.exe"),
            Path("vendor/libreoffice/program/soffice.exe"),
        ):
            candidate = (root / relative).resolve()
            key = os.path.normcase(str(candidate))
            if key not in seen:
                seen.add(key)
                yield candidate


def find_bundled_soffice() -> Optional[Path]:
    return next((path for path in bundled_soffice_candidates() if path.is_file()), None)


def find_system_soffice() -> Optional[Path]:
    candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    return next((Path(item) for item in candidates if item and Path(item).is_file()), None)


def find_soffice() -> Optional[Path]:
    return find_bundled_soffice() or find_system_soffice()


def force_bundled_office() -> bool:
    return os.environ.get(FORCE_BUNDLED_OFFICE_ENV, "").strip().lower() in {"1", "true", "yes", "on"}


def _decode_output(value: bytes | None) -> str:
    if not value:
        return ""
    for encoding in ("utf-8", "gb18030", "mbcs"):
        try:
            return value.decode(encoding).strip()
        except (LookupError, UnicodeDecodeError):
            continue
    return value.decode("utf-8", errors="replace").strip()


def convert_with_soffice(
    source_path: str | Path,
    output_dir: str | Path,
    convert_to: str,
    output_suffix: str,
    timeout_seconds: int = 180,
) -> Path:
    source = Path(source_path).resolve()
    destination = Path(output_dir).resolve()
    soffice = find_soffice()
    if soffice is None:
        raise RuntimeError("程序内置文档引擎缺失，请重新下载完整便携版压缩包并全部解压后运行。")

    destination.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="QunzhongVote-Office-") as profile_dir:
        command = [
            str(soffice),
            f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--nofirststartwizard",
            "--convert-to",
            convert_to,
            "--outdir",
            str(destination),
            str(source),
        ]
        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            completed = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=timeout_seconds,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(f"内置文档引擎转换超时（{timeout_seconds} 秒），请检查模板是否损坏或被其他程序占用。") from exc

    generated = destination / f"{source.stem}{output_suffix}"
    if completed.returncode != 0 or not generated.is_file():
        details = _decode_output(completed.stderr) or _decode_output(completed.stdout) or "未返回详细错误"
        raise RuntimeError(f"内置文档引擎转换失败：{details}")
    return generated
