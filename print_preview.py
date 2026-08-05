from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from PIL import Image

from office_runtime import convert_with_soffice, find_bundled_soffice, force_bundled_office


PDF_FORMAT = 17
POINTS_PER_MM = 72.0 / 25.4
WORD_PROGID = "Word.Application"
WPS_PROGIDS = ("KWPS.Application",)
COM_EXPORT_TIMEOUT_SECONDS = 120


def com_progid_registered(progid: str) -> bool:
    """Return whether a desktop Office COM application is registered."""
    if force_bundled_office():
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
            return True
    except (ImportError, OSError):
        return False


def find_wps_progid() -> Optional[str]:
    return next((progid for progid in WPS_PROGIDS if com_progid_registered(progid)), None)


def _export_with_com_direct(source: Path, target: Path, progid: str) -> Path:
    """Export through a Word-compatible desktop COM application."""
    import pythoncom
    import win32com.client

    initialized = False
    application = None
    document = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        application = win32com.client.DispatchEx(progid)
        application.Visible = False
        application.DisplayAlerts = 0
        document = application.Documents.Open(str(source), False, True, False)
        document.ExportAsFixedFormat(str(target), PDF_FORMAT)
    finally:
        try:
            if document is not None:
                document.Close(False)
        finally:
            try:
                if application is not None:
                    application.Quit()
            finally:
                if initialized:
                    pythoncom.CoUninitialize()

    if not target.is_file() or target.stat().st_size <= 0:
        raise RuntimeError("办公软件没有生成有效的 PDF 文件")
    return target


def run_com_export(progid: str, source_path: str | Path, target_path: str | Path, report_path: str | Path) -> int:
    """Child-process entry point used to isolate Word/WPS automation."""
    source = Path(source_path).resolve()
    target = Path(target_path).resolve()
    report = Path(report_path).resolve()
    report.parent.mkdir(parents=True, exist_ok=True)
    try:
        _export_with_com_direct(source, target, progid)
        result = {"success": True}
    except Exception as exc:
        target.unlink(missing_ok=True)
        result = {"success": False, "error": str(exc).strip() or exc.__class__.__name__}
    report.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    return 0 if result["success"] else 1


def _terminate_process_tree(process: subprocess.Popen) -> None:
    """Stop the exact helper process and its PyInstaller child process."""
    if process.poll() is not None:
        return
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            startupinfo=None,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    else:  # pragma: no cover - Windows is the supported release target
        process.kill()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:  # pragma: no cover - defensive fallback
        process.kill()
        process.wait(timeout=5)


def _export_with_com(source: Path, target: Path, progid: str) -> Path:
    """Run Office automation out of process so a broken COM server cannot freeze the UI."""
    with tempfile.TemporaryDirectory(prefix="QunzhongVote-COM-") as temp_dir:
        report = Path(temp_dir) / "result.json"
        if getattr(sys, "frozen", False):
            command = [sys.executable, "--com-export", str(report), progid, str(source), str(target)]
        else:
            command = [
                sys.executable,
                str(Path(__file__).resolve().with_name("app.py")),
                "--com-export",
                str(report),
                progid,
                str(source),
                str(target),
            ]
        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        process = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        try:
            return_code = process.wait(timeout=COM_EXPORT_TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired as exc:
            _terminate_process_tree(process)
            target.unlink(missing_ok=True)
            raise RuntimeError(f"办公软件转换超时（{COM_EXPORT_TIMEOUT_SECONDS} 秒），已自动停止并尝试其他引擎") from exc

        result: Dict[str, Any] = {}
        try:
            result = json.loads(report.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            pass
        if return_code != 0 or not result.get("success"):
            details = str(result.get("error") or "办公软件后台转换失败")
            target.unlink(missing_ok=True)
            raise RuntimeError(details)
        if not target.is_file() or target.stat().st_size <= 0:
            raise RuntimeError("办公软件没有生成有效的 PDF 文件")
    return target


def export_with_word(source: Path, target: Path) -> Path:
    return _export_with_com(source, target, WORD_PROGID)


def export_with_wps(source: Path, target: Path, progid: str) -> Path:
    return _export_with_com(source, target, progid)


def export_with_bundled_engine(source: Path, target: Path) -> Path:
    with tempfile.TemporaryDirectory(prefix="QunzhongVote-PDF-") as converted_dir:
        generated = convert_with_soffice(source, converted_dir, "pdf:writer_pdf_Export", ".pdf")
        generated.replace(target)
    if not target.is_file() or target.stat().st_size <= 0:
        raise RuntimeError("内置文档引擎没有生成有效的 PDF 文件")
    return target


def _attempt_backend(
    name: str,
    exporter: Callable[[], Path],
    target: Path,
    errors: List[str],
) -> Optional[Path]:
    target.unlink(missing_ok=True)
    try:
        return exporter()
    except Exception as exc:
        target.unlink(missing_ok=True)
        details = str(exc).strip() or exc.__class__.__name__
        errors.append(f"{name}：{details}")
        return None


def docx_to_pdf(docx_path: str | Path, pdf_path: str | Path | None = None) -> Path:
    """Export a DOCX using Word, WPS, then the bundled engine."""
    source = Path(docx_path).resolve()
    if not source.exists():
        raise FileNotFoundError(f"预览文件不存在：{source}")
    target = Path(pdf_path).resolve() if pdf_path else source.with_suffix(".pdf")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.unlink(missing_ok=True)
    errors: List[str] = []

    if not force_bundled_office() and com_progid_registered(WORD_PROGID):
        result = _attempt_backend(
            "Microsoft Word",
            lambda: export_with_word(source, target),
            target,
            errors,
        )
        if result is not None:
            return result

    if not force_bundled_office():
        wps_progid = find_wps_progid()
        if wps_progid:
            result = _attempt_backend(
                "WPS Writer",
                lambda: export_with_wps(source, target, wps_progid),
                target,
                errors,
            )
            if result is not None:
                return result

    result = _attempt_backend(
        "内置文档引擎",
        lambda: export_with_bundled_engine(source, target),
        target,
        errors,
    )
    if result is not None:
        return result

    bundled_state = "已找到" if find_bundled_soffice() else "未找到"
    details = "；".join(errors) if errors else "未返回详细错误"
    raise RuntimeError(f"无法生成真实打印预览。内置文档引擎{bundled_state}。{details}")


def page_size_label(width_points: float, height_points: float) -> str:
    width_mm = width_points / POINTS_PER_MM
    height_mm = height_points / POINTS_PER_MM
    portrait_width, portrait_height = sorted((width_mm, height_mm))
    known = {
        "A3": (297.0, 420.0),
        "A4": (210.0, 297.0),
        "A5": (148.0, 210.0),
        "Letter": (215.9, 279.4),
        "Legal": (215.9, 355.6),
    }
    name = "自定义纸张"
    for candidate, (known_width, known_height) in known.items():
        if abs(portrait_width - known_width) <= 2.0 and abs(portrait_height - known_height) <= 2.0:
            name = candidate
            break
    orientation = "横向" if width_mm > height_mm else "纵向"
    return f"{name} {width_mm:.1f}×{height_mm:.1f} mm（{orientation}）"


def render_pdf_pages(pdf_path: str | Path, zoom: float = 1.5) -> Tuple[List[Image.Image], List[Dict[str, Any]]]:
    try:
        import fitz
    except ImportError as exc:  # pragma: no cover - dependency error
        raise RuntimeError("缺少 PyMuPDF，无法显示真实 PDF 打印预览。") from exc

    document = fitz.open(str(Path(pdf_path).resolve()))
    images: List[Image.Image] = []
    page_info: List[Dict[str, Any]] = []
    try:
        matrix = fitz.Matrix(float(zoom), float(zoom))
        for page_index, page in enumerate(document):
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            images.append(image)
            page_info.append(
                {
                    "page": page_index,
                    "widthPoints": float(page.rect.width),
                    "heightPoints": float(page.rect.height),
                    "label": page_size_label(float(page.rect.width), float(page.rect.height)),
                }
            )
    finally:
        document.close()
    return images, page_info


def search_pdf_text(pdf_path: str | Path, text: str) -> List[Dict[str, float]]:
    if not str(text or "").strip():
        return []
    import fitz

    document = fitz.open(str(Path(pdf_path).resolve()))
    matches: List[Dict[str, float]] = []
    try:
        for page_index, page in enumerate(document):
            for rect in page.search_for(str(text).strip()):
                matches.append(
                    {
                        "page": page_index,
                        "x0": float(rect.x0),
                        "y0": float(rect.y0),
                        "x1": float(rect.x1),
                        "y1": float(rect.y1),
                    }
                )
    finally:
        document.close()
    return matches
