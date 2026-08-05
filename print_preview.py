from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image

from office_runtime import convert_with_soffice, find_bundled_soffice, force_bundled_office


PDF_FORMAT = 17
POINTS_PER_MM = 72.0 / 25.4


def docx_to_pdf(docx_path: str | Path, pdf_path: str | Path | None = None) -> Path:
    """Export a DOCX to fixed-layout PDF using Word or the bundled engine."""
    source = Path(docx_path).resolve()
    if not source.exists():
        raise FileNotFoundError(f"预览文件不存在：{source}")
    target = Path(pdf_path).resolve() if pdf_path else source.with_suffix(".pdf")
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()

    word_error: Optional[Exception] = None
    if not force_bundled_office():
        try:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
            ):
                pass
            word_registered = True
        except OSError:
            word_registered = False
    else:
        word_registered = False

    if word_registered:
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            word = None
            document = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                document = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
                document.ExportAsFixedFormat(str(target), PDF_FORMAT)
            finally:
                try:
                    if document is not None:
                        document.Close(False)
                finally:
                    try:
                        if word is not None:
                            word.Quit()
                    finally:
                        pythoncom.CoUninitialize()
            if target.exists():
                return target
        except Exception as exc:  # pragma: no cover - depends on local Office installation
            word_error = exc

    try:
        with tempfile.TemporaryDirectory(prefix="QunzhongVote-PDF-") as converted_dir:
            generated = convert_with_soffice(source, converted_dir, "pdf:writer_pdf_Export", ".pdf")
            generated.replace(target)
        return target
    except Exception as bundled_error:
        word_detail = f"；Microsoft Word 错误：{word_error}" if word_error else ""
        bundled_state = "已找到" if find_bundled_soffice() else "未找到"
        raise RuntimeError(
            f"无法生成真实打印预览。内置文档引擎{bundled_state}但转换失败：{bundled_error}{word_detail}"
        ) from bundled_error


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
