from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image


PDF_FORMAT = 17
POINTS_PER_MM = 72.0 / 25.4


def docx_to_pdf(docx_path: str | Path, pdf_path: str | Path | None = None) -> Path:
    """Export a DOCX through Word so preview geometry matches printing."""
    source = Path(docx_path).resolve()
    if not source.exists():
        raise FileNotFoundError(f"预览文件不存在：{source}")
    target = Path(pdf_path).resolve() if pdf_path else source.with_suffix(".pdf")
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()

    word_error: Optional[Exception] = None
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

    soffice_candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    soffice = next((item for item in soffice_candidates if item and Path(item).exists()), None)
    if soffice:
        completed = subprocess.run(
            [str(soffice), "--headless", "--convert-to", "pdf", "--outdir", str(target.parent), str(source)],
            capture_output=True,
            text=True,
            check=False,
        )
        generated = target.parent / f"{source.stem}.pdf"
        if completed.returncode == 0 and generated.exists():
            if generated != target:
                generated.replace(target)
            return target

    detail = f"\nWord 导出错误：{word_error}" if word_error else ""
    raise RuntimeError(
        "无法生成真实打印预览。请确认已安装 Microsoft Word；也可以安装 LibreOffice 作为备用转换器。" + detail
    )


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
