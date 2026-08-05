from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk

from PIL import ImageGrab


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def widget_text(widget) -> str:
    try:
        return str(widget.cget("text") or "")
    except Exception:
        return ""


def wait_until(app, predicate, timeout: float, message: str):
    deadline = time.time() + timeout
    while time.time() < deadline:
        app.update()
        if predicate():
            return
        time.sleep(0.05)
    raise TimeoutError(message)


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify that export preview is a rendered Word-to-PDF page.")
    parser.add_argument("template")
    parser.add_argument("data")
    parser.add_argument("report")
    parser.add_argument("screenshot")
    args = parser.parse_args()

    profile_root = tempfile.TemporaryDirectory()
    os.environ["LOCALAPPDATA"] = profile_root.name

    from app import VoteDocxApp, messagebox
    from vote_core import auto_mapping_for_template, config_pairs, read_vote_records_for_mapping, selected_mark_pair_refs

    messagebox.showinfo = lambda *_args, **_kwargs: None
    messagebox.showwarning = lambda *_args, **_kwargs: None
    messagebox.showerror = lambda *_args, **_kwargs: None

    template = str(Path(args.template).resolve())
    data = str(Path(args.data).resolve())
    app = VoteDocxApp()
    app.geometry("1240x780+10+10")
    app.deiconify()
    app.template_path.set(template)
    app.data_path.set(data)
    app.output_dir.set(str((PROJECT_DIR / "output" / "qa-v033-preview").resolve()))
    app.mapping = auto_mapping_for_template(template)
    app.records = read_vote_records_for_mapping(data, app.mapping)
    app.debug_completed = True

    main_texts = [widget_text(widget) for widget in descendants(app)]
    if any("打勾位置微调" in text or "选择打勾位置" in text for text in main_texts):
        raise AssertionError("Main page still exposes mark-position adjustment controls")

    preview_record = app.records[0]
    mark_refs = selected_mark_pair_refs(preview_record, app.mapping)
    if not mark_refs:
        raise AssertionError("Preview record did not resolve any visible mark positions")

    app.preview_docx()

    def preview_dialog():
        return next(
            (
                child
                for child in app.winfo_children()
                if isinstance(child, tk.Toplevel) and child.winfo_exists() and child.title() == "真实打印预览与用户信息调整"
            ),
            None,
        )

    wait_until(app, lambda: preview_dialog() is not None, 45, "True print preview window did not open")
    dialog = preview_dialog()
    dialog.deiconify()
    dialog.attributes("-topmost", True)
    dialog.lift()
    dialog.focus_force()
    app.update()

    canvases = [widget for widget in descendants(dialog) if isinstance(widget, tk.Canvas)]
    if len(canvases) != 1:
        raise AssertionError(f"Expected one print-preview canvas, found {len(canvases)}")
    canvas = canvases[0]
    image_items = [item for item in canvas.find_all() if canvas.type(item) == "image"]
    if not image_items:
        raise AssertionError("True print preview canvas contains no rendered PDF page image")

    listboxes = [widget for widget in descendants(dialog) if isinstance(widget, tk.Listbox)]
    adjustment_list = next(
        (
            widget
            for widget in listboxes
            if any(str(widget.get(index)).startswith("打勾 ·") for index in range(widget.size()))
        ),
        None,
    )
    if adjustment_list is None:
        raise AssertionError("True print preview does not list visible mark positions")
    mark_index = next(index for index in range(adjustment_list.size()) if str(adjustment_list.get(index)).startswith("打勾 ·"))
    adjustment_list.selection_clear(0, "end")
    adjustment_list.selection_set(mark_index)
    adjustment_list.event_generate("<<ListboxSelect>>")
    app.update()

    selected_ref = mark_refs[0]
    selected_config = app.mapping["options"][selected_ref["key"]]
    selected_pair = config_pairs(selected_config)[int(selected_ref["pairIndex"])]
    before_style = dict(selected_pair.get("markStyle") or selected_config.get("markStyle") or {})

    right_buttons = [
        widget
        for widget in descendants(dialog)
        if isinstance(widget, ttk.Button) and widget_text(widget) == "→"
    ]
    left_buttons = [
        widget
        for widget in descendants(dialog)
        if isinstance(widget, ttk.Button) and widget_text(widget) == "←"
    ]
    if len(right_buttons) != 1 or len(left_buttons) != 1:
        raise AssertionError(
            f"Expected one horizontal nudge pair, found left={len(left_buttons)}, right={len(right_buttons)}"
        )
    for _index in range(3):
        right_buttons[0].invoke()
    for _index in range(2):
        left_buttons[0].invoke()
    app.update()

    confirm_button = next(
        widget
        for widget in descendants(dialog)
        if isinstance(widget, ttk.Button) and widget_text(widget) == "确认预览，允许导出"
    )
    wait_until(app, lambda: str(confirm_button.cget("state")) == "normal", 45, "Exact print preview refresh did not finish")

    selected_config = app.mapping["options"][selected_ref["key"]]
    selected_pair = config_pairs(selected_config)[int(selected_ref["pairIndex"])]
    after_style = dict(selected_pair.get("markStyle") or selected_config.get("markStyle") or {})
    before_x = float(before_style.get("offsetX") or 0) * (1.0 if before_style.get("offsetUnits") == "pt" else 0.35)
    after_x = float(after_style.get("offsetX") or 0)
    if after_style.get("offsetUnits") != "pt" or abs(after_x - before_x - 1.0) > 0.01:
        raise AssertionError(f"Mark nudge did not persist as one point: before={before_style}, after={after_style}")

    dialog.geometry("1220x820+30+30")
    wait_until(
        app,
        lambda: dialog.winfo_ismapped() and dialog.winfo_width() >= 1000 and dialog.winfo_height() >= 700,
        5,
        "True print preview window was not mapped at its requested size",
    )
    time.sleep(0.5)
    app.update()
    x, y = dialog.winfo_rootx(), dialog.winfo_rooty()
    width, height = dialog.winfo_width(), dialog.winfo_height()
    screenshot = Path(args.screenshot).resolve()
    screenshot.parent.mkdir(parents=True, exist_ok=True)
    captured = ImageGrab.grab(bbox=(x, y, x + width, y + height))
    if captured.width < 1000 or captured.height < 700 or captured.getbbox() is None:
        raise AssertionError(f"Captured preview screenshot is invalid: {captured.size}")
    captured.save(screenshot)

    report = {
        "pdfPageImages": len(image_items),
        "visibleMarkAdjustments": sum(
            1 for index in range(adjustment_list.size()) if str(adjustment_list.get(index)).startswith("打勾 ·")
        ),
        "mainMarkAdjustmentRemoved": True,
        "markOffsetBeforePt": before_x,
        "markOffsetAfterPt": after_x,
        "exactRefreshCompleted": True,
        "screenshot": str(screenshot),
    }
    report_path = Path(args.report).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    return_button = next(
        widget
        for widget in descendants(dialog)
        if isinstance(widget, ttk.Button) and widget_text(widget) == "返回修改"
    )
    return_button.invoke()
    app.update()
    app.destroy()
    profile_root.cleanup()
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
