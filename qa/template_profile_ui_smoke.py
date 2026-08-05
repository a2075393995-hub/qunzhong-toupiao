from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import time
from pathlib import Path

from PIL import ImageGrab


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify template marks survive a real Tk application restart.")
    parser.add_argument("template")
    parser.add_argument("data")
    parser.add_argument("report")
    parser.add_argument("screenshot")
    args = parser.parse_args()

    profile_root = tempfile.TemporaryDirectory()
    os.environ["LOCALAPPDATA"] = profile_root.name

    from app import VoteDocxApp, filedialog, messagebox
    from template_profiles import mapping_area_counts

    template = str(Path(args.template).resolve())
    data = str(Path(args.data).resolve())
    filedialog.askopenfilename = lambda **_kwargs: template
    messagebox.showinfo = lambda *_args, **_kwargs: None
    messagebox.showwarning = lambda *_args, **_kwargs: None
    messagebox.showerror = lambda *_args, **_kwargs: None

    first = VoteDocxApp()
    first.withdraw()
    first.choose_template()
    first.data_path.set(data)
    first.load_data_preview()
    first.debug_template()
    first.option_list.selection_set(0)
    first.on_option_selected()
    cells = [cell for cell in first.preview_cells if isinstance(cell.get("target"), dict)]
    if len(cells) < 2:
        raise RuntimeError("Template preview did not expose two selectable regions")
    first.set_brush("judgment")
    first.on_cell_clicked(cells[0]["target"], str(cells[0].get("text") or ""))
    first.set_brush("mark")
    first.on_cell_clicked(cells[1]["target"], str(cells[1].get("text") or ""))
    before_counts = mapping_area_counts(first.mapping)
    before_tree = [first.mapping_tree.item(item, "values") for item in first.mapping_tree.get_children()]
    first.save_current_template_profile()
    first.update()
    time.sleep(0.15)
    first.update()
    first.destroy()

    second = VoteDocxApp()
    second.withdraw()
    second.choose_template()
    second.data_path.set(data)
    second.load_data_preview()
    second.debug_template()
    second.update_idletasks()
    after_counts = mapping_area_counts(second.mapping)
    after_tree = [second.mapping_tree.item(item, "values") for item in second.mapping_tree.get_children()]
    if before_counts != (1, 1) or after_counts != before_counts:
        raise AssertionError(f"Template areas were not restored: before={before_counts}, after={after_counts}")
    if before_tree != after_tree:
        raise AssertionError("Mapping tree changed after application restart")

    second.deiconify()
    second.geometry("1240x780+40+40")
    second.attributes("-topmost", True)
    second.lift()
    second.focus_force()
    second.update()
    time.sleep(1)
    second.update()
    x = second.winfo_rootx()
    y = second.winfo_rooty()
    width = second.winfo_width()
    height = second.winfo_height()
    screenshot = Path(args.screenshot).resolve()
    screenshot.parent.mkdir(parents=True, exist_ok=True)
    ImageGrab.grab(bbox=(x, y, x + width, y + height)).save(screenshot)
    second.attributes("-topmost", False)

    report = {
        "beforeCounts": before_counts,
        "afterCounts": after_counts,
        "mappingTreeRows": len(after_tree),
        "status": second.status_text.get(),
        "screenshot": str(screenshot),
    }
    report_path = Path(args.report).resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    second.destroy()
    profile_root.cleanup()
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
