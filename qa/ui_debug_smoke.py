from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from app import VoteDocxApp
from vote_core import blank_mapping, docx_page_count


def main() -> None:
    parser = argparse.ArgumentParser(description="Open the debug workspace with known template and data files.")
    parser.add_argument("template")
    parser.add_argument("data")
    parser.add_argument("--geometry-log")
    args = parser.parse_args()

    app = VoteDocxApp()
    app.template_path.set(args.template)
    app.data_path.set(args.data)
    app.mapping = blank_mapping()
    app.template_page_count = docx_page_count(args.template)
    app.load_data_preview()
    app.debug_template()
    app.update_file_status()

    if args.geometry_log:
        def write_geometry_log() -> None:
            widgets = []

            def visit(widget) -> None:
                item = {
                    "path": str(widget),
                    "class": widget.winfo_class(),
                    "x": widget.winfo_rootx(),
                    "y": widget.winfo_rooty(),
                    "width": widget.winfo_width(),
                    "height": widget.winfo_height(),
                    "mapped": bool(widget.winfo_ismapped()),
                }
                try:
                    item["text"] = widget.cget("text")
                except Exception:
                    pass
                widgets.append(item)
                for child in widget.winfo_children():
                    visit(child)

            visit(app)
            Path(args.geometry_log).write_text(json.dumps(widgets, ensure_ascii=False, indent=2), encoding="utf-8")

        app.after(1500, write_geometry_log)
    app.mainloop()


if __name__ == "__main__":
    main()
