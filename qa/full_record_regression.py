from __future__ import annotations

import argparse
import json
import sys
from copy import deepcopy
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from vote_core import auto_mapping_for_template, build_document_for_record, read_vote_records_for_mapping


def regression_mapping(template: str | Path, data: str | Path) -> tuple[dict, list]:
    detected = auto_mapping_for_template(template)
    records = read_vote_records_for_mapping(data, detected)
    if not records:
        raise RuntimeError("No records found in data source")

    # The supplied fixture uses result columns as slots while its option values
    # are global template option numbers (1..14). Flatten the slots so the
    # regression exercises every auto-detected template target directly.
    for record in records:
        flattened = [option for options in record.result_options.values() for option in options]
        if flattened:
            record.options = flattened
            record.result_options = {}

    return detected, records


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate every source record in pure and non-pure modes.")
    parser.add_argument("template")
    parser.add_argument("data")
    parser.add_argument("output_dir")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    mapping, records = regression_mapping(args.template, args.data)
    (output_dir / "mapping.json").write_text(
        json.dumps(mapping, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    report = {"records": len(records), "outputs": [], "warnings": []}
    for clean_mode, mode_name in ((False, "non_pure"), (True, "pure")):
        mode_dir = output_dir / mode_name
        mode_dir.mkdir(parents=True, exist_ok=True)
        current_mapping = deepcopy(mapping)
        current_mapping["cleanMode"] = clean_mode
        for record in records:
            document, warnings = build_document_for_record(args.template, current_mapping, record)
            filename = f"{record.row_no:03d}_{record.room}_{record.name}.docx"
            path = mode_dir / filename
            document.save(path)
            report["outputs"].append(str(path))
            report["warnings"].extend(warnings)

    (output_dir / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps({"records": len(records), "outputs": len(report["outputs"]), "warnings": len(report["warnings"])}, ensure_ascii=True))


if __name__ == "__main__":
    main()
