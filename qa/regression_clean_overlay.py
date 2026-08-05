from __future__ import annotations

import argparse
import sys
from copy import deepcopy
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from vote_core import auto_mapping_for_template, build_document_for_record, read_vote_records_for_mapping, split_options


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate paired pure/non-pure DOCX files for layout regression.")
    parser.add_argument("template")
    parser.add_argument("data")
    parser.add_argument("output_dir")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    mapping = auto_mapping_for_template(args.template)
    records = read_vote_records_for_mapping(args.data, mapping)
    if not records:
        raise RuntimeError("No records found in data source")

    record = records[0]
    raw_options = list(record.raw.values())[3] if len(record.raw) >= 4 else ""
    record.options = split_options(raw_options, normalize=True, dedupe=False)
    record.result_options = {}
    for clean_mode, filename in ((False, "non_pure.docx"), (True, "pure.docx")):
        current_mapping = deepcopy(mapping)
        current_mapping["cleanMode"] = clean_mode
        document, warnings = build_document_for_record(args.template, current_mapping, record)
        document.save(output_dir / filename)
        if warnings:
            print(f"{filename}: {'; '.join(warnings)}")


if __name__ == "__main__":
    main()
