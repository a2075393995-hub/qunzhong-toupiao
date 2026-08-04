from __future__ import annotations

import argparse
import json

import pdfplumber


ANCHORS = [
    "百家湖西花园小区 2026 年第一次业主大会临时会议",
    "议题一：",
    "议题二：",
    "议题三：",
    "议题四：",
    "二、填票说明：",
    "南京市江宁区百家湖西花园业主委员会",
    "一式三联：",
]


def rounded_box(match):
    return {
        key: round(match[key], 3)
        for key in ("x0", "top", "x1", "bottom")
        if key in match
    }


def inspect(path):
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        anchors = {}
        for text in ANCHORS:
            matches = page.search(text, regex=False) or []
            anchors[text] = [rounded_box(match) for match in matches]
        edges = [
            {
                "x0": round(edge["x0"], 3),
                "top": round(edge["top"], 3),
                "x1": round(edge["x1"], 3),
                "bottom": round(edge["bottom"], 3),
                "orientation": edge.get("orientation"),
            }
            for edge in page.edges
        ]
        return {
            "pages": len(pdf.pages),
            "width": round(page.width, 3),
            "height": round(page.height, 3),
            "anchors": anchors,
            "edges": edges,
        }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("reference")
    parser.add_argument("candidate")
    args = parser.parse_args()
    reference = inspect(args.reference)
    candidate = inspect(args.candidate)
    print(
        json.dumps(
            {
                "reference": reference,
                "candidate": candidate,
                "same_page_size": (
                    reference["width"], reference["height"]
                ) == (candidate["width"], candidate["height"]),
                "same_anchors": reference["anchors"] == candidate["anchors"],
                "same_edges": reference["edges"] == candidate["edges"],
            },
            ensure_ascii=True,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
