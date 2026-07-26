#!/usr/bin/env python3
"""Split a combined SF International label PDF by sender name."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from generate_dp_shipment_upload import read_sf_international_rows


SENDER_PATTERN = re.compile(r"寄方姓名\s*[:：]\s*([^\r\n]+)")
WAYBILL_PATTERN = re.compile(r"\bSF\d{10,20}\b", re.IGNORECASE)
INVALID_FILENAME_PATTERN = re.compile(r"[<>:\"/\\|?*\x00-\x1f]")


def date_folder_name(ship_date: str) -> str:
    match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", ship_date.strip())
    if not match:
        return ship_date.strip()
    year, month, day = match.groups()
    return f"{int(year)}年{int(month)}月{int(day)}日"


def default_output_root(ship_date: str) -> Path:
    return (
        Path.home()
        / "Desktop"
        / "Codex Folder"
        / "DropShipZone"
        / "Codex发货记录"
        / date_folder_name(ship_date)
    )


def unique_directory(path: Path) -> Path:
    if not path.exists():
        return path
    return path.with_name(f"{path.name}_{datetime.now().strftime('%H%M%S')}")


def safe_filename(value: str) -> str:
    cleaned = INVALID_FILENAME_PATTERN.sub("_", value.strip())
    cleaned = re.sub(r"\s+", " ", cleaned).rstrip(". ")
    return cleaned or "未命名寄方"


def extract_label_metadata(text: str, page_number: int) -> tuple[str, str]:
    sender_match = SENDER_PATTERN.search(text)
    if not sender_match or not sender_match.group(1).strip():
        raise ValueError(f"SF label page {page_number} has no usable 寄方姓名.")
    sender_name = sender_match.group(1).strip()

    waybills = sorted({value.upper() for value in WAYBILL_PATTERN.findall(text)})
    if len(waybills) != 1:
        raise ValueError(
            f"SF label page {page_number} must contain exactly one unique SF waybill; found {len(waybills)}."
        )
    return sender_name, waybills[0]


def inspect_labels(reader: PdfReader) -> list[dict[str, object]]:
    labels: list[dict[str, object]] = []
    for page_index, page in enumerate(reader.pages):
        sender_name, waybill = extract_label_metadata(page.extract_text() or "", page_index + 1)
        labels.append(
            {
                "page_index": page_index,
                "page_number": page_index + 1,
                "sender_name": sender_name,
                "waybill": waybill,
            }
        )
    if not labels:
        raise ValueError("SF label PDF contains no pages.")

    duplicate_waybills = sorted(
        value for value, count in Counter(label["waybill"] for label in labels).items() if count > 1
    )
    if duplicate_waybills:
        raise ValueError(
            "Duplicate SF waybills in label PDF: "
            + json.dumps(duplicate_waybills, ensure_ascii=False)
        )
    return labels


def validate_expected_waybills(
    labels: list[dict[str, object]],
    expected_waybills: list[str],
) -> None:
    normalized_expected = [value.strip().upper() for value in expected_waybills if value.strip()]
    duplicate_expected = sorted(
        value for value, count in Counter(normalized_expected).items() if count > 1
    )
    if duplicate_expected:
        raise ValueError(
            "Duplicate SF waybills in order file: "
            + json.dumps(duplicate_expected, ensure_ascii=False)
        )
    label_waybills = {str(label["waybill"]).upper() for label in labels}
    expected_set = set(normalized_expected)
    if label_waybills != expected_set:
        raise ValueError(
            "SF label waybill set does not match SF order file. "
            + json.dumps(
                {
                    "missing_labels": sorted(expected_set - label_waybills),
                    "extra_labels": sorted(label_waybills - expected_set),
                },
                ensure_ascii=False,
            )
        )


def split_labels(
    source_pdf: Path,
    ship_date: str,
    output_root: Path,
    sf_order_file: Path | None = None,
) -> tuple[Path, dict[str, object]]:
    reader = PdfReader(source_pdf)
    if reader.is_encrypted:
        raise ValueError("Encrypted SF label PDFs are not supported.")
    labels = inspect_labels(reader)
    expected_waybill_count = None
    if sf_order_file:
        shipment_rows = read_sf_international_rows(sf_order_file)
        expected_waybills = [row["Tracking Number"] for row in shipment_rows]
        validate_expected_waybills(labels, expected_waybills)
        expected_waybill_count = len(expected_waybills)

    groups: dict[str, list[dict[str, object]]] = {}
    for label in labels:
        groups.setdefault(str(label["sender_name"]), []).append(label)

    batch_parent = output_root / "顺丰面单按寄方姓名"
    batch_parent.mkdir(parents=True, exist_ok=True)
    batch_dir = unique_directory(batch_parent / f"顺丰面单_{ship_date}_{len(labels)}单")
    batch_dir.mkdir(parents=False, exist_ok=False)

    used_names: set[str] = set()
    group_reports: list[dict[str, object]] = []
    for group_index, (sender_name, sender_labels) in enumerate(groups.items(), start=1):
        filename_stem = f"{ship_date}_{safe_filename(sender_name)}_{len(sender_labels)}单"
        collision_key = filename_stem.casefold()
        if collision_key in used_names:
            filename_stem += f"_组{group_index}"
            collision_key = filename_stem.casefold()
        used_names.add(collision_key)

        output_pdf = batch_dir / f"{filename_stem}.pdf"
        writer = PdfWriter()
        for label in sender_labels:
            writer.add_page(reader.pages[int(label["page_index"])])
        writer.add_metadata(
            {
                "/Title": f"SF labels - {sender_name}",
                "/Subject": f"{len(sender_labels)} labels grouped by sender name",
            }
        )
        with output_pdf.open("wb") as stream:
            writer.write(stream)

        output_reader = PdfReader(output_pdf)
        if output_reader.is_encrypted or len(output_reader.pages) != len(sender_labels):
            raise ValueError(f"Split PDF validation failed: {output_pdf.name}")
        group_reports.append(
            {
                "sender_name": sender_name,
                "page_count": len(sender_labels),
                "waybill_count": len({str(label["waybill"]) for label in sender_labels}),
                "output_pdf": str(output_pdf),
            }
        )

    output_page_count = sum(int(group["page_count"]) for group in group_reports)
    if output_page_count != len(labels):
        raise ValueError(
            f"Split page count mismatch: source={len(labels)}, outputs={output_page_count}"
        )

    summary_path = batch_dir / "寄方姓名分组汇总.csv"
    with summary_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=["寄方姓名", "面单页数", "PDF文件名"])
        writer.writeheader()
        for group in group_reports:
            writer.writerow(
                {
                    "寄方姓名": group["sender_name"],
                    "面单页数": group["page_count"],
                    "PDF文件名": Path(str(group["output_pdf"])).name,
                }
            )

    report = {
        "status": "OK",
        "source_pdf": str(source_pdf),
        "sf_order_file": str(sf_order_file) if sf_order_file else None,
        "source_page_count": len(labels),
        "expected_waybill_count": expected_waybill_count,
        "unique_waybill_count": len({str(label["waybill"]) for label in labels}),
        "sender_group_count": len(group_reports),
        "output_page_count": output_page_count,
        "groups": group_reports,
    }
    (batch_dir / "校验成功报告.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return batch_dir, report


def main() -> int:
    parser = argparse.ArgumentParser(description="Split SF International labels by sender name.")
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--sf-international-file", type=Path, default=None)
    parser.add_argument("--date", required=True)
    parser.add_argument("--output-root", type=Path, default=None)
    args = parser.parse_args()

    try:
        output_root = args.output_root or default_output_root(args.date)
        batch_dir, _report = split_labels(
            args.pdf,
            args.date,
            output_root,
            args.sf_international_file,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"OK: {batch_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
