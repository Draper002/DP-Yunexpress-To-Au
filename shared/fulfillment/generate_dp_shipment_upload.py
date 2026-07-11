#!/usr/bin/env python3
"""Generate DropShipZone shipment upload file from YunExpress order export."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl


DEFAULT_CARRIER = "YunExpress"


def clean(value: object) -> str:
    return "" if value is None else str(value).strip()


def date_folder_name(ship_date: str) -> str:
    match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", clean(ship_date))
    if not match:
        return clean(ship_date)
    year, month, day = match.groups()
    return f"{int(year)}年{int(month)}月{int(day)}日"


def default_output_root(ship_date: str) -> Path:
    return Path.home() / "Desktop" / "Codex Folder" / "DropShipZone" / "Codex发货记录" / date_folder_name(ship_date)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stamp = datetime.now().strftime("%H%M%S")
    return path.with_name(f"{path.stem}_{stamp}{path.suffix}")


def read_dp_order_ids(csv_path: Path) -> set[str]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if "Order ID" not in (reader.fieldnames or []):
            raise ValueError("DP order CSV missing Order ID column.")
        order_ids = {clean(row.get("Order ID")) for row in reader if clean(row.get("Order ID"))}
    if not order_ids:
        raise ValueError("DP order CSV contains no Order ID values.")
    return order_ids


def read_yunexpress_rows(xlsx_path: Path, prefer_tracking: bool) -> list[dict[str, str]]:
    workbook = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet = workbook["订单信息"] if "订单信息" in workbook.sheetnames else workbook.worksheets[0]
    headers = {clean(sheet.cell(1, col).value): col for col in range(1, sheet.max_column + 1)}
    required = ["客户单号", "运单号"]
    missing_headers = [header for header in required if header not in headers]
    if missing_headers:
        raise ValueError(f"YunExpress order file missing columns: {missing_headers}")

    tracking_col = headers.get("跟踪号")
    rows: list[dict[str, str]] = []
    for row_idx in range(2, sheet.max_row + 1):
        order_id = clean(sheet.cell(row_idx, headers["客户单号"]).value)
        waybill = clean(sheet.cell(row_idx, headers["运单号"]).value)
        tracking = clean(sheet.cell(row_idx, tracking_col).value) if tracking_col else ""
        if not order_id and not waybill and not tracking:
            continue
        if not order_id:
            raise ValueError(f"YunExpress row {row_idx} has blank 客户单号.")
        if not waybill:
            raise ValueError(f"YunExpress row {row_idx} has blank 运单号 for {order_id}.")
        tracking_number = tracking if prefer_tracking and tracking else waybill
        if not tracking_number:
            raise ValueError(f"YunExpress row {row_idx} has no usable tracking/waybill for {order_id}.")
        rows.append({
            "Order ID": order_id,
            "Carrier": DEFAULT_CARRIER,
            "Tracking Number": tracking_number,
            "Source Field": "跟踪号" if prefer_tracking and tracking else "运单号",
        })
    if not rows:
        raise ValueError("No YunExpress shipment rows found.")
    return rows


def validate_rows(rows: list[dict[str, str]], expected_order_ids: set[str] | None) -> None:
    order_ids = [row["Order ID"] for row in rows]
    tracking_numbers = [row["Tracking Number"] for row in rows]
    duplicate_order_ids = sorted(order_id for order_id, count in Counter(order_ids).items() if count > 1)
    duplicate_tracking = sorted(value for value, count in Counter(tracking_numbers).items() if count > 1)
    if duplicate_order_ids:
        raise ValueError("Duplicate Order ID in YunExpress rows: " + json.dumps(duplicate_order_ids, ensure_ascii=False))
    if duplicate_tracking:
        raise ValueError("Duplicate tracking/waybill in YunExpress rows: " + json.dumps(duplicate_tracking, ensure_ascii=False))
    if expected_order_ids is not None and set(order_ids) != expected_order_ids:
        missing = sorted(expected_order_ids - set(order_ids))
        extra = sorted(set(order_ids) - expected_order_ids)
        raise ValueError(
            "YunExpress Order ID set does not match DP CSV. "
            + json.dumps({"missing": missing, "extra": extra}, ensure_ascii=False)
        )


def write_upload(template_path: Path, rows: list[dict[str, str]], date: str, output_root: Path) -> Path:
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.worksheets[0]
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row - 1)
    headers = ["Order ID", "Carrier", "Tracking Number"]
    for col_idx, header in enumerate(headers, start=1):
        sheet.cell(1, col_idx).value = header
    for row_idx, row in enumerate(rows, start=2):
        for col_idx, header in enumerate(headers, start=1):
            cell = sheet.cell(row_idx, col_idx)
            cell.value = row[header]
            cell.number_format = "@"

    output_dir = output_root / "DP回填模板"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = unique_path(output_dir / f"DP_发货回填_{date}_{len(rows)}orders_YunExpress_waybill.xlsx")
    workbook.save(output_path)
    return output_path


def validate_output(path: Path) -> dict[str, object]:
    workbook = openpyxl.load_workbook(path, data_only=True)
    sheet = workbook.worksheets[0]
    headers = [sheet.cell(1, col).value for col in range(1, 4)]
    rows = [
        [sheet.cell(row, col).value for col in range(1, 4)]
        for row in range(2, sheet.max_row + 1)
        if any(sheet.cell(row, col).value not in (None, "") for col in range(1, 4))
    ]
    blank_cells = [(idx + 2, col + 1) for idx, row in enumerate(rows) for col, value in enumerate(row) if value in (None, "")]
    return {
        "headers_ok": headers == ["Order ID", "Carrier", "Tracking Number"],
        "row_count": len(rows),
        "blank_cells_count": len(blank_cells),
        "carrier_counts": dict(Counter(row[1] for row in rows)),
        "tracking_unique_count": len({row[2] for row in rows}),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate DP shipment upload from YunExpress order export.")
    parser.add_argument("--yunexpress-xlsx", required=True, type=Path)
    parser.add_argument("--shipment-template-xlsx", required=True, type=Path)
    parser.add_argument("--dp-orders-csv", type=Path, default=None)
    parser.add_argument("--date", required=True)
    parser.add_argument("--prefer-tracking", action="store_true", help="Use 跟踪号 when present; default uses 运单号.")
    parser.add_argument("--output-root", type=Path, default=None)
    args = parser.parse_args()

    try:
        output_root = args.output_root or default_output_root(args.date)
        expected_order_ids = read_dp_order_ids(args.dp_orders_csv) if args.dp_orders_csv else None
        rows = read_yunexpress_rows(args.yunexpress_xlsx, args.prefer_tracking)
        validate_rows(rows, expected_order_ids)
        output_path = write_upload(args.shipment_template_xlsx, rows, args.date, output_root)
        validation = validate_output(output_path)
        if not validation["headers_ok"] or validation["blank_cells_count"] != 0:
            raise ValueError(f"Generated DP upload failed validation: {validation}")
        report = {
            "status": "OK",
            "source_tracking_field": sorted({row["Source Field"] for row in rows}),
            "output_path": str(output_path),
            "validation": validation,
        }
        output_path.with_suffix(".report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"OK: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

