import csv
import json
import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared" / "fulfillment"))

import order_consolidation as consolidation


FIELDS = [
    "Order ID",
    "Status",
    "Item Sku",
    "Item Quantity",
    "Ship to Name",
    "Street",
    "Suburb",
    "State",
    "Postcode",
    "Phone Number",
]


def order(order_id, sku="SKU-001", quantity="1", street="10 Test Street", phone="0412 345 678"):
    return {
        "Order ID": order_id,
        "Status": "processing",
        "Item Sku": sku,
        "Item Quantity": quantity,
        "Ship to Name": "Test Buyer",
        "Street": street,
        "Suburb": "Sydney",
        "State": "NSW",
        "Postcode": "2000",
        "Phone Number": phone,
    }


def write_orders(path, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)


class OrderConsolidationTests(unittest.TestCase):
    def test_merges_only_normalized_recipient_and_exact_sku(self):
        rows = [
            order("ORDER-001"),
            order("ORDER-002", quantity="2", street="  10  TEST street ", phone="(0412) 345-678"),
            order("ORDER-003", sku="SKU-002"),
            order("ORDER-004", street="11 Test Street"),
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "orders.csv"
            write_orders(source, rows)
            plan = consolidation.build_plan(
                rows,
                "2026-07-29",
                "yunexpress",
                source,
                {"SKU-001": {"chinese_name": "Item one"}},
            )

        self.assertEqual(plan["order_count"], 4)
        self.assertEqual(plan["package_count"], 3)
        self.assertEqual(plan["merged_group_count"], 1)
        self.assertEqual(plan["saved_parcel_count"], 1)
        merged = next(package for package in plan["packages"] if package["merged"])
        self.assertEqual(merged["carrier_reference"], "M260729001")
        self.assertEqual(merged["total_quantity"], 3)
        self.assertEqual(
            [line["order_id"] for line in merged["order_lines"]],
            ["ORDER-001", "ORDER-002"],
        )

    def test_disabled_plan_keeps_one_package_per_order(self):
        rows = [order("ORDER-001"), order("ORDER-002")]
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "orders.csv"
            write_orders(source, rows)
            plan = consolidation.build_plan(
                rows,
                "2026-07-29",
                "sf",
                source,
                merge_enabled=False,
            )

        self.assertEqual(plan["merged_group_count"], 0)
        self.assertEqual(
            [package["carrier_reference"] for package in plan["packages"]],
            ["ORDER-001", "ORDER-002"],
        )

    def test_expands_one_carrier_parcel_to_all_original_dp_orders(self):
        rows = [order("ORDER-001"), order("ORDER-002", quantity="2")]
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "orders.csv"
            write_orders(source, rows)
            plan = consolidation.build_plan(rows, "2026-07-29", "sf", source)

        reference = plan["packages"][0]["carrier_reference"]
        carrier_rows = [
            {
                "Order ID": reference,
                "Carrier": "SF INTERNATIONAL",
                "Tracking Number": "SF0000000000001",
                "Source Field": "SF waybill",
            }
        ]
        expanded = consolidation.expand_carrier_rows(carrier_rows, plan)

        self.assertEqual({row["Order ID"] for row in expanded}, {"ORDER-001", "ORDER-002"})
        self.assertEqual(Counter(row["Tracking Number"] for row in expanded), {"SF0000000000001": 2})
        self.assertEqual({row["Carrier Reference"] for row in expanded}, {reference})

    def test_plan_hash_prevents_cross_batch_reuse(self):
        rows = [order("ORDER-001"), order("ORDER-002")]
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "orders.csv"
            write_orders(source, rows)
            plan = consolidation.build_plan(rows, "2026-07-29", "yunexpress", source)
            plan_path = Path(temp_dir) / "plan.json"
            plan_path.write_text(json.dumps(plan), encoding="utf-8")
            write_orders(source, [order("OTHER-001")])

            with self.assertRaisesRegex(ValueError, "does not belong"):
                consolidation.load_plan(plan_path, source, "yunexpress", "2026-07-29")


if __name__ == "__main__":
    unittest.main()
