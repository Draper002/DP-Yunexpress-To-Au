import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared" / "fulfillment"))

import split_sf_labels_by_sender as splitter


class SplitSfLabelsBySenderTests(unittest.TestCase):
    def test_extracts_sender_and_deduplicates_repeated_waybill_text(self):
        text = "\n".join(
            [
                "顺丰国际电商",
                "SF6047737000001",
                "寄方姓名: 测试供应商",
                "订单号: TEST000001",
                "SF6047737000001",
            ]
        )

        sender, waybill = splitter.extract_label_metadata(text, 1)

        self.assertEqual(sender, "测试供应商")
        self.assertEqual(waybill, "SF6047737000001")

    def test_rejects_missing_sender(self):
        with self.assertRaises(ValueError):
            splitter.extract_label_metadata("SF6047737000001", 1)

    def test_rejects_multiple_waybills_on_one_page(self):
        text = "寄方姓名: 测试\nSF6047737000001\nSF6047737000002"
        with self.assertRaises(ValueError):
            splitter.extract_label_metadata(text, 1)

    def test_sanitizes_windows_filename_characters(self):
        self.assertEqual(splitter.safe_filename(' 供应商:A/B*? '), "供应商_A_B__")

    def test_rejects_label_waybill_set_mismatch(self):
        labels = [
            {"waybill": "SF6047737000001"},
            {"waybill": "SF6047737000002"},
        ]

        with self.assertRaisesRegex(ValueError, "does not match SF order file"):
            splitter.validate_expected_waybills(
                labels,
                ["SF6047737000001", "SF6047737000003"],
            )

    def test_accepts_exact_label_waybill_set(self):
        labels = [
            {"waybill": "SF6047737000001"},
            {"waybill": "SF6047737000002"},
        ]

        splitter.validate_expected_waybills(
            labels,
            ["sf6047737000002", "SF6047737000001"],
        )

    def test_partition_removes_merged_label_from_normal_sender_groups(self):
        merged_package = {
            "carrier_reference": "M260729001",
            "merged": True,
            "order_lines": [
                {"order_id": "ORDER-001", "quantity": 1},
                {"order_id": "ORDER-002", "quantity": 1},
            ],
        }
        normal_package = {
            "carrier_reference": "ORDER-003",
            "merged": False,
            "order_lines": [{"order_id": "ORDER-003", "quantity": 1}],
        }
        plan = {"packages": [merged_package, normal_package]}
        shipment_rows = [
            {"Order ID": "M260729001", "Tracking Number": "SF0000000000001"},
            {"Order ID": "ORDER-003", "Tracking Number": "SF0000000000002"},
        ]
        labels = [
            {"waybill": "SF0000000000001", "sender_name": "Supplier A", "page_index": 0},
            {"waybill": "SF0000000000002", "sender_name": "Supplier A", "page_index": 1},
        ]

        normal, merged = splitter.partition_labels_by_consolidation(
            labels,
            shipment_rows,
            plan,
        )

        self.assertEqual([label["waybill"] for label in normal], ["SF0000000000002"])
        self.assertEqual([label["waybill"] for label, _ in merged], ["SF0000000000001"])
        self.assertIs(merged[0][1], merged_package)


if __name__ == "__main__":
    unittest.main()
