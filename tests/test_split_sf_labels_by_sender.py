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


if __name__ == "__main__":
    unittest.main()
