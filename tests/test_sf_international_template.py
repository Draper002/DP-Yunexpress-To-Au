import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared" / "fulfillment"))

import generate_sf_international_template as sf


class SfInternationalTemplateTests(unittest.TestCase):
    def test_maps_required_australian_order_fields(self):
        orders = [
            {
                "Order ID": "TEST-ORDER-001",
                "Ship to Name": "Test Recipient",
                "Street": "1 Example Street",
                "Suburb": "Sydney",
                "State": "NSW",
                "Postcode": "2000",
                "Phone Number": "0400000000",
                "Item Sku": "TEST-SKU",
                "Item Quantity": "2",
            }
        ]
        sku_map = {
            "TEST-SKU": {
                "chinese_name": "测试商品",
                "english_name": "test product",
                "price_usd": 8.5,
                "unit_weight_kg": 0.32,
            }
        }

        rows = sf.build_sf_rows(orders, sku_map, "国际小包", "否")

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["country"], "AU")
        self.assertEqual(rows[0]["currency"], "USD/美元")
        self.assertEqual(rows[0]["total_weight"], 0.64)
        self.assertEqual(rows[0]["merchant_sku"], "TEST-SKU")
        self.assertEqual(rows[0]["phone"], "0400000000")
        self.assertEqual(rows[0]["mobile"], "0400000000")

    def test_rejects_unknown_business_type(self):
        with self.assertRaises(ValueError):
            sf.validate_settings("未定义渠道", "否")

    def test_rejects_unknown_battery_value(self):
        with self.assertRaises(ValueError):
            sf.validate_settings("国际小包", "不确定")


if __name__ == "__main__":
    unittest.main()
