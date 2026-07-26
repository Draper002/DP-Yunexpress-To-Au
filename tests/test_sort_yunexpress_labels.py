import csv
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared" / "fulfillment"))

import sort_yunexpress_labels_by_sku as sorter


class SortYunexpressLabelsTests(unittest.TestCase):
    def test_dp_mapping_ignores_non_processing_orders(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "orders.csv"
            with path.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.DictWriter(
                    stream,
                    fieldnames=["Order ID", "Status", "Item Sku"],
                )
                writer.writeheader()
                writer.writerow(
                    {"Order ID": "CURRENT-001", "Status": "processing", "Item Sku": "SKU-NEW"}
                )
                writer.writerow(
                    {"Order ID": "OLD-001", "Status": "shipped", "Item Sku": "SKU-OLD-A"}
                )
                writer.writerow(
                    {"Order ID": "OLD-001", "Status": "shipped", "Item Sku": "SKU-OLD-B"}
                )

            mapping = sorter.read_dp_orders(path)

        self.assertEqual(mapping, {"CURRENT-001": "SKU-NEW"})

    def test_validation_rejects_dp_order_without_yunexpress_shipment(self):
        labels = [sorter.LabelPdf(zip_name="YT0001.pdf", waybill="YT0001")]
        shipments = {
            "YT0001": sorter.Shipment(
                order_id="ORDER-001",
                waybill="YT0001",
                tracking="",
                status="",
            )
        }

        errors = sorter.validate(
            labels,
            shipments,
            {"ORDER-001": "SKU-001", "ORDER-002": "SKU-002"},
            {"SKU-001": "商品一", "SKU-002": "商品二"},
        )

        self.assertTrue(any("DP orders without YunExpress shipment" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
