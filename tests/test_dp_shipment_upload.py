import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared" / "fulfillment"))

import generate_dp_shipment_upload as shipment


class DpShipmentUploadTests(unittest.TestCase):
    def test_reads_sf_xlsx_with_dp_order_and_sf_waybill(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sf-orders.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["客户订单号", "平台订单号", "顺丰运单号", "订单状态"])
            sheet.append(["TEST000001", "TEST000001", "SF0000000000001", "已下单"])
            sheet.append(["TEST000002", "TEST000002", "SF0000000000002", "已下单"])
            workbook.save(path)

            rows = shipment.read_sf_international_rows(path)

        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Order ID"], "TEST000001")
        self.assertEqual(rows[0]["Carrier"], "SF INTERNATIONAL")
        self.assertEqual(rows[0]["Tracking Number"], "SF0000000000001")
        self.assertEqual(rows[0]["Source Field"], "顺丰运单号")

    def test_validates_sf_carrier_and_unique_tracking(self):
        rows = [
            {
                "Order ID": "TEST000001",
                "Carrier": "SF INTERNATIONAL",
                "Tracking Number": "SF0000000000001",
                "Source Field": "顺丰运单号",
            },
            {
                "Order ID": "TEST000002",
                "Carrier": "SF INTERNATIONAL",
                "Tracking Number": "SF0000000000002",
                "Source Field": "顺丰运单号",
            },
        ]

        shipment.validate_rows(rows, None, "SF International", "SF INTERNATIONAL")

    def test_rejects_duplicate_sf_waybill(self):
        rows = [
            {
                "Order ID": "TEST000001",
                "Carrier": "SF INTERNATIONAL",
                "Tracking Number": "SF0000000000001",
                "Source Field": "顺丰运单号",
            },
            {
                "Order ID": "TEST000002",
                "Carrier": "SF INTERNATIONAL",
                "Tracking Number": "SF0000000000001",
                "Source Field": "顺丰运单号",
            },
        ]

        with self.assertRaises(ValueError):
            shipment.validate_rows(rows, None, "SF International", "SF INTERNATIONAL")


if __name__ == "__main__":
    unittest.main()
