import csv
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared" / "fulfillment"))

import order_consolidation as consolidation
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

    def test_merged_label_is_routed_only_to_special_supplier_package(self):
        rows = []
        for order_id, quantity in (("ORDER-001", "1"), ("ORDER-002", "2")):
            rows.append(
                {
                    "Order ID": order_id,
                    "Status": "processing",
                    "Item Sku": "SKU-001",
                    "Item Quantity": quantity,
                    "Ship to Name": "Test Buyer",
                    "Street": "10 Test Street",
                    "Suburb": "Sydney",
                    "State": "NSW",
                    "Postcode": "2000",
                    "Phone Number": "0412345678",
                }
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            dp_csv = root / "orders.csv"
            with dp_csv.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
                writer.writeheader()
                writer.writerows(rows)

            plan = consolidation.build_plan(
                rows,
                "2026-07-29",
                "yunexpress",
                dp_csv,
                {"SKU-001": {"chinese_name": "Test product"}},
            )
            merge_map, _ = consolidation.write_plan_files(plan, root / "carrier.xlsx")
            reference = plan["packages"][0]["carrier_reference"]

            yun_orders = root / "yun-orders.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "订单信息"
            sheet.append(["客户单号", "运单号", "跟踪号", "订单状态"])
            sheet.append([reference, "YT000000000001", "YT000000000001", "已创建"])
            workbook.save(yun_orders)
            workbook.close()

            sku_file = root / "sku.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["No", "申报中文名", "English name", "SKU"])
            sheet.append([1, "Test product", "Test product", "SKU-001"])
            workbook.save(sku_file)
            workbook.close()

            labels_zip = root / "labels.zip"
            with zipfile.ZipFile(labels_zip, "w") as archive:
                archive.writestr("YT000000000001.pdf", b"%PDF-placeholder")

            output_dir = sorter.build_package(
                labels_zip,
                yun_orders,
                dp_csv,
                sku_file,
                "2026-07-29",
                root / "results",
                merge_map,
            )

            merged_folders = [path for path in (output_dir / "合并件").iterdir() if path.is_dir()]
            self.assertEqual(len(merged_folders), 1)
            merged_folder = merged_folders[0]
            self.assertEqual(len(list(merged_folder.glob("*.pdf"))), 1)
            self.assertTrue((merged_folder / "合并订单说明.xlsx").exists())
            supplier_zips = list(merged_folder.glob("*.zip"))
            self.assertEqual(len(supplier_zips), 1)
            with zipfile.ZipFile(supplier_zips[0]) as archive:
                self.assertEqual(len([name for name in archive.namelist() if name.endswith(".pdf")]), 1)
                self.assertIn("合并订单说明.xlsx", archive.namelist())
            self.assertTrue((output_dir / "合并订单与运单对应表.xlsx").exists())
            self.assertEqual(
                [path for path in output_dir.iterdir() if path.is_dir() and path.name != "合并件"],
                [],
            )


if __name__ == "__main__":
    unittest.main()
