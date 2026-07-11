# Shared fulfillment logic

这里是 Windows、macOS 和 Web 共用的发货处理逻辑，包含：

- `generate_yunexpress_template.py`
- `generate_dp_shipment_upload.py`
- `sort_yunexpress_labels_by_sku.py`

修改字段映射、运单匹配、SKU 分拣或校验规则时，只修改这里，并用三个平台分别验证。平台入口不要复制这些脚本。


