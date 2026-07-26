# Shared fulfillment logic

这里是 Windows、macOS 和 Web 共用的发货处理逻辑，包含：

- `generate_yunexpress_template.py`
- `generate_sf_international_template.py`
- `generate_dp_shipment_upload.py`
- `sort_yunexpress_labels_by_sku.py`
- `split_sf_labels_by_sender.py`

`generate_dp_shipment_upload.py` 同时支持云途 `.xlsx` 和顺丰国际 `.xls/.xlsx`。顺丰映射固定为：

```text
客户订单号 -> Order ID
SF INTERNATIONAL -> Carrier
顺丰运单号 -> Tracking Number
```

`split_sf_labels_by_sender.py` 读取每页面单中的“寄方姓名”和唯一顺丰运单号，并为每个寄方生成一个合并 PDF。桌面和 Web 入口会自动传入第 2 步顺丰订单数据，要求 PDF 运单号集合一个不少、一个不多。任何缺少寄方姓名、单页出现多个运单号、重复运单号或集合不一致都会停止输出。

修改字段映射、运单匹配、SKU 分拣或校验规则时，只修改这里，并用三个平台分别验证。平台入口不要复制这些脚本。
