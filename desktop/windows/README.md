# DP International Fulfillment To Au

DropShipZone 澳大利亚订单到云途或顺丰国际发货的本地桌面工具，同时维护 Windows 与 Apple Silicon macOS 版本。

## 当前功能

1. 第 1 步可选择云途或顺丰国际，并生成相应批量寄件 Excel。
2. 从云途或顺丰国际订单信息生成 DP 发货回填模板，并自动与第 1 步 DP 订单逐单核对。
3. 云途面单按 SKU 分拣并生成供应商 ZIP；顺丰合并面单先与第 2 步运单逐单核对，再按寄方姓名拆成供应商 PDF。
4. 第 1 步可勾选“合并相同收件人且相同 SKU 的订单”，生成前显示候选组数并要求确认。

顺丰国际当前三步均已接入：批量上传、运单号回填 DP、合并面单按寄方姓名拆分。

合单开启后，第 1 步同时生成“订单合并关系”Excel；第 2 步允许且仅允许同一合并组的多个 DP 订单共用一个运单号；第 3 步在结果根目录建立“合并件”，每组合并件包含一张面单、供应商说明和一个可直接发送的 ZIP。不同 SKU 不会自动合并。

## 当前状态

- Windows：业务流程已验证，OS 26 风格浅色界面已打包测试。
- macOS：M5 MacBook Air / macOS 26 的共用源码与 arm64 打包配置已完成，待 Mac 实机构建。
- Web：云途与顺丰三步、多账号隔离和 Docker 部署包已完成，见 [`../../web/README.md`](../../web/README.md) 和 [`../../web/docs/WEB_DEVELOPMENT_PROGRESS.md`](../../web/docs/WEB_DEVELOPMENT_PROGRESS.md)。
- 最新进度：见 [`../../docs/DEVELOPMENT_PROGRESS.md`](../../docs/DEVELOPMENT_PROGRESS.md)。
- Mac 交接：见 [`../macos/MACOS_M5_HANDOFF_ZH.md`](../macos/MACOS_M5_HANDOFF_ZH.md)。

## 源码结构

```text
app.py
../../shared/fulfillment/
  generate_yunexpress_template.py
  generate_sf_international_template.py
  generate_dp_shipment_upload.py
  order_consolidation.py
  sort_yunexpress_labels_by_sku.py
  split_sf_labels_by_sender.py
../../docs/
../macos/
../../web/
```

## 本地运行

```bash
python -m venv .venv
```

Windows：

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

macOS：

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

## 隐私与数据安全

仓库只保存源码和脱敏开发文档。严禁提交真实订单 CSV、买家姓名、地址、电话、云途订单、面单 PDF、SKU/模板工作簿及发货输出。`.gitignore` 已默认排除常见业务文件格式，但提交前仍必须人工检查。

## 固定业务参数

- 云途产品代码：`THPHR`
- 国家：`AU`
- 申报币种：`USD`
- DP 承运商：`YunExpress`

顺丰国际参数：

- 业务类型：从模板提供的 16 个渠道中明确选择。
- 是否带电：每批选择“否”或“是”。
- 国家：`AU`
- 申报币种：`USD/美元`
- DP 承运商：`SF INTERNATIONAL`
- DP `Order ID`：顺丰导出的“客户订单号”
- DP `Tracking Number`：顺丰导出的“顺丰运单号”

任何字段映射或匹配规则修改，都必须用同一批脱敏数据在 Windows 和 macOS 进行字段级对比。
