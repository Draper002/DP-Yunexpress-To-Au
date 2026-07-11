# DP&Yunexpress To Au

DropShipZone 澳大利亚订单到云途发货的本地桌面工具，同时维护 Windows 与 Apple Silicon macOS 版本。

## 当前功能

1. 从 DP 订单 CSV 生成云途批量寄件 Excel。
2. 从云途订单信息生成 DP 发货回填模板。
3. 按 SKU 分拣云途面单 PDF，并为每个供应商生成独立 ZIP。

## 当前状态

- Windows：业务流程已验证，OS 26 风格浅色界面已打包测试。
- macOS：M5 MacBook Air / macOS 26 的 arm64 版本待开发。
- 最新进度：见 [`docs/DEVELOPMENT_PROGRESS.md`](docs/DEVELOPMENT_PROGRESS.md)。
- Mac 交接：见 [`docs/MACOS_M5_HANDOFF_ZH.md`](docs/MACOS_M5_HANDOFF_ZH.md)。

## 源码结构

```text
app.py
scripts/
  generate_yunexpress_template.py
  generate_dp_shipment_upload.py
  sort_yunexpress_labels_by_sku.py
docs/
  DEVELOPMENT_PROGRESS.md
  MACOS_M5_HANDOFF_ZH.md
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

任何字段映射或匹配规则修改，都必须用同一批脱敏数据在 Windows 和 macOS 进行字段级对比。

