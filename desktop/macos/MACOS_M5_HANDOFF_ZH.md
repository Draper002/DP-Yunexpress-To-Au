# DP International Fulfillment To Au - MacBook M5 / macOS 26 开发交接文档

更新时间：`2026-07-26 21:10 CST (Asia/Shanghai)`

当前仓库已经提供 `desktop/macos/app.py`、Apple Silicon 打包配置和构建脚本。Mac 端下一位 Codex 不需要重新移植业务代码，重点是在 M5 实机完成运行、界面、输出一致性和 `.app` 构建验证。

## 1. 项目目标

将现有 Windows 程序 `DP&Yunexpress To Au` 移植到 Apple Silicon MacBook Air（M5、macOS 26），保持三个发货步骤、字段映射、输出文件内容和目录规则与 Windows 版本一致。

首个 Mac 版本只需供机主本人使用，目前没有 Apple Developer 开发者账户。交付形式优先为可双击运行的 `DP&Yunexpress To Au.app`，同时提供便于保存和传输的 `.dmg` 或 `.zip`。

## 2. 现有项目位置和文件

原始 Windows 源码目录（历史路径）：

```text
C:\Users\Admin\Desktop\Codex Folder\DropShipZone\DP_Yunexpress_To_Au_Tool
```

当前仓库的程序代码目录：

```text
desktop/windows/app.py
desktop/macos/
shared/fulfillment/
├── __init__.py
├── generate_yunexpress_template.py
├── generate_sf_international_template.py
├── generate_dp_shipment_upload.py
├── sort_yunexpress_labels_by_sku.py
└── split_sf_labels_by_sender.py
web/app.py
```

不要把 Windows 的以下内容当作 Mac 构建输入：

```text
build/
dist/
release*/
*.exe
Windows 生成的 *.spec
```

这些文件是 Windows 构建产物，不能在 macOS 上直接运行或转换。

## 3. 不允许改变的业务规则

### 第 1 步：生成物流平台上传模板

- 用户选择云途或顺丰国际，并提供本次 DP 订单 CSV。
- 自动沿用固定配置中的 SKU 商品库及对应物流模板。
- 产品代码固定为 `THPHR`。
- 国家固定为 `AU`。
- 云途申报币种固定为 `USD`；顺丰为 `USD/美元`。
- 订单状态筛选规则保持现有脚本实现，不擅自修改字段映射。
- 输出对应物流平台的批量寄件 Excel。

### 第 2 步：生成 DP 发货回填模板

- 用户选择云途或顺丰国际，并提供对应订单信息 Excel。
- 自动沿用固定配置中的 DP 发货模板。
- 云途承运商为 `YunExpress`；顺丰承运商为 `SF INTERNATIONAL`。
- 顺丰“客户订单号”映射 DP `Order ID`，“顺丰运单号”映射 `Tracking Number`。
- 运单号必须按订单正确回填，不允许仅按表格行号匹配。

### 第 3 步：按 SKU 分拣云途面单

- 用户本次只选择云途面单 ZIP。
- 自动沿用第 1 步 DP 订单 CSV、第 2 步云途订单信息及固定 SKU 商品库。
- ZIP 内每个 PDF 文件名为 `YT` 开头的云途运单号。
- 必须通过订单、运单号和 SKU 关系进行匹配。
- 输出文件夹命名规则：`日期+SKU+申报中文名+数量`。
- 每个 SKU 文件夹内包含该 SKU 的面单 PDF，以及一个包含这些 PDF 的供应商 ZIP。
- 单个 SKU 文件夹内不要放发货明细表。
- 总目录可以保留汇总表和校验报告。
- 任何未匹配、重复匹配、缺少 PDF 或数量不一致都必须明确报错或写入校验报告，不能静默忽略。

顺丰国际模式：

- 用户本次只选择顺丰合并面单 PDF。
- 按每页面单中的“寄方姓名”分组，不依赖订单表或 SKU 商品库。
- 每个寄方输出一个合并 PDF，文件名包含日期、寄方姓名和面单数量。
- 每页必须存在一个寄方姓名和一个唯一顺丰运单号；缺失或重复时停止输出。

## 4. 输出目录规则

Windows 当前规则：

```text
~/Desktop/Codex Folder/DropShipZone/Codex发货记录/YYYY年M月D日
```

Mac 版本建议保持同样的用户可见结构：

```text
~/Desktop/Codex Folder/DropShipZone/Codex发货记录/YYYY年M月D日
```

如果用户的 macOS 将桌面同步到 iCloud，必须使用 `Path.home() / "Desktop"` 实际测试；如果该目录不可写，应允许用户在“固定配置”中选择输出根目录并保存设置。

日期默认采用 Mac 当前系统日期，格式为 `YYYY-MM-DD`，允许手动修改。

## 5. macOS 适配状态

### 5.1 Finder 打开文件夹：已完成

`desktop/macos/app.py` 已继承共享桌面界面，并用 macOS `open` 命令覆盖两个目录打开动作：

```python
subprocess.Popen(["open", str(path)])
```

需要在 Mac 实机确认“打开输出目录”和“打开本次结果”均打开 Finder。

### 5.2 统一输出路径

以下四个文件都存在输出目录逻辑，必须统一，避免界面显示路径与脚本实际保存路径不一致：

```text
app.py
shared/fulfillment/generate_yunexpress_template.py
shared/fulfillment/generate_dp_shipment_upload.py
shared/fulfillment/sort_yunexpress_labels_by_sku.py
```

建议在 `shared/fulfillment/` 新建共享模块，例如 `paths.py`，集中处理：

- 输出根目录
- 日期目录名
- Downloads、Desktop 目录
- 用户自定义配置目录

### 5.3 固定配置持久化

建议将以下配置保存为 JSON：

- SKU 商品库路径
- 云途标准模板路径
- 顺丰国际标准模板路径
- DP 发货模板路径
- 输出根目录

Mac 建议配置文件位置：

```text
~/Library/Application Support/DP Yunexpress To Au/config.json
```

Windows 可继续使用用户目录下的应用配置文件。不要把某台电脑的绝对路径写进安装包。

### 5.4 字体和界面

- macOS 优先使用系统字体 `SF Pro`；Tkinter 无法指定时使用系统默认字体。
- 中文显示必须完整，不能出现方框或截断。
- 默认窗口一次显示三个步骤所需的全部内容，不要求用户手动放大窗口。
- 保持浅色、通透、单一系统蓝强调色的 macOS 26 风格。
- 不要为了视觉改造修改三个业务脚本的字段映射。

## 6. 推荐技术路线

### 第一阶段：最低成本可用版本

继续使用现有 Python、Tkinter、`openpyxl`、`xlrd` 和 `pypdf`：

- 优点：业务代码复用最多，开发和验证最快。
- 缺点：界面只能接近 macOS 风格，无法获得真正原生的液态玻璃效果。

第一阶段完成并验证业务正确后，再决定是否升级界面。

### 第二阶段：更高质量跨平台界面

建议将界面迁移到 PySide6，三个处理脚本继续复用：

- Windows 和 macOS 共用一套界面代码。
- 比 Tkinter 更容易实现圆角、层级、动画、图标和高 DPI 适配。
- 仍需分别在 Windows 和 Apple Silicon Mac 上构建安装包。

暂不建议立即使用 SwiftUI 重写全部程序，因为这会形成两套界面和打包体系，后续维护成本更高。只有明确要求完全原生 Mac 体验时再考虑 SwiftUI。

## 7. Mac 开发环境

在 M5 MacBook 上安装：

1. Xcode Command Line Tools。
2. Apple Silicon 原生 Python，建议 Python 3.11 或 3.12。
3. 创建独立虚拟环境。
4. 安装 `openpyxl`、`xlrd`、`pypdf` 和 `pyinstaller`。

示例命令：

```bash
xcode-select --install
cd "/源码所在目录"
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install openpyxl xlrd pypdf pyinstaller
python app.py
```

先从源码运行并完成测试，不要一开始就打包。

确认 Python 架构：

```bash
python -c "import platform; print(platform.machine())"
```

预期结果：

```text
arm64
```

## 8. Mac 打包流程

源码测试通过后执行：

```bash
source .venv/bin/activate
python -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --onedir \
  --name "DP&Yunexpress To Au" \
  --collect-all openpyxl \
  --collect-all xlrd \
  --collect-all pypdf \
  app.py
```

首轮建议使用 `--onedir`，更容易排查缺失依赖。稳定后再测试 `--onefile`，但 macOS 正式交付仍应以 `.app` 应用包为主。

预期产物：

```text
dist/DP&Yunexpress To Au.app
```

本机无开发者账户测试时，可先进行临时签名：

```bash
codesign --force --deep --sign - "dist/DP&Yunexpress To Au.app"
codesign --verify --deep --strict --verbose=2 "dist/DP&Yunexpress To Au.app"
```

如果系统拦截自己构建的应用，优先在“系统设置 -> 隐私与安全性”中确认允许打开。不要要求用户关闭 Gatekeeper。

创建简单 DMG 可使用 `hdiutil`：

```bash
mkdir -p dmg-stage
cp -R "dist/DP&Yunexpress To Au.app" dmg-stage/
ln -s /Applications dmg-stage/Applications
hdiutil create \
  -volname "DP&Yunexpress To Au" \
  -srcfolder dmg-stage \
  -ov \
  -format UDZO \
  "DP&Yunexpress To Au_macOS26_arm64.dmg"
```

没有 Apple Developer 账户时，该 DMG 仅作为本人测试版，不视为可无提示分发的正式安装包。

## 9. 必须完成的测试

使用同一批脱敏测试数据，在 Windows 和 Mac 分别运行并对比。

### 功能测试

- 第 1 步分别成功生成云途上传 Excel 和顺丰国际上传 XLSM。
- 云途及顺丰后台能够成功接受对应文件。
- 第 2 步分别从云途订单信息和顺丰 `.xls/.xlsx` 生成可由 DP 接受的回填模板。
- 顺丰 DP 回填的承运商必须为 `SF INTERNATIONAL`。
- 第 3 步能够从云途总面单 ZIP 正确分拣 PDF。
- 第 3 步能够从顺丰合并面单 PDF 按寄方姓名生成供应商 PDF。
- 每个 SKU 文件夹内 PDF 数量与文件夹名称中的数量一致。
- 每个供应商 ZIP 内文件数量、名称与同目录 PDF 完全一致。
- 打开输出目录、打开结果目录在 macOS 上有效。
- 日期默认使用 Mac 当前日期，手动修改后归档目录正确。
- 固定配置重启应用后仍保留。

### 一致性测试

- Windows 与 Mac 生成的关键单元格值一致。
- 订单号、运单号、SKU、数量和申报中文名的对应关系一致。
- 允许 XLSX 文件内部元数据时间不同，不以二进制哈希作为唯一判断标准。
- CSV、中文文件名、空格路径和较长路径均需测试。

### 异常测试

- 缺少必要列时给出中文错误提示。
- 找不到模板或 SKU 时不生成半成品。
- 运单号重复或找不到对应订单时停止或输出明确校验失败。
- ZIP 中缺少某个 PDF 时明确列出运单号。
- 输出目录无写入权限时提示用户重新选择目录。

测试日志和截图不得公开展示真实买家姓名、电话和地址。

## 10. Mac 版本交付物

Mac 端 Codex 完成后应提供：

```text
DP&Yunexpress To Au.app
DP&Yunexpress To Au_macOS26_arm64.dmg
Mac构建说明.md
测试结果.md
修改后的完整源码
```

测试结果至少记录：

- 使用的 Mac 型号和 macOS 版本
- Python、PyInstaller、openpyxl 版本
- 三个步骤是否通过
- 与 Windows 输出的字段级对比结果
- 尚未解决的问题

## 11. 给 MacBook Codex 的首条任务指令

可以将下面内容和本文件一起交给 MacBook 上的 Codex：

```text
请阅读“MacBook_M5_macOS26开发交接文档.md”和完整源码。先不要重写业务逻辑。

目标是在 M5 MacBook Air、macOS 26 上完成 DP&Yunexpress To Au 的 Apple Silicon arm64 版本。请先从源码运行，完成跨平台路径、打开文件夹、固定配置持久化和界面适配，再用同一批脱敏订单测试三个步骤。确认输出和 Windows 版本字段级一致后，打包为 .app 和 .dmg。

不要打印或在截图中展示买家姓名、电话、地址。遇到任何字段映射不确定时停止并询问，不要自行猜测。完成后提供修改清单、测试结果、.app、.dmg 和完整源码。
```

## 12. 两个独立 Codex 账号如何共同开发

### 推荐方案：私有 GitHub 仓库

两个 Codex 账号不需要互相访问聊天记录。将源码放进同一个私有 GitHub 仓库，用提交记录共享工作结果：

```text
main             已验证的稳定版本
windows-ui       Windows 界面调整
macos-arm64      Mac M5 / arm64 适配
```

建议流程：

1. Windows 端创建私有仓库，只提交源码、文档和脱敏测试数据。
2. Mac 端克隆仓库，从 `macos-arm64` 分支开发。
3. Mac Codex 每完成一个阶段就提交一次，例如“跨平台路径”“Mac 源码运行通过”“arm64 打包通过”。
4. Windows 端拉取该分支，审查业务脚本是否被意外修改。
5. 通过相同测试数据后合并到 `main`。

严禁把真实买家信息、真实订单 CSV、云途面单或包含地址的输出文件提交到 GitHub，即使仓库是私有的。仓库应添加 `.gitignore`：

```gitignore
.venv/
__pycache__/
build/
dist/
release*/
*.spec
*.exe
*.dmg
*.zip
*.csv
*.xlsx
*.xls
*.pdf
config.json
.DS_Store
```

测试文件应使用人工构造或脱敏数据，并单独确认不包含个人信息。

### 简单方案：源码 ZIP 手动传递

如果暂时不使用 GitHub，可以将源码和本交接文档压缩后传到 Mac。Mac 完成后再把“修改后的源码 ZIP + 测试报告”传回 Windows。

该方案操作简单，但无法清楚查看每次改了什么，也容易出现 Windows 和 Mac 源码版本不一致，只适合第一次移植。

### 不建议的方式

- 让两个 Codex 账号只依靠口头描述同步修改。
- 用真实订单文件作为公开协作样本。
- Windows 和 Mac 同时直接修改同一份无版本控制的源码。
- 将 `.exe` 发送到 Mac 后尝试转换为 `.app`。

## 13. 后续正式分发

当前没有 Apple Developer 账户，可以完成开发、本机测试和临时签名。

如果未来要把应用发给其他 Mac 用户，建议再完成：

- 加入 Apple Developer Program。
- 申请 Developer ID Application 证书。
- 启用 Hardened Runtime。
- 对 `.app` 或 `.dmg` 进行正式签名。
- 提交 Apple Notary Service 公证并将票据 stapling 到交付包。
- 在一台未安装开发环境的 Mac 上测试首次安装。

在没有开发者账户前，不需要为第一版本阻塞开发，但交付物应明确标记为“本机测试版”。
