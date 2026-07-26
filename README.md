# DP International Fulfillment To Au

DropShipZone 澳大利亚订单发货工具，支持 YunExpress 云途和 SF International 顺丰国际，并包含 Windows 桌面版、macOS Apple Silicon 版和云服务器网页版本。

## 代码目录

```text
desktop/
  windows/              Windows 桌面程序入口和打包配置
  macos/                Mac 入口、适配说明和 Mac 打包文件
web/                    Web 网页程序入口、依赖和网页进度
shared/
  fulfillment/          三个平台共用的订单、运单和 SKU 业务逻辑
docs/                   总体需求和开发进度文档
```

## 三个平台的职责

| 目录 | 作用 | 当前状态 |
|---|---|---|
| `desktop/windows/` | Windows `.exe` 桌面应用 | 已完成可测试版 |
| `desktop/macos/` | M5 MacBook / macOS 26 应用 | 双平台源码与 arm64 打包配置已完成，待 Mac 实机构建 |
| `web/` | 云服务器多账号网页应用 | 云途与顺丰三步流程已完成真实文件回归 |
| `shared/fulfillment/` | 三个平台共用的核心处理逻辑 | 云途、顺丰上传、DP 运单回填及两种面单整理均已验证 |

## 重要规则

业务脚本只维护 `shared/fulfillment/` 这一份。Windows、Mac 和 Web 只负责界面、文件上传下载、账号权限和平台适配。真实订单、买家信息、模板、面单和输出文件禁止提交到 Git。

## 支持的流程

| 步骤 | 云途 | 顺丰国际 |
|---|---|---|
| 1 | DP CSV -> 云途批量寄件 Excel | DP CSV -> 顺丰批量导入 XLSM |
| 2 | 云途订单信息 -> DP 回填模板 | 顺丰订单 `.xls/.xlsx` -> DP 回填模板，承运商为 `SF INTERNATIONAL` |
| 3 | 面单 ZIP -> 按 SKU 的供应商文件夹与 ZIP | 合并面单 PDF -> 按寄方姓名拆分的 PDF |

## 文档入口

- Windows：[desktop/windows/README.md](desktop/windows/README.md)
- Mac：[desktop/macos/README.md](desktop/macos/README.md)
- Mac 交接：[desktop/macos/MACOS_M5_HANDOFF_ZH.md](desktop/macos/MACOS_M5_HANDOFF_ZH.md)
- Web：[web/README.md](web/README.md)
- Web 进度：[web/docs/WEB_DEVELOPMENT_PROGRESS.md](web/docs/WEB_DEVELOPMENT_PROGRESS.md)
- 总体进度：[docs/DEVELOPMENT_PROGRESS.md](docs/DEVELOPMENT_PROGRESS.md)
