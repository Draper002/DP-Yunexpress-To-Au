# DP&Yunexpress To Au

DropShipZone -> YunExpress 澳大利亚发货工具，包含 Windows 桌面版、macOS Apple Silicon 版和云服务器网页版本。

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
| `desktop/macos/` | M5 MacBook / macOS 26 应用 | 开发中 |
| `web/` | 云服务器多账号网页应用 | MVP 开发中 |
| `shared/fulfillment/` | 三个平台共用的核心处理逻辑 | 已验证 |

## 重要规则

业务脚本只维护 `shared/fulfillment/` 这一份。Windows、Mac 和 Web 只负责界面、文件上传下载、账号权限和平台适配。真实订单、买家信息、模板、面单和输出文件禁止提交到 Git。

## 文档入口

- Windows：[desktop/windows/README.md](desktop/windows/README.md)
- Mac：[desktop/macos/README.md](desktop/macos/README.md)
- Mac 交接：[desktop/macos/MACOS_M5_HANDOFF_ZH.md](desktop/macos/MACOS_M5_HANDOFF_ZH.md)
- Web：[web/README.md](web/README.md)
- Web 进度：[web/docs/WEB_DEVELOPMENT_PROGRESS.md](web/docs/WEB_DEVELOPMENT_PROGRESS.md)
- 总体进度：[docs/DEVELOPMENT_PROGRESS.md](docs/DEVELOPMENT_PROGRESS.md)


