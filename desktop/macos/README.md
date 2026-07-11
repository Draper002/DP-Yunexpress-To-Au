# macOS / Apple Silicon

这里存放 Mac 版本的适配说明、Mac 专用打包配置和发布记录。

业务处理逻辑统一使用 `../../shared/fulfillment/`，不要在 Mac 目录复制一份业务脚本。这样 Windows、Mac 和 Web 三个平台的订单字段映射、运单匹配和 SKU 面单分拣规则保持一致。

当前交接文档：[`MACOS_M5_HANDOFF_ZH.md`](MACOS_M5_HANDOFF_ZH.md)。

M5 版本完成后，Mac 专用入口和打包文件放在本目录，通用业务代码仍放在 `shared/fulfillment/`。


