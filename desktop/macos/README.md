# macOS / Apple Silicon

这里存放 M 系列 Mac 的运行入口、依赖、打包配置和发布记录。

业务处理逻辑统一使用 `../../shared/fulfillment/`，不要在 Mac 目录复制一份业务脚本。这样 Windows、Mac 和 Web 三个平台的订单字段映射、运单匹配和 SKU 面单分拣规则保持一致。

## 当前功能

Mac 入口与 Windows 共用三步业务规则，并支持云途和顺丰国际：

1. DP 订单生成云途或顺丰国际上传模板。
2. 云途或顺丰订单信息生成 DP 运单回填文件。
3. 云途面单按 SKU 分拣；顺丰合并 PDF 按寄方姓名拆分。

Mac 专用代码只负责 Finder 打开目录和 Apple Silicon 打包，不复制字段映射。

## 从源码运行

```bash
cd desktop/macos
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

## 构建 M 芯片应用

```bash
cd desktop/macos
chmod +x build_macos_arm64.sh
./build_macos_arm64.sh
```

输出：`dist/DP International Fulfillment To Au.app`

当前交接文档：[`MACOS_M5_HANDOFF_ZH.md`](MACOS_M5_HANDOFF_ZH.md)。

由于当前没有 Apple Developer 账户，需先在 M5 MacBook 上构建、临时签名并本机验证，再决定是否制作未公证的测试 DMG。
