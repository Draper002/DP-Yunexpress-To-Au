# DP International Fulfillment To Au Web

云服务器多账号网页版本。支持云途和顺丰国际，并统一复用仓库根目录 `shared/fulfillment/` 的业务脚本。

## 当前流程

| 步骤 | 云途 | 顺丰国际 |
|---|---|---|
| 1 | DP 订单 CSV 生成云途批量寄件 Excel | DP 订单 CSV 生成顺丰批量导入 XLSM，可选业务类型和是否带电 |
| 2 | 云途订单信息生成 DP 回填模板，并核对第 1 步订单集合 | 顺丰订单 `.xls/.xlsx` 生成 DP 回填模板并核对批次，承运商固定为 `SF INTERNATIONAL` |
| 3 | 云途面单 ZIP 按 SKU 分拣 | 顺丰合并面单 PDF 按寄方姓名拆分 |

## 本地启动

```bash
cd web
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export DP_WEB_ADMIN_EMAIL="admin@example.com"
export DP_WEB_ADMIN_PASSWORD="change-this-before-deploy"
export DP_WEB_DATA_DIR="$(pwd)/data"
export DP_WEB_COOKIE_SECURE="0"
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

Windows PowerShell 使用 `$env:NAME="value"` 设置环境变量。首次登录后，管理员可创建员工账号并维护 4 份固定配置：SKU 商品库、云途标准模板、顺丰国际标准模板和 DP 发货模板。固定配置支持只替换其中一份。

## 数据策略

本版本使用 SQLite 和磁盘保存账号、处理结果及校验记录。各账号的处理记录、工作目录和下载权限相互隔离；管理员可以下载全部结果。第 2 步自动复用第 1 步 DP CSV，用户无需重复上传；订单集合不一致时停止生成，失败输入不会覆盖上一份已验证的批次文件。云途或顺丰第 3 步成功后清理对应批次目录。服务器仍需配置历史结果和未完成批次的定期清理策略。

正式公网部署必须启用 HTTPS、反向代理、持久化数据目录和备份。健康检查地址为 `/healthz`。

## 验证

```bash
python -m unittest discover -s tests -v
```

2026-07-26 已使用真实结构文件分别完成云途和顺丰国际 Web 回归，并验证错批次订单会被拦截；19 项自动化测试通过。测试运行目录位于本地 `artifacts/`，不进入 Git 或部署包。
