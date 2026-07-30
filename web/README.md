# DP International Fulfillment To Au Web

云服务器多账号网页版本。支持云途和顺丰国际，并统一复用仓库根目录 `shared/fulfillment/` 的业务脚本。

## 当前流程

| 步骤 | 云途 | 顺丰国际 |
|---|---|---|
| 1 | DP 订单 CSV 生成云途批量寄件 Excel | DP 订单 CSV 生成顺丰批量导入 XLSM，可选业务类型和是否带电 |
| 2 | 云途订单信息生成 DP 回填模板，并核对第 1 步订单集合 | 顺丰订单 `.xls/.xlsx` 生成 DP 回填模板并核对批次，承运商固定为 `SF INTERNATIONAL` |
| 3 | 云途面单 ZIP 按 SKU 分拣 | 顺丰合并面单与第 2 步运单逐单核对后，按寄方姓名拆分 |

第 1 步默认勾选“合并相同收件人 + 相同 SKU”。检测到合并组时，下载 ZIP 同时包含承运商上传模板和订单合并关系表。第 2 步自动复用服务器保存的合单计划，把同一运单号写入组内全部 DP 订单，并附合并运单对应表。第 3 步把合并面单单独整理到“合并件”供应商包；不同 SKU 不合并。

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

本版本使用 SQLite 和磁盘保存账号、处理结果及校验记录。各账号的处理记录、工作目录和下载权限相互隔离；管理员可以下载全部结果。第 2 步自动复用第 1 步 DP CSV 和合单计划；顺丰第 3 步自动复用第 2 步顺丰订单数据。用户无需重复上传，订单、计划、日期、平台或运单集合不一致时停止生成，失败输入不会覆盖上一份已验证的批次文件。云途或顺丰第 3 步成功后清理对应批次目录。服务器仍需配置历史结果和未完成批次的定期清理策略。

正式公网部署必须启用 HTTPS、反向代理、持久化数据目录和备份。健康检查地址为 `/healthz`。

## 验证

```bash
python -m unittest discover -s tests -v
```

2026-07-29 已增加合单、同运单多订单回填、云途合并件供应商包及顺丰合并件分流测试；30 项自动化测试通过。DP 后台是否接受多个订单共用同一运单号仍需用小批量真实订单完成最终平台验收。测试运行目录位于本地 `artifacts/`，不进入 Git 或部署包。
