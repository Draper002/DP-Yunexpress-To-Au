# DP&Yunexpress To Au Web

云服务器多账号网页版本。复用上级目录 `DP_Yunexpress_To_Au_Tool/scripts` 的三个业务脚本。

## 本地启动

```bash
cd webapp
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export DP_WEB_SECRET="replace-with-a-long-random-secret"
export DP_WEB_ADMIN_EMAIL="admin@example.com"
export DP_WEB_ADMIN_PASSWORD="change-this-before-deploy"
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

Windows PowerShell 使用 `$env:NAME="value"` 设置环境变量。首次登录后，管理员可创建员工账号并替换 3 份固定配置。

## 数据策略

本 MVP 使用 SQLite 和磁盘保存账号、处理结果及校验记录。订单 CSV、云途订单 Excel 和面单 ZIP 写入批次临时目录，第三步成功或失败后删除。正式公网部署必须启用 HTTPS、反向代理、持久化数据目录和定期清理任务。


