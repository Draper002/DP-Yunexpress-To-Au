# Web 部署命令

对应文件：`DP-International-Fulfillment-Web_20260730_2133.zip`

适用目录：`/www/wwwroot/dropshipzone`

先通过服务器面板把 ZIP 上传到上述目录，再逐行复制执行：

```bash
cd /www/wwwroot/dropshipzone || exit 1

PACKAGE="DP-International-Fulfillment-Web_20260730_2133.zip"
EXPECTED_SHA256="D3DB6A7BD0FEE66F0FF5FBC959310E8462E62BB951A0E3545D85BFA51BD1D7E1"

test -f "$PACKAGE" || { echo "未找到 $PACKAGE，请先上传到当前目录"; exit 1; }
ACTUAL_SHA256="$(sha256sum "$PACKAGE" | awk '{print toupper($1)}')"
test "$ACTUAL_SHA256" = "$EXPECTED_SHA256" || { echo "ZIP 哈希不一致：$ACTUAL_SHA256"; exit 1; }

unzip -oq "$PACKAGE" -d .
test -f shared/fulfillment/order_consolidation.py || { echo "合单模块未解压成功"; exit 1; }
grep -q 'same_recipient_same_sku_merge' web/app.py || { echo "Web 新功能代码未生效"; exit 1; }
test -f deployment/web/.env || { echo "缺少 deployment/web/.env，请恢复原部署配置后再继续"; exit 1; }

docker compose -f deployment/web/docker-compose.yml config >/dev/null || exit 1
docker compose -f deployment/web/docker-compose.yml up -d --build --force-recreate || exit 1
sleep 15

docker compose -f deployment/web/docker-compose.yml ps
curl -fsS http://127.0.0.1:8000/healthz && echo
curl -fsS https://dpz.zylgzx.cn/healthz && echo
docker compose -f deployment/web/docker-compose.yml logs --tail=80
```

健康接口应包含：

```json
{"status":"ok","platforms":["yunexpress","sf"],"features":["same_recipient_same_sku_merge"]}
```

以上命令不会删除 Docker 数据卷。不要执行 `docker compose down -v`，否则可能删除账号、固定配置和历史结果。
