# Web 部署包说明

这个目录提供 Docker 部署配置。完整部署包还必须包含仓库根目录的 `web/` 和 `shared/` 两个目录，不能只上传本目录。

构建上下文使用项目根目录，因此有效的 Docker 忽略规则位于项目根目录 `.dockerignore`；它会排除业务文件、构建产物和本地数据目录。

## 服务器要求

- Linux 云服务器
- Docker Engine
- Docker Compose Plugin
- 一个已经解析到服务器 IP 的域名

## 部署步骤

1. 将部署压缩包上传并解压到服务器，例如 `/opt/dp-international-fulfillment-to-au`。
2. 进入解压后的项目根目录，复制环境变量文件：

```bash
cp deployment/web/.env.example deployment/web/.env
```

3. 编辑 `deployment/web/.env`，务必修改管理员邮箱和管理员密码。安全模式会拒绝示例密码启动；该密码只在首次创建数据库时生效：

```bash
nano deployment/web/.env
```

4. 构建并启动：

```bash
docker compose -f deployment/web/docker-compose.yml up -d --build
```

当前版本依靠单个 Uvicorn 进程串行保护批次文件和共享脚本参数，请保持 Dockerfile 默认启动命令，不要额外增加 `--workers`。

5. 检查容器状态和健康接口：

```bash
docker compose -f deployment/web/docker-compose.yml ps
curl http://127.0.0.1:8000/healthz
```

6. 在 Nginx、Caddy 或云服务器反向代理中，将域名转发到：

```text
http://127.0.0.1:8000
```

必须使用 HTTPS。当前 Compose 已设置 `DP_WEB_COOKIE_SECURE=1`，因此登录 Cookie 只会在 HTTPS 下发送。

反向代理需允许至少 80 MB 上传；建议设置为 100 MB，并把请求超时设置为至少 300 秒，以覆盖较大的顺丰合并面单 PDF。

## 数据保存位置

Docker 数据卷 `dp_yunexpress_data` 保存：

- 管理员和员工账号
- 4 份固定配置：SKU 商品库、云途模板、顺丰国际模板、DP 发货模板
- 处理结果
- 校验记录

订单 CSV、物流订单 Excel、云途面单 ZIP 和顺丰面单 PDF 只在对应账号、对应平台的批次目录中保存。第 2 步自动使用第 1 步 DP CSV 校验订单集合；失败上传会删除，上一份已验证批次不会被覆盖。第 3 步成功后清理对应批次目录。首次部署建议先使用脱敏文件测试，不要直接上传真实订单。

## 平台流程

- 云途：DP CSV -> 云途批量模板 -> DP 运单回填 -> 按 SKU 分拣供应商面单包。
- 顺丰国际：DP CSV -> 顺丰批量模板 -> DP 运单回填 -> 按寄方姓名拆分合并面单 PDF。
- 顺丰 DP 回填的承运商固定为 `SF INTERNATIONAL`。
- 第 2 步不要求再次上传 DP CSV，但必须先由同一账号、同一承运平台成功完成第 1 步。
- 每个员工只看到自己的历史记录，也不能下载其他员工的结果；管理员保留全局下载权限。

## 当前版本边界

这是可部署测试版，不是最终生产版。正式使用前还需要补充登录限流、CSRF 防护、后台任务队列、结果过期清理、数据库备份和对象存储。
