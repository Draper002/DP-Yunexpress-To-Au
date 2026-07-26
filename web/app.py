from __future__ import annotations

import asyncio
import hashlib
import hmac
import html
import io
import os
import secrets
import shutil
import sqlite3
import sys
import threading
import time
from contextlib import asynccontextmanager, closing, redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path

import openpyxl
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from starlette.concurrency import run_in_threadpool

ROOT = Path(__file__).resolve().parents[1]
TOOL_ROOT = ROOT / "shared" / "fulfillment"
DATA = Path(os.getenv("DP_WEB_DATA_DIR", str(Path(__file__).parent / "data"))).resolve()
CONFIG = DATA / "config"
RESULTS = DATA / "results"
WORKING = DATA / "working"
DB_PATH = DATA / "app.sqlite3"
ADMIN_EMAIL = os.getenv("DP_WEB_ADMIN_EMAIL", "admin@example.com")
ADMIN_PASSWORD = os.getenv("DP_WEB_ADMIN_PASSWORD", "change-me")
MAX_UPLOAD = 80 * 1024 * 1024
COOKIE_SECURE = os.getenv("DP_WEB_COOKIE_SECURE", "0") == "1"
UNSAFE_ADMIN_PASSWORDS = {"change-me", "replace-before-starting"}
sys.path.insert(0, str(TOOL_ROOT))
import generate_dp_shipment_upload
import generate_sf_international_template
import generate_yunexpress_template
import sort_yunexpress_labels_by_sku
import split_sf_labels_by_sender

CONFIG_FILES = {
    "sku": "sku.xlsx",
    "yun_template": "yun_template.xlsx",
    "sf_template": "sf_template.xlsm",
    "dp_template": "dp_template.xlsx",
}
CONFIG_LABELS = {
    "sku": "SKU 商品库",
    "yun_template": "云途标准模板",
    "sf_template": "顺丰国际标准模板",
    "dp_template": "DP 发货模板",
}
PLATFORM_LABELS = {"yunexpress": "云途", "sf": "顺丰国际"}
SCRIPT_LOCK = threading.Lock()
WORKFLOW_LOCK = asyncio.Lock()
YUN_TEMPLATE_HEADERS = {
    "客户订单号",
    "产品代码",
    "收件人国家",
    "收件人姓名",
    "收件人地址",
    "收件人城市",
    "收件人省州",
    "邮编",
    "电话",
    "件数",
    "包裹总重量",
    "申报币种",
    "SKU1",
    "申报品名1",
    "中文申报品名1",
    "申报数量1",
    "申报FOB价1",
    "申报单重1",
}


@asynccontextmanager
async def lifespan(_app: FastAPI):
    if COOKIE_SECURE and ADMIN_PASSWORD in UNSAFE_ADMIN_PASSWORDS:
        raise RuntimeError("Set a private DP_WEB_ADMIN_PASSWORD before secure deployment.")
    init_db()
    for folder in (CONFIG, RESULTS, WORKING):
        folder.mkdir(parents=True, exist_ok=True)
    yield


app = FastAPI(title="DP International Fulfillment To Au Web", lifespan=lifespan)


def connection():
    DATA.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or secrets.token_bytes(16)
    value = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 240_000)
    return f"{salt.hex()}${value.hex()}"


def check_password(password: str, encoded: str) -> bool:
    try:
        salt, expected = encoded.split("$", 1)
        actual = hash_password(password, bytes.fromhex(salt)).split("$", 1)[1]
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False


def init_db():
    DATA.mkdir(parents=True, exist_ok=True)
    with closing(connection()) as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,email TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,role TEXT NOT NULL,active INTEGER NOT NULL DEFAULT 1,created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,user_id INTEGER NOT NULL,expires_at INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS runs(id INTEGER PRIMARY KEY,user_id INTEGER NOT NULL,step TEXT NOT NULL,status TEXT NOT NULL,date TEXT NOT NULL,result_path TEXT,report_path TEXT,created_at TEXT NOT NULL,error TEXT);
        """)
        if not conn.execute("SELECT 1 FROM users LIMIT 1").fetchone():
            conn.execute("INSERT INTO users(email,password_hash,role,created_at) VALUES(?,?,?,?)", (ADMIN_EMAIL, hash_password(ADMIN_PASSWORD), "admin", datetime.now().isoformat(timespec="seconds")))
        conn.execute("DELETE FROM sessions WHERE expires_at<=?", (int(time.time()),))
        conn.commit()


def user_for(request: Request):
    token = request.cookies.get("dp_session")
    if not token:
        return None
    with closing(connection()) as conn:
        return conn.execute("SELECT users.* FROM sessions JOIN users ON users.id=sessions.user_id WHERE token=? AND expires_at>? AND active=1", (token, int(time.time()))).fetchone()


async def save_upload(upload: UploadFile, folder: Path, extensions: set[str]) -> Path:
    suffix = Path(upload.filename or "").suffix.lower()
    if suffix not in extensions:
        raise ValueError("文件类型不正确")
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"{secrets.token_hex(10)}{suffix}"
    size = 0
    with target.open("wb") as out:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD:
                target.unlink(missing_ok=True)
                raise ValueError("文件超过 80 MB 限制")
            out.write(chunk)
    return target


def config_path(key: str) -> Path:
    if key not in CONFIG_FILES:
        raise KeyError(f"Unknown config key: {key}")
    return CONFIG / CONFIG_FILES[key]


def validate_config_upload(key: str, path: Path) -> None:
    if key == "sku":
        if not generate_yunexpress_template.read_sku_info(path):
            raise ValueError("SKU 商品库没有可用的 SKU 数据")
        return

    workbook = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=False,
        keep_vba=key == "sf_template",
    )
    try:
        if not workbook.worksheets:
            raise ValueError(f"{CONFIG_LABELS[key]}没有工作表")
        sheet = (
            workbook["标准上传格式"]
            if key == "sf_template" and "标准上传格式" in workbook.sheetnames
            else workbook.worksheets[0]
        )
        headers = {
            generate_yunexpress_template.clean(sheet.cell(1, column).value)
            for column in range(1, sheet.max_column + 1)
        }
        if key == "yun_template":
            missing = sorted(YUN_TEMPLATE_HEADERS - headers)
            if missing:
                raise ValueError("云途标准模板缺少字段：" + "、".join(missing))
        elif key == "sf_template":
            missing = sorted(set(generate_sf_international_template.HEADER.values()) - headers)
            if missing:
                raise ValueError("顺丰国际标准模板缺少字段：" + "、".join(missing))
    finally:
        workbook.close()


def normalize_platform(value: str) -> str:
    platform = (value or "").strip().lower()
    if platform not in PLATFORM_LABELS:
        raise ValueError("请选择云途或顺丰国际")
    return platform


def normalize_date(value: str) -> str:
    try:
        parsed = datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("发货日期格式必须为 YYYY-MM-DD") from exc
    if parsed.strftime("%Y-%m-%d") != value:
        raise ValueError("发货日期格式必须为 YYYY-MM-DD")
    return value


def required_config_keys(step: int, platform: str) -> tuple[str, ...]:
    platform = normalize_platform(platform)
    requirements = {
        (1, "yunexpress"): ("sku", "yun_template"),
        (1, "sf"): ("sku", "sf_template"),
        (2, "yunexpress"): ("dp_template",),
        (2, "sf"): ("dp_template",),
        (3, "yunexpress"): ("sku",),
        (3, "sf"): (),
    }
    try:
        return requirements[(step, platform)]
    except KeyError as exc:
        raise ValueError("未知处理步骤") from exc


def add_run(user_id: int, step: str, date: str, status: str, result: Path | None = None, report: Path | None = None, error: str = ""):
    with closing(connection()) as conn:
        conn.execute("INSERT INTO runs(user_id,step,status,date,result_path,report_path,created_at,error) VALUES(?,?,?,?,?,?,?,?)", (user_id, step, status, date, str(result) if result else None, str(report) if report else None, datetime.now().isoformat(timespec="seconds"), error[:3000]))
        conn.commit()


def script(module, args: list[str]) -> tuple[int, str, str]:
    with SCRIPT_LOCK:
        old = sys.argv[:]
        stdout, stderr = io.StringIO(), io.StringIO()
        try:
            sys.argv = [module.__name__] + args
            with redirect_stdout(stdout), redirect_stderr(stderr):
                code = module.main()
            return int(code or 0), stdout.getvalue(), stderr.getvalue()
        finally:
            sys.argv = old


def result_from(output: str) -> Path:
    paths = [line[3:].strip().strip('"') for line in output.splitlines() if line.startswith("OK:")]
    if not paths or not Path(paths[-1]).exists():
        raise ValueError(output.strip() or "处理脚本没有返回结果")
    return Path(paths[-1])


def batch_dir(user_id: int, platform: str) -> Path:
    path = WORKING / str(user_id) / normalize_platform(platform)
    path.mkdir(parents=True, exist_ok=True)
    return path


def field(name: str, label: str, accept: str, required: bool = True) -> str:
    required_attr = " required" if required else ""
    return f'<label class="file"><b>{label}</b><input name="{name}" type="file" accept="{accept}"{required_attr}><span>选择文件</span></label>'


def platform_selector() -> str:
    return '''<div class="platform-picker" role="group" aria-label="承运平台">
    <label><input type="radio" name="platform" value="yunexpress" checked><span>云途</span></label>
    <label><input type="radio" name="platform" value="sf"><span>顺丰国际</span></label>
    </div>'''


def sf_options() -> str:
    business_options = "".join(
        f'<option value="{html.escape(value)}">{html.escape(value)}</option>'
        for value in generate_sf_international_template.SF_BUSINESS_TYPES
    )
    battery_options = "".join(
        f'<label><input type="radio" name="battery" value="{html.escape(value)}"{(" checked" if value == "否" else "")}><span>{html.escape(value)}</span></label>'
        for value in generate_sf_international_template.BATTERY_OPTIONS
    )
    return f'''<div class="sf-only sf-options">
    <label>业务类型<select name="business_type"><option value="">请选择业务类型</option>{business_options}</select></label>
    <div><b>是否带电</b><div class="mini-segment">{battery_options}</div></div>
    </div>'''


def page(title: str, content: str, user=None) -> str:
    nav = f'<header><a class="brand" href="/">DP 国际发货 <small>TO AU</small></a><nav><a href="/">工作台</a>{"<a href=/admin>成员管理</a>" if user and user["role"] == "admin" else ""}<a href="/logout">退出</a></nav></header>' if user else '<header><a class="brand" href="/login">DP 国际发货 <small>TO AU</small></a></header>'
    return f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)} · DP International Fulfillment</title><style>{CSS}</style></head><body>{nav}<main>{content}</main><script>{SCRIPT}</script></body></html>'


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    if user_for(request):
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(page("登录", '<section class="login"><div class="mark">DP<span>&</span></div><p class="eyebrow">PRIVATE OPERATIONS</p><h1>登录发货工作台</h1><p class="lead">员工账号由管理员创建，订单文件和处理记录按账号隔离保存。</p><form method="post" action="/login"><label>邮箱<input name="email" type="email" required></label><label>密码<input name="password" type="password" required></label><button class="primary" type="submit">登录</button></form></section>'))


@app.post("/login")
def login(email: str = Form(...), password: str = Form(...)):
    with closing(connection()) as conn:
        row = conn.execute("SELECT * FROM users WHERE lower(email)=lower(?) AND active=1", (email.strip(),)).fetchone()
    if not row or not check_password(password, row["password_hash"]):
        return HTMLResponse(page("登录失败", '<section class="login"><h1>无法登录</h1><p class="lead">邮箱或密码不正确。</p><a class="button" href="/login">返回登录</a></section>'), status_code=401)
    token = secrets.token_urlsafe(32)
    with closing(connection()) as conn:
        conn.execute("INSERT INTO sessions VALUES(?,?,?)", (token, row["id"], int(time.time()) + 86400)); conn.commit()
    response = RedirectResponse("/", status_code=303)
    response.set_cookie("dp_session", token, httponly=True, samesite="lax", secure=COOKIE_SECURE, max_age=86400)
    return response


@app.get("/logout")
def logout(request: Request):
    token = request.cookies.get("dp_session")
    if token:
        with closing(connection()) as conn:
            conn.execute("DELETE FROM sessions WHERE token=?", (token,)); conn.commit()
    response = RedirectResponse("/login", status_code=303); response.delete_cookie("dp_session"); return response


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    user = user_for(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    with closing(connection()) as conn:
        runs = conn.execute(
            "SELECT * FROM runs WHERE user_id=? ORDER BY id DESC LIMIT 10",
            (user["id"],),
        ).fetchall()
    records = ""
    for record in runs:
        result_link = f'<a href="/download/{record["id"]}">下载结果</a>' if record["result_path"] and Path(record["result_path"]).exists() else "-"
        records += f'<tr><td>{html.escape(record["created_at"])}</td><td>{html.escape(record["step"])}</td><td><span class="status {record["status"].lower()}">{html.escape(record["status"])}</span></td><td>{html.escape(record["date"])}</td><td>{result_link}</td></tr>'
    records = records or '<tr><td colspan="5" class="empty">还没有处理记录，从第 1 步开始。</td></tr>'
    today = datetime.now().date().isoformat()
    config_ready = all(config_path(key).exists() for key in CONFIG_LABELS)
    content = f'''<section class="hero"><div><p class="eyebrow">AU FULFILLMENT DESK</p><h1>发货工作台</h1><p class="lead">选择云途或顺丰国际，按顺序完成寄件模板、DP 运单号回填和供应商面单整理。</p></div><div class="identity"><i></i><div><b>{html.escape(user["email"])}</b><small>{"管理员" if user["role"] == "admin" else "员工"} · 在线</small></div></div></section>
    <section class="config"><div><small>固定配置</small><b>{"全部就绪" if config_ready else "需要检查"}</b><em>SKU 商品库、两家承运商模板与 DP 发货模板，仅更新时维护</em></div><a class="button soft" href="/config">管理配置</a></section>
    <section class="steps">
      <article data-platform="yunexpress"><div class="step-head"><div class="no">01</div><span>寄件导入</span></div><h2><span class="yun-only">生成云途模板</span><span class="sf-only">生成顺丰模板</span></h2><p><span class="yun-only">DP 订单转换为云途批量寄件 Excel。</span><span class="sf-only">DP 订单转换为顺丰国际批量导入工作簿。</span></p><form class="workflow-form" data-platform="yunexpress" action="/step/1" method="post" enctype="multipart/form-data">{platform_selector()}{field("file","本次 DP 订单 CSV",".csv")}{sf_options()}<div class="row"><input name="date" type="date" value="{today}" required><button class="primary" type="submit">开始生成</button></div></form></article>
      <article data-platform="yunexpress"><div class="step-head"><div class="no">02</div><span>DP 回填</span></div><h2>生成 DP 回填模板</h2><p><span class="yun-only">上传云途订单信息 Excel，批量回填运单号。</span><span class="sf-only">上传顺丰订单数据，承运商自动填写 SF INTERNATIONAL。</span></p><form class="workflow-form" data-platform="yunexpress" action="/step/2" method="post" enctype="multipart/form-data">{platform_selector()}<label class="file"><b><span class="yun-only">本次云途订单信息 Excel</span><span class="sf-only">本次顺丰国际订单数据</span></b><input name="file" type="file" accept=".xls,.xlsx" required><span>选择文件</span></label><div class="row"><input name="date" type="date" value="{today}" required><button class="primary" type="submit">开始生成</button></div></form></article>
      <article data-platform="yunexpress"><div class="step-head"><div class="no">03</div><span>面单整理</span></div><h2><span class="yun-only">按 SKU 分拣</span><span class="sf-only">按寄方姓名拆分</span></h2><p><span class="yun-only">上传云途面单 ZIP，生成按 SKU 分组的供应商包。</span><span class="sf-only">上传顺丰面单 PDF，生成每个寄方姓名对应的 PDF。</span></p><form class="workflow-form" data-platform="yunexpress" action="/step/3" method="post" enctype="multipart/form-data">{platform_selector()}<label class="file"><b><span class="yun-only">本次云途面单 ZIP</span><span class="sf-only">本次顺丰国际面单 PDF</span></b><input name="file" type="file" accept=".zip,.pdf" required><span>选择文件</span></label><div class="row"><input name="date" type="date" value="{today}" required><button class="primary" type="submit">开始整理</button></div></form></article>
    </section><section class="history"><div class="section-head"><div><p class="eyebrow">ACTIVITY</p><h2>我的最近处理记录</h2></div><span>每个账号仅显示自己的批次</span></div><div class="table-wrap"><table><thead><tr><th>时间</th><th>步骤</th><th>状态</th><th>日期</th><th>结果</th></tr></thead><tbody>{records}</tbody></table></div></section>'''
    return HTMLResponse(page("工作台", content, user))


@app.get("/config", response_class=HTMLResponse)
def config_page(request: Request):
    user = user_for(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    status = "".join(f'<li><i></i><div><b>{label}</b><small>{"已上传，可替换" if config_path(key).exists() else "尚未上传"}</small></div></li>' for key, label in CONFIG_LABELS.items())
    edit_form = f'''<form action="/config" method="post" enctype="multipart/form-data">
    {field("sku","SKU 商品库",".xlsx",False)}
    {field("yun_template","云途标准模板",".xlsx",False)}
    {field("sf_template","顺丰国际标准模板",".xlsm",False)}
    {field("dp_template","DP 发货模板",".xlsx",False)}
    <p class="hint">只选择本次需要更新的文件，其余配置保持不变。</p><button class="primary" type="submit">保存所选配置</button></form>'''
    if user["role"] != "admin":
        edit_form = '<div class="notice">固定配置由管理员维护。当前账号可以查看配置状态并执行发货任务。</div>'
    content = f'<section class="head"><p class="eyebrow">CONFIGURATION</p><h1>固定配置</h1><p class="lead">这些文件不需要每天上传，只在 SKU 数据或平台模板发生变化时替换。</p></section><section class="two"><div class="panel"><h2>更新配置</h2>{edit_form}</div><div class="panel quiet"><p class="eyebrow">CURRENT</p><h2>当前状态</h2><ul>{status}</ul><p class="hint">保存后，下一次任务立即使用新文件。</p></div></section>'
    return HTMLResponse(page("固定配置", content, user))


@app.post("/config")
async def save_config(
    request: Request,
    sku: UploadFile | None = File(None),
    yun_template: UploadFile | None = File(None),
    sf_template: UploadFile | None = File(None),
    dp_template: UploadFile | None …123 tokens truncated…文件", '<section class="login"><h1>没有选择配置文件</h1><p class="lead">请选择至少一个需要更新的文件。</p><a class="button" href="/config">返回固定配置</a></section>', user), status_code=400)
    allowed = {"sf_template": {".xlsm"}}
    staging = CONFIG / f".staging-{secrets.token_hex(8)}"
    staged: dict[str, Path] = {}
    try:
        for key, upload in selected:
            path = await save_upload(upload, staging, allowed.get(key, {".xlsx"}))
            await run_in_threadpool(validate_config_upload, key, path)
            staged[key] = path
        async with WORKFLOW_LOCK:
            for key, path in staged.items():
                path.replace(config_path(key))
    except Exception as exc:
        return HTMLResponse(
            page(
                "配置保存失败",
                '<section class="login"><h1>配置未更新</h1>'
                f'<p class="lead">{html.escape(str(exc))}</p>'
                '<a class="button" href="/config">返回固定配置</a></section>',
                user,
            ),
            status_code=400,
        )
    finally:
        shutil.rmtree(staging, ignore_errors=True)
    return RedirectResponse("/config", status_code=303)


@app.get("/admin", response_class=HTMLResponse)
def admin_page(request: Request):
    user = user_for(request)
    if not user or user["role"] != "admin": return HTMLResponse("无权限", status_code=403)
    with closing(connection()) as conn: users = conn.execute("SELECT email,role,active,created_at FROM users ORDER BY id").fetchall()
    rows = "".join(f'<tr><td>{html.escape(u["email"])}</td><td>{"管理员" if u["role"] == "admin" else "员工"}</td><td>{"正常" if u["active"] else "停用"}</td><td>{u["created_at"]}</td></tr>' for u in users)
    content = f'<section class="head"><p class="eyebrow">TEAM ACCESS</p><h1>成员管理</h1><p class="lead">管理员创建员工账号，员工只能处理订单和下载结果。</p></section><section class="two"><div class="panel"><h2>创建员工账号</h2><form action="/admin/users" method="post"><label>员工邮箱<input name="email" type="email" required></label><label>初始密码<input name="password" type="password" minlength="8" required></label><button class="primary" type="submit">创建员工</button></form></div><div class="panel"><h2>账号列表</h2><table><thead><tr><th>邮箱</th><th>角色</th><th>状态</th><th>创建时间</th></tr></thead><tbody>{rows}</tbody></table></div></section>'
    return HTMLResponse(page("成员管理", content, user))


@app.post("/admin/users")
def create_user(request: Request, email: str = Form(...), password: str = Form(...)):
    user = user_for(request)
    if not user or user["role"] != "admin": return HTMLResponse("无权限", status_code=403)
    if len(password) < 8: return HTMLResponse("密码至少 8 位", status_code=400)
    try:
        with closing(connection()) as conn:
            conn.execute("INSERT INTO users(email,password_hash,role,created_at) VALUES(?,?,?,?)", (email.strip(), hash_password(password), "employee", datetime.now().isoformat(timespec="seconds"))); conn.commit()
    except sqlite3.IntegrityError: return HTMLResponse("邮箱已存在", status_code=400)
    return RedirectResponse("/admin", status_code=303)


def execute_workflow(
    user_id: int,
    step: int,
    date: str,
    source: Path,
    platform: str,
    business_type: str,
    battery: str,
) -> tuple[Path, Path | None]:
    batch = batch_dir(user_id, platform)
    output = RESULTS / str(user_id) / date.replace("-", "") / platform
    output.mkdir(parents=True, exist_ok=True)
    commit_target: Path | None = None

    if step == 1:
        commit_target = batch / "dp_orders.csv"
        if platform == "yunexpress":
            code, out, err = script(
                generate_yunexpress_template,
                [
                    "--dp-orders-csv", str(source),
                    "--sku-xlsx", str(config_path("sku")),
                    "--yunexpress-template-xlsx", str(config_path("yun_template")),
                    "--date", date,
                    "--output-root", str(output),
                ],
            )
        else:
            if business_type not in generate_sf_international_template.SF_BUSINESS_TYPES:
                raise ValueError("请选择本批顺丰国际订单的业务类型")
            if battery not in generate_sf_international_template.BATTERY_OPTIONS:
                raise ValueError("顺丰带电选项必须为“是”或“否”")
            code, out, err = script(
                generate_sf_international_template,
                [
                    "--dp-orders-csv", str(source),
                    "--sku-xlsx", str(config_path("sku")),
                    "--sf-template-xlsm", str(config_path("sf_template")),
                    "--business-type", business_type,
                    "--battery", battery,
                    "--date", date,
                    "--output-root", str(output),
                ],
            )
    elif step == 2:
        dp_orders = batch / "dp_orders.csv"
        if not dp_orders.exists():
            raise ValueError("为防止回填错批次，请先在同一账号和承运平台完成第 1 步")
        if platform == "yunexpress":
            commit_target = batch / "yun_orders.xlsx"
            source_args = ["--yunexpress-xlsx", str(source)]
        else:
            commit_target = batch / f"sf_orders{source.suffix.lower()}"
            source_args = ["--sf-international-file", str(source)]
        code, out, err = script(
            generate_dp_shipment_upload,
            source_args
            + [
                "--dp-orders-csv", str(dp_orders),
                "--shipment-template-xlsx", str(config_path("dp_template")),
                "--date", date,
                "--output-root", str(output),
            ],
        )
    elif platform == "yunexpress":
        dp_orders = batch / "dp_orders.csv"
        yun_orders = batch / "yun_orders.xlsx"
        if not dp_orders.exists() or not yun_orders.exists():
            raise ValueError("云途面单分拣需要先在同一账号完成第 1、2 步")
        code, out, err = script(
            sort_yunexpress_labels_by_sku,
            [
                "--zip", str(source),
                "--yunexpress-xlsx", str(yun_orders),
                "--dp-orders-csv", str(dp_orders),
                "--sku-xlsx", str(config_path("sku")),
                "--date", date,
                "--output-root", str(output),
            ],
        )
    else:
        code, out, err = script(
            split_sf_labels_by_sender,
            [
                "--pdf", str(source),
                "--date", date,
                "--output-root", str(output),
            ],
        )

    if code != 0:
        raise ValueError(err.strip() or out.strip() or "处理失败")
    if commit_target:
        if platform == "sf" and step == 2:
            for old in batch.glob("sf_orders.xls*"):
                if old != commit_target:
                    old.unlink(missing_ok=True)
        source.replace(commit_target)
    result = result_from(out)
    report = result.with_suffix(".report.json") if result.is_file() else result / "校验成功报告.json"
    package = (
        result
        if result.is_file()
        else Path(
            shutil.make_archive(
                str(result),
                "zip",
                root_dir=result.parent,
                base_dir=result.name,
            )
        )
    )
    if step == 3:
        shutil.rmtree(batch, ignore_errors=True)
    return package, report if report.exists() else None


async def process(
    request: Request,
    step: int,
    date: str,
    upload: UploadFile,
    platform: str,
    business_type: str = "",
    battery: str = "否",
):
    user = user_for(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    source: Path | None = None
    try:
        platform = normalize_platform(platform)
        date = normalize_date(date)
        required = required_config_keys(step, platform)
        missing = [CONFIG_LABELS[key] for key in required if not config_path(key).exists()]
        if missing:
            raise ValueError("请先在固定配置中上传：" + "、".join(missing))

        batch = batch_dir(user["id"], platform)
        extensions = {
            (1, "yunexpress"): {".csv"},
            (1, "sf"): {".csv"},
            (2, "yunexpress"): {".xlsx"},
            (2, "sf"): {".xls", ".xlsx"},
            (3, "yunexpress"): {".zip"},
            (3, "sf"): {".pdf"},
        }
        source = await save_upload(upload, batch, extensions[(step, platform)])
        async with WORKFLOW_LOCK:
            package, report = await run_in_threadpool(
                execute_workflow,
                user["id"],
                step,
                date,
                source,
                platform,
                business_type,
                battery,
            )
        step_name = f"步骤 {step} · {PLATFORM_LABELS[platform]}"
        add_run(user["id"], step_name, date, "SUCCESS", package, report)
        return FileResponse(package, filename=package.name)
    except Exception as exc:
        if source and source.exists():
            source.unlink(missing_ok=True)
        label = PLATFORM_LABELS.get(platform, "未知平台") if isinstance(platform, str) else "未知平台"
        add_run(user["id"], f"步骤 {step} · {label}", date, "FAILED", None, None, str(exc))
        return HTMLResponse(page("处理失败", f'<section class="login"><h1>处理失败</h1><p class="lead">{html.escape(str(exc))}</p><a class="button" href="/">返回工作台</a></section>', user), status_code=400)


@app.post("/step/1")
async def step1(
    request: Request,
    date: str = Form(...),
    platform: str = Form(...),
    business_type: str = Form(""),
    battery: str = Form("否"),
    file: UploadFile = File(...),
):
    return await process(request, 1, date, file, platform, business_type, battery)

@app.post("/step/2")
async def step2(request: Request, date: str = Form(...), platform: str = Form(...), file: UploadFile = File(...)):
    return await process(request, 2, date, file, platform)

@app.post("/step/3")
async def step3(request: Request, date: str = Form(...), platform: str = Form(...), file: UploadFile = File(...)):
    return await process(request, 3, date, file, platform)


@app.get("/healthz")
def healthz():
    return {"status": "ok", "platforms": list(PLATFORM_LABELS)}

@app.get("/download/{run_id}")
def download(request: Request, run_id: int):
    user = user_for(request)
    if not user: return RedirectResponse("/login", status_code=303)
    with closing(connection()) as conn:
        if user["role"] == "admin":
            row = conn.execute("SELECT result_path FROM runs WHERE id=?", (run_id,)).fetchone()
        else:
            row = conn.execute(
                "SELECT result_path FROM runs WHERE id=? AND user_id=?",
                (run_id, user["id"]),
            ).fetchone()
    if not row or not row["result_path"] or not Path(row["result_path"]).exists(): return HTMLResponse("结果不存在", status_code=404)
    path = Path(row["result_path"]); return FileResponse(path, filename=path.name)


CSS = r'''
:root{font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","PingFang SC","Microsoft YaHei UI",sans-serif;color:#18202a;background:#eef1f5;letter-spacing:0}*{box-sizing:border-box;letter-spacing:0}body{margin:0;min-height:100vh;background:#eef1f5}a{text-decoration:none;color:inherit}header{height:66px;padding:0 max(24px,4vw);display:flex;align-items:center;justify-content:space-between;background:rgba(250,251,253,.86);border-bottom:1px solid rgba(64,73,88,.16);backdrop-filter:blur(20px);position:sticky;top:0;z-index:3}.brand{font-size:18px;font-weight:760}.brand small{font-size:9px;color:#788291;margin-left:8px}nav{display:flex;gap:22px;color:#667181;font-size:13px;font-weight:650}nav a:hover{color:#0071e3}main{width:min(1200px,calc(100% - 32px));margin:auto;padding:34px 0 64px}.hero{display:flex;justify-content:space-between;align-items:end;gap:28px;margin-bottom:24px}.eyebrow{font-size:10px;color:#727d8d;font-weight:750;margin:0 0 9px}.hero h1,.head h1{font-size:34px;line-height:1.08;margin:0 0 11px}.lead{color:#626d7c;line-height:1.65;max-width:710px;margin:0;font-size:14px}.identity{display:flex;gap:10px;align-items:center;background:rgba(255,255,255,.72);padding:10px 13px;border:1px solid rgba(67,77,92,.15);border-radius:8px}.identity i{width:8px;height:8px;border-radius:50%;background:#23a566;box-shadow:0 0 0 3px #d8f2e4}.identity b,.identity small{display:block}.identity b{font-size:12px}.identity small{font-size:11px;color:#778293;margin-top:3px}.config{display:flex;align-items:center;justify-content:space-between;padding:15px 17px;background:rgba(255,255,255,.74);border:1px solid rgba(64,74,89,.14);border-radius:8px;margin-bottom:14px}.config small{color:#737e8d;margin-right:12px}.config b{font-size:14px}.config em{display:block;color:#7a8595;font-size:11px;font-style:normal;margin-top:4px}.steps{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.steps article,.panel,.history{background:rgba(255,255,255,.78);border:1px solid rgba(64,74,89,.14);border-radius:8px;padding:18px;box-shadow:0 10px 28px rgba(41,50,66,.06)}.step-head{height:25px;display:flex;align-items:center;justify-content:space-between;margin-bottom:13px}.step-head>span{font-size:10px;color:#7a8492;font-weight:700}.no{width:25px;height:25px;display:grid;place-items:center;border-radius:7px;background:#e5f1ff;color:#006edb;font-weight:800;font-size:11px}.steps h2,.panel h2,.history h2{font-size:17px;margin:0 0 7px}.steps>article>p{height:42px;color:#6e7887;font-size:12px;line-height:1.55;margin:0 0 13px}.workflow-form{display:grid;gap:10px}.platform-picker,.mini-segment{display:grid;grid-template-columns:1fr 1fr;background:#e9edf2;border:1px solid #dce1e7;border-radius:8px;padding:3px;gap:3px}.platform-picker label,.mini-segment label{cursor:pointer}.platform-picker input,.mini-segment input{position:absolute;opacity:0;pointer-events:none}.platform-picker span,.mini-segment span{height:30px;display:grid;place-items:center;border-radius:6px;color:#667181;font-size:11px;font-weight:700}.platform-picker input:checked+span,.mini-segment input:checked+span{background:#fff;color:#086dcc;box-shadow:0 1px 4px rgba(37,45,59,.14)}article[data-platform="yunexpress"] .sf-only{display:none}article[data-platform="sf"] .yun-only{display:none}.file{display:block;border:1px dashed #b8c2cf;border-radius:8px;padding:11px 12px;background:#f8fafc;cursor:pointer;overflow:hidden}.file:hover{border-color:#7ca9d7;background:#f4f8fc}.file b{display:block;font-size:12px;margin-bottom:7px}.file>span:last-child{display:block;color:#738092;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.file input{display:none}.sf-options{border:1px solid #dce2e9;background:#f7f9fb;border-radius:8px;padding:10px;gap:9px}.sf-options>label{display:grid;gap:6px;color:#687486;font-size:11px;font-weight:700}.sf-options select{width:100%}.sf-options>div{display:grid;grid-template-columns:auto 120px;align-items:center;gap:10px}.sf-options>div>b{font-size:11px;color:#687486}.mini-segment span{height:27px}.row{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px}.row input,input,select{min-width:0;border:1px solid #cdd4dd;background:#fff;border-radius:7px;padding:9px 10px;color:#202a36;font:inherit;font-size:12px}.row input:focus,input:focus,select:focus{outline:2px solid rgba(0,113,227,.2);border-color:#4795dc}.primary,.button{border:0;border-radius:7px;padding:9px 13px;font:inherit;font-weight:720;font-size:12px;cursor:pointer;display:inline-grid;place-items:center}.primary{background:#0878df;color:#fff}.primary:hover{background:#006dcc}.soft{background:#e8f2fc;color:#1267b5}.history{margin-top:14px}.section-head{display:flex;justify-content:space-between;align-items:end;margin-bottom:10px}.section-head>span{color:#7d8794;font-size:11px}.table-wrap{overflow-x:auto}table{width:100%;border-collapse:collapse;font-size:12px}th{text-align:left;color:#7d8794;font-size:10px}th,td{padding:11px 8px;border-bottom:1px solid #e3e7ec}td{color:#455164}.status{font-size:10px;font-weight:800;padding:4px 7px;border-radius:6px}.status.success{background:#e3f5eb;color:#237a50}.status.failed{background:#fce9e8;color:#ad4844}td a{color:#0876da;font-weight:700}.empty{text-align:center;color:#87919f;padding:26px}.head{margin-bottom:22px}.two{display:grid;grid-template-columns:1.15fr .85fr;gap:14px}.panel form{display:grid;gap:12px}.panel label,.login label{display:grid;gap:7px;color:#637083;font-size:12px;font-weight:700}.panel input,.login input{width:100%}.quiet{background:rgba(247,249,252,.86)}.quiet ul{padding:0;margin:10px 0 22px;list-style:none;display:grid;gap:14px}.quiet li{display:flex;gap:10px;align-items:center}.quiet li i{width:7px;height:7px;border-radius:50%;background:#0878df}.quiet li b,.quiet li small{display:block}.quiet li b{font-size:13px}.quiet li small,.hint{font-size:11px;color:#788494}.hint{line-height:1.55;margin:0}.notice{padding:14px;background:#f1f4f8;border:1px solid #dce2e9;border-radius:8px;color:#667385;font-size:12px;line-height:1.6}.login{max-width:410px;margin:7vh auto 0;background:rgba(255,255,255,.82);padding:30px;border:1px solid rgba(64,74,89,.14);border-radius:8px;box-shadow:0 18px 44px rgba(41,50,66,.09)}.mark{font-size:23px;font-weight:800;margin-bottom:30px}.mark span{color:#0878df}.login h1{font-size:27px;margin:0 0 10px}.login form{display:grid;gap:13px;margin-top:22px}@media(max-width:980px){.steps{grid-template-columns:1fr}.steps>article>p{height:auto;min-height:0}.two{grid-template-columns:1fr}}@media(max-width:620px){header{height:auto;min-height:62px;padding:13px 16px;align-items:flex-start;gap:12px}.brand small{display:none}nav{gap:12px;flex-wrap:wrap;justify-content:flex-end}main{width:min(100% - 20px,1200px);padding-top:22px}.hero{align-items:start;flex-direction:column}.hero h1,.head h1{font-size:29px}.identity{width:100%}.config{align-items:flex-start;gap:12px}.row{grid-template-columns:1fr}.row .primary{width:100%}.section-head{align-items:start;gap:10px;flex-direction:column}}
'''

SCRIPT = r'''
const localToday = new Date(Date.now() - new Date().getTimezoneOffset() * 60000)
  .toISOString()
  .slice(0, 10);
document.querySelectorAll('input[type="date"]').forEach((input) => {
  input.value = localToday;
});

document.querySelectorAll('.workflow-form').forEach((form) => {
  const article = form.closest('article');
  const fileInput = form.querySelector('input[type="file"]');
  const step = Number((form.action.match(/\/step\/(\d+)/) || [])[1] || 0);
  const accepts = {
    1: {yunexpress: '.csv', sf: '.csv'},
    2: {yunexpress: '.xlsx', sf: '.xls,.xlsx'},
    3: {yunexpress: '.zip', sf: '.pdf'}
  };
  const syncPlatform = () => {
    const selected = form.querySelector('input[name="platform"]:checked');
    const platform = selected ? selected.value : 'yunexpress';
    form.dataset.platform = platform;
    if (article) article.dataset.platform = platform;
    if (fileInput && accepts[step]) {
      fileInput.accept = accepts[step][platform];
      fileInput.value = '';
      const caption = fileInput.parentElement.querySelector(':scope > span:last-child');
      if (caption) caption.textContent = '选择文件';
    }
  };
  form.querySelectorAll('input[name="platform"]').forEach((input) => {
    input.addEventListener('change', syncPlatform);
  });
  if (fileInput) {
    fileInput.addEventListener('change', () => {
      const caption = fileInput.parentElement.querySelector(':scope > span:last-child');
      if (caption) caption.textContent = fileInput.files.length ? fileInput.files[0].name : '选择文件';
    });
  }
  syncPlatform();
});
'''
