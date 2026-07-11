from __future__ import annotations

import hashlib
import hmac
import html
import io
import os
import secrets
import shutil
import sqlite3
import sys
import tempfile
import time
from contextlib import closing, redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse

ROOT = Path(__file__).resolve().parents[1]
TOOL_ROOT = ROOT / "DP_Yunexpress_To_Au_Tool"
DATA = Path(os.getenv("DP_WEB_DATA_DIR", str(Path(__file__).parent / "data"))).resolve()
CONFIG = DATA / "config"
RESULTS = DATA / "results"
WORKING = DATA / "working"
DB_PATH = DATA / "app.sqlite3"
ADMIN_EMAIL = os.getenv("DP_WEB_ADMIN_EMAIL", "admin@example.com")
ADMIN_PASSWORD = os.getenv("DP_WEB_ADMIN_PASSWORD", "change-me")
MAX_UPLOAD = 80 * 1024 * 1024
sys.path.insert(0, str(TOOL_ROOT))
from scripts import generate_dp_shipment_upload, generate_yunexpress_template, sort_yunexpress_labels_by_sku

app = FastAPI(title="DP&Yunexpress To Au Web")
CONFIG_LABELS = {"sku": "SKU 商品库", "yun_template": "云途标准模板", "dp_template": "DP 发货模板"}


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
    return CONFIG / f"{key}.xlsx"


def add_run(user_id: int, step: str, date: str, status: str, result: Path | None = None, report: Path | None = None, error: str = ""):
    with closing(connection()) as conn:
        conn.execute("INSERT INTO runs(user_id,step,status,date,result_path,report_path,created_at,error) VALUES(?,?,?,?,?,?,?,?)", (user_id, step, status, date, str(result) if result else None, str(report) if report else None, datetime.now().isoformat(timespec="seconds"), error[:3000]))
        conn.commit()


def script(module, args: list[str]) -> tuple[int, str, str]:
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


def batch_dir(user_id: int) -> Path:
    path = WORKING / str(user_id)
    path.mkdir(parents=True, exist_ok=True)
    return path


def field(name: str, label: str, accept: str) -> str:
    return f'<label class="file"><b>{label}</b><input name="{name}" type="file" accept="{accept}" required><span>选择文件</span></label>'


def page(title: str, content: str, user=None) -> str:
    nav = f'<header><a class="brand" href="/">DP<span>&</span>Yunexpress <small>TO AU</small></a><nav><a href="/">工作台</a>{"<a href=/admin>成员管理</a>" if user and user["role"] == "admin" else ""}<a href="/logout">退出</a></nav></header>' if user else '<header><a class="brand" href="/login">DP<span>&</span>Yunexpress <small>TO AU</small></a></header>'
    return f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)} · DP&Yunexpress To Au</title><style>{CSS}</style></head><body>{nav}<main>{content}</main></body></html>'


@app.on_event("startup")
def startup():
    init_db()
    for folder in (CONFIG, RESULTS, WORKING):
        folder.mkdir(parents=True, exist_ok=True)


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    if user_for(request):
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(page("登录", '<section class="login"><div class="mark">DP<span>&</span></div><p class="eyebrow">PRIVATE OPERATIONS</p><h1>登录发货工作台</h1><p class="lead">员工账号由管理员创建，订单文件只在本次批次期间临时保存。</p><form method="post" action="/login"><label>邮箱<input name="email" type="email" required></label><label>密码<input name="password" type="password" required></label><button class="primary" type="submit">登录</button></form></section>'))


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
    response.set_cookie("dp_session", token, httponly=True, samesite="lax", max_age=86400)
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
        runs = conn.execute("SELECT * FROM runs ORDER BY id DESC LIMIT 10").fetchall()
    records = ""
    for record in runs:
        result_link = f'<a href="/download/{record["id"]}">下载结果</a>' if record["result_path"] and Path(record["result_path"]).exists() else "-"
        records += f'<tr><td>{html.escape(record["created_at"])}</td><td>{html.escape(record["step"])}</td><td><span class="status {record["status"].lower()}">{html.escape(record["status"])}</span></td><td>{html.escape(record["date"])}</td><td>{result_link}</td></tr>'
    records = records or '<tr><td colspan="5" class="empty">还没有处理记录，从第 1 步开始。</td></tr>'
    today = datetime.now().date().isoformat()
    content = f'''<section class="hero"><div><p class="eyebrow">AU FULFILLMENT DESK</p><h1>发货工作台</h1><p class="lead">把 DP 订单、云途寄件和供应商面单，整理成一条可校验的工作流。</p></div><div class="identity"><i></i><div><b>{html.escape(user["email"])}</b><small>{"管理员" if user["role"] == "admin" else "员工"} · 在线</small></div></div></section><section class="config"><div><small>固定配置</small><b>{"已配置" if all(config_path(k).exists() for k in CONFIG_LABELS) else "待配置"}</b><em>SKU 商品库、云途标准模板、DP 发货模板</em></div><a class="button soft" href="/config">管理固定配置</a></section><section class="steps"><article><div class="no">01</div><h2>生成云途模板</h2><p>上传本次 DP 订单 CSV，生成云途批量寄件 Excel。</p><form action="/step/1" method="post" enctype="multipart/form-data">{field("file","本次 DP 订单 CSV",".csv")}<div class="row"><input name="date" type="date" value="{today}" required><button class="primary" type="submit">开始生成</button></div></form></article><article><div class="no">02</div><h2>生成 DP 回填模板</h2><p>上传云途创建寄件订单后下载的订单信息 Excel。</p><form action="/step/2" method="post" enctype="multipart/form-data">{field("file","本次云途订单信息 Excel",".xlsx")}<div class="row"><input name="date" type="date" value="{today}" required><button class="primary" type="submit">开始生成</button></div></form></article><article><div class="no">03</div><h2>分拣供应商面单</h2><p>上传所有订单面单 ZIP，下载按 SKU 打包的结果压缩包。</p><form action="/step/3" method="post" enctype="multipart/form-data">{field("file","本次云途面单 ZIP",".zip")}<div class="row"><input name="date" type="date" value="{today}" required><button class="primary" type="submit">开始分拣</button></div></form></article></section><section class="history"><div class="section-head"><div><p class="eyebrow">ACTIVITY</p><h2>最近处理记录</h2></div><span>只保留结果和校验状态</span></div><table><thead><tr><th>时间</th><th>步骤</th><th>状态</th><th>日期</th><th>结果</th></tr></thead><tbody>{records}</tbody></table></section>'''
    return HTMLResponse(page("工作台", content, user))


@app.get("/config", response_class=HTMLResponse)
def config_page(request: Request):
    user = user_for(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    status = "".join(f'<li><i></i><div><b>{label}</b><small>{"已上传，可替换" if config_path(key).exists() else "尚未上传"}</small></div></li>' for key, label in CONFIG_LABELS.items())
    content = f'<section class="head"><p class="eyebrow">CONFIGURATION</p><h1>固定配置</h1><p class="lead">三份模板只需在更新时替换，不需要每天重复上传。</p></section><section class="two"><div class="panel"><h2>替换文件</h2><form action="/config" method="post" enctype="multipart/form-data">{field("sku","SKU 商品库",".xlsx")}{field("yun_template","云途标准模板",".xlsx")}{field("dp_template","DP 发货模板",".xlsx")}<button class="primary" type="submit">保存固定配置</button></form></div><div class="panel quiet"><p class="eyebrow">CURRENT</p><h2>当前状态</h2><ul>{status}</ul><p class="hint">替换后，下一次任务立即使用新文件。</p></div></section>'
    return HTMLResponse(page("固定配置", content, user))


@app.post("/config")
async def save_config(request: Request, sku: UploadFile = File(...), yun_template: UploadFile = File(...), dp_template: UploadFile = File(...)):
    user = user_for(request)
    if not user or user["role"] != "admin": return HTMLResponse("无权限", status_code=403)
    for key, upload in (("sku", sku), ("yun_template", yun_template), ("dp_template", dp_template)):
        path = await save_upload(upload, CONFIG, {".xlsx"}); path.replace(config_path(key))
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


async def process(request: Request, step: int, date: str, upload: UploadFile):
    user = user_for(request)
    if not user: return RedirectResponse("/login", status_code=303)
    if not all(config_path(k).exists() for k in CONFIG_LABELS): return HTMLResponse(page("需要配置", '<section class="login"><h1>请先上传固定配置</h1><a class="button" href="/config">去上传</a></section>', user), status_code=400)
    temp = Path(tempfile.mkdtemp(prefix="dp-upload-")); batch = batch_dir(user["id"]); output = RESULTS / date.replace("-", ""); output.mkdir(parents=True, exist_ok=True)
    try:
        ext = {1: {".csv"}, 2: {".xlsx"}, 3: {".zip"}}[step]
        source = await save_upload(upload, batch, ext)
        if step == 1:
            (batch / "dp_orders.csv").unlink(missing_ok=True); source.replace(batch / "dp_orders.csv")
            code, out, err = script(generate_yunexpress_template, ["--dp-orders-csv", str(batch / "dp_orders.csv"), "--sku-xlsx", str(config_path("sku")), "--yunexpress-template-xlsx", str(config_path("yun_template")), "--date", date, "--output-root", str(output)])
        elif step == 2:
            source.replace(batch / "yun_orders.xlsx")
            if not (batch / "dp_orders.csv").exists(): raise ValueError("请先完成第 1 步")
            code, out, err = script(generate_dp_shipment_upload, ["--yunexpress-xlsx", str(batch / "yun_orders.xlsx"), "--shipment-template-xlsx", str(config_path("dp_template")), "--dp-orders-csv", str(batch / "dp_orders.csv"), "--date", date, "--output-root", str(output)])
        else:
            source.replace(batch / "labels.zip")
            if not (batch / "dp_orders.csv").exists() or not (batch / "yun_orders.xlsx").exists(): raise ValueError("请先完成第 1、2 步")
            code, out, err = script(sort_yunexpress_labels_by_sku, ["--zip", str(batch / "labels.zip"), "--yunexpress-xlsx", str(batch / "yun_orders.xlsx"), "--dp-orders-csv", str(batch / "dp_orders.csv"), "--sku-xlsx", str(config_path("sku")), "--date", date, "--output-root", str(output)])
        if code != 0: raise ValueError(err.strip() or out.strip() or "处理失败")
        result = result_from(out)
        report = result.with_suffix(".report.json") if result.is_file() else result / "校验成功报告.json"
        package = result if result.is_file() else Path(shutil.make_archive(str(result), "zip", root_dir=result.parent, base_dir=result.name))
        add_run(user["id"], f"步骤 {step}", date, "SUCCESS", package, report if report.exists() else None)
        if step == 3: shutil.rmtree(batch, ignore_errors=True)
        return FileResponse(package, filename=package.name)
    except Exception as exc:
        add_run(user["id"], f"步骤 {step}", date, "FAILED", None, None, str(exc))
        if step == 3: shutil.rmtree(batch, ignore_errors=True)
        return HTMLResponse(page("处理失败", f'<section class="login"><h1>处理失败</h1><p class="lead">{html.escape(str(exc))}</p><a class="button" href="/">返回工作台</a></section>', user), status_code=400)
    finally: shutil.rmtree(temp, ignore_errors=True)


@app.post("/step/1")
async def step1(request: Request, date: str = Form(...), file: UploadFile = File(...)): return await process(request, 1, date, file)

@app.post("/step/2")
async def step2(request: Request, date: str = Form(...), file: UploadFile = File(...)): return await process(request, 2, date, file)

@app.post("/step/3")
async def step3(request: Request, date: str = Form(...), file: UploadFile = File(...)): return await process(request, 3, date, file)

@app.get("/download/{run_id}")
def download(request: Request, run_id: int):
    if not user_for(request): return RedirectResponse("/login", status_code=303)
    with closing(connection()) as conn: row = conn.execute("SELECT result_path FROM runs WHERE id=?", (run_id,)).fetchone()
    if not row or not row["result_path"] or not Path(row["result_path"]).exists(): return HTMLResponse("结果不存在", status_code=404)
    path = Path(row["result_path"]); return FileResponse(path, filename=path.name)


CSS = r'''
:root{font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","Microsoft YaHei UI",sans-serif;color:#142033;background:#edf4fb}*{box-sizing:border-box}body{margin:0;min-height:100vh;background:radial-gradient(circle at 12% 0%,#fff 0,#edf4fb 42%,#e8f1fb 100%)}a{text-decoration:none;color:inherit}header{height:76px;padding:0 5vw;display:flex;align-items:center;justify-content:space-between;background:rgba(250,253,255,.76);border-bottom:1px solid #d7e3ef;backdrop-filter:blur(18px);position:sticky;top:0;z-index:3}.brand{font-size:20px;font-weight:800;letter-spacing:-.7px}.brand span,.mark span{color:#0a84ff}.brand small{font-size:9px;color:#728198;letter-spacing:1.6px;margin-left:7px}nav{display:flex;gap:24px;color:#65758a;font-size:13px;font-weight:700}nav a:hover{color:#0a84ff}main{width:min(1180px,90vw);margin:auto;padding:54px 0 80px}.hero{display:flex;justify-content:space-between;align-items:end;gap:30px;margin-bottom:34px}.eyebrow{font-size:10px;letter-spacing:1.7px;color:#6d82a0;font-weight:800;margin:0 0 10px}.hero h1,.head h1{font-size:40px;letter-spacing:-1.8px;line-height:1;margin:0 0 14px}.lead{color:#66768a;line-height:1.7;max-width:620px;margin:0;font-size:14px}.identity{display:flex;gap:10px;align-items:center;background:#fff9;padding:11px 14px;border:1px solid #d8e4f1;border-radius:16px}.identity i{width:8px;height:8px;border-radius:50%;background:#2bb673;box-shadow:0 0 0 4px #dff5e9}.identity b,.identity small{display:block}.identity b{font-size:12px}.identity small{font-size:11px;color:#7c8da2;margin-top:3px}.config{display:flex;align-items:center;justify-content:space-between;padding:18px 22px;background:#ffffffb5;border:1px solid #d7e3f0;border-radius:18px;margin-bottom:20px}.config small{color:#70819a;margin-right:14px}.config b{font-size:14px}.config em{display:block;color:#7d8da2;font-size:11px;font-style:normal;margin-top:5px}.steps{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}.steps article,.panel,.history{background:#ffffffb8;border:1px solid #d6e3f1;border-radius:20px;padding:22px;box-shadow:0 16px 42px #38639012}.no{color:#0a84ff;font-weight:800;font-size:12px;letter-spacing:1px;margin-bottom:28px}.steps h2,.panel h2,.history h2{font-size:18px;letter-spacing:-.5px;margin:0 0 8px}.steps p{min-height:48px;color:#718197;font-size:13px;line-height:1.55;margin:0 0 18px}.file{display:block;border:1px dashed #b9cbe0;border-radius:12px;padding:12px 13px;background:#f7fbff;cursor:pointer}.file b{display:block;font-size:12px;margin-bottom:8px}.file input{display:none}.file span{color:#71849a;font-size:11px}.row{display:flex;gap:8px;margin-top:10px}.row input{flex:1}.row input,input{border:1px solid #cad9e8;background:#fbfdff;border-radius:10px;padding:10px 11px;color:#1b2b41;font:inherit;font-size:12px}.primary,.button{border:0;border-radius:11px;padding:10px 14px;font:inherit;font-weight:700;font-size:12px;cursor:pointer;display:inline-block}.primary{background:#0a84ff;color:#fff;box-shadow:0 7px 18px #0a84ff3b}.soft{background:#e9f3ff;color:#1465af}.history{margin-top:22px}.section-head{display:flex;justify-content:space-between;align-items:end;margin-bottom:14px}.section-head>span{color:#8291a3;font-size:11px}table{width:100%;border-collapse:collapse;font-size:12px}th{text-align:left;color:#8392a4;font-size:10px;letter-spacing:.8px;text-transform:uppercase}th,td{padding:12px 8px;border-bottom:1px solid #e5edf5}td{color:#43536a}.status{font-size:10px;font-weight:800;padding:4px 8px;border-radius:7px}.status.success{background:#e3f7ec;color:#258456}.status.failed{background:#fff0ef;color:#b34b47}td a{color:#0a84ff;font-weight:700}.empty{text-align:center;color:#8a99ab;padding:28px}.two{display:grid;grid-template-columns:1.2fr .8fr;gap:18px}.panel form{display:grid;gap:15px}.panel label,.login label{display:grid;gap:7px;color:#63748a;font-size:12px;font-weight:700}.panel input,.login input{width:100%}.quiet{background:#f6faffb5}.quiet ul{padding:0;margin:10px 0 24px;list-style:none;display:grid;gap:17px}.quiet li{display:flex;gap:11px;align-items:center}.quiet li i{width:7px;height:7px;border-radius:50%;background:#0a84ff}.quiet li b,.quiet li small{display:block}.quiet li b{font-size:13px}.quiet li small,.hint{font-size:11px;color:#7b8ca1}.login{max-width:410px;margin:8vh auto 0;background:#ffffffbf;padding:34px;border:1px solid #d5e2ef;border-radius:22px;box-shadow:0 22px 60px #36608f1a}.mark{font-size:25px;font-weight:800;margin-bottom:34px}.login h1{font-size:29px;letter-spacing:-1px;margin:0 0 11px}.login form{display:grid;gap:14px;margin-top:25px}@media(max-width:820px){.steps,.two{grid-template-columns:1fr}.hero{align-items:start;flex-direction:column}.hero h1,.head h1{font-size:34px}}
'''

