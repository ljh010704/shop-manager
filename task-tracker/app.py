import os
import threading
import time
from datetime import date, datetime, timedelta

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import Optional

import crypto

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data.db")

app = FastAPI()
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


def get_db():
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


@app.on_event("startup")
def startup():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            license_name TEXT NOT NULL,
            license_no TEXT DEFAULT '',
            holder_name TEXT DEFAULT '',
            expire_date TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS shops (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shop_name TEXT NOT NULL UNIQUE,
            license_id INTEGER DEFAULT NULL,
            group_name TEXT DEFAULT '',
            status TEXT DEFAULT 'new',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            new_shop_task_done INTEGER DEFAULT 0,
            FOREIGN KEY (license_id) REFERENCES licenses(id)
        );

        CREATE TABLE IF NOT EXISTS platform_accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shop_id INTEGER NOT NULL,
            platform_name TEXT NOT NULL,
            account TEXT DEFAULT '',
            password_enc TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            updated_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (shop_id) REFERENCES shops(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS daily_task_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_name TEXT NOT NULL,
            sort_order INTEGER DEFAULT 0,
            is_active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS new_shop_task_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_name TEXT NOT NULL,
            sort_order INTEGER DEFAULT 0,
            is_active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS daily_tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shop_id INTEGER NOT NULL,
            task_id INTEGER NOT NULL,
            task_date TEXT NOT NULL,
            is_completed INTEGER DEFAULT 0,
            completed_at TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            UNIQUE(shop_id, task_id, task_date),
            FOREIGN KEY (shop_id) REFERENCES shops(id) ON DELETE CASCADE,
            FOREIGN KEY (task_id) REFERENCES daily_task_templates(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS new_shop_tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shop_id INTEGER NOT NULL,
            task_id INTEGER NOT NULL,
            is_completed INTEGER DEFAULT 0,
            completed_at TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            UNIQUE(shop_id, task_id),
            FOREIGN KEY (shop_id) REFERENCES shops(id) ON DELETE CASCADE,
            FOREIGN KEY (task_id) REFERENCES new_shop_task_templates(id) ON DELETE CASCADE
        );
    """)

    daily_count = conn.execute("SELECT COUNT(*) FROM daily_task_templates").fetchone()[0]
    if daily_count == 0:
        defaults = [
            ("处理待发货订单", 1), ("处理售后/退款订单", 2),
            ("检查店铺违规", 3), ("检查体验分并优化", 4),
            ("上新商品（3-5个）", 5), ("优化商品标题/主图", 6),
            ("检查库存情况", 7), ("处理工单", 8),
            ("店铺数据复盘", 9), ("竞品分析", 10),
        ]
        conn.executemany("INSERT INTO daily_task_templates(task_name, sort_order) VALUES(?,?)", defaults)

    new_count = conn.execute("SELECT COUNT(*) FROM new_shop_task_templates").fetchone()[0]
    if new_count == 0:
        defaults = [
            ("完善店铺基本信息（名称、头像、简介）", 1),
            ("设置运费模板", 2), ("设置售后地址", 3),
            ("开通支付方式", 4), ("上传营业执照（如需要）", 5),
            ("完善店铺资质", 6), ("设置客服自动回复", 7),
            ("完善退换货规则", 8), ("上架第一批商品（5-10个）", 9),
            ("设置店铺优惠券", 10), ("开通运费险", 11),
            ("店铺装修", 12),
        ]
        conn.executemany("INSERT INTO new_shop_task_templates(task_name, sort_order) VALUES(?,?)", defaults)

    conn.commit()
    conn.close()


def ensure_daily_tasks(conn, shop_id, task_date):
    templates_list = conn.execute(
        "SELECT id FROM daily_task_templates WHERE is_active=1 ORDER BY sort_order"
    ).fetchall()
    for t in templates_list:
        conn.execute(
            "INSERT OR IGNORE INTO daily_tasks(shop_id, task_id, task_date) VALUES(?,?,?)",
            (shop_id, t['id'], task_date)
        )
    conn.commit()


def ensure_new_shop_tasks(conn, shop_id):
    templates_list = conn.execute(
        "SELECT id FROM new_shop_task_templates WHERE is_active=1 ORDER BY sort_order"
    ).fetchall()
    for t in templates_list:
        conn.execute(
            "INSERT OR IGNORE INTO new_shop_tasks(shop_id, task_id) VALUES(?,?)",
            (shop_id, t['id'])
        )
    conn.commit()


# ==================== 首页 ====================

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    conn = get_db()
    today = date.today().isoformat()

    shops = [dict(r) for r in conn.execute(
        "SELECT * FROM shops WHERE status != 'closed' ORDER BY shop_name"
    ).fetchall()]

    for shop in shops:
        ensure_daily_tasks(conn, shop['id'], today)
        total = conn.execute(
            "SELECT COUNT(*) FROM daily_tasks WHERE shop_id=? AND task_date=?",
            (shop['id'], today)
        ).fetchone()[0]
        done = conn.execute(
            "SELECT COUNT(*) FROM daily_tasks WHERE shop_id=? AND task_date=? AND is_completed=1",
            (shop['id'], today)
        ).fetchone()[0]
        shop['total_tasks'] = total
        shop['done_tasks'] = done
        shop['progress'] = round(done / total * 100) if total > 0 else 0

        if shop['status'] == 'new':
            new_total = conn.execute(
                "SELECT COUNT(*) FROM new_shop_tasks WHERE shop_id=?", (shop['id'],)
            ).fetchone()[0]
            new_done = conn.execute(
                "SELECT COUNT(*) FROM new_shop_tasks WHERE shop_id=? AND is_completed=1",
                (shop['id'],)
            ).fetchone()[0]
            shop['new_total'] = new_total
            shop['new_done'] = new_done
            shop['new_progress'] = round(new_done / new_total * 100) if new_total > 0 else 0
        else:
            shop['new_total'] = 0
            shop['new_done'] = 0
            shop['new_progress'] = 0

    total_shops = len(shops)
    total_done = sum(1 for s in shops if s['progress'] == 100)
    new_shops = [s for s in shops if s['status'] == 'new']

    conn.close()
    return templates.TemplateResponse(request, "index.html", {
        "request": request, "shops": shops, "today": today,
        "total_shops": total_shops, "total_done": total_done,
        "new_shops": new_shops
    })


# ==================== 营业执照管理 ====================

@app.get("/licenses", response_class=HTMLResponse)
def licenses_page(request: Request):
    conn = get_db()
    licenses = [dict(r) for r in conn.execute("SELECT * FROM licenses ORDER BY created_at DESC").fetchall()]
    for lic in licenses:
        count = conn.execute("SELECT COUNT(*) FROM shops WHERE license_id=?", (lic['id'],)).fetchone()[0]
        lic['shop_count'] = count
    conn.close()
    return templates.TemplateResponse(request, "licenses.html", {
        "request": request, "licenses": licenses
    })


@app.post("/api/licenses/add")
def add_license(request: Request, license_name: str = Form(...), license_no: str = Form(""),
                holder_name: str = Form(""), expire_date: str = Form(""), remark: str = Form("")):
    conn = get_db()
    conn.execute(
        "INSERT INTO licenses(license_name, license_no, holder_name, expire_date, remark) VALUES(?,?,?,?,?)",
        (license_name, license_no, holder_name, expire_date, remark)
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/licenses", status_code=302)


@app.post("/api/licenses/{license_id}/delete")
def delete_license(license_id: int):
    conn = get_db()
    conn.execute("UPDATE shops SET license_id=NULL WHERE license_id=?", (license_id,))
    conn.execute("DELETE FROM licenses WHERE id=?", (license_id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/licenses", status_code=302)


@app.post("/api/licenses/{license_id}/edit")
def edit_license(license_id: int, license_name: str = Form(...), license_no: str = Form(""),
                 holder_name: str = Form(""), expire_date: str = Form(""), remark: str = Form("")):
    conn = get_db()
    conn.execute(
        "UPDATE licenses SET license_name=?, license_no=?, holder_name=?, expire_date=?, remark=? WHERE id=?",
        (license_name, license_no, holder_name, expire_date, remark, license_id)
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/licenses", status_code=302)


# ==================== 店铺管理 ====================

@app.get("/shops", response_class=HTMLResponse)
def shops_page(request: Request, group: str = ""):
    conn = get_db()
    query = "SELECT s.*, l.license_name FROM shops s LEFT JOIN licenses l ON s.license_id=l.id"
    params = []
    if group:
        query += " WHERE s.group_name=?"
        params.append(group)
    query += " ORDER BY s.status, s.shop_name"
    shops = [dict(r) for r in conn.execute(query, params).fetchall()]

    all_groups = [r['group_name'] for r in conn.execute(
        "SELECT DISTINCT group_name FROM shops WHERE group_name != '' ORDER BY group_name"
    ).fetchall()]
    licenses = [dict(r) for r in conn.execute("SELECT id, license_name FROM licenses ORDER BY license_name").fetchall()]
    conn.close()
    return templates.TemplateResponse(request, "shops.html", {
        "request": request, "shops": shops, "groups": all_groups,
        "licenses": licenses, "selected_group": group
    })


@app.post("/api/shops/add")
def add_shop(request: Request, shop_name: str = Form(...), license_id: str = Form(""),
             group_name: str = Form(""), status: str = Form("new")):
    conn = get_db()
    lid = int(license_id) if license_id else None
    try:
        cur = conn.execute(
            "INSERT INTO shops(shop_name, license_id, group_name, status) VALUES(?,?,?,?)",
            (shop_name, lid, group_name, status)
        )
        shop_id = cur.lastrowid
        if status == 'new':
            ensure_new_shop_tasks(conn, shop_id)
        conn.commit()
    except Exception:
        pass
    conn.close()
    return RedirectResponse("/shops", status_code=302)


@app.post("/api/shops/{shop_id}/edit")
def edit_shop(shop_id: int, shop_name: str = Form(...), license_id: str = Form(""),
              group_name: str = Form(""), status: str = Form("new")):
    conn = get_db()
    lid = int(license_id) if license_id else None
    old = conn.execute("SELECT status FROM shops WHERE id=?", (shop_id,)).fetchone()
    conn.execute(
        "UPDATE shops SET shop_name=?, license_id=?, group_name=?, status=? WHERE id=?",
        (shop_name, lid, group_name, status, shop_id)
    )
    if old and old['status'] != 'new' and status == 'new':
        ensure_new_shop_tasks(conn, shop_id)
    if status == 'new':
        shop = conn.execute("SELECT new_shop_task_done FROM shops WHERE id=?", (shop_id,)).fetchone()
        if shop and not shop['new_shop_task_done']:
            ensure_new_shop_tasks(conn, shop_id)
    conn.commit()
    conn.close()
    return RedirectResponse("/shops", status_code=302)


@app.post("/api/shops/{shop_id}/delete")
def delete_shop(shop_id: int):
    conn = get_db()
    conn.execute("DELETE FROM daily_tasks WHERE shop_id=?", (shop_id,))
    conn.execute("DELETE FROM new_shop_tasks WHERE shop_id=?", (shop_id,))
    conn.execute("DELETE FROM platform_accounts WHERE shop_id=?", (shop_id,))
    conn.execute("DELETE FROM shops WHERE id=?", (shop_id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/shops", status_code=302)


# ==================== 平台账号管理 ====================

@app.get("/accounts", response_class=HTMLResponse)
def accounts_page(request: Request, shop_id: str = "", platform: str = ""):
    conn = get_db()
    shops = [dict(r) for r in conn.execute("SELECT id, shop_name FROM shops ORDER BY shop_name").fetchall()]

    accounts = []
    if shop_id:
        rows = conn.execute(
            "SELECT a.*, s.shop_name FROM platform_accounts a JOIN shops s ON a.shop_id=s.id WHERE a.shop_id=? ORDER BY a.platform_name",
            (int(shop_id),)
        ).fetchall()
        accounts = [dict(r) for r in rows]
    elif platform:
        rows = conn.execute(
            "SELECT a.*, s.shop_name FROM platform_accounts a JOIN shops s ON a.shop_id=s.id WHERE a.platform_name=? ORDER BY s.shop_name",
            (platform,)
        ).fetchall()
        accounts = [dict(r) for r in rows]
    else:
        rows = conn.execute(
            "SELECT a.*, s.shop_name FROM platform_accounts a JOIN shops s ON a.shop_id=s.id ORDER BY s.shop_name, a.platform_name"
        ).fetchall()
        accounts = [dict(r) for r in rows]

    for acc in accounts:
        acc['password_dec'] = crypto.decrypt(acc['password_enc'])

    all_platforms = [r['platform_name'] for r in conn.execute(
        "SELECT DISTINCT platform_name FROM platform_accounts ORDER BY platform_name"
    ).fetchall()]

    conn.close()
    return templates.TemplateResponse(request, "accounts.html", {
        "request": request, "accounts": accounts, "shops": shops,
        "all_platforms": all_platforms, "selected_shop": shop_id, "selected_platform": platform
    })


@app.post("/api/accounts/add")
def add_account(request: Request, shop_id: int = Form(...), platform_name: str = Form(...),
                account: str = Form(""), password: str = Form(""), remark: str = Form("")):
    conn = get_db()
    enc_pwd = crypto.encrypt(password)
    conn.execute(
        "INSERT INTO platform_accounts(shop_id, platform_name, account, password_enc, remark) VALUES(?,?,?,?,?)",
        (shop_id, platform_name, account, enc_pwd, remark)
    )
    conn.commit()
    conn.close()
    return RedirectResponse(f"/accounts?shop_id={shop_id}", status_code=302)


@app.post("/api/accounts/{acc_id}/edit")
def edit_account(acc_id: int, platform_name: str = Form(...), account: str = Form(""),
                 password: str = Form(""), remark: str = Form("")):
    conn = get_db()
    enc_pwd = crypto.encrypt(password)
    conn.execute(
        "UPDATE platform_accounts SET platform_name=?, account=?, password_enc=?, remark=?, updated_at=datetime('now','localtime') WHERE id=?",
        (platform_name, account, enc_pwd, remark, acc_id)
    )
    conn.commit()
    acc = conn.execute("SELECT shop_id FROM platform_accounts WHERE id=?", (acc_id,)).fetchone()
    conn.close()
    shop_id = acc['shop_id'] if acc else ""
    return RedirectResponse(f"/accounts?shop_id={shop_id}", status_code=302)


@app.post("/api/accounts/{acc_id}/delete")
def delete_account(acc_id: int):
    conn = get_db()
    acc = conn.execute("SELECT shop_id FROM platform_accounts WHERE id=?", (acc_id,)).fetchone()
    conn.execute("DELETE FROM platform_accounts WHERE id=?", (acc_id,))
    conn.commit()
    shop_id = acc['shop_id'] if acc else ""
    conn.close()
    return RedirectResponse(f"/accounts?shop_id={shop_id}", status_code=302)


@app.get("/api/accounts/{acc_id}/get_password")
def get_password(acc_id: int):
    conn = get_db()
    acc = conn.execute("SELECT password_enc FROM platform_accounts WHERE id=?", (acc_id,)).fetchone()
    conn.close()
    if acc:
        return {"password": crypto.decrypt(acc['password_enc'])}
    return {"password": ""}


# ==================== 每日任务模板管理 ====================

@app.get("/daily_task_templates", response_class=HTMLResponse)
def daily_task_templates_page(request: Request):
    conn = get_db()
    tasks = [dict(r) for r in conn.execute(
        "SELECT * FROM daily_task_templates ORDER BY sort_order"
    ).fetchall()]
    conn.close()
    return templates.TemplateResponse(request, "daily_tasks.html", {
        "request": request, "tasks": tasks
    })


@app.post("/api/daily_task_templates/add")
def add_daily_task(task_name: str = Form(...)):
    conn = get_db()
    max_order = conn.execute("SELECT COALESCE(MAX(sort_order),0) FROM daily_task_templates").fetchone()[0]
    conn.execute("INSERT INTO daily_task_templates(task_name, sort_order) VALUES(?,?)",
                 (task_name, max_order + 1))
    conn.commit()
    conn.close()
    return RedirectResponse("/daily_task_templates", status_code=302)


@app.post("/api/daily_task_templates/{task_id}/toggle")
def toggle_daily_task(task_id: int):
    conn = get_db()
    conn.execute("UPDATE daily_task_templates SET is_active = 1 - is_active WHERE id=?", (task_id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/daily_task_templates", status_code=302)


@app.post("/api/daily_task_templates/{task_id}/delete")
def delete_daily_task(task_id: int):
    conn = get_db()
    conn.execute("DELETE FROM daily_task_templates WHERE id=?", (task_id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/daily_task_templates", status_code=302)


# ==================== 新店任务模板管理 ====================

@app.get("/new_shop_task_templates", response_class=HTMLResponse)
def new_shop_task_templates_page(request: Request):
    conn = get_db()
    tasks = [dict(r) for r in conn.execute(
        "SELECT * FROM new_shop_task_templates ORDER BY sort_order"
    ).fetchall()]
    conn.close()
    return templates.TemplateResponse(request, "new_shop_tasks.html", {
        "request": request, "tasks": tasks
    })


@app.post("/api/new_shop_task_templates/add")
def add_new_shop_task(task_name: str = Form(...)):
    conn = get_db()
    max_order = conn.execute("SELECT COALESCE(MAX(sort_order),0) FROM new_shop_task_templates").fetchone()[0]
    conn.execute("INSERT INTO new_shop_task_templates(task_name, sort_order) VALUES(?,?)",
                 (task_name, max_order + 1))
    conn.commit()
    conn.close()
    return RedirectResponse("/new_shop_task_templates", status_code=302)


@app.post("/api/new_shop_task_templates/{task_id}/toggle")
def toggle_new_shop_task(task_id: int):
    conn = get_db()
    conn.execute("UPDATE new_shop_task_templates SET is_active = 1 - is_active WHERE id=?", (task_id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/new_shop_task_templates", status_code=302)


@app.post("/api/new_shop_task_templates/{task_id}/delete")
def delete_new_shop_task(task_id: int):
    conn = get_db()
    conn.execute("DELETE FROM new_shop_task_templates WHERE id=?", (task_id,))
    conn.commit()
    conn.close()
    return RedirectResponse("/new_shop_task_templates", status_code=302)


# ==================== 每日打卡 ====================

@app.get("/checkin", response_class=HTMLResponse)
def checkin_page(request: Request, shop_id: str = "", qdate: str = ""):
    conn = get_db()
    today = date.today().isoformat()
    selected_date = qdate if qdate else today

    shops = [dict(r) for r in conn.execute(
        "SELECT id, shop_name, group_name, status FROM shops WHERE status != 'closed' ORDER BY shop_name"
    ).fetchall()]

    tasks_data = []
    selected_shop = None
    if shop_id:
        selected_shop = next((s for s in shops if s['id'] == int(shop_id)), None)
        ensure_daily_tasks(conn, int(shop_id), selected_date)
        rows = conn.execute("""
            SELECT dt.id, dt.is_completed, dt.completed_at, dt.remark,
                   t.task_name, t.sort_order
            FROM daily_tasks dt
            JOIN daily_task_templates t ON dt.task_id = t.id
            WHERE dt.shop_id=? AND dt.task_date=?
            ORDER BY t.sort_order
        """, (int(shop_id), selected_date)).fetchall()
        tasks_data = [dict(r) for r in rows]

    conn.close()
    return templates.TemplateResponse(request, "checkin.html", {
        "request": request, "shops": shops, "tasks": tasks_data,
        "selected_shop": selected_shop, "selected_date": selected_date, "today": today
    })


@app.post("/api/checkin/toggle")
def checkin_toggle(shop_id: int = Form(...), task_id: int = Form(...), task_date: str = Form(...)):
    conn = get_db()
    row = conn.execute(
        "SELECT is_completed FROM daily_tasks WHERE shop_id=? AND task_id=? AND task_date=?",
        (shop_id, task_id, task_date)
    ).fetchone()
    if row:
        new_val = 0 if row['is_completed'] else 1
        completed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if new_val else ""
        conn.execute(
            "UPDATE daily_tasks SET is_completed=?, completed_at=? WHERE shop_id=? AND task_id=? AND task_date=?",
            (new_val, completed_at, shop_id, task_id, task_date)
        )
    conn.commit()
    conn.close()
    return {"ok": True}


@app.post("/api/checkin/complete_all")
def checkin_complete_all(shop_id: int = Form(...), task_date: str = Form(...)):
    conn = get_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "UPDATE daily_tasks SET is_completed=1, completed_at=? WHERE shop_id=? AND task_date=?",
        (now, shop_id, task_date)
    )
    conn.commit()
    conn.close()
    return {"ok": True}


# ==================== 新店任务打卡 ====================

@app.get("/new_shop_checkin", response_class=HTMLResponse)
def new_shop_checkin_page(request: Request, shop_id: str = ""):
    conn = get_db()
    new_shops = [dict(r) for r in conn.execute(
        "SELECT id, shop_name FROM shops WHERE status='new' ORDER BY shop_name"
    ).fetchall()]

    tasks_data = []
    selected_shop = None
    if shop_id:
        selected_shop = next((s for s in new_shops if s['id'] == int(shop_id)), None)
        ensure_new_shop_tasks(conn, int(shop_id))
        rows = conn.execute("""
            SELECT ns.id, ns.is_completed, ns.completed_at, ns.remark,
                   t.task_name, t.sort_order
            FROM new_shop_tasks ns
            JOIN new_shop_task_templates t ON ns.task_id = t.id
            WHERE ns.shop_id=?
            ORDER BY t.sort_order
        """, (int(shop_id),)).fetchall()
        tasks_data = [dict(r) for r in rows]

    conn.close()
    return templates.TemplateResponse(request, "new_shop_checkin.html", {
        "request": request, "new_shops": new_shops, "tasks": tasks_data,
        "selected_shop": selected_shop
    })


@app.post("/api/new_shop_checkin/toggle")
def new_shop_checkin_toggle(shop_id: int = Form(...), task_id: int = Form(...)):
    conn = get_db()
    row = conn.execute(
        "SELECT is_completed FROM new_shop_tasks WHERE shop_id=? AND task_id=?",
        (shop_id, task_id)
    ).fetchone()
    if row:
        new_val = 0 if row['is_completed'] else 1
        completed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if new_val else ""
        conn.execute(
            "UPDATE new_shop_tasks SET is_completed=?, completed_at=? WHERE shop_id=? AND task_id=?",
            (new_val, completed_at, shop_id, task_id)
        )
    conn.commit()
    conn.close()
    return {"ok": True}


@app.post("/api/new_shop_checkin/complete_all")
def new_shop_checkin_complete_all(shop_id: int = Form(...)):
    conn = get_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "UPDATE new_shop_tasks SET is_completed=1, completed_at=? WHERE shop_id=?",
        (now, shop_id)
    )
    conn.execute("UPDATE shops SET new_shop_task_done=1, status='normal' WHERE id=?", (shop_id,))
    conn.commit()
    conn.close()
    return {"ok": True}


# ==================== 已完成新店任务查询 ====================

@app.get("/new_shop_history", response_class=HTMLResponse)
def new_shop_history_page(request: Request, shop_id: str = ""):
    conn = get_db()
    completed_shops = [dict(r) for r in conn.execute("""
        SELECT id, shop_name, created_at
        FROM shops
        WHERE new_shop_task_done=1
        ORDER BY shop_name
    """).fetchall()]

    tasks_data = []
    selected_shop = None
    if shop_id:
        selected_shop = next((s for s in completed_shops if s['id'] == int(shop_id)), None)
        rows = conn.execute("""
            SELECT ns.is_completed, ns.completed_at, ns.remark,
                   t.task_name, t.sort_order
            FROM new_shop_tasks ns
            JOIN new_shop_task_templates t ON ns.task_id = t.id
            WHERE ns.shop_id=?
            ORDER BY t.sort_order
        """, (int(shop_id),)).fetchall()
        tasks_data = [dict(r) for r in rows]

    conn.close()
    return templates.TemplateResponse(request, "new_shop_history.html", {
        "request": request, "completed_shops": completed_shops,
        "tasks": tasks_data, "selected_shop": selected_shop
    })


# ==================== 统计报表 ====================

@app.get("/stats", response_class=HTMLResponse)
def stats_page(request: Request, days: int = 7):
    conn = get_db()
    today = date.today().isoformat()
    start_date = (date.today() - timedelta(days=days - 1)).isoformat()

    daily_stats = [dict(r) for r in conn.execute("""
        SELECT task_date,
               COUNT(DISTINCT shop_id) as total_shops,
               SUM(CASE WHEN is_completed=1 THEN 1 ELSE 0 END) as done_count,
               COUNT(*) as total_count
        FROM daily_tasks
        WHERE task_date >= ? AND task_date <= ?
        GROUP BY task_date ORDER BY task_date
    """, (start_date, today)).fetchall()]

    for d in daily_stats:
        d['completion_rate'] = round(d['done_count'] / d['total_count'] * 100) if d['total_count'] > 0 else 0

    new_shops_progress = [dict(r) for r in conn.execute("""
        SELECT s.shop_name,
               COUNT(*) as total,
               SUM(CASE WHEN ns.is_completed=1 THEN 1 ELSE 0 END) as done
        FROM new_shop_tasks ns
        JOIN shops s ON ns.shop_id=s.id
        WHERE s.status='new'
        GROUP BY s.id
    """).fetchall()]

    for ns in new_shops_progress:
        ns['progress'] = round(ns['done'] / ns['total'] * 100) if ns['total'] > 0 else 0

    all_groups = [r['group_name'] for r in conn.execute(
        "SELECT DISTINCT group_name FROM shops WHERE group_name != '' ORDER BY group_name"
    ).fetchall()]

    license_stats = [dict(r) for r in conn.execute("""
        SELECT l.license_name, COUNT(s.id) as shop_count
        FROM licenses l LEFT JOIN shops s ON s.license_id=l.id
        GROUP BY l.id ORDER BY shop_count DESC
    """).fetchall()]

    conn.close()
    return templates.TemplateResponse(request, "stats.html", {
        "request": request, "daily_stats": daily_stats, "days": days,
        "new_shops_progress": new_shops_progress, "license_stats": license_stats,
        "all_groups": all_groups, "today": today
    })


# ==================== 启动 ====================

if __name__ == "__main__":
    import uvicorn

    LAST_ACTIVITY = [time.time()]
    IDLE_TIMEOUT = 300

    @app.middleware("http")
    async def idle_timeout_middleware(request: Request, call_next):
        LAST_ACTIVITY[0] = time.time()
        response = await call_next(request)
        return response

    def check_idle():
        while True:
            time.sleep(10)
            if time.time() - LAST_ACTIVITY[0] > IDLE_TIMEOUT:
                print(f"空闲超过 {IDLE_TIMEOUT} 秒，自动关闭服务...")
                os._exit(0)

    t = threading.Thread(target=check_idle, daemon=True)
    t.start()

    uvicorn.run(app, host="0.0.0.0", port=8002)
