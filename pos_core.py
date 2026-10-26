# -*- coding: utf-8 -*-
"""Mobile Shop POS - database + business logic (no UI)."""
import sqlite3, os, zipfile, shutil, datetime as dt, hashlib, csv, smtplib, calendar, glob, html, json, hmac, secrets, base64, io, time, re
from email.message import EmailMessage

try:
    import openpyxl
except Exception:
    openpyxl = None
try:
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
except Exception:
    Credentials = None

APP_NAME = "MobilePOS_V5"
DATA_DIR = os.environ.get("MOBILEPOS_DATA") or os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
BACKUP_DIR = os.environ.get("MOBILEPOS_BACKUPS") or os.path.join(
    os.path.expanduser("~"), "Documents", APP_NAME + "_Backups")
RECEIPT_DIR = os.path.join(DATA_DIR, "receipts")
for _d in (DATA_DIR, BACKUP_DIR, RECEIPT_DIR):
    os.makedirs(_d, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "pos.db")

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY, name TEXT NOT NULL, category TEXT, barcode TEXT UNIQUE,
  cost REAL DEFAULT 0, price REAL DEFAULT 0, qty INTEGER DEFAULT 0, min_qty INTEGER DEFAULT 2,
  track_imei INTEGER DEFAULT 0, active INTEGER DEFAULT 1, service INTEGER DEFAULT 0, warranty INTEGER DEFAULT 0, alfa_service_id INTEGER, alfa_card_id INTEGER);
CREATE TABLE IF NOT EXISTS imeis(id INTEGER PRIMARY KEY, product_id INTEGER, imei TEXT UNIQUE,
  status TEXT DEFAULT 'in_stock', cost REAL DEFAULT 0, sale_id INTEGER, supplier_id INTEGER);
CREATE TABLE IF NOT EXISTS customers(id INTEGER PRIMARY KEY, name TEXT NOT NULL, phone TEXT, address TEXT,
  notes TEXT, points INTEGER DEFAULT 0, created TEXT);
CREATE TABLE IF NOT EXISTS suppliers(id INTEGER PRIMARY KEY, name TEXT NOT NULL, phone TEXT, notes TEXT);
CREATE TABLE IF NOT EXISTS sales(id INTEGER PRIMARY KEY, date TEXT, customer_id INTEGER, subtotal REAL, discount REAL,
  points_used INTEGER DEFAULT 0, points_value REAL DEFAULT 0, total REAL, kind TEXT, due_date TEXT, notes TEXT,
  points_earned INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS sale_items(id INTEGER PRIMARY KEY, sale_id INTEGER, product_id INTEGER, name TEXT,
  qty INTEGER, price REAL, cost REAL, imei TEXT, warranty_until TEXT);
CREATE TABLE IF NOT EXISTS payments(id INTEGER PRIMARY KEY, date TEXT, sale_id INTEGER, customer_id INTEGER,
  amount REAL, method TEXT, note TEXT, kind TEXT, cur TEXT DEFAULT 'USD', amount_lbp REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS installments(id INTEGER PRIMARY KEY, sale_id INTEGER, no INTEGER, due_date TEXT,
  amount REAL, paid REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS returns_c(id INTEGER PRIMARY KEY, date TEXT, sale_id INTEGER, item_id INTEGER,
  product_id INTEGER, qty INTEGER, amount REAL, cost REAL, imei TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS purchases(id INTEGER PRIMARY KEY, date TEXT, supplier_id INTEGER, invoice_no TEXT,
  total REAL, notes TEXT);
CREATE TABLE IF NOT EXISTS purchase_items(id INTEGER PRIMARY KEY, purchase_id INTEGER, product_id INTEGER,
  name TEXT, qty INTEGER, cost REAL, imeis TEXT);
CREATE TABLE IF NOT EXISTS supplier_payments(id INTEGER PRIMARY KEY, date TEXT, supplier_id INTEGER, amount REAL,
  method TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS supplier_returns(id INTEGER PRIMARY KEY, date TEXT, supplier_id INTEGER, product_id INTEGER,
  name TEXT, qty INTEGER, amount REAL, imei TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS expenses(id INTEGER PRIMARY KEY, date TEXT, category TEXT, amount REAL, note TEXT);
CREATE TABLE IF NOT EXISTS offers(id INTEGER PRIMARY KEY, product_id INTEGER, price REAL, start TEXT, end TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS repairs(id INTEGER PRIMARY KEY, date TEXT, customer_id INTEGER, device TEXT, imei TEXT,
  issue TEXT, price REAL DEFAULT 0, cost REAL DEFAULT 0, warranty_days INTEGER DEFAULT 0, status TEXT, sale_id INTEGER);
CREATE TABLE IF NOT EXISTS closes(id INTEGER PRIMARY KEY, date TEXT, float REAL, expected REAL, counted REAL, diff REAL,
  cash REAL, card REAL, transfer REAL, cnt INTEGER, revenue REAL,
  float_lbp REAL DEFAULT 0, expected_lbp REAL DEFAULT 0, counted_lbp REAL DEFAULT 0, diff_lbp REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS alfa_services(id INTEGER PRIMARY KEY, name TEXT NOT NULL, alfa REAL DEFAULT 0, price REAL DEFAULT 0, active INTEGER DEFAULT 1, fav INTEGER DEFAULT 0, provider TEXT DEFAULT 'alfa');
CREATE TABLE IF NOT EXISTS alfa_ledger(id INTEGER PRIMARY KEY, date TEXT, kind TEXT, amount REAL, cost REAL, note TEXT, ref_id INTEGER, provider TEXT DEFAULT 'alfa');
CREATE TABLE IF NOT EXISTS validity_sales(id INTEGER PRIMARY KEY, date TEXT, months INTEGER, price REAL, ret REAL, cards_cost REAL,
  cost_ret REAL, cards TEXT, names TEXT, sale_id INTEGER, provider TEXT DEFAULT 'alfa');
CREATE TABLE IF NOT EXISTS alfa_sales(id INTEGER PRIMARY KEY, date TEXT, kind TEXT, name TEXT, amt REAL, fee REAL, ded REAL,
  cost REAL, price REAL, profit REAL, sale_id INTEGER, provider TEXT DEFAULT 'alfa');
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, role TEXT DEFAULT 'cashier', pin TEXT, active INTEGER DEFAULT 1, username TEXT);
CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY, date TEXT, user TEXT, action TEXT, detail TEXT);
CREATE TABLE IF NOT EXISTS rate_history(id INTEGER PRIMARY KEY, date TEXT, rate REAL, user TEXT);
CREATE TABLE IF NOT EXISTS cash_moves(id INTEGER PRIMARY KEY, date TEXT, kind TEXT, amount REAL, cur TEXT DEFAULT 'USD', reason TEXT, user TEXT);
CREATE TABLE IF NOT EXISTS recurring(id INTEGER PRIMARY KEY, category TEXT, amount REAL, day INTEGER DEFAULT 1, note TEXT, last_ym TEXT, active INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS validity_presets(id INTEGER PRIMARY KEY, name TEXT, months INTEGER, cards TEXT, ret REAL, price REAL, provider TEXT DEFAULT 'alfa');
CREATE TABLE IF NOT EXISTS alfa_cards(id INTEGER PRIMARY KEY, name TEXT NOT NULL, cost REAL DEFAULT 0, active INTEGER DEFAULT 1, price REAL DEFAULT 0, provider TEXT DEFAULT 'alfa');
CREATE TABLE IF NOT EXISTS tradeins(id INTEGER PRIMARY KEY, date TEXT, product_id INTEGER, imei TEXT, model TEXT, cost REAL, price REAL,
  seller_name TEXT, seller_phone TEXT, id_note TEXT, condition TEXT, notes TEXT, user TEXT);
CREATE TABLE IF NOT EXISTS claims(id INTEGER PRIMARY KEY, date TEXT, sale_id INTEGER, imei TEXT, customer TEXT, phone TEXT, product TEXT, warranty_until TEXT,
  issue TEXT, status TEXT DEFAULT 'Open', supplier TEXT, resolution TEXT, notes TEXT, closed TEXT, user TEXT);
CREATE TABLE IF NOT EXISTS price_undo(id INTEGER PRIMARY KEY, product_id INTEGER, old_price REAL, batch TEXT);
CREATE TABLE IF NOT EXISTS images(id INTEGER PRIMARY KEY, kind TEXT, ref TEXT, mime TEXT, data BLOB, ver INTEGER DEFAULT 1, UNIQUE(kind, ref));
CREATE TABLE IF NOT EXISTS topups(id INTEGER PRIMARY KEY, sale_id INTEGER, name TEXT, price REAL, cost REAL, profit REAL, player TEXT, date TEXT);
CREATE TABLE IF NOT EXISTS topup_items(id INTEGER PRIMARY KEY, name TEXT UNIQUE, price REAL, profit REAL, uses INTEGER DEFAULT 0, last TEXT);
CREATE INDEX IF NOT EXISTS ix_sales_date ON sales(date);
CREATE INDEX IF NOT EXISTS ix_items_sale ON sale_items(sale_id);
CREATE INDEX IF NOT EXISTS ix_pay_sale ON payments(sale_id);
CREATE INDEX IF NOT EXISTS ix_imei_prod ON imeis(product_id);
"""

DEFAULTS = {
    "shop_name": "My Mobile Shop", "shop_phone": "", "shop_address": "", "currency": "$",
    "points_per": "10", "point_value": "0.10", "pin": "", "auto_backup": "1", "keep_backups": "30",
    "gdrive_folder": "", "auto_drive": "0", "email_to": "", "smtp_host": "smtp-mail.outlook.com",
    "smtp_port": "587", "country_code": "961", "smtp_user": "", "smtp_pass": "", "allow_negative": "0",
    "receipt_footer": "Thank you for your visit!", "lbp_rate": "90000", "alfa_fee": "0.14", "logo": "", "accent": "",
    "owner_phone": "", "autolock_min": "10", "backup_password": "", "receipt_width": "80mm", "auto_summary": "0",
    "max_discount_pct": "10", "lang": "en",
    "touch_fee": "0", "owner_recovery": "", "pic_fit": "contain", "backup_hours": "3", "lan_mode": "0", "lan_port": "8765", "last_backup": "", "last_offsite": "",
}

db = sqlite3.connect(DB_PATH, check_same_thread=False)
db.row_factory = sqlite3.Row
db.executescript(SCHEMA)
db.commit()

def migrate():
    """Upgrade a database made by an older version: add any missing columns, keep all data."""
    mem = sqlite3.connect(":memory:"); mem.executescript(SCHEMA)
    for (t,) in mem.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall():
        have = {r[1] for r in db.execute(f"PRAGMA table_info({t})")}
        for _, name, typ, _, dflt, _ in mem.execute(f"PRAGMA table_info({t})").fetchall():
            if name not in have:
                db.execute(f"ALTER TABLE {t} ADD COLUMN {name} {typ}" + (f" DEFAULT {dflt}" if dflt is not None else ""))
    db.execute("UPDATE repairs SET status='In repair' WHERE status='In progress'")
    db.commit(); mem.close()

migrate()

# ---------------------------------------------------------------- helpers
def q(sql, a=()): return db.execute(sql, a).fetchall()
def q1(sql, a=()): return db.execute(sql, a).fetchone()
def x(sql, a=()):
    cur = db.execute(sql, a); db.commit(); return cur.lastrowid
def val(sql, a=()):
    r = db.execute(sql, a).fetchone()
    return (r[0] if r and r[0] is not None else 0)
def today(): return dt.date.today().isoformat()
def now(): return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def parse_date(s):
    try: return dt.datetime.strptime(s, "%Y-%m-%d").date()
    except Exception: raise ValueError(f"Invalid date '{s}' (use YYYY-MM-DD)")
def add_months(d, n):
    y = d.year + (d.month - 1 + n) // 12; m = (d.month - 1 + n) % 12 + 1
    return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1]))

def S(k):
    r = db.execute("SELECT v FROM settings WHERE k=?", (k,)).fetchone()
    return r[0] if r else DEFAULTS.get(k, "")
def set_S(k, v):
    db.execute("INSERT INTO settings(k,v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v", (k, str(v)))
    db.commit()
def fS(k):
    try: return float(S(k))
    except Exception: return 0.0
def money(v): return f"{S('currency')}{(v or 0):,.2f}"
CURRENT = {"user": "", "role": "owner"}
def hash_pin(p, iters=120000):
    salt = secrets.token_hex(8)
    return f"pbkdf2${iters}${salt}${hashlib.pbkdf2_hmac('sha256', p.encode(), salt.encode(), iters).hex()}"
def verify_pin(p, stored):
    if not stored: return False
    if stored.startswith("pbkdf2$"):
        _, it, salt, h = stored.split("$")
        return hmac.compare_digest(hashlib.pbkdf2_hmac("sha256", p.encode(), salt.encode(), int(it)).hex(), h)
    return hmac.compare_digest(hashlib.sha256(p.encode()).hexdigest(), stored)   # older PIN format

# balance of a sale = total - returned - paid   (alias s must be the sales table)
BAL = ("(s.total - COALESCE((SELECT SUM(amount) FROM payments p WHERE p.sale_id=s.id),0)"
       " - COALESCE((SELECT SUM(amount) FROM returns_c r WHERE r.sale_id=s.id),0))")

# ---------------------------------------------------------------- products
def sync_qty(pid):
    if val("SELECT track_imei FROM products WHERE id=?", (pid,)):
        db.execute("UPDATE products SET qty=(SELECT COUNT(*) FROM imeis WHERE product_id=? AND status='in_stock') "
                   "WHERE id=?", (pid, pid))

def active_offer(pid):
    d = today()
    r = q1("SELECT price FROM offers WHERE product_id=? AND start<=? AND end>=? ORDER BY price LIMIT 1", (pid, d, d))
    return r[0] if r else None

def save_product(r, pid=None):
    track = 1 if r.get("track_imei") in (1, True, "1", "Yes") else 0
    service = 1 if r.get("service") in (1, True, "1", "Yes") else 0
    if service: track = 0
    bc = (r.get("barcode") or "").strip() or None
    name = (r.get("name") or "").strip()
    if not name: raise ValueError("Product name is required")
    f = lambda k: float(r.get(k) or 0)
    i = lambda k: int(float(r.get(k) or 0))
    try:
        if pid:
            db.execute("UPDATE products SET name=?,category=?,barcode=?,cost=?,price=?,min_qty=?,track_imei=?,service=?,warranty=? WHERE id=?",
                       (name, r.get("category") or "", bc, f("cost"), f("price"), i("min_qty"), track, service, i("warranty"), pid))
            if track: sync_qty(pid)
            elif service: db.execute("UPDATE products SET qty=0 WHERE id=?", (pid,))
            else: db.execute("UPDATE products SET qty=? WHERE id=?", (i("qty"), pid))
        else:
            pid = db.execute("INSERT INTO products(name,category,barcode,cost,price,qty,min_qty,track_imei,service,warranty) VALUES(?,?,?,?,?,?,?,?,?,?)",
                             (name, r.get("category") or "", bc, f("cost"), f("price"), 0 if (track or service) else i("qty"),
                              i("min_qty") if r.get("min_qty") not in (None, "") else 2, track, service, i("warranty"))).lastrowid
            if not bc: db.execute("UPDATE products SET barcode=? WHERE id=?", (f"2{pid:09d}", pid))
        db.commit()
    except sqlite3.IntegrityError:
        db.rollback(); raise ValueError("This barcode is already used by another product")
    return pid

def add_imeis(pid, imeis, cost=0, supplier_id=None):
    for im in imeis:
        im = im.strip()
        if not im: continue
        try:
            db.execute("INSERT INTO imeis(product_id,imei,status,cost,supplier_id) VALUES(?,?,?,?,?)",
                       (pid, im, "in_stock", cost, supplier_id))
        except sqlite3.IntegrityError:
            db.rollback(); raise ValueError(f"IMEI {im} already exists")
    sync_qty(pid); db.commit()

def find_exact(text):
    """barcode or IMEI scan -> (product_row, imei or None) or None"""
    r = q1("SELECT * FROM products WHERE barcode=? AND active=1", (text,))
    if r: return r, None
    r = q1("SELECT product_id FROM imeis WHERE imei=? AND status='in_stock'", (text,))
    if r: return q1("SELECT * FROM products WHERE id=?", (r[0],)), text
    return None

def search_products(text):
    t = f"%{text}%"
    return q("SELECT * FROM products WHERE active=1 AND (name LIKE ? OR barcode LIKE ? OR category LIKE ? OR id IN "
             "(SELECT product_id FROM imeis WHERE imei LIKE ? AND status='in_stock')) ORDER BY name LIMIT 60", (t, t, t, t))

# ---------------------------------------------------------------- sales
def create_sale(cart, customer_id, discount, points_used, paid, method, kind, due, notes, inst_n=3, cur="USD", lbp=0):
    if not cart: raise ValueError("Cart is empty")
    subtotal = round(sum(i["qty"] * i["price"] for i in cart), 2)
    pv = round(points_used * fS("point_value"), 2)
    if points_used:
        if not customer_id or points_used > val("SELECT points FROM customers WHERE id=?", (customer_id,)):
            raise ValueError("Not enough loyalty points")
    total = round(max(0.0, subtotal - discount - pv), 2)
    paid = total if kind == "cash" else round(min(max(paid, 0), total), 2)
    if cur == "LBP" and (kind == "cash" or not lbp): lbp = round(paid * fS("lbp_rate"))
    if cur != "LBP": lbp = 0
    balance = round(total - paid, 2)
    if balance <= 0.004: kind, balance = "cash", 0
    elif not customer_id: raise ValueError("Choose a customer for debt / installment sales")
    if balance > 0 and due: parse_date(due)
    need = {}
    for i in cart:
        if i.get("pid"): need[i["pid"]] = need.get(i["pid"], 0) + i["qty"]
    for pid, n in need.items():
        p = q1("SELECT * FROM products WHERE id=?", (pid,))
        if p["service"]: continue
        if p["track_imei"] and any(not c.get("imei") for c in cart if c.get("pid") == pid):
            raise ValueError(f"Select an IMEI for {p['name']}")
        if p["qty"] < n and S("allow_negative") != "1":
            raise ValueError(f"Not enough stock for {p['name']} (available {p['qty']})")
    card_lines = {}
    alfa_lines, need_p, avgs = {}, {}, {}       # service products take credit from the Alfa / Touch stock
    for idx, i in enumerate(cart):
        if i.get("pid"):
            sv = q1("SELECT sv.* FROM alfa_services sv JOIN products p ON p.alfa_service_id=sv.id WHERE p.id=?", (i["pid"],))
            if sv: alfa_lines[idx] = sv; need_p[sv["provider"]] = need_p.get(sv["provider"], 0.0) + sv["alfa"] * i["qty"]
            cd = q1("SELECT c.* FROM alfa_cards c JOIN products p ON p.alfa_card_id=c.id WHERE p.id=?", (i["pid"],))
            if cd: card_lines[idx] = cd
    for pv_, tot_ in need_p.items():
        st = alfa_state(pv_)
        if tot_ > st["amount"] + 0.00001: raise ValueError(f"Not enough {pname(pv_)} in stock ({st['amount']:.2f}$ left, needs {tot_:.2f}$)")
        avgs[pv_] = st["avg"]
    per = fS("points_per"); earned = int(total // per) if customer_id and per > 0 else 0
    with db:
        sid = db.execute("INSERT INTO sales(date,customer_id,subtotal,discount,points_used,points_value,total,kind,"
                         "due_date,notes,points_earned) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                         (now(), customer_id or None, subtotal, discount, points_used, pv, total, kind,
                          due or None, notes, earned)).lastrowid
        for idx, i in enumerate(cart):
            cost = i["cost"]; pid_ = i.get("pid") or None
            if idx in alfa_lines: cost = alfa_lines[idx]["alfa"] * avgs[alfa_lines[idx]["provider"]]
            if idx in card_lines: cost = card_lines[idx]["cost"]
            p = db.execute("SELECT service,warranty,track_imei FROM products WHERE id=?", (pid_,)).fetchone() if pid_ else None
            days = int(i.get("warranty") or (p["warranty"] if p else 0) or 0)
            wu = (dt.date.today() + dt.timedelta(days=days)).isoformat() if days > 0 else None
            if i.get("imei") and pid_:
                r = db.execute("SELECT cost FROM imeis WHERE imei=?", (i["imei"],)).fetchone()
                if r and r[0] and r[0] > 0: cost = r[0]
                db.execute("UPDATE imeis SET status='sold', sale_id=? WHERE imei=?", (sid, i["imei"]))
            db.execute("INSERT INTO sale_items(sale_id,product_id,name,qty,price,cost,imei,warranty_until) VALUES(?,?,?,?,?,?,?,?)",
                       (sid, pid_, i["name"], i["qty"], i["price"], cost, i.get("imei") or None, wu))
            if idx in card_lines:
                db.execute("INSERT INTO alfa_sales(date,kind,name,amt,fee,ded,cost,price,profit,sale_id,provider) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                           (now(), "card", i["name"], 0, 0, 0, cost * i["qty"], i["price"] * i["qty"], (i["price"] - cost) * i["qty"], sid, card_lines[idx]["provider"]))
            if idx in alfa_lines:
                sv = alfa_lines[idx]; ded = sv["alfa"] * i["qty"]; avg = avgs[sv["provider"]]
                db.execute("INSERT INTO alfa_ledger(date,kind,amount,cost,note,ref_id,provider) VALUES(?,?,?,?,?,?,?)", (now(), "sale", -ded, -ded * avg, sv["name"], sid, sv["provider"]))
                db.execute("INSERT INTO alfa_sales(date,kind,name,amt,fee,ded,cost,price,profit,sale_id,provider) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                           (now(), "svc", sv["name"], sv["alfa"], 0, ded, ded * avg, i["price"] * i["qty"], i["price"] * i["qty"] - ded * avg, sid, sv["provider"]))
            if pid_ and p and not p["service"]:
                if p["track_imei"]: sync_qty(pid_)
                else: db.execute("UPDATE products SET qty=qty-? WHERE id=?", (i["qty"], pid_))
        if customer_id:
            db.execute("UPDATE customers SET points=points-?+? WHERE id=?", (points_used, earned, customer_id))
        if paid > 0:
            db.execute("INSERT INTO payments(date,sale_id,customer_id,amount,method,note,kind,cur,amount_lbp) VALUES(?,?,?,?,?,?,?,?,?)",
                       (now(), sid, customer_id or None, paid, method, "", "initial", cur, lbp))
        if kind == "installment" and inst_n > 0:
            first = parse_date(due) if due else add_months(dt.date.today(), 1)
            base = round(balance / inst_n, 2)
            for n in range(1, inst_n + 1):
                amt = base if n < inst_n else round(balance - base * (inst_n - 1), 2)
                db.execute("INSERT INTO installments(sale_id,no,due_date,amount) VALUES(?,?,?,?)",
                           (sid, n, add_months(first, n - 1).isoformat(), amt))
    for pv_ in need_p: _alfa_zero_out(pv_)
    return sid

def sale_balance(sid): return val(f"SELECT {BAL} FROM sales s WHERE s.id=?", (sid,))

def add_payment(sale_id, amount, method="Cash", note="", cur="USD", lbp=0):
    if amount <= 0: raise ValueError("Amount must be positive")
    bal = sale_balance(sale_id)
    if amount > bal + 0.004:
        if cur == "LBP" and amount - bal <= 0.01: amount = round(bal, 2)
        else: raise ValueError("Amount is more than the remaining balance")
    if cur != "LBP": lbp = 0
    cid = val("SELECT customer_id FROM sales WHERE id=?", (sale_id,)) or None
    with db:
        db.execute("INSERT INTO payments(date,sale_id,customer_id,amount,method,note,kind,cur,amount_lbp) VALUES(?,?,?,?,?,?,?,?,?)",
                   (now(), sale_id, cid, amount, method, note, "payment", cur, lbp))
        left = amount
        for r in db.execute("SELECT id,amount,paid FROM installments WHERE sale_id=? AND paid<amount-0.004 "
                            "ORDER BY due_date,no", (sale_id,)).fetchall():
            take = min(left, r["amount"] - r["paid"])
            db.execute("UPDATE installments SET paid=paid+? WHERE id=?", (take, r["id"]))
            left -= take
            if left <= 0.004: break

def customer_balance(cid):
    return val(f"SELECT SUM(b) FROM (SELECT {BAL} b FROM sales s WHERE s.customer_id=?) WHERE b>0.004", (cid,))

def customer_pay(cid, amount, method="Cash", note="", cur="USD", lbp=0):
    debts = q(f"SELECT s.id, {BAL} AS bal FROM sales s WHERE s.customer_id=? AND {BAL}>0.004 ORDER BY s.date", (cid,))
    if amount <= 0: raise ValueError("Amount must be positive")
    tot = sum(r["bal"] for r in debts)
    if amount > tot + 0.004:
        if cur == "LBP" and amount - tot <= 0.01: amount = round(tot, 2)
        else: raise ValueError("Amount is more than the customer's debt")
    left = amount
    for r in debts:
        take = round(min(left, r["bal"]), 2)
        add_payment(r["id"], take, method, note, cur, (lbp * take / amount) if (lbp and amount) else 0)
        left -= take
        if left <= 0.004: break

def _alfa_return(it, s, qty, amount):
    # a returned Alfa / Touch service gives its credit back to the stock at the original cost; the activity log gets a negative line
    sv = cd = None
    if it["product_id"]:
        sv = q1("SELECT sv.* FROM alfa_services sv JOIN products p ON p.alfa_service_id=sv.id WHERE p.id=?", (it["product_id"],))
        cd = q1("SELECT c.* FROM alfa_cards c JOIN products p ON p.alfa_card_id=c.id WHERE p.id=?", (it["product_id"],))
    if sv: ded, cost, provider, name, amt = sv["alfa"] * qty, it["cost"] * qty, sv["provider"], sv["name"], sv["alfa"]
    elif cd: ded, cost, provider, name, amt = 0.0, it["cost"] * qty, cd["provider"], it["name"], 0
    else:
        row = q1("SELECT * FROM alfa_sales WHERE sale_id=? AND kind IN ('svc','usd','card') ORDER BY id LIMIT 1", (s["id"],))
        if not row: return
        ratio = qty / it["qty"] if it["qty"] else 1
        ded, cost, provider, name, amt = row["ded"] * ratio, row["cost"] * ratio, row["provider"], row["name"], row["amt"]
    if ded > 0:
        db.execute("INSERT INTO alfa_ledger(date,kind,amount,cost,note,ref_id,provider) VALUES(?,?,?,?,?,?,?)", (now(), "return", ded, cost, "Return: " + name, s["id"], provider))
    db.execute("INSERT INTO alfa_sales(date,kind,name,amt,fee,ded,cost,price,profit,sale_id,provider) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
               (now(), "return", "Return: " + name, amt, 0, -ded, -cost, -amount, -(amount - cost), s["id"], provider))

def return_item(item_id, qty, note=""):
    it = q1("SELECT * FROM sale_items WHERE id=?", (item_id,))
    s = q1("SELECT * FROM sales WHERE id=?", (it["sale_id"],))
    done = val("SELECT SUM(qty) FROM returns_c WHERE item_id=?", (item_id,))
    if qty <= 0 or qty + done > it["qty"]: raise ValueError("Invalid return quantity")
    if val("SELECT COUNT(*) FROM validity_sales WHERE sale_id=?", (it["sale_id"],)): raise ValueError("A Validity days sale cannot be returned - use Void instead")
    ratio = (s["total"] / s["subtotal"]) if s["subtotal"] else 1
    amount = round(qty * it["price"] * ratio, 2)
    newbal = sale_balance(s["id"]) - amount
    with db:
        db.execute("INSERT INTO returns_c(date,sale_id,item_id,product_id,qty,amount,cost,imei,note) VALUES(?,?,?,?,?,?,?,?,?)",
                   (now(), s["id"], item_id, it["product_id"], qty, amount, it["cost"], it["imei"], note))
        _alfa_return(it, s, qty, amount)
        if it["imei"] and it["product_id"]:
            db.execute("UPDATE imeis SET status='in_stock', sale_id=NULL WHERE imei=?", (it["imei"],))
            sync_qty(it["product_id"])
        elif it["product_id"]:
            db.execute("UPDATE products SET qty=qty+? WHERE id=? AND service=0", (qty, it["product_id"]))
        if newbal < -0.004:   # customer already paid -> cash refund
            db.execute("INSERT INTO payments(date,sale_id,customer_id,amount,method,note,kind) VALUES(?,?,?,?,?,?,?)",
                       (now(), s["id"], s["customer_id"], newbal, "Cash", note or "Refund", "refund"))
    return amount

def void_sale(sid):
    s = q1("SELECT * FROM sales WHERE id=?", (sid,))
    if not s: return
    with db:
        for it in db.execute("SELECT * FROM sale_items WHERE sale_id=?", (sid,)).fetchall():
            back = it["qty"] - (db.execute("SELECT COALESCE(SUM(qty),0) FROM returns_c WHERE item_id=?", (it["id"],)).fetchone()[0])
            if it["imei"] and it["product_id"]:
                db.execute("UPDATE imeis SET status='in_stock', sale_id=NULL WHERE imei=?", (it["imei"],))
                sync_qty(it["product_id"])
            elif back > 0 and it["product_id"]:
                db.execute("UPDATE products SET qty=qty+? WHERE id=? AND service=0", (back, it["product_id"]))
        if s["customer_id"]:
            db.execute("UPDATE customers SET points=MAX(0,points+?-?) WHERE id=?",
                       (s["points_used"], s["points_earned"], s["customer_id"]))
        for v in db.execute("SELECT cards FROM validity_sales WHERE sale_id=?", (sid,)).fetchall():
            for cpid, n in json.loads(v["cards"] or "{}").items():
                db.execute("UPDATE products SET qty=qty+? WHERE id=?", (n, int(cpid)))
        db.execute("DELETE FROM alfa_ledger WHERE ref_id=? AND kind IN ('validity','sale','return')", (sid,))
        db.execute("DELETE FROM validity_sales WHERE sale_id=?", (sid,))
        db.execute("DELETE FROM alfa_sales WHERE sale_id=?", (sid,))
        db.execute("DELETE FROM topups WHERE sale_id=?", (sid,))
        for t in ("payments", "installments", "returns_c", "sale_items"):
            db.execute(f"DELETE FROM {t} WHERE sale_id=?", (sid,))
        db.execute("DELETE FROM sales WHERE id=?", (sid,))

# ---------------------------------------------------------------- suppliers
def supplier_balance(sid):
    return (val("SELECT SUM(total) FROM purchases WHERE supplier_id=?", (sid,))
            - val("SELECT SUM(amount) FROM supplier_payments WHERE supplier_id=?", (sid,))
            - val("SELECT SUM(amount) FROM supplier_returns WHERE supplier_id=?", (sid,)))

def total_payable():
    return sum(max(0, supplier_balance(r[0])) for r in q("SELECT id FROM suppliers"))

def add_purchase(supplier_id, invoice_no, date, items, paid, method, notes):
    if not items: raise ValueError("Add at least one item")
    parse_date(date)
    total = round(sum(i["qty"] * i["cost"] for i in items), 2)
    try:
        with db:
            pid = db.execute("INSERT INTO purchases(date,supplier_id,invoice_no,total,notes) VALUES(?,?,?,?,?)",
                             (date, supplier_id, invoice_no, total, notes)).lastrowid
            for i in items:
                ims = i.get("imeis") or []
                db.execute("INSERT INTO purchase_items(purchase_id,product_id,name,qty,cost,imeis) VALUES(?,?,?,?,?,?)",
                           (pid, i["pid"], i["name"], i["qty"], i["cost"], ",".join(ims)))
                if ims:
                    for im in ims:
                        db.execute("INSERT INTO imeis(product_id,imei,status,cost,supplier_id) VALUES(?,?,?,?,?)",
                                   (i["pid"], im, "in_stock", i["cost"], supplier_id))
                    db.execute("UPDATE products SET cost=? WHERE id=?", (i["cost"], i["pid"]))
                    sync_qty(i["pid"])
                else:
                    db.execute("UPDATE products SET qty=qty+?, cost=? WHERE id=?", (i["qty"], i["cost"], i["pid"]))
            if paid > 0:
                db.execute("INSERT INTO supplier_payments(date,supplier_id,amount,method,note) VALUES(?,?,?,?,?)",
                           (date, supplier_id, paid, method, f"Invoice {invoice_no or pid}"))
    except sqlite3.IntegrityError:
        raise ValueError("One of the IMEIs already exists in the system")
    return pid

def pay_supplier(supplier_id, amount, method, note, date=None):
    if amount <= 0: raise ValueError("Amount must be positive")
    x("INSERT INTO supplier_payments(date,supplier_id,amount,method,note) VALUES(?,?,?,?,?)",
      (date or today(), supplier_id, amount, method, note))

def supplier_return(supplier_id, product_id, qty, amount, imei, note, date=None):
    p = q1("SELECT * FROM products WHERE id=?", (product_id,))
    with db:
        if p["track_imei"]:
            r = db.execute("SELECT id FROM imeis WHERE imei=? AND product_id=? AND status='in_stock'", (imei, product_id)).fetchone()
            if not r: raise ValueError("This IMEI is not in stock for that product")
            db.execute("UPDATE imeis SET status='returned' WHERE id=?", (r[0],)); qty = 1
            sync_qty(product_id)
        else:
            if p["qty"] < qty: raise ValueError("Not enough stock to return")
            db.execute("UPDATE products SET qty=qty-? WHERE id=?", (qty, product_id))
        db.execute("INSERT INTO supplier_returns(date,supplier_id,product_id,name,qty,amount,imei,note) VALUES(?,?,?,?,?,?,?,?)",
                   (date or today(), supplier_id, product_id, p["name"], qty, amount, imei or None, note))

# ---------------------------------------------------------------- reports
def total_debts():
    return val(f"SELECT SUM(b) FROM (SELECT {BAL} b FROM sales s) WHERE b>0.004")

def period_summary(d1, d2):
    r = (d1, d2)
    D = "substr(date,1,10) BETWEEN ? AND ?"
    s = {}
    s["rev"] = val(f"SELECT SUM(total) FROM sales WHERE {D}", r)
    s["ret"] = val(f"SELECT SUM(amount) FROM returns_c WHERE {D}", r)
    s["cogs"] = val("SELECT SUM(i.qty*i.cost) FROM sale_items i JOIN sales s ON s.id=i.sale_id "
                    "WHERE substr(s.date,1,10) BETWEEN ? AND ?", r)
    s["cogs_r"] = val(f"SELECT SUM(qty*cost) FROM returns_c WHERE {D}", r)
    s["exp"] = val("SELECT SUM(amount) FROM expenses WHERE date BETWEEN ? AND ?", r)
    s["cash"] = val(f"SELECT SUM(amount) FROM payments WHERE {D}", r)
    s["net_rev"] = s["rev"] - s["ret"]
    s["gross"] = s["net_rev"] - (s["cogs"] - s["cogs_r"])
    s["net"] = s["gross"] - s["exp"]
    return s

def _r(v): return round(v or 0, 2)
D_ = "substr(s.date,1,10) BETWEEN ? AND ?"

def rep_pnl(d1, d2):
    s = period_summary(d1, d2)
    inv_c = val("SELECT SUM(qty*cost) FROM products WHERE active=1 AND qty>0")
    inv_p = val("SELECT SUM(qty*price) FROM products WHERE active=1 AND qty>0")
    rows = [["Revenue (sales)", _r(s["rev"])], ["Customer returns", -_r(s["ret"])], ["Net revenue", _r(s["net_rev"])],
            ["Cost of goods sold", -_r(s["cogs"] - s["cogs_r"])], ["GROSS PROFIT", _r(s["gross"])],
            ["Operating expenses", -_r(s["exp"])],
            ["NET PROFIT", _r(s["net"])], ["", ""], ["Cash collected from sales", _r(s["cash"])],
            ["Customer debts outstanding (now)", _r(total_debts())], ["Owed to suppliers (now)", _r(total_payable())],
            ["Inventory value at cost (now)", _r(inv_c)], ["Inventory value at retail (now)", _r(inv_p)]]
    return ["Item", "Amount"], rows

def rep_sales_day(d1, d2):
    rows = q(f"SELECT substr(s.date,1,10), COUNT(*), ROUND(SUM(s.total),2), ROUND(SUM(s.total)-SUM((SELECT COALESCE(SUM(qty*cost),0) "
             f"FROM sale_items i WHERE i.sale_id=s.id)),2) FROM sales s WHERE {D_} GROUP BY 1 ORDER BY 1", (d1, d2))
    return ["Date", "Invoices", "Sales", "Gross profit"], [list(r) for r in rows]

def rep_top(d1, d2):
    rows = q(f"SELECT i.name, SUM(i.qty), ROUND(SUM(i.qty*i.price),2), ROUND(SUM(i.qty*(i.price-i.cost)),2) FROM sale_items i "
             f"JOIN sales s ON s.id=i.sale_id WHERE {D_} GROUP BY i.name ORDER BY SUM(i.qty) DESC LIMIT 100", (d1, d2))
    return ["Product", "Qty sold", "Sales", "Profit"], [list(r) for r in rows]

def rep_inventory(d1=None, d2=None):
    rows = q("SELECT name, category, qty, cost, price, ROUND(qty*cost,2), ROUND(qty*price,2) FROM products WHERE active=1 ORDER BY category,name")
    return ["Product", "Category", "Qty", "Cost", "Price", "Value at cost", "Value at retail"], [list(r) for r in rows]

def rep_phones(d1=None, d2=None):
    rows = q("SELECT p.name, i.imei, i.cost, p.price FROM imeis i JOIN products p ON p.id=i.product_id WHERE i.status='in_stock' ORDER BY p.name")
    return ["Phone", "IMEI", "Cost", "Price"], [list(r) for r in rows]

def rep_debts(d1=None, d2=None):
    rows = [[r["name"], r["phone"], round(customer_balance(r["id"]), 2)] for r in q("SELECT * FROM customers ORDER BY name")]
    return ["Customer", "Phone", "Debt"], [r for r in rows if r[2] > 0]

def rep_suppliers(d1=None, d2=None):
    rows = []
    for r in q("SELECT * FROM suppliers ORDER BY name"):
        rows.append([r["name"], _r(val("SELECT SUM(total) FROM purchases WHERE supplier_id=?", (r["id"],))),
                     _r(val("SELECT SUM(amount) FROM supplier_payments WHERE supplier_id=?", (r["id"],))),
                     _r(val("SELECT SUM(amount) FROM supplier_returns WHERE supplier_id=?", (r["id"],))), _r(supplier_balance(r["id"]))])
    return ["Supplier", "Purchases", "Paid", "Returns", "Balance owed"], rows

def rep_expenses(d1, d2):
    rows = q("SELECT category, COUNT(*), ROUND(SUM(amount),2) FROM expenses WHERE date BETWEEN ? AND ? GROUP BY category ORDER BY 3 DESC", (d1, d2))
    return ["Category", "Count", "Total"], [list(r) for r in rows]

def rep_low(d1=None, d2=None):
    rows = q("SELECT name, category, qty, min_qty FROM products WHERE active=1 AND qty<=min_qty ORDER BY qty")
    return ["Product", "Category", "Qty", "Alert level"], [list(r) for r in rows]

def rep_cash_method(d1, d2):
    rows = q("SELECT method, COUNT(*), ROUND(SUM(amount),2) FROM payments WHERE substr(date,1,10) BETWEEN ? AND ? GROUP BY method", (d1, d2))
    return ["Method", "Payments", "Total"], [list(r) for r in rows]

def rep_installments(d1=None, d2=None):
    rows = q("SELECT c.name, c.phone, i.sale_id, i.no, i.due_date, i.amount, i.paid, ROUND(i.amount-i.paid,2) FROM installments i "
             "JOIN sales s ON s.id=i.sale_id JOIN customers c ON c.id=s.customer_id WHERE i.paid<i.amount-0.004 ORDER BY i.due_date")
    return ["Customer", "Phone", "Invoice", "#", "Due", "Amount", "Paid", "Remaining"], [list(r) for r in rows]

def _dash_cat(it, alfa, isval):
    cat = it["pcat"] or ""; touch = cat.startswith("Touch")
    if it["alfa_service_id"]: return "Touch services" if touch else "Alfa services"
    if it["alfa_card_id"]: return "Touch cards" if touch else "Recharge cards"
    if not it["product_id"]:
        if val("SELECT COUNT(*) FROM topups WHERE sale_id=? AND name=?", (it["sale_id"], it["name"])): return "Game cards"
        if isval: return "Validity days"
        if alfa:
            t = alfa["provider"] == "touch"
            return ("Touch services" if t else "Alfa services") if alfa["kind"] in ("svc", "usd") else ("Touch cards" if t else "Recharge cards")
        return "Other"
    return cat or "Other"

def rep_dashboard(d1, d2):
    # one row per item per day, in the format the "A-TECH shop sales" dashboard reads
    agg = {}
    for s in q("SELECT * FROM sales WHERE substr(date,1,10) BETWEEN ? AND ? ORDER BY id", (d1, d2)):
        ratio = (s["total"] / s["subtotal"]) if s["subtotal"] else 1
        alfa = q1("SELECT kind, provider FROM alfa_sales WHERE sale_id=? AND kind!='return' ORDER BY id LIMIT 1", (s["id"],))
        isval = q1("SELECT 1 FROM validity_sales WHERE sale_id=?", (s["id"],))
        lines = []
        for it in q("SELECT i.*, p.category pcat, p.alfa_service_id, p.alfa_card_id FROM sale_items i LEFT JOIN products p ON p.id=i.product_id WHERE i.sale_id=?", (s["id"],)):
            rt = q1("SELECT COALESCE(SUM(qty),0) rq, COALESCE(SUM(amount),0) ra, COALESCE(SUM(cost*qty),0) rc FROM returns_c WHERE item_id=?", (it["id"],))
            lines.append((_dash_cat(it, alfa, isval), it["name"], it["qty"] - rt["rq"], it["qty"] * it["price"] * ratio - rt["ra"], it["qty"] * it["cost"] - rt["rc"]))
        tot = sum(l[3] for l in lines)
        usd = val("SELECT SUM(amount) FROM payments WHERE sale_id=? AND COALESCE(cur,'USD')!='LBP'", (s["id"],))
        lbp = val("SELECT SUM(amount) FROM payments WHERE sale_id=? AND cur='LBP'", (s["id"],))
        debt = max(0.0, sale_balance(s["id"]))
        for cat, name, qty, rev, cost in lines:
            share = (rev / tot) if tot else 0
            a = agg.setdefault((s["date"][:10], cat, name), [0, 0.0, 0.0, 0.0, 0.0, 0.0])
            a[0] += qty; a[1] += rev; a[2] += cost; a[3] += usd * share; a[4] += lbp * share; a[5] += debt * share
    out = [[k[0], k[1], k[2], v[0], round(v[1], 2), round(v[2], 2), round(v[3], 2), round(v[4], 2), round(v[5], 2)] for k, v in sorted(agg.items())]
    return ["day", "category", "item", "qty", "revenue", "cost", "paid_usd", "paid_lbp_usd", "on_debt"], out

def rep_topups(d1, d2):
    rs = q("SELECT substr(date,1,10), name, player, price, cost, profit FROM topups WHERE substr(date,1,10) BETWEEN ? AND ? ORDER BY id", (d1, d2))
    return ["Date", "Item", "Player ID / order", "Price", "Cost", "Profit"], [[r[0], r[1], r[2], round(r[3], 2), round(r[4], 2), round(r[5], 2)] for r in rs]

def rep_inventory_full(d1=None, d2=None):
    out = []
    for p in q("SELECT * FROM products WHERE active=1 ORDER BY category, name"):
        kind = "Alfa/Touch card" if p["alfa_card_id"] else "Alfa/Touch service" if p["alfa_service_id"] else "Service" if p["service"] else "Phone (IMEI)" if p["track_imei"] else "Item"
        stock = 0 if p["service"] else p["qty"]
        out.append([p["name"], p["category"], p["barcode"], kind, round(p["cost"], 2), round(p["price"], 2), stock, p["min_qty"],
                    round(stock * p["cost"], 2), round(stock * p["price"], 2)])
    return ["Name", "Category", "Barcode", "Type", "Cost", "Price", "In stock", "Alert level", "Stock value at cost", "Stock value at price"], out

def rep_imeis(d1=None, d2=None):
    return ["Phone", "IMEI", "Cost", "Price"], [list(r) for r in q("SELECT p.name, i.imei, ROUND(i.cost,2), ROUND(p.price,2) FROM imeis i JOIN products p ON p.id=i.product_id WHERE i.status='in_stock' ORDER BY p.name, i.imei")]

def rep_tradeins(d1, d2):
    out = []
    for t in q("SELECT t.*, i.status, i.sale_id FROM tradeins t LEFT JOIN imeis i ON i.imei=t.imei WHERE substr(t.date,1,10) BETWEEN ? AND ? ORDER BY t.id", (d1, d2)):
        sold = val("SELECT SUM(price) FROM sale_items WHERE sale_id=? AND imei=?", (t["sale_id"], t["imei"])) if t["sale_id"] else None
        out.append([t["date"][:10], t["model"], t["imei"], t["seller_name"], round(t["cost"], 2), round(t["price"], 2), "Sold" if t["status"] == "sold" else "In stock", round(sold - t["cost"], 2) if sold else ""])
    return ["Date", "Phone", "IMEI", "Bought from", "Paid", "Asking price", "Status", "Profit"], out

def rep_claims(d1, d2):
    return ["Date", "Phone", "IMEI", "Customer", "Issue", "Status", "Resolution"], [list(r) for r in q("SELECT substr(date,1,10), product, imei, customer, issue, status, COALESCE(resolution,'') FROM claims WHERE substr(date,1,10) BETWEEN ? AND ? ORDER BY id", (d1, d2))]

def rep_profit_cat(d1, d2):
    rows_ = q(f"SELECT COALESCE(p.category,'Services & Alfa'), SUM(i.qty), ROUND(SUM(i.qty*i.price),2), ROUND(SUM(i.qty*(i.price-i.cost)),2) FROM sale_items i "
              f"JOIN sales s ON s.id=i.sale_id LEFT JOIN products p ON p.id=i.product_id WHERE {D_} GROUP BY 1 ORDER BY 4 DESC", (d1, d2))
    return ["Category", "Qty sold", "Sales", "Profit"], [list(r) for r in rows_]

def rep_dead(d1=None, d2=None):
    cutoff = (dt.date.today() - dt.timedelta(days=60)).isoformat(); out = []
    for p in q("SELECT * FROM products WHERE active=1 AND service=0 AND qty>0 ORDER BY name"):
        last = val("SELECT MAX(substr(s.date,1,10)) FROM sale_items i JOIN sales s ON s.id=i.sale_id WHERE i.product_id=?", (p["id"],)) or ""
        if not last or last < cutoff: out.append([p["name"], p["category"], p["qty"], round(p["qty"] * p["cost"], 2), last or "never sold"])
    return ["Product", "Category", "Qty", "Value at cost", "Last sold"], out

def rep_reorder(d1=None, d2=None):
    out = []
    for p in q("SELECT * FROM products WHERE active=1 AND service=0 AND qty<=min_qty ORDER BY name"):
        r = q1("SELECT sp.name sname, i.cost FROM purchase_items i JOIN purchases pu ON pu.id=i.purchase_id LEFT JOIN suppliers sp ON sp.id=pu.supplier_id "
               "WHERE i.product_id=? ORDER BY pu.date DESC, pu.id DESC LIMIT 1", (p["id"],))
        out.append([(r["sname"] if r and r["sname"] else "(no supplier yet)"), p["name"], p["qty"], p["min_qty"], max(1, p["min_qty"] * 2 - p["qty"]), round(r["cost"] if r else p["cost"], 2)])
    out.sort(key=lambda r: (r[0], r[1]))
    return ["Supplier", "Product", "In stock", "Alert level", "Suggested order", "Last cost"], out

def rep_cash_moves(d1, d2):
    rows_ = q("SELECT substr(date,1,16), CASE kind WHEN 'in' THEN 'Cash in' ELSE 'Cash out' END, amount, COALESCE(cur,'USD'), reason, user FROM cash_moves WHERE substr(date,1,10) BETWEEN ? AND ? ORDER BY id", (d1, d2))
    return ["Time", "Type", "Amount", "Currency", "Reason", "By"], [list(r) for r in rows_]

def rep_alfa(d1, d2):
    out = []
    for r in q("SELECT substr(date,1,10), CASE provider WHEN 'touch' THEN 'Touch ' ELSE '' END || 'Validity days', months||' MONTH - '||names, ret, ROUND(cost_ret,4), price, 0 FROM validity_sales WHERE substr(date,1,10) BETWEEN ? AND ?", (d1, d2)): out.append(list(r))
    for r in q("SELECT substr(date,1,10), CASE provider WHEN 'touch' THEN 'Touch ' ELSE 'Alfa ' END || CASE kind WHEN 'card' THEN 'card sale' WHEN 'return' THEN 'return' ELSE 'sale' END, name, -ded, ROUND(cost,4), price, ROUND(profit,4) FROM alfa_sales WHERE substr(date,1,10) BETWEEN ? AND ?", (d1, d2)): out.append(list(r))
    out.sort(key=lambda r: r[0])
    return ["Date", "Type", "Item", "Credit change", "Cost", "Price", "Profit"], out

REPORTS = {
    "Profit & Loss": rep_pnl, "Sales by day": rep_sales_day, "Top selling products": rep_top,
    "Cash by payment method": rep_cash_method, "Inventory valuation": rep_inventory,
    "Phones in stock (IMEI)": rep_phones, "Low stock": rep_low, "Customer debts": rep_debts,
    "Unpaid installments": rep_installments, "Sales for dashboard": rep_dashboard, "Top-up sales": rep_topups, "Inventory list": rep_inventory_full, "Phones in stock (IMEIs)": rep_imeis, "Trade-ins": rep_tradeins, "Warranty claims": rep_claims, "Alfa and Touch activity": rep_alfa,
    "Profit by category": rep_profit_cat, "Dead stock (60 days)": rep_dead, "Reorder list by supplier": rep_reorder, "Cash in / out": rep_cash_moves, "Supplier balances": rep_suppliers, "Expenses by category": rep_expenses,
}

def export_rows(path, title, cols, rows):
    if path.lower().endswith(".xlsx"):
        if openpyxl is None: raise RuntimeError("openpyxl is not installed - save as .csv instead")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = title[:30]
        ws.append([title]); ws.append([]); ws.append(cols)
        for r in rows: ws.append(r)
        for i, c in enumerate(cols):
            ws.column_dimensions[chr(65 + i)].width = max(14, len(str(c)) + 4)
        wb.save(path)
    else:
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f); w.writerow(cols); w.writerows(rows)

# ---------------------------------------------------------------- receipts / lookups
# ---------------------------------------------------------------- backup / restore / cloud
def make_backup(tag="backup"):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    tmp = os.path.join(DATA_DIR, "_tmp_backup.db")
    if os.path.exists(tmp): os.remove(tmp)
    dst = sqlite3.connect(tmp); db.backup(dst); dst.close()
    path = os.path.join(BACKUP_DIR, f"pos_{tag}_{dt.datetime.now():%Y%m%d_%H%M%S}.zip")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(tmp, "pos.db")
    os.remove(tmp)
    path = _encrypt_if_needed(path)
    keep = int(fS("keep_backups") or 30)
    files = sorted(glob.glob(os.path.join(BACKUP_DIR, "pos_*.zip*")), key=os.path.getmtime, reverse=True)
    hourly = [f for f in files if os.path.basename(f).startswith("pos_hourly_")]
    others = [f for f in files if f not in hourly]
    for old in hourly[24:] + others[keep:]:
        try: os.remove(old)
        except OSError: pass
    set_S("last_backup", now())
    return path

def _fernet(salt):
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=200000)
    return Fernet(base64.urlsafe_b64encode(kdf.derive(S("backup_password").encode())))

def _encrypt_if_needed(path):
    if not S("backup_password"): return path
    try: import cryptography  # noqa
    except Exception: raise RuntimeError("A backup password is set but the encryption library is missing in this build")
    salt = os.urandom(16)
    with open(path, "rb") as f: token = _fernet(salt).encrypt(f.read())
    with open(path + ".enc", "wb") as f: f.write(salt + token)
    os.remove(path)
    return path + ".enc"

def _decrypt_file(path):
    if not S("backup_password"): raise ValueError("This backup is password protected - enter the backup password in Settings first")
    with open(path, "rb") as f: raw = f.read()
    try: return _fernet(raw[:16]).decrypt(raw[16:])
    except Exception: raise ValueError("Wrong backup password")

def list_backups():
    return sorted(glob.glob(os.path.join(BACKUP_DIR, "pos_*.zip*")), key=os.path.getmtime, reverse=True)

def auto_backup(force=False):
    if S("auto_backup") != "1" and not force: return None
    key = f"{dt.date.today():%Y%m%d}"
    if not force and any(key in os.path.basename(f) for f in list_backups()): return None
    return make_backup("auto")

def restore_backup(zpath):
    make_backup("before_restore")
    tmp = os.path.join(DATA_DIR, "_restore.db")
    if zpath.endswith(".enc"):
        with zipfile.ZipFile(io.BytesIO(_decrypt_file(zpath))) as z:
            with open(tmp, "wb") as f: f.write(z.read("pos.db"))
    else:
        with zipfile.ZipFile(zpath) as z:
            with open(tmp, "wb") as f: f.write(z.read("pos.db"))
    src = sqlite3.connect(tmp); src.backup(db); src.close(); os.remove(tmp)
    db.executescript(SCHEMA); db.commit(); migrate()

def to_drive(path):
    """Copy to a Google Drive for Desktop folder if set, otherwise upload through the Google Drive API."""
    gf = S("gdrive_folder")
    if gf:
        if not os.path.isdir(gf): raise RuntimeError(f"Folder not found: {gf}")
        shutil.copy2(path, gf); set_S("last_offsite", now()); return f"Copied to {gf} (Google Drive will sync it)"
    if Credentials is None: raise RuntimeError("Google libraries are not installed in this build")
    scopes = ["https://www.googleapis.com/auth/drive.file"]
    tok, cred = os.path.join(DATA_DIR, "token.json"), os.path.join(DATA_DIR, "credentials.json")
    creds = Credentials.from_authorized_user_file(tok, scopes) if os.path.exists(tok) else None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token: creds.refresh(Request())
        else:
            if not os.path.exists(cred):
                raise RuntimeError("credentials.json not found.\nPut your Google OAuth file in:\n" + DATA_DIR +
                                   "\n(or set a Google Drive for Desktop folder in Settings)")
            creds = InstalledAppFlow.from_client_secrets_file(cred, scopes).run_local_server(port=0)
        with open(tok, "w") as f: f.write(creds.to_json())
    svc = build("drive", "v3", credentials=creds)
    name = f"{APP_NAME} Backups"
    res = svc.files().list(q=f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
                           fields="files(id)").execute().get("files", [])
    fid = res[0]["id"] if res else svc.files().create(
        body={"name": name, "mimeType": "application/vnd.google-apps.folder"}, fields="id").execute()["id"]
    svc.files().create(body={"name": os.path.basename(path), "parents": [fid]},
                       media_body=MediaFileUpload(path), fields="id").execute()
    set_S("last_offsite", now())
    return f"Uploaded to Google Drive folder '{name}'"

def send_mail(subject, body, path=None):
    to = S("email_to").strip()
    if not to: raise RuntimeError("Set 'Send backups to this email' in Settings first")
    if S("smtp_user"):
        msg = EmailMessage(); msg["Subject"] = subject; msg["From"] = S("smtp_user"); msg["To"] = to
        msg.set_content(body)
        if path:
            with open(path, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="zip" if path.endswith(".zip") else "octet-stream", filename=os.path.basename(path))
        with smtplib.SMTP(S("smtp_host"), int(S("smtp_port") or 587), timeout=60) as sm:
            sm.starttls(); sm.login(S("smtp_user"), S("smtp_pass")); sm.send_message(msg)
        return f"Email sent to {to}"
    try:
        import win32com.client
    except Exception:
        raise RuntimeError("Outlook automation needs pywin32 (included in the .exe build), or fill the SMTP settings.")
    ol = win32com.client.Dispatch("Outlook.Application"); m = ol.CreateItem(0)
    m.To = to; m.Subject = subject; m.Body = body
    if path: m.Attachments.Add(path)
    m.Send()
    return f"Sent through Outlook to {to}"

def email_backup(path):
    r = send_mail(f"{S('shop_name')} backup {today()}", "POS backup attached.", path)
    set_S("last_offsite", now()); return r

def hourly_backup():
    # automatic backup every N hours (Settings); also keeps a "pos_latest" copy in the Google Drive folder when one is set
    h = fS("backup_hours")
    if h <= 0 or S("auto_backup") != "1": return None
    fl = [f for f in list_backups() if os.path.basename(f).startswith("pos_hourly_")]
    if fl and time.time() - os.path.getmtime(fl[0]) < h * 3600: return None
    p = make_backup("hourly"); gf = S("gdrive_folder")
    if gf and os.path.isdir(gf):
        shutil.copy2(p, os.path.join(gf, "pos_latest" + (".zip.enc" if p.endswith(".enc") else ".zip"))); set_S("last_offsite", now())
    return p

def import_products(path):
    n = 0
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            name = (r.get("name") or "").strip()
            if not name: continue
            def fl(k):
                try: return float(r.get(k) or 0)
                except ValueError: return 0.0
            save_product({"name": name, "category": r.get("category") or "Accessory", "barcode": r.get("barcode") or "",
                          "cost": fl("cost"), "price": fl("price"), "qty": int(fl("qty")), "min_qty": int(fl("min_qty") or 2),
                          "track": "No"})
            n += 1
    return n


# ---------------------------------------------------------------- lookups used by the UI
def rows(sql, a=()):
    return [dict(r) for r in db.execute(sql, a).fetchall()]

def sale_detail(sid):
    r = q1("SELECT s.*, c.name cname, c.phone cphone, c.points cpoints FROM sales s LEFT JOIN customers c ON c.id=s.customer_id WHERE s.id=?", (sid,))
    if not r: raise ValueError("Invoice not found")
    s = dict(r)
    s["items"] = rows("SELECT i.*, COALESCE((SELECT SUM(qty) FROM returns_c x WHERE x.item_id=i.id),0) AS returned FROM sale_items i WHERE sale_id=?", (sid,))
    s["paid"] = round(val("SELECT SUM(amount) FROM payments WHERE sale_id=?", (sid,)), 2)
    s["balance"] = round(sale_balance(sid), 2)
    s["installments"] = rows("SELECT * FROM installments WHERE sale_id=? ORDER BY no", (sid,))
    s["payments"] = rows("SELECT date,amount,method,kind,cur,amount_lbp FROM payments WHERE sale_id=? ORDER BY id", (sid,))
    return s

def sales_list(d1, d2, text=""):
    return rows(f"SELECT s.id, s.date, s.customer_id, COALESCE(c.name,'Walk-in') customer, s.total, s.kind, "
                f"COALESCE((SELECT SUM(amount) FROM payments p WHERE p.sale_id=s.id),0) paid, {BAL} balance, "
                f"s.total - COALESCE((SELECT SUM(qty*cost) FROM sale_items i WHERE i.sale_id=s.id),0) - COALESCE((SELECT SUM(amount-qty*cost) FROM returns_c r WHERE r.sale_id=s.id),0) profit, "
                f"(SELECT group_concat(qty||' x '||name || COALESCE(' ('||imei||')',''), char(10)) FROM sale_items i WHERE i.sale_id=s.id) items "
                f"FROM sales s LEFT JOIN customers c ON c.id=s.customer_id "
                f"WHERE substr(s.date,1,10) BETWEEN ? AND ? AND (COALESCE(c.name,'Walk-in') LIKE ? OR CAST(s.id AS TEXT)=? OR "
                f"s.id IN (SELECT sale_id FROM sale_items WHERE name LIKE ? OR imei LIKE ?)) ORDER BY s.id DESC",
                (d1, d2, f"%{text}%", text.strip().lstrip('#'), f"%{text}%", f"%{text}%"))

def warranty_search(text):
    t = f"%{text.strip()}%"
    if not text.strip(): return []
    return rows("SELECT s.id sale_id, s.date, i.name, i.imei, i.warranty_until, c.name customer, c.phone FROM sale_items i "
                "JOIN sales s ON s.id=i.sale_id LEFT JOIN customers c ON c.id=s.customer_id "
                "WHERE i.imei LIKE ? OR CAST(s.id AS TEXT)=? OR c.name LIKE ? OR c.phone LIKE ? ORDER BY s.id DESC LIMIT 30",
                (t, text.strip().lstrip('#'), t, t))

def drawer():
    last = q1("SELECT * FROM closes ORDER BY id DESC LIMIT 1")
    since = last["date"] if last else "0000-00-00 00:00:00"
    usd = val("SELECT SUM(amount) FROM payments WHERE date>? AND method='Cash' AND COALESCE(cur,'USD')!='LBP'", (since,))
    lbp = val("SELECT SUM(amount_lbp) FROM payments WHERE date>? AND method='Cash' AND cur='LBP'", (since,))
    mv = lambda cur: val("SELECT SUM(CASE WHEN kind='in' THEN amount ELSE -amount END) FROM cash_moves WHERE date>? AND COALESCE(cur,'USD')=?", (since, cur))
    card = val("SELECT SUM(amount) FROM payments WHERE date>? AND method='Card'", (since,))
    transfer = val("SELECT SUM(amount) FROM payments WHERE date>? AND method NOT IN ('Cash','Card')", (since,))
    return {"since": last["date"] if last else None, "last_counted": last["counted"] if last else 0,
            "last_counted_lbp": (last["counted_lbp"] or 0) if last else 0,
            "cash": round(usd + mv("USD"), 2), "cash_lbp": round(lbp + mv("LBP")), "moves_usd": round(mv("USD"), 2), "moves_lbp": round(mv("LBP")),
            "card": round(card, 2), "transfer": round(transfer, 2),
            "count": val("SELECT COUNT(*) FROM sales WHERE date>?", (since,)),
            "revenue": round(val("SELECT SUM(total) FROM sales WHERE date>?", (since,)), 2),
            "history": rows("SELECT * FROM closes ORDER BY id DESC LIMIT 8")}

def drawer_close(float_, counted, float_lbp=0, counted_lbp=0):
    d = drawer(); expected = round(float_ + d["cash"], 2); exp_l = round(float_lbp + d["cash_lbp"])
    x("INSERT INTO closes(date,float,expected,counted,diff,cash,card,transfer,cnt,revenue,float_lbp,expected_lbp,counted_lbp,diff_lbp) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
      (now(), float_, expected, counted, round(counted - expected, 2), d["cash"], d["card"], d["transfer"], d["count"], d["revenue"],
       float_lbp, exp_l, counted_lbp, round(counted_lbp - exp_l)))

def repair_deliver(rid, method="Cash"):
    r = q1("SELECT * FROM repairs WHERE id=?", (rid,))
    if not r: raise ValueError("Ticket not found")
    if r["status"] == "Delivered": raise ValueError("Already delivered")
    line = {"pid": None, "name": "Repair: " + (r["device"] or "") + (" - " + r["issue"] if r["issue"] else ""), "qty": 1,
            "price": r["price"], "cost": r["cost"], "imei": r["imei"], "warranty": r["warranty_days"]}
    sid = create_sale([line], r["customer_id"], 0, 0, r["price"], method, "cash", None, f"Repair ticket #{rid}")
    x("UPDATE repairs SET status='Delivered', sale_id=? WHERE id=?", (sid, rid))
    return sid

def dashboard():
    run_recurring()
    d = today(); sm = period_summary(d, d); lim = (dt.date.today() + dt.timedelta(days=7)).isoformat()
    return {"sales_count": val("SELECT COUNT(*) FROM sales WHERE substr(date,1,10)=?", (d,)), "revenue": round(sm["net_rev"], 2),
            "profit": round(sm["gross"] - sm["exp"], 2), "cash": round(sm["cash"], 2),
            "alfa": alfa_state(), "touch": alfa_state("touch"), "last_backup": S("last_backup"), "last_offsite": S("last_offsite"), "low_count": val("SELECT COUNT(*) FROM products WHERE active=1 AND service=0 AND qty<=min_qty"),
            "debts": round(total_debts(), 2), "payable": round(total_payable(), 2),
            "overdue": val("SELECT COUNT(*) FROM installments WHERE paid<amount-0.004 AND due_date<?", (d,)),
            "low": rows("SELECT id,name,category,qty,min_qty FROM products WHERE active=1 AND service=0 AND qty<=min_qty ORDER BY qty LIMIT 12"),
            "due": rows("SELECT c.name customer, c.phone, i.sale_id, i.due_date, ROUND(i.amount-i.paid,2) remaining FROM installments i "
                        "JOIN sales s ON s.id=i.sale_id JOIN customers c ON c.id=s.customer_id WHERE i.paid<i.amount-0.004 AND i.due_date<=? "
                        "ORDER BY i.due_date LIMIT 12", (lim,)),
            "repairs_open": val("SELECT COUNT(*) FROM repairs WHERE status!='Delivered'")}


# ---------------------------------------------------------------- Alfa credit: validity days + Alfa services
# Stock of Alfa dollars lives in alfa_ledger (+ amount / + cost when it comes in, - when sold).
# Cost of 1$ = total cost / total amount (weighted average). Every sale is also a normal sale (profit, debts, receipts).
PROVIDERS = {"alfa": {"name": "Alfa", "fee": "alfa_fee", "svc_cat": "Services", "card_cat": "Recharge Cards", "inv": "%Alfa%"},
             "touch": {"name": "Touch", "fee": "touch_fee", "svc_cat": "Touch Services", "card_cat": "Touch Cards", "inv": "%Touch%"}}
def prov(p): return p if p in PROVIDERS else "alfa"
def pname(p): return PROVIDERS[prov(p)]["name"]

def _alfa_zero_out(provider="alfa"):
    # When a credit stock (Alfa / Touch) is empty any leftover cost is meaningless; clear it with a visible "reset" entry.
    provider = prov(provider)
    amt = val("SELECT SUM(amount) FROM alfa_ledger WHERE provider=?", (provider,)); cost = val("SELECT SUM(cost) FROM alfa_ledger WHERE provider=?", (provider,))
    if amt <= 0.00001 and (abs(cost) > 0.00001 or abs(amt) > 0.0000001):
        db.execute("INSERT INTO alfa_ledger(date,kind,amount,cost,note,ref_id,provider) VALUES(?,?,?,?,?,NULL,?)",
                   (now(), "reset", -amt, -cost, f"Leftover cost cleared (the {pname(provider)} stock was empty)", provider)); db.commit()

def alfa_state(provider="alfa"):
    provider = prov(provider); _alfa_zero_out(provider)
    amt = val("SELECT SUM(amount) FROM alfa_ledger WHERE provider=?", (provider,)); cost = val("SELECT SUM(cost) FROM alfa_ledger WHERE provider=?", (provider,))
    return {"amount": round(amt, 4), "cost": round(cost, 4), "avg": (cost / amt) if amt > 0.00001 else 0.0}

def lbp_rate():
    r = fS("lbp_rate")
    if r <= 0: raise ValueError("Set the LBP exchange rate in Settings first")
    return r

def _settle(price, cur, received, cid):
    cur = "LBP" if cur == "LBP" else "USD"
    rate = lbp_rate() if cur == "LBP" else 0
    if received in (None, ""):
        paid, lbp = price, (round(price * rate) if rate else 0)
    else:
        r = float(received)
        if r < 0: raise ValueError("Invalid amount received")
        paid, lbp = (r / rate, r) if cur == "LBP" else (r, 0)
    kind = "cash" if paid >= price - 0.004 else "debt"
    if kind == "debt" and not cid: raise ValueError("Choose a customer to record the debt")
    return cur, paid, lbp, kind

def alfa_adjust(amount, cost, note="", provider="alfa"):
    provider = prov(provider); _alfa_zero_out(provider)
    if abs(amount) < 0.00001 and abs(cost) < 0.00001: raise ValueError("Enter the amount and/or its cost")
    x("INSERT INTO alfa_ledger(date,kind,amount,cost,note,ref_id,provider) VALUES(?,?,?,?,?,NULL,?)", (now(), "adjust", amount, cost, note or "Manual adjustment", provider))

def validity_sale(months, card_ids, ret, price, cur="USD", received=None, cid=None, notes="", provider="alfa"):
    provider = prov(provider); _alfa_zero_out(provider)
    months = int(months)
    if not 1 <= months <= 12: raise ValueError("Choose 1 to 12 months")
    if ret <= 0: raise ValueError(f"Enter the {pname(provider)} dollars returned")
    if price < 0: raise ValueError("Invalid price")
    if not card_ids: raise ValueError("Choose at least one card")
    # a card is "t:<id>" (card type bought from an app/agent, no stock), "p:<id>" (inventory product) or "c:<cost>:<name>" (one-off)
    need, cards_cost, counts = {}, 0.0, {}
    for k in card_ids:
        k = str(k).strip()
        if k.isdigit(): k = "p:" + k
        kind_, _, rest = k.partition(":")
        if kind_ == "p":
            need[int(rest)] = need.get(int(rest), 0) + 1
        elif kind_ == "t":
            t = q1("SELECT * FROM alfa_cards WHERE id=? AND active=1 AND provider=?", (rest, provider))
            if not t: raise ValueError("That card type no longer exists")
            cards_cost += t["cost"]; counts[t["name"]] = counts.get(t["name"], 0) + 1
        elif kind_ == "c":
            cs, _, nm = rest.partition(":")
            try: cv = float(cs)
            except ValueError: raise ValueError("Invalid card cost")
            if cv < 0: raise ValueError("Invalid card cost")
            cards_cost += cv; key_ = f"{(nm or 'Card')} (${cv:g})"; counts[key_] = counts.get(key_, 0) + 1
        else: raise ValueError("Invalid card selected")
    for pid, n in need.items():
        p = q1("SELECT * FROM products WHERE id=? AND active=1", (pid,))
        if not p or p["service"] or p["track_imei"]: raise ValueError("Invalid card selected")
        if p["qty"] < n and S("allow_negative") != "1": raise ValueError(f"Not enough {p['name']} in Inventory (available {p['qty']})")
        cards_cost += p["cost"] * n; counts[p["name"]] = counts.get(p["name"], 0) + n
    names = [f"{nm} x{n}" if n > 1 else nm for nm, n in counts.items()]
    cur, paid, lbp, kind = _settle(price, cur, received, cid)
    cost_ret = cards_cost - price
    label = ("Touch " if provider == "touch" else "") + f"Validity days {months} MONTH"
    sid = create_sale([{"pid": None, "name": label, "qty": 1, "price": price, "cost": price, "imei": None}], cid, 0, 0, paid,
                      "Cash", kind, None, notes or " + ".join(names), 3, cur, lbp)
    with db:
        for pid, n in need.items(): db.execute("UPDATE products SET qty=qty-? WHERE id=?", (n, pid))
        db.execute("INSERT INTO alfa_ledger(date,kind,amount,cost,note,ref_id,provider) VALUES(?,?,?,?,?,?,?)", (now(), "validity", ret, cost_ret, label, sid, provider))
        db.execute("INSERT INTO validity_sales(date,months,price,ret,cards_cost,cost_ret,cards,names,sale_id,provider) VALUES(?,?,?,?,?,?,?,?,?,?)",
                   (now(), months, price, ret, cards_cost, cost_ret, json.dumps({str(k): v for k, v in need.items()}), " + ".join(names), sid, provider))
    return sid

def alfa_sell(stype, service_id, amount, price, cur="USD", received=None, cid=None, notes="", provider="alfa"):
    provider = prov(provider); pn = pname(provider)
    if price < 0: raise ValueError("Invalid price")
    if stype == "svc":
        s = q1("SELECT * FROM alfa_services WHERE id=? AND active=1 AND provider=?", (service_id, provider))
        if not s: raise ValueError("Choose a service")
        amt, fee, name = s["alfa"], 0.0, s["name"]
    else:
        if amount not in (0.5, 1, 1.5, 2, 2.5, 3): raise ValueError("Amount must be 0.5, 1, 1.5, 2, 2.5 or 3")
        amt, fee, name = amount, fS(PROVIDERS[provider]["fee"]), f"{pn} dollars {amount:g}$"
    ded = amt + fee
    st = alfa_state(provider)
    if ded > st["amount"] + 0.00001: raise ValueError(f"Not enough {pn} in stock ({st['amount']:.2f}$ left, needs {ded:.2f}$)")
    cost = ded * st["avg"]
    cur, paid, lbp, kind = _settle(price, cur, received, cid)
    sid = create_sale([{"pid": None, "name": name, "qty": 1, "price": price, "cost": cost, "imei": None}], cid, 0, 0, paid,
                      "Cash", kind, None, notes, 3, cur, lbp)
    with db:
        db.execute("INSERT INTO alfa_ledger(date,kind,amount,cost,note,ref_id,provider) VALUES(?,?,?,?,?,?,?)", (now(), "sale", -ded, -cost, name, sid, provider))
        db.execute("INSERT INTO alfa_sales(date,kind,name,amt,fee,ded,cost,price,profit,sale_id,provider) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                   (now(), stype, name, amt, fee, ded, cost, price, price - cost, sid, provider))
    _alfa_zero_out(provider)
    return sid

def card_sell(card_id, price, cur="USD", received=None, cid=None, notes="", provider="alfa"):
    # selling a credit card itself to a customer (cards are bought from an app/agent, so there is no stock)
    provider = prov(provider)
    t = q1("SELECT * FROM alfa_cards WHERE id=? AND active=1 AND provider=?", (card_id, provider))
    if not t: raise ValueError("Choose a card")
    if price < 0: raise ValueError("Invalid price")
    cur, paid, lbp, kind = _settle(price, cur, received, cid)
    name = t["name"] if "card" in t["name"].lower() else pname(provider) + " card " + t["name"]
    sid = create_sale([{"pid": None, "name": name, "qty": 1, "price": price, "cost": t["cost"], "imei": None}], cid, 0, 0, paid, "Cash", kind, None, notes, 3, cur, lbp)
    x("INSERT INTO alfa_sales(date,kind,name,amt,fee,ded,cost,price,profit,sale_id,provider) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
      (now(), "card", name, 0, 0, 0, t["cost"], price, price - t["cost"], sid, provider))
    return sid

def alfa_log(limit=40, provider="alfa"):
    provider = prov(provider)
    a = rows("SELECT date, CASE kind WHEN 'card' THEN 'Card sale' WHEN 'return' THEN 'Return' ELSE ? END kind, name item, -ded alfa, cost, price, profit, sale_id FROM alfa_sales WHERE provider=? ORDER BY id DESC LIMIT ?", (pname(provider) + " sale", provider, limit))
    v = rows("SELECT date, 'Validity days' kind, months||' MONTH - '||COALESCE(names,'') item, ret alfa, cost_ret cost, price, 0 profit, sale_id FROM validity_sales WHERE provider=? ORDER BY id DESC LIMIT ?", (provider, limit))
    return sorted(a + v, key=lambda r: r["date"], reverse=True)[:limit]

def wipe_all():
    # erases everything the user entered (products, sales, customers, Alfa...). Settings, PIN and logo are kept.
    with db:
        for (t,) in db.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' AND name NOT IN ('settings','users','audit','rate_history','images','topup_items')").fetchall():
            db.execute(f"DELETE FROM {t}")
        db.execute("DELETE FROM images WHERE kind NOT IN ('tab','sq')")
    db.execute("VACUUM")

def alfa_set(amount, avg, note="", provider="alfa"):
    # set the credit stock to an exact amount and cost per 1$ (creates the adjustment needed)
    provider = prov(provider)
    if amount < 0 or avg < 0: raise ValueError("Values cannot be negative")
    st = alfa_state(provider)
    da, dc = round(amount - st["amount"], 6), round(amount * avg - st["cost"], 6)
    if abs(da) < 0.000001 and abs(dc) < 0.000001: raise ValueError("Nothing to change - the stock already has these values")
    x("INSERT INTO alfa_ledger(date,kind,amount,cost,note,ref_id,provider) VALUES(?,?,?,?,?,NULL,?)", (now(), "adjust", da, dc, note or "Stock set manually", provider))

def alfa_adjust_delete(rid):
    r = q1("SELECT * FROM alfa_ledger WHERE id=? AND kind='adjust'", (rid,))
    if not r: raise ValueError("Adjustment not found")
    if alfa_state(r["provider"])["amount"] - r["amount"] < -0.00001: raise ValueError("Cannot remove it: that credit was already used by later sales")
    x("DELETE FROM alfa_ledger WHERE id=?", (rid,))

# ---------------------------------------------------------------- users, audit, rate history
def users_count(): return val("SELECT COUNT(*) FROM users WHERE active=1")

def norm_username(u): return re.sub(r"\s+", "", (u or "").strip().lower())
USERNAME_RE = re.compile(r"^[a-z0-9._-]{3,32}$")

def migrate_pin():
    old = db.execute("SELECT v FROM settings WHERE k='pin'").fetchone()
    if old and old[0] and users_count() == 0:
        db.execute("INSERT INTO users(name,role,pin,username) VALUES('Owner','owner',?,'owner')", (old[0],))
        db.execute("UPDATE settings SET v='' WHERE k='pin'"); db.commit()

def migrate_users():
    # every user needs a unique username: older users get their name (lower case, no spaces)
    used = set()
    for r in q("SELECT id, name, username FROM users ORDER BY id"):
        u = norm_username(r["username"]) or norm_username(r["name"])
        if not USERNAME_RE.match(u): u = f"user{r['id']}"
        base, n = u, 2
        while u in used: u = f"{base}{n}"; n += 1
        used.add(u)
        if u != r["username"]: db.execute("UPDATE users SET username=? WHERE id=?", (u, r["id"]))
    db.execute("CREATE UNIQUE INDEX IF NOT EXISTS ix_users_username ON users(username COLLATE NOCASE)"); db.commit()

def ensure_rate_row():
    if not val("SELECT COUNT(*) FROM rate_history") and fS("lbp_rate") > 0:
        db.execute("INSERT INTO rate_history(date,rate,user) VALUES(?,?,?)", (now(), fS("lbp_rate"), "system")); db.commit()

migrate_pin(); migrate_users(); ensure_rate_row()

def check_owner_pin(p):
    if not users_count(): return True
    return any(verify_pin(p or "", r["pin"]) for r in q("SELECT pin FROM users WHERE active=1 AND role='owner'"))
check_pin = check_owner_pin

def audit(action, detail="", user=None):
    try:
        db.execute("INSERT INTO audit(date,user,action,detail) VALUES(?,?,?,?)", (now(), CURRENT["user"] if user is None else user, action, (detail or "")[:400])); db.commit()
    except Exception: pass

def user_save(uid, name, role, password, username=None):
    name = (name or "").strip()
    if not name: raise ValueError("Enter the full name")
    uname = norm_username(username) or norm_username(name)
    if not USERNAME_RE.match(uname): raise ValueError("The username needs 3 to 32 letters, numbers, dots or dashes, with no spaces")
    if role not in ("owner", "cashier"): role = "cashier"
    if users_count() == 0: role = "owner"
    password = password or ""
    if password and len(password) < 6: raise ValueError("The password needs at least 6 characters")
    try:
        if uid:
            u = q1("SELECT * FROM users WHERE id=?", (uid,))
            if not u: raise ValueError("User not found")
            if u["role"] == "owner" and role != "owner" and val("SELECT COUNT(*) FROM users WHERE active=1 AND role='owner' AND id!=?", (uid,)) == 0:
                raise ValueError("There must be at least one owner")
            if password: x("UPDATE users SET name=?, username=?, role=?, pin=? WHERE id=?", (name, uname, role, hash_pin(password), uid))
            else: x("UPDATE users SET name=?, username=?, role=? WHERE id=?", (name, uname, role, uid))
            return uid
        if not password: raise ValueError("Enter a password (at least 6 characters)")
        return x("INSERT INTO users(name,role,pin,username) VALUES(?,?,?,?)", (name, role, hash_pin(password), uname))
    except sqlite3.IntegrityError:
        db.rollback(); raise ValueError("This username or name is already used by another user")

def user_delete(uid):
    u = q1("SELECT * FROM users WHERE id=?", (uid,))
    if not u: return
    if u["role"] == "owner" and val("SELECT COUNT(*) FROM users WHERE active=1 AND role='owner' AND id!=?", (uid,)) == 0:
        raise ValueError("You cannot delete the only owner")
    x("DELETE FROM users WHERE id=?", (uid,))

def login(username, password):
    u = q1("SELECT * FROM users WHERE username=? COLLATE NOCASE AND active=1", (norm_username(username),))
    if not u or not verify_pin(password or "", u["pin"]): raise ValueError("Wrong username or password")
    return {"id": u["id"], "name": u["name"], "role": u["role"], "username": u["username"]}

def recovery_new():
    code = "-".join("".join(secrets.choice("ABCDEFGHJKMNPQRSTUVWXYZ23456789") for _ in range(4)) for _ in range(3))
    set_S("owner_recovery", hash_pin(code)); return code

def recover_owner(username, code, new_password):
    stored = S("owner_recovery")
    if not stored or not verify_pin((code or "").strip().upper(), stored): raise ValueError("Wrong recovery code")
    u = q1("SELECT * FROM users WHERE username=? COLLATE NOCASE AND role='owner' AND active=1", (norm_username(username),))
    if not u: raise ValueError("There is no owner with this username")
    if len(new_password or "") < 6: raise ValueError("The new password needs at least 6 characters")
    x("UPDATE users SET pin=? WHERE id=?", (hash_pin(new_password), u["id"]))
    return recovery_new()

# ---------------------------------------------------------------- cash moves, recurring expenses, presets
def run_recurring():
    d = dt.date.today(); ym = f"{d:%Y-%m}"
    for r in q("SELECT * FROM recurring WHERE active=1"):
        due = min(max(1, int(r["day"] or 1)), calendar.monthrange(d.year, d.month)[1])
        if r["last_ym"] == ym or d.day < due: continue
        db.execute("INSERT INTO expenses(date,category,amount,note) VALUES(?,?,?,?)", (f"{ym}-{due:02d}", r["category"], r["amount"], ("Recurring. " + (r["note"] or "")).strip()))
        db.execute("UPDATE recurring SET last_ym=? WHERE id=?", (ym, r["id"]))
    db.commit()

def stock_count_apply(items):
    n = 0
    for it in items:
        p = q1("SELECT * FROM products WHERE id=?", (it["pid"],))
        if not p or p["service"] or p["track_imei"]: continue
        counted = int(float(it["counted"]))
        if counted < 0: raise ValueError("Counted quantity cannot be negative")
        if counted != p["qty"]:
            db.execute("UPDATE products SET qty=? WHERE id=?", (counted, p["id"])); n += 1
            audit("stock_count", f"{p['name']}: {p['qty']} -> {counted}")
    db.commit(); return n

def import_v4():
    src = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "MobilePOS", "pos.db")
    if not os.path.exists(src): raise ValueError("No POS V4 data was found on this PC (looked in " + src + ")")
    make_backup("before_import_v4")
    s = sqlite3.connect(src); s.backup(db); s.close()
    db.executescript(SCHEMA); db.commit(); migrate(); migrate_pin(); migrate_users(); ensure_rate_row()

# ---------------------------------------------------------------- day summary / weekly chart
def day_summary(d):
    sm = period_summary(d, d); on = "substr(date,1,10)=?"
    cash_usd = val(f"SELECT SUM(amount) FROM payments WHERE {on} AND method='Cash' AND COALESCE(cur,'USD')!='LBP'", (d,))
    cash_lbp = val(f"SELECT SUM(amount_lbp) FROM payments WHERE {on} AND method='Cash' AND cur='LBP'", (d,))
    mv = lambda k, cur: val(f"SELECT SUM(amount) FROM cash_moves WHERE {on} AND kind=? AND COALESCE(cur,'USD')=?", (d, k, cur))
    return {"date": d, "sales_count": val(f"SELECT COUNT(*) FROM sales WHERE {on}", (d,)), "revenue": round(sm["net_rev"], 2),
            "gross": round(sm["gross"], 2), "expenses": round(sm["exp"], 2), "profit": round(sm["gross"] - sm["exp"], 2),
            "cash_usd": round(cash_usd, 2), "cash_lbp": round(cash_lbp),
            "card": round(val(f"SELECT SUM(amount) FROM payments WHERE {on} AND method='Card'", (d,)), 2),
            "transfer": round(val(f"SELECT SUM(amount) FROM payments WHERE {on} AND method NOT IN ('Cash','Card')", (d,)), 2),
            "in_usd": round(mv("in", "USD"), 2), "out_usd": round(mv("out", "USD"), 2), "in_lbp": round(mv("in", "LBP")), "out_lbp": round(mv("out", "LBP")),
            "new_debt": round(val(f"SELECT SUM(b) FROM (SELECT {BAL} b FROM sales s WHERE substr(s.date,1,10)=?) WHERE b>0.004", (d,)), 2),
            "collected": round(val(f"SELECT SUM(amount) FROM payments WHERE {on} AND kind='payment'", (d,)), 2),
            "alfa_sales": val(f"SELECT COUNT(*) FROM alfa_sales WHERE {on} AND provider='alfa' AND kind!='return'", (d,)), "alfa_profit": round(val(f"SELECT SUM(profit) FROM alfa_sales WHERE {on} AND provider='alfa'", (d,)), 2),
            "validity": val(f"SELECT COUNT(*) FROM validity_sales WHERE {on} AND provider='alfa'", (d,)), "alfa": alfa_state(),
            "touch_sales": val(f"SELECT COUNT(*) FROM alfa_sales WHERE {on} AND provider='touch' AND kind!='return'", (d,)), "touch_profit": round(val(f"SELECT SUM(profit) FROM alfa_sales WHERE {on} AND provider='touch'", (d,)), 2),
            "touch_validity": val(f"SELECT COUNT(*) FROM validity_sales WHERE {on} AND provider='touch'", (d,)), "touch": alfa_state("touch"),
            "low_count": val("SELECT COUNT(*) FROM products WHERE active=1 AND service=0 AND qty<=min_qty"),
            "debts": round(total_debts(), 2)}

def day_text(d):
    s = day_summary(d); c = S("currency") or "$"; a = s["alfa"]
    return "\n".join([f"{S('shop_name')} - summary for {d}", "",
        f"Sales: {s['sales_count']}   Revenue: {c}{s['revenue']:.2f}   Profit: {c}{s['profit']:.2f}   Expenses: {c}{s['expenses']:.2f}",
        f"Cash USD: {c}{s['cash_usd']:.2f}   Cash L.L: {s['cash_lbp']:,.0f}   Card: {c}{s['card']:.2f}   Transfer: {c}{s['transfer']:.2f}",
        f"Cash in/out USD: +{s['in_usd']:.2f} / -{s['out_usd']:.2f}   L.L: +{s['in_lbp']:,.0f} / -{s['out_lbp']:,.0f}",
        f"Validity days sold: {s['validity']}   Alfa sales: {s['alfa_sales']}   Alfa profit: {c}{s['alfa_profit']:.2f}",
        f"Alfa stock: {a['amount']:.2f}$ (cost of 1$ = {a['avg']:.3f})",
        f"Touch: validity days {s['touch_validity']}   sales {s['touch_sales']}   profit {c}{s['touch_profit']:.2f}   stock {s['touch']['amount']:.2f}$" if (s['touch_sales'] or s['touch_validity'] or s['touch']['amount']) else "",
        f"New debts today: {c}{s['new_debt']:.2f}   Debts collected: {c}{s['collected']:.2f}   All customer debts: {c}{s['debts']:.2f}",
        f"Low stock items: {s['low_count']}"])

def send_summary(d=None): return send_mail(f"{S('shop_name')} summary {d or today()}", day_text(d or today()))

def week_chart():
    out = []
    for i in range(6, -1, -1):
        d = (dt.date.today() - dt.timedelta(days=i)).isoformat(); sm = period_summary(d, d)
        out.append({"date": d, "rev": round(sm["net_rev"], 2), "profit": round(sm["gross"], 2)})
    return out


def _upsert_linked(p, col, link_id, name, price, cost, category):
    if p:
        db.execute("UPDATE products SET name=?, price=?, cost=?, category=?, service=1, active=1, barcode=COALESCE(barcode, ?) WHERE id=?", (name, price, cost, category, f"2{p['id']:09d}", p["id"]))
    else:
        pid = db.execute(f"INSERT INTO products(name,category,cost,price,qty,min_qty,track_imei,service,warranty,{col}) VALUES(?,?,?,?,?,?,?,?,?,?)",
                         (name, category, cost, price, 0, 0, 0, 1, 0, link_id)).lastrowid
        db.execute("UPDATE products SET barcode=? WHERE id=?", (f"2{pid:09d}", pid))

def sync_alfa_products():
    # Starred services -> products in the provider's services category ("Services" / "Touch Services");
    # card types that have a selling price -> "Recharge Cards" / "Touch Cards". Managed in their own tabs, no stock.
    for sv in q("SELECT * FROM alfa_services"):
        P = PROVIDERS[prov(sv["provider"])]
        p = q1("SELECT * FROM products WHERE alfa_service_id=?", (sv["id"],))
        if sv["active"] and sv["fav"]: _upsert_linked(p, "alfa_service_id", sv["id"], sv["name"], sv["price"], 0.0, P["svc_cat"])
        elif p and p["active"]: db.execute("UPDATE products SET active=0, barcode=NULL WHERE id=?", (p["id"],))
    for cd in q("SELECT * FROM alfa_cards"):
        P = PROVIDERS[prov(cd["provider"])]
        p = q1("SELECT * FROM products WHERE alfa_card_id=?", (cd["id"],))
        name = cd["name"] if "card" in cd["name"].lower() else P["name"] + " card " + cd["name"]
        if cd["active"] and (cd["price"] or 0) > 0: _upsert_linked(p, "alfa_card_id", cd["id"], name, cd["price"], cd["cost"], P["card_cat"])
        elif p and p["active"]: db.execute("UPDATE products SET active=0, barcode=NULL WHERE id=?", (p["id"],))
    db.commit()

sync_alfa_products()


# ---------------------------------------------------------------- trade-in (used phones) and warranty claims
def tradein_save(a):
    model = (a.get("model") or "").strip(); imei = re.sub(r"\s+", "", str(a.get("imei") or ""))
    if not model: raise ValueError("Enter the phone model")
    if not imei: raise ValueError("Enter the IMEI")
    cost = float(a.get("cost") or 0); price = float(a.get("price") or 0)
    if cost <= 0: raise ValueError("Enter what you paid the seller")
    if price <= 0: raise ValueError("Enter the selling price")
    if q1("SELECT 1 FROM imeis WHERE imei=?", (imei,)): raise ValueError("This IMEI is already in the system")
    name = model if "used" in model.lower() else model + " (used)"
    p = q1("SELECT * FROM products WHERE active=1 AND name=? AND category='Used phones' AND price=? AND track_imei=1", (name, price))
    pid = p["id"] if p else save_product({"name": name, "category": "Used phones", "cost": cost, "price": price, "track_imei": True, "min_qty": 0,
                                          "warranty": int(float(a.get("warranty_days") or 0))})
    add_imeis(pid, [imei], cost)
    tid = x("INSERT INTO tradeins(date,product_id,imei,model,cost,price,seller_name,seller_phone,id_note,condition,notes,user) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
            (now(), pid, imei, model, cost, price, (a.get("seller_name") or "").strip(), (a.get("seller_phone") or "").strip(), (a.get("id_note") or "").strip(),
             (a.get("condition") or "").strip(), (a.get("notes") or "").strip(), CURRENT["user"]))
    if str(a.get("pay", "1")) not in ("0", "No", "False", "false"):
        cur = "LBP" if a.get("paid_cur") == "LBP" else "USD"
        amt = round(cost * lbp_rate()) if cur == "LBP" else cost
        x("INSERT INTO cash_moves(date,kind,amount,cur,reason,user) VALUES(?,?,?,?,?,?)", (now(), "out", amt, cur, f"Trade-in: {model} IMEI {imei}", CURRENT["user"]))
    return tid

def tradeins_list(limit=60):
    return rows("SELECT t.*, i.status, i.sale_id, (SELECT SUM(si.price) FROM sale_items si WHERE si.sale_id=i.sale_id AND si.imei=t.imei) sold_price "
                "FROM tradeins t LEFT JOIN imeis i ON i.imei=t.imei ORDER BY t.id DESC LIMIT ?", (limit,))

CLAIM_STATUSES = ["Open", "Sent to supplier", "Repaired", "Replaced", "Rejected", "Closed"]

def claim_save(a):
    imei = re.sub(r"\s+", "", str(a.get("imei") or "")); issue = (a.get("issue") or "").strip()
    if not issue: raise ValueError("Describe the problem")
    it = q1("SELECT i.*, s.date sdate, s.customer_id FROM sale_items i JOIN sales s ON s.id=i.sale_id WHERE i.imei=? ORDER BY i.id DESC LIMIT 1", (imei,)) if imei else None
    if not it and a.get("sale_id"): it = q1("SELECT i.*, s.date sdate, s.customer_id FROM sale_items i JOIN sales s ON s.id=i.sale_id WHERE i.sale_id=? ORDER BY i.id LIMIT 1", (a["sale_id"],))
    cust = q1("SELECT * FROM customers WHERE id=?", (it["customer_id"],)) if it and it["customer_id"] else None
    return x("INSERT INTO claims(date,sale_id,imei,customer,phone,product,warranty_until,issue,status,supplier,user) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
             (now(), it["sale_id"] if it else None, imei or (it["imei"] if it else ""), (a.get("customer") or (cust["name"] if cust else "")).strip(),
              (a.get("phone") or (cust["phone"] if cust else "") or "").strip(), (a.get("product") or (it["name"] if it else "")).strip(),
              it["warranty_until"] if it else None, issue, "Open", (a.get("supplier") or "").strip(), CURRENT["user"]))

def claim_update(a):
    st = a.get("status") if a.get("status") in CLAIM_STATUSES else "Open"
    closed = now() if st in ("Repaired", "Replaced", "Rejected", "Closed") else None
    x("UPDATE claims SET status=?, resolution=?, notes=?, supplier=?, closed=? WHERE id=?",
      (st, (a.get("resolution") or "").strip(), (a.get("notes") or "").strip(), (a.get("supplier") or "").strip(), closed, a["id"]))

def claims_list(status=None):
    if status and status != "All": return rows("SELECT * FROM claims WHERE status=? ORDER BY id DESC LIMIT 200", (status,))
    return rows("SELECT * FROM claims ORDER BY (status IN ('Open','Sent to supplier')) DESC, id DESC LIMIT 200")

# ---------------------------------------------------------------- bulk price change
def price_bulk(category, pct, round_to=0):
    pct = float(pct); round_to = float(round_to or 0)
    if pct == 0 or abs(pct) > 200: raise ValueError("Enter a percentage between -200 and 200 (not 0)")
    batch = f"{now()} | {category or 'All categories'} | {pct:+g}%"; n = 0
    db.execute("DELETE FROM price_undo")
    for r in q("SELECT id, price FROM products WHERE active=1 AND alfa_service_id IS NULL AND alfa_card_id IS NULL AND (?='' OR category=?)", (category or "", category or "")):
        new = r["price"] * (1 + pct / 100.0)
        if round_to > 0: new = round(new / round_to) * round_to
        new = max(0.0, round(new, 2))
        if abs(new - r["price"]) > 0.0001:
            db.execute("INSERT INTO price_undo(product_id,old_price,batch) VALUES(?,?,?)", (r["id"], r["price"], batch))
            db.execute("UPDATE products SET price=? WHERE id=?", (new, r["id"])); n += 1
    db.commit(); return n

def price_undo_info():
    r = q1("SELECT COUNT(*) n, MAX(batch) batch FROM price_undo")
    return {"count": r["n"], "batch": r["batch"]}

def price_undo():
    n = 0
    for r in q("SELECT * FROM price_undo"):
        db.execute("UPDATE products SET price=? WHERE id=?", (r["old_price"], r["product_id"])); n += 1
    db.execute("DELETE FROM price_undo"); db.commit(); return n


# ---------------------------------------------------------------- pictures (products, services, card types, screens) - managed in Settings
IMG_PRODUCT = ("COALESCE((SELECT id||'.'||ver FROM images WHERE kind='product' AND ref=CAST(p.id AS TEXT)),"
               "(SELECT id||'.'||ver FROM images WHERE kind='service' AND ref=CAST(p.alfa_service_id AS TEXT)),"
               "(SELECT id||'.'||ver FROM images WHERE kind='card' AND ref=CAST(p.alfa_card_id AS TEXT)))")

def image_set(kind, ref, data, mime):
    ref = str(ref); r = q1("SELECT id, ver FROM images WHERE kind=? AND ref=?", (kind, ref))
    if r:
        x("UPDATE images SET data=?, mime=?, ver=ver+1 WHERE id=?", (data, mime, r["id"])); return f"{r['id']}.{r['ver'] + 1}"
    iid = x("INSERT INTO images(kind,ref,mime,data,ver) VALUES(?,?,?,?,1)", (kind, ref, mime, data)); return f"{iid}.1"

def image_delete(kind, ref): x("DELETE FROM images WHERE kind=? AND ref=?", (kind, str(ref)))

def tab_pics(): return {r["ref"]: f"{r['id']}.{r['ver']}" for r in q("SELECT id, ref, ver FROM images WHERE kind='tab'")}
def sq_pics(): return {r["ref"]: f"{r['id']}.{r['ver']}" for r in q("SELECT id, ref, ver FROM images WHERE kind='sq'")}


# ---------------------------------------------------------------- top-ups typed in the cart (game / app cards): name, price and profit; cost = price - profit
def topup_record(sid, name, price, cost, player, owner):
    x("INSERT INTO topups(sale_id,name,price,cost,profit,player,date) VALUES(?,?,?,?,?,?,?)", (sid, name, price, cost, round(price - cost, 4), player or "", now()))
    r = q1("SELECT id FROM topup_items WHERE name=?", (name,))
    if r:
        if owner: x("UPDATE topup_items SET price=?, profit=?, uses=uses+1, last=? WHERE id=?", (price, round(price - cost, 4), now(), r["id"]))
        else: x("UPDATE topup_items SET uses=uses+1, last=? WHERE id=?", (now(), r["id"]))
    elif owner: x("INSERT INTO topup_items(name,price,profit,uses,last) VALUES(?,?,?,1,?)", (name, price, round(price - cost, 4), now()))


# ---------------------------------------------------------------- program updates from a file (no rebuild of the .exe)
# Versions live in DATA_DIR/app/<version>/ (pos_core.py, pos_server.py, index.html); app/current.json says which one the launcher starts.
APP_DIR = os.path.join(DATA_DIR, "app")
BUILTIN_VERSION = os.environ.get("POS_BUILTIN_VERSION", "")
UPDATE_FILES = ("pos_core.py", "pos_server.py", "index.html")

def _ptr():
    try:
        with open(os.path.join(APP_DIR, "current.json"), encoding="utf-8") as f: return {"version": "builtin", "prev": None, "status": "ok", "tries": 0, **json.load(f)}
    except Exception: return {"version": "builtin", "prev": None, "status": "ok", "tries": 0}

def _save_ptr(p):
    os.makedirs(APP_DIR, exist_ok=True); tmp = os.path.join(APP_DIR, "current.json.tmp")
    with open(tmp, "w", encoding="utf-8") as f: json.dump(p, f)
    os.replace(tmp, os.path.join(APP_DIR, "current.json"))

def update_mark_ok():
    p = _ptr()
    if p.get("status") == "new": p["status"] = "ok"; p["tries"] = 0; _save_ptr(p)

def _vt(v): return tuple(int(n) for n in re.findall(r"\d+", v or ""))

def update_installed():
    out = []
    if os.path.isdir(APP_DIR):
        for n in os.listdir(APP_DIR):
            p = os.path.join(APP_DIR, n)
            if os.path.isdir(p) and all(os.path.isfile(os.path.join(p, f)) for f in UPDATE_FILES): out.append({"version": n, "mtime": os.path.getmtime(p)})
    return sorted(out, key=lambda r: _vt(r["version"]), reverse=True)

def update_info(running=""):
    p = _ptr()
    return {"builtin": BUILTIN_VERSION or "built-in", "current": p.get("version", "builtin"), "running": running, "status": p.get("status"), "note": p.get("note", ""), "installed": update_installed()}

def update_read(raw):
    if len(raw) > 30_000_000: raise ValueError("This file is too big to be an update")
    try: z = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile: raise ValueError("This is not a zip file")
    found = {}
    for n in z.namelist():
        base = os.path.basename(n)
        if base in UPDATE_FILES + ("POS_V5_NOTES.md",) and not n.endswith("/") and (base not in found or len(n) < len(found[base][0])): found[base] = (n, None)
    miss = [f for f in UPDATE_FILES if f not in found]
    if miss: raise ValueError("The zip is missing: " + ", ".join(miss) + ". It must contain the program files (index.html, pos_core.py, pos_server.py).")
    files = {b: z.read(found[b][0]) for b in found}
    for b in ("pos_core.py", "pos_server.py"):
        try: compile(files[b].decode("utf-8"), b, "exec")
        except Exception as e: raise ValueError(f"{b} has an error and cannot be installed ({type(e).__name__})")
    if b"__TOKEN__" not in files["index.html"]: raise ValueError("index.html is not a valid program page")
    m = re.search(r'^VERSION\s*=\s*"([^"]+)"', files["pos_server.py"].decode("utf-8"), re.M)
    ver = m.group(1) if m else ""
    if not re.fullmatch(r"V?\d+(\.\d+){0,3}", ver): raise ValueError("The zip has no valid version number")
    notes = ""
    if "POS_V5_NOTES.md" in files:
        nm = re.search(r"^## " + re.escape(ver) + r"\s*\n(.*?)(?=^## |\Z)", files["POS_V5_NOTES.md"].decode("utf-8", "replace"), re.S | re.M)
        notes = (nm.group(1).strip() if nm else "")[:900]
    return {"version": ver, "files": files, "notes": notes}

def update_check(raw, running):
    u = update_read(raw); rel = "same" if u["version"] == running else ("newer" if _vt(u["version"]) > _vt(running) else "older")
    return {"version": u["version"], "relation": rel, "notes": u["notes"], "running": running}

def update_install(raw):
    u = update_read(raw); ver = u["version"]; p = _ptr()
    make_backup("before_update")
    folder = os.path.join(APP_DIR, ver); os.makedirs(folder, exist_ok=True)
    for f in UPDATE_FILES:
        tmp = os.path.join(folder, f + ".tmp")
        with open(tmp, "wb") as fh: fh.write(u["files"][f])
        os.replace(tmp, os.path.join(folder, f))
    prev = p.get("version", "builtin")
    _save_ptr({"version": ver, "prev": prev if prev != ver else p.get("prev"), "status": "new", "tries": 0})
    keep = {ver, prev, "builtin"}
    for r in update_installed()[3:]:
        if r["version"] not in keep: shutil.rmtree(os.path.join(APP_DIR, r["version"]), ignore_errors=True)
    return ver

def update_use(version):
    p = _ptr()
    if version != "builtin" and version not in [r["version"] for r in update_installed()]: raise ValueError("That version is not installed")
    _save_ptr({"version": version, "prev": p.get("version") if p.get("version") != version else p.get("prev"), "status": "ok", "tries": 0})

def scrub_audit():
    # older builds wrote the typed password into the activity log when a user was created or edited: hide it
    for r in q("SELECT id, detail FROM audit WHERE detail LIKE '%\"password\"%' OR detail LIKE '%\"b64\"%' OR detail LIKE '%\"code\"%'"):
        d = re.sub(r'("(?:password|b64|code)"\s*:\s*)"[^"]*("|$)', r'\1"***"', r["detail"])
        if d != r["detail"]: db.execute("UPDATE audit SET detail=? WHERE id=?", (d, r["id"]))
    db.commit()

scrub_audit()
