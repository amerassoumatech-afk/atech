# -*- coding: utf-8 -*-
"""Shop POS - local API + desktop window launcher.  Run:  python pos_server.py"""
import sqlite3, json, os, sys, re, threading, secrets, time, traceback, webbrowser, http.server, tempfile, base64, io, csv, socket, ipaddress
import pos_core as core
from pos_core import q, q1, x, val, rows, today, S

VERSION = "V5.20"
API, LOCK, TOKEN = {}, threading.RLock(), secrets.token_hex(16)
SESSION = {"id": None, "user": "", "role": "cashier", "sid": ""}   # the session of the request being handled
SESSIONS = {}
LAN = {"on": False, "port": 0}
TAILSCALE = ipaddress.ip_network("100.64.0.0/10")
FAILS = {}
PUBLIC = {"ping", "boot", "users_public", "login", "logout", "recover"}
OWNER_ONLY = {"update_info", "update_check", "update_install", "update_use", "app_restart", "image_set", "image_delete", "dash_rows", "price_bulk", "price_undo", "price_undo_info", "lan_info", "backup_upload", "report", "report_list", "report_export", "settings_save", "users", "user_save", "user_delete", "audit_list", "rate_history", "import_v4",
              "wipe_all", "backup_now", "backup_list", "backup_restore", "backup_drive", "backup_email", "import_csv", "alfa_adjustments", "open_folder",
              "recovery_new", "recurring_list", "recurring_save", "recurring_delete", "summary_send", "purchase_create", "supplier_save", "supplier_pay", "supplier_get", "suppliers",
              "product_save", "imeis_add", "offer_save", "expense_save", "expenses", "offers", "stock_add", "alfa_card_save", "alfa_card_delete"}
SENSITIVE = {"topup_item_delete", "tradein_save", "claim_delete", "sale_void", "sale_return", "product_delete", "imei_remove", "expense_delete", "customer_delete", "offer_delete", "alfa_adjust", "alfa_set",
             "alfa_adjust_delete", "cash_move_delete", "stock_count_apply", "supplier_return", "repair_delete", "alfa_service_save", "alfa_service_delete", "preset_delete"}
AUDITED = SENSITIVE | {"sale_create", "sale_pay", "customer_pay", "settings_save", "wipe_all", "backup_restore", "import_v4", "validity_sale", "alfa_sell", "purchase_create",
                       "supplier_pay", "cash_move_save", "drawer_close", "user_save", "user_delete", "expense_save", "product_save", "imeis_add", "repair_deliver",
                       "recurring_save", "recurring_delete", "offer_save", "customer_save", "backup_restore", "stock_add", "update_install", "update_use", "app_restart", "price_bulk", "price_undo", "claim_save", "claim_update"}
HIDE_KEYS = {"pin", "_opin", "password", "code", "b64", "logo", "confirm", "smtp_pass", "backup_password", "text", "cart"}
PING = {"t": time.time(), "seen": False}
EXPORT_DIR = os.path.join(os.path.expanduser("~"), "Documents", core.APP_NAME + "_Exports")
OFFER = "(SELECT price FROM offers o WHERE o.product_id=p.id AND o.start<=? AND o.end>=? ORDER BY price LIMIT 1)"

def api(f): API[f.__name__] = f; return f
def num(v, d=0.0):
    try: return float(str(v).replace(",", "")) if v not in (None, "") else d
    except ValueError: return d
def resource(name):
    here = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)      # an installed update brings its own index.html next to this file
    if os.path.isfile(here): return here
    return os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))), name)
def open_path(p):
    if os.environ.get("MOBILEPOS_NOOPEN"): return
    if os.name == "nt": os.startfile(p)
    else: webbrowser.open("file://" + p)

def _pay_amt(a):
    amt = num(a["amount"])
    if a.get("cur") == "LBP": return round(amt / core.lbp_rate(), 2), "LBP", amt
    return amt, "USD", 0

# ------------------------------------------------------------ app / settings
def role(): return SESSION["role"] if core.users_count() else "owner"
def me_info(): return {"user": SESSION["user"] if core.users_count() else "", "role": role(), "has_users": core.users_count() > 0}

@api
def ping(a): PING["t"] = time.time(); PING["seen"] = True; return True

@api
def boot(a):
    core.run_recurring(); core.update_mark_ok()
    st = {k: S(k) for k in core.DEFAULTS if k not in ("pin", "smtp_pass", "backup_password", "owner_recovery")}
    st["has_smtp_pass"] = bool(S("smtp_pass")); st["has_recovery"] = bool(S("owner_recovery")); st["has_backup_password"] = bool(S("backup_password"))
    return {"settings": st, "data_dir": core.DATA_DIR, "backup_dir": core.BACKUP_DIR, "today": today(),
            "version": VERSION, "drive_api": core.Credentials is not None, "excel": core.openpyxl is not None, "me": me_info(), "pics": core.tab_pics(), "sqpics": core.sq_pics()}

@api
def settings_save(a):
    if a.get("lan_mode") == "1" and not core.users_count(): raise ValueError("Create the Owner user first (Settings > Users) - the shop network needs the sign-in screen")
    if "lan_port" in a and not (1024 <= num(a["lan_port"]) <= 65535): raise ValueError("The network port must be between 1024 and 65535")
    old = num(S("lbp_rate"))
    for k in core.DEFAULTS:
        if k in ("pin",) or k not in a: continue
        if k in ("smtp_pass", "backup_password"):
            if not a[k]: continue
            if a[k] == "__clear__": core.set_S(k, ""); continue
        core.set_S(k, a[k])
    if "lbp_rate" in a and num(a["lbp_rate"]) > 0 and num(a["lbp_rate"]) != old:
        x("INSERT INTO rate_history(date,rate,user) VALUES(?,?,?)", (core.now(), num(a["lbp_rate"]), core.CURRENT["user"] or "owner"))
    return True

@api
def lan_info(a): return {"wanted": S("lan_mode") == "1", "on": LAN["on"], "port": LAN["port"], "urls": [f"http://{ip}:{LAN['port']}/" for ip in local_ips()], "has_users": core.users_count() > 0}
@api
def backup_upload(a):
    name = os.path.basename(a.get("name") or "")
    if not re.fullmatch(r"pos_[A-Za-z0-9_\-]+\.zip(\.enc)?", name): raise ValueError("This is not a POS backup file (it should be named like pos_..._....zip)")
    path = os.path.join(core.BACKUP_DIR, name)
    with open(path, "wb") as f: f.write(base64.b64decode(a["b64"]))
    if name.endswith(".zip") and not core.zipfile.is_zipfile(path): os.remove(path); raise ValueError("The file is damaged")
    return name
@api
def users_public(a): return {"has_users": core.users_count() > 0}

@api
def login(a):
    name = core.norm_username(a.get("username") or a.get("name")); f = FAILS.get(name, [0, 0])
    if f[1] > time.time(): raise ValueError(f"Too many wrong attempts. Try again in {int(f[1] - time.time()) + 1} seconds")
    try: u = core.login(name, a.get("password") if a.get("password") is not None else a.get("pin"))
    except ValueError:
        f[0] += 1
        if f[0] >= 5: f = [0, time.time() + 60]
        FAILS[name] = f; core.audit("login_failed", name, user=name); raise
    FAILS.pop(name, None); sid = secrets.token_hex(16); SESSIONS[sid] = {"id": u["id"], "user": u["name"], "role": u["role"]}
    SESSION.update(SESSIONS[sid], sid=sid); core.CURRENT.update(user=u["name"], role=u["role"])
    core.audit("login", "", user=u["name"]); return {**me_info(), "session": sid}

@api
def recover(a):
    key = "recover"; f = FAILS.get(key, [0, 0])
    if f[1] > time.time(): raise ValueError(f"Too many wrong attempts. Try again in {int(f[1] - time.time()) + 1} seconds")
    try: code = core.recover_owner(a.get("username"), a.get("code"), a.get("password"))
    except ValueError as e:
        if "recovery code" in str(e) or "no owner" in str(e):
            f[0] += 1
            if f[0] >= 5: f = [0, time.time() + 60]
            FAILS[key] = f; core.audit("recover_failed", str(a.get("username"))[:40], user="")
        raise
    FAILS.pop(key, None); core.audit("owner_password_reset", core.norm_username(a.get("username")), user="recovery code"); return {"recovery": code}

@api
def recovery_new(a): return core.recovery_new()

@api
def logout(a): SESSIONS.pop(SESSION.get("sid"), None); SESSION.update(id=None, user="", role="cashier", sid=""); core.CURRENT.update(user="", role="owner"); return True

@api
def users(a): return rows("SELECT id,name,username,role FROM users WHERE active=1 ORDER BY role DESC, name")
@api
def user_save(a):
    first = core.users_count() == 0
    uid = core.user_save(a.get("id") or None, a.get("name"), a.get("role"), a.get("password") if a.get("password") is not None else (a.get("pin") or ""), a.get("username"))
    u = q1("SELECT * FROM users WHERE id=?", (uid,)); sid = ""; rec = ""
    if first:
        sid = secrets.token_hex(16); SESSIONS[sid] = {"id": u["id"], "user": u["name"], "role": "owner"}
        SESSION.update(SESSIONS[sid], sid=sid); core.CURRENT.update(user=u["name"], role="owner"); rec = core.recovery_new()
    for v in SESSIONS.values():
        if str(v["id"]) == str(uid): v.update(user=u["name"], role=u["role"])
    return {"id": uid, "session": sid, "recovery": rec}
@api
def user_delete(a):
    core.user_delete(a["id"])
    for k in [k for k, v in SESSIONS.items() if str(v["id"]) == str(a["id"])]: SESSIONS.pop(k, None)
    if str(SESSION["id"]) == str(a["id"]): SESSION.update(id=None, user="", role="cashier", sid="")
    return True
@api
def audit_list(a): return rows("SELECT * FROM audit ORDER BY id DESC LIMIT ?", (int(num(a.get("limit"), 200)),))
@api
def rate_history(a): return rows("SELECT * FROM rate_history ORDER BY id DESC LIMIT 60")
@api
def import_v4(a): core.import_v4(); return True
@api
def open_url(a): webbrowser.open(a["url"]); return True
@api
def open_folder(a):
    open_path({"data": core.DATA_DIR, "backups": core.BACKUP_DIR, "exports": EXPORT_DIR}[a["which"]]); return True

# ------------------------------------------------------------ products
@api
def products(a):
    sql = f"SELECT p.*, {OFFER} AS offer, (SELECT alfa FROM alfa_services sv WHERE sv.id=p.alfa_service_id) AS alfa_amt, {core.IMG_PRODUCT} AS img FROM products p WHERE p.active=1"; args = [today(), today()]
    t = (a.get("q") or "").strip()
    if t:
        sql += " AND (p.name LIKE ? OR p.barcode LIKE ? OR p.category LIKE ? OR p.id IN (SELECT product_id FROM imeis WHERE imei LIKE ? AND status='in_stock'))"
        args += [f"%{t}%"] * 4
    if a.get("cat") and a["cat"] != "All": sql += " AND p.category=?"; args.append(a["cat"])
    if a.get("low"): sql += " AND p.service=0 AND p.qty<=p.min_qty"
    return rows(sql + " ORDER BY p.category, p.name", args)

@api
def product_save(a):
    if a.get("id") and val("SELECT COALESCE(alfa_service_id, alfa_card_id) FROM products WHERE id=?", (a["id"],)): raise ValueError("This item is managed in the Alfa tab (services) or Validity days > Manage card types (cards)")
    return core.save_product(a, a.get("id") or None)
@api
def product_delete(a):
    if val("SELECT COALESCE(alfa_service_id, alfa_card_id) FROM products WHERE id=?", (a["id"],)): raise ValueError("This item is managed in the Alfa tab (services) or Validity days > Manage card types (cards)")
    x("UPDATE products SET active=0, barcode=NULL WHERE id=?", (a["id"],)); return True
@api
def stock_add(a):
    p = q1("SELECT * FROM products WHERE id=?", (a["id"],))
    if not p or p["service"] or p["track_imei"]: raise ValueError("Stock can't be added this way to this product")
    n = int(num(a.get("qty")))
    if n <= 0: raise ValueError("Enter a quantity")
    x("UPDATE products SET qty=qty+? WHERE id=?", (n, p["id"])); return p["qty"] + n
@api
def imeis(a): return rows("SELECT id,imei,status,cost FROM imeis WHERE product_id=? ORDER BY status, id", (a["product_id"],))
@api
def imeis_add(a):
    p = q1("SELECT * FROM products WHERE id=?", (a["product_id"],))
    core.add_imeis(p["id"], [s for s in re.split(r"[\s,;]+", a.get("text") or "") if s], p["cost"]); return True
@api
def imei_remove(a):
    r = q1("SELECT product_id,status FROM imeis WHERE id=?", (a["id"],))
    if r and r["status"] == "in_stock":
        x("DELETE FROM imeis WHERE id=?", (a["id"],)); core.sync_qty(r["product_id"]); core.db.commit()
    return True

@api
def scan(a):
    r = core.find_exact((a.get("code") or "").strip())
    if not r: return None
    d = rows(f"SELECT p.*, {OFFER} AS offer FROM products p WHERE id=?", (today(), today(), r[0]["id"]))[0]
    return {"product": d, "imei": r[1]}

# ------------------------------------------------------------ sales
@api
def customers_min(a): return rows("SELECT id,name,phone,points FROM customers ORDER BY name")

@api
def sale_create(a):
    cart = []
    for l in a["cart"]:
        pid = l.get("pid") or None; qty = int(num(l.get("qty"), 1))
        if pid:
            p = q1("SELECT * FROM products WHERE id=?", (pid,))
            cart.append({"pid": pid, "name": p["name"], "qty": qty, "price": num(l.get("price"), p["price"]), "cost": p["cost"], "imei": l.get("imei") or None})
        else:
            cost_ = num(l.get("cost"))
            if l.get("topup") and role() != "owner":          # cashiers can only sell saved top-ups; the profit comes from the saved item
                it_ = q1("SELECT * FROM topup_items WHERE name=?", (l["name"],))
                if not it_: raise ValueError("OWNER_PIN")
                cost_ = num(l["price"]) - it_["profit"]
            cart.append({"pid": None, "name": l["name"], "qty": qty, "price": num(l["price"]), "cost": cost_, "imei": l.get("imei") or None,
                         "topup": bool(l.get("topup")), "player": (l.get("player") or "").strip()})
    cur_ = "LBP" if a.get("cur") == "LBP" else "USD"; kind_ = a.get("kind") or "cash"; paid_in = num(a.get("paid"))
    lbp_ = paid_in if (cur_ == "LBP" and kind_ != "cash") else 0
    paid_usd = (paid_in / core.lbp_rate()) if lbp_ else paid_in
    sid = core.create_sale(cart, a.get("customer_id") or None, num(a.get("discount")), int(num(a.get("points_used"))), paid_usd,
                           a.get("method") or "Cash", a.get("kind") or "cash", a.get("due") or None,
                           " | ".join([t for t in [(a.get("notes") or "").strip()] + [f"Player ID {l['player']} ({l['name']})" for l in cart if l.get("topup") and l.get("player")] if t]),
                           int(num(a.get("inst_n"), 3)), cur_, lbp_)
    for l in cart:
        if l.get("topup"):
            for _ in range(int(l["qty"])): core.topup_record(sid, l["name"], l["price"], l["cost"], l.get("player"), role() == "owner")
    return core.sale_detail(sid)

@api
def sales_list(a): return core.sales_list(a["d1"], a["d2"], a.get("q") or "")
@api
def sale_get(a): return core.sale_detail(a["id"])
@api
def sale_void(a): core.void_sale(a["id"]); return True
@api
def sale_return(a): return core.return_item(a["item_id"], int(num(a["qty"], 1)), a.get("note", ""))
@api
def sale_pay(a): u, c, l = _pay_amt(a); core.add_payment(a["sale_id"], u, a.get("method", "Cash"), a.get("note", ""), c, l); return True

@api
def drawer(a): return core.drawer()
@api
def drawer_close(a):
    core.drawer_close(num(a.get("float")), num(a.get("counted")), num(a.get("float_lbp")), num(a.get("counted_lbp")))
    def bg():
        try:
            p = core.make_backup("close")
            if S("auto_drive") == "1": core.to_drive(p)
        except Exception: pass
        try:
            if S("auto_summary") == "1": core.send_summary(today())
        except Exception: pass
    threading.Thread(target=bg, daemon=True).start()
    return True
@api
def dashboard(a): return core.dashboard()

# ------------------------------------------------------------ customers / debts
@api
def customers(a):
    t = f"%{a.get('q') or ''}%"; d = today(); out = []
    for c in rows("SELECT * FROM customers WHERE name LIKE ? OR phone LIKE ? ORDER BY name", (t, t)):
        c["balance"] = round(core.customer_balance(c["id"]), 2)
        c["spent"] = round(val("SELECT SUM(total) FROM sales WHERE customer_id=?", (c["id"],)), 2)
        c["purchases"] = val("SELECT COUNT(*) FROM sales WHERE customer_id=?", (c["id"],))
        c["warranties"] = val("SELECT COUNT(*) FROM sale_items i JOIN sales s ON s.id=i.sale_id WHERE s.customer_id=? AND i.warranty_until>=?", (c["id"], d))
        if a.get("debtors") and c["balance"] <= 0: continue
        out.append(c)
    return out

@api
def customer_save(a):
    if not (a.get("name") or "").strip(): raise ValueError("Name is required")
    if a.get("id"):
        x("UPDATE customers SET name=?,phone=?,address=?,notes=? WHERE id=?", (a["name"].strip(), a.get("phone", ""), a.get("address", ""), a.get("notes", ""), a["id"])); return a["id"]
    return x("INSERT INTO customers(name,phone,address,notes,created) VALUES(?,?,?,?,?)", (a["name"].strip(), a.get("phone", ""), a.get("address", ""), a.get("notes", ""), today()))

@api
def customer_delete(a):
    if val("SELECT COUNT(*) FROM sales WHERE customer_id=?", (a["id"],)): raise ValueError("This customer has invoices and cannot be deleted")
    x("DELETE FROM customers WHERE id=?", (a["id"],)); return True

@api
def customer_get(a):
    cid = a["id"]; c = dict(q1("SELECT * FROM customers WHERE id=?", (cid,))); c["balance"] = round(core.customer_balance(cid), 2)
    return {"customer": c,
            "sales": rows(f"SELECT s.id,s.date,s.total,s.kind,s.due_date,{core.BAL} balance FROM sales s WHERE s.customer_id=? ORDER BY s.id DESC", (cid,)),
            "installments": rows("SELECT i.*, ROUND(i.amount-i.paid,2) remaining FROM installments i JOIN sales s ON s.id=i.sale_id WHERE s.customer_id=? ORDER BY i.due_date", (cid,)),
            "payments": rows("SELECT date,sale_id,amount,method,note FROM payments WHERE customer_id=? ORDER BY id DESC LIMIT 100", (cid,))}

@api
def customer_pay(a): u, c, l = _pay_amt(a); core.customer_pay(a["id"], u, a.get("method", "Cash"), a.get("note", ""), c, l); return True
@api
def warranty(a): return core.warranty_search(a.get("q") or "")

@api
def installments(a):
    d = today(); lim = (core.dt.date.today() + core.dt.timedelta(days=7)).isoformat(); f = a.get("flt", "Unpaid"); out = []
    for r in rows("SELECT i.id,c.id cid,c.name customer,c.phone,i.sale_id,i.no,i.due_date,i.amount,i.paid,ROUND(i.amount-i.paid,2) remaining FROM installments i "
                  "JOIN sales s ON s.id=i.sale_id JOIN customers c ON c.id=s.customer_id ORDER BY i.due_date"):
        r["status"] = "Paid" if r["remaining"] <= 0.004 else ("Overdue" if r["due_date"] < d else "Upcoming")
        if f == "Paid" and r["status"] != "Paid": continue
        if f == "Unpaid" and r["status"] == "Paid": continue
        if f == "Overdue" and r["status"] != "Overdue": continue
        if f == "Due in 7 days" and (r["status"] == "Paid" or r["due_date"] > lim): continue
        out.append(r)
    return out

@api
def debts(a):
    d = today()
    out = rows(f"SELECT s.id,substr(s.date,1,10) date,c.id cid,c.name customer,c.phone,s.total,{core.BAL} balance,s.due_date,s.kind FROM sales s "
               f"JOIN customers c ON c.id=s.customer_id WHERE {core.BAL}>0.004 ORDER BY s.due_date")
    for r in out: r["status"] = "Overdue" if (r["due_date"] and r["due_date"] < d) else "Open"
    return out

# ------------------------------------------------------------ suppliers
@api
def suppliers(a):
    out = []
    for r in rows("SELECT * FROM suppliers ORDER BY name"):
        i = r["id"]
        r.update(purchases=round(val("SELECT SUM(total) FROM purchases WHERE supplier_id=?", (i,)), 2), paid=round(val("SELECT SUM(amount) FROM supplier_payments WHERE supplier_id=?", (i,)), 2),
                 returns=round(val("SELECT SUM(amount) FROM supplier_returns WHERE supplier_id=?", (i,)), 2), balance=round(core.supplier_balance(i), 2))
        out.append(r)
    return out

@api
def supplier_save(a):
    if not (a.get("name") or "").strip(): raise ValueError("Name is required")
    if a.get("id"): x("UPDATE suppliers SET name=?,phone=?,notes=? WHERE id=?", (a["name"].strip(), a.get("phone", ""), a.get("notes", ""), a["id"])); return a["id"]
    return x("INSERT INTO suppliers(name,phone,notes) VALUES(?,?,?)", (a["name"].strip(), a.get("phone", ""), a.get("notes", "")))

@api
def supplier_get(a):
    i = a["id"]
    return {"supplier": dict(q1("SELECT * FROM suppliers WHERE id=?", (i,))), "balance": round(core.supplier_balance(i), 2),
            "purchases": rows("SELECT id,date,invoice_no,total FROM purchases WHERE supplier_id=? ORDER BY date DESC, id DESC", (i,)),
            "items": rows("SELECT p.date, COALESCE(p.invoice_no,p.id) invoice, i.name, i.qty, i.cost, i.imeis FROM purchase_items i JOIN purchases p ON p.id=i.purchase_id WHERE p.supplier_id=? ORDER BY p.date DESC, p.id DESC", (i,)),
            "payments": rows("SELECT date,amount,method,note FROM supplier_payments WHERE supplier_id=? ORDER BY date DESC, id DESC", (i,)),
            "returns": rows("SELECT date,name,qty,amount,imei,note FROM supplier_returns WHERE supplier_id=? ORDER BY date DESC, id DESC", (i,))}

@api
def purchase_create(a):
    items = []
    for l in a["items"]:
        p = q1("SELECT * FROM products WHERE id=?", (l["pid"],)); ims = [s for s in re.split(r"[\s,;]+", l.get("imeis") or "") if s]
        if p["track_imei"]:
            if not ims: raise ValueError(f"Enter the IMEIs for {p['name']}")
            if len(set(ims)) != len(ims): raise ValueError("Duplicate IMEI in the list")
            qty = len(ims)
        else: qty = int(num(l["qty"], 1))
        items.append({"pid": p["id"], "name": p["name"], "qty": qty, "cost": num(l.get("cost"), p["cost"]), "imeis": ims})
    return core.add_purchase(a["supplier_id"], a.get("invoice_no", ""), a.get("date") or today(), items, num(a.get("paid")), a.get("method") or "Cash", a.get("notes", ""))

# ------------------------------------------------------------ supplier invoice import (Excel / CSV)
HDR = [("name", ["item", "product", "description", "desc", "name", "البيان", "الصنف", "المنتج", "الوصف", "اسم"]),
       ("qty", ["qty", "quantity", "pcs", "count", "الكمية", "كمية", "العدد"]),
       ("barcode", ["barcode", "ean", "sku", "code", "الباركود", "باركود", "الرمز", "كود"]),
       ("imei", ["imei"]),
       ("category", ["category", "الفئة", "فئة"]),
       ("price", ["selling", "sell", "retail", "سعر البيع", "البيع"]),
       ("cost", ["unit cost", "cost", "unit price", "rate", "price", "الكلفة", "التكلفة", "الافرادي", "الإفرادي", "سعر الوحدة", "السعر", "سعر"])]
def _n(v): return re.sub(r"\s+", " ", str(v if v is not None else "").lower()).strip()
def _num(v):
    t = str(v if v is not None else "").strip().replace(" ", "")
    t = t.replace(",", ".") if re.fullmatch(r"-?[^0-9.,\-]*\d+,\d{1,2}[^0-9.,]*", t) else t.replace(",", "")
    t = re.sub(r"[^0-9.\-]", "", t)
    try: return float(t) if t not in ("", "-", ".") else 0.0
    except ValueError: return 0.0
def _read_table(name, raw):
    if name.lower().endswith((".xlsx", ".xlsm")):
        if core.openpyxl is None: raise ValueError("Excel files are not supported in this build - save the invoice as CSV")
        ws = core.openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True).worksheets[0]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    try: txt = raw.decode("utf-8-sig")
    except UnicodeDecodeError: txt = raw.decode("cp1256", errors="replace")
    try: dia = csv.Sniffer().sniff(txt[:3000], delimiters=",;\t")
    except csv.Error: dia = csv.excel
    return list(csv.reader(io.StringIO(txt), dia))
def _find_cols(table):
    for i, r in enumerate(table[:25]):
        cells = [_n(c) for c in r]; cols = {}; used = set()
        for key, syns in HDR:
            for j, c in enumerate(cells):
                if c and j not in used and any(sy in c for sy in syns): cols[key] = j; used.add(j); break
        if "name" in cols and "qty" in cols and "cost" in cols: return i, cols
    return None, None

@api
def invoice_parse(a):
    table = _read_table(a.get("name") or "", base64.b64decode(a["b64"]))
    hi, cols = _find_cols(table)
    if hi is None: raise ValueError("I could not find the columns. The first row needs headers like: Item, Qty, Cost (optional: Barcode, IMEI, Selling price, Category).")
    items, order = {}, []
    for r in table[hi + 1:]:
        g = lambda k: (r[cols[k]] if k in cols and cols[k] < len(r) else None)
        nm = str(g("name") or "").strip()
        if not nm or re.search(r"\b(total|subtotal|grand total)\b|المجموع|الاجمالي|الإجمالي", nm.lower()): continue
        ims = [t for t in re.split(r"[\s,;]+", str(g("imei") or "").replace(".0", "")) if t]
        qty = _num(g("qty"))
        if not ims and qty <= 0: continue
        bc = str(g("barcode") or "").strip()
        if bc.endswith(".0"): bc = bc[:-2]
        key = bc or nm.lower()
        it = items.get(key)
        if not it:
            it = items[key] = {"name": nm, "barcode": bc, "qty": 0, "cost": _num(g("cost")), "price": _num(g("price")) or None, "category": str(g("category") or "").strip(), "imeis": []}
            order.append(key)
        if ims: it["imeis"] += ims; it["qty"] = len(it["imeis"])
        else: it["qty"] += int(round(qty))
    out = []
    for k in order:
        it = items[k]
        p = q1("SELECT * FROM products WHERE active=1 AND barcode=?", (it["barcode"],)) if it["barcode"] else None
        if not p: p = q1("SELECT * FROM products WHERE active=1 AND lower(name)=?", (it["name"].lower(),))
        it["pid"] = p["id"] if p else None; it["matched"] = p["name"] if p else ""
        it["track_imei"] = bool(p["track_imei"]) if p else bool(it["imeis"])
        it["error"] = "This phone needs IMEIs" if it["track_imei"] and not it["imeis"] else ("" if it["qty"] > 0 else "No quantity")
        out.append(it)
    if not out: raise ValueError("No items found in this file")
    return out

@api
def supplier_pay(a): core.pay_supplier(a["id"], num(a["amount"]), a.get("method", "Cash"), a.get("note", ""), a.get("date") or None); return True
@api
def supplier_return(a):
    core.supplier_return(a["supplier_id"], a["product_id"], int(num(a.get("qty"), 1)), num(a["amount"]), a.get("imei") or None, a.get("note", "")); return True

# ------------------------------------------------------------ expenses / offers / repairs
@api
def expenses(a): return rows("SELECT * FROM expenses WHERE date BETWEEN ? AND ? ORDER BY date DESC, id DESC", (a["d1"], a["d2"]))
@api
def expense_save(a):
    if num(a.get("amount")) <= 0: raise ValueError("Enter an amount")
    core.parse_date(a.get("date") or today())
    if a.get("id"): x("UPDATE expenses SET date=?,category=?,amount=?,note=? WHERE id=?", (a["date"], a["category"], num(a["amount"]), a.get("note", ""), a["id"])); return a["id"]
    return x("INSERT INTO expenses(date,category,amount,note) VALUES(?,?,?,?)", (a.get("date") or today(), a["category"], num(a["amount"]), a.get("note", "")))
@api
def expense_delete(a): x("DELETE FROM expenses WHERE id=?", (a["id"],)); return True

@api
def offers(a):
    d = today()
    out = rows("SELECT o.id,o.product_id,p.name,p.price regular,o.price,o.start,o.end,o.note FROM offers o JOIN products p ON p.id=o.product_id ORDER BY o.end DESC")
    for r in out:
        r["status"] = "Active" if r["start"] <= d <= r["end"] else ("Upcoming" if r["start"] > d else "Expired")
        r["pct"] = round((1 - r["price"] / r["regular"]) * 100) if r["regular"] else 0
    return out
@api
def offer_save(a):
    core.parse_date(a["start"]); core.parse_date(a["end"])
    if a["end"] < a["start"]: raise ValueError("The end date is before the start date")
    if num(a.get("price")) < 0: raise ValueError("Invalid price")
    if a.get("id"): x("UPDATE offers SET product_id=?,price=?,start=?,end=?,note=? WHERE id=?", (a["product_id"], num(a["price"]), a["start"], a["end"], a.get("note", ""), a["id"])); return a["id"]
    return x("INSERT INTO offers(product_id,price,start,end,note) VALUES(?,?,?,?,?)", (a["product_id"], num(a["price"]), a["start"], a["end"], a.get("note", "")))
@api
def offer_delete(a): x("DELETE FROM offers WHERE id=?", (a["id"],)); return True

@api
def repairs(a):
    return rows("SELECT r.*, COALESCE(c.name, NULLIF(r.customer,''), 'Walk-in') customer FROM repairs r LEFT JOIN customers c ON c.id=r.customer_id ORDER BY r.id DESC")
@api
def repair_save(a):
    if not (a.get("device") or "").strip(): raise ValueError("Device is required")
    if a.get("price") in (None, ""): raise ValueError("Price is required")
    v = (a.get("customer_id") or None, a["device"].strip(), (a.get("imei") or "").strip(), (a.get("issue") or "").strip(), num(a["price"]), num(a.get("cost")), int(num(a.get("warranty_days"))))
    if a.get("id"): x("UPDATE repairs SET customer_id=?,device=?,imei=?,issue=?,price=?,cost=?,warranty_days=? WHERE id=?", v + (a["id"],)); return a["id"]
    return x("INSERT INTO repairs(customer_id,device,imei,issue,price,cost,warranty_days,date,status) VALUES(?,?,?,?,?,?,?,?,'Received')", v + (today(),))
@api
def repair_status(a): x("UPDATE repairs SET status=? WHERE id=? AND status!='Delivered'", (a["status"], a["id"])); return True
@api
def repair_deliver(a): return core.sale_detail(core.repair_deliver(a["id"], a.get("method", "Cash")))
@api
def repair_delete(a): x("DELETE FROM repairs WHERE id=?", (a["id"],)); return True

# ------------------------------------------------------------ reports
@api
def report_list(a): return list(core.REPORTS)
@api
def report(a):
    cols, rs = core.REPORTS[a["name"]](a["d1"], a["d2"]); return {"cols": cols, "rows": rs}
@api
def report_export(a):
    cols, rs = core.REPORTS[a["name"]](a["d1"], a["d2"]); os.makedirs(EXPORT_DIR, exist_ok=True)
    ext = "xlsx" if a.get("fmt") != "csv" and core.openpyxl else "csv"
    path = os.path.join(EXPORT_DIR, re.sub(r"[^A-Za-z0-9]+", "_", a["name"]).strip("_") + f"_{today()}.{ext}")
    core.export_rows(path, a["name"], cols, rs); open_path(path); return path

# ------------------------------------------------------------ backup
def _bpath(name): return os.path.join(core.BACKUP_DIR, os.path.basename(name))
@api
def backup_now(a): return os.path.basename(core.make_backup("manual"))
@api
def backup_list(a):
    return [{"name": os.path.basename(f), "kb": round(os.path.getsize(f) / 1024), "date": time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(f)))} for f in core.list_backups()]
@api
def backup_restore(a): core.restore_backup(_bpath(a["name"])); return True
@api
def backup_drive(a): return core.to_drive(_bpath(a["name"]) if a.get("name") else core.make_backup("manual"))
@api
def backup_email(a): return core.email_backup(_bpath(a["name"]) if a.get("name") else core.make_backup("manual"))
@api
def import_csv(a):
    fd, p = tempfile.mkstemp(suffix=".csv"); os.close(fd)
    with open(p, "w", encoding="utf-8") as f: f.write(a["text"])
    try: return core.import_products(p)
    finally: os.remove(p)

# ------------------------------------------------------------ Alfa credit / validity days / wipe
@api
def alfa_state(a): return core.alfa_state(core.prov(a.get("provider")))
@api
def alfa_services(a): return rows("SELECT s.*, (SELECT id||'.'||ver FROM images WHERE kind='service' AND ref=CAST(s.id AS TEXT)) AS img FROM alfa_services s WHERE s.active=1 AND s.provider=? ORDER BY s.name", (core.prov(a.get("provider")),))
@api
def alfa_service_save(a):
    name = (a.get("name") or "").strip()
    if not name: raise ValueError("Service name is required")
    if num(a.get("alfa")) <= 0: raise ValueError("Enter how many Alfa dollars the service uses")
    if num(a.get("price")) < 0: raise ValueError("Invalid price")
    fav = 1 if str(a.get("fav")) in ("1", "Yes", "true", "True") else 0
    if a.get("id"): x("UPDATE alfa_services SET name=?,alfa=?,price=?,fav=? WHERE id=?", (name, num(a["alfa"]), num(a.get("price")), fav, a["id"])); core.sync_alfa_products(); return a["id"]
    r = x("INSERT INTO alfa_services(name,alfa,price,fav,provider) VALUES(?,?,?,?,?)", (name, num(a["alfa"]), num(a.get("price")), fav, core.prov(a.get("provider")))); core.sync_alfa_products(); return r
@api
def alfa_service_delete(a): x("UPDATE alfa_services SET active=0 WHERE id=?", (a["id"],)); core.sync_alfa_products(); return True
@api
def alfa_adjust(a): core.alfa_adjust(num(a.get("amount")), num(a.get("cost")), a.get("note", ""), core.prov(a.get("provider"))); return True
@api
def alfa_log(a): return core.alfa_log(40, core.prov(a.get("provider")))
@api
def alfa_set(a): core.alfa_set(num(a.get("amount")), num(a.get("avg")), a.get("note", ""), core.prov(a.get("provider"))); return True
@api
def alfa_adjustments(a): return rows("SELECT id,date,amount,cost,note,kind FROM alfa_ledger WHERE kind IN ('adjust','reset') AND provider=? ORDER BY id DESC LIMIT 30", (core.prov(a.get("provider")),))
@api
def alfa_adjust_delete(a): core.alfa_adjust_delete(a["id"]); return True
@api
def alfa_cards(a): return rows("SELECT c.id,c.name,c.cost,c.price, (SELECT id||'.'||ver FROM images WHERE kind='card' AND ref=CAST(c.id AS TEXT)) AS img FROM alfa_cards c WHERE c.active=1 AND c.provider=? ORDER BY c.cost, c.name", (core.prov(a.get("provider")),))
@api
def alfa_card_save(a):
    name = (a.get("name") or "").strip()
    if not name: raise ValueError("Card name is required")
    if num(a.get("cost")) < 0: raise ValueError("Invalid cost")
    if a.get("id"): x("UPDATE alfa_cards SET name=?,cost=?,price=? WHERE id=?", (name, num(a.get("cost")), num(a.get("price")), a["id"])); core.sync_alfa_products(); return a["id"]
    r = x("INSERT INTO alfa_cards(name,cost,price,provider) VALUES(?,?,?,?)", (name, num(a.get("cost")), num(a.get("price")), core.prov(a.get("provider")))); core.sync_alfa_products(); return r
@api
def alfa_card_delete(a): x("UPDATE alfa_cards SET active=0 WHERE id=?", (a["id"],)); core.sync_alfa_products(); return True
@api
def validity_cards(a):
    return rows(f"SELECT p.id,p.name,p.cost,p.qty, {core.IMG_PRODUCT} AS img FROM products p WHERE p.active=1 AND p.service=0 AND p.track_imei=0 AND p.category LIKE ? ORDER BY p.name", (core.PROVIDERS[core.prov(a.get("provider"))]["inv"],))
@api
def validity_sale(a):
    sid = core.validity_sale(a.get("months"), [str(i) for i in (a.get("cards") or []) if str(i).strip()], num(a.get("ret")), num(a.get("price")),
                             a.get("cur") or "USD", a.get("received"), a.get("customer_id") or None, "", core.prov(a.get("provider")))
    return core.sale_detail(sid)
@api
def alfa_sell(a):
    pr_ = core.prov(a.get("provider"))
    if a.get("type") == "card":
        sid = core.card_sell(a.get("card_id"), num(a.get("price")), a.get("cur") or "USD", a.get("received"), a.get("customer_id") or None, "", pr_)
        return core.sale_detail(sid)
    sid = core.alfa_sell(a.get("type") or "svc", a.get("service_id") or None, num(a.get("amount")), num(a.get("price")),
                         a.get("cur") or "USD", a.get("received"), a.get("customer_id") or None, "", pr_)
    return core.sale_detail(sid)
@api
def wipe_all(a):
    if (a.get("confirm") or "").strip().upper() != "DELETE": raise ValueError("Type DELETE to confirm")
    if core.users_count() and not core.check_owner_pin(a.get("pin", "")): raise ValueError("Wrong owner password")
    path = core.make_backup("before_wipe")
    core.wipe_all()
    return os.path.basename(path)

# ------------------------------------------------------------ cash moves, recurring, presets, summary, count
@api
def cash_move_save(a):
    if num(a.get("amount")) <= 0: raise ValueError("Enter an amount")
    if a.get("kind") not in ("in", "out"): raise ValueError("Choose cash in or cash out")
    return x("INSERT INTO cash_moves(date,kind,amount,cur,reason,user) VALUES(?,?,?,?,?,?)",
             (core.now(), a["kind"], num(a["amount"]), "LBP" if a.get("cur") == "LBP" else "USD", (a.get("reason") or "").strip(), core.CURRENT["user"]))
@api
def cash_moves(a): return rows("SELECT * FROM cash_moves WHERE substr(date,1,10) BETWEEN ? AND ? ORDER BY id DESC", (a.get("d1") or today(), a.get("d2") or today()))
@api
def cash_move_delete(a): x("DELETE FROM cash_moves WHERE id=?", (a["id"],)); return True

@api
def recurring_list(a): return rows("SELECT * FROM recurring WHERE active=1 ORDER BY category")
@api
def recurring_save(a):
    if num(a.get("amount")) <= 0: raise ValueError("Enter an amount")
    day = min(28, max(1, int(num(a.get("day"), 1))))
    if a.get("id"): x("UPDATE recurring SET category=?,amount=?,day=?,note=? WHERE id=?", (a["category"], num(a["amount"]), day, a.get("note", ""), a["id"])); return a["id"]
    return x("INSERT INTO recurring(category,amount,day,note) VALUES(?,?,?,?)", (a["category"], num(a["amount"]), day, a.get("note", "")))
@api
def recurring_delete(a): x("UPDATE recurring SET active=0 WHERE id=?", (a["id"],)); return True

@api
def validity_presets(a): return rows("SELECT * FROM validity_presets WHERE provider=? ORDER BY name", (core.prov(a.get("provider")),))
@api
def preset_save(a):
    name = (a.get("name") or "").strip()
    if not name: raise ValueError("Give the preset a name")
    cards = [str(i) for i in (a.get("cards") or []) if str(i).strip()]
    if not cards: raise ValueError("Choose at least one card first")
    return x("INSERT INTO validity_presets(name,months,cards,ret,price,provider) VALUES(?,?,?,?,?,?)", (name, int(num(a.get("months"), 1)), json.dumps(cards), num(a.get("ret")), num(a.get("price")), core.prov(a.get("provider"))))
@api
def preset_delete(a): x("DELETE FROM validity_presets WHERE id=?", (a["id"],)); return True

@api
def day_summary(a): return core.day_summary(a.get("date") or today())
@api
def summary_send(a): return core.send_summary(a.get("date") or today())
@api
def week_chart(a): return core.week_chart()
@api
def stock_count_apply(a): return core.stock_count_apply(a.get("items") or [])

def needs_owner(name, a):
    if name in SENSITIVE: return True
    if name == "sale_create":
        sub = sum(num(l.get("price")) * int(num(l.get("qty"), 1)) for l in a.get("cart", []))
        under = 0.0
        for l in a.get("cart", []):
            if l.get("pid"):
                p = q1("SELECT price FROM products WHERE id=?", (l["pid"],))
                off = val("SELECT MIN(price) FROM offers WHERE product_id=? AND start<=? AND end>=?", (l["pid"], today(), today())) or p["price"]
                under += max(0.0, (min(p["price"], off) - num(l.get("price"))) * int(num(l.get("qty"), 1)))
        limit = num(S("max_discount_pct"), 10) / 100.0 * sub
        return (num(a.get("discount")) + under) > limit + 0.004
    return False

def redact(name, out):
    # cashiers never see costs or profits
    def clean(r):
        if isinstance(r, dict):
            for k in ("cost", "profit", "gross", "avg"):
                if k in r: r[k] = 0
        return r
    if name in ("products", "scan", "validity_cards", "alfa_cards", "topup_items", "alfa_log", "sales_list", "imeis", "alfa_services"):
        if isinstance(out, dict):
            clean(out.get("product") or {})
        elif isinstance(out, list):
            for r in out: clean(r)
    elif name in ("sale_get", "sale_create", "validity_sale", "alfa_sell", "repair_deliver"):
        for it in out.get("items", []): clean(it)
    elif name == "alfa_state": clean(out); out["cost"] = 0
    elif name == "dashboard": out["profit"] = None
    elif name == "day_summary":
        for k in ("gross", "profit", "alfa_profit", "touch_profit", "expenses"): out[k] = None
        out["alfa"]["cost"] = 0; out["alfa"]["avg"] = 0; out["touch"]["cost"] = 0; out["touch"]["avg"] = 0
    elif name == "week_chart":
        for r in out: r["profit"] = None
    return out

def summarize(a, out):
    d = {k: v for k, v in a.items() if k not in HIDE_KEYS}
    t = json.dumps(d, ensure_ascii=False, default=str)
    if "_opin" in a: t += " [owner password used]"
    if isinstance(out, dict) and out.get("id") and "total" in out: t += f" -> #{out['id']} total {out['total']}"
    elif isinstance(out, int) and not isinstance(out, bool): t += f" -> {out}"
    return t[:380]

# ------------------------------------------------------------ trade-in, warranty claims, bulk prices
@api
def tradein_save(a): return core.tradein_save(a)
@api
def tradeins(a): return core.tradeins_list()
@api
def claim_save(a): return core.claim_save(a)
@api
def claim_update(a): core.claim_update(a); return True
@api
def claims(a): return core.claims_list(a.get("status"))
@api
def claim_delete(a): x("DELETE FROM claims WHERE id=?", (a["id"],)); return True
@api
def price_bulk(a): return core.price_bulk(a.get("category") or "", num(a.get("pct")), num(a.get("round_to")))
@api
def price_undo_info(a): return core.price_undo_info()
@api
def price_undo(a): return core.price_undo()
@api
def categories(a): return [r[0] for r in q("SELECT DISTINCT category FROM products WHERE active=1 AND alfa_service_id IS NULL AND alfa_card_id IS NULL AND COALESCE(category,'')!='' ORDER BY category")]

@api
def dash_rows(a):
    # sales lines of the last ~6 months in the "Sales for dashboard" format; the Dashboard tab does the maths in the browser
    d2 = today(); d1 = (core.dt.date.today() - core.dt.timedelta(days=190)).isoformat()
    cols, rs = core.REPORTS["Sales for dashboard"](d1, d2)
    return [dict(zip(cols, r)) for r in rs]

@api
def image_set(a):
    kind = a.get("kind")
    if kind not in ("product", "service", "card", "tab", "sq"): raise ValueError("Unknown picture type")
    data = base64.b64decode(a["b64"])
    if len(data) > 2_000_000: raise ValueError("This picture is too big")
    return core.image_set(kind, a["ref"], data, "image/png" if a.get("mime") == "image/png" else "image/jpeg")
@api
def image_delete(a): core.image_delete(a.get("kind"), a["ref"]); return True

@api
def topup_items(a): return rows("SELECT * FROM topup_items ORDER BY uses DESC, last DESC LIMIT 16")
@api
def topup_item_delete(a): x("DELETE FROM topup_items WHERE id=?", (a["id"],)); return True

# ------------------------------------------------------------ program updates from a file
def _owner_pw(a):
    if core.users_count() and not core.check_owner_pin(a.get("password", "")): raise ValueError("Wrong owner password")
@api
def update_info(a): return core.update_info(VERSION)
@api
def update_check(a): return core.update_check(base64.b64decode(a["b64"]), VERSION)
@api
def update_install(a): _owner_pw(a); return core.update_install(base64.b64decode(a["b64"]))
@api
def update_use(a): _owner_pw(a); core.update_use(a["version"]); return True
@api
def app_restart(a):
    if os.environ.get("POS_NO_RESTART"): return "skipped"
    def go():
        time.sleep(1.2)
        try:
            import subprocess
            if getattr(sys, "frozen", False): cmd = [sys.executable]
            else: cmd = [sys.executable, os.environ.get("POS_LAUNCHER") or os.path.abspath(__file__)]
            env = {k: v for k, v in os.environ.items() if not k.startswith(("_MEI", "_PYI"))}
            env["POS_RESTART_DELAY"] = "2.5"; env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
            subprocess.Popen(cmd, env=env, close_fds=True, creationflags=(0x00000008 | 0x00000200) if os.name == "nt" else 0)
        finally:
            os._exit(0)
    threading.Thread(target=go, daemon=True).start(); return "restarting"

# ------------------------------------------------------------ http
class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        self.send_response(code); self.send_header("Content-Type", ctype); self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store"); self.end_headers(); self.wfile.write(body)
    def _host_ok(self):
        h = (self.headers.get("Host") or "").split(":")[0]
        if h in ("127.0.0.1", "localhost"): return True
        if not LAN["on"]: return False
        try:
            ip = ipaddress.ip_address(self.client_address[0])
            if not (ip.is_private or ip in TAILSCALE): return False     # only devices inside the shop network, or on your private Tailscale network
        except ValueError: return False
        return re.fullmatch(r"(10\.\d+\.\d+\.\d+|192\.168\.\d+\.\d+|172\.(1[6-9]|2\d|3[01])\.\d+\.\d+|100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])\.\d+\.\d+)", h) is not None
    def do_GET(self):
        if self.path.split("?")[0] in ("/", "/index.html") and self._host_ok():
            with open(resource("index.html"), encoding="utf-8") as f: html = f.read().replace("__TOKEN__", TOKEN)
            self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
        elif self.path.startswith("/img/") and self._host_ok():
            try:
                qs = dict(kv.split("=", 1) for kv in (self.path.split("?", 1)[1] if "?" in self.path else "").split("&") if "=" in kv)
                if qs.get("t") != TOKEN: return self._send(403, b"{}")
                with LOCK: r = q1("SELECT mime, data FROM images WHERE id=?", (int(self.path.split("?")[0][5:]),))
            except Exception: return self._send(404, b"{}")
            if not r: return self._send(404, b"{}")
            body = bytes(r["data"]); self.send_response(200); self.send_header("Content-Type", r["mime"] or "image/jpeg"); self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "private, max-age=31536000, immutable"); self.end_headers(); self.wfile.write(body)
        else: self._send(404, b"{}")
    def do_POST(self):
        name = self.path.split("/api/", 1)[-1]
        if not self._host_ok() or self.headers.get("X-Token") != TOKEN or name not in API: return self._send(403, b"{}")
        sid = self.headers.get("X-Session") or ""; sess = SESSIONS.get(sid)
        try:
            a = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
            with LOCK:
                SESSION.update(sess if sess else {"id": None, "user": "", "role": "cashier"}, sid=sid if sess else "")
                core.CURRENT.update(user=SESSION["user"], role=SESSION["role"] if sess else "owner")
                try:
                    if name not in PUBLIC:
                        if core.users_count() and not SESSION["user"]: raise ValueError("LOCKED")
                        r = role()
                        if name in OWNER_ONLY and r != "owner": raise ValueError("This action is for the owner only")
                        if r != "owner" and needs_owner(name, a):
                            if "_opin" in a and not core.check_owner_pin(a.get("_opin")): raise ValueError("Wrong owner password")
                            if "_opin" not in a: raise ValueError("OWNER_PIN")
                    try: res = API[name](a)
                    except sqlite3.OperationalError as e:      # old database: upgrade it, then retry once
                        if "no such" not in str(e): raise
                        core.migrate(); res = API[name](a)
                    if name in AUDITED: core.audit(name, summarize(a, res))
                    if name not in PUBLIC and role() != "owner": res = redact(name, res)
                    out = {"ok": True, "data": res}
                except sqlite3.OperationalError as e: raise
        except ValueError as e: out = {"ok": False, "error": str(e)}
        except Exception as e:
            traceback.print_exc(); out = {"ok": False, "error": f"{type(e).__name__}: {e}"}
        self._send(200, json.dumps(out, default=str).encode("utf-8"))

def local_ips():
    ips = set()
    try:
        so = socket.socket(socket.AF_INET, socket.SOCK_DGRAM); so.connect(("10.255.255.255", 1)); ips.add(so.getsockname()[0]); so.close()
    except Exception: pass
    try: ips.update(i[4][0] for i in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET))
    except Exception: pass
    return sorted(i for i in ips if not i.startswith("127.") and not i.startswith("169.254."))

def serve():
    lan = S("lan_mode") == "1" and core.users_count() > 0       # network mode needs the sign-in screen (at least the Owner user)
    LAN["on"] = lan
    try: srv = http.server.ThreadingHTTPServer(("0.0.0.0" if lan else "127.0.0.1", int(num(S("lan_port"), 8765)) if lan else 0), Handler)
    except OSError: srv = http.server.ThreadingHTTPServer(("0.0.0.0" if lan else "127.0.0.1", 0), Handler)
    LAN["port"] = srv.server_address[1]
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, f"http://127.0.0.1:{srv.server_address[1]}/"

def on_exit():
    try:
        if S("auto_backup") == "1":
            fl = core.list_backups()
            if not fl or time.time() - os.path.getmtime(fl[0]) > 1800:
                p = core.make_backup("exit")
                if S("auto_drive") == "1":
                    try: core.to_drive(p)
                    except Exception: pass
    except Exception: pass

def backup_loop():
    while True:
        time.sleep(60)
        try:
            with LOCK: core.hourly_backup()
        except Exception: pass

def main():
    srv, url = serve()
    threading.Thread(target=backup_loop, daemon=True).start()
    try:
        p = core.auto_backup()
        if p and S("auto_drive") == "1": threading.Thread(target=lambda: core.to_drive(p), daemon=True).start()
    except Exception: pass
    try:
        import webview
        webview.create_window(S("shop_name") or "Shop POS", url, width=1320, height=840, min_size=(980, 640))
        webview.start()
    except Exception:
        webbrowser.open(url); start = time.time()      # fallback: default browser; exits when the page is closed
        while True:
            time.sleep(2)
            if (PING["seen"] and time.time() - PING["t"] > 30) or (not PING["seen"] and time.time() - start > 180): break
    on_exit()

if __name__ == "__main__":
    main()
