"""
Blue Sky Smog - Desktop App  (PyQt6 rewrite)
Requires: pip install pyqt6 requests reportlab pymupdf
"""

import os, sys, json, uuid, sqlite3, threading, time, re, textwrap, urllib.request
from pathlib import Path
from datetime import datetime, timedelta

APP_VERSION = "1.1.27"
_UPDATE_API  = "https://api.github.com/repos/blueskysmog1/bluesky-smog-installer/releases/latest"
_DOWNLOAD_URL = "https://github.com/blueskysmog1/bluesky-smog-installer/releases/latest/download/BlueSkyDesktop_Setup.exe"

# â"€â"€ PyQt6 â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QStackedWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QGroupBox, QScrollArea, QFrame, QTabWidget,
    QCheckBox, QRadioButton, QSpinBox,
    QMessageBox, QFileDialog, QDialog, QDialogButtonBox,
    QMenu, QSizePolicy, QSplitter, QProgressBar,
    QCalendarWidget, QDateEdit,
)
from PyQt6.QtCore import (
    Qt, QTimer, QThread, pyqtSignal, QDate, QPoint, QRect, QObject,
)
from PyQt6.QtGui import (
    QFont, QColor, QIcon, QPixmap, QBrush, QAction, QImage,
    QPainter, QPen, QPalette, QCursor,
)

try:
    import win32print, win32api
    _WIN32_PRINT = True
except ImportError:
    _WIN32_PRINT = False

try:
    import fitz
    _FITZ_OK = True
except ImportError:
    _FITZ_OK = False

try:
    import requests
except ImportError:
    requests = None

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.graphics.barcode import code128

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  CONFIG & CONSTANTS
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

API_BASE      = "https://api.blueskysmog.net"

def _icon_path():
    if getattr(sys, "_MEIPASS", None):
        return os.path.join(sys._MEIPASS, "logo.ico")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.ico")

DEVICE_ID     = f"DESKTOP-{uuid.getnode()}"
SYNC_INTERVAL = 10
APP_NAME      = "BlueSkyDesktop"
_base         = os.environ.get("APPDATA") or os.path.expanduser("~")
DATA_DIR      = Path(_base) / APP_NAME
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH       = str(DATA_DIR / "local.db")
CREDS_FILE    = str(DATA_DIR / "creds.json")
LOG_FILE      = str(DATA_DIR / "sync.log")

import logging as _logging
_log_handler = _logging.FileHandler(LOG_FILE, encoding="utf-8")
_log_handler.setFormatter(_logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
_sync_log = _logging.getLogger("sync")
_sync_log.setLevel(_logging.DEBUG)
_sync_log.addHandler(_log_handler)

def slog(msg):
    try: _sync_log.info(msg)
    except Exception: pass

PRIMARY  = "#005B99"
ACCENT   = "#00AEEF"
BG       = "#F5F7FB"
TEXT     = "#111827"
WHITE    = "#FFFFFF"
RED      = "#DC2626"
GREEN    = "#16A34A"
TODAY_BG = "#EFF6FF"
DARK_HDR = "#37474F"

DEFAULT_BUSINESS = {
    "name": "", "address_line1": "", "address_line2": "",
    "phone": "", "email": "", "ard": "", "card_fee": 5.00,
    "logo_path": "",
    "invoice_notice": (
        "I authorize {business_name} to perform the indicated services. "
        "I am responsible for removing all valuable property from my vehicle prior to service. "
        "I shall inspect my vehicle on the premises after services are rendered."
    ),
}
DEFAULT_SERVICES = {
    "Smog Test":           {"price": 51.75, "cert_fee": 8.25},
    "Clean Truck OBDII":   {"price": 120.0,  "cert_fee": 0.0},
    "Clean Truck Opacity": {"price": 180.0,  "cert_fee": 0.0},
}

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  DATABASE
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS customers (
            customer_id TEXT PRIMARY KEY, first_name TEXT NOT NULL DEFAULT '',
            last_name TEXT NOT NULL DEFAULT '', company_name TEXT NOT NULL DEFAULT '',
            phone TEXT NOT NULL DEFAULT '', email TEXT NOT NULL DEFAULT '',
            address TEXT NOT NULL DEFAULT '', city TEXT NOT NULL DEFAULT '',
            state TEXT NOT NULL DEFAULT '', zip TEXT NOT NULL DEFAULT '',
            referral_code TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL, synced INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS vehicles (
            vehicle_id TEXT PRIMARY KEY, customer_id TEXT NOT NULL,
            vin TEXT NOT NULL DEFAULT '', plate TEXT NOT NULL DEFAULT '',
            make TEXT NOT NULL DEFAULT '', model TEXT NOT NULL DEFAULT '',
            year TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL,
            FOREIGN KEY(customer_id) REFERENCES customers(customer_id)
        );
        CREATE INDEX IF NOT EXISTS idx_vehicles_plate ON vehicles(plate);
        CREATE INDEX IF NOT EXISTS idx_vehicles_vin   ON vehicles(vin);
        CREATE TABLE IF NOT EXISTS invoices (
            invoice_id TEXT PRIMARY KEY, invoice_number INTEGER NOT NULL DEFAULT 0,
            customer_id TEXT NOT NULL DEFAULT '', customer_name TEXT NOT NULL DEFAULT '',
            first_name TEXT NOT NULL DEFAULT '', last_name TEXT NOT NULL DEFAULT '',
            company_name TEXT NOT NULL DEFAULT '', invoice_date TEXT NOT NULL DEFAULT '',
            plate TEXT NOT NULL DEFAULT '', vin TEXT NOT NULL DEFAULT '',
            year TEXT NOT NULL DEFAULT '', make TEXT NOT NULL DEFAULT '',
            model TEXT NOT NULL DEFAULT '', amount_cents INTEGER NOT NULL DEFAULT 0,
            payment_method TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'DRAFT',
            notes TEXT NOT NULL DEFAULT '', is_estimate INTEGER NOT NULL DEFAULT 0,
            from_mobile INTEGER NOT NULL DEFAULT 0, pdf_path TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL, synced INTEGER NOT NULL DEFAULT 0
        );
        CREATE INDEX IF NOT EXISTS idx_invoices_date     ON invoices(invoice_date);
        CREATE INDEX IF NOT EXISTS idx_invoices_customer ON invoices(customer_id);
        CREATE INDEX IF NOT EXISTS idx_invoices_plate    ON invoices(plate);
        CREATE TABLE IF NOT EXISTS invoice_lines (
            id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id TEXT NOT NULL,
            vin TEXT NOT NULL DEFAULT '', plate TEXT NOT NULL DEFAULT '',
            odometer TEXT NOT NULL DEFAULT '', year TEXT NOT NULL DEFAULT '',
            make TEXT NOT NULL DEFAULT '', model TEXT NOT NULL DEFAULT '',
            service TEXT NOT NULL DEFAULT '', result TEXT NOT NULL DEFAULT '',
            cert TEXT NOT NULL DEFAULT '', discount REAL NOT NULL DEFAULT 0.0,
            price REAL NOT NULL DEFAULT 0.0,
            FOREIGN KEY(invoice_id) REFERENCES invoices(invoice_id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS accounts (
            company_name TEXT PRIMARY KEY, total_owed REAL NOT NULL DEFAULT 0.0,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS account_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_name TEXT NOT NULL,
            entry_date TEXT NOT NULL, type TEXT NOT NULL, amount REAL NOT NULL,
            note TEXT NOT NULL DEFAULT '', invoice_id TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS outbox (
            id INTEGER PRIMARY KEY AUTOINCREMENT, event_id TEXT NOT NULL UNIQUE,
            entity TEXT NOT NULL, action TEXT NOT NULL,
            payload TEXT NOT NULL, created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS sync_state (key TEXT PRIMARY KEY, val TEXT NOT NULL);
    """)
    cols = {row[1] for row in c.execute("PRAGMA table_info(invoice_lines)").fetchall()}
    if "remote_item_id" not in cols:
        c.execute("ALTER TABLE invoice_lines ADD COLUMN remote_item_id TEXT NOT NULL DEFAULT ''")
    c.execute("DROP INDEX IF EXISTS idx_invoice_lines_remote_item_id")
    c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_invoice_lines_remote_item_id_partial "
              "ON invoice_lines(remote_item_id) WHERE remote_item_id != ''")
    c.execute("CREATE INDEX IF NOT EXISTS idx_invoice_lines_invoice_id ON invoice_lines(invoice_id)")
    conn.commit(); conn.close()

def migrate_db():
    conn = get_db(); c = conn.cursor()
    inv_cols = {row[1] for row in c.execute("PRAGMA table_info(invoices)").fetchall()}
    for col, defn in [
        ("veh_state","TEXT NOT NULL DEFAULT 'CA'"), ("owner_first","TEXT NOT NULL DEFAULT ''"),
        ("owner_last","TEXT NOT NULL DEFAULT ''"),  ("account_id","TEXT NOT NULL DEFAULT ''"),
        ("po_number","TEXT NOT NULL DEFAULT ''"),    ("test_result","TEXT NOT NULL DEFAULT ''"),
        ("cert_number","TEXT NOT NULL DEFAULT ''"),
    ]:
        if col not in inv_cols:
            c.execute(f"ALTER TABLE invoices ADD COLUMN {col} {defn}")
    acct_cols = {row[1] for row in c.execute("PRAGMA table_info(accounts)").fetchall()}
    for col, defn in [
        ("contact_name","TEXT NOT NULL DEFAULT ''"), ("phone","TEXT NOT NULL DEFAULT ''"),
        ("email","TEXT NOT NULL DEFAULT ''"),        ("address1","TEXT NOT NULL DEFAULT ''"),
        ("address2","TEXT NOT NULL DEFAULT ''"),     ("city","TEXT NOT NULL DEFAULT ''"),
        ("state","TEXT NOT NULL DEFAULT ''"),        ("zip","TEXT NOT NULL DEFAULT ''"),
        ("account_status","TEXT NOT NULL DEFAULT 'Active'"),
        ("tax_exempt","INTEGER NOT NULL DEFAULT 0"), ("require_po","INTEGER NOT NULL DEFAULT 0"),
        ("payment_types","TEXT NOT NULL DEFAULT '[]'"),
        ("custom_pricing","TEXT NOT NULL DEFAULT '{}'"),
        ("track_vehicles","INTEGER NOT NULL DEFAULT 0"),
    ]:
        if col not in acct_cols:
            c.execute(f"ALTER TABLE accounts ADD COLUMN {col} {defn}")
    # account_history migrations
    hist_cols = {row[1] for row in c.execute("PRAGMA table_info(account_history)").fetchall()}
    if "payment_number" not in hist_cols:
        c.execute("ALTER TABLE account_history ADD COLUMN payment_number TEXT NOT NULL DEFAULT ''")
    if "payment_id" not in hist_cols:
        c.execute("ALTER TABLE account_history ADD COLUMN payment_id TEXT NOT NULL DEFAULT ''")
    if "partial_json" not in hist_cols:
        c.execute("ALTER TABLE account_history ADD COLUMN partial_json TEXT NOT NULL DEFAULT '{}'")
    c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_ah_payment_id "
              "ON account_history(payment_id) WHERE payment_id != ''")
    # customers migration
    cust_cols = {row[1] for row in c.execute("PRAGMA table_info(customers)").fetchall()}
    if "discount_percent" not in cust_cols:
        c.execute("ALTER TABLE customers ADD COLUMN discount_percent REAL NOT NULL DEFAULT 0.0")
    if "discount_type" not in cust_cols:
        c.execute("ALTER TABLE customers ADD COLUMN discount_type TEXT NOT NULL DEFAULT 'PERCENT'")
    conn.commit(); conn.close()

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  CREDENTIALS & API
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

def load_creds():
    try:
        with open(CREDS_FILE, "r") as f: return json.load(f)
    except Exception: return {}

def save_creds(d):
    with open(CREDS_FILE, "w") as f: json.dump(d, f, indent=2)

def _hdrs():
    c = load_creds()
    h = {"Content-Type": "application/json"}
    if c.get("token"):    h["x-token"]    = c["token"]
    if c.get("username"): h["x-username"] = c["username"]
    if c.get("password"): h["x-password"] = c["password"]
    return h

def api_login(username, password):
    r = requests.get(f"{API_BASE}/v1/auth/login",
                     headers={"x-username": username, "x-password": password}, timeout=10)
    if not r.ok:
        try: detail = r.json().get("detail", "Sign in failed.")
        except Exception: detail = "Sign in failed."
        raise ValueError(detail)
    d = r.json()
    if not d.get("success"): raise ValueError("Login rejected")
    return d["token"], d["company_id"], d.get("company_name","")

def api_push(events):
    formatted = [{"event_id": ev["event_id"], "seq": i, "entity": ev["entity"],
                  "action": ev["action"], "payload": ev["payload"]}
                 for i, ev in enumerate(events)]
    r = requests.post(f"{API_BASE}/v1/sync/push",
                      json={"device_id": DEVICE_ID, "events": formatted},
                      headers=_hdrs(), timeout=15)
    r.raise_for_status(); return r.json()

def api_pull(since_seq):
    r = requests.get(f"{API_BASE}/v1/sync/pull/{DEVICE_ID}",
                     params={"since_seq": since_seq}, headers=_hdrs(), timeout=15)
    r.raise_for_status(); return r.json().get("events", [])

def api_subscription_status():
    try:
        r = requests.get(f"{API_BASE}/v1/subscription/status", headers=_hdrs(), timeout=8)
        if r.status_code == 200: return r.json()
    except Exception: pass
    return {}

def api_decode_vin(vin):
    try:
        url = f"https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVin/{vin}?format=json"
        req = urllib.request.Request(url, headers={"User-Agent": "BlueSkyDesktop"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read())
        res = {i["Variable"]: i["Value"] for i in data["Results"]}
        return res.get("Model Year",""), res.get("Make",""), res.get("Model","")
    except Exception: return "","",""

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  LOCAL DB HELPERS
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

def now_iso(): return datetime.utcnow().isoformat()
def get_setting(conn, key, default=""):
    r = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return r["value"] if r else default
def set_setting(conn, key, value):
    conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", (key, str(value))); conn.commit()
def get_last_seq(conn):
    r = conn.execute("SELECT val FROM sync_state WHERE key='last_seq'").fetchone()
    return int(r["val"]) if r else 0
def set_last_seq(conn, seq):
    conn.execute("INSERT OR REPLACE INTO sync_state VALUES ('last_seq',?)", (str(seq),)); conn.commit()
def enqueue(conn, entity, action, payload: dict):
    conn.execute("INSERT OR IGNORE INTO outbox(event_id,entity,action,payload,created_at) VALUES(?,?,?,?,?)",
                 (str(uuid.uuid4()), entity, action, json.dumps(payload), now_iso())); conn.commit()
def get_or_create_customer_id(conn, first, last, company):
    key = f"{first} {last}".strip().upper() or company.upper()
    r = conn.execute(
        "SELECT customer_id FROM customers WHERE UPPER(first_name||' '||last_name)=? OR UPPER(company_name)=?",
        (key, key)).fetchone()
    if r: return r["customer_id"]
    cid = str(uuid.uuid4())
    conn.execute("INSERT INTO customers(customer_id,first_name,last_name,company_name,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                 (cid, first, last, company, now_iso(), now_iso())); conn.commit()
    return cid
def upsert_customer(conn, first, last, company, phone="", email="",
                    address="", city="", state="", zip_="", synced=0, customer_id=None,
                    discount_percent=0.0, discount_type="PERCENT"):
    if not customer_id: customer_id = get_or_create_customer_id(conn, first, last, company)
    conn.execute("UPDATE customers SET first_name=?,last_name=?,company_name=?,phone=?,email=?,address=?,city=?,state=?,zip=?,discount_percent=?,discount_type=?,updated_at=?,synced=? WHERE customer_id=?",
                 (first,last,company,phone,email,address,city,state,zip_,discount_percent,discount_type,now_iso(),synced,customer_id))
    if conn.execute("SELECT changes()").fetchone()[0] == 0:
        conn.execute("INSERT OR IGNORE INTO customers(customer_id,first_name,last_name,company_name,phone,email,address,city,state,zip,discount_percent,discount_type,created_at,updated_at,synced) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                     (customer_id,first,last,company,phone,email,address,city,state,zip_,discount_percent,discount_type,now_iso(),now_iso(),synced))
    conn.commit(); return customer_id
def upsert_vehicle(conn, customer_id, vin, plate, make, model, year, vehicle_id=None):
    if vehicle_id:
        r = conn.execute("SELECT vehicle_id FROM vehicles WHERE vehicle_id=?", (vehicle_id,)).fetchone()
        if r:
            conn.execute("UPDATE vehicles SET customer_id=?,vin=?,plate=?,make=?,model=?,year=?,updated_at=? WHERE vehicle_id=?",
                         (customer_id,vin,plate,make,model,year,now_iso(),vehicle_id)); conn.commit(); return vehicle_id
        conn.execute("INSERT OR IGNORE INTO vehicles(vehicle_id,customer_id,vin,plate,make,model,year,updated_at) VALUES(?,?,?,?,?,?,?,?)",
                     (vehicle_id,customer_id,vin,plate,make,model,year,now_iso())); conn.commit(); return vehicle_id
    r = conn.execute("SELECT vehicle_id FROM vehicles WHERE (plate!='' AND plate=?) OR (vin!='' AND vin=?)", (plate,vin)).fetchone()
    if r:
        conn.execute("UPDATE vehicles SET customer_id=?,vin=?,plate=?,make=?,model=?,year=?,updated_at=? WHERE vehicle_id=?",
                     (customer_id,vin,plate,make,model,year,now_iso(),r["vehicle_id"])); conn.commit(); return r["vehicle_id"]
    vid = str(uuid.uuid4())
    conn.execute("INSERT INTO vehicles(vehicle_id,customer_id,vin,plate,make,model,year,updated_at) VALUES(?,?,?,?,?,?,?,?)",
                 (vid,customer_id,vin,plate,make,model,year,now_iso())); conn.commit(); return vid
def get_next_invoice_number(conn): return 0
def get_business_settings(conn):
    raw = get_setting(conn,"business","")
    if raw:
        try:
            b = json.loads(raw)
            for k,v in DEFAULT_BUSINESS.items(): b.setdefault(k,v)
            return b
        except Exception: pass
    return dict(DEFAULT_BUSINESS)
def get_services(conn):
    raw = get_setting(conn,"services","")
    if raw:
        try: return json.loads(raw)
        except: pass
    return dict(DEFAULT_SERVICES)
def get_printer_setting(conn):
    raw = get_setting(conn,"printer_setting","")
    default = {"mode":"pdf","printer_name":"","copies":1,"auto_print":False}
    if raw:
        try: d=json.loads(raw); default.update(d)
        except: pass
    return default
def display_customer_name(first="",last="",company="",customer_name=""):
    person = f"{first} {last}".strip()
    return (company or person or customer_name or "Customer").strip()
def format_phone(raw):
    digits = re.sub(r"\D","",raw or "")
    if len(digits)==10: return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    if len(digits)==11 and digits[0]=="1": return f"{digits[1:4]}-{digits[4:7]}-{digits[7:]}"
    return raw or ""
def safe_filename_part(text):
    text = (text or "").strip()
    if not text: return "Customer"
    text = re.sub(r"[^A-Za-z0-9]+","_",text)
    return re.sub(r"_+","_",text).strip("_") or "Customer"

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  SYNC ENGINE  (identical to original)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

class SyncEngine:
    def __init__(self):
        self._lock=threading.Lock(); self._running=False
        self._on_change=None; self._on_suspended=None
        self._last_pull_time=None; self._last_pull_count=0; self._last_push_count=0
    def set_on_change(self,cb): self._on_change=cb
    def set_on_suspended(self,cb): self._on_suspended=cb
    def start(self):
        self._running=True; t=threading.Thread(target=self._loop,daemon=True); t.start()
    def stop(self): self._running=False
    def force_pull_from_zero(self, notify_cb=None):
        """Clear all local data and re-pull everything from seq=0.
        Fetches events FIRST before clearing so a network failure never
        leaves the database empty.  notify_cb(msg) is called on completion."""
        slog("[ForcePull] Waiting for sync lock...")
        if not self._lock.acquire(blocking=True, timeout=30):
            msg="Could not acquire sync lock - try again in a moment."
            slog(f"[ForcePull] {msg}")
            if notify_cb: notify_cb(msg)
            return
        slog("[ForcePull] Lock acquired - fetching events from server...")
        conn=get_db()
        try:
            # â"€â"€ Step 1: fetch ALL events first (safe - no data deleted yet) â"€â"€
            events=api_pull(0)
            n=len(events) if events else 0
            slog(f"[ForcePull] Fetched {n} events from server")
            # â"€â"€ Step 2: only now wipe and rebuild â"€â"€
            # Delete in reverse-FK order: children before parents
            # vehicles.customer_id -> customers, so vehicles must go before customers
            for t in ("invoice_lines","invoices","vehicles","customers"):
                conn.execute(f"DELETE FROM {t}")
            set_last_seq(conn,0); conn.commit()
            # â"€â"€ Step 3: apply events in dependency order â"€â"€
            # Bucket first, then apply: customers -> vehicles -> invoices -> items
            # This avoids FK violations when events arrive out of entity order.
            if events:
                new_seq=0
                cust_evs=[]; veh_evs=[]; inv_evs=[]; item_evs=[]; pay_evs=[]
                for ev in events:
                    seq=int(ev.get("seq",0))
                    entity=ev.get("entity",""); action=ev.get("action","")
                    payload=ev.get("payload",{})
                    if isinstance(payload,str):
                        try: payload=json.loads(payload)
                        except: payload={}
                    if entity=="customer" and action=="upsert":
                        cust_evs.append(payload)
                    elif entity=="vehicle" and action=="upsert":
                        veh_evs.append(payload)
                    elif entity=="invoice" and action in ("upsert","finalize","delete"):
                        inv_evs.append((action,payload))
                    elif entity=="invoice_item" and action in ("upsert","insert","create"):
                        item_evs.append(payload)
                    elif entity=="account_payment" and action in ("upsert","delete"):
                        pay_evs.append((action,payload))
                    new_seq=max(new_seq,seq)
                # Pass 1 - customers
                for p in cust_evs:
                    try: self._merge_customer(conn,p)
                    except Exception as e2: slog(f"[ForcePull] customer SKIPPED err={e2}")
                # Pass 2 - vehicles (reference customers)
                for p in veh_evs:
                    try: self._merge_vehicle(conn,p)
                    except Exception as e2: slog(f"[ForcePull] vehicle SKIPPED err={e2}")
                # Pass 3 - invoices (reference customers/vehicles)
                for act,p in inv_evs:
                    try:
                        if act=="delete": self._delete_invoice(conn,p)
                        else: self._merge_invoice(conn,p)
                    except Exception as e2: slog(f"[ForcePull] invoice SKIPPED err={e2}")
                # Pass 4 - invoice_items (reference invoices)
                for p in item_evs:
                    try: self._merge_invoice_item(conn,p)
                    except Exception as e2: slog(f"[ForcePull] invoice_item SKIPPED err={e2}")
                # Pass 5 - payments (after invoices so company lookup works)
                for act,p in pay_evs:
                    try:
                        if act=="delete": self._merge_payment_delete(conn,p)
                        else: self._merge_payment(conn,p)
                    except Exception as e2: slog(f"[ForcePull] payment SKIPPED err={e2}")
                set_last_seq(conn,new_seq)
                self._last_pull_count=n; self._last_pull_time=datetime.now()
                if self._on_change:
                    try: self._on_change()
                    except: pass
            msg=f"Re-pull complete.\n{n} events loaded from server." if n>0 else \
                "Re-pull complete.\nNo events found on server (database may be empty)."
            slog(f"[ForcePull] Done - {msg}")
            if notify_cb: notify_cb(msg)
        except Exception as e:
            msg=f"Re-pull FAILED: {e}"
            slog(f"[ForcePull] {msg}")
            if notify_cb: notify_cb(msg)
        finally:
            conn.close()
            try: self._lock.release()
            except: pass
    def _loop(self):
        while self._running:
            try: self._flush(); self._pull()
            except Exception as e: slog(f"[Loop] unhandled: {e}")
            time.sleep(SYNC_INTERVAL)
    def _flush(self):
        if not requests: return
        conn=get_db()
        try:
            rows=conn.execute("SELECT id,event_id,entity,action,payload FROM outbox ORDER BY id LIMIT 50").fetchall()
            if not rows: return
            events=[]; ids=[]
            for row in rows:
                try:
                    events.append({"event_id":row["event_id"],"entity":row["entity"],"action":row["action"],"payload":json.loads(row["payload"])})
                    ids.append(row["id"])
                except Exception: pass
            api_push(events)
            conn.execute(f"DELETE FROM outbox WHERE id IN ({','.join('?'*len(ids))})",ids); conn.commit()
            self._last_push_count=len(events)
        except Exception as e:
            slog(f"[Push] FAILED err={e}")
            try:
                if hasattr(e,'response') and e.response is not None and e.response.status_code==403:
                    msg=e.response.json().get("detail","Account suspended.")
                    if self._on_suspended: self._on_suspended(msg)
            except Exception: pass
        finally: conn.close()
    def _pull(self):
        """Non-blocking pull - skips if the lock is already held."""
        if not requests: return
        if not self._lock.acquire(blocking=False): return
        conn=get_db()
        try:
            self._do_pull(conn)
        except Exception as e: slog(f"[Pull] FAILED err={e}")
        finally:
            conn.close()
            try: self._lock.release()
            except: pass
    def _do_pull(self,conn):
        """Inner pull logic - caller must already hold self._lock."""
        if not requests: return
        since=get_last_seq(conn); events=api_pull(since)
        if not events: return
        new_seq=since; item_events=[]
        for ev in events:
            seq=int(ev.get("seq",0)); entity=ev.get("entity",""); action=ev.get("action",""); payload=ev.get("payload",{})
            if isinstance(payload,str):
                try: payload=json.loads(payload)
                except: payload={}
            try:
                if entity=="customer" and action=="upsert": self._merge_customer(conn,payload)
                elif entity=="vehicle" and action=="upsert": self._merge_vehicle(conn,payload)
                elif entity=="invoice" and action in ("upsert","finalize"): self._merge_invoice(conn,payload)
                elif entity=="invoice" and action=="delete": self._delete_invoice(conn,payload)
                elif entity=="invoice_item" and action in ("upsert","insert","create"): item_events.append(payload)
                elif entity=="account_payment" and action=="upsert": self._merge_payment(conn,payload)
                elif entity=="account_payment" and action=="delete": self._merge_payment_delete(conn,payload)
            except Exception as ev_err: slog(f"[Pull] seq={seq} SKIPPED err={ev_err}")
            new_seq=max(new_seq,seq)
        for payload in item_events:
            try: self._merge_invoice_item(conn,payload)
            except Exception as ev_err: slog(f"[Pull] invoice_item SKIPPED err={ev_err}")
        set_last_seq(conn,new_seq)
        self._last_pull_count=len(events); self._last_pull_time=datetime.now()
        slog(f"[Pull] fetched {len(events)} events, new_seq={new_seq}")
        if self._on_change:
            try: self._on_change()
            except: pass
    def _merge_payment(self,conn,p):
        """Apply a remote account_payment upsert to local account_history."""
        payment_id=p.get("payment_id","")
        if not payment_id: return
        # Skip if already applied
        if conn.execute("SELECT 1 FROM account_history WHERE payment_id=?",(payment_id,)).fetchone(): return
        company_name=p.get("company_name","")
        customer_id=p.get("customer_id","")
        # Resolve company_name from customer_id if missing
        if not company_name and customer_id:
            r=conn.execute("SELECT company_name FROM customers WHERE customer_id=?",(customer_id,)).fetchone()
            if r: company_name=(r["company_name"] or "").strip()
        # Resolve company_name from referenced invoice if still missing
        if not company_name:
            for iid in p.get("invoice_id","").split(","):
                iid=iid.strip()
                if not iid: continue
                r=conn.execute("SELECT company_name,account_id FROM invoices WHERE invoice_id=?",(iid,)).fetchone()
                if r:
                    company_name=(r["company_name"] or r["account_id"] or "").strip()
                    if company_name: break
        if not company_name: return
        amount_cents=int(p.get("amount_cents",0))
        amount=amount_cents/100.0
        entry_date=p.get("entry_date",now_iso()[:10])
        note=p.get("note",""); invoice_id=p.get("invoice_id",""); payment_number=p.get("payment_number","")
        partial_json=p.get("partial_json","{}")
        conn.execute(
            "INSERT OR IGNORE INTO account_history(company_name,entry_date,type,amount,note,invoice_id,payment_number,payment_id,partial_json) "
            "VALUES(?,?,?,?,?,?,?,?,?)",
            (company_name,entry_date,"payment",amount,note,invoice_id,payment_number,payment_id,partial_json))
        conn.execute("UPDATE accounts SET total_owed=MAX(0,total_owed-?),updated_at=? WHERE company_name=?",
                     (amount,now_iso(),company_name))
        conn.commit()
        slog(f"[Payment] merged {payment_number} for {company_name} ${amount:.2f}")

    def _merge_payment_delete(self,conn,p):
        """Apply a remote account_payment delete to local account_history."""
        payment_id=p.get("payment_id","")
        if not payment_id: return
        row=conn.execute("SELECT amount,company_name FROM account_history WHERE payment_id=?",(payment_id,)).fetchone()
        if not row: return
        amount=row["amount"]; company_name=row["company_name"]
        conn.execute("DELETE FROM account_history WHERE payment_id=?",(payment_id,))
        conn.execute("UPDATE accounts SET total_owed=total_owed+?,updated_at=? WHERE company_name=?",
                     (amount,now_iso(),company_name))
        conn.commit()
        slog(f"[Payment] deleted payment_id={payment_id} for {company_name}")

    def _merge_vehicle(self,conn,p):
        vid=p.get("vehicle_id",""); cid=p.get("customer_id","")
        if not cid: return
        upsert_vehicle(conn,customer_id=cid,vin=p.get("vin",""or""),plate=p.get("plate",""or""),
                       make=p.get("make",""or""),model=p.get("model",""or""),year=p.get("year",""or""),
                       vehicle_id=vid if vid else None)
    def _delete_invoice(self,conn,p):
        iid=p.get("invoice_id","")
        if not iid: return
        conn.execute("DELETE FROM invoice_lines WHERE invoice_id=?",(iid,))
        conn.execute("DELETE FROM invoices WHERE invoice_id=?",(iid,)); conn.commit()
    def _merge_customer(self,conn,p):
        try:
            first=p.get("first_name",""or""); last=p.get("last_name",""or""); company=p.get("company_name",""or"")
            name=(p.get("name",""or"")).strip()
            if name and not first and not last and not company:
                parts=name.strip().split(" ",1); first=parts[0]; last=parts[1] if len(parts)>1 else ""
            cid=(p.get("customer_id",""or"")).strip()
            disc=float(p.get("discount_percent") or 0)
            disc_type=str(p.get("discount_type") or "PERCENT").upper()
            if disc_type not in ("PERCENT","FLAT"): disc_type="PERCENT"
            upsert_customer(conn,first,last,company,phone=format_phone(p.get("phone")or""),
                            email=p.get("email")or"",address=p.get("address")or"",city=p.get("city")or"",
                            state=p.get("state")or"",zip_=p.get("zip")or"",synced=1,
                            customer_id=cid if cid else None,
                            discount_percent=disc, discount_type=disc_type)
        except Exception as e: slog(f"[Merge] customer FAILED err={e}")
    def _merge_invoice(self,conn,p):
        try:
            iid=p.get("invoice_id","")
            if not iid: return
            existing=conn.execute("SELECT from_mobile,synced,status,invoice_number FROM invoices WHERE invoice_id=?",(iid,)).fetchone()
            incoming_num=int(p.get("invoice_number",0)or 0)
            if existing and not existing["from_mobile"]:
                if incoming_num and incoming_num!=existing["invoice_number"]:
                    conn.execute("UPDATE invoices SET invoice_number=?,synced=1 WHERE invoice_id=?",(incoming_num,iid)); conn.commit()
                elif not existing["synced"]:
                    conn.execute("UPDATE invoices SET synced=1 WHERE invoice_id=?",(iid,)); conn.commit()
                return
            first=p.get("first_name",""or""); last=p.get("last_name",""or""); company=p.get("company_name",""or"")
            name=p.get("customer_name",""or"")or f"{first} {last}".strip()
            cid=(p.get("customer_id",""or"")).strip()
            if cid:
                if not conn.execute("SELECT 1 FROM customers WHERE customer_id=?",(cid,)).fetchone():
                    conn.execute("INSERT OR IGNORE INTO customers(customer_id,first_name,last_name,company_name,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                                 (cid,first,last,company,now_iso(),now_iso())); conn.commit()
            else: cid=get_or_create_customer_id(conn,first,last,company)
            status=(p.get("status",""or"")or"DRAFT").upper(); is_est=1 if status=="ESTIMATE" else 0
            inv_num=p.get("invoice_number",0)or 0; plate=p.get("plate",""or""); vin=p.get("vin",""or"")
            year=p.get("year",""or""); make=p.get("make",""or""); model=p.get("model",""or"")
            notes=p.get("notes")or""; pay_method=p.get("payment_method")or""; invoice_date=p.get("invoice_date")or""; amount_cents=int(p.get("amount_cents")or 0)
            conn.execute("INSERT OR IGNORE INTO invoices(invoice_id,invoice_number,customer_id,customer_name,first_name,last_name,company_name,invoice_date,plate,vin,year,make,model,amount_cents,payment_method,status,notes,is_estimate,from_mobile,created_at,updated_at,synced) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?,1)",
                         (iid,inv_num,cid,name,first,last,company,invoice_date,plate,vin,year,make,model,amount_cents,pay_method,status,notes,is_est,now_iso(),now_iso()))
            conn.execute("UPDATE invoices SET invoice_number=?,customer_id=?,customer_name=?,first_name=?,last_name=?,company_name=?,invoice_date=?,plate=?,vin=?,year=?,make=?,model=?,amount_cents=?,payment_method=?,status=?,notes=?,is_estimate=?,from_mobile=1,updated_at=?,synced=1 WHERE invoice_id=?",
                         (inv_num,cid,name,first,last,company,invoice_date,plate,vin,year,make,model,amount_cents,pay_method,status,notes,is_est,now_iso(),iid))
            if plate or vin: upsert_vehicle(conn,cid,vin,plate,make,model,year)
            conn.commit()
        except Exception as e: slog(f"[Merge] invoice FAILED err={e}")
    def _merge_invoice_item(self,conn,p):
        try:
            invoice_id=p.get("invoice_id","")or p.get("parent_id","")
            if not invoice_id: return
            remote_item_id=p.get("id","")or p.get("item_id","")or""
            service=p.get("service","")or p.get("name","")or p.get("description","")or""
            qty=p.get("qty",1)or 1
            try: qty=float(qty)
            except: qty=1.0
            unit_price_cents=p.get("unit_price_cents",None)
            if unit_price_cents is None: unit_price_cents=p.get("price_cents",None)
            if unit_price_cents is None:
                raw_price=p.get("price",p.get("unit_price",0))
                try:
                    raw_price=float(raw_price)
                    # Old mobile bug: stored amount_cents (e.g. 5000) as price instead
                    # of dollars (e.g. 50.0).
                    # Heuristic 1: if price exactly matches invoice's amount_cents -> cents
                    # Heuristic 2: if price is a whole number â‰¥ 500 -> almost certainly cents
                    #   (a smog shop service priced at $500+ would use the new unit_price_cents field)
                    if raw_price>0 and raw_price==int(raw_price):
                        _is_cents=False
                        try:
                            inv_row=conn.execute("SELECT amount_cents FROM invoices WHERE invoice_id=?",(invoice_id,)).fetchone()
                            if inv_row and inv_row["amount_cents"] and abs(raw_price-float(inv_row["amount_cents"]))<1.0:
                                _is_cents=True
                        except: pass
                        if not _is_cents and raw_price>=500:
                            _is_cents=True   # large whole number without unit_price_cents -> cents
                        if _is_cents:
                            raw_price=raw_price/100.0
                    price=raw_price*qty
                except: price=0.0
            else:
                try: price=(float(unit_price_cents)/100.0)*qty
                except: price=0.0
            vin=p.get("vin","")or""; plate=p.get("plate","")or""; odometer=p.get("odometer","")or p.get("odo","")or""
            year=p.get("year","")or""; make=p.get("make","")or""; model=p.get("model","")or""
            if not (vin or plate or year or make or model):
                inv=conn.execute("SELECT * FROM invoices WHERE invoice_id=?",(invoice_id,)).fetchone()
                if inv:
                    vin=(inv["vin"]or"").strip(); plate=(inv["plate"]or"").strip()
                    year=(inv["year"]or"").strip(); make=(inv["make"]or"").strip(); model=(inv["model"]or"").strip()
            result=p.get("result","")or p.get("status","")or""
            cert=p.get("cert","")or p.get("certificate","")or""
            try:
                disc_raw=p.get("discount",None)
                if disc_raw is None:
                    dc=p.get("discount_cents",0)
                    discount=float(dc or 0)/100.0
                else:
                    discount=float(disc_raw or 0)
            except: discount=0.0
            if remote_item_id:
                existing=conn.execute("SELECT id FROM invoice_lines WHERE remote_item_id=?",(remote_item_id,)).fetchone()
                if existing:
                    conn.execute("UPDATE invoice_lines SET invoice_id=?,vin=?,plate=?,odometer=?,year=?,make=?,model=?,service=?,result=?,cert=?,discount=?,price=? WHERE remote_item_id=?",
                                 (invoice_id,vin,plate,odometer,year,make,model,service,result,cert,discount,price,remote_item_id))
                else:
                    conn.execute("INSERT OR REPLACE INTO invoice_lines(invoice_id,vin,plate,odometer,year,make,model,service,result,cert,discount,price,remote_item_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                 (invoice_id,vin,plate,odometer,year,make,model,service,result,cert,discount,price,remote_item_id))
            else:
                conn.execute("INSERT INTO invoice_lines(invoice_id,vin,plate,odometer,year,make,model,service,result,cert,discount,price,remote_item_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                             (invoice_id,vin,plate,odometer,year,make,model,service,result,cert,discount,price,""))
            try: conn.commit()
            except: conn.rollback(); conn.commit()
        except Exception as e: slog(f"[Merge] invoice_item FAILED err={e}")

SYNC = SyncEngine()

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  PDF  (identical to original)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

def print_pdf(pdf_path, printer_name="", copies=1, parent_widget=None):
    import subprocess, shutil
    # 1. SumatraPDF — silent direct-to-printer, no dialog
    sumatra = shutil.which("SumatraPDF") or r"C:\Program Files\SumatraPDF\SumatraPDF.exe"
    if os.path.exists(sumatra):
        try:
            target = printer_name or "default"
            for _ in range(copies): subprocess.Popen([sumatra,"-print-to",target,"-silent",pdf_path])
            return
        except Exception: pass
    # 2. win32print GDI path (only available when pywin32 is installed)
    if _WIN32_PRINT:
        try:
            import fitz, win32print, win32ui, win32con
            from PIL import Image; import io
            pn = printer_name or win32print.GetDefaultPrinter()
            for _ in range(copies):
                hprinter=win32print.OpenPrinter(pn)
                try:
                    hdc=win32ui.CreateDC(); hdc.CreatePrinterDC(pn)
                    dpi_x=hdc.GetDeviceCaps(88); dpi_y=hdc.GetDeviceCaps(90)
                    page_w=hdc.GetDeviceCaps(110); page_h=hdc.GetDeviceCaps(111)
                    doc=fitz.open(pdf_path); hdc.StartDoc(pdf_path)
                    for page_num in range(len(doc)):
                        hdc.StartPage(); page=doc[page_num]
                        zoom=min(dpi_x,dpi_y)/72.0; mat=fitz.Matrix(zoom,zoom)
                        pix=page.get_pixmap(matrix=mat,alpha=False); img_data=pix.tobytes("ppm")
                        pil_img=Image.open(io.BytesIO(img_data)).convert("RGB")
                        img_w,img_h=pil_img.size; scale=min(page_w/img_w,page_h/img_h)
                        new_w=int(img_w*scale); new_h=int(img_h*scale)
                        pil_img=pil_img.resize((new_w,new_h),Image.LANCZOS)
                        try:
                            import PIL.ImageWin as _iw; dib=_iw.Dib(pil_img)
                            dib.draw(hdc.GetHandleAttrib(),(0,0,new_w,new_h))
                        except Exception:
                            tmp=os.path.join(os.environ.get("TEMP",os.getcwd()),"_bs_print.bmp")
                            pil_img.save(tmp,"BMP"); bmp=win32ui.CreateBitmap(); bmp.LoadBitmapFile(tmp)
                            mem_dc=hdc.CreateCompatibleDC(); mem_dc.SelectObject(bmp)
                            hdc.StretchBlt((0,0),(new_w,new_h),mem_dc,(0,0),(new_w,new_h),win32con.SRCCOPY)
                            mem_dc.DeleteDC()
                        hdc.EndPage()
                    doc.close(); hdc.EndDoc(); hdc.DeleteDC()
                finally: win32print.ClosePrinter(hprinter)
            return
        except Exception: pass
    # 3. Qt + fitz — direct print to named printer, no dialog, always available
    if printer_name:
        try:
            import fitz
            from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo
            from PyQt6.QtGui import QImage, QPainter
            from PyQt6.QtCore import Qt, QRect
            infos = QPrinterInfo.availablePrinters()
            info = next((p for p in infos if p.printerName() == printer_name), None)
            if info is None and infos:
                info = infos[0]  # best guess if name changed
            if info:
                qp = QPrinter(info, QPrinter.PrinterMode.HighResolution)
                qp.setCopyCount(copies)
                doc = fitz.open(pdf_path)
                painter = QPainter()
                if painter.begin(qp):
                    for page_num in range(len(doc)):
                        if page_num > 0: qp.newPage()
                        page = doc[page_num]
                        dpi = max(qp.resolution(), 150)
                        zoom = dpi / 72.0
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        img = QImage(pix.samples, pix.width, pix.height, pix.stride,
                                     QImage.Format.Format_RGB888)
                        vp = painter.viewport()
                        scaled = img.scaled(vp.width(), vp.height(),
                                            Qt.AspectRatioMode.KeepAspectRatio,
                                            Qt.TransformationMode.SmoothTransformation)
                        painter.drawImage(0, 0, scaled)
                    painter.end()
                doc.close()
                return
        except Exception: pass
    # 4. Last resort — opens in default viewer (cross-platform)
    try:
        if sys.platform == "win32":
            os.startfile(pdf_path, "print")
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", pdf_path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", pdf_path])
    except Exception: pass

def build_invoice_pdf_path(inv_dir,invoice_date,first="",last="",company="",customer_name="",is_estimate=False):
    prefix="estimate" if is_estimate else "invoice"
    customer=display_customer_name(first=first,last=last,company=company,customer_name=customer_name)
    return os.path.join(inv_dir,f"{prefix}_{safe_filename_part(customer)}_{safe_filename_part(invoice_date or datetime.today().strftime('%Y-%m-%d'))}.pdf")

def output_invoice(invoice_id,conn,pdf_path,print_setting=None):
    if not generate_invoice_pdf(invoice_id,conn,pdf_path): return False
    if print_setting and print_setting.get("mode")=="printer":
        printer_name=""; copies=int(print_setting.get("copies",2))
        if _WIN32_PRINT: printer_name=print_setting.get("printer_name")or win32print.GetDefaultPrinter()
        print_pdf(pdf_path,printer_name=printer_name,copies=copies)
    return True

def _best_vehicle_for_invoice(conn,inv):
    def _get(key):
        try: return (inv[key] or "").strip()
        except (IndexError, KeyError): return ""
    vin=_get("vin"); plate=_get("plate")
    customer_id=_get("customer_id"); invoice_id=_get("invoice_id")
    if vin:
        row=conn.execute("SELECT * FROM vehicles WHERE vin=? ORDER BY updated_at DESC LIMIT 1",(vin,)).fetchone()
        if row: return row
    if plate:
        row=conn.execute("SELECT * FROM vehicles WHERE plate=? ORDER BY updated_at DESC LIMIT 1",(plate,)).fetchone()
        if row: return row
    if invoice_id:
        line=conn.execute("SELECT vin,plate,year,make,model FROM invoice_lines WHERE invoice_id=? AND (vin!='' OR plate!='' OR year!='' OR make!='') LIMIT 1",(invoice_id,)).fetchone()
        if line: return line
    if customer_id:
        rows=conn.execute("SELECT * FROM vehicles WHERE customer_id=? ORDER BY updated_at DESC",(customer_id,)).fetchall()
        if len(rows)==1: return rows[0]
    return None

def draw_header(c, biz, title, subtitle=""):
    w,h=LETTER; logo_path=(biz.get("logo_path")or"").strip(); biz_x=36
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(ImageReader(logo_path),36,h-100,width=82,height=72,preserveAspectRatio=True,mask="auto"); biz_x=126
        except Exception: pass
    c.setFillColor(colors.black); c.setFont("Helvetica-Bold",14)
    c.drawString(biz_x,h-28,biz.get("name","BLUE SKY SMOG")); c.setFont("Helvetica",9); info_y=h-43
    for line in filter(None,[biz.get("address_line1",""),biz.get("address_line2",""),
                              f"Phone: {biz.get('phone','')}" if biz.get("phone") else "",
                              f"ARD #: {biz.get('ard','')}" if biz.get("ard") else ""]):
        c.drawString(biz_x,info_y,str(line)); info_y-=12
    c.setFont("Helvetica-Bold",15); c.drawRightString(w-36,h-28,title)
    if subtitle: c.setFont("Helvetica",9); c.drawRightString(w-36,h-44,subtitle)
    c.setStrokeColor(colors.HexColor("#0097A7")); c.setLineWidth(1.5); c.line(36,h-108,w-36,h-108)
    c.setLineWidth(1); c.setStrokeColor(colors.black)

def generate_invoice_pdf(invoice_id, conn, out_path):
    inv=conn.execute("SELECT * FROM invoices WHERE invoice_id=?",(invoice_id,)).fetchone()
    if not inv: return False
    lines=conn.execute("SELECT * FROM invoice_lines WHERE invoice_id=? ORDER BY id",(invoice_id,)).fetchall()
    biz=get_business_settings(conn)
    cust=conn.execute("SELECT * FROM customers WHERE customer_id=?",(inv["customer_id"],)).fetchone()
    os.makedirs(os.path.dirname(out_path),exist_ok=True)
    c=canvas.Canvas(out_path,pagesize=LETTER); w,h=LETTER
    is_estimate=bool(inv["is_estimate"]); title="ESTIMATE" if is_estimate else "INVOICE"
    inv_num=inv["invoice_number"]or"PENDING"
    FS_BODY=9; FS_LABEL=9; FS_BOLD=10; FS_TOTAL=11; FS_GTOTAL=13; FS_NOTICE=5.5; LINE_H=13; BARCODE_RESERVE=52
    def page_header(page_title=None):
        draw_header(c,biz,page_title or title,"")
        c.setFont("Helvetica",FS_LABEL)
        c.drawRightString(w-170,h-58,f"{title.title()} #: {inv_num}")
        c.drawRightString(w-40,h-58,f"Date: {inv['invoice_date']}")
    page_header(); y=h-125
    c.setFont("Helvetica-Bold",FS_BOLD+2); c.drawString(40,y,"Bill To:"); x=160
    c.setFont("Helvetica-Bold",FS_BOLD)
    company=(inv["company_name"]or(cust["company_name"] if cust else "")or"").strip()
    person_name=f"{inv['first_name']} {inv['last_name']}".strip()or(inv["customer_name"]or"")
    address=cust["address"] if cust else""; city_line=" ".join(filter(None,[cust["city"],cust["state"],cust["zip"]])) if cust else""
    phone=cust["phone"] if cust else""; email=cust["email"] if cust else""
    if company: c.drawString(x,y,f"Company: {company}"); y-=LINE_H
    if person_name: c.setFont("Helvetica",FS_BODY); c.drawString(x,y,person_name); y-=LINE_H
    c.setFont("Helvetica",FS_BODY)
    if address: c.drawString(x,y,f"Address: {address}"); y-=LINE_H
    if city_line: c.drawString(x,y,city_line); y-=LINE_H
    if phone: c.drawString(x,y,f"Phone: {phone}"); y-=LINE_H
    if email: c.drawString(x,y,f"Email: {email}"); y-=LINE_H
    y-=8; c.line(40,y,w-40,y); y-=16
    c.setFont("Helvetica-Bold",FS_BOLD)
    c.drawString(70,y,"Vehicle / Service Performed"); c.drawRightString(w-70,y,"Amount")
    y-=10; c.line(70,y,w-70,y); y-=LINE_H; subtotal=0.0
    def ensure_space(min_y=BARCODE_RESERVE+10):
        nonlocal y
        if y>=min_y: return
        c.showPage(); page_header(title+" (cont.)"); y=h-110
        c.setFont("Helvetica-Bold",FS_BOLD); c.drawString(70,y,"Vehicle / Service Performed")
        c.drawRightString(w-70,y,"Amount"); y-=10; c.line(70,y,w-70,y); y-=LINE_H
    hdr_vin=(inv["vin"]or"").strip(); hdr_plate=(inv["plate"]or"").strip()
    hdr_year=(inv["year"]or"").strip(); hdr_make=(inv["make"]or"").strip(); hdr_model=(inv["model"]or"").strip()
    for line in lines:
        ensure_space(); svc_name=(line["service"]or"").strip()
        is_fee_line=svc_name in ("Credit Card Fee","Card Fee","CC Fee")
        if is_fee_line: vin_l=plate_l=odo_l=year_l=make_l=model_l=""
        else:
            vin_l=(line["vin"]or"").strip()or hdr_vin; plate_l=(line["plate"]or"").strip()or hdr_plate
            odo_l=(line["odometer"]or"").strip(); year_l=(line["year"]or"").strip()or hdr_year
            make_l=(line["make"]or"").strip()or hdr_make; model_l=(line["model"]or"").strip()or hdr_model
        result=(line["result"]or"").strip(); cert=(line["cert"]or"").strip()
        disc=float(line["discount"]or 0); price=float(line["price"]or 0); subtotal+=price
        info_parts=[]
        if vin_l: info_parts.append(f"VIN: {vin_l}")
        if plate_l: info_parts.append(f"Plate: {plate_l}")
        if odo_l: info_parts.append(f"Odometer: {odo_l}")
        if info_parts: c.setFont("Helvetica-Bold",FS_BOLD); c.drawString(70,y,"    ".join(info_parts)); y-=LINE_H
        vehicle_line="   ".join(filter(None,[f"Year: {year_l}" if year_l else"",f"Make: {make_l}" if make_l else"",f"Model: {model_l}" if model_l else""]))
        if vehicle_line: c.setFont("Helvetica",FS_BODY); c.drawString(70,y,vehicle_line); y-=LINE_H
        c.setFont("Helvetica",FS_BODY)
        service_text=svc_name or"Service Performed"
        if result and svc_name=="Smog Test": service_text+=f" ({result})"
        if cert: service_text+=f"  Cert: {cert}"
        c.drawString(90,y,f"Service: {service_text}"); c.drawRightString(w-70,y,f"${price:,.2f}"); y-=LINE_H
        if disc>0: c.drawString(90,y,"Discount"); c.drawRightString(w-70,y,f"-${disc:,.2f}"); y-=LINE_H
        y-=6
    if not lines:
        vrow=_best_vehicle_for_invoice(conn,inv)
        vin_f=(inv["vin"]or"").strip()or(vrow["vin"]if vrow else"")
        plate_f=(inv["plate"]or"").strip()or(vrow["plate"]if vrow else"")
        year_f=(inv["year"]or"").strip()or(vrow["year"]if vrow else"")
        make_f=(inv["make"]or"").strip()or(vrow["make"]if vrow else"")
        model_f=(inv["model"]or"").strip()or(vrow["model"]if vrow else"")
        info_parts=[]
        if vin_f: info_parts.append(f"VIN: {vin_f}")
        if plate_f: info_parts.append(f"Plate: {plate_f}")
        if info_parts: c.setFont("Helvetica-Bold",FS_BOLD); c.drawString(70,y,"    ".join(info_parts)); y-=LINE_H
        vehicle_line="   ".join(filter(None,[f"Year: {year_f}" if year_f else"",f"Make: {make_f}" if make_f else"",f"Model: {model_f}" if model_f else""]))
        if vehicle_line: c.setFont("Helvetica",FS_BODY); c.drawString(70,y,vehicle_line); y-=LINE_H
        notes_text=(inv["notes"]or"").strip(); service_text=notes_text if notes_text else"Smog Inspection"
        subtotal=float(inv["amount_cents"]or 0)/100.0
        c.setFont("Helvetica",FS_BODY); c.drawString(90,y,f"Service: {service_text}")
        c.drawRightString(w-70,y,f"${subtotal:,.2f}"); y-=LINE_H+4
    c.line(40,y,w-40,y); y-=18
    inv_total=float(inv["amount_cents"]or 0)/100.0; total_due=inv_total if inv_total>0 else subtotal
    c.setFont("Helvetica-Bold",FS_TOTAL); c.drawRightString(w-40,y,f"Subtotal: ${total_due:,.2f}"); y-=16
    c.setFont("Helvetica-Bold",FS_GTOTAL); c.drawRightString(w-40,y,f"Grand Total: ${total_due:,.2f}"); y-=24
    if not is_estimate and inv["payment_method"]:
        c.setFont("Helvetica",FS_BODY); c.drawString(70,y,f"Payment Method: {inv['payment_method']}"); y-=LINE_H
    if inv["notes"]:
        y-=10; c.setFont("Helvetica-Bold",FS_LABEL); c.drawString(70,y,"Notes:"); y-=LINE_H
        c.setFont("Helvetica",FS_BODY)
        for line_t in textwrap.wrap(inv["notes"],95): c.drawString(90,y,line_t); y-=LINE_H-1
    notice_raw=biz.get("invoice_notice",DEFAULT_BUSINESS.get("invoice_notice",""))
    biz_name=biz.get("name","").strip(); notice_text=notice_raw.replace("{business_name}",biz_name)
    if notice_text.strip():
        y-=6; c.setFont("Helvetica",FS_NOTICE)
        for notice_line in notice_text.split("\n"):
            notice_line=notice_line.strip()
            if not notice_line: y-=3; continue
            for wrapped in textwrap.wrap(notice_line,150)or[""]:
                ensure_space(BARCODE_RESERVE+8); c.drawString(40,y,wrapped); y-=7
        y-=4
    ensure_space(BARCODE_RESERVE+70); y-=50
    c.setFont("Helvetica",FS_BODY); c.drawString(70,y,"Customer Signature:"); c.line(190,y,w-40,y)
    y-=14; c.setFont("Helvetica",FS_NOTICE); c.drawString(190,y,"X")
    try:
        PAGE_WIDTH=LETTER[0]; LEFT_MARGIN=180; RIGHT_MARGIN=180; BOTTOM_Y=12; BAR_HEIGHT=18
        TARGET_WIDTH=PAGE_WIDTH-LEFT_MARGIN-RIGHT_MARGIN
        vin_b=(inv["vin"]or hdr_vin or"").strip(); plate_b=(inv["plate"]or hdr_plate or"").strip()
        year_b=(inv["year"]or hdr_year or"").strip(); make_b=(inv["make"]or hdr_make or"").strip()
        model_b=(inv["model"]or hdr_model or"").strip()
        # DMV-format barcode: %0 + VIN(17) + PLATE — matches what handheld scanners expect
        if vin_b:
            barcode_value = f"%0{vin_b}{plate_b}"
        elif plate_b:
            barcode_value = plate_b
        else:
            barcode_value = f"INV{inv_num}"
        # Compute barWidth so the barcode fills TARGET_WIDTH.
        # Code128B: start(11) + data(11n) + check(11) + stop(13) + quiet_zones(10+10) = 55+11n modules
        _n_mod = 55 + 11 * len(barcode_value)
        _bw    = max(TARGET_WIDTH / _n_mod, 0.5)
        barcode=code128.Code128(barcode_value,barHeight=BAR_HEIGHT,barWidth=_bw)
        barcode.drawOn(c,LEFT_MARGIN,BOTTOM_Y)
    except Exception: pass
    c.save()
    conn.execute("UPDATE invoices SET pdf_path=? WHERE invoice_id=?",(out_path,invoice_id)); conn.commit()
    return True

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  STYLESHEET
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

APP_STYLE = """
QWidget { font-family: 'Segoe UI', Arial, sans-serif; font-size: 10pt; font-weight: bold; color: #111827; }
QMainWindow, QDialog { background: #F5F7FB; }
QScrollArea > QWidget > QWidget { background: #F5F7FB; }

/* Buttons */
QPushButton {
    border: none; border-radius: 5px; padding: 6px 14px;
    font-weight: bold; font-size: 10pt;
}
QPushButton#primary   { background:#005B99; color:white; }
QPushButton#primary:hover   { background:#0073C2; }
QPushButton#success   { background:#2E7D32; color:white; }
QPushButton#success:hover   { background:#1B5E20; }
QPushButton#danger    { background:#DC2626; color:white; }
QPushButton#danger:hover    { background:#991B1B; }
QPushButton#secondary { background:#455A64; color:white; }
QPushButton#secondary:hover { background:#263238; }
QPushButton#accent    { background:#1976D2; color:white; }
QPushButton#accent:hover    { background:#1565C0; }

/* Inputs */
QLineEdit, QComboBox, QTextEdit, QPlainTextEdit, QSpinBox {
    background: white; border: 1px solid #D1D5DB; border-radius: 4px;
    padding: 4px 6px; selection-background-color: #005B99;
}
QLineEdit:focus, QComboBox:focus, QTextEdit:focus { border-color: #005B99; }
QComboBox::drop-down { border: none; padding-right: 4px; }
QComboBox QAbstractItemView {
    background: white; color: #111827;
    selection-background-color: #BFDBFE; selection-color: #111827;
    border: 1px solid #D1D5DB;
}
QComboBox QAbstractItemView::item { color: #111827; padding: 4px 6px; }

/* Table */
QTableWidget {
    background: white; gridline-color: #E5E7EB;
    border: 1px solid #D1D5DB; border-radius: 4px;
    alternate-background-color: #EFF6FF;
    selection-background-color: #BFDBFE;
    selection-color: #111827;
}
QTableWidget::item { padding: 2px 6px; }
QHeaderView::section {
    background: #005B99; color: white; font-weight: bold;
    padding: 6px 8px; border: none; border-right: 1px solid #0073C2;
}

/* GroupBox */
QGroupBox {
    font-weight: bold; border: 1px solid #D1D5DB;
    border-radius: 6px; margin-top: 8px; padding-top: 8px;
}
QGroupBox::title {
    subcontrol-origin: margin; left: 10px;
    background: #005B99; color: white;
    padding: 2px 8px; border-radius: 3px;
}

/* Tab */
QTabWidget::pane { border: 1px solid #D1D5DB; border-radius: 4px; }
QTabBar::tab {
    background: #E5E7EB; padding: 6px 16px; border-top-left-radius: 4px;
    border-top-right-radius: 4px; margin-right: 2px;
}
QTabBar::tab:selected { background: #005B99; color: white; }

/* Radio buttons - show a clean ring; filled border = checked */
QRadioButton { spacing: 8px; }
QRadioButton::indicator { width: 16px; height: 16px; border-radius: 8px; }
QRadioButton::indicator:unchecked { border: 2px solid #9CA3AF; background: white; }
QRadioButton::indicator:checked   { border: 5px solid #005B99; background: white; }

/* Checkboxes */
QCheckBox { spacing: 8px; }
QCheckBox::indicator { width: 16px; height: 16px; border-radius: 3px; }
QCheckBox::indicator:unchecked { border: 2px solid #9CA3AF; background: white; }
QCheckBox::indicator:checked   { border: 2px solid #005B99; background: #005B99; }

/* ScrollBar */
QScrollBar:vertical { width: 10px; background: #F1F5F9; }
QScrollBar::handle:vertical { background: #CBD5E1; border-radius: 5px; min-height: 20px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

/* Context Menu */
QMenu {
    background: white; border: 1px solid #D1D5DB;
    border-radius: 4px; padding: 4px 0;
    color: #111827;
}
QMenu::item { padding: 6px 24px 6px 16px; }
QMenu::item:selected { background: #DBEAFE; color: #1D4ED8; }
QMenu::separator { height: 1px; background: #E5E7EB; margin: 2px 8px; }

/* Autocomplete / completer dropdown */
QAbstractItemView {
    background: white; border: 1px solid #D1D5DB; color: #111827;
}
QAbstractItemView::item:selected, QAbstractItemView::item:hover {
    background: #DBEAFE; color: #111827;
}
QComboBox QAbstractItemView::item:selected {
    background: #DBEAFE; color: #111827;
}
QDateEdit { color: #111827; background: white; selection-background-color: #DBEAFE; selection-color: #111827; }
"""

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  WORKER THREADS
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

class VinWorker(QThread):
    done = pyqtSignal(str, str, str)
    def __init__(self, vin): super().__init__(); self.vin = vin
    def run(self):
        yr, mk, md = api_decode_vin(self.vin)
        self.done.emit(yr, mk, md)

class ZipWorker(QThread):
    done = pyqtSignal(str, str)
    def __init__(self, zip_code): super().__init__(); self.zip_code = zip_code
    def run(self):
        try:
            url = f"https://api.zippopotam.us/us/{self.zip_code}"
            with urllib.request.urlopen(url, timeout=3) as resp:
                data = json.loads(resp.read())
            city  = data["places"][0]["place name"].upper()
            state = data["places"][0]["state abbreviation"].upper()
            self.done.emit(city, state)
        except Exception: self.done.emit("", "")

class SyncWorker(QThread):
    status_changed = pyqtSignal(str)
    data_changed   = pyqtSignal()
    suspended      = pyqtSignal(str)
    def run(self):
        while True:
            try:
                SYNC._flush()
                SYNC._pull()
            except Exception: pass
            time.sleep(SYNC_INTERVAL)

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  HELPERS
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

def btn(text, style="primary", parent=None):
    b = QPushButton(text, parent)
    b.setObjectName(style)
    b.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
    return b

def make_header(biz_name, show_back=False, back_cb=None, sync_label=None):
    """Returns a blue header QWidget."""
    w = QWidget(); w.setFixedHeight(44)
    w.setStyleSheet(f"background:{PRIMARY};")
    h = QHBoxLayout(w); h.setContentsMargins(8,4,8,4); h.setSpacing(8)
    if show_back and back_cb:
        b = QPushButton("< Back"); b.setObjectName("secondary")
        b.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        b.clicked.connect(back_cb); h.addWidget(b)
    name_lbl = QLabel(biz_name.upper())
    name_lbl.setStyleSheet("color:white; font-size:13pt; font-weight:bold;")
    h.addWidget(name_lbl); h.addStretch()
    if sync_label:
        h.addWidget(sync_label)
    return w

def _upper_entry(le: QLineEdit):
    """Force uppercase on a QLineEdit."""
    le.textChanged.connect(lambda t: le.setText(t.upper()) if t != t.upper() else None)

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  ADMIN COMPANY DETAIL DIALOG  (master only)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

class AdminCompanyDialog(QDialog):
    """Detail view for a single company in the master admin dashboard."""
    def __init__(self, username, co_info, monthly, sub_status, parent=None):
        super().__init__(parent)
        self.username   = username
        self.co_info    = co_info
        self.monthly    = monthly
        self.sub_status = sub_status
        self.setWindowTitle(co_info.get("company_name", username))
        self.setMinimumWidth(720); self.setMinimumHeight(640)
        self._build()

    def _master_api(self, method, path, **kwargs):
        creds   = load_creds()
        headers = {"x-username": creds.get("username",""), "x-password": creds.get("password","")}
        r = getattr(requests, method)(f"{API_BASE}{path}", headers=headers, timeout=15, **kwargs)
        r.raise_for_status()
        return r.json()

    def _build(self):
        outer = QVBoxLayout(self)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        content = QWidget(); cl = QVBoxLayout(content)
        cl.setContentsMargins(16,16,16,16); cl.setSpacing(12)
        scroll.setWidget(content); outer.addWidget(scroll)

        # Title
        co_name = self.co_info.get("company_name", self.username)
        tl = QLabel(co_name); tl.setStyleSheet(f"color:{PRIMARY}; font-size:14pt; font-weight:bold;")
        cl.addWidget(tl)

        # Compute last-activity / inactivity
        last_act_raw = (self.co_info.get("last_activity","") or "")
        last_date    = None
        try:
            if last_act_raw:
                last_date = datetime.fromisoformat(last_act_raw.replace("Z",""))
        except Exception:
            pass
        days_since  = (datetime.now() - last_date).days if last_date else None
        inactive30  = days_since is not None and days_since >= 30
        if last_date is None:
            last_act_str = "No activity recorded"
        elif days_since == 0:
            last_act_str = "Today"
        elif days_since == 1:
            last_act_str = "Yesterday"
        else:
            last_act_str = f"{days_since} days ago"

        # Inactivity warning banner
        if inactive30:
            warn_box = QFrame()
            warn_box.setStyleSheet(
                "background:#FFF3E0; border:1px solid #FFA500; border-radius:6px;")
            wl = QHBoxLayout(warn_box); wl.setContentsMargins(10,8,10,8)
            warn_lbl = QLabel(f"⚠  This account has been inactive for {days_since} days.")
            warn_lbl.setStyleSheet("color:#CC6600; font-weight:bold; font-size:10pt;")
            wl.addWidget(warn_lbl)
            cl.addWidget(warn_box)

        # â"€â"€ Account Info â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        sus       = self.co_info.get("is_suspended", False)
        plan      = self.sub_status.get("plan","?")
        is_exempt = (plan == "owner")
        info_grp  = QGroupBox("Account Info")
        ig        = QGridLayout(info_grp)
        info_rows = [
            ("Username",        f"@{self.co_info.get('username','')}",  None),
            ("Company",         co_name,                                 None),
            ("Member Since",    (self.co_info.get("created_at","") or "")[:10], None),
            ("Last Activity",   last_act_str,                            "#CC6600" if inactive30 else None),
            ("Total Invoices",  str(self.co_info.get("invoice_count",0)), None),
            ("Status",          "Active" if not sus else "Suspended",    None),
            ("Billing",         "Exempt - No Billing" if is_exempt else f"Plan: {plan}", None),
        ]
        for i,(k,v,color) in enumerate(info_rows):
            ig.addWidget(QLabel(k+":"), i, 0)
            vl = QLabel(v)
            style = "font-weight:bold;"
            if color:
                style += f" color:{color};"
            vl.setStyleSheet(style)
            ig.addWidget(vl, i, 1)
        cl.addWidget(info_grp)

        # â"€â"€ Billing Summary â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        if self.monthly:
            now_month      = datetime.now().strftime("%Y-%m")
            this_month_row = next((m for m in self.monthly if m["month"]==now_month), None)
            all_cents      = sum(m.get("total_cents",0) for m in self.monthly)
            all_inv        = sum(m.get("invoice_count",0) for m in self.monthly)
            this_cents     = this_month_row["total_cents"]    if this_month_row else 0
            this_inv_cnt   = this_month_row["invoice_count"]  if this_month_row else 0
            per_inv_fee    = round(this_inv_cnt * 0.15, 2)
            due            = max(per_inv_fee, 40.00) if this_inv_cnt > 0 else 0.0

            bill_grp = QGroupBox("Billing Summary"); bill_lay = QVBoxLayout(bill_grp)
            cards_h  = QHBoxLayout()
            for label, amount, count, color in [
                ("This Month", this_cents/100, this_inv_cnt, PRIMARY),
                ("All Time",   all_cents/100,  all_inv,      GREEN),
            ]:
                card = QFrame(); card.setFrameShape(QFrame.Shape.StyledPanel)
                card.setStyleSheet("background:#2c3e5a; border-radius:6px;")
                c2 = QVBoxLayout(card); c2.setContentsMargins(12,8,12,8)
                lbl_title = QLabel(label); lbl_title.setStyleSheet("color:white; font-size:10pt;")
                c2.addWidget(lbl_title)
                al = QLabel(f"${amount:,.2f}"); al.setStyleSheet(f"color:{color}; font-size:16pt; font-weight:bold;")
                lbl_count = QLabel(f"{count} invoices"); lbl_count.setStyleSheet("color:#b0c4de; font-size:9pt;")
                c2.addWidget(al); c2.addWidget(lbl_count)
                cards_h.addWidget(card)
            bill_lay.addLayout(cards_h)

            if not is_exempt and this_inv_cnt > 0:
                due_box = QFrame()
                due_box.setStyleSheet("background:#3a2800; border-radius:6px; border:1px solid #FFA500;")
                dl = QVBoxLayout(due_box); dl.setContentsMargins(12,8,12,8)
                due_title = QLabel("Amount Due This Month"); due_title.setStyleSheet("color:white; font-size:10pt;")
                dl.addWidget(due_title)
                da = QLabel(f"${due:,.2f}"); da.setStyleSheet("color:#FFA500; font-size:20pt; font-weight:bold;")
                dl.addWidget(da)
                nl = QLabel(f"Flat rate - {this_inv_cnt} x $0.15 = ${per_inv_fee:.2f} (min $40.00)")
                nl.setStyleSheet("color:#FFA500; font-size:9pt;"); dl.addWidget(nl)
                bill_lay.addWidget(due_box)
            cl.addWidget(bill_grp)

            # Monthly breakdown table
            mb_grp = QGroupBox("Monthly Breakdown"); mb_lay = QVBoxLayout(mb_grp)
            mb_tbl = QTableWidget(len(self.monthly), 3)
            mb_tbl.setHorizontalHeaderLabels(["Month","Invoices","Revenue"])
            mb_tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            mb_tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            mb_tbl.verticalHeader().setVisible(False); mb_tbl.setMaximumHeight(200)
            for i, m in enumerate(self.monthly):
                is_cur = (m["month"] == now_month)
                for col, val in enumerate([m["month"], str(m.get("invoice_count",0)),
                                           f"${m.get('total_cents',0)/100:,.2f}"]):
                    item = QTableWidgetItem(val)
                    if is_cur:
                        item.setForeground(QColor(PRIMARY))
                        item.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
                    mb_tbl.setItem(i, col, item)
            mb_lay.addWidget(mb_tbl); cl.addWidget(mb_grp)

        # â"€â"€ Subscription Status â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        sub_grp = QGroupBox("Subscription Status"); sub_lay = QVBoxLayout(sub_grp)
        can = self.sub_status.get("can_create", False)
        sh  = QHBoxLayout()
        badge = QLabel(f"  {plan.upper()}  ")
        badge.setStyleSheet(f"background:{GREEN}; color:white; border-radius:4px; padding:2px 6px; font-weight:bold;")
        sh.addWidget(badge); sh.addWidget(QLabel(f"  Plan: {plan}"))
        can_lbl = QLabel("Can create" if can else "Read-only")
        can_lbl.setStyleSheet(f"color:{GREEN if can else RED};")
        sh.addWidget(can_lbl); sh.addStretch(); sub_lay.addLayout(sh)
        sub_lay.addWidget(QLabel("Override Plan:"))
        ph = QHBoxLayout()
        for pn in ("trial","grace","locked","monthly","per_invoice"):
            pb = btn(pn, "secondary"); pb.setFixedHeight(28)
            pb.clicked.connect(lambda chk=False, p2=pn: self._override_plan(p2))
            ph.addWidget(pb)
        ph.addStretch(); sub_lay.addLayout(ph); cl.addWidget(sub_grp)

        # â"€â"€ Admin Notes â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        notes_grp = QGroupBox("Admin Notes"); notes_lay = QVBoxLayout(notes_grp)
        self._notes_e = QTextEdit(self.co_info.get("admin_notes",""))
        self._notes_e.setMaximumHeight(100); notes_lay.addWidget(self._notes_e)
        sn_b = btn("Save Notes","primary"); sn_b.clicked.connect(self._save_notes)
        notes_lay.addWidget(sn_b); cl.addWidget(notes_grp)

        # â"€â"€ Action Buttons â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        acts = QHBoxLayout()
        if sus:
            ub = btn("Unsuspend","success"); ub.clicked.connect(self._unsuspend); acts.addWidget(ub)
        else:
            sb2 = btn("Suspend","danger"); sb2.clicked.connect(self._suspend); acts.addWidget(sb2)
        acts.addStretch()
        del_b = btn("Delete This Company","danger"); del_b.clicked.connect(self._delete_company)
        acts.addWidget(del_b); cl.addLayout(acts)
        cl.addStretch()

        # Close button
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        bb.rejected.connect(self.reject); outer.addWidget(bb)

    def _override_plan(self, plan):
        if QMessageBox.question(self,"Override Plan",f"Set plan to '{plan}' for @{self.username}?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        try:
            self._master_api("post", f"/v1/master/company/{self.username}/subscription",
                             json={"plan": plan, "reset_invoice_count": False})
            QMessageBox.information(self,"Done",f"Plan set to '{plan}'.")
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")

    def _save_notes(self):
        try:
            self._master_api("post", f"/v1/master/company/{self.username}/notes",
                             json={"notes": self._notes_e.toPlainText()})
            QMessageBox.information(self,"Saved","Notes saved.")
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")

    def _suspend(self):
        if QMessageBox.question(self,"Suspend",
                f"Suspend @{self.username}? They will be locked out immediately.",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        try:
            self._master_api("post", f"/v1/master/company/{self.username}/suspend")
            QMessageBox.information(self,"Done",f"@{self.username} has been suspended.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")

    def _unsuspend(self):
        try:
            self._master_api("post", f"/v1/master/company/{self.username}/unsuspend")
            QMessageBox.information(self,"Done",f"@{self.username} has been unsuspended.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")

    def _delete_company(self):
        name = self.co_info.get("company_name", self.username)
        if QMessageBox.question(self,"DELETE COMPANY",
                f"PERMANENTLY DELETE '{name}' (@{self.username})?\nThis cannot be undone.",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        try:
            self._master_api("delete", f"/v1/master/company/{self.username}")
            QMessageBox.information(self,"Deleted",f"Company @{self.username} has been deleted.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  LOGIN DIALOG
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Blue Sky Smog - Sign In")
        self.setFixedWidth(400)
        self.setModal(True)
        self._token = None
        self._build()

    def _build(self):
        lay = QVBoxLayout(self); lay.setSpacing(12); lay.setContentsMargins(32,32,32,32)
        title = QLabel("BLUE SKY SMOG"); title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(f"color:{PRIMARY}; font-size:18pt; font-weight:bold;")
        lay.addWidget(title)
        sub = QLabel("Sign in to continue"); sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet("color:#6B7280;"); lay.addWidget(sub)

        form = QFormLayout(); form.setSpacing(8)
        saved = load_creds()
        self._user_e = QLineEdit(saved.get("username",""))
        self._pass_e = QLineEdit(saved.get("password",""))
        self._pass_e.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("Username:", self._user_e)
        form.addRow("Password:", self._pass_e)
        lay.addLayout(form)

        self._err_lbl = QLabel(""); self._err_lbl.setStyleSheet("color:red;")
        self._err_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._err_lbl.setWordWrap(True); lay.addWidget(self._err_lbl)

        sign_btn = btn("Sign In", "primary"); sign_btn.clicked.connect(self._do_login)
        lay.addWidget(sign_btn)
        self._pass_e.returnPressed.connect(self._do_login)

        reg_btn = btn("Create Account", "secondary"); reg_btn.clicked.connect(self._do_register)
        lay.addWidget(reg_btn)

        sub_btn = btn("Subscribe - $40/month", "primary")
        sub_btn.clicked.connect(self._do_subscribe)
        lay.addWidget(sub_btn)

        portal_btn = btn("Manage Subscription", "secondary")
        portal_btn.clicked.connect(self._do_portal)
        lay.addWidget(portal_btn)

    def _do_login(self):
        u = self._user_e.text().strip(); p = self._pass_e.text().strip()
        if not u or not p: self._err_lbl.setText("Enter username and password."); return
        self._err_lbl.setText("Signing in...")
        QApplication.processEvents()
        try:
            if not requests: raise ValueError("No network library (requests not installed).")
            token, company_id, company_name = api_login(u, p)
            save_creds({"username":u,"password":p,"token":token,
                        "company_id":company_id,"company_name":company_name})
            self._token = token; self.accept()
        except Exception as e:
            self._err_lbl.setText(str(e))

    def _do_register(self):
        dlg = QDialog(self); dlg.setWindowTitle("Create Account"); dlg.setFixedWidth(420); dlg.setModal(True)
        lay = QVBoxLayout(dlg); lay.setSpacing(10); lay.setContentsMargins(28,28,28,28)

        title = QLabel("Create a New Account"); title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(f"color:{PRIMARY}; font-size:14pt; font-weight:bold;"); lay.addWidget(title)

        info = QLabel("A free trial is included. A $40/month subscription is required to continue after your trial.")
        info.setWordWrap(True); info.setStyleSheet("color:#6B7280; font-size:10pt;")
        info.setAlignment(Qt.AlignmentFlag.AlignCenter); lay.addWidget(info)

        form = QFormLayout(); form.setSpacing(8)
        co_e   = QLineEdit(); co_e.setPlaceholderText("Your business name")
        user_e = QLineEdit(); user_e.setPlaceholderText("Lowercase letters and numbers")
        pass_e = QLineEdit(); pass_e.setEchoMode(QLineEdit.EchoMode.Password); pass_e.setPlaceholderText("At least 6 characters")
        pass2_e= QLineEdit(); pass2_e.setEchoMode(QLineEdit.EchoMode.Password); pass2_e.setPlaceholderText("Repeat password")
        form.addRow("Business Name:", co_e)
        form.addRow("Username:",      user_e)
        form.addRow("Password:",      pass_e)
        form.addRow("Confirm:",       pass2_e)
        lay.addLayout(form)

        err_lbl = QLabel(""); err_lbl.setStyleSheet("color:red;"); err_lbl.setWordWrap(True)
        err_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); lay.addWidget(err_lbl)

        def _submit():
            co   = co_e.text().strip()
            u    = user_e.text().strip().lower()
            p    = pass_e.text()
            p2   = pass2_e.text()
            if not co:   err_lbl.setText("Business name is required."); return
            if len(u)<3: err_lbl.setText("Username must be at least 3 characters."); return
            if len(p)<6: err_lbl.setText("Password must be at least 6 characters."); return
            if p != p2:  err_lbl.setText("Passwords do not match."); return
            err_lbl.setText("Creating account...")
            QApplication.processEvents()
            try:
                if not requests: raise ValueError("No network library (requests not installed).")
                r = requests.post(f"{API_BASE}/v1/auth/register",
                    json={"username": u, "password": p, "company_name": co},
                    timeout=15)
                r.raise_for_status()
                data = r.json()
                if not data.get("success"): raise ValueError("Registration failed.")
                token = data.get("token","")
                save_creds({"username": u, "password": p, "token": token,
                            "company_id": data.get("company_id",""),
                            "company_name": data.get("company_name", co)})
                self._token = token
                dlg.accept()
                self.accept()
            except Exception as ex:
                msg = str(ex)
                if "409" in msg or "already taken" in msg:
                    err_lbl.setText("That username is already taken. Choose another.")
                elif "400" in msg:
                    err_lbl.setText("Invalid username or password requirements.")
                else:
                    err_lbl.setText(f"Error: {msg}")

        create_btn = btn("Create Account", "primary"); create_btn.clicked.connect(_submit)
        pass2_e.returnPressed.connect(_submit)
        lay.addWidget(create_btn)
        cancel_btn = btn("Cancel", "secondary"); cancel_btn.clicked.connect(dlg.reject)
        lay.addWidget(cancel_btn)
        dlg.exec()

    def _do_subscribe(self):
        u = self._user_e.text().strip(); p = self._pass_e.text().strip()
        if not u or not p:
            self._err_lbl.setText("Enter your username and password first, then click Subscribe.")
            return
        self._err_lbl.setText("Opening checkout...")
        QApplication.processEvents()
        try:
            if not requests:
                raise ValueError("Network library (requests) not installed.")
            r = requests.post(
                f"{API_BASE}/v1/subscription/checkout",
                json={"plan": "monthly"},
                headers={"x-username": u, "x-password": p},
                timeout=15,
            )
            r.raise_for_status()
            url = r.json().get("checkout_url", "")
            if url:
                import webbrowser
                webbrowser.open(url)
                self._err_lbl.setText("Checkout page opened in your browser.")
            else:
                self._err_lbl.setText("No checkout URL returned from server.")
        except Exception as e:
            self._err_lbl.setText(f"Subscribe error: {e}")

    def _do_portal(self):
        u = self._user_e.text().strip(); p = self._pass_e.text().strip()
        if not u or not p:
            self._err_lbl.setText("Enter your username and password first, then click Manage Subscription.")
            return
        self._err_lbl.setText("Opening billing portal...")
        QApplication.processEvents()
        try:
            if not requests:
                raise ValueError("Network library (requests) not installed.")
            r = requests.post(
                f"{API_BASE}/v1/subscription/portal",
                headers={"x-username": u, "x-password": p},
                timeout=15,
            )
            r.raise_for_status()
            url = r.json().get("portal_url", "")
            if url:
                import webbrowser
                webbrowser.open(url)
                self._err_lbl.setText("Billing portal opened in your browser.")
            else:
                self._err_lbl.setText("No portal URL returned from server.")
        except Exception as e:
            self._err_lbl.setText(f"Manage Subscription error: {e}")

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  PDF VIEWER DIALOG
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

class PdfViewerDialog(QDialog):
    def __init__(self, pdf_path, ps, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path; self.ps = ps
        self.setWindowTitle(f"Invoice - {os.path.basename(pdf_path)}")
        self.resize(700, 900); self._build()

    def _build(self):
        lay = QVBoxLayout(self); lay.setContentsMargins(8,8,8,8); lay.setSpacing(6)
        tb = QHBoxLayout()
        print_btn = btn("Print", "primary"); print_btn.clicked.connect(self._print)
        open_btn  = btn("Open in System Viewer", "secondary")
        open_btn.clicked.connect(lambda: os.startfile(self.pdf_path))
        tb.addWidget(print_btn); tb.addWidget(open_btn); tb.addStretch()
        lay.addLayout(tb)
        self._scroll = QScrollArea(); self._scroll.setWidgetResizable(True)
        self._content = QWidget(); self._vlay = QVBoxLayout(self._content)
        self._vlay.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self._scroll.setWidget(self._content); lay.addWidget(self._scroll)
        QTimer.singleShot(100, self._render)

    def _render(self):
        if not _FITZ_OK:
            lbl = QLabel("PDF viewer requires PyMuPDF.\n\nRun: pip install pymupdf")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); self._vlay.addWidget(lbl); return
        try:
            doc = fitz.open(self.pdf_path)
            for page_num in range(len(doc)):
                page = doc[page_num]
                zoom = min(2.0, (self._scroll.width() - 30) / page.rect.width)
                zoom = max(1.0, zoom)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img = QImage(pix.samples, pix.width, pix.height,
                             pix.stride, QImage.Format.Format_RGB888)
                pm = QPixmap.fromImage(img)
                lbl = QLabel(); lbl.setPixmap(pm)
                lbl.setStyleSheet("border:1px solid #ccc; margin:8px;")
                self._vlay.addWidget(lbl)
            doc.close()
        except Exception as e:
            lbl = QLabel(f"Error rendering PDF: {e}")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); self._vlay.addWidget(lbl)

    def _print(self):
        copies = int(self.ps.get("copies",2)); printer_name = self.ps.get("printer_name","")
        if _WIN32_PRINT and not printer_name:
            try: printer_name = win32print.GetDefaultPrinter()
            except: pass
        print_pdf(self.pdf_path, printer_name=printer_name, copies=copies, parent_widget=self)

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  MAIN APPLICATION
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

class App(QMainWindow):
    _sync_signal   = pyqtSignal()
    _susp_signal   = pyqtSignal(str)
    _fp_signal     = pyqtSignal(str)   # force-pull result message

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Blue Sky Smog")
        try: self.setWindowIcon(QIcon(_icon_path()))
        except Exception: pass

        self.db = get_db()
        _app_data = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "BlueSkyDesktop")
        self.inv_dir = os.path.join(_app_data, "Invoices")
        os.makedirs(self.inv_dir, exist_ok=True)

        self._editing_id      = None
        self._lines_data      = []
        self._acct_names      = []
        self._acct_index      = 0
        self._sub_status      = {"status":"trial","can_create":True,"warning":""}
        self._is_master       = (load_creds().get("username","") == "bluesky_master")
        self._current_screen  = ""
        self._customer_touched= False
        try: self._zoom = json.loads(get_setting(self.db,"zoom_levels","{}"))
        except: self._zoom = {}

        self._sync_label = QLabel("Checking...")
        self._sync_label.setStyleSheet("color:#90EE90; font-size:9pt;")

        self._stack = QStackedWidget()
        self.setCentralWidget(self._stack)
        self._screens = {}

        self._build_doc_list_screen()
        self._build_estimate_entry_screen()
        self._build_account_setup_screen()
        self._build_reports_screen()
        self._build_settings_screen()
        self._build_customers_screen()
        if self._is_master:
            self._build_admin_screen()

        # Clock
        self._clock_timer = QTimer(self); self._clock_timer.timeout.connect(self._tick_clock)
        self._clock_timer.start(1000)

        # Sync
        self._sync_signal.connect(self._on_sync_change)
        self._susp_signal.connect(self._on_account_suspended)
        self._fp_signal.connect(self._on_force_pull_done)
        SYNC.set_on_change(lambda: self._sync_signal.emit())
        SYNC.set_on_suspended(lambda msg: self._susp_signal.emit(msg))
        SYNC.start()

        # Restore geometry
        try:
            geo = get_setting(self.db,"window_geometry","")
            if geo: self.restoreGeometry(bytes.fromhex(geo))
            else:   self.showMaximized()
        except Exception: self.showMaximized()

        QTimer.singleShot(800, lambda: threading.Thread(target=self._startup_sync, daemon=True).start())
        QTimer.singleShot(2000, self._refresh_sub_status)

        self.show_screen("doc_list")

    def closeEvent(self, event):
        try: set_setting(self.db,"window_geometry", bytes(self.saveGeometry()).hex())
        except Exception: pass
        SYNC.stop(); self.db.close(); event.accept()

    def show_screen(self, name, **kw):
        if name not in self._screens: return
        self._stack.setCurrentWidget(self._screens[name])
        self._current_screen = name
        # Apply saved zoom for this screen
        size = self._zoom.get(name, 10)
        self._screens[name].setStyleSheet(f"font-size: {size}pt;")
        handler = getattr(self, f"_on_show_{name}", None)
        if handler: handler(**kw)

    def _tick_clock(self):
        now = datetime.now().strftime("%a, %m/%d/%Y  %I:%M %p")
        for lbl in getattr(self,"_clock_labels",[]):
            try: lbl.setText(now)
            except Exception: pass

    def _add_clock_label(self, parent_layout):
        lbl = QLabel("")
        lbl.setStyleSheet("color:white; font-size:9pt;")
        if not hasattr(self,"_clock_labels"): self._clock_labels=[]
        self._clock_labels.append(lbl)
        parent_layout.addWidget(lbl)

    def _get_display_company_name(self):
        """Return the best available company name: settings → creds → fallback."""
        biz = get_business_settings(self.db).get("name", "").strip()
        if biz:
            return biz.upper()
        creds_name = load_creds().get("company_name", "").strip()
        if creds_name:
            return creds_name.upper()
        return "BLUE SKY SMOG"

    def _refresh_header_name(self):
        """Update the header label with the latest company name."""
        if hasattr(self, "_header_name_lbl"):
            self._header_name_lbl.setText(self._get_display_company_name())

    def _make_header(self, show_back=False):
        """Returns header_widget with zoom buttons."""
        hdr = QWidget(); hdr.setFixedHeight(44)
        hdr.setStyleSheet(f"background:{PRIMARY};")
        hl = QHBoxLayout(hdr); hl.setContentsMargins(8,4,8,4); hl.setSpacing(8)
        biz_name = self._get_display_company_name()
        name_lbl = QLabel(biz_name)
        name_lbl.setStyleSheet("color:white; font-size:13pt; font-weight:bold;")
        self._header_name_lbl = name_lbl
        hl.addWidget(name_lbl)
        sync_b = btn("Sync","accent"); sync_b.clicked.connect(self._manual_sync_now)
        hl.addWidget(sync_b)
        hl.addWidget(self._sync_label)
        hl.addStretch()
        # Zoom controls
        zm = QLabel("Zoom:"); zm.setStyleSheet("color:white; font-size:9pt;"); hl.addWidget(zm)
        _zoom_btn_style = ("background:#FFFFFF; color:#111827; border:none; border-radius:4px;"
                           " padding:2px 6px; font-size:9pt; font-weight:bold;")
        zm_out = QPushButton("A-"); zm_out.setFixedWidth(36); zm_out.setFixedHeight(26)
        zm_out.setStyleSheet(_zoom_btn_style)
        zm_out.clicked.connect(self._zoom_out); hl.addWidget(zm_out)
        zm_in  = QPushButton("A+"); zm_in.setFixedWidth(36); zm_in.setFixedHeight(26)
        zm_in.setStyleSheet(_zoom_btn_style)
        zm_in.clicked.connect(self._zoom_in); hl.addWidget(zm_in)
        self._add_clock_label(hl)
        return hdr

    def _zoom_in(self):
        s = self._current_screen; cur = self._zoom.get(s, 10)
        self._apply_zoom(s, min(cur + 1, 18))

    def _zoom_out(self):
        s = self._current_screen; cur = self._zoom.get(s, 10)
        self._apply_zoom(s, max(cur - 1, 7))

    def _apply_zoom(self, screen, size):
        self._zoom[screen] = size
        if screen in self._screens:
            self._screens[screen].setStyleSheet(f"font-size: {size}pt;")
        set_setting(self.db, "zoom_levels", json.dumps(self._zoom))
        # Re-populate tables that set explicit fonts on cells (so PASS/FAIL scales)
        if screen == "doc_list":
            self.refresh_doc_list()

    # â"€â"€ Column width persistence â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
    def _register_table(self, key, table):
        """Load saved column widths and connect resize signal for auto-save."""
        try:
            saved = json.loads(get_setting(self.db, f"col_widths_{key}", "{}"))
            for col_str, w in saved.items():
                col = int(col_str)
                if col < table.columnCount():
                    table.setColumnWidth(col, int(w))
        except Exception:
            pass
        table.horizontalHeader().sectionResized.connect(
            lambda logical, old, new, k=key, t=table: self._save_col_widths(k, t))

    def _save_col_widths(self, key, table):
        widths = {str(c): table.columnWidth(c) for c in range(table.columnCount())}
        try: set_setting(self.db, f"col_widths_{key}", json.dumps(widths))
        except Exception: pass

    # â"€â"€ Sync helpers â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
    def _on_sync_change(self):
        t = SYNC._last_pull_time
        since = t.strftime("%I:%M %p") if t else "never"
        cnt = SYNC._last_pull_count
        self._sync_label.setText(f"synced {since} (+{cnt})")
        self._refresh_header_name()
        if self._current_screen == "doc_list":
            self.refresh_doc_list()
        elif self._current_screen == "estimate_entry":
            self._refresh_acct_id_dropdown()

    def _on_force_pull_done(self, msg):
        QMessageBox.information(self, "Force Re-pull", msg)

    def _update_sync_status(self):
        t = SYNC._last_pull_time
        since = t.strftime("%I:%M %p") if t else "never"
        self._sync_label.setText(f"synced {since}")

    def _startup_sync(self):
        try: SYNC._flush(); SYNC._pull()
        except Exception: pass

    def _manual_sync_now(self):
        def _bg():
            try: SYNC._flush(); SYNC._pull()
            except Exception: pass
        threading.Thread(target=_bg, daemon=True).start()

    def _refresh_sub_status(self):
        def _fetch():
            s = api_subscription_status()
            if s: self._sub_status = s
        threading.Thread(target=_fetch, daemon=True).start()

    def _on_account_suspended(self, msg):
        self._sub_status = {"status":"locked","can_create":False,"warning":msg}
        QMessageBox.critical(self,"Account Suspended",f"{msg}\n\nYou are now in read-only mode.")

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: DOCUMENT LIST
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_doc_list_screen(self):
        w = QWidget(); self._screens["doc_list"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        lay.addWidget(self._make_header())
        self._stack.addWidget(w)

        # Search bar
        sb = QWidget(); sb_h = QHBoxLayout(sb); sb_h.setContentsMargins(8,4,8,4)
        sb_h.addWidget(QLabel("Search:"))
        self._dl_search = QLineEdit(); self._dl_search.setPlaceholderText("plate, name, #, date...")
        self._dl_search.setMaximumWidth(300)
        self._dl_search.textChanged.connect(self.refresh_doc_list); sb_h.addWidget(self._dl_search)
        sb_h.addWidget(QLabel("  Show:"))
        self._dl_filter = QComboBox()
        self._dl_filter.addItems(["All","Invoices","Estimates","Mobile"])
        self._dl_filter.currentTextChanged.connect(self.refresh_doc_list); sb_h.addWidget(self._dl_filter)
        sb_h.addStretch()
        self._dl_count_lbl = QLabel(""); self._dl_count_lbl.setStyleSheet(f"color:{PRIMARY}; font-weight:bold;")
        sb_h.addWidget(self._dl_count_lbl)
        lay.addWidget(sb)

        # Table
        cols = ["License #","Vehicle","State","#","Type","Customer","Amount","Result","Date","Paid With"]
        self._dl_table = QTableWidget(0, len(cols)); self._dl_table.setHorizontalHeaderLabels(cols)
        self._dl_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self._dl_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._dl_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._dl_table.setAlternatingRowColors(True)
        self._dl_table.verticalHeader().setVisible(False)
        self._dl_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._dl_table.customContextMenuRequested.connect(self._dl_context_menu)
        self._dl_table.doubleClicked.connect(self._dl_open)
        widths = [95,190,50,70,85,170,85,75,100,95]
        for i,w2 in enumerate(widths): self._dl_table.setColumnWidth(i,w2)
        self._register_table("doc_list", self._dl_table)
        lay.addWidget(self._dl_table)
        self._dl_inv_ids = []  # parallel list of invoice_ids per row

        # Bottom bar
        bb = QWidget(); bb.setObjectName("darkBar"); bb.setStyleSheet("QWidget#darkBar { background: #37474F; }")
        bb_h = QHBoxLayout(bb); bb_h.setContentsMargins(8,6,8,6); bb_h.setSpacing(6)
        for text, style, cb in [
            ("NEW ESTIMATE",  "success",    self._new_estimate_action),
            ("ACCOUNTS",      "secondary",  lambda: self.show_screen("account_setup")),
            ("CUSTOMERS",     "secondary",  lambda: self.show_screen("customers")),
            ("REPORTS",       "secondary",  lambda: self.show_screen("reports")),
            ("SYSTEM SETUP",  "secondary",  lambda: self.show_screen("settings")),
            ("LOGOUT",        "danger",     self._do_logout),
        ]:
            b = btn(text, style); b.clicked.connect(cb); bb_h.addWidget(b)
        if self._is_master:
            adm_b = btn("ADMIN BACKEND", "primary")
            adm_b.clicked.connect(lambda: self.show_screen("admin"))
            bb_h.addWidget(adm_b)
        bb_h.addStretch()
        lay.addWidget(bb)

    def _on_show_doc_list(self):
        self._refresh_header_name()
        self.refresh_doc_list()
        self._update_sync_status()

    def refresh_doc_list(self):
        q    = self._dl_search.text().strip().lower()
        filt = self._dl_filter.currentText()
        today = datetime.today().strftime("%Y-%m-%d")
        sql = ("""
            SELECT i.invoice_id,i.invoice_number,i.invoice_date,i.plate,i.vin,i.year,i.make,i.model,
                   i.veh_state,i.customer_name,i.first_name,i.last_name,i.company_name,i.customer_id,
                   i.amount_cents,i.payment_method,i.is_estimate,i.from_mobile,
                   CASE WHEN i.test_result != '' THEN i.test_result
                        ELSE COALESCE((
                            SELECT CASE WHEN MAX(CASE WHEN il.result IN ('FAIL','RETEST') THEN 1 ELSE 0 END)=1
                                        THEN 'FAIL'
                                        WHEN COUNT(CASE WHEN il.result!='' THEN 1 END)>0 THEN 'PASS'
                                        ELSE '' END
                            FROM invoice_lines il
                            WHERE il.invoice_id=i.invoice_id
                        ),'')
                   END AS test_result
            FROM invoices i WHERE 1=1""")
        if filt == "Invoices":    sql += " AND i.is_estimate=0"
        elif filt == "Estimates": sql += " AND i.is_estimate=1"
        elif filt == "Mobile":    sql += " AND i.from_mobile=1"
        sql += " ORDER BY i.is_estimate ASC, i.invoice_date ASC, i.invoice_number ASC"
        rows = self.db.execute(sql).fetchall()

        self._dl_table.setRowCount(0); self._dl_inv_ids = []; shown = 0
        for row in rows:
            cname = row["customer_name"] or f"{row['first_name']} {row['last_name']}".strip()
            plate = (row["plate"] or "").strip(); vin = (row["vin"] or "").strip()
            yr = (row["year"] or "").strip(); mk = (row["make"] or "").strip(); md = (row["model"] or "").strip()
            if not (plate or yr or mk or md):
                vrow = _best_vehicle_for_invoice(self.db, row)
                if vrow:
                    plate = vrow["plate"] or ""; yr = vrow["year"] or ""
                    mk = vrow["make"] or ""; md = vrow["model"] or ""
            ymm   = " ".join(filter(None,[yr,mk,md]))
            typ   = "ESTIMATE" if row["is_estimate"] else "INVOICE"
            amt   = f"${row['amount_cents']/100:,.2f}"
            num   = str(row["invoice_number"]) if row["invoice_number"] else "-"
            result= "" if row["is_estimate"] else (row["test_result"] or "").upper()
            paid  = (row["payment_method"] or "").upper()
            state = (row["veh_state"] or "CA").upper()
            if q:
                blob = " ".join([num,row["invoice_date"],plate,vin,ymm,cname,typ,paid]).lower()
                if q not in blob: continue

            r = self._dl_table.rowCount(); self._dl_table.insertRow(r)
            self._dl_inv_ids.append(row["invoice_id"])
            values = [plate,ymm,state,num,typ,cname,amt,result,row["invoice_date"],paid]
            is_today = (row["invoice_date"] == today)
            for col, val in enumerate(values):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter |
                    (Qt.AlignmentFlag.AlignRight if col==6 else
                     Qt.AlignmentFlag.AlignCenter if col in (2,3,4,7,8,9) else
                     Qt.AlignmentFlag.AlignLeft))
                if is_today:
                    item.setBackground(QColor(TODAY_BG))
                # Per-cell Result column coloring (col 7)
                if col == 7:
                    _zsz = self._zoom.get(self._current_screen, 10)
                    if result == "PASS":
                        item.setForeground(QColor(GREEN))
                        item.setFont(QFont("Segoe UI", _zsz, QFont.Weight.Bold))
                    elif result in ("FAIL","RETEST"):
                        item.setForeground(QColor(RED))
                        item.setFont(QFont("Segoe UI", _zsz, QFont.Weight.Bold))
                if row["is_estimate"] and col == 4:
                    item.setForeground(QColor(RED))
                self._dl_table.setItem(r, col, item)
            shown += 1

        self._dl_count_lbl.setText(f"{shown} Documents")
        self._dl_table.scrollToBottom()

    def _dl_selected_id(self):
        rows = self._dl_table.selectedItems()
        if not rows: return None
        r = self._dl_table.currentRow()
        if 0 <= r < len(self._dl_inv_ids): return self._dl_inv_ids[r]
        return None

    def _dl_open(self):
        iid = self._dl_selected_id()
        if iid: self._open_pdf_for_invoice(iid)

    def _dl_context_menu(self, pos):
        iid = self._dl_selected_id()
        if not iid: return
        menu = QMenu(self)
        menu.addAction("View PDF",    lambda: self._open_pdf_for_invoice(iid))
        menu.addAction("Edit",        lambda: self.show_screen("estimate_entry", invoice_id=iid))
        menu.addAction("Print PDF",   lambda: self._open_pdf_for_invoice(iid))
        menu.addSeparator()
        a = menu.addAction("Delete...", lambda: self._dl_delete(iid))
        a.setIcon(QIcon()); a.setIconVisibleInMenu(False)
        menu.exec(self._dl_table.viewport().mapToGlobal(pos))

    def _dl_delete(self, iid):
        row = self.db.execute("SELECT invoice_number,invoice_date,customer_name,first_name,last_name,is_estimate FROM invoices WHERE invoice_id=?",(iid,)).fetchone()
        if not row: return
        cname = row["customer_name"] or f"{row['first_name']} {row['last_name']}".strip()
        label = "Estimate" if row["is_estimate"] else "Invoice"
        num   = row["invoice_number"] or iid[:8]
        if QMessageBox.question(self,"Delete",
                f"Delete {label} #{num} - {cname}?\n\nCannot be undone.",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No
                ) != QMessageBox.StandardButton.Yes: return
        if requests:
            try:
                r = requests.delete(f"{API_BASE}/v1/invoices/{iid}", headers=_hdrs(), timeout=15); r.raise_for_status()
            except Exception as e: QMessageBox.warning(self,"Server",f"Server delete failed: {e}\nDeleted locally only.")
        pdf_row = self.db.execute("SELECT pdf_path FROM invoices WHERE invoice_id=?",(iid,)).fetchone()
        if pdf_row and pdf_row["pdf_path"] and os.path.exists(pdf_row["pdf_path"]):
            try: os.remove(pdf_row["pdf_path"])
            except: pass
        self.db.execute("DELETE FROM invoice_lines WHERE invoice_id=?",(iid,))
        self.db.execute("DELETE FROM invoices WHERE invoice_id=?",(iid,)); self.db.commit()
        self.refresh_doc_list()

    def _show_all(self):
        self._dl_search.clear(); self._dl_filter.setCurrentText("All")

    def _new_estimate_action(self):
        if not self._sub_status.get("can_create",True):
            QMessageBox.critical(self,"Subscription Required","Your free trial has ended.\n\nPlease subscribe to continue."); return
        self.show_screen("estimate_entry", invoice_id=None)

    def _do_logout(self):
        if QMessageBox.question(self,"Logout","Sign out and return to login?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No
                ) != QMessageBox.StandardButton.Yes: return
        save_creds({})
        SYNC.stop()
        try:
            for t in ("invoices","invoice_lines","customers","vehicles","outbox"):
                self.db.execute(f"DELETE FROM {t}")
            self.db.commit(); set_last_seq(self.db,0)
        except Exception: pass
        self.db.close()
        import subprocess
        subprocess.Popen([sys.executable]+sys.argv)
        self.close()

    def _open_pdf_for_invoice(self, iid):
        row = self.db.execute("SELECT is_estimate,invoice_date,customer_name,first_name,last_name,company_name FROM invoices WHERE invoice_id=?",(iid,)).fetchone()
        if not row: return
        pdf = build_invoice_pdf_path(self.inv_dir, row["invoice_date"],
            company=row["company_name"], first=row["first_name"],
            last=row["last_name"], customer_name=row["customer_name"],
            is_estimate=bool(row["is_estimate"]))
        ps = get_printer_setting(self.db)
        if not generate_invoice_pdf(iid, self.db, pdf):
            QMessageBox.critical(self,"Error","Could not generate PDF."); return
        dlg = PdfViewerDialog(pdf, ps, self); dlg.exec()

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: INVOICE / ESTIMATE ENTRY
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_estimate_entry_screen(self):
        outer = QWidget(); self._screens["estimate_entry"] = outer
        outer_lay = QVBoxLayout(outer); outer_lay.setContentsMargins(0,0,0,0); outer_lay.setSpacing(0)
        outer_lay.addWidget(self._make_header(show_back=True))
        self._stack.addWidget(outer)

        # Doc title bar
        title_bar = QWidget(); tb_h = QHBoxLayout(title_bar); tb_h.setContentsMargins(8,4,8,4)
        self._ee_type_lbl = QLabel("NEW ESTIMATE")
        self._ee_type_lbl.setStyleSheet(f"color:{PRIMARY}; font-size:12pt; font-weight:bold;")
        self._ee_num_lbl = QLabel(""); tb_h.addWidget(self._ee_type_lbl); tb_h.addWidget(self._ee_num_lbl)
        tb_h.addSpacing(20); tb_h.addWidget(QLabel("Date:"))
        self._inv_date_e = QLineEdit(datetime.today().strftime("%Y-%m-%d")); self._inv_date_e.setMaximumWidth(110)
        self._inv_date_e.setReadOnly(True)
        cal_btn = btn("Cal","secondary"); cal_btn.setFixedWidth(30); cal_btn.clicked.connect(self._pick_inv_date)
        tb_h.addWidget(self._inv_date_e); tb_h.addWidget(cal_btn); tb_h.addStretch()
        outer_lay.addWidget(title_bar)

        # Scrollable body
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        body = QWidget(); body_lay = QVBoxLayout(body); body_lay.setContentsMargins(10,4,10,4); body_lay.setSpacing(6)
        scroll.setWidget(body); outer_lay.addWidget(scroll)

        def grp(title): g = QGroupBox(title); return g

        # â"€â"€ Vehicle â"€â"€
        vg = grp("Vehicle Info"); vlay = QGridLayout(vg)
        self._f_plate = QLineEdit(); _upper_entry(self._f_plate); self._f_plate.setMaximumWidth(110)
        self._f_vstate = QLineEdit("CA"); _upper_entry(self._f_vstate); self._f_vstate.setMaximumWidth(50)
        self._f_vin = QLineEdit(); _upper_entry(self._f_vin); self._f_vin.setMinimumWidth(200)
        self._f_year  = QLineEdit(); _upper_entry(self._f_year); self._f_year.setMaximumWidth(60)
        self._f_make  = QLineEdit(); _upper_entry(self._f_make); self._f_make.setMaximumWidth(120)
        self._f_model = QLineEdit(); _upper_entry(self._f_model); self._f_model.setMaximumWidth(140)
        self._f_odo   = QLineEdit(); _upper_entry(self._f_odo); self._f_odo.setMaximumWidth(100)
        for c,(lbl_txt,w2) in enumerate([("License #",self._f_plate),("State",self._f_vstate),("VIN",self._f_vin)]):
            vlay.addWidget(QLabel(lbl_txt),0,c*2); vlay.addWidget(w2,0,c*2+1)
        for c,(lbl_txt,w2) in enumerate([("Year",self._f_year),("Make",self._f_make),("Model",self._f_model),("Odometer",self._f_odo)]):
            vlay.addWidget(QLabel(lbl_txt),1,c*2); vlay.addWidget(w2,1,c*2+1)
        body_lay.addWidget(vg)
        self._f_vin.editingFinished.connect(self._vin_lookup)
        # Plate field: auto-detects DMV barcode format (%0+VIN+PLATE) on Enter/focus-out
        self._f_plate.editingFinished.connect(self._plate_lookup)

        # â"€â"€ Customer â"€â"€
        cg = grp("Customer Info"); clay = QGridLayout(cg)
        self._f_acct_id = QComboBox(); self._f_acct_id.setEditable(True); self._f_acct_id.setMaximumWidth(180)
        self._f_first   = QLineEdit(); _upper_entry(self._f_first)
        self._f_last    = QLineEdit(); _upper_entry(self._f_last)
        self._f_company = QLineEdit(); _upper_entry(self._f_company)
        self._f_po      = QLineEdit(); _upper_entry(self._f_po); self._f_po.setMaximumWidth(120)
        clay.addWidget(QLabel("Account ID"),0,0); clay.addWidget(self._f_acct_id,0,1)
        clay.addWidget(QLabel("First"),0,2); clay.addWidget(self._f_first,0,3)
        clay.addWidget(QLabel("Last"),0,4); clay.addWidget(self._f_last,0,5)
        clay.addWidget(QLabel("Company"),1,0); clay.addWidget(self._f_company,1,1,1,3)
        clay.addWidget(QLabel("PO #"),1,4); clay.addWidget(self._f_po,1,5)
        body_lay.addWidget(cg)
        self._f_acct_id.currentTextChanged.connect(self._acct_id_changed)
        for w2 in (self._f_first,self._f_last,self._f_company):
            w2.textChanged.connect(self._autocomplete_customer)
            w2.editingFinished.connect(self._fill_customer)

        # â"€â"€ Address â"€â"€
        ag = grp("Address"); alay = QGridLayout(ag)
        self._f_addr  = QLineEdit(); _upper_entry(self._f_addr)
        self._f_zip   = QLineEdit(); _upper_entry(self._f_zip); self._f_zip.setMaximumWidth(80)
        self._f_city  = QLineEdit(); _upper_entry(self._f_city)
        self._f_state = QLineEdit(); _upper_entry(self._f_state); self._f_state.setMaximumWidth(50)
        self._f_phone = QLineEdit(); self._f_phone.setMaximumWidth(140)
        self._f_email = QLineEdit(); _upper_entry(self._f_email)
        alay.addWidget(QLabel("Address"),0,0); alay.addWidget(self._f_addr,0,1,1,3)
        alay.addWidget(QLabel("ZIP"),0,4); alay.addWidget(self._f_zip,0,5)
        alay.addWidget(QLabel("City"),0,6); alay.addWidget(self._f_city,0,7)
        alay.addWidget(QLabel("State"),0,8); alay.addWidget(self._f_state,0,9)
        alay.addWidget(QLabel("Phone"),1,0); alay.addWidget(self._f_phone,1,1)
        alay.addWidget(QLabel("Email"),1,2); alay.addWidget(self._f_email,1,3,1,5)
        body_lay.addWidget(ag)
        self._f_zip.editingFinished.connect(self._zip_lookup)
        self._f_phone.editingFinished.connect(self._fmt_phone)

        # â"€â"€ Notes â"€â"€
        ng = grp("Notes"); nlay = QVBoxLayout(ng)
        self._f_notes = QTextEdit(); self._f_notes.setMaximumHeight(70)
        nlay.addWidget(self._f_notes); body_lay.addWidget(ng)

        # â"€â"€ Service line adder â"€â"€
        sg = grp("Add Service Line"); slay = QHBoxLayout(sg)
        slay.addWidget(QLabel("Service:"))
        self._f_svc = QComboBox(); self._f_svc.setMinimumWidth(180)
        slay.addWidget(self._f_svc)
        slay.addWidget(QLabel("Result:"))
        self._f_result = QComboBox(); self._f_result.addItems(["Pass","Fail","Retest"])
        slay.addWidget(self._f_result)
        slay.addWidget(QLabel("Cert #:")); self._f_cert = QLineEdit(); self._f_cert.setMaximumWidth(120)
        slay.addWidget(self._f_cert)
        slay.addWidget(QLabel("Disc $:")); self._f_disc = QLineEdit("0"); self._f_disc.setMaximumWidth(70)
        slay.addWidget(self._f_disc)
        add_b = btn("Add Line","primary"); add_b.clicked.connect(self._add_line); slay.addWidget(add_b)
        body_lay.addWidget(sg)

        # â"€â"€ Service lines table â"€â"€
        lg = grp("Service Lines"); llay = QVBoxLayout(lg)
        self._lines_table = QTableWidget(0,6)
        self._lines_table.setHorizontalHeaderLabels(["VIN","Service","Result","Cert #","Discount","Price"])
        self._lines_table.setMaximumHeight(150)
        self._lines_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._lines_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._lines_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._lines_table.doubleClicked.connect(self._edit_line)
        llay.addWidget(self._lines_table)
        tot_h = QHBoxLayout()
        rm_b = btn("Remove Selected","secondary"); rm_b.clicked.connect(self._remove_line); tot_h.addWidget(rm_b)
        tot_h.addStretch()
        self._total_lbl = QLabel("Total: $0.00")
        self._total_lbl.setStyleSheet(f"color:{PRIMARY}; font-size:12pt; font-weight:bold;")
        tot_h.addWidget(self._total_lbl); llay.addLayout(tot_h); body_lay.addWidget(lg)

        # â"€â"€ Payment â"€â"€
        pg = grp("Payment"); play = QHBoxLayout(pg)
        play.addWidget(QLabel("Method:"))
        self._f_pay = QComboBox()
        self._f_pay.addItems(["","CASH","VISA","MASTERCARD","DISCOVER","AMEX","CHECK","CHARGE"])
        self._f_pay.currentTextChanged.connect(self._payment_changed); play.addWidget(self._f_pay)
        play.addStretch()
        self._ee_total_big = QLabel("TOTAL: $0.00")
        self._ee_total_big.setStyleSheet(f"color:{PRIMARY}; font-size:14pt; font-weight:bold;")
        play.addWidget(self._ee_total_big); body_lay.addWidget(pg)
        body_lay.addStretch()

        # â"€â"€ Bottom button bar â"€â"€
        bb = QWidget(); bb.setObjectName("darkBar"); bb.setStyleSheet("QWidget#darkBar { background: #37474F; }")
        bb_h = QHBoxLayout(bb); bb_h.setContentsMargins(8,6,8,6); bb_h.setSpacing(6)
        for text, style, cb in [
            ("DOCUMENT LIST","secondary",  lambda: self.show_screen("doc_list")),
            ("SAVE ESTIMATE","secondary",  self._save_estimate_action),
            ("ISSUE",        "success",    self._issue_action),
            ("PRINT PDF",    "secondary",  self._open_selected_pdf),
            ("DELETE",       "danger",     self._delete_document_action),
            ("CLEAR",        "danger",     self._clear_form),
        ]:
            b = btn(text,style); b.clicked.connect(cb); bb_h.addWidget(b)
        bb_h.addStretch(); outer_lay.addWidget(bb)

    def _on_show_estimate_entry(self, invoice_id=None):
        self._refresh_acct_id_dropdown()
        svc_names = list(get_services(self.db).keys())
        self._f_svc.clear(); self._f_svc.addItems(svc_names)
        if invoice_id: self._load_invoice_into_form(invoice_id)
        else: self._clear_form()

    def _refresh_acct_id_dropdown(self):
        names = [r[0] for r in self.db.execute("SELECT company_name FROM accounts ORDER BY company_name").fetchall()]
        cur = self._f_acct_id.currentText()
        self._f_acct_id.blockSignals(True)
        self._f_acct_id.clear(); self._f_acct_id.addItems([""] + names)
        self._f_acct_id.setCurrentText(cur)
        self._f_acct_id.blockSignals(False)

    def _acct_id_changed(self, name):
        name = name.strip()
        if not name: return
        row = self.db.execute("SELECT * FROM accounts WHERE UPPER(company_name)=?",(name.upper(),)).fetchone()
        if not row: return
        def si(w2, val):
            if not w2.text().strip() and val: w2.setText(val)
        si(self._f_company, row["company_name"]); si(self._f_phone, row["phone"])
        si(self._f_email, row["email"]); si(self._f_addr, row["address1"])
        si(self._f_city, row["city"]); si(self._f_state, row["state"]); si(self._f_zip, row["zip"])

    def _autocomplete_customer(self):
        self._customer_touched = True

    def _fill_customer(self):
        if not self._customer_touched: return
        first = self._f_first.text().strip().upper()
        last  = self._f_last.text().strip().upper()
        co    = self._f_company.text().strip().upper()
        if not (first or last or co): return
        key = f"{first} {last}".strip() or co
        cust = self.db.execute(
            "SELECT * FROM customers WHERE UPPER(first_name||' '||last_name)=? OR UPPER(company_name)=? LIMIT 1",
            (key,key)).fetchone()
        if not cust: return
        def si(w2, val):
            if not w2.text().strip() and val: w2.setText(val)
        si(self._f_phone,cust["phone"]); si(self._f_email,cust["email"])
        si(self._f_addr,cust["address"]); si(self._f_city,cust["city"])
        si(self._f_state,cust["state"]); si(self._f_zip,cust["zip"])

    def _vin_lookup(self):
        vin = self._f_vin.text().strip().upper()
        if len(vin) != 17: return
        # Local DB first
        vrow = self.db.execute("SELECT * FROM vehicles WHERE vin=? LIMIT 1",(vin,)).fetchone()
        if vrow:
            if not self._f_year.text(): self._f_year.setText(vrow["year"] or "")
            if not self._f_make.text(): self._f_make.setText(vrow["make"] or "")
            if not self._f_model.text(): self._f_model.setText(vrow["model"] or "")
            return
        self._vin_worker = VinWorker(vin); self._vin_worker.done.connect(self._vin_done); self._vin_worker.start()

    def _vin_done(self, yr, mk, md):
        if yr and not self._f_year.text(): self._f_year.setText(yr)
        if mk and not self._f_make.text(): self._f_make.setText(mk)
        if md and not self._f_model.text(): self._f_model.setText(md)

    def _plate_lookup(self):
        raw = self._f_plate.text().strip().upper()
        if not raw: return
        # ── DMV / invoice barcode detection ──────────────────────────────
        # Physical barcode scanners type data into this field.
        # Format: [garbage_prefix] + VIN(17 chars) + PLATE(rest)
        # The garbage prefix can be %0, $0, or any other non-alphanumeric chars.
        # Strip all leading non-alphanumeric characters, then check for 17+ chars.
        stripped = re.sub(r'^[^A-Z0-9]+', '', raw)
        if stripped != raw and len(stripped) >= 17:
            vin   = stripped[:17]
            plate = stripped[17:].strip()
            self._f_plate.setText(plate)
            if vin:
                self._f_vin.setText(vin)
                self._vin_lookup()
            if plate:
                vrow = self.db.execute("SELECT * FROM vehicles WHERE UPPER(plate)=? ORDER BY updated_at DESC LIMIT 1",(plate,)).fetchone()
                if vrow:
                    if not self._f_year.text():  self._f_year.setText(vrow["year"] or "")
                    if not self._f_make.text():  self._f_make.setText(vrow["make"] or "")
                    if not self._f_model.text(): self._f_model.setText(vrow["model"] or "")
            return
        # ── Normal plate lookup ───────────────────────────────────────────
        vrow = self.db.execute("SELECT * FROM vehicles WHERE UPPER(plate)=? ORDER BY updated_at DESC LIMIT 1",(raw,)).fetchone()
        if not vrow: return
        if not self._f_vin.text():   self._f_vin.setText(vrow["vin"] or "")
        if not self._f_year.text():  self._f_year.setText(vrow["year"] or "")
        if not self._f_make.text():  self._f_make.setText(vrow["make"] or "")
        if not self._f_model.text(): self._f_model.setText(vrow["model"] or "")

    def _zip_lookup(self):
        z = self._f_zip.text().strip()
        if len(z) != 5 or not z.isdigit(): return
        if self._f_city.text().strip() and self._f_state.text().strip(): return
        self._zip_worker = ZipWorker(z); self._zip_worker.done.connect(lambda c,s: (self._f_city.setText(c) if c else None, self._f_state.setText(s) if s else None))
        self._zip_worker.start()

    def _fmt_phone(self):
        self._f_phone.setText(format_phone(self._f_phone.text()))

    def _pick_inv_date(self):
        dlg = QDialog(self); dlg.setWindowTitle("Pick Date")
        lay = QVBoxLayout(dlg)
        cal = QCalendarWidget(); cal.setSelectedDate(QDate.fromString(self._inv_date_e.text(),"yyyy-MM-dd"))
        lay.addWidget(cal)
        ok = btn("OK","primary"); ok.clicked.connect(dlg.accept); lay.addWidget(ok)
        if dlg.exec(): self._inv_date_e.setText(cal.selectedDate().toString("yyyy-MM-dd"))

    def _get_service_price(self, svc_name, result):
        svcs = get_services(self.db); s = svcs.get(svc_name,{}); base = float(s.get("price",0))
        try:
            acct = self._f_acct_id.currentText().strip().upper()
            if acct:
                row = self.db.execute("SELECT custom_pricing FROM accounts WHERE UPPER(company_name)=?",(acct,)).fetchone()
                if row:
                    cp = json.loads(row["custom_pricing"] or "{}")
                    if svc_name in cp: base = float(cp[svc_name])
        except Exception: pass
        if svc_name == "Smog Test" and result == "Pass":
            base += float(s.get("cert_fee",8.25))
        return base

    def _add_line(self):
        svc = self._f_svc.currentText().strip(); result = self._f_result.currentText().strip()
        if not svc: QMessageBox.warning(self,"Missing","Select a service."); return
        vin   = self._f_vin.text().strip();   plate = self._f_plate.text().strip()
        odo   = self._f_odo.text().strip();    year  = self._f_year.text().strip()
        make  = self._f_make.text().strip();   model = self._f_model.text().strip()
        cert  = self._f_cert.text().strip()
        try: disc = float(self._f_disc.text().strip() or 0)
        except: disc = 0.0
        price = max(self._get_service_price(svc, result) - disc, 0)
        d = dict(vin=vin,plate=plate,odometer=odo,year=year,make=make[:8],model=model[:10],
                 service=svc,result=result,cert=cert,discount=disc,price=price,remote_item_id="")
        self._lines_data.append(d)
        r = self._lines_table.rowCount(); self._lines_table.insertRow(r)
        for col, val in enumerate([vin,svc,result,cert,f"${disc:.2f}",f"${price:.2f}"]):
            self._lines_table.setItem(r,col,QTableWidgetItem(val))
        self._update_total(); self._payment_changed()
        for w2 in (self._f_vin,self._f_plate,self._f_odo,self._f_year,self._f_make,self._f_model,self._f_cert):
            w2.clear()
        self._f_disc.setText("0"); self._f_result.setCurrentText("Pass")

    def _edit_line(self):
        r = self._lines_table.currentRow()
        if r < 0 or r >= len(self._lines_data): return
        d = self._lines_data[r]
        self._f_svc.setCurrentText(d.get("service",""))
        self._f_result.setCurrentText(d.get("result","Pass"))
        self._f_cert.setText(d.get("cert",""))
        self._f_disc.setText(str(d.get("discount",0)))
        self._lines_data.pop(r); self._lines_table.removeRow(r); self._update_total()

    def _remove_line(self):
        r = self._lines_table.currentRow()
        if r < 0 or r >= len(self._lines_data): return
        self._lines_data.pop(r); self._lines_table.removeRow(r); self._update_total()

    def _payment_changed(self, text=None):
        pay = self._f_pay.currentText().upper()
        self._lines_data = [d for d in self._lines_data if d["service"] != "Credit Card Fee"]
        self._lines_table.setRowCount(0)
        for d in self._lines_data:
            r = self._lines_table.rowCount(); self._lines_table.insertRow(r)
            for col,val in enumerate([d.get("vin",""),d["service"],d["result"],d["cert"],f"${d['discount']:.2f}",f"${d['price']:.2f}"]):
                self._lines_table.setItem(r,col,QTableWidgetItem(val))
        if pay not in ("","CASH","CHECK","CHARGE"):
            biz = get_business_settings(self.db); fee = float(biz.get("card_fee",5.0))
            d = dict(vin="",plate="",odometer="",year="",make="",model="",service="Credit Card Fee",
                     result="",cert="",discount=0.0,price=fee,remote_item_id="")
            self._lines_data.append(d)
            r = self._lines_table.rowCount(); self._lines_table.insertRow(r)
            for col,val in enumerate(["","Credit Card Fee","","","$0.00",f"${fee:.2f}"]):
                self._lines_table.setItem(r,col,QTableWidgetItem(val))
        self._update_total()

    def _update_total(self):
        total = sum(d["price"] for d in self._lines_data)
        self._total_lbl.setText(f"Total: ${total:,.2f}")
        self._ee_total_big.setText(f"TOTAL: ${total:,.2f}")

    def _collect_form(self):
        def u(w2): return w2.text().strip().upper()
        return {
            "first":     u(self._f_first),    "last":    u(self._f_last),
            "company":   u(self._f_company),  "addr":    u(self._f_addr),
            "city":      u(self._f_city),      "state":   u(self._f_state),
            "zip":       u(self._f_zip),       "phone":   format_phone(self._f_phone.text().strip()),
            "email":     u(self._f_email),     "veh_state": u(self._f_vstate) or "CA",
            "acct_id":   self._f_acct_id.currentText().strip().upper(),
            "po":        u(self._f_po),
            "date":      self._inv_date_e.text(),
            "pay":       self._f_pay.currentText().upper(),
            "notes":     self._f_notes.toPlainText().strip(),
        }

    def _clear_form(self):
        self._editing_id = None; self._customer_touched = False
        for w2 in (self._f_first,self._f_last,self._f_company,self._f_addr,self._f_city,
                   self._f_state,self._f_zip,self._f_phone,self._f_email,self._f_plate,
                   self._f_year,self._f_make,self._f_model,self._f_vin,self._f_odo,self._f_po,
                   self._f_cert,self._f_disc):
            w2.clear()
        self._f_vstate.setText("CA"); self._f_disc.setText("0")
        self._f_result.setCurrentText("Pass"); self._f_pay.setCurrentIndex(0)
        self._f_acct_id.setCurrentText(""); self._f_notes.clear()
        self._inv_date_e.setText(datetime.today().strftime("%Y-%m-%d"))
        self._lines_data.clear(); self._lines_table.setRowCount(0)
        self._update_total(); self._ee_type_lbl.setText("NEW ESTIMATE"); self._ee_num_lbl.setText("")

    def _load_invoice_into_form(self, invoice_id):
        self._clear_form(); self._editing_id = invoice_id
        inv = self.db.execute("SELECT * FROM invoices WHERE invoice_id=?",(invoice_id,)).fetchone()
        if not inv: return
        lines = self.db.execute("SELECT * FROM invoice_lines WHERE invoice_id=? ORDER BY id",(invoice_id,)).fetchall()
        self._f_first.setText(inv["first_name"] or "")
        self._f_last.setText(inv["last_name"] or "")
        self._f_company.setText(inv["company_name"] or "")
        self._f_acct_id.setCurrentText(inv["account_id"] or "")
        self._f_po.setText(inv["po_number"] or "")
        self._inv_date_e.setText(inv["invoice_date"] or datetime.today().strftime("%Y-%m-%d"))
        cust = self.db.execute("SELECT * FROM customers WHERE customer_id=?",(inv["customer_id"],)).fetchone()
        if cust:
            self._f_phone.setText(cust["phone"] or ""); self._f_email.setText(cust["email"] or "")
            self._f_addr.setText(cust["address"] or ""); self._f_city.setText(cust["city"] or "")
            self._f_state.setText(cust["state"] or ""); self._f_zip.setText(cust["zip"] or "")
        self._f_vstate.setText(inv["veh_state"] or "CA")
        self._f_notes.setPlainText(inv["notes"] or "")
        self._f_pay.setCurrentText(inv["payment_method"] or "")
        self._lines_data.clear(); self._lines_table.setRowCount(0)
        for line in lines:
            d = dict(vin=line["vin"]or"",plate=line["plate"]or"",odometer=line["odometer"]or"",
                     year=line["year"]or"",make=line["make"]or"",model=line["model"]or"",
                     service=line["service"]or"",result=line["result"]or"",cert=line["cert"]or"",
                     discount=float(line["discount"]or 0),price=float(line["price"]or 0),
                     remote_item_id=line["remote_item_id"] if "remote_item_id" in line.keys() else "")
            self._lines_data.append(d)
            r = self._lines_table.rowCount(); self._lines_table.insertRow(r)
            for col,val in enumerate([d["vin"],d["service"],d["result"],d["cert"],f"${d['discount']:.2f}",f"${d['price']:.2f}"]):
                self._lines_table.setItem(r,col,QTableWidgetItem(val))
        self._update_total()
        if inv["is_estimate"]:
            self._ee_type_lbl.setText("ESTIMATE")
        else:
            self._ee_type_lbl.setText("INVOICE")
        num = inv["invoice_number"] or ""
        self._ee_num_lbl.setText(f"  #{num}" if num else "  #PENDING")
        # Populate vehicle from first line
        if lines:
            l = lines[0]
            self._f_vin.setText(l["vin"] or inv["vin"] or "")
            self._f_plate.setText(l["plate"] or inv["plate"] or "")
            self._f_year.setText(l["year"] or inv["year"] or "")
            self._f_make.setText(l["make"] or inv["make"] or "")
            self._f_model.setText(l["model"] or inv["model"] or "")

    def _save_estimate_action(self): self._save_doc(is_estimate=True)

    def _issue_action(self):
        if not self._lines_data: QMessageBox.warning(self,"No Lines","Add at least one service line."); return
        if self._editing_id:
            inv = self.db.execute("SELECT is_estimate FROM invoices WHERE invoice_id=?",(self._editing_id,)).fetchone()
            if inv and not inv["is_estimate"]:
                ans = QMessageBox.question(
                    self, "Re-issue Invoice",
                    "This document has already been issued as an invoice.\n\n"
                    "Do you want to re-issue it now (e.g. to correct lines or payment method)?\n"
                    "The invoice number will be preserved.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if ans == QMessageBox.StandardButton.Yes:
                    self._save_doc(is_estimate=False)
                return
        if QMessageBox.question(self,"Issue Invoice","Issue as a new invoice?\nSelect a payment method first.",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self._save_doc(is_estimate=False)

    def _open_selected_pdf(self):
        if self._editing_id: self._open_pdf_for_invoice(self._editing_id)
        else: QMessageBox.information(self,"No Invoice","Save or issue first.")

    def _delete_document_action(self):
        if not self._editing_id: QMessageBox.information(self,"Nothing","No document loaded."); return
        self._dl_delete(self._editing_id); self._clear_form(); self.show_screen("doc_list")

    def _save_doc(self, is_estimate):
        if not self._sub_status.get("can_create",True):
            QMessageBox.critical(self,"Subscription Required","Your free trial has ended."); return
        if not self._lines_data: QMessageBox.warning(self,"Error","Add at least one line."); return
        fd = self._collect_form()
        if not is_estimate and not fd["pay"]:
            QMessageBox.warning(self,"Missing","Select a payment method."); return

        cid = upsert_customer(self.db,fd["first"],fd["last"],fd["company"],
                              phone=fd["phone"],email=fd["email"],address=fd["addr"],
                              city=fd["city"],state=fd["state"],zip_=fd["zip"])
        _seen = set()
        for d in self._lines_data:
            if not (d["vin"] or d["plate"]): continue
            key = d["vin"] or d["plate"]
            if key in _seen: continue; _seen.add(key)
            vid = upsert_vehicle(self.db,cid,d["vin"],d["plate"],d["make"],d["model"],d["year"])
            enqueue(self.db,"vehicle","upsert",{"vehicle_id":vid,"customer_id":cid,"vin":d["vin"],
                "plate":d["plate"],"make":d["make"],"model":d["model"],"year":d["year"],
                "odometer":d["odometer"],"service_type":""})

        total_cents = int(sum(d["price"] for d in self._lines_data) * 100)
        status = "ESTIMATE" if is_estimate else ("CHARGE" if fd["pay"]=="CHARGE" else "PAID")
        cname  = fd["company"] or f"{fd['first']} {fd['last']}".strip() or "Customer"
        plate  = next((d["plate"] for d in self._lines_data if d["plate"]),"")
        vin    = next((d["vin"]   for d in self._lines_data if d["vin"]),"")
        yr     = next((d["year"]  for d in self._lines_data if d["year"]),"")
        mk     = next((d["make"]  for d in self._lines_data if d["make"]),"")
        md     = next((d["model"] for d in self._lines_data if d["model"]),"")
        if is_estimate: agg_result = ""
        else:
            results = [d["result"].upper() for d in self._lines_data if d.get("result","").strip()]
            agg_result = "FAIL" if any(r in ("FAIL","RETEST") for r in results) else ("PASS" if results else "")

        charge_acct_co = ""
        if fd["pay"] == "CHARGE" and not is_estimate:
            charge_acct_co = (fd["acct_id"] or fd["company"] or f"{fd['first']} {fd['last']}".strip()).upper()
        effective_acct_id = charge_acct_co if charge_acct_co else fd["acct_id"]

        if self._editing_id:
            iid = self._editing_id
            inv_num = self.db.execute("SELECT invoice_number FROM invoices WHERE invoice_id=?",(iid,)).fetchone()["invoice_number"]
            self.db.execute("DELETE FROM invoice_lines WHERE invoice_id=?",(iid,))
        else:
            iid = str(uuid.uuid4()); inv_num = get_next_invoice_number(self.db)

        self.db.execute("""
            INSERT OR REPLACE INTO invoices
            (invoice_id,invoice_number,customer_id,customer_name,first_name,last_name,
             company_name,invoice_date,plate,vin,year,make,model,amount_cents,
             payment_method,status,notes,is_estimate,from_mobile,created_at,updated_at,synced,
             veh_state,account_id,po_number,test_result)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0,?,?,0,?,?,?,?)
        """,(iid,inv_num,cid,cname,fd["first"],fd["last"],fd["company"],fd["date"],
             plate,vin,yr,mk,md,total_cents,fd["pay"],status,fd["notes"],
             1 if is_estimate else 0,now_iso(),now_iso(),
             fd["veh_state"],effective_acct_id,fd["po"],agg_result))

        for d in self._lines_data:
            if not d.get("remote_item_id"): d["remote_item_id"] = str(uuid.uuid4())
            self.db.execute("""
                INSERT INTO invoice_lines
                (invoice_id,vin,plate,odometer,year,make,model,service,result,cert,discount,price,remote_item_id)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,(iid,d["vin"],d["plate"],d["odometer"],d["year"],d["make"],d["model"],
                 d["service"],d["result"],d["cert"],d["discount"],d["price"],d["remote_item_id"]))
        self.db.commit()

        # AR tracking
        if charge_acct_co:
            try:
                co = charge_acct_co
                contact_name = f"{fd['first']} {fd['last']}".strip()
                self.db.execute("""
                    INSERT INTO accounts(company_name,total_owed,updated_at,contact_name,phone,email,address1,city,state,zip)
                    VALUES(?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(company_name) DO UPDATE SET
                        total_owed=total_owed+excluded.total_owed, updated_at=excluded.updated_at,
                        contact_name=CASE WHEN contact_name='' THEN excluded.contact_name ELSE contact_name END,
                        phone=CASE WHEN phone='' THEN excluded.phone ELSE phone END,
                        email=CASE WHEN email='' THEN excluded.email ELSE email END,
                        address1=CASE WHEN address1='' THEN excluded.address1 ELSE address1 END,
                        city=CASE WHEN city='' THEN excluded.city ELSE city END,
                        state=CASE WHEN state='' THEN excluded.state ELSE state END,
                        zip=CASE WHEN zip='' THEN excluded.zip ELSE zip END
                """,(co,total_cents/100,now_iso(),contact_name,fd["phone"],fd["email"],fd["addr"],fd["city"],fd["state"],fd["zip"]))
                self.db.execute("INSERT INTO account_history(company_name,entry_date,type,amount,invoice_id) VALUES(?,?,?,?,?)",
                                (co,fd["date"],"charge",total_cents/100,iid))
                self.db.commit()
                if self._current_screen == "account_setup":
                    self._on_show_account_setup(co)
            except Exception as e:
                import traceback
                QMessageBox.critical(self,"AR Error",f"Could not update account balance:\n{e}\n\n{traceback.format_exc()}")

        # Enqueue for sync
        enqueue(self.db,"invoice","upsert",{"invoice_id":iid,"invoice_number":inv_num or 0,
            "customer_id":cid,"customer_name":cname,"first_name":fd["first"],"last_name":fd["last"],
            "company_name":fd["company"],"invoice_date":fd["date"],"plate":plate,"vin":vin,
            "year":yr,"make":mk,"model":md,"amount_cents":total_cents,"payment_method":fd["pay"],
            "status":status,"notes":fd["notes"],"is_estimate":1 if is_estimate else 0})

        # Generate PDF and show viewer
        pdf = build_invoice_pdf_path(self.inv_dir,fd["date"],company=fd["company"],
                                     first=fd["first"],last=fd["last"],customer_name=cname,is_estimate=is_estimate)
        ps = get_printer_setting(self.db)
        generate_invoice_pdf(iid,self.db,pdf)
        self._editing_id = iid
        self._ee_type_lbl.setText("ESTIMATE" if is_estimate else "INVOICE")

        # Background sync to get real invoice number
        def _bg_sync():
            try: SYNC._flush(); SYNC._pull()
            except Exception: pass
        threading.Thread(target=_bg_sync, daemon=True).start()

        if ps.get("auto_print") and not is_estimate:
            print_pdf(pdf, printer_name=ps.get("printer_name",""), copies=int(ps.get("copies",2)), parent_widget=self)
        dlg = PdfViewerDialog(pdf, ps, self); dlg.exec()
        self.show_screen("doc_list")

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: ACCOUNTS / AR
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_account_setup_screen(self):
        w = QWidget(); self._screens["account_setup"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        lay.addWidget(self._make_header(show_back=True))
        self._stack.addWidget(w)

        # Toolbar
        tb = QWidget(); tb_h = QHBoxLayout(tb); tb_h.setContentsMargins(8,4,8,4); tb_h.setSpacing(6)
        tb_h.addWidget(QLabel("Account:"))
        self._acct_combo = QComboBox(); self._acct_combo.setMinimumWidth(200)
        self._acct_combo.currentTextChanged.connect(self._acct_selected); tb_h.addWidget(self._acct_combo)
        for sym, cb in [("<<",self._acct_first),("<",self._acct_prev),(">",self._acct_next),(">>",self._acct_last)]:
            b = btn(sym,"secondary"); b.setFixedWidth(30); b.clicked.connect(cb); tb_h.addWidget(b)
        tb_h.addSpacing(12)
        new_b = btn("NEW ACCOUNT","success"); new_b.clicked.connect(self._new_acct_dialog); tb_h.addWidget(new_b)
        edit_b = btn("EDIT ACCOUNT","primary"); edit_b.clicked.connect(self._edit_acct_dialog); tb_h.addWidget(edit_b)
        del_b = btn("DELETE","danger"); del_b.clicked.connect(self._delete_acct_action); tb_h.addWidget(del_b)
        tb_h.addStretch()
        lay.addWidget(tb)

        # Body (scrollable)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        body = QWidget(); body_lay = QVBoxLayout(body); body_lay.setContentsMargins(12,8,12,8); body_lay.setSpacing(8)
        scroll.setWidget(body); lay.addWidget(scroll)

        # Customer info section
        info_grp = QGroupBox("Customer Info"); info_grid = QGridLayout(info_grp)
        self._av_name    = QLineEdit(); self._av_name.setReadOnly(True)
        self._av_contact = QLineEdit(); self._av_contact.setReadOnly(True)
        self._av_phone   = QLineEdit(); self._av_phone.setReadOnly(True)
        self._av_email   = QLineEdit(); self._av_email.setReadOnly(True)
        self._av_addr1   = QLineEdit(); self._av_addr1.setReadOnly(True)
        self._av_city    = QLineEdit(); self._av_city.setReadOnly(True)
        self._av_state   = QLineEdit(); self._av_state.setReadOnly(True)
        self._av_zip     = QLineEdit(); self._av_zip.setReadOnly(True)
        self._av_status  = QLineEdit(); self._av_status.setReadOnly(True)
        for row_idx,(lbl_txt,w2) in enumerate([("Account Name",self._av_name),("Contact",self._av_contact),
                ("Phone",self._av_phone),("Email",self._av_email),("Address",self._av_addr1),
                ("City",self._av_city),("State",self._av_state),("ZIP",self._av_zip),("Status",self._av_status)]):
            info_grid.addWidget(QLabel(lbl_txt),row_idx,0); info_grid.addWidget(w2,row_idx,1)
        body_lay.addWidget(info_grp)

        # Balance + action buttons
        bal_h = QHBoxLayout()
        self._acct_balance_lbl = QLabel("Balance Owed: $0.00")
        self._acct_balance_lbl.setStyleSheet(f"color:{PRIMARY}; font-size:12pt; font-weight:bold;")
        bal_h.addWidget(self._acct_balance_lbl)
        pay_b = btn("Post Payment","primary"); pay_b.clicked.connect(self._acct_post_payment); bal_h.addWidget(pay_b)
        prt_b = btn("Print Statement","secondary"); prt_b.clicked.connect(self._acct_print_statement); bal_h.addWidget(prt_b)
        bal_h.addStretch(); body_lay.addLayout(bal_h)

        # Customer History table (invoices + payments combined)
        hist_grp = QGroupBox("Customer History")
        hist_lay = QVBoxLayout(hist_grp)
        self._acct_hist_table = QTableWidget(0, 6)
        self._acct_hist_table.setHorizontalHeaderLabels(["Date","Type","Reference #","Invoice","Payment","Notes"])
        self._acct_hist_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._acct_hist_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._acct_hist_table.setAlternatingRowColors(True)
        hh = self._acct_hist_table.horizontalHeader()
        hh.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        for ci, w2 in enumerate([90, 80, 100, 90, 90]):
            self._acct_hist_table.setColumnWidth(ci, w2)
        self._acct_hist_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._acct_hist_table.customContextMenuRequested.connect(self._acct_hist_context_menu)
        self._acct_hist_table.doubleClicked.connect(self._acct_hist_double_click)
        self._register_table("acct_hist", self._acct_hist_table)
        hist_lay.addWidget(self._acct_hist_table); body_lay.addWidget(hist_grp)
        body_lay.addStretch()
        bb = QWidget(); bb.setObjectName("darkBar"); bb.setStyleSheet("QWidget#darkBar { background: #37474F; }")
        bb_h = QHBoxLayout(bb); bb_h.setContentsMargins(8,6,8,6)
        dl_b = btn("DOCUMENTS LIST","secondary"); dl_b.clicked.connect(lambda: self.show_screen("doc_list")); bb_h.addWidget(dl_b)
        bb_h.addStretch(); lay.addWidget(bb)

    def _on_show_account_setup(self, company_name=None):
        self._acct_names = [r[0] for r in self.db.execute("SELECT company_name FROM accounts ORDER BY company_name").fetchall()]
        self._acct_combo.blockSignals(True)
        self._acct_combo.clear(); self._acct_combo.addItems(self._acct_names)
        if company_name and company_name in self._acct_names:
            self._acct_index = self._acct_names.index(company_name)
            self._acct_combo.setCurrentText(company_name)
        elif self._acct_names:
            self._acct_index = 0; self._acct_combo.setCurrentIndex(0)
        self._acct_combo.blockSignals(False)
        if self._acct_names: self._load_acct(self._acct_names[self._acct_index])

    def _acct_selected(self, name):
        if name and name in self._acct_names:
            self._acct_index = self._acct_names.index(name); self._load_acct(name)

    def _load_acct(self, company_name):
        row = self.db.execute("SELECT * FROM accounts WHERE company_name=?",(company_name,)).fetchone()
        if not row: return
        self._av_name.setText(row["company_name"] or "")
        self._av_contact.setText(row["contact_name"] or "")
        self._av_phone.setText(row["phone"] or ""); self._av_email.setText(row["email"] or "")
        self._av_addr1.setText(row["address1"] or ""); self._av_city.setText(row["city"] or "")
        self._av_state.setText(row["state"] or ""); self._av_zip.setText(row["zip"] or "")
        self._av_status.setText(row["account_status"] or "Active")
        bal = row["total_owed"] or 0.0
        self._acct_balance_lbl.setText(f"Balance Owed: ${bal:,.2f}")
        self._refresh_acct_history(company_name)

    def _refresh_acct_history(self, company_name):
        """Populate Customer History with invoices and payments merged by date."""
        bal = self.db.execute("SELECT total_owed FROM accounts WHERE company_name=?",(company_name,)).fetchone()
        bal_val = bal["total_owed"] if bal else 0.0
        self._acct_balance_lbl.setText(f"Balance Owed: ${bal_val:,.2f}")

        # Collect all invoice UUIDs referenced in payments (to flag paid invoices)
        paid_ids = set()
        for ph in self.db.execute(
            "SELECT invoice_id FROM account_history WHERE company_name=? AND type='payment'",
            (company_name,)).fetchall():
            for uid in (ph["invoice_id"] or "").split(","):
                uid = uid.strip()
                if uid: paid_ids.add(uid)

        # Invoices for this account
        inv_rows = self.db.execute("""
            SELECT invoice_id, invoice_number, invoice_date, amount_cents FROM invoices
            WHERE (UPPER(company_name)=UPPER(?) OR UPPER(account_id)=UPPER(?)) AND is_estimate=0
            ORDER BY invoice_date DESC
        """, (company_name, company_name)).fetchall()

        # Payments from account_history (include id/payment_id/invoice_id for delete)
        pay_rows = self.db.execute("""
            SELECT id, entry_date, payment_number, amount, note, payment_id, invoice_id
            FROM account_history
            WHERE company_name=? AND type='payment'
            ORDER BY entry_date DESC, id DESC
        """, (company_name,)).fetchall()

        # Merge: build list of (date, sort_key, row_data)
        # row_data = (date, type_str, ref, invoice_amt, payment_amt, notes, is_paid, meta)
        # meta is None for invoice rows; dict with id/payment_id/invoice_id/amount for payment rows
        entries = []
        for r in inv_rows:
            is_paid = r["invoice_id"] in paid_ids
            entries.append((r["invoice_date"], 0,
                (r["invoice_date"], "Invoice", str(r["invoice_number"] or "-"),
                 f"${r['amount_cents']/100:,.2f}", "", "", is_paid, None)))
        for r in pay_rows:
            meta = {
                "id":         r["id"],
                "payment_id": r["payment_id"] or "",
                "invoice_id": r["invoice_id"] or "",
                "amount":     r["amount"],
            }
            entries.append((r["entry_date"], 1,
                (r["entry_date"], "Payment", r["payment_number"] or "-",
                 "", f"${r['amount']:,.2f}", r["note"] or "", False, meta)))

        entries.sort(key=lambda x: (x[0], x[1]), reverse=True)

        INV_BG   = QColor("#EFF6FF")   # light blue for invoices
        INV_PAID = QColor("#F0FDF4")   # light green for paid invoices
        PAY_BG   = QColor("#ECFDF5")   # green tint for payments
        PAY_FG   = QColor("#15803D")

        self._acct_hist_table.setRowCount(0)
        for _, _, row_data in entries:
            date_s, type_s, ref_s, inv_s, pay_s, note_s, is_paid, meta = row_data
            ri = self._acct_hist_table.rowCount(); self._acct_hist_table.insertRow(ri)
            for ci, val in enumerate([date_s, type_s, ref_s, inv_s, pay_s, note_s]):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter |
                    (Qt.AlignmentFlag.AlignRight if ci in (3,4) else Qt.AlignmentFlag.AlignLeft))
                if type_s == "Payment":
                    item.setBackground(PAY_BG)
                    if ci == 4: item.setForeground(PAY_FG)
                else:
                    item.setBackground(INV_PAID if is_paid else INV_BG)
                # Store payment metadata on column 0 for use by the context menu
                if ci == 0 and meta is not None:
                    item.setData(Qt.ItemDataRole.UserRole, meta)
                self._acct_hist_table.setItem(ri, ci, item)

    def _acct_hist_context_menu(self, pos):
        """Right-click menu on Customer History table - mark/unmark invoices paid, delete payments."""
        tbl = self._acct_hist_table
        idx = tbl.indexAt(pos)
        if not idx.isValid(): return

        row = idx.row()
        type_item = tbl.item(row, 1)
        if not type_item: return
        type_str = type_item.text()

        # â"€â"€ Payment row: offer Delete â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        if type_str == "Payment":
            first_item = tbl.item(row, 0)
            meta = first_item.data(Qt.ItemDataRole.UserRole) if first_item else None
            if not meta: return

            ref_item = tbl.item(row, 2)
            pay_num  = ref_item.text().strip() if ref_item else "-"
            amt      = meta["amount"]

            menu = QMenu(tbl)
            del_act = menu.addAction(f"Delete Payment {pay_num}")
            action = menu.exec(tbl.viewport().mapToGlobal(pos))
            if action is None: return

            reply = QMessageBox.question(
                self, "Delete Payment",
                f"Delete payment  {pay_num}  (${amt:,.2f})?\n\n"
                "The account balance will be restored and any invoices it was\n"
                "applied to will revert to unpaid.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes: return

            if not self._acct_names or self._acct_index < 0: return
            name = self._acct_names[self._acct_index]

            self.db.execute("DELETE FROM account_history WHERE id=?", (meta["id"],))
            self.db.execute(
                "UPDATE accounts SET total_owed=total_owed+?,updated_at=? WHERE company_name=?",
                (amt, now_iso(), name))
            if meta["payment_id"]:
                enqueue(self.db, "account_payment", "delete",
                        {"payment_id": meta["payment_id"], "company_name": name})
            self.db.commit()
            self._load_acct(name)
            return

        if type_str != "Invoice": return   # ignore any other row types

        ref_item = tbl.item(row, 2)   # invoice number string
        inv_item = tbl.item(row, 3)   # "$X,XXX.XX"
        if not ref_item: return
        inv_num_str = ref_item.text().strip()

        # Check background colour - if light-green it's already paid
        bg = tbl.item(row, 0).background().color() if tbl.item(row, 0) else None
        already_paid = (bg == QColor("#F0FDF4")) if bg else False

        menu = QMenu(tbl)
        if already_paid:
            unmark_act = menu.addAction("Unmark as Paid")
        else:
            mark_act = menu.addAction("Mark as Paid")

        action = menu.exec(tbl.viewport().mapToGlobal(pos))
        if action is None: return

        if not self._acct_names or self._acct_index < 0: return
        name = self._acct_names[self._acct_index]

        # Resolve invoice number -> UUID + amount
        try: inv_num_int = int(inv_num_str)
        except: QMessageBox.warning(self, "Error", f"Could not parse invoice number '{inv_num_str}'."); return

        r = self.db.execute(
            "SELECT invoice_id, amount_cents FROM invoices WHERE invoice_number=? AND "
            "(UPPER(company_name)=UPPER(?) OR UPPER(account_id)=UPPER(?))",
            (inv_num_int, name, name)).fetchone()
        if not r:
            QMessageBox.warning(self, "Not Found", f"Invoice #{inv_num_str} not found for this account.")
            return

        inv_uuid    = r["invoice_id"]
        inv_dollars = (r["amount_cents"] or 0) / 100

        if already_paid:
            # --- Unmark: remove the MARK-PAID entry that references this invoice
            rows_del = self.db.execute(
                "SELECT id, invoice_id FROM account_history "
                "WHERE company_name=? AND type='payment' AND payment_number='PAID'",
                (name,)).fetchall()
            deleted = False
            for del_row in rows_del:
                ids = [x.strip() for x in (del_row["invoice_id"] or "").split(",")]
                if inv_uuid in ids:
                    pid = del_row["payment_id"] if "payment_id" in del_row.keys() else ""
                    self.db.execute("DELETE FROM account_history WHERE id=?", (del_row["id"],))
                    # Restore balance
                    self.db.execute(
                        "UPDATE accounts SET total_owed=total_owed+?,updated_at=? WHERE company_name=?",
                        (inv_dollars, now_iso(), name))
                    if pid:
                        enqueue(self.db, "account_payment", "delete",
                                {"payment_id": pid, "company_name": name})
                    deleted = True
                    break
            if not deleted:
                QMessageBox.information(self, "Info",
                    "This invoice was marked paid via a regular payment - edit that payment to remove it.")
                return
            self.db.commit()
            self._load_acct(name)
        else:
            # --- Mark as paid
            reply = QMessageBox.question(
                self, "Mark as Paid",
                f"Mark Invoice #{inv_num_str} (${inv_dollars:,.2f}) as paid?\n\n"
                "This will reduce the account balance by the invoice amount.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes: return

            payment_id = str(uuid.uuid4())
            cust_row = self.db.execute(
                "SELECT customer_id FROM customers WHERE UPPER(company_name)=UPPER(?) LIMIT 1",
                (name,)).fetchone()
            customer_id = cust_row["customer_id"] if cust_row else ""
            today = datetime.today().strftime("%Y-%m-%d")
            self.db.execute(
                "INSERT INTO account_history(company_name,entry_date,type,amount,note,invoice_id,payment_number,payment_id) "
                "VALUES(?,?,?,?,?,?,?,?)",
                (name, today, "payment", inv_dollars, "", inv_uuid, "PAID", payment_id))
            self.db.execute(
                "UPDATE accounts SET total_owed=MAX(0,total_owed-?),updated_at=? WHERE company_name=?",
                (inv_dollars, now_iso(), name))
            enqueue(self.db, "account_payment", "upsert", {
                "payment_id": payment_id, "customer_id": customer_id,
                "company_name": name, "entry_date": today,
                "amount_cents": int(inv_dollars * 100), "note": "",
                "invoice_id": inv_uuid, "payment_number": "PAID",
            })
            self.db.commit()
            self._load_acct(name)

    def _acct_hist_double_click(self, idx):
        """Double-click an unpaid invoice row to open it in the invoice editor."""
        row = idx.row()
        type_item = self._acct_hist_table.item(row, 1)
        if not type_item or type_item.text() != "Invoice": return

        # Only open if not already paid (light-green bg means paid)
        bg = self._acct_hist_table.item(row, 0).background().color() if self._acct_hist_table.item(row, 0) else None
        if bg == QColor("#F0FDF4"): return   # already paid - ignore double-click

        ref_item = self._acct_hist_table.item(row, 2)
        if not ref_item: return
        try: inv_num_int = int(ref_item.text().strip())
        except: return

        if not self._acct_names or self._acct_index < 0: return
        name = self._acct_names[self._acct_index]

        r = self.db.execute(
            "SELECT invoice_id FROM invoices WHERE invoice_number=? AND "
            "(UPPER(company_name)=UPPER(?) OR UPPER(account_id)=UPPER(?))",
            (inv_num_int, name, name)).fetchone()
        if not r: return

        self.show_screen("estimate_entry", invoice_id=r["invoice_id"])

    def _acct_first(self):
        if self._acct_names: self._acct_index=0; self._acct_combo.setCurrentIndex(0)
    def _acct_last(self):
        if self._acct_names: self._acct_index=len(self._acct_names)-1; self._acct_combo.setCurrentIndex(self._acct_index)
    def _acct_prev(self):
        if self._acct_index>0: self._acct_index-=1; self._acct_combo.setCurrentIndex(self._acct_index)
    def _acct_next(self):
        if self._acct_index<len(self._acct_names)-1: self._acct_index+=1; self._acct_combo.setCurrentIndex(self._acct_index)

    def _new_acct_dialog(self): self._acct_edit_dialog(None)
    def _edit_acct_dialog(self):
        if not self._acct_names or self._acct_index<0: return
        self._acct_edit_dialog(self._acct_names[self._acct_index])

    def _acct_edit_dialog(self, existing_name):
        dlg = QDialog(self); dlg.setWindowTitle("Account" if existing_name else "New Account")
        dlg.setMinimumWidth(460); lay = QVBoxLayout(dlg); form = QFormLayout()
        fields = {}
        row = self.db.execute("SELECT * FROM accounts WHERE company_name=?",(existing_name,)).fetchone() if existing_name else None
        def fe(key, default=""):
            e = QLineEdit(str(row[key] if row and row[key] else default))
            fields[key]=e; return e
        form.addRow("Account Name:", fe("company_name")); form.addRow("Contact:", fe("contact_name"))
        form.addRow("Phone:", fe("phone")); form.addRow("Email:", fe("email"))
        form.addRow("Address:", fe("address1")); form.addRow("City:", fe("city"))
        form.addRow("State:", fe("state")); form.addRow("ZIP:", fe("zip"))
        lay.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lay.addWidget(bb)
        if dlg.exec() != QDialog.DialogCode.Accepted: return
        name = fields["company_name"].text().strip().upper()
        if not name: QMessageBox.warning(self,"Required","Account Name is required."); return
        self.db.execute("""
            INSERT INTO accounts(company_name,total_owed,updated_at,contact_name,phone,email,address1,city,state,zip)
            VALUES(?,0,?,?,?,?,?,?,?,?)
            ON CONFLICT(company_name) DO UPDATE SET
                contact_name=excluded.contact_name,phone=excluded.phone,email=excluded.email,
                address1=excluded.address1,city=excluded.city,state=excluded.state,zip=excluded.zip,updated_at=excluded.updated_at
        """,(name,now_iso(),fields["contact_name"].text(),fields["phone"].text(),fields["email"].text(),
             fields["address1"].text(),fields["city"].text(),fields["state"].text(),fields["zip"].text()))
        self.db.commit(); self._on_show_account_setup(name)

    def _delete_acct_action(self):
        if not self._acct_names or self._acct_index<0: return
        name = self._acct_names[self._acct_index]
        if QMessageBox.question(self,"Delete Account",f"Delete account '{name}'?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            self.db.execute("DELETE FROM accounts WHERE company_name=?",(name,))
            self.db.execute("DELETE FROM account_history WHERE company_name=?",(name,))
            self.db.commit(); self._on_show_account_setup()

    def _acct_post_payment(self):
        if not self._acct_names or self._acct_index < 0: return
        name = self._acct_names[self._acct_index]

        # Auto-generate next payment number
        row_max = self.db.execute(
            "SELECT MAX(CAST(REPLACE(payment_number,'PMT-','') AS INTEGER)) "
            "FROM account_history WHERE payment_number LIKE 'PMT-%'").fetchone()
        next_num = (row_max[0] or 0) + 1
        default_pmt = f"PMT-{next_num:04d}"

        dlg = QDialog(self); dlg.setWindowTitle("Post Payment"); dlg.setMinimumWidth(440)
        lay = QVBoxLayout(dlg)
        form = QFormLayout()
        pmt_num_e = QLineEdit(default_pmt)
        date_e    = QLineEdit(datetime.today().strftime("%Y-%m-%d"))
        amt_e     = QLineEdit()
        note_e    = QLineEdit()
        form.addRow("Payment #:", pmt_num_e)
        form.addRow("Date:",      date_e)
        form.addRow("Amount $:",  amt_e)
        form.addRow("Notes:",     note_e)
        lay.addLayout(form)
        info_lbl = QLabel("Payment is applied automatically to the oldest unpaid invoices first.")
        info_lbl.setWordWrap(True)
        info_lbl.setStyleSheet("color:#6B7280;font-size:10pt;padding:6px 0;")
        lay.addWidget(info_lbl)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lay.addWidget(bb)
        if dlg.exec() != QDialog.DialogCode.Accepted: return

        try:
            amt = float(amt_e.text().strip().replace(",","").replace("$",""))
        except:
            QMessageBox.warning(self,"Invalid","Enter a valid dollar amount."); return
        if amt <= 0:
            QMessageBox.warning(self,"Invalid","Amount must be greater than zero."); return

        pmt_num  = pmt_num_e.text().strip() or default_pmt
        date_str = date_e.text().strip() or datetime.today().strftime("%Y-%m-%d")
        notes    = note_e.text().strip()

        # â"€â"€ Build paid_ids set â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        paid_ids = set()
        for pr in self.db.execute(
                "SELECT invoice_id FROM account_history WHERE company_name=? AND type='payment'",
                (name,)).fetchall():
            for uid in (pr["invoice_id"] or "").split(","):
                u = uid.strip()
                if u: paid_ids.add(u)

        # â"€â"€ Accumulate prior partial amounts per invoice UUID â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        partial_applied: dict = {}   # uuid -> total dollars already partially applied
        for pr in self.db.execute(
                "SELECT partial_json FROM account_history WHERE company_name=? AND type='payment'",
                (name,)).fetchall():
            try:
                pj = json.loads(pr["partial_json"] or "{}")
                for iid, v in pj.items():
                    partial_applied[iid] = partial_applied.get(iid, 0.0) + float(v)
            except Exception: pass

        # â"€â"€ Get unpaid invoices oldest first, with effective remaining â"€â"€â"€â"€â"€â"€â"€â"€â"€
        unpaid = []
        for r in self.db.execute("""
            SELECT invoice_id, invoice_number, amount_cents FROM invoices
            WHERE (UPPER(company_name)=UPPER(?) OR UPPER(account_id)=UPPER(?))
              AND is_estimate=0
            ORDER BY invoice_date ASC, invoice_number ASC
        """, (name, name)).fetchall():
            if r["invoice_id"] in paid_ids: continue
            full_amt   = (r["amount_cents"] or 0) / 100.0
            already    = partial_applied.get(r["invoice_id"], 0.0)
            remaining  = round(full_amt - already, 2)
            if remaining > 0.005:
                unpaid.append({"uuid": r["invoice_id"], "num": r["invoice_number"],
                               "full": full_amt, "remaining": remaining,
                               "was_partial": already > 0.005})

        # â"€â"€ Auto-apply algorithm â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        left       = amt
        fully_paid = []    # list of invoice dicts that this payment fully covers
        partial    = None  # single invoice dict with extra keys: applied, shortfall
        for inv in unpaid:
            if left < 0.005: break
            if left >= inv["remaining"] - 0.005:       # fully covers remaining balance
                fully_paid.append(inv)
                left = round(left - inv["remaining"], 2)
            else:                                       # partial coverage
                partial = {**inv, "applied": left,
                           "shortfall": round(inv["remaining"] - left, 2)}
                left = 0
                break

        # â"€â"€ Confirmation dialog â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        lines = [f"<b>Payment of ${amt:,.2f} - breakdown:</b><br>"]
        for inv in fully_paid:
            tag = " (completes partial)" if inv["was_partial"] else ""
            lines.append(f"&nbsp;&nbsp;Invoice #{inv['num']}  "
                         f"(${inv['full']:,.2f})&nbsp;-&nbsp;<b>PAID IN FULL{tag}</b>")
        if partial:
            lines.append(f"&nbsp;&nbsp;! Invoice #{partial['num']}  "
                         f"(${partial['full']:,.2f})&nbsp;-&nbsp;"
                         f"${partial['applied']:,.2f} applied,&nbsp;"
                         f"<b>${partial['shortfall']:,.2f} still owed</b>&nbsp;(stays unpaid)")
        if not fully_paid and partial is None:
            if unpaid:
                lines.append("&nbsp;&nbsp;No invoices will be marked paid (amount is less than any single invoice).")
            else:
                lines.append("&nbsp;&nbsp;No open invoices - payment reduces balance only.")

        cdlg = QDialog(self); cdlg.setWindowTitle("Confirm Payment"); cdlg.setMinimumWidth(480)
        cl = QVBoxLayout(cdlg)
        lbl = QLabel("<br>".join(lines))
        lbl.setTextFormat(Qt.TextFormat.RichText); lbl.setWordWrap(True)
        cl.addWidget(lbl)
        cbb = QDialogButtonBox(QDialogButtonBox.StandardButton.Yes|QDialogButtonBox.StandardButton.No)
        cbb.accepted.connect(cdlg.accept); cbb.rejected.connect(cdlg.reject)
        cl.addWidget(cbb)
        if cdlg.exec() != QDialog.DialogCode.Accepted: return

        # â"€â"€ Build fields for DB insert â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        invoice_id_str = ",".join(inv["uuid"] for inv in fully_paid)

        new_partial_json = "{}"
        if partial:
            new_partial_json = json.dumps({partial["uuid"]: partial["applied"]})
            pn = f"[Partial ${partial['applied']:,.2f} toward Inv #{partial['num']} - ${partial['shortfall']:,.2f} remaining]"
            notes = f"{notes}  {pn}".strip() if notes else pn

        payment_id = str(uuid.uuid4())
        cust_row = self.db.execute(
            "SELECT customer_id FROM customers WHERE UPPER(company_name)=UPPER(?) LIMIT 1",
            (name,)).fetchone()
        customer_id = cust_row["customer_id"] if cust_row else ""

        self.db.execute(
            "UPDATE accounts SET total_owed=MAX(0,total_owed-?),updated_at=? WHERE company_name=?",
            (amt, now_iso(), name))
        self.db.execute(
            "INSERT INTO account_history"
            "(company_name,entry_date,type,amount,note,invoice_id,payment_number,payment_id,partial_json) "
            "VALUES(?,?,?,?,?,?,?,?,?)",
            (name, date_str, "payment", amt, notes, invoice_id_str, pmt_num, payment_id, new_partial_json))
        enqueue(self.db, "account_payment", "upsert", {
            "payment_id":   payment_id, "customer_id": customer_id,
            "company_name": name,       "entry_date":  date_str,
            "amount_cents": int(amt * 100), "note":    notes,
            "invoice_id":   invoice_id_str, "payment_number": pmt_num,
            "partial_json": new_partial_json,
        })
        self.db.commit()
        self._load_acct(name)

    def _acct_print_statement(self):
        """Print account statement as PDF - choose all history or outstanding only."""
        if not self._acct_names or self._acct_index < 0: return
        name = self._acct_names[self._acct_index]

        dlg = QDialog(self); dlg.setWindowTitle("Print Statement"); dlg.setMinimumWidth(320)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel(f"<b>Account:</b> {name}"))
        lay.addSpacing(8)
        all_rb  = QRadioButton("Full Customer History (all invoices + payments)")
        out_rb  = QRadioButton("Outstanding Balance Only (unpaid invoices)")
        all_rb.setChecked(True)
        lay.addWidget(all_rb); lay.addWidget(out_rb)
        lay.addSpacing(8)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lay.addWidget(bb)
        if dlg.exec() != QDialog.DialogCode.Accepted: return

        outstanding_only = out_rb.isChecked()
        self._generate_acct_statement_pdf(name, outstanding_only)

    def _generate_acct_statement_pdf(self, company_name, outstanding_only):
        try:
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.pagesizes import letter
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib import colors as rl_colors
            from reportlab.lib.units import inch
            import tempfile, os

            # Gather data
            acct = self.db.execute("SELECT * FROM accounts WHERE company_name=?",(company_name,)).fetchone()
            bal_val = (acct["total_owed"] if acct else 0.0) or 0.0

            paid_ids = set()
            for ph in self.db.execute(
                "SELECT invoice_id FROM account_history WHERE company_name=? AND type='payment'",
                (company_name,)).fetchall():
                for uid in (ph["invoice_id"] or "").split(","):
                    uid=uid.strip()
                    if uid: paid_ids.add(uid)

            inv_rows = self.db.execute("""
                SELECT invoice_id, invoice_number, invoice_date, amount_cents FROM invoices
                WHERE (UPPER(company_name)=UPPER(?) OR UPPER(account_id)=UPPER(?)) AND is_estimate=0
                ORDER BY invoice_date
            """, (company_name, company_name)).fetchall()

            pay_rows = self.db.execute("""
                SELECT entry_date, payment_number, amount, note FROM account_history
                WHERE company_name=? AND type='payment'
                ORDER BY entry_date, id
            """, (company_name,)).fetchall()

            # Build PDF
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False,
                                              dir=os.path.expanduser("~"))
            tmp_path = tmp.name; tmp.close()

            doc = SimpleDocTemplate(tmp_path, pagesize=letter,
                                    leftMargin=0.75*inch, rightMargin=0.75*inch,
                                    topMargin=0.75*inch, bottomMargin=0.75*inch)
            styles = getSampleStyleSheet()
            story = []

            story.append(Paragraph(f"<b>Account Statement - {company_name}</b>", styles["Title"]))
            story.append(Spacer(1, 6))
            if acct:
                story.append(Paragraph(f"Contact: {acct['contact_name'] or ''}  |  "
                                       f"Phone: {acct['phone'] or ''}  |  "
                                       f"Email: {acct['email'] or ''}", styles["Normal"]))
            story.append(Paragraph(f"Balance Owed: <b>${bal_val:,.2f}</b>", styles["Normal"]))
            story.append(Spacer(1, 12))

            HDR_BLUE = rl_colors.HexColor("#005B99")
            PAY_GRN  = rl_colors.HexColor("#DCFCE7")
            INV_BLUE = rl_colors.HexColor("#DBEAFE")
            INV_PAID = rl_colors.HexColor("#F0FDF4")

            if outstanding_only:
                story.append(Paragraph("<b>Outstanding Invoices</b>", styles["Heading2"]))
                tbl_data = [["Invoice #", "Date", "Amount"]]
                total_out = 0.0
                for r in inv_rows:
                    if r["invoice_id"] in paid_ids: continue
                    amt = r["amount_cents"] / 100.0
                    total_out += amt
                    tbl_data.append([str(r["invoice_number"] or "-"),
                                     r["invoice_date"],
                                     f"${amt:,.2f}"])
                tbl_data.append(["", "TOTAL OUTSTANDING", f"${total_out:,.2f}"])
                col_w = [1.2*inch, 2.0*inch, 1.5*inch]
                t = Table(tbl_data, colWidths=col_w, repeatRows=1)
                t.setStyle(TableStyle([
                    ("BACKGROUND", (0,0), (-1,0), HDR_BLUE),
                    ("TEXTCOLOR",  (0,0), (-1,0), rl_colors.white),
                    ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
                    ("GRID",       (0,0), (-1,-1), 0.5, rl_colors.grey),
                    ("BACKGROUND", (0,-1), (-1,-1), rl_colors.HexColor("#FEF9C3")),
                    ("FONTNAME",   (0,-1), (-1,-1), "Helvetica-Bold"),
                    ("ALIGN",      (2,0),  (-1,-1), "RIGHT"),
                ]))
                story.append(t)
            else:
                story.append(Paragraph("<b>Customer History</b>", styles["Heading2"]))
                tbl_data = [["Date", "Type", "Reference #", "Invoice", "Payment", "Notes"]]
                # Merge invoices + payments sorted by date
                entries = []
                for r in inv_rows:
                    entries.append((r["invoice_date"], 0,
                        ("Invoice", str(r["invoice_number"] or "-"),
                         f"${r['amount_cents']/100:,.2f}", "", "",
                         r["invoice_id"] in paid_ids)))
                for r in pay_rows:
                    entries.append((r["entry_date"], 1,
                        ("Payment", r["payment_number"] or "-",
                         "", f"${r['amount']:,.2f}", r["note"] or "", False)))
                entries.sort(key=lambda x: (x[0], x[1]))
                row_styles = []
                for idx, (date_s, _, rd) in enumerate(entries):
                    type_s, ref_s, inv_s, pay_s, note_s, is_paid = rd
                    tbl_data.append([date_s, type_s, ref_s, inv_s, pay_s, note_s])
                    ri = idx + 1
                    if type_s == "Payment":
                        row_styles.append(("BACKGROUND", (0,ri), (-1,ri), PAY_GRN))
                    elif is_paid:
                        row_styles.append(("BACKGROUND", (0,ri), (-1,ri), INV_PAID))
                    else:
                        row_styles.append(("BACKGROUND", (0,ri), (-1,ri), INV_BLUE))

                col_w = [0.85*inch, 0.75*inch, 1.0*inch, 0.85*inch, 0.85*inch, 2.2*inch]
                t = Table(tbl_data, colWidths=col_w, repeatRows=1)
                base_style = [
                    ("BACKGROUND", (0,0), (-1,0), HDR_BLUE),
                    ("TEXTCOLOR",  (0,0), (-1,0), rl_colors.white),
                    ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
                    ("GRID",       (0,0), (-1,-1), 0.5, rl_colors.grey),
                    ("ALIGN",      (3,0), (4,-1),  "RIGHT"),
                    ("FONTSIZE",   (0,0), (-1,-1), 8),
                ]
                t.setStyle(TableStyle(base_style + row_styles))
                story.append(t)

            doc.build(story)
            ps = get_printer_setting(self.db)
            viewer = PdfViewerDialog(tmp_path, ps, self)
            viewer.exec()
        except Exception as e:
            QMessageBox.critical(self, "Print Error", f"Could not generate statement:\n{e}")

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: REPORTS
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_reports_screen(self):
        w = QWidget(); self._screens["reports"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        lay.addWidget(self._make_header(show_back=True)); self._stack.addWidget(w)
        body = QWidget(); body_lay = QVBoxLayout(body); body_lay.setContentsMargins(16,16,16,16); body_lay.setSpacing(12)
        lay.addWidget(body)
        dr = QHBoxLayout(); dr.addWidget(QLabel("Begin:"))
        self._rpt_begin = QDateEdit(QDate(datetime.today().year, datetime.today().month, 1))
        self._rpt_begin.setCalendarPopup(True); self._rpt_begin.setDisplayFormat("yyyy-MM-dd")
        self._rpt_begin.setMaximumWidth(140); dr.addWidget(self._rpt_begin)
        dr.addWidget(QLabel("End:"))
        self._rpt_end = QDateEdit(QDate.currentDate())
        self._rpt_end.setCalendarPopup(True); self._rpt_end.setDisplayFormat("yyyy-MM-dd")
        self._rpt_end.setMaximumWidth(140); dr.addWidget(self._rpt_end)
        dr.addStretch(); body_lay.addLayout(dr)
        rpt_grid = QGridLayout()
        for i,(lbl_txt,cb) in enumerate([
            ("Month-To-Date",      lambda: self._run_report("mtd")),
            ("Daily",              lambda: self._run_report("daily")),
            ("Weekly",             lambda: self._run_report("weekly")),
            ("By Payment Type",    lambda: self._run_report("by_pay")),
            ("By Service",         lambda: self._run_report("by_svc")),
            ("Open Estimates",     lambda: self._run_report("estimates")),
            ("Account Balances",   lambda: self._run_report("balances")),
        ]):
            b = btn(lbl_txt,"primary"); b.setMinimumHeight(48); b.clicked.connect(cb)
            rpt_grid.addWidget(b, i//3, i%3)
        body_lay.addLayout(rpt_grid); body_lay.addStretch()
        bb = QWidget(); bb.setObjectName("darkBar"); bb.setStyleSheet("QWidget#darkBar { background: #37474F; }")
        bb_h = QHBoxLayout(bb); bb_h.setContentsMargins(8,6,8,6); bb_h.setSpacing(6)
        dl_b = btn("DOCUMENTS LIST","secondary"); dl_b.clicked.connect(lambda: self.show_screen("doc_list")); bb_h.addWidget(dl_b)
        bb_h.addStretch(); lay.addWidget(bb)

    def _run_report(self, rpt_type):
        begin_date = self._rpt_begin.date()
        end_date   = self._rpt_end.date()

        # Override date range based on report type
        if rpt_type == "mtd":
            begin = QDate(begin_date.year(), begin_date.month(), 1).toString("yyyy-MM-dd")
            import calendar
            last_day = calendar.monthrange(begin_date.year(), begin_date.month())[1]
            end = QDate(begin_date.year(), begin_date.month(), last_day).toString("yyyy-MM-dd")
        elif rpt_type == "weekly":
            begin = begin_date.toString("yyyy-MM-dd")
            end   = begin_date.addDays(7).toString("yyyy-MM-dd")
        elif rpt_type == "daily":
            begin = begin_date.toString("yyyy-MM-dd")
            end   = begin
        else:
            begin = begin_date.toString("yyyy-MM-dd")
            end   = end_date.toString("yyyy-MM-dd")

        try:
            all_rows = self.db.execute("""
                SELECT invoice_number,invoice_date,customer_name,first_name,last_name,
                       company_name,amount_cents,payment_method,is_estimate,test_result
                FROM invoices WHERE invoice_date>=? AND invoice_date<=? AND is_estimate=?
                ORDER BY invoice_date
            """,(begin,end,1 if rpt_type=="estimates" else 0)).fetchall()
        except Exception as e:
            QMessageBox.warning(self,"Error",str(e)); return

        title_str = "Outstanding Account Balances" if rpt_type == "balances" else f"Report - {rpt_type}  ({begin} to {end})"
        dlg = QDialog(self); dlg.setWindowTitle(title_str); dlg.resize(860,620)
        lay = QVBoxLayout(dlg)

        tbl = QTableWidget(0,0)
        tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        tbl.setAlternatingRowColors(True)
        total = 0.0

        if rpt_type == "by_pay":
            # Group by payment type: Date Range | Payment Type | # Transactions | Total
            tbl.setColumnCount(4)
            tbl.setHorizontalHeaderLabels(["Date Range","Payment Type","# Transactions","Total"])
            tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            from collections import defaultdict
            groups = defaultdict(lambda: {"count":0,"total":0.0})
            for row in all_rows:
                pt = (row["payment_method"] or "UNKNOWN").upper()
                groups[pt]["count"] += 1
                groups[pt]["total"] += row["amount_cents"]/100
                total += row["amount_cents"]/100
            for pt, g in sorted(groups.items()):
                r = tbl.rowCount(); tbl.insertRow(r)
                for col,val in enumerate([f"{begin} - {end}", pt, str(g["count"]), f"${g['total']:,.2f}"]):
                    tbl.setItem(r,col,QTableWidgetItem(val))

        elif rpt_type == "by_svc":
            # Group by service: Service Type | # Transactions | Total
            tbl.setColumnCount(3)
            tbl.setHorizontalHeaderLabels(["Service Type","# Transactions","Total"])
            tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            try:
                from collections import defaultdict
                svc_groups = defaultdict(lambda: {"count":0,"total":0.0})
                # Use il.price directly - mobile app will send correct per-line prices going forward
                lines = self.db.execute(
                    "SELECT il.service, il.price FROM invoice_lines il "
                    "JOIN invoices i ON il.invoice_id=i.invoice_id "
                    "WHERE i.invoice_date>=? AND i.invoice_date<=? AND i.is_estimate=0",
                    (begin, end)).fetchall()
                for ln in lines:
                    svc = (ln["service"] or "UNKNOWN").upper()
                    svc_groups[svc]["count"] += 1
                    svc_groups[svc]["total"] += float(ln["price"] or 0)
                if not svc_groups:
                    for row in all_rows:
                        svc_groups["SMOG CHECK"]["count"] += 1
                        svc_groups["SMOG CHECK"]["total"] += row["amount_cents"]/100
                for svc, g in sorted(svc_groups.items()):
                    r = tbl.rowCount(); tbl.insertRow(r)
                    total += g["total"]
                    for col,val in enumerate([svc, str(g["count"]), f"${g['total']:,.2f}"]):
                        tbl.setItem(r,col,QTableWidgetItem(val))
            except Exception as ex:
                for row in all_rows:
                    r = tbl.rowCount(); tbl.insertRow(r)
                    for col,val in enumerate(["SMOG CHECK","1",f"${row['amount_cents']/100:,.2f}"]):
                        tbl.setItem(r,col,QTableWidgetItem(val))
                    total += row["amount_cents"]/100

        elif rpt_type == "balances":
            rows = self.db.execute(
                "SELECT company_name, contact_name, phone, email, total_owed "
                "FROM accounts WHERE total_owed > 0 ORDER BY total_owed DESC"
            ).fetchall()
            tbl.setColumnCount(5)
            tbl.setHorizontalHeaderLabels(["Company", "Contact", "Phone", "Email", "Balance Owed"])
            tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
            for row in rows:
                r = tbl.rowCount(); tbl.insertRow(r)
                total += row["total_owed"]
                for col, val in enumerate([
                    row["company_name"] or "",
                    row["contact_name"] or "",
                    row["phone"] or "",
                    row["email"] or "",
                    f"${row['total_owed']:,.2f}",
                ]):
                    item = QTableWidgetItem(val)
                    if col == 4:
                        item.setForeground(QColor(RED))
                        item.setFont(QFont("", -1, QFont.Weight.Bold))
                    tbl.setItem(r, col, item)

        else:
            # Standard report: #, Date, Customer, Amount, Payment, Result
            tbl.setColumnCount(6)
            tbl.setHorizontalHeaderLabels(["#","Date","Customer","Amount","Payment","Result"])
            tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
            for row in all_rows:
                cname = row["customer_name"] or f"{row['first_name']} {row['last_name']}".strip()
                r = tbl.rowCount(); tbl.insertRow(r)
                amt = row["amount_cents"]/100; total += amt
                result = (row["test_result"] or "").upper()
                for col,val in enumerate([str(row["invoice_number"] or "-"),row["invoice_date"],cname,
                                          f"${amt:,.2f}",(row["payment_method"] or ""),result]):
                    item = QTableWidgetItem(val)
                    if col == 5:
                        if result == "PASS":   item.setForeground(QColor(GREEN))
                        elif result in ("FAIL","RETEST"): item.setForeground(QColor(RED))
                    tbl.setItem(r,col,item)

        lay.addWidget(tbl)
        if rpt_type == "balances":
            summary = QLabel(f"Accounts with Outstanding Balances: {tbl.rowCount()}   |   Total Owed: ${total:,.2f}")
        else:
            summary = QLabel(f"Period: {begin} -> {end}   |   Records: {tbl.rowCount()}   |   Total: ${total:,.2f}")
        summary.setStyleSheet(f"color:{PRIMARY}; font-weight:bold; font-size:11pt;"); lay.addWidget(summary)

        btn_row = QHBoxLayout()
        pr_b = QPushButton("Print Report"); pr_b.setObjectName("primary")
        def _print_rpt():
            from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            pdlg = QPrintDialog(printer, dlg)
            if pdlg.exec() == QDialog.DialogCode.Accepted:
                from PyQt6.QtGui import QTextDocument
                html = f"<h3>{dlg.windowTitle()}</h3><table border='1' cellpadding='4' style='border-collapse:collapse;width:100%'>"
                hdrs = [tbl.horizontalHeaderItem(c).text() for c in range(tbl.columnCount())]
                html += "<tr>" + "".join(f"<th style='background:#005B99;color:white'>{h}</th>" for h in hdrs) + "</tr>"
                for r in range(tbl.rowCount()):
                    html += "<tr>" + "".join(f"<td>{tbl.item(r,c).text() if tbl.item(r,c) else ''}</td>" for c in range(tbl.columnCount())) + "</tr>"
                html += f"</table><p><b>{summary.text()}</b></p>"
                doc = QTextDocument(); doc.setHtml(html)
                doc.print(printer)
        pr_b.clicked.connect(_print_rpt); btn_row.addWidget(pr_b)
        cl_b = QPushButton("Close"); cl_b.setObjectName("secondary")
        cl_b.clicked.connect(dlg.reject); btn_row.addWidget(cl_b)
        btn_row.addStretch(); lay.addLayout(btn_row)
        dlg.exec()

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: SETTINGS
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_settings_screen(self):
        w = QWidget(); self._screens["settings"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        lay.addWidget(self._make_header(show_back=True)); self._stack.addWidget(w)
        tabs = QTabWidget(); lay.addWidget(tabs)
        bb = QWidget(); bb.setObjectName("darkBar"); bb.setStyleSheet("QWidget#darkBar { background: #37474F; }")
        bb_h = QHBoxLayout(bb); bb_h.setContentsMargins(8,6,8,6)
        dl_b = btn("DOCUMENTS LIST","secondary"); dl_b.clicked.connect(lambda: self.show_screen("doc_list")); bb_h.addWidget(dl_b)
        bb_h.addStretch(); lay.addWidget(bb)

        # â"€â"€ Tab 1: Sync/Account â"€â"€
        t1 = QWidget(); t1_lay = QVBoxLayout(t1); t1_lay.setContentsMargins(16,16,16,16)
        saved = load_creds()
        t1_lay.addWidget(QLabel("Sync / Account Settings", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        form1 = QFormLayout()
        self._s_user = QLineEdit(saved.get("username","")); form1.addRow("Username:", self._s_user)
        self._s_pass = QLineEdit(saved.get("password","")); self._s_pass.setEchoMode(QLineEdit.EchoMode.Password)
        form1.addRow("Password:", self._s_pass)
        t1_lay.addLayout(form1)
        self._s_err = QLabel(""); self._s_err.setStyleSheet("color:red;"); t1_lay.addWidget(self._s_err)
        btn_h = QHBoxLayout()
        si_b = btn("Sign In","primary"); si_b.clicked.connect(self._settings_signin); btn_h.addWidget(si_b)
        so_b = btn("Sign Out","danger");  so_b.clicked.connect(self._settings_signout); btn_h.addWidget(so_b)
        def _do_fp():
            threading.Thread(
                target=SYNC.force_pull_from_zero,
                kwargs={"notify_cb": lambda msg: self._fp_signal.emit(msg)},
                daemon=True
            ).start()
        fp_b = btn("Force Re-pull","secondary"); fp_b.clicked.connect(_do_fp); btn_h.addWidget(fp_b)
        btn_h.addStretch(); t1_lay.addLayout(btn_h)
        t1_lay.addSpacing(16)
        manage_sub_b = btn("Manage Subscription", "primary")
        manage_sub_b.clicked.connect(self._settings_manage_subscription)
        t1_lay.addWidget(manage_sub_b)
        t1_lay.addSpacing(16)
        t1_lay.addWidget(QLabel("-- Danger Zone --", font=QFont("Segoe UI",10,QFont.Weight.Bold)))
        clr_b = btn("Clear Local Database","danger"); clr_b.clicked.connect(self._clear_local_db); t1_lay.addWidget(clr_b)
        t1_lay.addStretch(); tabs.addTab(t1,"Sync / Account")

        # â"€â"€ Tab 2: Business â"€â"€
        t2 = QWidget(); t2_lay = QVBoxLayout(t2); t2_lay.setContentsMargins(16,16,16,16)
        biz = get_business_settings(self.db)
        form2 = QFormLayout()
        self._biz = {}
        for key,lbl_txt in [("name","Business Name"),("address_line1","Address 1"),("address_line2","Address 2"),
                             ("phone","Phone"),("email","Email"),("ard","ARD #"),("card_fee","Card Fee $")]:
            e = QLineEdit(str(biz.get(key,"")))
            self._biz[key]=e; form2.addRow(f"{lbl_txt}:",e)
        self._biz_notice = QTextEdit(biz.get("invoice_notice","")); self._biz_notice.setMaximumHeight(90)
        form2.addRow("Invoice Notice:", self._biz_notice)
        logo_h = QHBoxLayout(); logo_h.addWidget(QLabel("Logo Path:"))
        self._biz_logo = QLineEdit(biz.get("logo_path","")); logo_h.addWidget(self._biz_logo)
        logo_btn = btn("Browse","secondary"); logo_btn.clicked.connect(self._browse_logo); logo_h.addWidget(logo_btn)
        t2_lay.addLayout(form2); t2_lay.addLayout(logo_h)
        save_b = btn("Save Business Info","primary"); save_b.clicked.connect(self._save_biz); t2_lay.addWidget(save_b)
        t2_lay.addStretch(); tabs.addTab(t2,"Business Info")

        # â"€â"€ Tab 3: Printer â"€â"€
        t3 = QWidget(); t3_lay = QVBoxLayout(t3); t3_lay.setContentsMargins(16,16,16,16)
        ps = get_printer_setting(self.db)
        self._pr_pdf = QRadioButton("Save PDF only"); self._pr_printer = QRadioButton("Send to printer")
        (self._pr_printer if ps.get("mode")=="printer" else self._pr_pdf).setChecked(True)
        t3_lay.addWidget(self._pr_pdf); t3_lay.addWidget(self._pr_printer)
        form3 = QFormLayout()
        self._pr_name = QComboBox()
        # Populate with system printers via Qt (always works in the built exe)
        try:
            from PyQt6.QtPrintSupport import QPrinterInfo
            _printers = [p.printerName() for p in QPrinterInfo.availablePrinters()]
        except Exception:
            _printers = []
        # Fallback: win32print
        if not _printers and _WIN32_PRINT:
            try:
                _printers = [p[2] for p in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL|win32print.PRINTER_ENUM_CONNECTIONS)]
            except Exception: pass
        # Fallback: PowerShell (Windows) or lpstat (Mac/Linux)
        if not _printers:
            try:
                import subprocess as _sp
                if sys.platform == "win32":
                    _r = _sp.run(["powershell","-Command",
                                  "Get-Printer | Select-Object -ExpandProperty Name"],
                                 capture_output=True,text=True,timeout=8)
                    _printers = [p.strip() for p in _r.stdout.strip().splitlines() if p.strip()]
                else:
                    _r = _sp.run(["lpstat","-a"], capture_output=True, text=True, timeout=8)
                    _printers = [l.split()[0] for l in _r.stdout.strip().splitlines() if l.strip()]
            except Exception: pass
        # Always add PDF option on Mac
        if sys.platform == "darwin" and "Save as PDF" not in _printers:
            _printers.insert(0, "Save as PDF")
        if _printers:
            self._pr_name.addItems(_printers)
            saved_pr = ps.get("printer_name","")
            if saved_pr in _printers:
                self._pr_name.setCurrentText(saved_pr)
        else:
            self._pr_name.setEditable(True)   # last resort: let user type
        form3.addRow("Printer:", self._pr_name)
        self._pr_copies = QSpinBox(); self._pr_copies.setRange(1,10); self._pr_copies.setValue(int(ps.get("copies",2)))
        form3.addRow("Copies:", self._pr_copies)
        self._pr_auto = QCheckBox("Auto-print after issuing invoice"); self._pr_auto.setChecked(bool(ps.get("auto_print")))
        t3_lay.addLayout(form3); t3_lay.addWidget(self._pr_auto)
        save_pr_b = btn("Save Printer Settings","primary"); save_pr_b.clicked.connect(self._save_printer); t3_lay.addWidget(save_pr_b)
        t3_lay.addStretch(); tabs.addTab(t3,"Printer")

        # ── Tab 4: Services ──
        t4 = QWidget(); t4_lay = QVBoxLayout(t4); t4_lay.setContentsMargins(16,16,16,16)
        t4_lay.addWidget(QLabel("Services & Prices", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        t4_lay.addWidget(QLabel("Add or edit the services shown in the invoice dropdown."))
        self._svc_table = QTableWidget(0, 3)
        self._svc_table.setHorizontalHeaderLabels(["Service Name", "Price ($)", "Cert Fee ($)"])
        self._svc_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._svc_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._svc_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        svcs = get_services(self.db)
        for svc_name, svc_data in svcs.items():
            row = self._svc_table.rowCount(); self._svc_table.insertRow(row)
            self._svc_table.setItem(row, 0, QTableWidgetItem(svc_name))
            self._svc_table.setItem(row, 1, QTableWidgetItem(str(svc_data.get("price", 0.0))))
            self._svc_table.setItem(row, 2, QTableWidgetItem(str(svc_data.get("cert_fee", 0.0))))
        t4_lay.addWidget(self._svc_table)
        svc_btn_h = QHBoxLayout()
        add_svc_b = btn("Add Service", "secondary"); add_svc_b.clicked.connect(self._svc_add_row); svc_btn_h.addWidget(add_svc_b)
        del_svc_b = btn("Delete Selected", "danger"); del_svc_b.clicked.connect(self._svc_del_row); svc_btn_h.addWidget(del_svc_b)
        svc_btn_h.addStretch(); t4_lay.addLayout(svc_btn_h)
        save_svc_b = btn("Save Services", "primary"); save_svc_b.clicked.connect(self._save_services); t4_lay.addWidget(save_svc_b)
        t4_lay.addStretch(); tabs.addTab(t4, "Services")

    def _svc_add_row(self):
        row = self._svc_table.rowCount(); self._svc_table.insertRow(row)
        self._svc_table.setItem(row, 0, QTableWidgetItem("New Service"))
        self._svc_table.setItem(row, 1, QTableWidgetItem("0.00"))
        self._svc_table.setItem(row, 2, QTableWidgetItem("0.00"))
        self._svc_table.editItem(self._svc_table.item(row, 0))

    def _svc_del_row(self):
        row = self._svc_table.currentRow()
        if row >= 0: self._svc_table.removeRow(row)

    def _save_services(self):
        svcs = {}
        for row in range(self._svc_table.rowCount()):
            name = (self._svc_table.item(row, 0) or QTableWidgetItem("")).text().strip()
            if not name: continue
            try: price = float((self._svc_table.item(row, 1) or QTableWidgetItem("0")).text())
            except: price = 0.0
            try: cert_fee = float((self._svc_table.item(row, 2) or QTableWidgetItem("0")).text())
            except: cert_fee = 0.0
            svcs[name] = {"price": price, "cert_fee": cert_fee}
        set_setting(self.db, "services", json.dumps(svcs))
        # Refresh the invoice form service dropdown immediately
        svc_names = list(svcs.keys())
        self._f_svc.clear(); self._f_svc.addItems(svc_names)
        QMessageBox.information(self, "Saved", "Services saved successfully.")

    def _reload_svc_table(self):
        svcs = get_services(self.db)
        self._svc_table.setRowCount(0)
        for svc_name, svc_data in svcs.items():
            row = self._svc_table.rowCount(); self._svc_table.insertRow(row)
            self._svc_table.setItem(row, 0, QTableWidgetItem(svc_name))
            self._svc_table.setItem(row, 1, QTableWidgetItem(str(svc_data.get("price", 0.0))))
            self._svc_table.setItem(row, 2, QTableWidgetItem(str(svc_data.get("cert_fee", 0.0))))

    def _on_show_settings(self):
        self._reload_svc_table()

    def _settings_signin(self):
        u = self._s_user.text().strip(); p = self._s_pass.text().strip()
        if not u or not p: self._s_err.setText("Enter username and password."); return
        self._s_err.setText("Signing in..."); QApplication.processEvents()
        try:
            token,company_id,company_name = api_login(u,p)
            save_creds({"username":u,"password":p,"token":token,"company_id":company_id,"company_name":company_name})
            self._s_err.setText(""); QMessageBox.information(self,"Signed In",f"Signed in as {u}")
        except Exception as e: self._s_err.setText(str(e))

    def _settings_signout(self):
        if QMessageBox.question(self,"Sign Out","Sign out?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            save_creds({})

    def _settings_manage_subscription(self):
        try:
            data = json.dumps({}).encode()
            req = urllib.request.Request(
                f"{API_BASE}/v1/subscription/portal",
                data=data,
                headers={**_hdrs(), "Content-Type": "application/json"},
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                result = json.loads(resp.read())
            url = result.get("portal_url", "")
            msg = result.get("message", "")
            if url:
                import webbrowser
                webbrowser.open(url)
            elif msg:
                QMessageBox.information(self, "Subscription", msg)
            else:
                QMessageBox.warning(self, "Error", "No portal URL returned from server.")
        except Exception as e:
            QMessageBox.warning(self, "Manage Subscription Error", str(e))

    def _clear_local_db(self):
        if QMessageBox.question(self,"Clear Database","This will erase ALL local data. Continue?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            for t in ("invoices","invoice_lines","customers","vehicles","outbox"):
                self.db.execute(f"DELETE FROM {t}")
            self.db.commit(); set_last_seq(self.db,0)
            QMessageBox.information(self,"Cleared","Local database cleared.")

    def _browse_logo(self):
        path,_ = QFileDialog.getOpenFileName(self,"Select Logo","","Images (*.png *.jpg *.ico *.bmp)")
        if path: self._biz_logo.setText(path)

    def _save_biz(self):
        biz = {k: e.text() if isinstance(e,QLineEdit) else e.toPlainText() for k,e in self._biz.items()}
        biz["invoice_notice"] = self._biz_notice.toPlainText()
        biz["logo_path"] = self._biz_logo.text()
        try: biz["card_fee"] = float(biz.get("card_fee",5.0))
        except: biz["card_fee"] = 5.0
        set_setting(self.db,"business",json.dumps(biz))
        # Sync business info to mobile devices
        enqueue(self.db,"company_settings","upsert",{
            "co_name":              biz.get("name",""),
            "co_addr":              biz.get("address_line1",""),
            "co_city":              biz.get("address_line2",""),
            "co_phone":             biz.get("phone",""),
            "co_email":             biz.get("email",""),
            "co_ard":               biz.get("ard",""),
            "invoice_notice":       biz.get("invoice_notice",""),
            "card_surcharge_value": str(biz.get("card_fee","")),
            "card_surcharge_type":  biz.get("card_surcharge_type","percent"),
        })
        QMessageBox.information(self,"Saved","Business info saved and will sync to mobile.")

    def _save_printer(self):
        ps = {"mode":"printer" if self._pr_printer.isChecked() else "pdf",
              "printer_name":self._pr_name.currentText(),"copies":self._pr_copies.value(),
              "auto_print":self._pr_auto.isChecked()}
        set_setting(self.db,"printer_setting",json.dumps(ps))
        QMessageBox.information(self,"Saved","Printer settings saved.")

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: CUSTOMERS
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_customers_screen(self):
        w = QWidget(); self._screens["customers"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        lay.addWidget(self._make_header(show_back=True)); self._stack.addWidget(w)
        sb = QWidget(); sb_h = QHBoxLayout(sb); sb_h.setContentsMargins(8,4,8,4)
        sb_h.addWidget(QLabel("Search:"))
        self._cust_search = QLineEdit(); self._cust_search.setMaximumWidth(260)
        self._cust_search.textChanged.connect(self._refresh_customers); sb_h.addWidget(self._cust_search)
        sb_h.addStretch(); lay.addWidget(sb)
        cols = ["Name","Company","Phone","Email","City"]
        self._cust_table = QTableWidget(0,len(cols)); self._cust_table.setHorizontalHeaderLabels(cols)
        self._cust_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._cust_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._cust_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._register_table("customers", self._cust_table)
        self._cust_table.doubleClicked.connect(self._cust_view)
        lay.addWidget(self._cust_table)
        bb = QWidget(); bb.setObjectName("darkBar"); bb.setStyleSheet("QWidget#darkBar { background: #37474F; }"); bb_h = QHBoxLayout(bb)
        bb_h.setContentsMargins(8,6,8,6); bb_h.setSpacing(6)
        for text,cb in [("View",self._cust_view),("Edit",self._cust_edit),("New Customer",self._cust_new)]:
            b = btn(text,"primary" if text!="New Customer" else "success"); b.clicked.connect(cb); bb_h.addWidget(b)
        del_b = btn("Delete","danger"); del_b.clicked.connect(self._cust_delete); bb_h.addWidget(del_b)
        dl_b = btn("DOCUMENTS LIST","secondary"); dl_b.clicked.connect(lambda: self.show_screen("doc_list")); bb_h.addWidget(dl_b)
        bb_h.addStretch(); lay.addWidget(bb)
        self._cust_ids = []
        self._cust_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._cust_table.customContextMenuRequested.connect(self._cust_context_menu)

    def _on_show_customers(self): self._refresh_customers()

    def _refresh_customers(self):
        q = self._cust_search.text().strip().lower()
        rows = self.db.execute("SELECT * FROM customers ORDER BY last_name,first_name,company_name").fetchall()
        self._cust_table.setRowCount(0); self._cust_ids=[]
        for row in rows:
            name = f"{row['first_name']} {row['last_name']}".strip() or row["company_name"]
            co = row["company_name"] or ""; phone=row["phone"] or ""; email=row["email"] or ""; city=row["city"] or ""
            if q and q not in " ".join([name,co,phone,email]).lower(): continue
            r = self._cust_table.rowCount(); self._cust_table.insertRow(r); self._cust_ids.append(row["customer_id"])
            for col,val in enumerate([name,co,phone,email,city]):
                self._cust_table.setItem(r,col,QTableWidgetItem(val))

    def _cust_selected_id(self):
        r = self._cust_table.currentRow()
        return self._cust_ids[r] if 0<=r<len(self._cust_ids) else None

    def _cust_view(self):
        cid = self._cust_selected_id()
        if not cid: return
        cust = self.db.execute("SELECT * FROM customers WHERE customer_id=?",(cid,)).fetchone()
        if not cust: return
        invs = self.db.execute("SELECT invoice_number,invoice_date,amount_cents,payment_method,is_estimate FROM invoices WHERE customer_id=? ORDER BY invoice_date DESC",(cid,)).fetchall()
        dlg = QDialog(self); dlg.setWindowTitle(f"Customer - {cust['first_name']} {cust['last_name']}"); dlg.resize(600,450)
        lay = QVBoxLayout(dlg)
        info = QLabel(f"<b>{cust['first_name']} {cust['last_name']}</b>  {cust['company_name']}<br>"
                      f"Phone: {cust['phone']}  Email: {cust['email']}<br>"
                      f"{cust['address']} {cust['city']} {cust['state']} {cust['zip']}")
        info.setTextFormat(Qt.TextFormat.RichText); lay.addWidget(info)
        tbl = QTableWidget(0,5); tbl.setHorizontalHeaderLabels(["#","Date","Type","Amount","Payment"])
        tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        total = 0.0
        for row in invs:
            r=tbl.rowCount(); tbl.insertRow(r); amt=row["amount_cents"]/100; total+=amt
            for col,val in enumerate([str(row["invoice_number"] or "-"),row["invoice_date"],"EST" if row["is_estimate"] else "INV",f"${amt:,.2f}",row["payment_method"] or ""]):
                tbl.setItem(r,col,QTableWidgetItem(val))
        lay.addWidget(tbl)
        lay.addWidget(QLabel(f"Total Invoiced: ${total:,.2f}", font=QFont("Segoe UI",11,QFont.Weight.Bold)))
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        bb.rejected.connect(dlg.reject); lay.addWidget(bb); dlg.exec()

    def _cust_edit(self):
        cid = self._cust_selected_id()
        if not cid: return
        cust = self.db.execute("SELECT * FROM customers WHERE customer_id=?",(cid,)).fetchone()
        if not cust: return
        dlg = QDialog(self); dlg.setWindowTitle("Edit Customer"); lay=QVBoxLayout(dlg); form=QFormLayout()
        fields={}
        for key,lbl_txt,val in [("first_name","First",cust["first_name"]),("last_name","Last",cust["last_name"]),
                                 ("company_name","Company",cust["company_name"]),("phone","Phone",cust["phone"]),
                                 ("email","Email",cust["email"]),("address","Address",cust["address"]),
                                 ("city","City",cust["city"]),("state","State",cust["state"]),("zip","ZIP",cust["zip"])]:
            e=QLineEdit(val or ""); fields[key]=e; form.addRow(f"{lbl_txt}:",e)
        disc_val = str(cust["discount_percent"] or "").rstrip("0").rstrip(".") if cust["discount_percent"] else ""
        disc_e = QLineEdit(disc_val); fields["discount_percent"] = disc_e
        disc_type_cb = QComboBox(); disc_type_cb.addItems(["%  (Percent)", "$  (Flat Amount)"])
        saved_type = (cust["discount_type"] or "PERCENT").upper()
        disc_type_cb.setCurrentIndex(0 if saved_type == "PERCENT" else 1)
        disc_row = QHBoxLayout(); disc_row.addWidget(disc_e); disc_row.addWidget(disc_type_cb)
        disc_widget = QWidget(); disc_widget.setLayout(disc_row)
        form.addRow("Discount:", disc_widget)
        lay.addLayout(form)
        bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lay.addWidget(bb)
        if dlg.exec()!=QDialog.DialogCode.Accepted: return
        disc = float(fields["discount_percent"].text().strip() or 0)
        disc_type = "PERCENT" if disc_type_cb.currentIndex() == 0 else "FLAT"
        upsert_customer(self.db,fields["first_name"].text(),fields["last_name"].text(),fields["company_name"].text(),
                        phone=format_phone(fields["phone"].text()),email=fields["email"].text(),address=fields["address"].text(),
                        city=fields["city"].text(),state=fields["state"].text(),zip_=fields["zip"].text(),
                        customer_id=cid, discount_percent=disc, discount_type=disc_type)
        enqueue(self.db, "customer", "upsert", {
            "customer_id": cid,
            "first_name": fields["first_name"].text(), "last_name": fields["last_name"].text(),
            "company_name": fields["company_name"].text(), "phone": format_phone(fields["phone"].text()),
            "email": fields["email"].text(), "address": fields["address"].text(),
            "city": fields["city"].text(), "state": fields["state"].text(), "zip": fields["zip"].text(),
            "discount_percent": disc, "discount_type": disc_type,
        })
        self._refresh_customers()

    def _cust_context_menu(self, pos):
        cid = self._cust_selected_id()
        if not cid: return
        menu = QMenu(self)
        menu.addAction("View",   self._cust_view)
        menu.addAction("Edit",   self._cust_edit)
        menu.addSeparator()
        menu.addAction("Delete...", self._cust_delete)
        menu.exec(self._cust_table.viewport().mapToGlobal(pos))

    def _cust_delete(self):
        cid = self._cust_selected_id()
        if not cid: return
        cust = self.db.execute("SELECT * FROM customers WHERE customer_id=?", (cid,)).fetchone()
        if not cust: return
        name = f"{cust['first_name']} {cust['last_name']}".strip() or cust["company_name"] or cid
        inv_count = self.db.execute(
            "SELECT COUNT(*) FROM invoices WHERE customer_id=?", (cid,)
        ).fetchone()[0]
        extra = (f"\n\nWarning: this customer has {inv_count} invoice(s) that will also be deleted."
                 if inv_count else "")
        if QMessageBox.question(
            self, "Delete Customer",
            f"Permanently delete customer '{name}'?{extra}\n\nThis cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) != QMessageBox.StandardButton.Yes:
            return
        if requests:
            try:
                r = requests.delete(f"{API_BASE}/v1/customers/{cid}", headers=_hdrs(), timeout=15)
                r.raise_for_status()
            except Exception as e:
                QMessageBox.warning(self, "Server", f"Server delete failed: {e}\nDeleted locally only.")
        self.db.execute(
            "DELETE FROM invoice_lines WHERE invoice_id IN (SELECT invoice_id FROM invoices WHERE customer_id=?)", (cid,))
        self.db.execute("DELETE FROM invoices WHERE customer_id=?", (cid,))
        self.db.execute("DELETE FROM vehicles WHERE customer_id=?", (cid,))
        self.db.execute("DELETE FROM customers WHERE customer_id=?", (cid,))
        self.db.commit()
        enqueue(self.db, "customer", "delete", {"customer_id": cid})
        self._refresh_customers()


    def _cust_new(self):
        dlg = QDialog(self); dlg.setWindowTitle("New Customer"); lay=QVBoxLayout(dlg); form=QFormLayout()
        fields={}
        for key,lbl_txt in [("first_name","First"),("last_name","Last"),("company_name","Company"),
                             ("phone","Phone"),("email","Email"),("address","Address"),("city","City"),("state","State"),("zip","ZIP")]:
            e=QLineEdit(); fields[key]=e; form.addRow(f"{lbl_txt}:",e)
        disc_e = QLineEdit(); fields["discount_percent"] = disc_e
        disc_type_cb = QComboBox(); disc_type_cb.addItems(["%  (Percent)", "$  (Flat Amount)"])
        disc_row2 = QHBoxLayout(); disc_row2.addWidget(disc_e); disc_row2.addWidget(disc_type_cb)
        disc_widget2 = QWidget(); disc_widget2.setLayout(disc_row2)
        form.addRow("Discount:", disc_widget2)
        lay.addLayout(form)
        bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lay.addWidget(bb)
        if dlg.exec()!=QDialog.DialogCode.Accepted: return
        disc = float(fields["discount_percent"].text().strip() or 0)
        disc_type = "PERCENT" if disc_type_cb.currentIndex() == 0 else "FLAT"
        cid = upsert_customer(self.db,fields["first_name"].text(),fields["last_name"].text(),fields["company_name"].text(),
                        phone=format_phone(fields["phone"].text()),email=fields["email"].text(),address=fields["address"].text(),
                        city=fields["city"].text(),state=fields["state"].text(),zip_=fields["zip"].text(),
                        discount_percent=disc, discount_type=disc_type)
        enqueue(self.db, "customer", "upsert", {
            "customer_id": cid,
            "first_name": fields["first_name"].text(), "last_name": fields["last_name"].text(),
            "company_name": fields["company_name"].text(), "phone": format_phone(fields["phone"].text()),
            "email": fields["email"].text(), "address": fields["address"].text(),
            "city": fields["city"].text(), "state": fields["state"].text(), "zip": fields["zip"].text(),
            "discount_percent": disc, "discount_type": disc_type,
        })
        self._refresh_customers()


    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: ADMIN BACKEND  (master account only)
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _master_api(self, method, path, **kwargs):
        """Make a master-authenticated API call."""
        creds   = load_creds()
        headers = {"x-username": creds.get("username",""), "x-password": creds.get("password","")}
        r = getattr(requests, method)(f"{API_BASE}{path}", headers=headers, timeout=15, **kwargs)
        r.raise_for_status()
        return r.json()

    def _build_admin_screen(self):
        w = QWidget(); self._screens["admin"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        lay.addWidget(self._make_header(show_back=True))
        self._stack.addWidget(w)

        body = QWidget(); body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(16,16,16,16); body_lay.setSpacing(12)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setWidget(body)
        lay.addWidget(scroll)

        # Title row
        title_h = QHBoxLayout()
        tl = QLabel("MASTER DASHBOARD")
        tl.setStyleSheet(f"color:{PRIMARY}; font-size:16pt; font-weight:bold;")
        title_h.addWidget(tl)
        ref_b = btn("Refresh","secondary"); ref_b.clicked.connect(self._admin_refresh)
        title_h.addWidget(ref_b)
        new_b = btn("+ New Account","success"); new_b.clicked.connect(self._admin_create_account)
        title_h.addWidget(new_b); title_h.addStretch(); body_lay.addLayout(title_h)

        # Stats cards
        stats_h = QHBoxLayout(); stats_h.setSpacing(12)
        self._adm_stat_co  = self._admin_stat_card(stats_h, "Companies",      "-", PRIMARY)
        self._adm_stat_inv = self._admin_stat_card(stats_h, "Total Invoices",  "-", GREEN)
        self._adm_stat_ex  = self._admin_stat_card(stats_h, "Exempt Accounts", "-", "#FFA500")
        self._adm_stat_sus = self._admin_stat_card(stats_h, "Suspended",       "-", RED)
        body_lay.addLayout(stats_h)

        # Exemptions
        ex_grp = QGroupBox("Subscription Exemptions"); ex_lay = QVBoxLayout(ex_grp)
        self._adm_ex_tbl = QTableWidget(0, 2)
        self._adm_ex_tbl.setHorizontalHeaderLabels(["Username","Added"])
        self._adm_ex_tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._adm_ex_tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._adm_ex_tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._adm_ex_tbl.setMaximumHeight(130); ex_lay.addWidget(self._adm_ex_tbl)
        ex_btns = QHBoxLayout()
        add_ex  = btn("+ Add Exemption","success"); add_ex.clicked.connect(self._admin_add_exempt)
        rem_ex  = btn("Remove Selected","danger");  rem_ex.clicked.connect(self._admin_remove_exempt)
        ex_btns.addWidget(add_ex); ex_btns.addWidget(rem_ex); ex_btns.addStretch()
        ex_lay.addLayout(ex_btns); body_lay.addWidget(ex_grp)

        # Companies table
        co_grp = QGroupBox("Companies"); co_lay = QVBoxLayout(co_grp)
        self._adm_co_tbl = QTableWidget(0, 6)
        self._adm_co_tbl.setHorizontalHeaderLabels(["Company","Username","Invoices","Status","Last Seen",""])
        self._adm_co_tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._adm_co_tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._adm_co_tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._adm_co_tbl.setAlternatingRowColors(True)
        self._adm_co_tbl.verticalHeader().setVisible(False)
        self._adm_co_tbl.doubleClicked.connect(lambda idx: self._admin_open_row(idx.row()))
        co_lay.addWidget(self._adm_co_tbl); body_lay.addWidget(co_grp)
        body_lay.addStretch()
        bb = QWidget(); bb.setObjectName("darkBar"); bb.setStyleSheet("QWidget#darkBar { background: #37474F; }")
        bb_h = QHBoxLayout(bb); bb_h.setContentsMargins(8,6,8,6)
        dl_b = btn("DOCUMENTS LIST","secondary"); dl_b.clicked.connect(lambda: self.show_screen("doc_list")); bb_h.addWidget(dl_b)
        bb_h.addStretch(); lay.addWidget(bb)

    def _admin_stat_card(self, parent_layout, label, value, color):
        card = QFrame(); card.setFrameShape(QFrame.Shape.StyledPanel)
        card.setStyleSheet("background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px;")
        cl = QVBoxLayout(card); cl.setContentsMargins(12,12,12,12)
        val_lbl = QLabel(value)
        val_lbl.setStyleSheet(f"color:{color}; font-size:22pt; font-weight:bold;")
        val_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_lbl = QLabel(label); lbl_lbl.setStyleSheet("color:#374151; font-size:9pt;")
        lbl_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(val_lbl); cl.addWidget(lbl_lbl)
        parent_layout.addWidget(card)
        return val_lbl

    def _on_show_admin(self):
        self._admin_refresh()

    def _admin_refresh(self):
        if not requests:
            QMessageBox.warning(self,"Offline","No internet connection."); return
        try:
            data      = self._master_api("get", "/v1/master/companies")
            companies = data.get("companies", [])
            stats     = data.get("stats", {})
            ex_data   = self._master_api("get", "/v1/master/exempt")
            exempts   = ex_data.get("exempt", [])
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed to load admin data:\n{e}"); return

        self._adm_companies = companies
        ex_count = sum(1 for e in exempts if e["username"] != "bluesky_master")
        self._adm_stat_co.setText(str(stats.get("total_companies", len(companies))))
        self._adm_stat_inv.setText(str(stats.get("total_invoices", 0)))
        self._adm_stat_ex.setText(str(ex_count))
        self._adm_stat_sus.setText(str(stats.get("suspended_count", 0)))

        # Exemptions table
        self._adm_ex_tbl.setRowCount(0)
        for e in exempts:
            r = self._adm_ex_tbl.rowCount(); self._adm_ex_tbl.insertRow(r)
            self._adm_ex_tbl.setItem(r,0,QTableWidgetItem(e["username"]))
            self._adm_ex_tbl.setItem(r,1,QTableWidgetItem((e.get("added_at","") or "")[:10]))

        # Companies table
        self._adm_co_tbl.setRowCount(0)
        for co in companies:
            sus       = co.get("is_suspended", False)
            last_act  = (co.get("last_activity","") or "")
            last_date = None
            try:
                if last_act:
                    last_date = datetime.fromisoformat(last_act.replace("Z",""))
            except Exception:
                pass
            days_since  = (datetime.now() - last_date).days if last_date else None
            inactive30  = days_since is not None and days_since >= 30

            if days_since is None:
                last_str = "Never"
            elif days_since == 0:
                last_str = "Today"
            elif days_since == 1:
                last_str = "Yesterday"
            else:
                last_str = f"{days_since}d ago"

            r = self._adm_co_tbl.rowCount(); self._adm_co_tbl.insertRow(r)
            for col, text in enumerate([
                co.get("company_name",""),
                "@"+co.get("username",""),
                str(co.get("invoice_count",0)),
                "Suspended" if sus else "Active",
                last_str,
            ]):
                item = QTableWidgetItem(text)
                if col == 3:
                    item.setForeground(QColor(RED if sus else GREEN))
                if not sus and inactive30:
                    item.setForeground(QColor("#CC6600"))
                    item.setBackground(QColor("#FFF3E0"))
                elif sus:
                    item.setBackground(QColor("#FFEBEE"))
                self._adm_co_tbl.setItem(r, col, item)

            uname = co.get("username","")
            vb = btn("View","primary"); vb.setFixedHeight(26)
            vb.clicked.connect(lambda chk=False, u=uname: self._admin_open_company_by(u))
            self._adm_co_tbl.setCellWidget(r,5,vb)

    def _admin_open_row(self, row):
        if not hasattr(self,"_adm_companies") or row < 0 or row >= len(self._adm_companies): return
        self._admin_open_company_by(self._adm_companies[row].get("username",""))

    def _admin_open_company_by(self, username):
        if not requests: return
        try:
            monthly = self._master_api("get", f"/v1/master/company/{username}/monthly")
            sub     = self._master_api("get", f"/v1/master/company/{username}/subscription")
            co      = next((c for c in getattr(self,"_adm_companies",[]) if c.get("username")==username), {})
            dlg = AdminCompanyDialog(username, co, monthly.get("monthly",[]), sub, self)
            dlg.exec()
            self._admin_refresh()
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed to load company detail:\n{e}")

    def _admin_add_exempt(self):
        from PyQt6.QtWidgets import QInputDialog
        username, ok = QInputDialog.getText(self,"Add Exemption","Username to exempt from billing:")
        if not ok or not username.strip(): return
        try:
            self._master_api("post", f"/v1/master/exempt/{username.strip().lower()}")
            self._admin_refresh()
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")

    def _admin_remove_exempt(self):
        row = self._adm_ex_tbl.currentRow()
        if row < 0: QMessageBox.information(self,"Select","Select an account first."); return
        username = self._adm_ex_tbl.item(row,0).text()
        if username == "bluesky_master":
            QMessageBox.warning(self,"Protected","Cannot remove bluesky_master exemption."); return
        if QMessageBox.question(self,"Remove Exemption",
                f"Remove exemption for @{username}?\nThey will go back to normal trial/subscription rules.",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        try:
            self._master_api("delete", f"/v1/master/exempt/{username}")
            self._admin_refresh()
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")

    def _admin_create_account(self):
        dlg = QDialog(self); dlg.setWindowTitle("Create New Shop Account"); dlg.setMinimumWidth(380)
        lay = QVBoxLayout(dlg); form = QFormLayout()
        c_e = QLineEdit(); u_e = QLineEdit()
        p_e = QLineEdit(); p_e.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("Company Name:", c_e); form.addRow("Username:", u_e); form.addRow("Password:", p_e)
        lay.addLayout(form)
        err_lbl = QLabel(""); err_lbl.setStyleSheet(f"color:{RED};"); lay.addWidget(err_lbl)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lay.addWidget(bb)
        if dlg.exec() != QDialog.DialogCode.Accepted: return
        c = c_e.text().strip(); u = u_e.text().strip().lower(); p = p_e.text().strip()
        if not c or not u or not p:
            QMessageBox.warning(self,"Required","All fields are required."); return
        if len(p) < 6:
            QMessageBox.warning(self,"Password","Password must be at least 6 characters."); return
        try:
            self._master_api("post","/v1/master/create_account",
                             json={"username":u,"password":p,"company_name":c})
            QMessageBox.information(self,"Created",f"Account @{u} created successfully.")
            self._admin_refresh()
        except Exception as e:
            QMessageBox.critical(self,"Error",f"Failed:\n{e}")


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  ENTRY POINT
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

def _hide_console():
    """Hide the Windows console window and remove it from the taskbar."""
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            GWL_EXSTYLE    = -20
            WS_EX_TOOLWINDOW = 0x00000080   # tool window - no taskbar button
            WS_EX_APPWINDOW  = 0x00040000   # force taskbar button
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            style = (style | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
            ctypes.windll.user32.ShowWindow(hwnd, 0)   # SW_HIDE
    except Exception:
        pass


_SKIP_VERSION_FILE = os.path.join(
    os.environ.get("APPDATA") or os.path.expanduser("~"),
    "BlueSkyDesktop", "skipped_update.txt"
)

def _get_skipped_version():
    try:
        with open(_SKIP_VERSION_FILE, "r") as f:
            return f.read().strip()
    except Exception:
        return ""

def _set_skipped_version(ver):
    try:
        os.makedirs(os.path.dirname(_SKIP_VERSION_FILE), exist_ok=True)
        with open(_SKIP_VERSION_FILE, "w") as f:
            f.write(ver)
    except Exception:
        pass


class _UpdateWorker(QObject):
    """Emits update_found(tag) on the main thread via Qt's queued signal mechanism."""
    update_found = pyqtSignal(str)

    def check(self):
        try:
            req = urllib.request.Request(_UPDATE_API,
                headers={"User-Agent": "BlueSkyDesktop"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read())
            tag = data.get("tag_name", "").lstrip("v")
            if not tag:
                return
            def ver_tuple(v):
                try: return tuple(int(x) for x in v.split("."))
                except: return (0,)
            # Skip if not newer than current, or if user already dismissed this version
            if ver_tuple(tag) <= ver_tuple(APP_VERSION):
                return
            if tag == _get_skipped_version():
                return
            self.update_found.emit(tag)
        except Exception:
            pass


def _check_for_update(parent=None):
    """Check GitHub for a newer release in a background thread using a Qt signal."""
    worker = _UpdateWorker()

    def _show(tag):
        msg = QMessageBox(parent)
        msg.setWindowTitle("Update Available")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText(
            f"<b>Blue Sky Smog v{tag} is available.</b><br><br>"
            f"You are running v{APP_VERSION}.<br>"
            "Download the new installer to update."
        )
        dl_btn  = msg.addButton("Download Now", QMessageBox.ButtonRole.AcceptRole)
        skip_btn = msg.addButton("Skip This Version", QMessageBox.ButtonRole.RejectRole)
        msg.addButton("Later", QMessageBox.ButtonRole.RejectRole)
        msg.exec()
        clicked = msg.clickedButton()
        if clicked == dl_btn:
            import webbrowser
            webbrowser.open(_DOWNLOAD_URL)
        elif clicked == skip_btn:
            _set_skipped_version(tag)

    # Qt automatically queues cross-thread signal delivery to the main thread
    worker.update_found.connect(_show, Qt.ConnectionType.QueuedConnection)
    t = threading.Thread(target=worker.check, daemon=True)
    t.start()
    # Keep worker alive until thread finishes
    threading.Thread(target=lambda: (t.join(), None), daemon=True).start()


def main():
    _hide_console()   # hide & remove from taskbar before anything else
    import traceback
    try:
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        app.setStyleSheet(APP_STYLE)
        init_db()
        migrate_db()
        creds = load_creds()
        if not (creds.get("token") or (creds.get("username") and creds.get("password"))):
            dlg = LoginDialog()
            if dlg.exec() != QDialog.DialogCode.Accepted:
                sys.exit(0)
        window = App()
        window.showMaximized()
        _hide_console()
        # Check for updates 3 seconds after launch so UI is fully loaded
        QTimer.singleShot(3000, lambda: _check_for_update(window))
        sys.exit(app.exec())
    except Exception as e:
        try:
            QMessageBox.critical(None, "Startup Error",
                f"Failed to start:\n{e}\n\n{traceback.format_exc()}")
        except Exception:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()


