"""
Blue Sky Smog - Desktop App  (PyQt6 rewrite)
Requires: pip install pyqt6 requests reportlab pymupdf
"""

import os, sys, json, uuid, sqlite3, threading, time, re, textwrap, urllib.request, subprocess
from pathlib import Path
from datetime import datetime, timedelta

APP_VERSION = "1.2.12"
_UPDATE_API  = "https://api.github.com/repos/blueskysmog1/bluesky-smog-mac/releases/latest"
_DOWNLOAD_URL = "https://github.com/blueskysmog1/bluesky-smog-mac/releases/latest/download/BlueSkyDesktop.dmg"

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
    QCalendarWidget, QDateEdit, QGraphicsDropShadowEffect,
    QGraphicsView, QGraphicsScene,
)
from PyQt6.QtCore import (
    Qt, QTimer, QThread, pyqtSignal, QDate, QPoint, QRect, QRectF, QObject, QUrl,
)
from PyQt6.QtGui import (
    QFont, QColor, QIcon, QPixmap, QBrush, QAction, QImage,
    QPainter, QPen, QPalette, QCursor, QDesktopServices,
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

# When frozen by PyInstaller, point requests at the bundled certifi CA bundle
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    _ca = os.path.join(sys._MEIPASS, "certifi", "cacert.pem")
    if os.path.isfile(_ca):
        os.environ.setdefault("SSL_CERT_FILE", _ca)
        os.environ.setdefault("REQUESTS_CA_BUNDLE", _ca)

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
SYNC_INTERVAL = 8
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

# Design system (new)
CLR_NAVY   = "#0D2B4E"
CLR_BLUE   = "#1B5FA8"
CLR_BFAINT = "#E8F0FB"
CLR_SURFACE= "#F4F7FB"
CLR_CARD   = "#FFFFFF"
CLR_BORDER = "#D0DCF0"
CLR_TEXT   = "#1A2332"
CLR_TSUB   = "#5A6E8A"
CLR_TMUTED = "#8FA3BC"
CLR_PASS   = "#1E7E34"
CLR_PASSBG = "#E6F4EA"
CLR_FAIL   = "#B91C1C"
CLR_FAILBG = "#FEE2E2"
CLR_WARN   = "#92400E"
CLR_WARNBG = "#FEF3C7"
_INTERVAL_OPTS = [("No reminder", None),("90 Days",90),("6 Months",183),("1 Year",365),("2 Years",730)]
CLR_EST    = "#5B21B6"
CLR_ESTBG  = "#EDE9FE"

# Legacy aliases (used throughout existing code)
PRIMARY  = CLR_BLUE
ACCENT   = "#4A90D9"
BG       = CLR_SURFACE
TEXT     = CLR_TEXT
WHITE    = CLR_CARD
RED      = CLR_FAIL
GREEN    = CLR_PASS
TODAY_BG = CLR_BFAINT
DARK_HDR = CLR_NAVY

# Global application stylesheet
_APP_STYLE = f"""
QMainWindow {{ background: {CLR_SURFACE}; }}
QWidget {{ color: {CLR_TEXT}; }}
QScrollArea {{ border: none; background: {CLR_SURFACE}; }}
QScrollArea > QWidget > QWidget {{ background: {CLR_SURFACE}; }}

QTableWidget {{
    background: {CLR_CARD};
    alternate-background-color: {CLR_SURFACE};
    border: 1px solid {CLR_BORDER};
    gridline-color: #EDF2F8;
    outline: none;
}}
QTableWidget::item {{ padding: 5px 8px; border: none; }}
QTableWidget::item:selected {{ background: {CLR_BFAINT}; color: {CLR_TEXT}; }}
QHeaderView::section {{
    background: {CLR_BLUE};
    color: white;
    padding: 5px 8px;
    font-weight: bold;
    border: none;
    border-right: 1px solid rgba(255,255,255,0.15);
}}
QHeaderView::section:last {{ border-right: none; }}

QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox {{
    background: #F8FAFD;
    border: 1.5px solid #B8CCE8;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 10.5pt;
    font-weight: 700;
    color: {CLR_BLUE};
    selection-background-color: {CLR_BFAINT};
    min-height: 26px;
}}
QLineEdit:focus, QTextEdit:focus, QSpinBox:focus {{ border-color: {CLR_BLUE}; background: {CLR_CARD}; }}
QLineEdit:read-only {{ background: {CLR_SURFACE}; color: {CLR_TSUB}; }}
QComboBox {{
    background: #F8FAFD;
    border: 1.5px solid #B8CCE8;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 10.5pt;
    font-weight: 700;
    color: {CLR_BLUE};
    min-height: 26px;
}}
QComboBox:focus {{ border-color: {CLR_BLUE}; background: {CLR_CARD}; }}
QComboBox::drop-down {{ border: none; width: 18px; }}
QComboBox QAbstractItemView {{
    background: {CLR_CARD};
    border: 1px solid {CLR_BORDER};
    selection-background-color: {CLR_BFAINT};
    selection-color: {CLR_TEXT};
    outline: none;
}}
QPushButton {{
    background: {CLR_BLUE};
    color: white;
    border: none;
    border-radius: 5px;
    padding: 5px 14px;
    font-weight: 600;
}}
QPushButton:hover {{ background: #1450A0; }}
QPushButton:pressed {{ background: #0D3A7A; }}
QPushButton:disabled {{ background: {CLR_BORDER}; color: {CLR_TMUTED}; }}
QGroupBox {{
    border: 1.5px solid {CLR_BORDER};
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 8px;
    font-weight: 600;
    color: {CLR_TEXT};
    background: {CLR_CARD};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
    background: {CLR_CARD};
    color: {CLR_TSUB};
    font-size: 8pt;
    font-weight: 700;
    text-transform: uppercase;
}}
QTabWidget::pane {{ border: 1px solid {CLR_BORDER}; background: {CLR_CARD}; border-radius: 6px; }}
QTabBar::tab {{
    padding: 7px 18px;
    border: 1px solid {CLR_BORDER};
    border-bottom: none;
    background: {CLR_SURFACE};
    color: {CLR_TSUB};
    margin-right: 2px;
    border-radius: 5px 5px 0 0;
}}
QTabBar::tab:selected {{ background: {CLR_CARD}; color: {CLR_BLUE}; font-weight: 700; }}
QScrollBar:vertical {{
    width: 8px; background: {CLR_SURFACE}; margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {CLR_BORDER}; border-radius: 4px; min-height: 20px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{
    height: 8px; background: {CLR_SURFACE};
}}
QScrollBar::handle:horizontal {{
    background: {CLR_BORDER}; border-radius: 4px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
QLabel {{ background: transparent; color: {CLR_TEXT}; }}
QCheckBox, QRadioButton {{ color: {CLR_TEXT}; spacing: 6px; background: transparent; }}
QDialog {{ background: {CLR_CARD}; }}
QMessageBox {{ background: {CLR_CARD}; }}
"""

DEFAULT_BUSINESS = {
    "name": "", "address_line1": "", "address_line2": "",
    "phone": "", "email": "", "website": "", "ard": "", "card_fee": 5.00,
    "logo_path": "", "qr_path": "",
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
    # vehicles migration — test interval / due date synced from mobile
    veh_cols = {row[1] for row in c.execute("PRAGMA table_info(vehicles)").fetchall()}
    if "test_interval_days" not in veh_cols:
        c.execute("ALTER TABLE vehicles ADD COLUMN test_interval_days INTEGER")
    if "next_test_due" not in veh_cols:
        c.execute("ALTER TABLE vehicles ADD COLUMN next_test_due TEXT NOT NULL DEFAULT ''")
    if "deleted" not in veh_cols:
        c.execute("ALTER TABLE vehicles ADD COLUMN deleted INTEGER NOT NULL DEFAULT 0")
    # Migrate data from old column names used in v1.1.x (service_interval_days / next_due)
    veh_cols = {row[1] for row in c.execute("PRAGMA table_info(vehicles)").fetchall()}
    if "service_interval_days" in veh_cols:
        c.execute("""UPDATE vehicles SET test_interval_days=service_interval_days
                     WHERE (test_interval_days IS NULL OR test_interval_days=0)
                       AND service_interval_days > 0""")
    if "next_due" in veh_cols:
        c.execute("""UPDATE vehicles SET next_test_due=next_due
                     WHERE (next_test_due='' OR next_test_due IS NULL)
                       AND next_due != ''""")
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

def api_pull(since_seq, limit=None, timeout=30):
    params = {"since_seq": since_seq}
    if limit: params["limit"] = limit
    r = requests.get(f"{API_BASE}/v1/sync/pull/{DEVICE_ID}",
                     params=params, headers=_hdrs(), timeout=timeout)
    r.raise_for_status(); return r.json().get("events", [])

def api_subscription_status():
    try:
        r = requests.get(f"{API_BASE}/v1/subscription/status", headers=_hdrs(), timeout=8)
        if r.status_code == 200: return r.json()
    except Exception: pass
    return {}

def api_decode_vin(vin):
    try:
        r = requests.get(f"https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVin/{vin}?format=json", timeout=8)
        data = r.json()
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
        # Not found by vehicle_id — check plate/VIN to avoid creating a duplicate row
        if plate or vin:
            r2 = conn.execute("SELECT vehicle_id FROM vehicles WHERE (plate!='' AND plate=?) OR (vin!='' AND vin=?)", (plate,vin)).fetchone()
            if r2:
                conn.execute("UPDATE vehicles SET customer_id=?,vin=?,plate=?,make=?,model=?,year=?,updated_at=? WHERE vehicle_id=?",
                             (customer_id,vin,plate,make,model,year,now_iso(),r2["vehicle_id"])); conn.commit(); return r2["vehicle_id"]
        conn.execute("INSERT OR IGNORE INTO vehicles(vehicle_id,customer_id,vin,plate,make,model,year,updated_at) VALUES(?,?,?,?,?,?,?,?)",
                     (vehicle_id,customer_id,vin,plate,make,model,year,now_iso())); conn.commit(); return vehicle_id
    r = conn.execute("SELECT vehicle_id FROM vehicles WHERE (plate!='' AND plate=?) OR (vin!='' AND vin=?)", (plate,vin)).fetchone()
    if r:
        conn.execute("UPDATE vehicles SET customer_id=?,vin=?,plate=?,make=?,model=?,year=?,updated_at=? WHERE vehicle_id=?",
                     (customer_id,vin,plate,make,model,year,now_iso(),r["vehicle_id"])); conn.commit(); return r["vehicle_id"]
    vid = str(uuid.uuid4())
    conn.execute("INSERT INTO vehicles(vehicle_id,customer_id,vin,plate,make,model,year,updated_at) VALUES(?,?,?,?,?,?,?,?)",
                 (vid,customer_id,vin,plate,make,model,year,now_iso())); conn.commit(); return vid
def get_next_invoice_number(conn): return 0  # server assigns invoice numbers via sync; 0 is placeholder
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
        """Re-sync everything from seq=0 WITHOUT deleting local data.
        Resets last_seq to 0 then replays all server events as upserts so
        existing local records are never at risk.  Large histories are pulled
        500 events at a time so a single request can never time out or OOM.
        notify_cb(msg) is called with progress updates and on completion."""
        slog("[ForcePull] Waiting for sync lock...")
        if not self._lock.acquire(blocking=True, timeout=30):
            msg="Could not acquire sync lock - try again in a moment."
            slog(f"[ForcePull] {msg}")
            if notify_cb: notify_cb(msg)
            return
        conn=get_db()
        try:
            # Reset the sequence pointer so we replay from the beginning.
            # Local data stays intact — upserts are idempotent and delete
            # events will still remove anything that was deleted on the server.
            set_last_seq(conn, 0); conn.commit()
            slog("[ForcePull] last_seq reset to 0 - replaying events from server...")

            PAGE = 500
            total = 0; since = 0
            while True:
                page = api_pull(since, limit=PAGE, timeout=60)
                if not page: break
                # Apply this page in dependency order
                cust_evs=[]; veh_evs=[]; inv_evs=[]; item_evs=[]; pay_evs=[]
                new_seq = since
                for ev in page:
                    seq=int(ev.get("seq",0))
                    entity=ev.get("entity",""); action=ev.get("action","")
                    payload=ev.get("payload",{})
                    if isinstance(payload,str):
                        try: payload=json.loads(payload)
                        except: payload={}
                    if entity=="customer" and action in ("upsert","delete"):
                        cust_evs.append((action,payload))
                    elif entity=="vehicle" and action in ("upsert","delete"):
                        veh_evs.append((action,payload))
                    elif entity=="invoice" and action in ("upsert","finalize","delete"):
                        inv_evs.append((action,payload))
                    elif entity=="invoice_item" and action in ("upsert","insert","create"):
                        item_evs.append(payload)
                    elif entity=="invoice_item" and action=="delete":
                        item_evs.append(("delete",payload))
                    elif entity=="account_payment" and action in ("upsert","delete"):
                        pay_evs.append((action,payload))
                    new_seq=max(new_seq,seq)
                for act,p in cust_evs:
                    try:
                        if act=="delete": self._delete_customer(conn,p)
                        else: self._merge_customer(conn,p)
                    except Exception as e2: slog(f"[ForcePull] customer SKIPPED err={e2}")
                for act,p in veh_evs:
                    try:
                        if act=="delete": self._delete_vehicle(conn,p)
                        else: self._merge_vehicle(conn,p)
                    except Exception as e2: slog(f"[ForcePull] vehicle SKIPPED err={e2}")
                for act,p in inv_evs:
                    try:
                        if act=="delete": self._delete_invoice(conn,p)
                        else: self._merge_invoice(conn,p)
                    except Exception as e2: slog(f"[ForcePull] invoice SKIPPED err={e2}")
                for ev in item_evs:
                    try:
                        if isinstance(ev,tuple): self._delete_invoice_item(conn,ev[1])
                        else: self._merge_invoice_item(conn,ev)
                    except Exception as e2: slog(f"[ForcePull] item SKIPPED err={e2}")
                for act,p in pay_evs:
                    try:
                        if act=="delete": self._merge_payment_delete(conn,p)
                        else: self._merge_payment(conn,p)
                    except Exception as e2: slog(f"[ForcePull] payment SKIPPED err={e2}")
                set_last_seq(conn, new_seq); conn.commit()
                total += len(page)
                slog(f"[ForcePull] Page done: {len(page)} events (total {total})")
                if notify_cb:
                    notify_cb(f"Syncing… {total} records processed")
                if len(page) < PAGE: break  # final page
                since = new_seq

            self._last_pull_count=total; self._last_pull_time=datetime.now()
            if self._on_change:
                try: self._on_change()
                except: pass
            msg=f"Re-sync complete.\n{total} events replayed from server." if total>0 else \
                "Re-sync complete.\nNo events found on server."
            slog(f"[ForcePull] Done - {msg}")
            if notify_cb: notify_cb(msg)
        except Exception as e:
            msg=f"Re-sync FAILED: {e}"
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
        if not events:
            self._last_pull_time=datetime.now(); self._last_pull_count=0
            if self._on_change:
                try: self._on_change()
                except: pass
            return
        new_seq=since; item_events=[]
        for ev in events:
            seq=int(ev.get("seq",0)); entity=ev.get("entity",""); action=ev.get("action",""); payload=ev.get("payload",{})
            if isinstance(payload,str):
                try: payload=json.loads(payload)
                except: payload={}
            try:
                if entity=="customer" and action=="upsert": self._merge_customer(conn,payload)
                elif entity=="customer" and action=="delete": self._delete_customer(conn,payload)
                elif entity=="vehicle" and action=="upsert": self._merge_vehicle(conn,payload)
                elif entity=="vehicle" and action=="delete": self._delete_vehicle(conn,payload)
                elif entity=="invoice" and action in ("upsert","finalize"): self._merge_invoice(conn,payload)
                elif entity=="invoice" and action=="delete": self._delete_invoice(conn,payload)
                elif entity=="invoice_item" and action in ("upsert","insert","create"): item_events.append(payload)
                elif entity=="invoice_item" and action=="delete": self._delete_invoice_item(conn,payload)
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
        rec_type=p.get("type","payment")  # 'payment' or 'adjustment'
        conn.execute(
            "INSERT OR IGNORE INTO account_history(company_name,entry_date,type,amount,note,invoice_id,payment_number,payment_id,partial_json) "
            "VALUES(?,?,?,?,?,?,?,?,?)",
            (company_name,entry_date,rec_type,amount,note,invoice_id,payment_number,payment_id,partial_json))
        if rec_type=="adjustment":
            # Zero the balance directly — adjustment sets total_owed to 0
            conn.execute("UPDATE accounts SET total_owed=0,updated_at=? WHERE company_name=?",
                         (now_iso(),company_name))
            slog(f"[Adjustment] balance zeroed for {company_name}")
        elif rec_type=="charge":
            # Charge ADDS to the balance (invoice placed on account from mobile)
            conn.execute("UPDATE accounts SET total_owed=total_owed+?,updated_at=? WHERE company_name=?",
                         (amount,now_iso(),company_name))
            slog(f"[Charge] balance increased for {company_name} ${amount:.2f}")
        else:
            conn.execute("UPDATE accounts SET total_owed=MAX(0,total_owed-?),updated_at=? WHERE company_name=?",
                         (amount,now_iso(),company_name))
            slog(f"[Payment] merged {payment_number} for {company_name} ${amount:.2f}")
        conn.commit()

    def _delete_customer(self,conn,p):
        cid=p.get("customer_id","")
        if not cid: return
        conn.execute("DELETE FROM customers WHERE customer_id=?",(cid,))
        conn.commit()
        slog(f"[Customer] deleted customer_id={cid}")
    def _delete_vehicle(self,conn,p):
        vid=p.get("vehicle_id","")
        if not vid: return
        conn.execute("DELETE FROM vehicles WHERE vehicle_id=?",(vid,))
        conn.commit()
        slog(f"[Vehicle] deleted vehicle_id={vid}")
    def _delete_invoice_item(self,conn,p):
        item_id=p.get("item_id","")
        if not item_id: return
        conn.execute("DELETE FROM invoice_lines WHERE remote_item_id=?",(item_id,))
        conn.commit()
        slog(f"[InvoiceItem] deleted item_id={item_id}")
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
        upsert_vehicle(conn,customer_id=cid,vin=p.get("vin",""),plate=p.get("plate",""),
                       make=p.get("make",""),model=p.get("model",""),year=p.get("year",""),
                       vehicle_id=vid if vid else None)
        interval=p.get("test_interval_days"); due=(p.get("next_test_due")or"").strip()
        if vid and (interval is not None or due):
            conn.execute("UPDATE vehicles SET test_interval_days=?,next_test_due=? WHERE vehicle_id=?",
                         (interval,due,vid)); conn.commit()
    def _delete_invoice(self,conn,p):
        iid=p.get("invoice_id","")
        if not iid: return
        conn.execute("DELETE FROM invoice_lines WHERE invoice_id=?",(iid,))
        conn.execute("DELETE FROM invoices WHERE invoice_id=?",(iid,)); conn.commit()
    def _merge_customer(self,conn,p):
        try:
            first=p.get("first_name",""); last=p.get("last_name",""); company=p.get("company_name","")
            name=(p.get("name","")).strip()
            if name and not first and not last and not company:
                parts=name.strip().split(" ",1); first=parts[0]; last=parts[1] if len(parts)>1 else ""
            cid=(p.get("customer_id","")).strip()
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
                incoming_status=(p.get("status","")or"").upper()
                existing_status=(existing["status"]or"ESTIMATE").upper()
                # Allow mobile to finalize a desktop-created estimate (status change ESTIMATE→*)
                is_finalization = existing_status=="ESTIMATE" and incoming_status not in ("","ESTIMATE","DRAFT")
                if not is_finalization:
                    if incoming_num and incoming_num!=existing["invoice_number"]:
                        conn.execute("UPDATE invoices SET invoice_number=?,synced=1 WHERE invoice_id=?",(incoming_num,iid)); conn.commit()
                    elif not existing["synced"]:
                        conn.execute("UPDATE invoices SET synced=1 WHERE invoice_id=?",(iid,)); conn.commit()
                    return
            first=p.get("first_name",""); last=p.get("last_name",""); company=p.get("company_name","")
            name=p.get("customer_name","")or f"{first} {last}".strip()
            # Split customer_name into first/last when individual fields are absent (mobile invoices)
            if name and not first and not last and not company:
                parts=name.strip().split(" ",1); first=parts[0]; last=parts[1] if len(parts)>1 else ""
            cid=(p.get("customer_id","")).strip()
            if cid:
                if not conn.execute("SELECT 1 FROM customers WHERE customer_id=?",(cid,)).fetchone():
                    conn.execute("INSERT OR IGNORE INTO customers(customer_id,first_name,last_name,company_name,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                                 (cid,first,last,company,now_iso(),now_iso())); conn.commit()
            else: cid=get_or_create_customer_id(conn,first,last,company)
            status=(p.get("status","")or"DRAFT").upper(); is_est=1 if status=="ESTIMATE" else 0
            inv_num=p.get("invoice_number",0)or 0; plate=p.get("plate",""); vin=p.get("vin","")
            year=p.get("year",""); make=p.get("make",""); model=p.get("model","")
            notes=p.get("notes")or""; pay_method=p.get("payment_method")or""; invoice_date=p.get("invoice_date")or""; amount_cents=int(p.get("amount_cents")or 0)
            po_number=p.get("po_number","")or""
            conn.execute("INSERT OR IGNORE INTO invoices(invoice_id,invoice_number,customer_id,customer_name,first_name,last_name,company_name,invoice_date,plate,vin,year,make,model,amount_cents,payment_method,status,notes,is_estimate,po_number,from_mobile,created_at,updated_at,synced) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?,1)",
                         (iid,inv_num,cid,name,first,last,company,invoice_date,plate,vin,year,make,model,amount_cents,pay_method,status,notes,is_est,po_number,now_iso(),now_iso()))
            conn.execute("UPDATE invoices SET invoice_number=?,customer_id=?,customer_name=?,first_name=?,last_name=?,company_name=?,invoice_date=?,plate=?,vin=?,year=?,make=?,model=?,amount_cents=?,payment_method=?,status=?,notes=?,is_estimate=?,po_number=?,from_mobile=1,updated_at=?,synced=1 WHERE invoice_id=?",
                         (inv_num,cid,name,first,last,company,invoice_date,plate,vin,year,make,model,amount_cents,pay_method,status,notes,is_est,po_number,now_iso(),iid))
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
            conn.commit()
        except Exception as e: slog(f"[Merge] invoice_item FAILED err={e}")

SYNC = SyncEngine()

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  PDF  (identical to original)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

def print_pdf(pdf_path, printer_name="", copies=1, parent_widget=None, silent=False):
    import subprocess, shutil
    # 1. SumatraPDF — silent, specific printer, most reliable
    sumatra = shutil.which("SumatraPDF") or r"C:\Program Files\SumatraPDF\SumatraPDF.exe"
    if os.path.exists(sumatra):
        try:
            target = printer_name or "default"
            for _ in range(copies): subprocess.Popen([sumatra,"-print-to",target,"-silent",pdf_path])
            return
        except Exception: pass
    # 2. Qt native printing — QPrinter routes directly to the named printer via Qt's
    #    printing stack (already bundled; same system used by printer settings screen)
    try:
        from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo
        from PyQt6.QtGui import QPainter, QImage
        from PyQt6.QtCore import QRect
        import fitz
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        if printer_name:
            target = printer_name.strip().lower()
            for info in QPrinterInfo.availablePrinters():
                if info.printerName().strip().lower() == target:
                    printer = QPrinter(info, QPrinter.PrinterMode.HighResolution)
                    break
        printer.setCopyCount(copies)
        printer.setFullPage(True)
        doc = fitz.open(pdf_path)
        painter = QPainter()
        if not painter.begin(printer):
            raise RuntimeError(f"Could not open printer '{printer.printerName()}'")
        dpi = min(max(printer.resolution(), 150), 300)
        zoom = dpi / 72.0
        for page_num in range(len(doc)):
            if page_num > 0: printer.newPage()
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img = QImage(bytes(pix.samples), pix.width, pix.height,
                         pix.stride, QImage.Format.Format_RGB888)
            painter.drawImage(QRect(0, 0, printer.width(), printer.height()), img)
        painter.end(); doc.close()
        return
    except Exception as e:
        if parent_widget:
            QMessageBox.warning(parent_widget, "Print Error",
                f"Could not print to '{printer_name}':\n{e}\n\nTry installing SumatraPDF.")
    # 3. Last resort — system default printer only, skipped for silent auto-print
    if not silent:
        try:
            if sys.platform == "win32":
                os.startfile(pdf_path, "print")
            elif sys.platform == "darwin":
                subprocess.Popen(["open", pdf_path])
            else:
                subprocess.Popen(["xdg-open", pdf_path])
        except Exception: pass

def build_invoice_pdf_path(inv_dir,invoice_date,first="",last="",company="",customer_name="",is_estimate=False,inv_num=None):
    prefix="EST" if is_estimate else "INV"
    customer=display_customer_name(first=first,last=last,company=company,customer_name=customer_name)
    num_part = safe_filename_part(str(inv_num)) if inv_num else safe_filename_part(invoice_date or datetime.today().strftime('%Y-%m-%d'))
    return os.path.join(inv_dir,f"{safe_filename_part(customer)}_{prefix}{num_part}.pdf")

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
    w, h = LETTER
    # ── Logo (left) ──────────────────────────────────────────────────────────
    logo_path = (biz.get("logo_path") or "").strip(); biz_x = 36
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(ImageReader(logo_path), 36, h-100, width=82, height=72,
                        preserveAspectRatio=True)
            biz_x = 126
        except Exception:
            try:
                c.drawImage(ImageReader(logo_path), 36, h-100, width=82, height=72,
                            preserveAspectRatio=True, mask="auto")
                biz_x = 126
            except Exception: pass
    # ── Business info (left) ─────────────────────────────────────────────────
    c.setFillColor(colors.black); c.setFont("Helvetica-Bold", 14)
    c.drawString(biz_x, h-28, biz.get("name", "BLUE SKY SMOG"))
    c.setFont("Helvetica", 9); info_y = h-43
    email   = (biz.get("email")   or "").strip()
    website = (biz.get("website") or "").strip()
    for line in filter(None, [
        biz.get("address_line1", ""),
        biz.get("address_line2", ""),
        f"Phone: {biz.get('phone','')}" if biz.get("phone") else "",
        f"Email: {email}"   if email   else "",
        f"Web:   {website}" if website else "",
        f"ARD #: {biz.get('ard','')}" if biz.get("ard") else "",
    ]):
        c.drawString(biz_x, info_y, str(line)); info_y -= 12
    # ── QR code (right) ──────────────────────────────────────────────────────
    has_qr = False
    qr_path = (biz.get("qr_path") or "").strip()
    if qr_path and os.path.exists(qr_path):
        try:
            c.drawImage(ImageReader(qr_path), w-108, h-100, width=68, height=68,
                        preserveAspectRatio=True)
            has_qr = True
        except Exception: pass
    # ── Title (right) ────────────────────────────────────────────────────────
    title_x = w-116 if has_qr else w-36
    c.setFont("Helvetica-Bold", 15); c.drawRightString(title_x, h-28, title)
    if subtitle: c.setFont("Helvetica", 9); c.drawRightString(title_x, h-44, subtitle)
    c.setStrokeColor(colors.HexColor("#0097A7")); c.setLineWidth(1.5); c.line(36, h-108, w-36, h-108)
    c.setLineWidth(1); c.setStrokeColor(colors.black)
    return has_qr

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
        has_qr = draw_header(c,biz,page_title or title,"")
        date_x = w-116 if has_qr else w-40
        c.setFont("Helvetica",FS_LABEL)
        c.drawRightString(w-170,h-58,f"{title.title()} #: {inv_num}")
        c.drawRightString(date_x,h-58,f"Date: {inv['invoice_date']}")
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
    prev_vin_pdf=""
    for line in lines:
        ensure_space(); svc_name=(line["service"]or"").strip()
        is_fee_line=svc_name in ("Credit Card Fee","Card Fee","CC Fee")
        is_cert_line=svc_name=="Certificate"
        if is_fee_line: vin_l=plate_l=odo_l=year_l=make_l=model_l=""
        else:
            vin_l=(line["vin"]or"").strip()or hdr_vin; plate_l=(line["plate"]or"").strip()or hdr_plate
            odo_l=(line["odometer"]or"").strip(); year_l=(line["year"]or"").strip()or hdr_year
            make_l=(line["make"]or"").strip()or hdr_make; model_l=(line["model"]or"").strip()or hdr_model
        result=(line["result"]or"").strip(); cert=(line["cert"]or"").strip()
        disc=float(line["discount"]or 0); price=float(line["price"]or 0); subtotal+=price
        same_vehicle=is_cert_line and vin_l and vin_l==prev_vin_pdf
        if not same_vehicle:
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
        if same_vehicle and is_cert_line:
            cert_x=90+c.stringWidth("Service: ","Helvetica",FS_BODY)
            c.drawString(cert_x,y,service_text); c.drawRightString(w-70,y,f"${price:,.2f}"); y-=LINE_H
        else:
            c.drawString(90,y,f"Service: {service_text}"); c.drawRightString(w-70,y,f"${price:,.2f}"); y-=LINE_H
        if disc>0: c.drawString(90,y,"Discount"); c.drawRightString(w-70,y,f"-${disc:,.2f}"); y-=LINE_H
        y-=6
        if not is_fee_line and vin_l: prev_vin_pdf=vin_l
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
        PAGE_WIDTH=LETTER[0]; LEFT_MARGIN=72; RIGHT_MARGIN=72; BOTTOM_Y=18; BAR_HEIGHT=18
        TARGET_WIDTH=PAGE_WIDTH-LEFT_MARGIN-RIGHT_MARGIN
        vin_b=(inv["vin"]or hdr_vin or"").strip(); plate_b=(inv["plate"]or hdr_plate or"").strip()
        year_b=(inv["year"]or hdr_year or"").strip(); make_b=(inv["make"]or hdr_make or"").strip()
        model_b=(inv["model"]or hdr_model or"").strip()
        barcode_value="\x1e".join([vin_b,plate_b,year_b,make_b,model_b,str(inv_num)])
        if not barcode_value.replace("\x1e","").strip(): barcode_value=f"INV{inv_num}-{inv['invoice_date']}"
        barcode=code128.Code128(barcode_value,barHeight=BAR_HEIGHT,barWidth=0.6)
        raw_width=max(float(getattr(barcode,"width",1.0)),1.0); scale_x=TARGET_WIDTH/raw_width
        c.saveState(); c.translate(LEFT_MARGIN,BOTTOM_Y); c.scale(scale_x,1.0); barcode.drawOn(c,0,0); c.restoreState()
    except Exception: pass
    c.save()
    conn.execute("UPDATE invoices SET pdf_path=? WHERE invoice_id=?",(out_path,invoice_id)); conn.commit()
    return True

class _UpComboBox(QComboBox):
    """QComboBox whose popup opens upward so it stays on screen."""
    def showPopup(self):
        super().showPopup()
        popup = self.findChild(QFrame)
        if popup:
            from PyQt6.QtCore import QPoint as _QP
            g = self.mapToGlobal(_QP(0, 0))
            popup.move(g.x(), g.y() - popup.height())

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  STYLESHEET
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

APP_STYLE = """
QWidget { font-size: 10pt; font-weight: bold; color: #111827; }
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
    background: #F8FAFD; border: 1.5px solid #B8CCE8; border-radius: 6px;
    padding: 6px 10px; font-size: 10.5pt; font-weight: 700;
    color: #1B5FA8; min-height: 26px; selection-background-color: #005B99;
}
QLineEdit:focus, QComboBox:focus, QTextEdit:focus { border-color: #1B5FA8; background: white; }
QComboBox::drop-down { border: none; padding-right: 4px; }

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

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
#  HELPERS
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

_BTN_STYLES = {
    "primary":   f"background:{CLR_BLUE};color:white;",
    "secondary": f"background:{CLR_SURFACE};color:{CLR_TSUB};border:1px solid {CLR_BORDER};",
    "success":   f"background:{CLR_PASS};color:white;",
    "danger":    f"background:{CLR_FAIL};color:white;",
    "accent":    f"background:#4A90D9;color:white;",
    "purple":    f"background:{CLR_EST};color:white;",
    "warning":   f"background:{CLR_WARN};color:white;",
}

def btn(text, style="primary", parent=None):
    b = QPushButton(text, parent)
    b.setObjectName(style)
    b.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
    base = _BTN_STYLES.get(style, _BTN_STYLES["primary"])
    b.setStyleSheet(f"QPushButton {{ {base} border-radius:5px; padding:5px 12px; font-weight:600; }}")
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
        self.setMinimumWidth(860); self.setMinimumHeight(900)
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

        # â"€â"€ Account Info â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
        sus       = self.co_info.get("is_suspended", False)
        plan      = self.sub_status.get("plan","?")
        is_exempt = (plan == "owner")
        info_grp  = QGroupBox("Account Info")
        ig        = QGridLayout(info_grp)
        for i,(k,v) in enumerate([
            ("Username",        f"@{self.co_info.get('username','')}"),
            ("Company",         co_name),
            ("Member Since",    (self.co_info.get("created_at","") or "")[:10]),
            ("Total Invoices",  str(self.co_info.get("invoice_count",0))),
            ("Status",          "Active" if not sus else "Suspended"),
            ("Billing",         "Exempt - No Billing" if is_exempt else f"Plan: {plan}"),
        ]):
            ig.addWidget(QLabel(k+":"), i, 0)
            vl = QLabel(v); vl.setStyleSheet("font-weight:bold;")
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
            mb_tbl.verticalHeader().setVisible(False)
            mb_tbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
            mb_tbl.setMinimumHeight(min(len(self.monthly), 12) * 28 + 32)
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


class SmogMasterImportDialog(QDialog):
    """Import customer records from a Smog Master CSV export."""
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Import from Smog Master"); self.setMinimumWidth(520)
        lay = QVBoxLayout(self); lay.setSpacing(10); lay.setContentsMargins(24,24,24,24)
        lay.addWidget(QLabel("Import Customers from Smog Master",
                             font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        desc = QLabel("Select a Smog Master CSV export file.")
        desc2 = QLabel("Expected columns: LastName, FirstName, Company, Phone, Email, Address, City, State, ZIP")
        desc3 = QLabel("(Column order does not matter - headers are matched by name.)")
        lay.addWidget(desc); lay.addWidget(desc2); lay.addWidget(desc3)
        file_row = QHBoxLayout()
        self._path_e = QLineEdit(); self._path_e.setPlaceholderText("Select file...")
        self._path_e.setReadOnly(True); file_row.addWidget(self._path_e)
        brw = btn("Browse","secondary"); brw.clicked.connect(self._browse); file_row.addWidget(brw)
        lay.addLayout(file_row)
        self._log = QTextEdit(); self._log.setReadOnly(True); self._log.setMaximumHeight(200)
        lay.addWidget(self._log)
        self._prog = QLabel(""); lay.addWidget(self._prog)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self._ok_btn = bb.button(QDialogButtonBox.StandardButton.Ok)
        self._ok_btn.setText("Import"); self._ok_btn.setEnabled(False)
        bb.accepted.connect(self._do_import); bb.rejected.connect(self.reject); lay.addWidget(bb)

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Smog Master Export",
                                              "", "CSV Files (*.csv);;All Files (*)")
        if path:
            self._path_e.setText(path); self._ok_btn.setEnabled(True)
            self._log.setText(f"File selected: {path}\nClick Import to begin.")

    def _do_import(self):
        import csv
        path = self._path_e.text().strip()
        if not path or not os.path.isfile(path):
            QMessageBox.warning(self, "No File", "Select a file first."); return
        self._ok_btn.setEnabled(False); self._log.clear()
        imported = 0; skipped = 0; errors = 0
        try:
            with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
                reader = csv.DictReader(f)
                def _col(row, *keys):
                    for k in keys:
                        for rk in row:
                            if rk.strip().lower() == k.lower():
                                return (row[rk] or "").strip()
                    return ""
                for i, row in enumerate(reader):
                    try:
                        first   = _col(row, "firstname", "first_name", "first")
                        last    = _col(row, "lastname",  "last_name",  "last")
                        company = _col(row, "company",   "company_name", "businessname", "business")
                        phone   = format_phone(_col(row, "phone", "phonenumber", "phone1"))
                        email   = _col(row, "email")
                        addr    = _col(row, "address",   "address1", "streetaddress")
                        city    = _col(row, "city")
                        state   = _col(row, "state")
                        zip_    = _col(row, "zip",  "zipcode", "postalcode")
                        if not (first or last or company):
                            skipped += 1; continue
                        cid = upsert_customer(self.db, first, last, company,
                                        phone=phone, email=email, address=addr,
                                        city=city, state=state, zip_=zip_)
                        enqueue(self.db, "customer", "upsert", {
                            "customer_id": cid, "first_name": first, "last_name": last,
                            "company_name": company, "phone": phone, "email": email,
                            "address": addr, "city": city, "state": state, "zip": zip_,
                            "discount_percent": 0, "discount_type": "PERCENT",
                        })
                        imported += 1
                    except Exception as row_e:
                        errors += 1
                        self._log.append(f"Row {i+2} error: {row_e}")
            self._log.append(f"\nDone \u2014 {imported} imported, {skipped} skipped (no name), {errors} errors.")
            self._prog.setText(f"Imported {imported} customers.")
            if imported > 0:
                self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Import Error", str(e))
            self._ok_btn.setEnabled(True)


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

        fp_lbl = QLabel('<a href="#">Forgot Password?</a>')
        fp_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        fp_lbl.setOpenExternalLinks(False)
        fp_lbl.linkActivated.connect(lambda _: self._do_forgot_password())
        lay.addWidget(fp_lbl)

        reg_btn = btn("Create Account", "secondary"); reg_btn.clicked.connect(self._do_register)
        lay.addWidget(reg_btn)

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

    def _do_forgot_password(self):
        dlg = QDialog(self); dlg.setWindowTitle("Reset Password"); dlg.setFixedWidth(380)
        lay = QVBoxLayout(dlg); lay.setSpacing(10); lay.setContentsMargins(24,24,24,24)
        hdr = QLabel("Reset Password"); hdr.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hdr.setStyleSheet(f"color:{PRIMARY}; font-size:14pt; font-weight:bold;")
        lay.addWidget(hdr)
        info = QLabel("Enter your username and we'll send a 6-digit code to the email on your account.")
        info.setWordWrap(True); info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info.setStyleSheet("color:#6B7280; font-size:9pt;"); lay.addWidget(info)

        form1 = QFormLayout(); form1.setSpacing(8)
        usr_e = QLineEdit(self._user_e.text().strip())
        form1.addRow("Username:", usr_e); lay.addLayout(form1)

        err_lbl = QLabel(""); err_lbl.setStyleSheet("color:red;")
        err_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); err_lbl.setWordWrap(True)
        lay.addWidget(err_lbl)

        # Step 2 widgets (hidden until code is sent)
        step2 = QWidget(); s2l = QFormLayout(step2); s2l.setSpacing(8)
        code_e  = QLineEdit(); code_e.setPlaceholderText("6-digit code from email"); code_e.setMaxLength(6)
        newpw_e = QLineEdit(); newpw_e.setEchoMode(QLineEdit.EchoMode.Password); newpw_e.setPlaceholderText("New password (8+ chars)")
        conpw_e = QLineEdit(); conpw_e.setEchoMode(QLineEdit.EchoMode.Password); conpw_e.setPlaceholderText("Confirm new password")
        s2l.addRow("Code:", code_e); s2l.addRow("New Password:", newpw_e); s2l.addRow("Confirm:", conpw_e)
        step2.setVisible(False); lay.addWidget(step2)

        send_btn  = btn("Send Code", "primary"); lay.addWidget(send_btn)
        reset_btn = btn("Reset Password", "primary"); reset_btn.setVisible(False); lay.addWidget(reset_btn)
        back_btn  = btn("Back", "secondary"); back_btn.setVisible(False); lay.addWidget(back_btn)
        close_btn = btn("Close", "secondary"); lay.addWidget(close_btn)
        close_btn.clicked.connect(dlg.reject)

        def _send():
            u = usr_e.text().strip().lower()
            if not u: err_lbl.setText("Enter your username."); return
            err_lbl.setText("Sending…"); QApplication.processEvents()
            try:
                r = requests.post(f"{API_BASE}/v1/auth/forgot_password",
                                  json={"username": u}, timeout=10)
                if r.status_code == 200:
                    err_lbl.setStyleSheet("color:green;")
                    err_lbl.setText("Code sent! Check your email.")
                    usr_e.setEnabled(False); step2.setVisible(True)
                    send_btn.setVisible(False); reset_btn.setVisible(True); back_btn.setVisible(True)
                elif r.status_code == 404:
                    err_lbl.setStyleSheet("color:red;")
                    err_lbl.setText("No account found with that username, or no email on file.")
                else:
                    err_lbl.setStyleSheet("color:red;")
                    err_lbl.setText(f"Error: {r.status_code}")
            except Exception as e:
                err_lbl.setStyleSheet("color:red;"); err_lbl.setText(str(e))

        def _reset():
            u = usr_e.text().strip().lower()
            code = code_e.text().strip(); newpw = newpw_e.text(); con = conpw_e.text()
            if len(code) != 6: err_lbl.setStyleSheet("color:red;"); err_lbl.setText("Enter the 6-digit code."); return
            if len(newpw) < 8: err_lbl.setStyleSheet("color:red;"); err_lbl.setText("Password must be at least 8 characters."); return
            if newpw != con: err_lbl.setStyleSheet("color:red;"); err_lbl.setText("Passwords do not match."); return
            err_lbl.setText("Resetting…"); QApplication.processEvents()
            try:
                r = requests.post(f"{API_BASE}/v1/auth/verify_reset_code",
                                  json={"username": u, "code": code, "new_password": newpw}, timeout=10)
                if r.status_code == 200:
                    QMessageBox.information(dlg, "Password Reset",
                        "Your password has been reset. Sign in with your new password.")
                    dlg.accept()
                else:
                    err_lbl.setStyleSheet("color:red;")
                    err_lbl.setText("Invalid or expired code. Request a new one.")
            except Exception as e:
                err_lbl.setStyleSheet("color:red;"); err_lbl.setText(str(e))

        def _back():
            step2.setVisible(False); send_btn.setVisible(True)
            reset_btn.setVisible(False); back_btn.setVisible(False)
            usr_e.setEnabled(True)
            err_lbl.setStyleSheet("color:red;"); err_lbl.setText("")

        send_btn.clicked.connect(_send); reset_btn.clicked.connect(_reset); back_btn.clicked.connect(_back)
        dlg.exec()

    def _do_register(self):
        dlg = QDialog(self); dlg.setWindowTitle("Create Account"); dlg.setFixedWidth(380)
        lay = QVBoxLayout(dlg); lay.setSpacing(10); lay.setContentsMargins(24,24,24,24)
        hdr = QLabel("Create Your Account"); hdr.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hdr.setStyleSheet(f"color:{PRIMARY}; font-size:14pt; font-weight:bold;")
        lay.addWidget(hdr)
        form = QFormLayout(); form.setSpacing(8)
        co_e  = QLineEdit(); co_e.setPlaceholderText("e.g. Valley Smog Pros")
        usr_e = QLineEdit(); usr_e.setPlaceholderText("lowercase, no spaces")
        pw_e  = QLineEdit(); pw_e.setEchoMode(QLineEdit.EchoMode.Password)
        pw2_e = QLineEdit(); pw2_e.setEchoMode(QLineEdit.EchoMode.Password)
        for w in (co_e, usr_e, pw_e, pw2_e): w.setMinimumWidth(240)
        form.addRow("Shop / Company Name:", co_e)
        form.addRow("Username:", usr_e)
        form.addRow("Password:", pw_e)
        form.addRow("Confirm Password:", pw2_e)
        lay.addLayout(form)
        err_lbl = QLabel(""); err_lbl.setStyleSheet("color:red;")
        err_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); err_lbl.setWordWrap(True)
        lay.addWidget(err_lbl)
        note = QLabel("A 30-day free trial starts automatically.\nNo credit card required to sign up.")
        note.setAlignment(Qt.AlignmentFlag.AlignCenter)
        note.setStyleSheet("color:#6B7280; font-size:8pt;"); lay.addWidget(note)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Create Account")
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lay.addWidget(bb)

        while True:
            if dlg.exec() != QDialog.DialogCode.Accepted: return
            co   = co_e.text().strip()
            usr  = usr_e.text().strip().lower()
            pw   = pw_e.text()
            pw2  = pw2_e.text()
            if not co or not usr or not pw:
                err_lbl.setText("All fields are required."); continue
            if len(usr) < 3:
                err_lbl.setText("Username must be at least 3 characters."); continue
            if len(pw) < 6:
                err_lbl.setText("Password must be at least 6 characters."); continue
            if pw != pw2:
                err_lbl.setText("Passwords do not match."); continue
            if not requests:
                err_lbl.setText("Network library not available."); continue
            err_lbl.setText("Creating account..."); QApplication.processEvents()
            try:
                r = requests.post(f"{API_BASE}/v1/auth/register",
                                  json={"username": usr, "password": pw, "company_name": co},
                                  timeout=15)
                r.raise_for_status()
                data = r.json()
                token      = data.get("token","")
                company_id = data.get("company_id","")
                company_name = data.get("company_name", co)
                save_creds({"username": usr, "password": pw, "token": token,
                            "company_id": company_id, "company_name": company_name})
                self._token = token
                QMessageBox.information(self, "Account Created",
                    f"Welcome to Blue Sky Smog, {company_name}!\n\n"
                    "Your 30-day free trial has started.\n"
                    "You are now signed in.")
                self.accept()
                return
            except Exception as ex:
                msg = str(ex)
                try:
                    body = ex.response.json() if hasattr(ex, "response") and ex.response else {}
                    msg = body.get("detail", msg)
                except Exception: pass
                err_lbl.setText(f"Registration failed: {msg}")

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

class BarChartWidget(QWidget):
    """Custom painted bar chart for the reports screen."""
    barClicked = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._data   = []   # list of (label, pass_cnt, fail_cnt)
        self._sel    = -1
        self.setMinimumHeight(150)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def setData(self, data):
        self._data = data; self._sel = -1; self.update()

    def selectedIndex(self): return self._sel

    def paintEvent(self, event):
        if not self._data: return
        p = QPainter(self); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        W = self.width(); H = self.height()
        BAR_H = H - 22      # reserve 22px for labels at bottom
        n = len(self._data)
        slot = W / n
        max_v = max((pv + fv for _, pv, fv in self._data), default=1) * 1.12

        # Faint grid lines
        for f in (0.33, 0.66, 1.0):
            y = int(BAR_H * (1 - f * 0.88))
            p.setPen(QPen(QColor("#E8F0FB"), 1))
            p.drawLine(0, y, W, y)

        for i, (label, pv, fv) in enumerate(self._data):
            cx   = i * slot + slot / 2
            bw   = max(3, int(slot * (0.38 if n <= 14 else 0.32)))
            lx   = int(cx - bw - 1)
            isSel  = self._sel == i
            dimmed = self._sel >= 0 and not isSel

            # Pass bar
            ph = int(pv / max_v * BAR_H * 0.88) if pv > 0 else 0
            if ph > 0:
                c = QColor("#2471C8") if isSel else (QColor("#B5CEED") if dimmed else QColor(CLR_BLUE))
                p.fillRect(lx, BAR_H - ph, bw, ph, c)
            # Fail bar
            fh = int(fv / max_v * BAR_H * 0.88) if fv > 0 else 0
            if fh > 0:
                c = QColor("#D63A3A") if isSel else (QColor("#EDA8A8") if dimmed else QColor(CLR_FAIL))
                p.fillRect(lx + bw + 2, BAR_H - fh, bw, fh, c)
            # Selection border
            if isSel:
                p.setPen(QPen(QColor(CLR_BLUE), 1.5))
                p.drawRect(lx - 2, 2, bw * 2 + 6, BAR_H - 4)
                p.setPen(QPen())

            # Label below bar area
            lc = QColor(CLR_BLUE) if isSel else (QColor("#C0D0E4") if dimmed else QColor(CLR_TMUTED))
            p.setPen(QPen(lc))
            f2 = QFont("Segoe UI", 8 if n > 20 else 9)
            if isSel: f2.setBold(True)
            p.setFont(f2)
            p.drawText(
                int(cx - slot / 2), BAR_H, int(slot), 22,
                Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                label
            )
        p.end()

    def mousePressEvent(self, event):
        if not self._data: return
        n   = len(self._data)
        idx = int(event.position().x() / (self.width() / n))
        if 0 <= idx < n:
            self._sel = -1 if self._sel == idx else idx
            self.update(); self.barClicked.emit(self._sel)


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
        open_btn.clicked.connect(lambda: (
            os.startfile(self.pdf_path) if sys.platform == "win32"
            else subprocess.Popen(["open", self.pdf_path])
        ))
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
        # One-time dedup: remove duplicate vehicle rows keeping the one with the latest next_test_due
        try:
            self.db.execute("""
                DELETE FROM vehicles WHERE vehicle_id NOT IN (
                    SELECT vehicle_id FROM vehicles v1
                    WHERE updated_at = (
                        SELECT MAX(v2.updated_at) FROM vehicles v2
                        WHERE (v1.vin!='' AND v2.vin=v1.vin)
                           OR (v1.vin='' AND v1.plate!='' AND v2.plate=v1.plate AND v2.vin='')
                    )
                    AND vehicle_id = (
                        SELECT MIN(v3.vehicle_id) FROM vehicles v3
                        WHERE v3.updated_at = (
                            SELECT MAX(v4.updated_at) FROM vehicles v4
                            WHERE (v1.vin!='' AND v4.vin=v1.vin)
                               OR (v1.vin='' AND v1.plate!='' AND v4.plate=v1.plate AND v4.vin='')
                        )
                        AND (v1.vin!='' AND v3.vin=v1.vin OR v1.vin='' AND v3.plate=v1.plate)
                    )
                )
            """)
            self.db.commit()
        except Exception:
            pass
        _app_data = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "BlueSkyDesktop")
        self.inv_dir = os.path.join(os.path.expanduser("~"), "Documents", "Invoices")
        os.makedirs(self.inv_dir, exist_ok=True)

        self._editing_id      = None
        self._editing_is_estimate = True
        self._lines_data      = []
        self._acct_names      = []
        self._acct_index      = 0
        self._sub_status      = {"status":"trial","can_create":True,"warning":""}
        self._is_master       = (load_creds().get("username","") == "bluesky_master")
        self._current_screen  = ""
        self._customer_touched= False
        try: self._zoom = json.loads(get_setting(self.db,"zoom_levels","{}"))
        except: self._zoom = {}

        # Apply global stylesheet
        QApplication.instance().setStyleSheet(_APP_STYLE)

        # ── New sidebar layout ──────────────────────────────────────
        central = QWidget(); central.setObjectName("appRoot")
        _mh = QHBoxLayout(central); _mh.setContentsMargins(0,0,0,0); _mh.setSpacing(0)
        self.setCentralWidget(central)

        # Sidebar (built after _screens dict exists)
        self._screens = {}
        self._sb_btns = {}

        # Right panel: topbar + stack + statusbar
        right = QWidget(); right.setObjectName("rightPanel")
        right.setStyleSheet(f"QWidget#rightPanel {{ background: {CLR_SURFACE}; }}")
        _rv = QVBoxLayout(right); _rv.setContentsMargins(0,0,0,0); _rv.setSpacing(0)

        # Topbar
        self._topbar = QWidget(); self._topbar.setObjectName("topbar")
        self._topbar.setFixedHeight(46)
        self._topbar.setStyleSheet(
            f"QWidget#topbar {{ background:{CLR_CARD}; border-bottom:1px solid {CLR_BORDER}; }}")
        _tbl = QHBoxLayout(self._topbar); _tbl.setContentsMargins(16,0,16,0); _tbl.setSpacing(8)
        self._topbar_title_lbl = QLabel("Documents")
        self._topbar_title_lbl.setStyleSheet(
            f"color:{CLR_TEXT}; font-size:12pt; font-weight:600; background:transparent;")
        _tbl.addWidget(self._topbar_title_lbl); _tbl.addStretch()
        self._topbar_right = QWidget()
        self._topbar_right.setStyleSheet("background:transparent;")
        _trh = QHBoxLayout(self._topbar_right); _trh.setContentsMargins(0,0,0,0); _trh.setSpacing(8)
        self._topbar_right_layout = _trh
        _tbl.addWidget(self._topbar_right)
        _rv.addWidget(self._topbar)

        # Page stack
        self._stack = QStackedWidget()
        _rv.addWidget(self._stack)

        # Status bar
        self._sbar = QWidget(); self._sbar.setObjectName("statusbar")
        self._sbar.setFixedHeight(24)
        self._sbar.setStyleSheet(f"QWidget#statusbar {{ background:{CLR_NAVY}; }}")
        _sbl = QHBoxLayout(self._sbar); _sbl.setContentsMargins(14,0,14,0); _sbl.setSpacing(14)
        self._sync_label = QLabel("Checking...")
        self._sync_label.setStyleSheet("color:rgba(100,220,130,0.85); font-size:9pt; background:transparent;")
        self._sbar_clock = QLabel("")
        self._sbar_clock.setStyleSheet("color:rgba(255,255,255,0.5); font-size:9pt; background:transparent;")
        _sbl.addWidget(self._sync_label); _sbl.addStretch(); _sbl.addWidget(self._sbar_clock)
        _rv.addWidget(self._sbar)

        # Build screens first (sidebar references self._new_estimate_action etc.)
        self._build_doc_list_screen()
        self._build_estimate_entry_screen()
        self._build_account_setup_screen()
        self._build_reports_screen()
        self._build_vehicles_due_screen()
        self._build_settings_screen()
        self._build_customers_screen()
        if self._is_master:
            self._build_admin_screen()

        # Build sidebar after screens so callbacks are valid
        self._sidebar = self._make_sidebar()
        _mh.addWidget(self._sidebar)
        _mh.addWidget(right)

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

        self._refresh_sidebar_name()
        self.show_screen("doc_list")

    def closeEvent(self, event):
        try: set_setting(self.db,"window_geometry", bytes(self.saveGeometry()).hex())
        except Exception: pass
        SYNC.stop()
        # Clear token so next launch always requires sign-in (keeps username/password pre-filled)
        if not getattr(self, '_logging_out', False):
            creds = load_creds()
            if creds.get("token"):
                save_creds({"username": creds.get("username",""), "password": creds.get("password","")})
        self.db.close(); event.accept()

    def show_screen(self, name, **kw):
        if name not in self._screens: return
        self._stack.setCurrentWidget(self._screens[name])
        self._current_screen = name
        size = self._zoom.get(name, 10)
        self._screens[name].setStyleSheet(f"font-size: {size}pt;")
        # Update sidebar active item
        for k, b in self._sb_btns.items():
            active = (k == name)
            if active:
                b.setStyleSheet(self._SB_ACTIVE_STYLE)
            else:
                b.setStyleSheet(self._SB_NAV_STYLE)
        handler = getattr(self, f"_on_show_{name}", None)
        if handler: handler(**kw)

    def _set_page_title(self, title, action_widgets=None):
        self._topbar_title_lbl.setText(title)
        while self._topbar_right_layout.count():
            item = self._topbar_right_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        if action_widgets:
            for w in action_widgets:
                self._topbar_right_layout.addWidget(w)

    def _tick_clock(self):
        now = datetime.now().strftime("%a, %m/%d/%Y  %I:%M %p")
        try: self._sbar_clock.setText(now)
        except Exception: pass
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
        biz = get_business_settings(self.db).get("name","").strip()
        if biz: return biz.upper()
        creds_name = load_creds().get("company_name","").strip()
        if creds_name: return creds_name.upper()
        return "BLUE SKY SMOG"

    def _refresh_sidebar_name(self):
        try:
            name = self._get_display_company_name()
            self._sb_biz_name.setText(name)
            biz = get_business_settings(self.db)
            sub = biz.get("address_line2","").strip() or biz.get("address_line1","").strip() or ""
            sub = re.sub(r'\s+\d{5}(-\d{4})?\s*$', '', sub).strip()
            self._sb_biz_sub.setText(sub)
            # Try to load logo
            logo_path = get_business_settings(self.db).get("logo_path","")
            if logo_path and os.path.isfile(logo_path):
                pix = QPixmap(logo_path).scaled(
                    30, 30, Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation)
                self._sb_logo_lbl.setPixmap(pix)
                self._sb_logo_lbl.setText("")
            else:
                self._sb_logo_lbl.setPixmap(QPixmap())
                self._sb_logo_lbl.setText("✦")
        except Exception: pass

    def _make_header(self, show_back=False):
        """Legacy stub – sidebar + topbar now replace the old header. Returns invisible 0-height widget."""
        hdr = QWidget(); hdr.setFixedHeight(0); hdr.setVisible(False)
        return hdr

    def _make_sidebar(self):
        NAV   = f"""QPushButton {{
            color:rgba(255,255,255,.6); background:transparent;
            border:none; border-left:3px solid transparent;
            text-align:left; padding:7px 13px; font-size:10pt;
        }}
        QPushButton:hover {{ color:rgba(255,255,255,.9); background:rgba(255,255,255,.07); }}"""
        ACTIVE= f"""QPushButton {{
            color:white; background:rgba(255,255,255,.12);
            border:none; border-left:3px solid #4A90D9;
            text-align:left; padding:7px 10px; font-size:10pt; font-weight:700;
        }}"""
        self._SB_NAV_STYLE    = NAV
        self._SB_ACTIVE_STYLE = ACTIVE

        sb = QWidget(); sb.setObjectName("sidebar"); sb.setFixedWidth(172)
        sb.setStyleSheet(f"QWidget#sidebar {{ background:{CLR_NAVY}; }}")
        lay = QVBoxLayout(sb); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)

        # Business header
        hdr = QWidget()
        hdr.setStyleSheet(f"border-bottom:1px solid rgba(255,255,255,.1); background:{CLR_NAVY};")
        hh = QHBoxLayout(hdr); hh.setContentsMargins(10,10,10,10); hh.setSpacing(8)
        self._sb_logo_lbl = QLabel()
        self._sb_logo_lbl.setFixedSize(34,34)
        self._sb_logo_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._sb_logo_lbl.setStyleSheet(
            "background:rgba(255,255,255,.12);border:1.5px solid rgba(255,255,255,.2);"
            "border-radius:7px;color:white;font-size:12pt;")
        self._sb_logo_lbl.setText("✦")
        hh.addWidget(self._sb_logo_lbl)
        nc = QVBoxLayout(); nc.setSpacing(1)
        self._sb_biz_name = QLabel(self._get_display_company_name())
        self._sb_biz_name.setStyleSheet("color:white;font-weight:700;font-size:10pt;background:transparent;")
        biz0 = get_business_settings(self.db)
        sub0 = biz0.get("address_line2","").strip() or biz0.get("address_line1","").strip() or ""
        sub0 = re.sub(r'\s+\d{5}(-\d{4})?\s*$', '', sub0).strip()
        self._sb_biz_sub  = QLabel(sub0)
        self._sb_biz_sub.setStyleSheet("color:rgba(255,255,255,.4);font-size:8pt;background:transparent;")
        nc.addWidget(self._sb_biz_name); nc.addWidget(self._sb_biz_sub)
        hh.addLayout(nc); lay.addWidget(hdr)

        def ns(text):
            lbl = QLabel(text)
            lbl.setStyleSheet(
                f"color:rgba(255,255,255,.3);font-size:8pt;letter-spacing:2px;"
                f"padding:6px 13px 2px;background:{CLR_NAVY};")
            lay.addWidget(lbl)

        def nb(key, text, cb):
            b = QPushButton(text); b.setFlat(True); b.setStyleSheet(NAV)
            b.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            b.clicked.connect(cb); lay.addWidget(b)
            self._sb_btns[key] = b

        ns("WORKSPACE")
        nb("doc_list",       "  Documents",   lambda: self.show_screen("doc_list"))
        nb("customers",      "  Customers",   lambda: self.show_screen("customers"))
        ns("MANAGE")
        nb("estimate_entry", "  New invoice", self._new_estimate_action)
        nb("reports",        "  Reports",     lambda: self.show_screen("reports"))
        nb("vehicles_due",   "  Vehicles Due",lambda: self.show_screen("vehicles_due"))
        nb("account_setup",  "  Accounts",   lambda: self.show_screen("account_setup"))
        lay.addStretch()

        sep = QWidget(); sep.setFixedHeight(1)
        sep.setStyleSheet(f"background:rgba(255,255,255,.1);"); lay.addWidget(sep)
        nb("settings", "  Settings", lambda: self.show_screen("settings"))

        # Sync button with status dot
        sync_b = QPushButton("  Sync"); sync_b.setFlat(True); sync_b.setStyleSheet(NAV)
        sync_b.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        sync_b.clicked.connect(self._manual_sync_now); lay.addWidget(sync_b)

        if self._is_master:
            nb("admin", "  Admin", lambda: self.show_screen("admin"))

        # Zoom controls in sidebar bottom
        zoom_w = QWidget(); zoom_w.setStyleSheet(f"background:{CLR_NAVY};")
        zh = QHBoxLayout(zoom_w); zh.setContentsMargins(10,4,10,4); zh.setSpacing(4)
        zs = "background:rgba(255,255,255,.15);color:white;border:none;border-radius:3px;padding:2px 6px;font-size:9pt;font-weight:bold;"
        zm_o = QPushButton("A-"); zm_o.setStyleSheet(zs); zm_o.setFixedWidth(30); zm_o.clicked.connect(self._zoom_out)
        zm_i = QPushButton("A+"); zm_i.setStyleSheet(zs); zm_i.setFixedWidth(30); zm_i.clicked.connect(self._zoom_in)
        zh.addWidget(zm_o); zh.addWidget(zm_i); zh.addStretch()
        lay.addWidget(zoom_w)

        # Logout
        lo = QPushButton("  Logout"); lo.setFlat(True)
        lo.setStyleSheet(NAV.replace("rgba(255,255,255,.6)","rgba(255,120,120,.85)"))
        lo.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        lo.clicked.connect(self._do_logout); lay.addWidget(lo)

        return sb

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
        if screen == "doc_list":
            self.refresh_doc_list()
        elif screen == "estimate_entry" and hasattr(self, '_ee_view'):
            scale = size / 10.0
            self._ee_view.resetTransform()
            self._ee_view.scale(scale, scale)
            if hasattr(self, '_ee_body_w'):
                vw = self._ee_view.viewport().width()
                new_w = max(int(vw / scale) - 8, 700)
                self._ee_body_w.setFixedWidth(new_w)
                self._ee_body_w.adjustSize()
                self._ee_scene.setSceneRect(QRectF(0, 0, new_w, self._ee_body_w.height()))

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
        self._sync_label.setText(f"● Synced {since} (+{cnt})")
        self._refresh_sidebar_name()
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
        self._stack.addWidget(w)

        # ── Page header: title + search + buttons ────────────────────────
        hdr = QWidget(); hdr.setStyleSheet(f"background:{CLR_CARD};border-bottom:1px solid {CLR_BORDER};")
        hdr_h = QHBoxLayout(hdr); hdr_h.setContentsMargins(20,10,16,10); hdr_h.setSpacing(12)
        _ttl = QLabel("Documents")
        _ttl.setStyleSheet(f"color:{CLR_TEXT};font-size:14pt;font-weight:700;background:transparent;")
        hdr_h.addWidget(_ttl); hdr_h.addSpacing(8)
        self._dl_search = QLineEdit()
        self._dl_search.setPlaceholderText("Search plate, VIN, customer...")
        self._dl_search.setStyleSheet(
            f"QLineEdit{{background:{CLR_SURFACE};border:1px solid {CLR_BORDER};border-radius:6px;"
            f"padding:4px 10px;color:{CLR_TEXT};}}")
        self._dl_search.setMinimumWidth(200); self._dl_search.setMaximumWidth(310)
        self._dl_search.textChanged.connect(self.refresh_doc_list)
        hdr_h.addWidget(self._dl_search); hdr_h.addStretch()
        self._dl_count_lbl = QLabel("")
        self._dl_count_lbl.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;background:transparent;")
        hdr_h.addWidget(self._dl_count_lbl)
        _show_all_b = QPushButton("Show All")
        _show_all_b.setStyleSheet(
            f"QPushButton{{background:transparent;border:1px solid {CLR_BORDER};border-radius:5px;"
            f"color:{CLR_TEXT};padding:4px 12px;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}")
        _show_all_b.clicked.connect(self._dl_show_all)
        hdr_h.addWidget(_show_all_b)
        _new_b = QPushButton("+ New")
        _new_b.setStyleSheet(
            f"QPushButton{{background:{CLR_BLUE};border:none;border-radius:5px;"
            f"color:white;padding:4px 14px;font-weight:600;}}"
            f"QPushButton:hover{{background:{CLR_NAVY};}}")
        _new_b.clicked.connect(self._new_estimate_action)
        hdr_h.addWidget(_new_b)
        lay.addWidget(hdr)

        # ── Filter pills ─────────────────────────────────────────────────
        pills_bar = QWidget()
        pills_bar.setStyleSheet(f"background:{CLR_SURFACE};border-bottom:1px solid {CLR_BORDER};")
        pills_h = QHBoxLayout(pills_bar); pills_h.setContentsMargins(20,6,16,6); pills_h.setSpacing(6)
        self._dl_time = "30d"
        self._dl_time_btns = {}
        for key, label in [("30d","Last 30 days"), ("week","This week"), ("all","All time")]:
            b = QPushButton(label)
            b.setStyleSheet(self._dl_pill_style(key == "30d"))
            self._dl_time_btns[key] = b
            b.clicked.connect(lambda _, k=key: self._dl_set_time(k))
            pills_h.addWidget(b)
        pills_h.addStretch()
        self._dl_result = "all"
        self._dl_result_btns = {}
        for key, label in [("pass","Pass"), ("fail","Fail"), ("est","Estimates"), ("all","All")]:
            b = QPushButton(label)
            b.setStyleSheet(self._dl_pill_style(key == "all"))
            self._dl_result_btns[key] = b
            b.clicked.connect(lambda _, k=key: self._dl_set_result(k))
            pills_h.addWidget(b)
        lay.addWidget(pills_bar)

        # ── Table ─────────────────────────────────────────────────────────
        cols = ["Date","Plate","Customer","VIN","Year / Make","Amount","Status"]
        self._dl_table = QTableWidget(0, len(cols)); self._dl_table.setHorizontalHeaderLabels(cols)
        hh = self._dl_table.horizontalHeader()
        self._dl_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._dl_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._dl_table.setAlternatingRowColors(True)
        self._dl_table.verticalHeader().setVisible(False)
        self._dl_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._dl_table.customContextMenuRequested.connect(self._dl_context_menu)
        self._dl_table.doubleClicked.connect(self._dl_open)
        for i, w2 in enumerate([85, 85, 0, 0, 130, 80, 90]):
            if w2: self._dl_table.setColumnWidth(i, w2)
        self._register_table("doc_list", self._dl_table)
        lay.addWidget(self._dl_table)
        self._dl_inv_ids = []

    def _on_show_doc_list(self):
        self._set_page_title("Documents")
        self.refresh_doc_list()

    def _dl_pill_style(self, active):
        if active:
            return (f"QPushButton{{background:{CLR_BLUE};color:white;border:none;border-radius:11px;"
                    f"padding:3px 13px;font-size:9pt;}}"
                    f"QPushButton:hover{{background:{CLR_NAVY};}}")
        return (f"QPushButton{{background:transparent;color:{CLR_TEXT};border:1px solid {CLR_BORDER};"
                f"border-radius:11px;padding:3px 13px;font-size:9pt;}}"
                f"QPushButton:hover{{background:{CLR_BFAINT};}}")

    def _dl_set_time(self, key):
        self._dl_time = key
        for k, b in self._dl_time_btns.items(): b.setStyleSheet(self._dl_pill_style(k == key))
        self.refresh_doc_list()

    def _dl_set_result(self, key):
        self._dl_result = key
        for k, b in self._dl_result_btns.items(): b.setStyleSheet(self._dl_pill_style(k == key))
        self.refresh_doc_list()

    def _dl_show_all(self):
        self._dl_time = "all"; self._dl_result = "all"
        for k, b in self._dl_time_btns.items(): b.setStyleSheet(self._dl_pill_style(k == "all"))
        for k, b in self._dl_result_btns.items(): b.setStyleSheet(self._dl_pill_style(k == "all"))
        self.refresh_doc_list()

    def refresh_doc_list(self):
        q     = self._dl_search.text().strip().lower()
        today = datetime.today().strftime("%Y-%m-%d")
        sql   = ("SELECT invoice_id,invoice_number,invoice_date,plate,vin,year,make,model,"
                 "veh_state,customer_name,first_name,last_name,company_name,customer_id,"
                 "amount_cents,payment_method,is_estimate,from_mobile,test_result,cert_number "
                 "FROM invoices WHERE 1=1")
        time_f = getattr(self, '_dl_time', '30d')
        if time_f == "30d":
            cutoff = (datetime.today() - timedelta(days=30)).strftime("%Y-%m-%d")
            sql += f" AND invoice_date >= '{cutoff}'"
        elif time_f == "week":
            td = datetime.today()
            wk = (td - timedelta(days=(td.weekday() + 1) % 7)).strftime("%Y-%m-%d")
            sql += f" AND invoice_date >= '{wk}'"
        result_f = getattr(self, '_dl_result', 'all')
        if result_f == "pass":   sql += " AND is_estimate=0 AND test_result='PASS'"
        elif result_f == "fail": sql += " AND is_estimate=0 AND test_result IN ('FAIL','RETEST')"
        elif result_f == "est":  sql += " AND is_estimate=1"
        sql += " ORDER BY invoice_date ASC, invoice_number ASC"
        rows = self.db.execute(sql).fetchall()

        # Pre-fetch invoice_lines VIN + result in one bulk query (mobile invoices store
        # these in invoice_lines rather than in the invoices table itself)
        inv_ids = [r["invoice_id"] for r in rows]
        il_vin_map: dict = {}; il_result_map: dict = {}
        if inv_ids:
            ph = ",".join("?" * len(inv_ids))
            for il in self.db.execute(
                f"SELECT invoice_id,vin,result FROM invoice_lines WHERE invoice_id IN ({ph})",
                inv_ids
            ).fetchall():
                iid = il["invoice_id"]
                if iid not in il_vin_map and (il["vin"] or "").strip():
                    il_vin_map[iid] = il["vin"].strip()
                if iid not in il_result_map and (il["result"] or "").strip():
                    il_result_map[iid] = il["result"].strip().upper()

        self._dl_table.setUpdatesEnabled(False)
        self._dl_table.setRowCount(0); self._dl_inv_ids = []; shown = 0
        for row in rows:
            cname = row["customer_name"] or row["company_name"] or \
                    f"{row['first_name']} {row['last_name']}".strip()
            plate = (row["plate"] or "").strip()
            vin   = (row["vin"] or "").strip()
            yr    = (row["year"] or "").strip()
            mk    = (row["make"] or "").strip()
            ymk   = f"{yr} {mk}".strip()
            if not plate and not ymk:
                vrow = _best_vehicle_for_invoice(self.db, row)
                if vrow:
                    plate = (vrow["plate"] or "").strip()
                    yr    = (vrow["year"] or "").strip()
                    mk    = (vrow["make"] or "").strip()
                    ymk   = f"{yr} {mk}".strip()
                    if not vin: vin = (vrow["vin"] or "").strip()
            if not vin:
                vin = il_vin_map.get(row["invoice_id"], "")
            num    = row["invoice_number"] or 0
            is_est = bool(row["is_estimate"])
            result = "" if is_est else (row["test_result"] or "").upper()
            if not result and not is_est:
                raw = il_result_map.get(row["invoice_id"], "")
                if raw == "PASSED": raw = "PASS"
                elif raw == "FAILED": raw = "FAIL"
                result = raw
            status = "ESTIMATE" if is_est else result
            try:
                date_str = datetime.strptime(row["invoice_date"], "%Y-%m-%d").strftime("%m/%d/%y")
            except Exception:
                date_str = row["invoice_date"]
            if q:
                blob = " ".join([str(num), row["invoice_date"], plate, vin, ymk, cname, status]).lower()
                if q not in blob: continue

            r = self._dl_table.rowCount(); self._dl_table.insertRow(r)
            self._dl_inv_ids.append(row["invoice_id"])
            is_today = (row["invoice_date"] == today)
            amt_cents = row["amount_cents"] or 0
            amt_str = f"${amt_cents/100:.2f}" if amt_cents else "—"
            values   = [date_str, plate, cname, vin, ymk, amt_str, status]
            _zsz = self._zoom.get(self._current_screen, 10)
            for col, val in enumerate(values):
                item = QTableWidgetItem(val)
                if col in (0, 6):
                    align = Qt.AlignmentFlag.AlignCenter
                elif col == 5:
                    align = Qt.AlignmentFlag.AlignRight
                else:
                    align = Qt.AlignmentFlag.AlignLeft
                item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | align)
                if is_today: item.setBackground(QColor(TODAY_BG))
                if col == 1:  # Plate: blue
                    item.setForeground(QColor(CLR_BLUE))
                if col == 6:  # Status badge
                    if status == "PASS":
                        item.setForeground(QColor(CLR_PASS)); item.setBackground(QColor(CLR_PASSBG))
                        item.setFont(QFont("Segoe UI", _zsz, QFont.Weight.Bold))
                    elif status in ("FAIL","RETEST"):
                        item.setForeground(QColor(CLR_FAIL)); item.setBackground(QColor(CLR_FAILBG))
                        item.setFont(QFont("Segoe UI", _zsz, QFont.Weight.Bold))
                    elif status == "ESTIMATE":
                        item.setForeground(QColor(CLR_EST)); item.setBackground(QColor(CLR_ESTBG))
                        item.setFont(QFont("Segoe UI", _zsz, QFont.Weight.Bold))
                self._dl_table.setItem(r, col, item)
            shown += 1

        self._dl_table.setUpdatesEnabled(True)
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
        if not iid: return
        row = self.db.execute("SELECT is_estimate FROM invoices WHERE invoice_id=?", (iid,)).fetchone()
        if row and row["is_estimate"]:
            self.show_screen("estimate_entry", invoice_id=iid)
        else:
            self._open_pdf_for_invoice(iid)

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
        pdf_row = self.db.execute("SELECT pdf_path FROM invoices WHERE invoice_id=?",(iid,)).fetchone()
        if pdf_row and pdf_row["pdf_path"] and os.path.exists(pdf_row["pdf_path"]):
            try: os.remove(pdf_row["pdf_path"])
            except: pass
        self.db.execute("DELETE FROM invoice_lines WHERE invoice_id=?",(iid,))
        self.db.execute("DELETE FROM invoices WHERE invoice_id=?",(iid,)); self.db.commit()
        # Push delete event directly to server so mobile sees it on the next pull.
        # Fall back to outbox if the push fails (background sync will retry).
        del_event_id = str(uuid.uuid4())
        del_payload  = {"invoice_id": iid}
        pushed = False
        if requests:
            try:
                api_push([{"event_id": del_event_id, "seq": 0,
                           "entity": "invoice", "action": "delete",
                           "payload": del_payload}])
                pushed = True
            except Exception as e:
                slog(f"[Delete] direct push failed, falling back to outbox: {e}")
        if not pushed:
            enqueue(self.db, "invoice", "delete", del_payload)
        self.refresh_doc_list()

    def _new_estimate_action(self):
        if not self._sub_status.get("can_create",True):
            QMessageBox.critical(self,"Subscription Required","Your free trial has ended.\n\nPlease subscribe to continue."); return
        self.show_screen("estimate_entry", invoice_id=None)

    def _do_logout(self, confirmed=False):
        if not confirmed:
            if QMessageBox.question(self,"Logout","Sign out and return to login?",
                    QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No
                    ) != QMessageBox.StandardButton.Yes: return
        # Keep username/password so login dialog pre-fills, but clear token
        creds = load_creds()
        save_creds({"username": creds.get("username",""), "password": creds.get("password","")})
        SYNC.stop()
        try:
            for t in ("invoices","invoice_lines","customers","vehicles","outbox"):
                self.db.execute(f"DELETE FROM {t}")
            self.db.commit(); set_last_seq(self.db,0)
        except Exception: pass
        self._logging_out = True
        # Signal main loop to show login again
        QApplication.instance()._show_login = True
        self.close()

    def _open_pdf_for_invoice(self, iid):
        row = self.db.execute(
            "SELECT is_estimate,invoice_date,customer_name,first_name,last_name,company_name,pdf_path FROM invoices WHERE invoice_id=?",
            (iid,)).fetchone()
        if not row: return
        ps = get_printer_setting(self.db)
        cached = (row["pdf_path"] or "").strip()
        if cached and os.path.isfile(cached):
            dlg = PdfViewerDialog(cached, ps, self); dlg.exec()
            return
        pdf = build_invoice_pdf_path(self.inv_dir, row["invoice_date"],
            company=row["company_name"], first=row["first_name"],
            last=row["last_name"], customer_name=row["customer_name"],
            is_estimate=bool(row["is_estimate"]))
        if not generate_invoice_pdf(iid, self.db, pdf):
            QMessageBox.critical(self,"Error","Could not generate PDF."); return
        dlg = PdfViewerDialog(pdf, ps, self); dlg.exec()

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: INVOICE / ESTIMATE ENTRY
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_estimate_entry_screen(self):
        outer = QWidget(); self._screens["estimate_entry"] = outer
        outer_lay = QVBoxLayout(outer); outer_lay.setContentsMargins(0,0,0,0); outer_lay.setSpacing(0)
        self._stack.addWidget(outer)

        # ── GHOST WIDGET: hidden fields kept for save/load/clear compatibility ──
        _ghost = QWidget(); _ghost.setVisible(False)
        _gl = QVBoxLayout(_ghost); _gl.setContentsMargins(0,0,0,0)
        self._ee_type_lbl = QLabel("NEW ESTIMATE"); _gl.addWidget(self._ee_type_lbl)
        self._ee_num_lbl  = QLabel("");              _gl.addWidget(self._ee_num_lbl)
        self._f_acct_id   = QComboBox(); self._f_acct_id.setEditable(True); _gl.addWidget(self._f_acct_id)
        self._f_po        = QLineEdit(); _upper_entry(self._f_po);           _gl.addWidget(self._f_po)
        self._f_vstate    = QLineEdit("CA"); _upper_entry(self._f_vstate);   _gl.addWidget(self._f_vstate)

        self._f_svc       = QComboBox(); self._f_svc.setMinimumWidth(180);   _gl.addWidget(self._f_svc)
        self._f_result    = QComboBox(); self._f_result.addItems(["Pass","Fail","Retest"]); _gl.addWidget(self._f_result)
        self._f_disc      = QLineEdit("0");                                  _gl.addWidget(self._f_disc)
        self._total_lbl   = QLabel("Total: $0.00");                          _gl.addWidget(self._total_lbl)
        self._ee_total_big= QLabel("TOTAL: $0.00");                          _gl.addWidget(self._ee_total_big)
        self._lines_table = QTableWidget(0,6)
        self._lines_table.setHorizontalHeaderLabels(["VIN","Service","Result","Cert #","Discount","Price"])
        self._lines_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._lines_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._lines_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._lines_table.doubleClicked.connect(self._edit_line)
        self._lines_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._lines_table.customContextMenuRequested.connect(self._line_context_menu)
        _gl.addWidget(self._lines_table)
        self._f_acct_id.currentTextChanged.connect(self._acct_id_changed)
        outer_lay.addWidget(_ghost)

        # ── PAGE HEADER ──
        ee_hdr = QWidget()
        ee_hdr.setStyleSheet(f"background:{CLR_SURFACE};border-bottom:1px solid {CLR_BORDER};")
        ee_hdr_h = QHBoxLayout(ee_hdr); ee_hdr_h.setContentsMargins(20,10,20,10)
        self._ee_hdr_lbl = QLabel("New Invoice")
        self._ee_hdr_lbl.setStyleSheet(
            f"font-size:15pt;font-weight:700;color:{CLR_TEXT};background:transparent;")
        _cancel_b = QPushButton("✕  Cancel")
        _cancel_b.setStyleSheet(
            f"QPushButton{{background:transparent;color:{CLR_TEXT};"
            f"border:1px solid {CLR_BORDER};border-radius:6px;padding:4px 14px;font-size:10pt;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}")
        _cancel_b.clicked.connect(lambda: (self._clear_form(), self.show_screen("doc_list")))
        ee_hdr_h.addWidget(self._ee_hdr_lbl); ee_hdr_h.addStretch(); ee_hdr_h.addWidget(_cancel_b)
        outer_lay.addWidget(ee_hdr)

        # ── CARD HELPERS ──
        _inp = (f"QLineEdit{{border:1px solid {CLR_BORDER};border-radius:4px;"
                f"padding:5px 8px;background:{CLR_CARD};color:{CLR_TEXT};font-size:10pt;}}"
                f"QLineEdit:focus{{border-color:{CLR_BLUE};}}")
        def _lbl(txt):
            l = QLabel(txt)
            l.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;background:transparent;")
            return l
        _inp_ss = (
            f"QWidget{{background:{CLR_CARD};}}"
            f"QLineEdit{{background:#F8FAFD;border:1px solid #B8CCE8;"
            f"border-radius:6px;padding:6px 10px;color:{CLR_BLUE};font-size:10pt;font-weight:700;}}"
            f"QLineEdit:focus{{border-color:{CLR_BLUE};background:{CLR_CARD};}}"
            f"QTextEdit{{background:#F8FAFD;border:1px solid #B8CCE8;"
            f"border-radius:6px;padding:5px 10px;color:{CLR_BLUE};font-size:10pt;}}"
            f"QTextEdit:focus{{border-color:{CLR_BLUE};background:{CLR_CARD};}}"
            f"QComboBox{{background:#F8FAFD;border:1px solid #B8CCE8;"
            f"border-radius:6px;padding:5px 10px;color:{CLR_BLUE};}}"
            f"QComboBox:focus{{border-color:{CLR_BLUE};}}"
        )
        def _card(icon_txt, title_txt):
            c = QFrame()
            c.setObjectName("invoiceCard")
            c.setStyleSheet(
                f"QFrame#invoiceCard{{background:{CLR_CARD};border:1px solid #C8D8EE;border-radius:10px;}}")
            vl = QVBoxLayout(c); vl.setContentsMargins(0,0,0,0); vl.setSpacing(0)
            ch = QWidget(); ch.setStyleSheet(f"QWidget{{background:{CLR_CARD};}}")
            chh = QHBoxLayout(ch); chh.setContentsMargins(16,12,16,10); chh.setSpacing(6)
            ic = QLabel(icon_txt)
            ic.setStyleSheet(f"color:{CLR_BLUE};font-size:11pt;background:transparent;")
            tl = QLabel(title_txt)
            tl.setStyleSheet(f"font-weight:700;font-size:10pt;color:{CLR_TEXT};background:transparent;")
            chh.addWidget(ic); chh.addWidget(tl); chh.addStretch()
            vl.addWidget(ch)
            sep = QWidget(); sep.setFixedHeight(1)
            sep.setStyleSheet("background:#EEF3FA;"); vl.addWidget(sep)
            body = QWidget(); body.setStyleSheet(_inp_ss)
            bl = QVBoxLayout(body); bl.setContentsMargins(16,14,16,14); bl.setSpacing(10)
            vl.addWidget(body, 1)
            return c, bl, chh

        # ── ZOOMABLE GRAPHICS VIEW + GRID ──
        self._ee_view = QGraphicsView()
        self._ee_view.setStyleSheet("QGraphicsView{border:none;background:#E6EDF7;}")
        self._ee_view.viewport().setStyleSheet("background:#E6EDF7;")
        self._ee_view.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self._ee_scene = QGraphicsScene()
        self._ee_view.setScene(self._ee_scene)
        body_w = QWidget()
        body_w.setStyleSheet("background:#E6EDF7;")
        body_w.setMinimumWidth(700)
        grid = QGridLayout(body_w)
        grid.setContentsMargins(16,16,16,16); grid.setSpacing(12)
        grid.setColumnStretch(0, 5); grid.setColumnStretch(1, 4)
        self._ee_body_w = body_w
        self._ee_proxy  = self._ee_scene.addWidget(body_w)

        def _ee_resize(event):
            scale = self._zoom.get("estimate_entry", 10) / 10.0
            vw = self._ee_view.viewport().width()
            new_w = max(int(vw / scale) - 8, 700)
            body_w.setFixedWidth(new_w)
            body_w.adjustSize()
            self._ee_scene.setSceneRect(QRectF(0, 0, new_w, body_w.height()))
            type(self._ee_view).resizeEvent(self._ee_view, event)

        self._ee_view.resizeEvent = _ee_resize
        outer_lay.addWidget(self._ee_view, 1)

        # ── CUSTOMER CARD ──
        cust_c, cust_bl, _ = _card(">>", "Customer")
        self._f_first   = QLineEdit(); _upper_entry(self._f_first);   self._f_first.setPlaceholderText("First name")
        self._f_last    = QLineEdit(); _upper_entry(self._f_last);    self._f_last.setPlaceholderText("Last name")
        self._f_company = QLineEdit(); _upper_entry(self._f_company); self._f_company.setPlaceholderText("Company name")
        self._f_addr  = QLineEdit(); _upper_entry(self._f_addr);  self._f_addr.setPlaceholderText("Address")
        self._f_zip   = QLineEdit(); _upper_entry(self._f_zip);   self._f_zip.setPlaceholderText("Zip")
        self._f_city  = QLineEdit(); _upper_entry(self._f_city);  self._f_city.setPlaceholderText("City")
        self._f_state = QLineEdit(); _upper_entry(self._f_state); self._f_state.setPlaceholderText("State"); self._f_state.setMaximumWidth(60)
        self._f_phone = QLineEdit(); self._f_phone.setPlaceholderText("Phone")
        self._f_email = QLineEdit(); _upper_entry(self._f_email); self._f_email.setPlaceholderText("Email")
        self._f_notes = QTextEdit(); self._f_notes.setMaximumHeight(56)
        self._f_notes.setPlaceholderText("Contact / notes (optional)")
        for w2 in (self._f_first, self._f_last, self._f_company):
            w2.textChanged.connect(self._autocomplete_customer)
            w2.editingFinished.connect(self._fill_customer)
        self._f_zip.editingFinished.connect(self._zip_lookup)
        self._f_phone.editingFinished.connect(self._fmt_phone)
        # City | State row
        city_row = QHBoxLayout(); city_row.setSpacing(6)
        city_row.addWidget(self._f_city, 1); city_row.addWidget(self._f_state)
        # Phone | Email row
        pe_row = QHBoxLayout(); pe_row.setSpacing(6)
        pe_row.addWidget(self._f_phone, 1); pe_row.addWidget(self._f_email, 1)
        name_row = QHBoxLayout(); name_row.setSpacing(6)
        name_row.addWidget(self._f_first, 1); name_row.addWidget(self._f_last, 1)
        cust_bl.addWidget(_lbl("First / Last name")); cust_bl.addLayout(name_row)
        cust_bl.addWidget(_lbl("Company name")); cust_bl.addWidget(self._f_company)
        cust_bl.addWidget(_lbl("Address")); cust_bl.addWidget(self._f_addr)
        cust_bl.addWidget(_lbl("City / State / Zip"))
        zip_row = QHBoxLayout(); zip_row.setSpacing(6)
        zip_row.addLayout(city_row, 3); zip_row.addWidget(self._f_zip, 1)
        cust_bl.addLayout(zip_row)
        cust_bl.addWidget(_lbl("Phone / Email")); cust_bl.addLayout(pe_row)
        cust_bl.addWidget(_lbl("Notes")); cust_bl.addWidget(self._f_notes)
        grid.addWidget(cust_c, 1, 0)

        # ── INSPECTION RESULT CARD ──
        insp_c, insp_bl, _ = _card("☑", "Inspection result")
        self._ee_pass_btn = QPushButton("PASS")
        self._ee_fail_btn = QPushButton("FAIL")
        self._ee_pass_btn.clicked.connect(lambda: self._ee_toggle_result("Pass"))
        self._ee_fail_btn.clicked.connect(lambda: self._ee_toggle_result("Fail"))
        self._ee_toggle_result("", init=True)  # set initial inactive styles
        tog_row = QHBoxLayout(); tog_row.setSpacing(0)
        tog_row.addWidget(self._ee_pass_btn); tog_row.addWidget(self._ee_fail_btn); tog_row.addStretch()
        self._f_cert = QLineEdit(); self._f_cert.setPlaceholderText("Pending")
        self._ee_tech = QLineEdit(); self._ee_tech.setPlaceholderText("J. Hernandez")
        ct_row = QHBoxLayout(); ct_row.setSpacing(10)
        ct_lbl_row = QHBoxLayout(); ct_lbl_row.setSpacing(10)
        ct_lbl_row.addWidget(_lbl("Certificate #"), 1); ct_lbl_row.addWidget(_lbl("Technician"), 1)
        ct_row.addWidget(self._f_cert, 1); ct_row.addWidget(self._ee_tech, 1)
        insp_bl.addWidget(_lbl("Result")); insp_bl.addLayout(tog_row)
        insp_bl.addLayout(ct_lbl_row); insp_bl.addLayout(ct_row)
        insp_bl.addStretch()
        grid.addWidget(insp_c, 0, 1)

        # ── VEHICLE CARD ──
        veh_c, veh_bl, _ = _card(">>", "Vehicle")
        # Row 1: Plate | Test date
        self._f_plate = QLineEdit(); _upper_entry(self._f_plate)
        self._f_plate.setPlaceholderText("License plate")
        self._f_plate.setStyleSheet(
            f"QLineEdit{{border:1px solid #B8CCE8;border-radius:6px;"
            f"padding:6px 10px;background:#F8FAFD;color:{CLR_BLUE};"
            f"font-size:10pt;font-weight:700;}}"
            f"QLineEdit:focus{{border-color:{CLR_BLUE};background:{CLR_CARD};}}")
        self._inv_date_e = QLineEdit(datetime.today().strftime("%Y-%m-%d"))
        self._inv_date_e.setReadOnly(True)
        cal_btn = QPushButton("..."); cal_btn.setFixedWidth(32)
        cal_btn.setStyleSheet(
            f"QPushButton{{border:1px solid {CLR_BORDER};border-radius:4px;"
            f"background:{CLR_CARD};padding:4px;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}")
        cal_btn.clicked.connect(self._pick_inv_date)
        date_w = QWidget(); date_h = QHBoxLayout(date_w)
        date_h.setContentsMargins(0,0,0,0); date_h.setSpacing(4)
        date_h.addWidget(self._inv_date_e, 1); date_h.addWidget(cal_btn)
        self._no_plate_cb = QCheckBox("No Plate")
        self._no_plate_cb.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;background:transparent;")
        self._no_plate_cb.toggled.connect(self._on_no_plate_toggled)
        plate_w = QWidget(); plate_col = QVBoxLayout(plate_w)
        plate_col.setContentsMargins(0,0,0,0); plate_col.setSpacing(3)
        plate_col.addWidget(self._f_plate); plate_col.addWidget(self._no_plate_cb)
        r1 = QGridLayout(); r1.setSpacing(8)
        r1.addWidget(_lbl("License plate"),0,0); r1.addWidget(plate_w,1,0)
        r1.addWidget(_lbl("Test date"),0,1);      r1.addWidget(date_w,1,1)
        r1.setColumnStretch(0,2); r1.setColumnStretch(1,2)
        # Row 2: VIN
        self._f_vin = QLineEdit(); _upper_entry(self._f_vin); self._f_vin.setPlaceholderText("VIN")
        # Row 3: Year | Make | Model
        self._f_year  = QLineEdit(); _upper_entry(self._f_year);  self._f_year.setPlaceholderText("Year")
        self._f_make  = QLineEdit(); _upper_entry(self._f_make);  self._f_make.setPlaceholderText("Make")
        self._f_model = QLineEdit(); _upper_entry(self._f_model); self._f_model.setPlaceholderText("Model")
        r3 = QGridLayout(); r3.setSpacing(8)
        r3.addWidget(_lbl("Year"),0,0);  r3.addWidget(self._f_year,1,0)
        r3.addWidget(_lbl("Make"),0,1);  r3.addWidget(self._f_make,1,1)
        r3.addWidget(_lbl("Model"),0,2); r3.addWidget(self._f_model,1,2)
        # Row 4: Odometer + Next Test Due (interval shortcut + direct date picker)
        self._f_odo  = QLineEdit(); _upper_entry(self._f_odo);  self._f_odo.setPlaceholderText("Odometer")
        self._f_test_interval = QComboBox()
        for lbl_txt, _ in _INTERVAL_OPTS: self._f_test_interval.addItem(lbl_txt)
        self._f_next_due_date = QDateEdit()
        self._f_next_due_date.setCalendarPopup(True)
        self._f_next_due_date.setDisplayFormat("MM/dd/yyyy")
        self._f_next_due_date.setDate(QDate.currentDate())  # default = today (invoice date)
        def _update_next_due_from_interval(idx):
            _, days = _INTERVAL_OPTS[idx]
            if days is not None:
                inv_date_str = self._inv_date_e.text().strip()
                try:
                    base = datetime.strptime(inv_date_str, "%Y-%m-%d").date()
                except Exception:
                    base = datetime.now().date()
                self._f_next_due_date.setDate(QDate(*(base + timedelta(days=days)).timetuple()[:3]))
        self._f_test_interval.currentIndexChanged.connect(_update_next_due_from_interval)
        r4 = QHBoxLayout(); r4.setSpacing(8)
        odo_col = QVBoxLayout(); odo_col.setSpacing(4)
        odo_col.addWidget(_lbl("Odometer")); odo_col.addWidget(self._f_odo)
        int_col = QVBoxLayout(); int_col.setSpacing(4)
        int_col.addWidget(_lbl("Reminder Interval")); int_col.addWidget(self._f_test_interval)
        due_col = QVBoxLayout(); due_col.setSpacing(4)
        due_col.addWidget(_lbl("Next Test Due Date")); due_col.addWidget(self._f_next_due_date)
        r4.addLayout(odo_col, 1); r4.addLayout(int_col, 1); r4.addLayout(due_col, 1)
        # Row 5: Service/test type
        self._veh_svc_cmb = QComboBox()
        # Row 6: Add to Invoice button
        add_to_inv_btn = QPushButton("Add to Invoice")
        add_to_inv_btn.setStyleSheet(
            f"QPushButton{{background:{CLR_BLUE};color:white;border:none;"
            f"border-radius:6px;padding:7px 14px;font-size:10pt;font-weight:700;}}"
            f"QPushButton:hover{{background:{CLR_NAVY};}}")
        def _add_to_invoice():
            self._f_svc.setCurrentText(self._veh_svc_cmb.currentText())
            self._add_line()
        add_to_inv_btn.clicked.connect(_add_to_invoice)
        veh_bl.addLayout(r1)
        veh_bl.addWidget(_lbl("VIN")); veh_bl.addWidget(self._f_vin)
        veh_bl.addLayout(r3); veh_bl.addLayout(r4)
        veh_bl.addWidget(_lbl("Service / Test type")); veh_bl.addWidget(self._veh_svc_cmb)
        veh_bl.addWidget(add_to_inv_btn); veh_bl.addStretch()
        self._f_vin.editingFinished.connect(self._vin_lookup)
        self._f_plate.editingFinished.connect(self._plate_lookup)
        grid.addWidget(veh_c, 0, 0)

        # ── BILLING CARD ──
        bill_c, bill_bl, _ = _card("$", "Billing")

        self._disc_info_lbl = QLabel("")
        self._disc_info_lbl.setStyleSheet(f"color:{CLR_PASS}; font-size:9pt; font-style:italic;")
        bill_bl.addWidget(self._disc_info_lbl)
        self._ee_billing_lines = QVBoxLayout(); self._ee_billing_lines.setSpacing(2)
        bill_bl.addLayout(self._ee_billing_lines)
        bill_bl.addStretch()

        # Payment method row
        pay_sep = QWidget(); pay_sep.setFixedHeight(1)
        pay_sep.setStyleSheet(f"background:{CLR_BORDER};")
        self._f_pay = _UpComboBox()
        self._f_pay.addItems(["","CASH","VISA","MASTERCARD","DISCOVER","AMEX","CHECK","CHARGE"])
        self._f_pay.currentTextChanged.connect(self._payment_changed)
        pay_row = QHBoxLayout(); pay_row.setSpacing(8)
        pay_row.addWidget(_lbl("Payment")); pay_row.addWidget(self._f_pay, 1)
        bill_bl.addWidget(pay_sep)
        bill_bl.addLayout(pay_row)

        # Total footer
        tot_sep = QWidget(); tot_sep.setFixedHeight(1)
        tot_sep.setStyleSheet(f"background:{CLR_BORDER};")
        tot_footer = QWidget()
        tot_footer.setStyleSheet(
            f"background:{CLR_NAVY};border-radius:4px;border:none;")
        tf_h = QHBoxLayout(tot_footer); tf_h.setContentsMargins(12,8,12,8)
        tf_l = QLabel("Total")
        tf_l.setStyleSheet("color:white;font-weight:700;font-size:11pt;background:transparent;")
        self._ee_bill_total = QLabel("$0.00")
        self._ee_bill_total.setStyleSheet("color:white;font-weight:700;font-size:11pt;background:transparent;")
        tf_h.addWidget(tf_l); tf_h.addStretch(); tf_h.addWidget(self._ee_bill_total)
        bill_bl.addWidget(tot_sep); bill_bl.addWidget(tot_footer)
        grid.addWidget(bill_c, 1, 1)

        # ── BOTTOM ACTION BAR ──
        bb = QWidget()
        bb.setStyleSheet(f"background:{CLR_CARD};border-top:1px solid {CLR_BORDER};")
        bb_h = QHBoxLayout(bb); bb_h.setContentsMargins(16,8,16,8); bb_h.setSpacing(8)

        self._ee_clear_btn = QPushButton("Cancel")
        self._ee_clear_btn.setStyleSheet(
            f"QPushButton{{background:transparent;color:{CLR_TEXT};"
            f"border:1px solid {CLR_BORDER};border-radius:6px;padding:6px 16px;font-size:10pt;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}")
        self._ee_clear_btn.clicked.connect(self._clear_form)

        auto_b = QPushButton("Auto-prints via  printer settings  ▾")
        auto_b.setStyleSheet(
            f"QPushButton{{background:transparent;color:{CLR_TMUTED};"
            f"border:none;font-size:9pt;padding:6px 4px;}}"
            f"QPushButton:hover{{color:{CLR_TEXT};}}")

        self._ee_delete_btn = QPushButton("Delete")
        self._ee_delete_btn.setStyleSheet(
            f"QPushButton{{background:transparent;color:{CLR_FAIL};"
            f"border:none;font-size:9pt;padding:6px 8px;}}"
            f"QPushButton:hover{{text-decoration:underline;}}")
        self._ee_delete_btn.clicked.connect(self._delete_document_action)

        self._ee_print_btn = QPushButton("Print PDF")
        self._ee_print_btn.setStyleSheet(
            f"QPushButton{{background:transparent;color:{CLR_TEXT};"
            f"border:1px solid {CLR_BORDER};border-radius:6px;padding:6px 14px;font-size:10pt;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}")
        self._ee_print_btn.clicked.connect(self._open_selected_pdf)

        self._ee_estimate_btn = QPushButton("Issue estimate")
        self._ee_estimate_btn.setStyleSheet(
            f"QPushButton{{background:{CLR_SURFACE};color:{CLR_TEXT};"
            f"border:1px solid {CLR_BORDER};border-radius:6px;padding:6px 18px;"
            f"font-size:10pt;font-weight:600;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}")
        self._ee_estimate_btn.clicked.connect(self._save_estimate_action)

        self._ee_issue_btn = QPushButton("Issue invoice")
        self._ee_issue_btn.setStyleSheet(
            f"QPushButton{{background:{CLR_BLUE};color:white;border:none;"
            f"border-radius:6px;padding:6px 18px;font-size:10pt;font-weight:600;}}"
            f"QPushButton:hover{{background:{CLR_NAVY};}}")
        self._ee_issue_btn.clicked.connect(self._issue_action)

        bb_h.addWidget(self._ee_clear_btn); bb_h.addWidget(auto_b); bb_h.addStretch()
        bb_h.addWidget(self._ee_delete_btn); bb_h.addWidget(self._ee_print_btn)
        bb_h.addWidget(self._ee_estimate_btn); bb_h.addWidget(self._ee_issue_btn)
        outer_lay.addWidget(bb)

    def _ee_toggle_result(self, result, init=False):
        if not init:
            self._f_result.setCurrentText(result)
        pas = (result == "Pass")
        fail = (result == "Fail")
        self._ee_pass_btn.setStyleSheet(
            (f"QPushButton{{background:{CLR_PASSBG};color:{CLR_PASS};border:1px solid {CLR_PASS};"
             f"border-radius:4px 0 0 4px;padding:7px 22px;font-weight:700;font-size:10pt;}}")
            if pas else
            (f"QPushButton{{background:{CLR_CARD};color:{CLR_TEXT};border:1px solid {CLR_BORDER};"
             f"border-radius:4px 0 0 4px;padding:7px 22px;font-size:10pt;}}"
             f"QPushButton:hover{{background:{CLR_BFAINT};}}"))
        self._ee_fail_btn.setStyleSheet(
            (f"QPushButton{{background:{CLR_FAILBG};color:{CLR_FAIL};border:1px solid {CLR_FAIL};"
             f"border-radius:0 4px 4px 0;padding:7px 22px;font-weight:700;font-size:10pt;}}")
            if fail else
            (f"QPushButton{{background:{CLR_CARD};color:{CLR_TEXT};border:1px solid {CLR_BORDER};"
             f"border-radius:0 4px 4px 0;padding:7px 22px;font-size:10pt;}}"
             f"QPushButton:hover{{background:{CLR_BFAINT};}}"))

    def _refresh_billing_display(self):
        while self._ee_billing_lines.count():
            item = self._ee_billing_lines.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        svcs = get_services(self.db)
        for idx, d in enumerate(self._lines_data):
            rw = QWidget(); rw.setStyleSheet("background:transparent;")
            rv = QVBoxLayout(rw); rv.setContentsMargins(0,2,0,4); rv.setSpacing(3)
            id_txt = d.get("vin","") or d.get("plate","")
            if id_txt:
                id_lbl = QLabel(id_txt)
                id_lbl.setStyleSheet(f"color:{CLR_TSUB};font-size:8pt;background:transparent;")
                rv.addWidget(id_lbl)
            rh = QHBoxLayout(); rh.setContentsMargins(0,0,0,0); rh.setSpacing(6)
            sl = QLabel(d["service"])
            sl.setStyleSheet(f"color:{CLR_TEXT};font-size:10pt;background:transparent;")
            rh.addWidget(sl, 1)
            is_cc_fee = (d.get("service","") == "Credit Card Fee")
            if is_cc_fee:
                # Auth # field for credit card authorization number
                auth_e = QLineEdit(d.get("cert",""))
                auth_e.setPlaceholderText("Auth # (optional)")
                auth_e.setFixedWidth(130)
                auth_e.setStyleSheet(
                    f"QLineEdit{{background:#F8FAFD;border:1px solid #B8CCE8;"
                    f"border-radius:4px;padding:2px 6px;color:{CLR_BLUE};font-size:9pt;}}")
                auth_e.textChanged.connect(lambda txt, i=idx: self._set_line_cert(i, txt))
                rh.addWidget(auth_e)
            else:
                # Pass/Fail/Retest dropdown
                res_cmb = QComboBox()
                _res_opts = ["Pass","Fail","Retest"]
                res_cmb.addItems(_res_opts)
                res_cmb.setCurrentText(d.get("result","Pass") or "Pass")
                res_cmb.setFixedWidth(82)
                res_cmb.setStyleSheet(
                    f"QComboBox{{background:#F8FAFD;border:1px solid #B8CCE8;"
                    f"border-radius:4px;padding:2px 6px;color:{CLR_BLUE};font-size:9pt;}}")
                res_cmb.activated.connect(lambda index, i=idx, opts=_res_opts: self._change_line_result(i, opts[index]) if index < len(opts) else None)
                rh.addWidget(res_cmb)
                # Cert # field (show if service has cert_fee)
                cert_fee = d.get("cert_fee", 0) or float(svcs.get(d.get("service",""), {}).get("cert_fee", 0))
                if cert_fee > 0:
                    cert_e = QLineEdit(d.get("cert",""))
                    cert_e.setPlaceholderText("Cert #")
                    cert_e.setFixedWidth(88)
                    cert_e.setStyleSheet(
                        f"QLineEdit{{background:#F8FAFD;border:1px solid #B8CCE8;"
                        f"border-radius:4px;padding:2px 6px;color:{CLR_BLUE};font-size:9pt;}}")
                    cert_e.textChanged.connect(lambda txt, i=idx: self._set_line_cert(i, txt))
                    rh.addWidget(cert_e)
            pl = QLabel(f"${d['price']:,.2f}")
            pl.setStyleSheet(f"color:{CLR_TEXT};font-size:10pt;background:transparent;")
            can_remove = self._editing_is_estimate or self._editing_id is None
            if can_remove:
                rm = QPushButton("✕ Remove"); rm.setFixedHeight(22)
                rm.setStyleSheet(
                    f"QPushButton{{background:transparent;color:{CLR_TSUB};border:1px solid {CLR_BORDER};"
                    f"border-radius:4px;padding:1px 7px;font-size:8pt;}}"
                    f"QPushButton:hover{{background:{CLR_FAILBG};color:{CLR_FAIL};border-color:{CLR_FAIL};}}")
                rm.clicked.connect(lambda checked, i=idx: self._remove_billing_line(i))
                rh.addWidget(pl); rh.addWidget(rm)
            else:
                rh.addWidget(pl)
            rv.addLayout(rh)
            rw.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            rw.customContextMenuRequested.connect(lambda pos, i=idx: self._billing_row_context_menu(i))
            self._ee_billing_lines.addWidget(rw)
        total = sum(d["price"] for d in self._lines_data)
        self._ee_bill_total.setText(f"${total:,.2f}")

    def _change_line_result(self, idx, new_result):
        if idx >= len(self._lines_data): return
        d = self._lines_data[idx]
        d["result"] = new_result
        new_price = max(self._get_service_price(d["service"], new_result) - d.get("discount", 0), 0)
        d["price"] = new_price
        if new_result in ("Fail", "Retest"):
            d["cert"] = ""
        QTimer.singleShot(0, self._update_total)

    def _remove_billing_line(self, idx):
        if not (self._editing_is_estimate or self._editing_id is None): return
        if idx < len(self._lines_data):
            self._lines_data.pop(idx)
            if idx < self._lines_table.rowCount():
                self._lines_table.removeRow(idx)
        self._update_total()

    def _billing_row_context_menu(self, idx):
        can_edit = self._editing_is_estimate or self._editing_id is None
        menu = QMenu(self)
        if can_edit:
            act_remove = menu.addAction("Remove Line")
        else:
            act_locked = menu.addAction("Invoice finalized — no changes allowed")
            act_locked.setEnabled(False)
        from PyQt6.QtGui import QCursor
        chosen = menu.exec(QCursor.pos())
        if can_edit and chosen == act_remove:
            self._remove_billing_line(idx)

    def _set_line_cert(self, idx, txt):
        if idx < len(self._lines_data):
            self._lines_data[idx]["cert"] = txt

    def _on_show_estimate_entry(self, invoice_id=None):
        # Apply saved zoom to graphics view
        if hasattr(self, '_ee_view'):
            scale = self._zoom.get("estimate_entry", 10) / 10.0
            self._ee_view.resetTransform()
            self._ee_view.scale(scale, scale)
        self._refresh_acct_id_dropdown()
        svc_names = list(get_services(self.db).keys())
        self._f_svc.clear(); self._f_svc.addItems(svc_names)
        if hasattr(self, '_veh_svc_cmb'):
            self._veh_svc_cmb.clear(); self._veh_svc_cmb.addItems(svc_names)
        if invoice_id:
            self._load_invoice_into_form(invoice_id)
            # Sync the new visible header label
            if hasattr(self, '_ee_hdr_lbl'):
                typ = self._ee_type_lbl.text().title()
                num = self._ee_num_lbl.text().strip()
                self._ee_hdr_lbl.setText(f"{typ}{(' ' + num) if num else ''}")
            # Sync company display if only first/last was set
            if hasattr(self, '_f_company') and not self._f_company.text().strip():
                full = f"{self._f_first.text()} {self._f_last.text()}".strip()
                if full: self._f_company.setText(full)
        else:
            self._clear_form()
            self._set_page_title("New Invoice")

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

    def _apply_cust_discount(self, cust_row):
        """Store discount from customer record and show it on the form."""
        if not cust_row:
            self._cust_discount_pct  = 0.0
            self._cust_discount_type = "PERCENT"
            if hasattr(self, '_disc_info_lbl'): self._disc_info_lbl.setText("")
            return
        self._cust_discount_pct  = float(cust_row["discount_percent"] or 0.0)
        self._cust_discount_type = (cust_row["discount_type"] or "PERCENT").upper()
        if hasattr(self, '_disc_info_lbl'):
            if self._cust_discount_pct:
                if self._cust_discount_type == "FLAT":
                    self._disc_info_lbl.setText(f"Customer discount: ${self._cust_discount_pct:.2f} off each line")
                else:
                    self._disc_info_lbl.setText(f"Customer discount: {self._cust_discount_pct:.0f}% — applied automatically")
            else:
                self._disc_info_lbl.setText("")

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
        self._apply_cust_discount(cust)

    def _vin_lookup(self):
        vin = self._f_vin.text().strip().upper()
        if len(vin) != 17: return
        # Local DB first
        vrow = self.db.execute("SELECT * FROM vehicles WHERE vin=? LIMIT 1",(vin,)).fetchone()
        if vrow:
            if not self._f_year.text(): self._f_year.setText(vrow["year"] or "")
            if not self._f_make.text(): self._f_make.setText(vrow["make"] or "")
            if not self._f_model.text(): self._f_model.setText(vrow["model"] or "")
            self._prefill_test_interval(vrow)
            # When plate is absent/NONE, also fill customer via VIN
            plate = self._f_plate.text().strip().upper()
            if not plate or plate == "NONE":
                cust_row = self.db.execute(
                    "SELECT c.* FROM customers c "
                    "JOIN invoices i ON i.customer_id=c.customer_id "
                    "WHERE UPPER(i.vin)=? ORDER BY i.invoice_date DESC, i.updated_at DESC LIMIT 1",
                    (vin,)).fetchone()
                if not cust_row:
                    cid = vrow["customer_id"] if "customer_id" in vrow.keys() else None
                    if cid:
                        cust_row = self.db.execute("SELECT * FROM customers WHERE customer_id=?",(cid,)).fetchone()
                if cust_row:
                    def si(w2, val):
                        if not w2.text().strip() and val: w2.setText(val)
                    si(self._f_first,   cust_row["first_name"])
                    si(self._f_last,    cust_row["last_name"])
                    si(self._f_company, cust_row["company_name"])
                    si(self._f_phone,   cust_row["phone"])
                    si(self._f_email,   cust_row["email"])
                    si(self._f_addr,    cust_row["address"])
                    si(self._f_city,    cust_row["city"])
                    si(self._f_state,   cust_row["state"])
                    si(self._f_zip,     cust_row["zip"])
                    self._apply_cust_discount(cust_row)
            return
        if requests:
            self._vin_worker = VinWorker(vin); self._vin_worker.done.connect(self._vin_done); self._vin_worker.start()

    def _vin_done(self, yr, mk, md):
        if yr and not self._f_year.text(): self._f_year.setText(yr)
        if mk and not self._f_make.text(): self._f_make.setText(mk)
        if md and not self._f_model.text(): self._f_model.setText(md)

    def _plate_lookup(self):
        plate = self._f_plate.text().strip().upper()
        if not plate or plate == "NONE": return
        vin = self._f_vin.text().strip().upper()
        if vin:
            vrow = self.db.execute(
                "SELECT * FROM vehicles WHERE UPPER(plate)=? AND vin=? ORDER BY updated_at DESC LIMIT 1",
                (plate, vin)).fetchone()
            if not vrow:
                vrow = self.db.execute(
                    "SELECT * FROM vehicles WHERE UPPER(plate)=? ORDER BY updated_at DESC LIMIT 1",
                    (plate,)).fetchone()
        else:
            vrow = self.db.execute(
                "SELECT * FROM vehicles WHERE UPPER(plate)=? ORDER BY updated_at DESC LIMIT 1",
                (plate,)).fetchone()
        if not vrow: return
        if not self._f_vin.text():   self._f_vin.setText(vrow["vin"] or "")
        if not self._f_year.text():  self._f_year.setText(vrow["year"] or "")
        if not self._f_make.text():  self._f_make.setText(vrow["make"] or "")
        if not self._f_model.text(): self._f_model.setText(vrow["model"] or "")
        self._prefill_test_interval(vrow)
        # Look up customer from most recent invoice for this plate+VIN combo, fall back to plate only
        if vin:
            cust_row = self.db.execute(
                "SELECT c.* FROM customers c "
                "JOIN invoices i ON i.customer_id=c.customer_id "
                "WHERE UPPER(i.plate)=? AND UPPER(i.vin)=? ORDER BY i.invoice_date DESC, i.updated_at DESC LIMIT 1",
                (plate, vin)).fetchone()
            if not cust_row:
                cust_row = self.db.execute(
                    "SELECT c.* FROM customers c "
                    "JOIN invoices i ON i.customer_id=c.customer_id "
                    "WHERE UPPER(i.plate)=? ORDER BY i.invoice_date DESC, i.updated_at DESC LIMIT 1",
                    (plate,)).fetchone()
        else:
            cust_row = self.db.execute(
                "SELECT c.* FROM customers c "
                "JOIN invoices i ON i.customer_id=c.customer_id "
                "WHERE UPPER(i.plate)=? ORDER BY i.invoice_date DESC, i.updated_at DESC LIMIT 1",
                (plate,)).fetchone()
        if not cust_row:
            cid = vrow["customer_id"] if "customer_id" in vrow.keys() else None
            if cid:
                cust_row = self.db.execute("SELECT * FROM customers WHERE customer_id=?",(cid,)).fetchone()
        if cust_row:
            def si(w2, val):
                if not w2.text().strip() and val: w2.setText(val)
            si(self._f_first,   cust_row["first_name"])
            si(self._f_last,    cust_row["last_name"])
            si(self._f_company, cust_row["company_name"])
            si(self._f_phone,   cust_row["phone"])
            si(self._f_email,   cust_row["email"])
            si(self._f_addr,    cust_row["address"])
            si(self._f_city,    cust_row["city"])
            si(self._f_state,   cust_row["state"])
            si(self._f_zip,     cust_row["zip"])
            self._apply_cust_discount(cust_row)

    def _prefill_test_interval(self, vrow):
        interval = vrow["test_interval_days"] if "test_interval_days" in vrow.keys() else None
        for i, (_, days) in enumerate(_INTERVAL_OPTS):
            if days == interval:
                self._f_test_interval.setCurrentIndex(i)
                break
        else:
            self._f_test_interval.setCurrentIndex(0)
        # Prefill the next due date from the vehicle record if available
        if hasattr(self, '_f_next_due_date'):
            next_due = (vrow["next_test_due"] if "next_test_due" in vrow.keys() else None) or ""
            if next_due:
                qd = QDate.fromString(next_due.strip(), "yyyy-MM-dd")
                if qd.isValid():
                    self._f_next_due_date.setDate(qd)
                    return
            # Fall back to today if no due date stored
            self._f_next_due_date.setDate(QDate.currentDate())

    def _on_no_plate_toggled(self, checked):
        if checked:
            self._f_plate.setText("NONE")
            self._f_plate.setEnabled(False)
        else:
            if self._f_plate.text().strip().upper() == "NONE":
                self._f_plate.clear()
            self._f_plate.setEnabled(True)

    def _zip_lookup(self):
        z = self._f_zip.text().strip()
        if len(z) != 5 or not z.isdigit(): return
        if self._f_city.text().strip() and self._f_state.text().strip(): return
        self._zip_worker = ZipWorker(z)
        self._zip_worker.done.connect(lambda c,s: (self._f_city.setText(c) if c else None, self._f_state.setText(s) if s else None))
        self._zip_worker.start()

    def _fmt_phone(self):
        self._f_phone.setText(format_phone(self._f_phone.text()))

    def _pick_inv_date(self):
        dlg = QDialog(self); dlg.setWindowTitle("Pick Date")
        lay = QVBoxLayout(dlg)
        cal = QCalendarWidget(); cal.setSelectedDate(QDate.fromString(self._inv_date_e.text(),"yyyy-MM-dd"))
        lay.addWidget(cal)
        ok = btn("OK","primary"); ok.clicked.connect(dlg.accept); lay.addWidget(ok)
        if dlg.exec():
            self._inv_date_e.setText(cal.selectedDate().toString("yyyy-MM-dd"))
            if hasattr(self, '_f_next_due_date') and self._f_test_interval.currentIndex() == 0:
                self._f_next_due_date.setDate(cal.selectedDate())

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
        cert_fee = float(s.get("cert_fee", 0))
        if cert_fee > 0 and result == "Pass":
            base += cert_fee
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
        svcs_map = get_services(self.db); s = svcs_map.get(svc, {})
        cert_fee = float(s.get("cert_fee", 0))
        base_price = self._get_service_price(svc, result)
        # Auto-apply customer discount if no manual discount entered
        if disc == 0.0:
            pct  = getattr(self, '_cust_discount_pct',  0.0)
            dtype = getattr(self, '_cust_discount_type', 'PERCENT')
            if pct:
                disc = round(base_price * pct / 100, 2) if dtype != 'FLAT' else pct
        price = max(base_price - disc, 0)
        d = dict(vin=vin,plate=plate,odometer=odo,year=year,make=make[:8],model=model[:10],
                 service=svc,result=result,cert=cert,discount=disc,price=price,
                 cert_fee=cert_fee,remote_item_id="")
        self._lines_data.append(d)
        r = self._lines_table.rowCount(); self._lines_table.insertRow(r)
        for col, val in enumerate([vin,svc,result,cert,f"${disc:.2f}",f"${price:.2f}"]):
            self._lines_table.setItem(r,col,QTableWidgetItem(val))
        self._update_total(); self._payment_changed()
        self._f_cert.clear(); self._f_disc.setText("0")
        for w2 in (self._f_plate, self._f_vin, self._f_year, self._f_make, self._f_model, self._f_odo):
            w2.clear()
        if hasattr(self, '_no_plate_cb') and self._no_plate_cb.isChecked():
            self._f_plate.setText("NONE")
        if hasattr(self, '_veh_svc_cmb'): self._veh_svc_cmb.setCurrentIndex(0)

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

    def _line_context_menu(self, pos):
        r = self._lines_table.indexAt(pos).row()
        if r < 0 or r >= len(self._lines_data): return
        can_edit = self._editing_is_estimate or self._editing_id is None
        menu = QMenu(self)
        if can_edit:
            act_remove = menu.addAction("Remove Line")
            act_edit   = menu.addAction("Edit Line (load back)")
        else:
            act_locked = menu.addAction("Invoice finalized — no changes allowed")
            act_locked.setEnabled(False)
        chosen = menu.exec(self._lines_table.viewport().mapToGlobal(pos))
        if not can_edit: return
        self._lines_table.setCurrentCell(r, 0)
        if chosen == act_remove:
            self._remove_line()
        elif chosen == act_edit:
            self._edit_line()

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
                     result="",cert="",discount=0.0,price=fee,cert_fee=0,remote_item_id="")
            self._lines_data.append(d)
            r = self._lines_table.rowCount(); self._lines_table.insertRow(r)
            for col,val in enumerate(["","Credit Card Fee","","","$0.00",f"${fee:.2f}"]):
                self._lines_table.setItem(r,col,QTableWidgetItem(val))
        self._update_total()

    def _update_total(self):
        total = sum(d["price"] for d in self._lines_data)
        self._total_lbl.setText(f"Total: ${total:,.2f}")
        self._ee_total_big.setText(f"TOTAL: ${total:,.2f}")
        if hasattr(self, '_ee_billing_lines'):
            self._refresh_billing_display()

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
        self._editing_id = None; self._editing_is_estimate = True; self._customer_touched = False
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
        if hasattr(self, '_ee_hdr_lbl'): self._ee_hdr_lbl.setText("New Invoice")
        if hasattr(self, '_ee_toggle_result'): self._ee_toggle_result("")
        if getattr(self, '_ee_tech', None): self._ee_tech.clear()
        if hasattr(self, '_f_test_interval'): self._f_test_interval.setCurrentIndex(0)
        if hasattr(self, '_f_next_due_date'): self._f_next_due_date.setDate(QDate.currentDate())
        if hasattr(self, '_no_plate_cb'): self._no_plate_cb.setChecked(False)
        self._cust_discount_pct  = 0.0
        self._cust_discount_type = "PERCENT"
        if hasattr(self, '_disc_info_lbl'): self._disc_info_lbl.setText("")

    def _load_invoice_into_form(self, invoice_id):
        self._clear_form(); self._editing_id = invoice_id
        inv = self.db.execute("SELECT * FROM invoices WHERE invoice_id=?",(invoice_id,)).fetchone()
        if not inv: return
        self._editing_is_estimate = bool(inv["is_estimate"])
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
        num = inv["invoice_number"] or ""
        num_str = f" #{num}" if num else " #PENDING"
        if inv["is_estimate"]:
            self._ee_type_lbl.setText("ESTIMATE")
            self._set_page_title(f"Estimate{num_str}")
        else:
            self._ee_type_lbl.setText("INVOICE")
            self._set_page_title(f"Invoice{num_str}")
        self._ee_num_lbl.setText(f"  {num_str}")
        # Populate vehicle from first line
        if lines:
            l = lines[0]
            self._f_vin.setText(l["vin"] or inv["vin"] or "")
            plate_val = l["plate"] or inv["plate"] or ""
            self._f_plate.setText(plate_val)
            if hasattr(self, '_no_plate_cb') and plate_val.upper() == "NONE":
                self._no_plate_cb.setChecked(True)
            self._f_year.setText(l["year"] or inv["year"] or "")
            self._f_make.setText(l["make"] or inv["make"] or "")
            self._f_model.setText(l["model"] or inv["model"] or "")
            plate = (l["plate"] or inv["plate"] or "").upper()
            vin   = (l["vin"]   or inv["vin"]   or "").upper()
            if plate:
                vrow = self.db.execute("SELECT * FROM vehicles WHERE UPPER(plate)=? LIMIT 1",(plate,)).fetchone()
                if vrow: self._prefill_test_interval(vrow)
            elif vin:
                vrow = self.db.execute("SELECT * FROM vehicles WHERE UPPER(vin)=? LIMIT 1",(vin,)).fetchone()
                if vrow: self._prefill_test_interval(vrow)

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

        _existing_cid = get_or_create_customer_id(self.db, fd["first"], fd["last"], fd["company"])
        _existing_cust = self.db.execute("SELECT discount_percent, discount_type FROM customers WHERE customer_id=?", (_existing_cid,)).fetchone()
        _disc      = (_existing_cust["discount_percent"] or 0.0) if _existing_cust else 0.0
        _disc_type = (_existing_cust["discount_type"]    or "PERCENT") if _existing_cust else "PERCENT"
        cid = upsert_customer(self.db,fd["first"],fd["last"],fd["company"],
                              phone=fd["phone"],email=fd["email"],address=fd["addr"],
                              city=fd["city"],state=fd["state"],zip_=fd["zip"],
                              discount_percent=_disc, discount_type=_disc_type)
        _seen = set()
        for d in self._lines_data:
            if not (d["vin"] or d["plate"]): continue
            key = d["vin"] or d["plate"]
            if key in _seen: continue; _seen.add(key)
            vid = upsert_vehicle(self.db,cid,d["vin"],d["plate"],d["make"],d["model"],d["year"])
            v_row = self.db.execute(
                "SELECT service_type FROM vehicles WHERE vehicle_id=?", (vid,)).fetchone()
            v_svc = (v_row["service_type"] or "") if v_row else ""
            enqueue(self.db,"vehicle","upsert",{"vehicle_id":vid,"customer_id":cid,"vin":d["vin"],
                "plate":d["plate"],"make":d["make"],"model":d["model"],"year":d["year"],
                "odometer":d["odometer"],"service_type":v_svc})

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
            _inv_row = self.db.execute("SELECT invoice_number FROM invoices WHERE invoice_id=?",(iid,)).fetchone()
            inv_num = _inv_row["invoice_number"] if _inv_row else 0
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

        # Build final lines: for invoice Pass lines with cert_fee, split into base + cert line
        final_lines = []
        for d in self._lines_data:
            cert_fee = d.get("cert_fee", 0)
            if not is_estimate and cert_fee > 0 and (d.get("result","") or "").upper() == "PASS":
                base_line = dict(d); base_line["price"] = max(d["price"] - cert_fee, 0)
                cert_line = dict(vin=d["vin"],plate=d["plate"],odometer="",
                                 year=d["year"],make=d["make"],model=d["model"],
                                 service="Certificate",result="Pass",cert=d.get("cert",""),
                                 discount=0.0,price=cert_fee,cert_fee=0,remote_item_id="")
                final_lines.append(base_line); final_lines.append(cert_line)
            else:
                final_lines.append(d)
        for d in final_lines:
            if not d.get("remote_item_id"): d["remote_item_id"] = str(uuid.uuid4())
            self.db.execute("""
                INSERT INTO invoice_lines
                (invoice_id,vin,plate,odometer,year,make,model,service,result,cert,discount,price,remote_item_id)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,(iid,d["vin"],d["plate"],d["odometer"],d["year"],d["make"],d["model"],
                 d["service"],d["result"],d["cert"],d["discount"],d["price"],d["remote_item_id"]))
        self.db.commit()

        # Update vehicle next_test_due if a due date is selected (interval != "No reminder")
        if not is_estimate and hasattr(self, '_f_test_interval') and hasattr(self, '_f_next_due_date'):
            interval_idx = self._f_test_interval.currentIndex()
            interval_days = _INTERVAL_OPTS[interval_idx][1]
            if interval_days is not None and (plate or vin):
                try:
                    # Use the directly-edited date from the calendar picker
                    qd = self._f_next_due_date.date()
                    next_due = f"{qd.year():04d}-{qd.month():02d}-{qd.day():02d}"
                    if plate:
                        vid_row = self.db.execute("SELECT vehicle_id FROM vehicles WHERE UPPER(plate)=? LIMIT 1",(plate.upper(),)).fetchone()
                    else:
                        vid_row = self.db.execute("SELECT vehicle_id FROM vehicles WHERE UPPER(vin)=? LIMIT 1",(vin.upper(),)).fetchone()
                    if vid_row:
                        self.db.execute("UPDATE vehicles SET test_interval_days=?,next_test_due=? WHERE vehicle_id=?",
                                        (interval_days, next_due, vid_row["vehicle_id"]))
                        self.db.commit()
                        enqueue(self.db,"vehicle","upsert",{"vehicle_id":vid_row["vehicle_id"],
                            "customer_id":cid,
                            "vin":vin,"plate":plate,"year":yr,"make":mk,"model":md,
                            "test_interval_days":interval_days,"next_test_due":next_due,
                            "next_due":next_due,"service_interval_days":interval_days})
                except Exception: pass

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

        # Enqueue for sync — invoice header (full fields so mobile has complete record)
        enqueue(self.db,"invoice","upsert",{"invoice_id":iid,"invoice_number":inv_num or 0,
            "customer_id":cid,"customer_name":cname,"first_name":fd["first"],"last_name":fd["last"],
            "company_name":fd["company"],"invoice_date":fd["date"],"plate":plate,"vin":vin,
            "year":yr,"make":mk,"model":md,"amount_cents":total_cents,"payment_method":fd["pay"],
            "status":status,"notes":fd["notes"],"is_estimate":1 if is_estimate else 0,
            "finalized":0 if is_estimate else 1,
            "account_id":effective_acct_id or "",
            "po_number":fd.get("po","") or "",
            "owner_first":fd["first"],"owner_last":fd["last"]})
        # Enqueue invoice line items so mobile shows individual services
        for d in final_lines:
            enqueue(self.db,"invoice_item","upsert",{
                "item_id":     d["remote_item_id"],
                "invoice_id":  iid,
                "name":        d["service"],
                "service":     d["service"],
                "qty":         1,
                "unit_price_cents": int(round(d["price"] * 100)),
                "discount":    d.get("discount",0),
                "discount_cents": int(round(d.get("discount",0) * 100)),
                "result":      d.get("result",""),
                "cert":        d.get("cert",""),
                "odometer":    d.get("odometer",""),
                "vin":         d.get("vin",""),
                "plate":       d.get("plate",""),
                "year":        d.get("year",""),
                "make":        d.get("make",""),
                "model":       d.get("model",""),
                "tech_name":   "",
            })
        # Enqueue customer so mobile has full contact details
        _cust_row = self.db.execute("SELECT * FROM customers WHERE customer_id=?", (cid,)).fetchone()
        if _cust_row:
            enqueue(self.db,"customer","upsert",{
                "customer_id":     cid,
                "first_name":      _cust_row["first_name"]  or "",
                "last_name":       _cust_row["last_name"]   or "",
                "company_name":    _cust_row["company_name"] or "",
                "phone":           _cust_row["phone"]        or "",
                "email":           _cust_row["email"]        or "",
                "address":         _cust_row["address"]      or "",
                "city":            _cust_row["city"]         or "",
                "state":           _cust_row["state"]        or "",
                "zip":             _cust_row["zip"]          or "",
                "discount_percent": float(_cust_row["discount_percent"] or 0),
                "discount_type":   _cust_row["discount_type"] or "PERCENT",
            })

        ps = get_printer_setting(self.db)
        if is_estimate:
            import tempfile as _tf
            _fd, pdf = _tf.mkstemp(suffix=".pdf", prefix="EST_")
            os.close(_fd)
        else:
            pdf = build_invoice_pdf_path(self.inv_dir,fd["date"],company=fd["company"],
                                         first=fd["first"],last=fd["last"],customer_name=cname,
                                         is_estimate=False,inv_num=inv_num)
        generate_invoice_pdf(iid,self.db,pdf)
        self._editing_id = iid
        self._ee_type_lbl.setText("ESTIMATE" if is_estimate else "INVOICE")

        # Background sync to get real invoice number
        def _bg_sync():
            try: SYNC._flush(); SYNC._pull()
            except Exception: pass
        threading.Thread(target=_bg_sync, daemon=True).start()

        if ps.get("auto_print"):
            pname = (ps.get("printer_name") or "").strip()
            print_pdf(pdf, printer_name=pname, copies=int(ps.get("copies",2)), parent_widget=self, silent=True)
        if is_estimate:
            try: os.remove(pdf)
            except Exception: pass
        self.show_screen("doc_list")

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: ACCOUNTS / AR
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _build_account_setup_screen(self):
        w = QWidget(); self._screens["account_setup"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        self._stack.addWidget(w)

        # Toolbar
        tb = QWidget(); tb.setStyleSheet(f"background:{CLR_CARD};border-bottom:1px solid {CLR_BORDER};")
        tb_h = QHBoxLayout(tb); tb_h.setContentsMargins(10,6,10,6); tb_h.setSpacing(8)
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

    def _on_show_account_setup(self, company_name=None):
        self._set_page_title("Accounts")
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
                "partial_json": "{}",
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

    def _build_vehicles_due_screen(self):
        w = QWidget(); self._screens["vehicles_due"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(14,12,14,12); lay.setSpacing(8)
        self._stack.addWidget(w)

        # ── Header row ───────────────────────────────────────────────
        hdr_row = QHBoxLayout()
        ttl = QLabel("Vehicles Due"); ttl.setStyleSheet("font-size:15pt; font-weight:700;")
        hdr_row.addWidget(ttl)
        hdr_row.addStretch()
        ref_b = QPushButton("⟳ Refresh"); ref_b.setObjectName("secondary")
        hdr_row.addWidget(ref_b)
        lay.addLayout(hdr_row)

        # ── Filter row ───────────────────────────────────────────────
        filt_row = QHBoxLayout()
        filt_row.addWidget(QLabel("Show:"))
        self._vd_cmb = QComboBox()
        for fk, fl in [("30days","30 Days"),("60days","60 Days"),("90days","90 Days")]:
            self._vd_cmb.addItem(fl, fk)
        filt_row.addWidget(self._vd_cmb); filt_row.addStretch()
        self._vd_summary = QLabel("")
        self._vd_summary.setStyleSheet(f"color:{PRIMARY}; font-weight:bold;")
        filt_row.addWidget(self._vd_summary)
        lay.addLayout(filt_row)

        # ── Table ────────────────────────────────────────────────────
        self._vd_tbl = QTableWidget(0, 8)
        self._vd_tbl.setHorizontalHeaderLabels(
            ["Due Date","Days","Plate","Year","Make / Model","Customer","Email","Reminder"])
        self._vd_tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._vd_tbl.setAlternatingRowColors(True)
        self._vd_tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        hh = self._vd_tbl.horizontalHeader()
        hh.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        self._vd_tbl.setColumnWidth(7, 80)
        lay.addWidget(self._vd_tbl)

        def _populate(fkey):
            today = datetime.now().date()
            drop_before = (today - timedelta(days=274)).strftime("%Y-%m-%d")
            if fkey == '60days':
                lo, hi = drop_before, (today + timedelta(days=60)).strftime("%Y-%m-%d")
            elif fkey == '90days':
                lo, hi = drop_before, (today + timedelta(days=90)).strftime("%Y-%m-%d")
            else:  # 30days default
                lo, hi = drop_before, (today + timedelta(days=30)).strftime("%Y-%m-%d")

            rows = self.db.execute("""
                SELECT v.vehicle_id, v.plate, v.vin, v.year, v.make, v.model,
                       v.next_test_due, c.first_name, c.last_name, c.company_name,
                       c.email
                FROM vehicles v
                LEFT JOIN customers c ON v.customer_id = c.customer_id
                WHERE v.deleted = 0 AND v.next_test_due != '' AND v.next_test_due IS NOT NULL
                  AND v.next_test_due >= ? AND v.next_test_due <= ?
                ORDER BY v.next_test_due ASC
            """, (lo, hi)).fetchall()

            self._vd_tbl.setRowCount(0)
            overdue_ct = 0
            for vr in rows:
                due_str = (vr["next_test_due"] or "").strip()
                try:
                    due_date = datetime.strptime(due_str, "%Y-%m-%d").date()
                    days_left = (due_date - today).days
                    due_display = due_date.strftime("%m/%d/%Y")
                    days_display = f"{days_left}d" if days_left >= 0 else f"{abs(days_left)}d overdue"
                    if days_left < 0: overdue_ct += 1
                except ValueError:
                    due_display = due_str; days_left = 999; days_display = ""

                mk_mod = " ".join(filter(None, [vr["make"], vr["model"]]))
                cname  = vr["company_name"] or f"{vr['first_name'] or ''} {vr['last_name'] or ''}".strip()
                email  = vr["email"] or ""

                r_idx = self._vd_tbl.rowCount(); self._vd_tbl.insertRow(r_idx)
                for col, val in enumerate([due_display, days_display, vr["plate"] or "",
                                           str(vr["year"] or ""), mk_mod, cname, email]):
                    item = QTableWidgetItem(val)
                    if days_left < 0:
                        item.setForeground(QColor("#B91C1C"))
                        if col == 1: item.setFont(QFont("", -1, QFont.Weight.Bold))
                    elif days_left <= 30:
                        item.setForeground(QColor("#D97706"))
                    self._vd_tbl.setItem(r_idx, col, item)

                # Bell reminder button (email only on desktop)
                bell_b = QPushButton("🔔")
                bell_b.setToolTip("Send email reminder")
                bell_b.setFixedWidth(50)
                bell_b.setStyleSheet("border:none; font-size:14pt;")
                if email:
                    veh_label = f"{vr['year'] or ''} {mk_mod}".strip()
                    subj = f"Your smog check is coming up — {veh_label}"
                    body = (f"Hi {vr['first_name'] or cname},\n\n"
                            f"This is a reminder that your {veh_label} "
                            f"(plate {vr['plate'] or 'N/A'}) is due for a smog check on "
                            f"{due_display}.\n\nGive us a call to schedule!\n\nBlue Sky Smog")
                    import urllib.parse as _up
                    mailto = (f"mailto:{_up.quote(email, safe='@')}"
                              f"?subject={_up.quote(subj)}"
                              f"&body={_up.quote(body)}")
                    bell_b.clicked.connect(
                        lambda _, url=mailto: (
                            os.startfile(url) if sys.platform == "win32"
                            else subprocess.Popen(["open", url])
                        ))
                else:
                    bell_b.setEnabled(False)
                    bell_b.setToolTip("No email on file")
                self._vd_tbl.setCellWidget(r_idx, 7, bell_b)

            self._vd_summary.setText(
                f"Vehicles: {self._vd_tbl.rowCount()}   |   Overdue: {overdue_ct}")

        self._vd_populate = _populate
        self._vd_cmb.currentIndexChanged.connect(
            lambda _: _populate(self._vd_cmb.currentData()))
        ref_b.clicked.connect(lambda: _populate(self._vd_cmb.currentData()))

        # Initial population
        _populate("all")

    def _build_reports_screen(self):
        w = QWidget(); self._screens["reports"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        self._stack.addWidget(w)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        body = QWidget(); body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(14,12,14,12); body_lay.setSpacing(10)
        scroll.setWidget(body); lay.addWidget(scroll)

        # ── Period selector ──────────────────────────────────────────
        period_row = QHBoxLayout(); period_row.setSpacing(10)
        self._rpt_period = "w"   # "w" / "m" / "y"

        def _PTAB(key, text):
            b = QPushButton(text)
            b.setCheckable(True); b.setChecked(key == "w")
            b.setFixedHeight(30)
            b.setStyleSheet(f"""
                QPushButton {{ background:{CLR_SURFACE}; color:{CLR_TSUB};
                    border:1px solid {CLR_BORDER}; border-radius:5px; padding:0 14px; font-weight:500; }}
                QPushButton:checked {{ background:{CLR_BLUE}; color:white; border-color:{CLR_BLUE}; font-weight:700; }}
            """)
            b.clicked.connect(lambda _, k=key: self._rpt_set_period(k))
            return b
        self._rpt_tab_w = _PTAB("w","Week")
        self._rpt_tab_m = _PTAB("m","Month")
        self._rpt_tab_y = _PTAB("y","Year")
        for t in [self._rpt_tab_w, self._rpt_tab_m, self._rpt_tab_y]:
            period_row.addWidget(t)

        # Period picker button — opens dropdown matching mockup
        period_row.addSpacing(16)
        self._rpt_period_btn = QPushButton("This Week  ▾")
        self._rpt_period_btn.setStyleSheet(
            f"QPushButton{{background:{CLR_CARD};border:1px solid {CLR_BORDER};"
            f"border-radius:6px;color:{CLR_TEXT};padding:4px 14px;font-weight:600;font-size:10pt;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}")
        self._rpt_period_btn.clicked.connect(self._rpt_show_period_popup)
        period_row.addWidget(self._rpt_period_btn)
        period_row.addStretch()
        # Keep hidden QDateEdits for _run_report compatibility
        self._rpt_begin = QDateEdit(QDate(datetime.today().year, datetime.today().month, 1))
        self._rpt_begin.setCalendarPopup(True); self._rpt_begin.setVisible(False)
        self._rpt_end   = QDateEdit(QDate.currentDate())
        self._rpt_end.setCalendarPopup(True); self._rpt_end.setVisible(False)
        self._rpt_anchor = datetime.today()
        body_lay.addLayout(period_row)

        # ── Stat cards ───────────────────────────────────────────────
        CARD_STYLE = f"""
            QWidget {{ background:{CLR_CARD}; border:1px solid {CLR_BORDER};
                border-radius:7px; }}
        """
        stat_row = QHBoxLayout(); stat_row.setSpacing(8)
        def _stat_card(label, val_attr, sub_attr):
            card = QWidget(); card.setStyleSheet(CARD_STYLE)
            cl = QVBoxLayout(card); cl.setContentsMargins(12,10,12,10); cl.setSpacing(2)
            lbl = QLabel(label)
            lbl.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;font-weight:700;letter-spacing:1px;border:none;")
            val = QLabel("—")
            val.setStyleSheet(f"color:{CLR_TEXT};font-size:18pt;font-weight:700;border:none;")
            sub = QLabel("")
            sub.setStyleSheet(f"color:{CLR_TMUTED};font-size:9pt;border:none;")
            cl.addWidget(lbl); cl.addWidget(val); cl.addWidget(sub)
            setattr(self, val_attr, val); setattr(self, sub_attr, sub)
            stat_row.addWidget(card)
        _stat_card("INSPECTIONS",  "_rpt_s_insp",  "_rpt_s_insp_sub")
        _stat_card("PASS RATE",    "_rpt_s_pass",  "_rpt_s_pass_sub")
        _stat_card("REVENUE",      "_rpt_s_rev",   "_rpt_s_rev_sub")
        _stat_card("AVG / DAY",    "_rpt_s_avg",   "_rpt_s_avg_sub")
        body_lay.addLayout(stat_row)

        # ── Bar chart ────────────────────────────────────────────────
        chart_card = QWidget(); chart_card.setStyleSheet(
            f"background:{CLR_CARD};border:1px solid {CLR_BORDER};border-radius:7px;")
        cc_lay = QVBoxLayout(chart_card); cc_lay.setContentsMargins(0,0,0,0); cc_lay.setSpacing(0)

        ch_hdr = QWidget(); ch_hdr.setStyleSheet(
            f"background:{CLR_CARD};border-bottom:1px solid {CLR_BORDER};border-radius:7px 7px 0 0;")
        ch_h = QHBoxLayout(ch_hdr); ch_h.setContentsMargins(12,8,12,8); ch_h.setSpacing(8)
        self._rpt_chart_title = QLabel("Inspections by day")
        self._rpt_chart_title.setStyleSheet(f"color:{CLR_TEXT};font-weight:700;font-size:11pt;")
        ch_h.addWidget(self._rpt_chart_title)
        ch_h.addStretch()
        self._rpt_chart_hint = QLabel("Click a bar to filter below")
        self._rpt_chart_hint.setStyleSheet(f"color:{CLR_TMUTED};font-size:9pt;font-style:italic;")
        ch_h.addWidget(self._rpt_chart_hint)
        # Legend
        for color, text in [(CLR_BLUE,"Pass"),(CLR_FAIL,"Fail")]:
            dot = QLabel("●"); dot.setStyleSheet(f"color:{color};font-size:9pt;")
            ch_h.addWidget(dot); tl = QLabel(text)
            tl.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;"); ch_h.addWidget(tl)
        cc_lay.addWidget(ch_hdr)

        self._rpt_chart = BarChartWidget()
        self._rpt_chart.setMinimumHeight(160)
        self._rpt_chart.setStyleSheet(f"background:{CLR_CARD};border:none;")
        self._rpt_chart.barClicked.connect(self._rpt_bar_clicked)
        chart_wrap = QWidget()
        chart_wrap.setStyleSheet(f"background:{CLR_CARD};border:none;")
        cw_lay = QVBoxLayout(chart_wrap); cw_lay.setContentsMargins(10,8,10,8)
        cw_lay.addWidget(self._rpt_chart)
        cc_lay.addWidget(chart_wrap)
        body_lay.addWidget(chart_card)

        # ── Bottom tables row ────────────────────────────────────────
        tbl_row = QHBoxLayout(); tbl_row.setSpacing(10)

        def _small_card(title, icon_txt, cols, attr):
            tc = QWidget(); tc.setStyleSheet(
                f"background:{CLR_CARD};border:1px solid {CLR_BORDER};border-radius:7px;")
            tl = QVBoxLayout(tc); tl.setContentsMargins(0,0,0,0); tl.setSpacing(0)
            th = QWidget(); th.setStyleSheet(
                f"border-bottom:1px solid {CLR_BORDER};background:{CLR_CARD};border-radius:7px 7px 0 0;")
            thh = QHBoxLayout(th); thh.setContentsMargins(10,7,10,7); thh.setSpacing(6)
            il = QLabel(icon_txt); il.setStyleSheet(f"color:{CLR_BLUE};font-size:12pt;")
            tl2 = QLabel(title); tl2.setStyleSheet(f"color:{CLR_TEXT};font-weight:700;")
            thh.addWidget(il); thh.addWidget(tl2); thh.addStretch()
            self.__dict__[attr+"_sub"] = QLabel("All period")
            self.__dict__[attr+"_sub"].setStyleSheet(f"color:{CLR_TMUTED};font-size:9pt;")
            thh.addWidget(self.__dict__[attr+"_sub"])
            tl.addWidget(th)
            t = QTableWidget(0, len(cols)); t.setHorizontalHeaderLabels(cols)
            t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            t.verticalHeader().setVisible(False)
            t.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            t.setAlternatingRowColors(True)
            t.setStyleSheet(f"border:none;border-radius:0 0 7px 7px;")
            tl.addWidget(t); tbl_row.addWidget(tc)
            setattr(self, attr, t)
        _small_card("Top Customers", "#", ["Customer","Tests","Revenue"], "_rpt_cust_tbl")
        _small_card("By Test Type",  "✓",  ["Test Type","Count","Pass Rate"], "_rpt_tt_tbl")
        body_lay.addLayout(tbl_row)

        # ── Quick report buttons ─────────────────────────────────────
        qr_grp = QWidget(); qr_grp.setStyleSheet(
            f"background:{CLR_CARD};border:1px solid {CLR_BORDER};border-radius:7px;")
        qr_lay = QVBoxLayout(qr_grp); qr_lay.setContentsMargins(12,10,12,12); qr_lay.setSpacing(8)
        qr_lay.addWidget(QLabel("Quick Reports",
            styleSheet=f"color:{CLR_TEXT};font-weight:700;font-size:11pt;"))
        qr_grid = QGridLayout(); qr_grid.setSpacing(6)
        for i,(lbl_txt,cb) in enumerate([
            ("Month-To-Date",         lambda: self._run_report("mtd")),
            ("Daily",                 lambda: self._run_report("daily")),
            ("Weekly",                lambda: self._run_report("weekly")),
            ("By Payment Type",       lambda: self._run_report("by_pay")),
            ("By Service",            lambda: self._run_report("by_svc")),
            ("Open Estimates",        lambda: self._run_report("estimates")),
            ("Account Balances",      lambda: self._run_report("balances")),
            ("Vehicles Due (90 Days)",lambda: self._run_report("vehicles_due")),
        ]):
            b = btn(lbl_txt,"secondary"); b.setMinimumHeight(36); b.clicked.connect(cb)
            qr_grid.addWidget(b, i//4, i%4)
        qr_lay.addLayout(qr_grid)
        body_lay.addWidget(qr_grp)
        body_lay.addStretch()

    def _on_show_vehicles_due(self):
        self._set_page_title("Vehicles Due")
        self._vd_populate(self._vd_cmb.currentData() or "all")

    def _on_show_reports(self):
        self._set_page_title("Reports")
        self._rpt_refresh_dashboard()

    def _rpt_set_period(self, key):
        self._rpt_period = key
        self._rpt_anchor = datetime.today()
        for k, b in [("w",self._rpt_tab_w),("m",self._rpt_tab_m),("y",self._rpt_tab_y)]:
            b.setChecked(k == key)
        titles = {"w":"Inspections by day","m":"Inspections by date","y":"Inspections by month"}
        self._rpt_chart_title.setText(titles[key])
        self._rpt_refresh_dashboard()

    def _rpt_update_period_label(self, start, end):
        period = self._rpt_period
        if period == "w":
            lbl = f"{start.strftime('%b %d')} – {end.strftime('%b %d, %Y')}  ▾"
        elif period == "m":
            lbl = f"{start.strftime('%B %Y')}  ▾"
        else:
            lbl = f"{start.year}  ▾"
        self._rpt_period_btn.setText(lbl)

    def _rpt_show_period_popup(self):
        import calendar as _cal
        period  = self._rpt_period
        anchor  = getattr(self, '_rpt_anchor', datetime.today())
        today   = datetime.today()

        popup = QDialog(self, Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint)
        popup.setStyleSheet(
            f"QDialog{{background:{CLR_CARD};border:1px solid {CLR_BORDER};border-radius:8px;}}"
            f"QPushButton{{background:transparent;border:none;color:{CLR_TEXT};padding:4px 6px;"
            f"border-radius:4px;font-size:9pt;}}"
            f"QPushButton:hover{{background:{CLR_BFAINT};}}"
            f"QPushButton[selected=true]{{background:{CLR_BLUE};color:#ffffff;font-weight:600;}}"
            f"QPushButton[nav=true]{{font-size:11pt;padding:2px 8px;}}"
            f"QLabel{{color:{CLR_TSUB};font-size:8pt;font-weight:600;padding:2px 4px;}}"
        )

        outer = QVBoxLayout(popup)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(6)

        # ── WEEK mode: scrollable list of last 8 weeks ──────────────────────
        if period == "w":
            hdr = QLabel("SELECT A WEEK")
            hdr.setAlignment(Qt.AlignmentFlag.AlignCenter)
            outer.addWidget(hdr)

            cur_sun = anchor - timedelta(days=(anchor.weekday() + 1) % 7)
            weeks = []
            base = today - timedelta(days=(today.weekday() + 1) % 7)
            for i in range(8):
                s = base - timedelta(weeks=i)
                weeks.append(s)

            for s in weeks:
                e = s + timedelta(days=6)
                lbl_txt = f"{s.strftime('%b %d')} - {e.strftime('%b %d, %Y')}"
                b = QPushButton(lbl_txt)
                b.setProperty("selected", s.date() == cur_sun.date())
                b.style().unpolish(b); b.style().polish(b)
                def _pick_week(checked=False, _s=s):
                    self._rpt_anchor = _s
                    self._rpt_refresh_dashboard()
                    popup.close()
                b.clicked.connect(_pick_week)
                outer.addWidget(b)

        # ── MONTH mode: year nav + 4×3 grid ────────────────────────────────
        elif period == "m":
            nav_year = [anchor.year]  # mutable container for popup-scoped year nav

            month_grid_widget = QWidget()
            month_grid = QGridLayout(month_grid_widget)
            month_grid.setSpacing(4)

            MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun",
                           "Jul","Aug","Sep","Oct","Nov","Dec"]

            def _rebuild_month_grid():
                while month_grid.count():
                    item = month_grid.takeAt(0)
                    if item.widget(): item.widget().deleteLater()
                yr = nav_year[0]
                year_lbl.setText(str(yr))
                for idx, name in enumerate(MONTH_NAMES):
                    row, col = divmod(idx, 3)
                    m = idx + 1
                    is_future = (yr > today.year) or (yr == today.year and m > today.month)
                    is_sel    = (yr == anchor.year and m == anchor.month)
                    b = QPushButton(name)
                    b.setProperty("selected", is_sel)
                    if is_future:
                        b.setEnabled(False)
                        b.setStyleSheet(f"color:{CLR_TMUTED};")
                    else:
                        def _pick_month(checked=False, _y=yr, _m=m):
                            self._rpt_anchor = datetime(_y, _m, 1)
                            self._rpt_refresh_dashboard()
                            popup.close()
                        b.clicked.connect(_pick_month)
                    b.style().unpolish(b); b.style().polish(b)
                    month_grid.addWidget(b, row, col)

            nav_row = QHBoxLayout()
            prev_yr = QPushButton("◀"); prev_yr.setProperty("nav", True)
            year_lbl = QLabel(str(anchor.year))
            year_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            year_lbl.setStyleSheet(f"color:{CLR_TEXT};font-size:10pt;font-weight:600;")
            next_yr = QPushButton("▶"); next_yr.setProperty("nav", True)

            def _prev_year():
                nav_year[0] -= 1; _rebuild_month_grid()
            def _next_year():
                if nav_year[0] < today.year: nav_year[0] += 1; _rebuild_month_grid()

            prev_yr.clicked.connect(_prev_year)
            next_yr.clicked.connect(_next_year)
            nav_row.addWidget(prev_yr); nav_row.addWidget(year_lbl, 1)
            nav_row.addWidget(next_yr)
            outer.addLayout(nav_row)
            outer.addWidget(month_grid_widget)
            _rebuild_month_grid()

        # ── YEAR mode: decade nav + 3-col grid ─────────────────────────────
        else:
            decade_base = [(anchor.year // 10) * 10]  # mutable

            year_grid_widget = QWidget()
            year_grid = QGridLayout(year_grid_widget)
            year_grid.setSpacing(4)

            def _rebuild_year_grid():
                while year_grid.count():
                    item = year_grid.takeAt(0)
                    if item.widget(): item.widget().deleteLater()
                base = decade_base[0]
                decade_lbl.setText(f"{base}s")
                years = list(range(base, base + 10))
                for idx, yr in enumerate(years):
                    row, col = divmod(idx, 3) if idx < 9 else (3, 0)
                    is_future = yr > today.year
                    is_sel    = yr == anchor.year
                    b = QPushButton(str(yr))
                    b.setProperty("selected", is_sel)
                    if is_future:
                        b.setEnabled(False)
                        b.setStyleSheet(f"color:{CLR_TMUTED};")
                    else:
                        def _pick_year(checked=False, _y=yr):
                            self._rpt_anchor = datetime(_y, self._rpt_anchor.month, 1)
                            self._rpt_refresh_dashboard()
                            popup.close()
                        b.clicked.connect(_pick_year)
                    b.style().unpolish(b); b.style().polish(b)
                    year_grid.addWidget(b, row, col)

            nav_row = QHBoxLayout()
            prev_dec = QPushButton("◀"); prev_dec.setProperty("nav", True)
            decade_lbl = QLabel(f"{decade_base[0]}s")
            decade_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            decade_lbl.setStyleSheet(f"color:{CLR_TEXT};font-size:10pt;font-weight:600;")
            next_dec = QPushButton("▶"); next_dec.setProperty("nav", True)

            def _prev_decade():
                decade_base[0] -= 10; _rebuild_year_grid()
            def _next_decade():
                if decade_base[0] + 10 <= today.year: decade_base[0] += 10; _rebuild_year_grid()

            prev_dec.clicked.connect(_prev_decade)
            next_dec.clicked.connect(_next_decade)
            nav_row.addWidget(prev_dec); nav_row.addWidget(decade_lbl, 1)
            nav_row.addWidget(next_dec)
            outer.addLayout(nav_row)
            outer.addWidget(year_grid_widget)
            _rebuild_year_grid()

        popup.adjustSize()
        btn_rect = self._rpt_period_btn.rect()
        global_pos = self._rpt_period_btn.mapToGlobal(btn_rect.bottomLeft())
        popup.move(global_pos)
        popup.exec()

    def _rpt_refresh_dashboard(self):
        anchor = getattr(self, '_rpt_anchor', datetime.today())
        period = self._rpt_period

        if period == "w":
            start = anchor - timedelta(days=(anchor.weekday() + 1) % 7)  # Sunday
            end   = start + timedelta(days=6)
            dates = [start + timedelta(days=i) for i in range(7)]
            labels= ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]
        elif period == "m":
            start = anchor.replace(day=1)
            import calendar
            last_day = calendar.monthrange(anchor.year, anchor.month)[1]
            end   = anchor.replace(day=last_day)
            dates = [anchor.replace(day=d) for d in range(1, last_day+1)]
            labels= [str(d) for d in range(1, last_day+1)]
        else:  # year
            start = anchor.replace(month=1, day=1)
            end   = anchor.replace(month=12, day=31)
            dates = [anchor.replace(month=m, day=1) for m in range(1,13)]
            labels= ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

        self._rpt_update_period_label(start, end)

        start_s = start.strftime("%Y-%m-%d")
        end_s   = end.strftime("%Y-%m-%d")
        rows = self.db.execute(
            "SELECT i.invoice_id, i.invoice_date, i.test_result, i.amount_cents, "
            "i.customer_name, i.first_name, i.last_name, i.company_name FROM invoices i "
            "WHERE i.is_estimate=0 AND i.invoice_date>=? AND i.invoice_date<=? ORDER BY i.invoice_date",
            (start_s, end_s)
        ).fetchall()

        # Per-vehicle counts from invoice_lines (counts each vehicle on multi-vehicle invoices)
        inv_ids = [r["invoice_id"] for r in rows]
        line_results = {}  # invoice_id -> [result, ...]
        if inv_ids:
            ph = ",".join("?" * len(inv_ids))
            for il in self.db.execute(
                f"SELECT invoice_id, result FROM invoice_lines "
                f"WHERE invoice_id IN ({ph}) "
                f"AND service NOT IN ('Credit Card Fee','Certificate') "
                f"AND result IS NOT NULL AND result!=''",
                inv_ids
            ).fetchall():
                line_results.setdefault(il["invoice_id"], []).append(il["result"].strip().upper())

        # Chart data
        if period == "w":
            def _bucket(r): return (datetime.strptime(r["invoice_date"],"%Y-%m-%d") - start).days
        elif period == "m":
            def _bucket(r): return datetime.strptime(r["invoice_date"],"%Y-%m-%d").day - 1
        else:
            def _bucket(r): return datetime.strptime(r["invoice_date"],"%Y-%m-%d").month - 1

        pass_cnt = [0]*len(labels); fail_cnt = [0]*len(labels)
        total_rev = 0; pass_total = 0
        for r in rows:
            try: idx = _bucket(r)
            except Exception: continue
            if 0 <= idx < len(labels):
                results = line_results.get(r["invoice_id"], [])
                if not results:
                    # Fall back to invoice-level result for mobile-synced or old records
                    hdr = (r["test_result"] or "").upper().strip()
                    if hdr in ("PASS","PASSED"): results = ["PASS"]
                    elif hdr in ("FAIL","FAILED","RETEST"): results = ["FAIL"]
                for result in results:
                    if result in ("PASS","PASSED"):
                        pass_cnt[idx] += 1; pass_total += 1
                    elif result in ("FAIL","FAILED","RETEST"):
                        fail_cnt[idx] += 1
                total_rev += r["amount_cents"]

        chart_data = [(labels[i], pass_cnt[i], fail_cnt[i]) for i in range(len(labels))]
        self._rpt_chart.setData(chart_data)

        total_insp = sum(pass_cnt) + sum(fail_cnt)
        pass_rate  = f"{int(pass_total/total_insp*100)}%" if total_insp else "—"
        rev_str    = f"${total_rev/100:,.2f}"
        unique_days = len(set(r["invoice_date"] for r in rows)) or 1
        avg_str    = f"{total_insp/unique_days:.1f}" if total_insp else "—"

        self._rpt_s_insp.setText(str(total_insp))
        self._rpt_s_pass.setText(pass_rate)
        self._rpt_s_rev.setText(rev_str)
        self._rpt_s_avg.setText(avg_str)
        period_lbl = {"w":"this week","m":"this month","y":"this year"}[period]
        self._rpt_s_insp_sub.setText(f"Total {period_lbl}")
        self._rpt_s_pass_sub.setText(f"{pass_total} pass / {sum(fail_cnt)} fail")
        self._rpt_s_rev_sub.setText(f"Gross revenue")
        self._rpt_s_avg_sub.setText(f"Avg per day")

        # Populate tables (all-period view)
        self._rpt_fill_tables(rows, "All period")

    def _rpt_bar_clicked(self, bar_idx):
        if bar_idx < 0:
            self._rpt_refresh_dashboard(); return
        anchor = getattr(self, '_rpt_anchor', datetime.today()); period = self._rpt_period
        _sel = ("SELECT invoice_id, invoice_date, test_result, amount_cents, "
                "customer_name, first_name, last_name, company_name FROM invoices")
        if period == "w":
            start = anchor - timedelta(days=(anchor.weekday() + 1) % 7)
            d = (start + timedelta(days=bar_idx)).strftime("%Y-%m-%d")
            rows = self.db.execute(_sel + " WHERE is_estimate=0 AND invoice_date=?", (d,)).fetchall()
            lbl = d
        elif period == "m":
            d = anchor.replace(day=bar_idx+1).strftime("%Y-%m-%d")
            rows = self.db.execute(_sel + " WHERE is_estimate=0 AND invoice_date=?", (d,)).fetchall()
            lbl = d
        else:
            m = bar_idx + 1
            start_s = anchor.replace(month=m, day=1).strftime("%Y-%m-%d")
            import calendar; last_day = calendar.monthrange(anchor.year, m)[1]
            end_s = anchor.replace(month=m, day=last_day).strftime("%Y-%m-%d")
            rows = self.db.execute(_sel + " WHERE is_estimate=0 AND invoice_date>=? AND invoice_date<=?",
                (start_s, end_s)).fetchall()
            lbl = datetime(anchor.year, m, 1).strftime("%B %Y")
        self._rpt_fill_tables(rows, lbl)

    def _rpt_fill_tables(self, rows, period_lbl):
        # Top customers
        from collections import Counter
        cust_counts = Counter(); cust_rev = {}
        svc_counts = Counter(); svc_pass = Counter()
        for r in rows:
            name = r["customer_name"] or f"{r['first_name']} {r['last_name']}".strip() or r.get("company_name","?")
            cust_counts[name] += 1
            cust_rev[name] = cust_rev.get(name, 0) + r["amount_cents"]
        t = self._rpt_cust_tbl; t.setRowCount(0)
        self._rpt_cust_tbl_sub.setText(period_lbl)
        for i,(name,cnt) in enumerate(cust_counts.most_common(8)):
            t.insertRow(i)
            for col,val in enumerate([name, str(cnt), f"${cust_rev[name]/100:,.2f}"]):
                t.setItem(i, col, QTableWidgetItem(val))

        # By test type — pull service names from invoice_lines
        t2 = self._rpt_tt_tbl; t2.setRowCount(0)
        self._rpt_tt_tbl_sub.setText(period_lbl)
        inv_ids = [r["invoice_id"] for r in rows]
        if inv_ids:
            svc_rows = self.db.execute(
                f"SELECT il.service, il.result FROM invoice_lines il "
                f"JOIN invoices i ON il.invoice_id=i.invoice_id "
                f"WHERE i.is_estimate=0 AND i.invoice_id IN ({','.join('?'*len(inv_ids))})",
                inv_ids
            ).fetchall()
            svc_cnt = Counter(); svc_pass2 = Counter()
            for sr in svc_rows:
                svc = sr["service"] or "Unknown"
                svc_cnt[svc]+=1
                if (sr["result"] or "").upper() in ("PASS","PASSED"): svc_pass2[svc]+=1
            for i,(svc,cnt) in enumerate(svc_cnt.most_common(6)):
                pr = f"{int(svc_pass2[svc]/cnt*100)}%" if cnt else "N/A"
                t2.insertRow(i)
                for col,val in enumerate([svc, str(cnt), pr]):
                    t2.setItem(i, col, QTableWidgetItem(val))

    def _run_report(self, rpt_type):
        begin_date = self._rpt_begin.date()
        end_date   = self._rpt_end.date()

        # Override date range based on report type
        today = QDate.currentDate()
        if rpt_type == "mtd":
            begin = QDate(today.year(), today.month(), 1).toString("yyyy-MM-dd")
            end   = today.toString("yyyy-MM-dd")
        elif rpt_type == "weekly":
            # Start of current week (Sunday)
            dow   = today.dayOfWeek() % 7  # Qt: 1=Mon…7=Sun → 0=Sun…6=Sat
            begin = today.addDays(-dow).toString("yyyy-MM-dd")
            end   = today.addDays(6 - dow).toString("yyyy-MM-dd")
        elif rpt_type == "daily":
            begin = today.toString("yyyy-MM-dd")
            end   = begin
        else:
            begin = begin_date.toString("yyyy-MM-dd")
            end   = end_date.toString("yyyy-MM-dd")

        all_rows = []
        if rpt_type != "vehicles_due":
            try:
                all_rows = self.db.execute("""
                    SELECT invoice_number,invoice_date,customer_name,first_name,last_name,
                           company_name,amount_cents,payment_method,is_estimate,test_result
                    FROM invoices WHERE invoice_date>=? AND invoice_date<=? AND is_estimate=?
                    ORDER BY invoice_date
                """,(begin,end,1 if rpt_type=="estimates" else 0)).fetchall()
            except Exception as e:
                QMessageBox.warning(self,"Error",str(e)); return

        if rpt_type == "balances":
            title_str = "Outstanding Account Balances"
        elif rpt_type == "vehicles_due":
            title_str = "Vehicles Due for Testing — Next 90 Days"
        else:
            title_str = f"Report - {rpt_type}  ({begin} to {end})"
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

        elif rpt_type == "vehicles_due":
            _today = datetime.now().date()
            # Drop vehicles more than 9 months overdue
            _drop_before = (_today - timedelta(days=274)).strftime("%Y-%m-%d")

            # Filter control row
            _vdue_filters = [
                ("30days",  "30 Days"),
                ("60days",  "60 Days"),
                ("90days",  "90 Days"),
            ]
            _filter_row = QHBoxLayout()
            _filter_row.addWidget(QLabel("Show:"))
            _vdue_cmb = QComboBox()
            for _fk, _fl in _vdue_filters:
                _vdue_cmb.addItem(_fl, _fk)
            # Default to last used filter; remap old keys to '30days'
            _valid_keys = {k for k, _ in _vdue_filters}
            _saved_raw   = getattr(self, '_veh_due_filter', '30days')
            _saved_filter = _saved_raw if _saved_raw in _valid_keys else '30days'
            _vdue_cmb.setCurrentIndex(next((i for i,(k,_) in enumerate(_vdue_filters) if k==_saved_filter), 0))
            _filter_row.addWidget(_vdue_cmb)
            _filter_row.addStretch()
            lay.addLayout(_filter_row)

            def _vdue_query(fkey):
                t = _today; d = _drop_before
                if fkey == '60days':
                    lo = d; hi = (t + timedelta(days=60)).strftime("%Y-%m-%d")
                elif fkey == '90days':
                    lo = d; hi = (t + timedelta(days=90)).strftime("%Y-%m-%d")
                else:  # 30days default
                    lo = d; hi = (t + timedelta(days=30)).strftime("%Y-%m-%d")
                return self.db.execute("""
                    SELECT v.plate, v.vin, v.year, v.make, v.model, v.next_test_due,
                           c.first_name, c.last_name, c.company_name, c.phone
                    FROM vehicles v
                    LEFT JOIN customers c ON v.customer_id = c.customer_id
                    WHERE v.next_test_due != '' AND v.next_test_due IS NOT NULL
                      AND v.next_test_due >= ? AND v.next_test_due <= ?
                    ORDER BY v.next_test_due ASC
                """, (lo, hi)).fetchall()

            def _populate_vdue_tbl(fkey):
                self._veh_due_filter = fkey
                tbl.setRowCount(0)
                veh_rows = _vdue_query(fkey)
                for vr in veh_rows:
                    due_str = (vr["next_test_due"] or "").strip()
                    try:
                        due_date = datetime.strptime(due_str, "%Y-%m-%d").date()
                        days_left = (due_date - _today).days
                        due_display = due_date.strftime("%m/%d/%Y")
                        days_display = f"{days_left}d" if days_left >= 0 else f"{abs(days_left)}d overdue"
                    except ValueError:
                        due_display = due_str; days_left = 999; days_display = ""
                    mk_mod = " ".join(filter(None, [vr["make"], vr["model"]]))
                    cname = vr["company_name"] or f"{vr['first_name'] or ''} {vr['last_name'] or ''}".strip()
                    r_idx = tbl.rowCount(); tbl.insertRow(r_idx)
                    for col, val in enumerate([due_display, days_display, vr["plate"] or "",
                                               str(vr["year"] or ""), mk_mod, cname, vr["phone"] or ""]):
                        item = QTableWidgetItem(val)
                        if days_left < 0:
                            item.setForeground(QColor("#B91C1C"))
                            if col == 1: item.setFont(QFont("", -1, QFont.Weight.Bold))
                        elif days_left <= 30:
                            item.setForeground(QColor("#D97706"))
                        tbl.setItem(r_idx, col, item)

            tbl.setColumnCount(7)
            tbl.setHorizontalHeaderLabels(["Due Date","Days","Plate","Year","Make / Model","Customer","Phone"])
            tbl.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
            tbl.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
            _populate_vdue_tbl(_saved_filter)
            _vdue_cmb.currentIndexChanged.connect(
                lambda _: _populate_vdue_tbl(_vdue_cmb.currentData()))

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
        elif rpt_type == "vehicles_due":
            overdue_ct = sum(1 for r in range(tbl.rowCount())
                             if tbl.item(r,1) and "overdue" in (tbl.item(r,1).text() or ""))
            summary = QLabel(f"Vehicles shown: {tbl.rowCount()}   |   Overdue: {overdue_ct}")
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

    def _stt_tab_style(self, active):
        if active:
            return (f"QPushButton{{background:{CLR_BFAINT};color:{CLR_BLUE};border:none;"
                    f"border-left:3px solid {CLR_BLUE};text-align:left;padding:8px 16px;"
                    f"font-weight:600;font-size:10pt;}}")
        return (f"QPushButton{{background:transparent;color:{CLR_TEXT};border:none;"
                f"border-left:3px solid transparent;text-align:left;padding:8px 16px;font-size:10pt;}}"
                f"QPushButton:hover{{background:{CLR_BFAINT};}}")

    def _stt_show(self, key):
        idx = self._stt_idx.get(key, 0)
        self._stt_stack.setCurrentIndex(idx)
        for k, b in self._stt_btns.items(): b.setStyleSheet(self._stt_tab_style(k == key))

    def _build_settings_screen(self):
        w = QWidget(); self._screens["settings"] = w
        root_lay = QVBoxLayout(w); root_lay.setContentsMargins(0,0,0,0); root_lay.setSpacing(0)
        self._stack.addWidget(w)

        # ── Page header ──────────────────────────────────────────────────
        hdr = QWidget(); hdr.setStyleSheet(f"background:{CLR_CARD};border-bottom:1px solid {CLR_BORDER};")
        hdr_h = QHBoxLayout(hdr); hdr_h.setContentsMargins(20,10,16,10)
        _ttl = QLabel("Settings")
        _ttl.setStyleSheet(f"color:{CLR_TEXT};font-size:14pt;font-weight:700;background:transparent;")
        hdr_h.addWidget(_ttl); hdr_h.addStretch()
        root_lay.addWidget(hdr)

        # ── Body: left sidebar + right stack ─────────────────────────────
        body = QWidget(); body_lay = QHBoxLayout(body); body_lay.setContentsMargins(0,0,0,0); body_lay.setSpacing(0)
        root_lay.addWidget(body, 1)
        sidebar = QWidget(); sidebar.setFixedWidth(172)
        sidebar.setStyleSheet(f"background:{CLR_SURFACE};border-right:1px solid {CLR_BORDER};")
        sb_lay = QVBoxLayout(sidebar); sb_lay.setContentsMargins(0,12,0,12); sb_lay.setSpacing(2)
        body_lay.addWidget(sidebar)
        self._stt_stack = QStackedWidget(); body_lay.addWidget(self._stt_stack, 1)
        _TAB_ITEMS = [
            ("business","Business"), ("logo","Logo"), ("services","Services"),
            ("printer","Printer"),   ("billing","Billing"),
            ("import","Import"),     ("sync","Sync"),
        ]
        self._stt_btns = {}
        for key, label in _TAB_ITEMS:
            b = QPushButton(label); b.setStyleSheet(self._stt_tab_style(False))
            b.clicked.connect(lambda _, k=key: self._stt_show(k))
            self._stt_btns[key] = b; sb_lay.addWidget(b)
        sb_lay.addStretch()

        # ── Business tab ──────────────────────────────────────────────────
        t_biz = QScrollArea(); t_biz.setWidgetResizable(True)
        biz_inner = QWidget(); biz_body = QVBoxLayout(biz_inner)
        biz_body.setContentsMargins(28,24,28,24); biz_body.setSpacing(16)
        card = QWidget()
        card.setStyleSheet(f"background:{CLR_CARD};border:1px solid {CLR_BORDER};border-radius:8px;")
        cl = QVBoxLayout(card); cl.setContentsMargins(20,16,20,20); cl.setSpacing(10)
        cl.addWidget(QLabel("Business information", font=QFont("Segoe UI",11,QFont.Weight.Bold)))
        biz = get_business_settings(self.db); self._biz = {}
        # Parse legacy address_line2 into city/state/zip if separate keys not yet saved
        if not biz.get("city") and biz.get("address_line2"):
            import re as _re
            _m = _re.match(r'^(.+?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$', biz["address_line2"].strip())
            if _m: biz["city"], biz["state"], biz["zip"] = _m.group(1).strip(), _m.group(2), _m.group(3)
        for key, lbl_txt, biz_key in [("name","Business name","name"),
                                       ("address_line1","Address","address_line1")]:
            lw = QLabel(lbl_txt); lw.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;font-weight:600;")
            e = QLineEdit(str(biz.get(biz_key,""))); self._biz[key] = e
            cl.addWidget(lw); cl.addWidget(e)
        # City / State / Zip row with zip auto-fill
        lw_csz = QLabel("City / State / ZIP"); lw_csz.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;font-weight:600;")
        cl.addWidget(lw_csz)
        csz_row = QHBoxLayout(); csz_row.setSpacing(6)
        _e_city  = QLineEdit(str(biz.get("city",""))); _e_city.setPlaceholderText("City"); self._biz["city"] = _e_city
        _e_state = QLineEdit(str(biz.get("state",""))); _e_state.setPlaceholderText("ST"); _e_state.setMaximumWidth(50); self._biz["state"] = _e_state
        _e_biz_zip = QLineEdit(str(biz.get("zip",""))); _e_biz_zip.setPlaceholderText("ZIP"); _e_biz_zip.setMaximumWidth(80); self._biz["zip"] = _e_biz_zip
        def _biz_zip_lookup():
            z = _e_biz_zip.text().strip()
            if len(z) != 5 or not z.isdigit(): return
            if _e_city.text().strip() and _e_state.text().strip(): return
            self._biz_zip_worker = ZipWorker(z)
            self._biz_zip_worker.done.connect(lambda c,s: (_e_city.setText(c.title()) if c else None, _e_state.setText(s) if s else None))
            self._biz_zip_worker.start()
        _e_biz_zip.editingFinished.connect(_biz_zip_lookup)
        csz_row.addWidget(_e_city, 3); csz_row.addWidget(_e_state); csz_row.addWidget(_e_biz_zip)
        cl.addLayout(csz_row)
        row2 = QHBoxLayout(); row2.setSpacing(12)
        for key, lbl_txt, biz_key in [("ard","License number","ard"), ("phone","Phone","phone")]:
            col = QVBoxLayout()
            lw = QLabel(lbl_txt); lw.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;font-weight:600;")
            e = QLineEdit(str(biz.get(biz_key,""))); self._biz[key] = e
            col.addWidget(lw); col.addWidget(e); row2.addLayout(col, 1)
        cl.addLayout(row2)
        self._biz["email"] = QLineEdit(str(biz.get("email",""))); self._biz["email"].setVisible(False)
        _lw_web = QLabel("Website"); _lw_web.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;font-weight:600;")
        self._biz["website"] = QLineEdit(str(biz.get("website","")))
        self._biz["website"].setPlaceholderText("https://")
        cl.addWidget(_lw_web); cl.addWidget(self._biz["website"])
        self._biz["card_fee"] = QLineEdit(str(biz.get("card_fee", "5.00")))
        self._biz_notice = QTextEdit(biz.get("invoice_notice","")); self._biz_notice.setVisible(False)
        self._biz_logo = QLineEdit(biz.get("logo_path","")); self._biz_logo.setVisible(False)
        sv_row = QHBoxLayout()
        _sv = QPushButton("Save")
        _sv.setStyleSheet(f"QPushButton{{background:{CLR_BLUE};color:white;border:none;border-radius:5px;"
                          f"padding:6px 20px;font-weight:600;}}")
        _sv.clicked.connect(self._save_biz); sv_row.addWidget(_sv); sv_row.addStretch()
        cl.addLayout(sv_row); biz_body.addWidget(card); biz_body.addStretch()
        t_biz.setWidget(biz_inner); self._stt_stack.addWidget(t_biz)  # index 0

        # ── Logo tab ──────────────────────────────────────────────────────
        t_logo = QWidget(); tl_lay = QVBoxLayout(t_logo); tl_lay.setContentsMargins(28,24,28,24); tl_lay.setSpacing(14)
        tl_lay.addWidget(QLabel("Logo", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        lo_row = QHBoxLayout(); lo_row.addWidget(self._biz_logo)
        lo_btn = btn("Browse","secondary"); lo_btn.clicked.connect(self._browse_logo); lo_row.addWidget(lo_btn)
        tl_lay.addLayout(lo_row)
        # Preview
        self._logo_preview = QLabel()
        self._logo_preview.setFixedSize(160,160)
        self._logo_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._logo_preview.setStyleSheet(
            f"border:2px dashed {CLR_BORDER};border-radius:8px;background:{CLR_CARD};color:#888;font-size:9pt;")
        self._logo_preview.setText("No logo selected")
        existing_logo = biz.get("logo_path","")
        if existing_logo and os.path.isfile(existing_logo):
            _pix = QPixmap(existing_logo).scaled(
                156,156,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            self._logo_preview.setPixmap(_pix)
        tl_lay.addWidget(self._logo_preview)
        tl_lay.addSpacing(16)
        tl_lay.addWidget(QLabel("Payment QR Code (printed top-right of invoice)", font=QFont("Segoe UI",10,QFont.Weight.Bold)))
        tl_lay.addWidget(QLabel("Upload a QR code image (e.g. Venmo, Zelle, CashApp). Leave blank to hide.",
                                styleSheet=f"color:{CLR_TSUB};font-size:9pt;"))
        self._biz_qr = QLineEdit(biz.get("qr_path",""))
        qr_row = QHBoxLayout(); qr_row.addWidget(self._biz_qr)
        qr_btn = btn("Browse","secondary"); qr_btn.clicked.connect(self._browse_qr); qr_row.addWidget(qr_btn)
        tl_lay.addLayout(qr_row)
        self._qr_preview = QLabel()
        self._qr_preview.setFixedSize(120,120)
        self._qr_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._qr_preview.setStyleSheet(
            f"border:2px dashed {CLR_BORDER};border-radius:8px;background:{CLR_CARD};color:#888;font-size:9pt;")
        self._qr_preview.setText("No QR selected")
        existing_qr = biz.get("qr_path","")
        if existing_qr and os.path.isfile(existing_qr):
            _qpix = QPixmap(existing_qr).scaled(116,116,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            self._qr_preview.setPixmap(_qpix)
        tl_lay.addWidget(self._qr_preview)
        _sv_logo = btn("Save","primary"); _sv_logo.clicked.connect(self._save_biz)
        tl_lay.addWidget(_sv_logo); tl_lay.addStretch()
        self._stt_stack.addWidget(t_logo)  # index 1

        # ── Services tab ──────────────────────────────────────────────────
        t_svc = QWidget(); ts_lay = QVBoxLayout(t_svc); ts_lay.setContentsMargins(28,24,28,24); ts_lay.setSpacing(10)
        ts_lay.addWidget(QLabel("Services & Prices", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        ts_lay.addWidget(QLabel("Add or edit the services shown in the invoice dropdown."))
        self._svc_table = QTableWidget(0, 3)
        self._svc_table.setHorizontalHeaderLabels(["Service Name", "Price ($)", "Cert Fee ($)"])
        self._svc_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._svc_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._svc_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        svcs = get_services(self.db)
        for svc_name, svc_data in svcs.items():
            _r = self._svc_table.rowCount(); self._svc_table.insertRow(_r)
            self._svc_table.setItem(_r, 0, QTableWidgetItem(svc_name))
            self._svc_table.setItem(_r, 1, QTableWidgetItem(str(svc_data.get("price", 0.0))))
            self._svc_table.setItem(_r, 2, QTableWidgetItem(str(svc_data.get("cert_fee", 0.0))))
        ts_lay.addWidget(self._svc_table)
        svc_btn_h = QHBoxLayout()
        add_svc_b = btn("Add Service", "secondary"); add_svc_b.clicked.connect(self._svc_add_row); svc_btn_h.addWidget(add_svc_b)
        del_svc_b = btn("Delete Selected", "danger"); del_svc_b.clicked.connect(self._svc_del_row); svc_btn_h.addWidget(del_svc_b)
        svc_btn_h.addStretch(); ts_lay.addLayout(svc_btn_h)
        save_svc_b = btn("Save Services", "primary"); save_svc_b.clicked.connect(self._save_services); ts_lay.addWidget(save_svc_b)
        ts_lay.addSpacing(16)
        ts_lay.addWidget(QLabel("Card Surcharge", font=QFont("Segoe UI",10,QFont.Weight.Bold)))
        card_fee_form = QFormLayout(); card_fee_form.addRow("Card fee ($):", self._biz["card_fee"])
        ts_lay.addLayout(card_fee_form)
        save_fee_b = btn("Save Card Fee", "primary"); save_fee_b.clicked.connect(self._save_biz); ts_lay.addWidget(save_fee_b)
        ts_lay.addStretch(); self._stt_stack.addWidget(t_svc)  # index 2

        # ── Printer tab ───────────────────────────────────────────────────
        t3 = QWidget(); t3_lay = QVBoxLayout(t3); t3_lay.setContentsMargins(28,24,28,24); t3_lay.setSpacing(10)
        t3_lay.addWidget(QLabel("Printer", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        ps = get_printer_setting(self.db)
        self._pr_pdf = QRadioButton("Save PDF only"); self._pr_printer = QRadioButton("Send to printer")
        (self._pr_printer if ps.get("mode")=="printer" else self._pr_pdf).setChecked(True)
        t3_lay.addWidget(self._pr_pdf); t3_lay.addWidget(self._pr_printer)
        form3 = QFormLayout(); self._pr_name = QComboBox()
        _printers = []
        try:
            from PyQt6.QtPrintSupport import QPrinterInfo
            _printers = [p.printerName() for p in QPrinterInfo.availablePrinters()]
        except Exception: pass
        if not _printers and _WIN32_PRINT:
            try:
                _printers = [p[2] for p in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL|win32print.PRINTER_ENUM_CONNECTIONS)]
            except Exception: pass
        if not _printers:
            try:
                if sys.platform == "win32":
                    _r = subprocess.run(["powershell","-Command",
                                  "Get-Printer | Select-Object -ExpandProperty Name"],
                                 capture_output=True,text=True,timeout=8)
                    _printers = [p.strip() for p in _r.stdout.strip().splitlines() if p.strip()]
                else:
                    _r = subprocess.run(["lpstat","-a"], capture_output=True, text=True, timeout=8)
                    _printers = [l.split()[0] for l in _r.stdout.strip().splitlines() if l.strip()]
            except Exception: pass
        if sys.platform == "darwin" and "Save as PDF" not in _printers:
            _printers.insert(0, "Save as PDF")
        if _printers:
            self._pr_name.addItems(_printers)
            saved_pr = ps.get("printer_name","")
            if saved_pr in _printers: self._pr_name.setCurrentText(saved_pr)
        else:
            self._pr_name.setEditable(True)
        form3.addRow("Printer:", self._pr_name)
        self._pr_copies = QSpinBox(); self._pr_copies.setRange(1,10); self._pr_copies.setValue(int(ps.get("copies",2)))
        form3.addRow("Copies:", self._pr_copies)
        self._pr_auto = QCheckBox("Auto-print after issuing invoice"); self._pr_auto.setChecked(bool(ps.get("auto_print")))
        t3_lay.addLayout(form3); t3_lay.addWidget(self._pr_auto)
        save_pr_b = btn("Save Printer Settings","primary"); save_pr_b.clicked.connect(self._save_printer)
        t3_lay.addWidget(save_pr_b); t3_lay.addStretch(); self._stt_stack.addWidget(t3)  # index 3

        # ── Billing tab ───────────────────────────────────────────────────
        t_bill = QWidget(); tb_lay = QVBoxLayout(t_bill); tb_lay.setContentsMargins(28,24,28,24); tb_lay.setSpacing(12)
        tb_lay.addWidget(QLabel("Billing & Subscription", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        _sub_status_card = QFrame()
        _sub_status_card.setStyleSheet(f"background:{CLR_CARD};border:1px solid {CLR_BORDER};border-radius:8px;")
        _ssc_lay = QVBoxLayout(_sub_status_card); _ssc_lay.setContentsMargins(16,14,16,14); _ssc_lay.setSpacing(6)
        try:
            _ss = api_subscription_status() if requests else {}
        except Exception: _ss = {}
        _plan = (_ss.get("plan") or "—").upper()
        _can  = _ss.get("can_create", True)
        _plan_lbl = QLabel(f"Plan: <b>{_plan}</b>")
        _plan_lbl.setTextFormat(Qt.TextFormat.RichText)
        _status_lbl = QLabel(f"Status: <b>{'Active' if _can else 'Limited'}</b>")
        _status_lbl.setTextFormat(Qt.TextFormat.RichText)
        _status_lbl.setStyleSheet(f"color:{CLR_PASS if _can else CLR_FAIL};")
        _ssc_lay.addWidget(_plan_lbl); _ssc_lay.addWidget(_status_lbl)
        tb_lay.addWidget(_sub_status_card)
        sub_b2 = btn("Subscribe — $39.99/month", "primary"); sub_b2.clicked.connect(self._settings_subscribe)
        tb_lay.addWidget(sub_b2)
        tb_lay.addWidget(QLabel("Already subscribed? Manage or cancel your plan below.",
                                styleSheet="color:#888;font-size:11px;"))
        manage_b2 = btn("Manage Subscription", "secondary"); manage_b2.clicked.connect(self._settings_manage_subscription)
        tb_lay.addWidget(manage_b2)
        tb_lay.addStretch(); self._stt_stack.addWidget(t_bill)  # index 4

        # ── Import tab ────────────────────────────────────────────────────
        t_imp = QWidget(); ti_lay = QVBoxLayout(t_imp); ti_lay.setContentsMargins(28,24,28,24); ti_lay.setSpacing(12)
        ti_lay.addWidget(QLabel("Import", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        ti_lay.addWidget(QLabel("Import customer records from Smog Master (.smc / .csv)"))
        imp_b = btn("Import from Smog Master","primary"); imp_b.clicked.connect(self._smog_master_import)
        ti_lay.addWidget(imp_b); ti_lay.addStretch(); self._stt_stack.addWidget(t_imp)  # index 5

        # ── Sync / Account tab ────────────────────────────────────────────
        t1 = QWidget(); t1_lay = QVBoxLayout(t1); t1_lay.setContentsMargins(28,24,28,24); t1_lay.setSpacing(10)
        t1_lay.addWidget(QLabel("Sync / Account", font=QFont("Segoe UI",12,QFont.Weight.Bold)))
        saved = load_creds()
        _acct_info = QLabel(f"Signed in as: <b>{saved.get('username','—')}</b>")
        _acct_info.setTextFormat(Qt.TextFormat.RichText); t1_lay.addWidget(_acct_info)
        self._s_err = QLabel(""); self._s_err.setStyleSheet("color:red;"); t1_lay.addWidget(self._s_err)
        btn_h = QHBoxLayout()
        so_b = btn("Sign Out","danger"); so_b.clicked.connect(self._settings_signout); btn_h.addWidget(so_b)
        def _do_fp():
            threading.Thread(
                target=SYNC.force_pull_from_zero,
                kwargs={"notify_cb": lambda msg: self._fp_signal.emit(msg)},
                daemon=True
            ).start()
        fp_b = btn("Force Re-pull","secondary"); fp_b.clicked.connect(_do_fp); btn_h.addWidget(fp_b)
        btn_h.addStretch(); t1_lay.addLayout(btn_h)
        t1_lay.addSpacing(16)

        # ── Account Email ─────────────────────────────────────────────
        t1_lay.addWidget(QLabel("Account Email", font=QFont("Segoe UI",10,QFont.Weight.Bold)))
        t1_lay.addWidget(QLabel("Used for password resets via the mobile app.", styleSheet="color:#666;font-size:11px"))
        self._acct_email = QLineEdit(); self._acct_email.setPlaceholderText("e.g. you@example.com")
        ae_form = QFormLayout(); ae_form.setSpacing(8)
        ae_form.addRow("Email:", self._acct_email)
        t1_lay.addLayout(ae_form)
        self._ae_msg = QLabel(""); self._ae_msg.setWordWrap(True); t1_lay.addWidget(self._ae_msg)
        ae_save_b = btn("Save Email","primary"); ae_save_b.clicked.connect(self._update_email_action)
        t1_lay.addWidget(ae_save_b)
        t1_lay.addSpacing(16)

        # ── Change Password ───────────────────────────────────────────
        t1_lay.addWidget(QLabel("Change Password", font=QFont("Segoe UI",10,QFont.Weight.Bold)))
        self._cp_cur = QLineEdit(); self._cp_cur.setPlaceholderText("Current password"); self._cp_cur.setEchoMode(QLineEdit.EchoMode.Password)
        self._cp_new = QLineEdit(); self._cp_new.setPlaceholderText("New password (8+ chars)"); self._cp_new.setEchoMode(QLineEdit.EchoMode.Password)
        self._cp_con = QLineEdit(); self._cp_con.setPlaceholderText("Confirm new password"); self._cp_con.setEchoMode(QLineEdit.EchoMode.Password)
        cp_form = QFormLayout(); cp_form.setSpacing(8)
        cp_form.addRow("Current:", self._cp_cur)
        cp_form.addRow("New:", self._cp_new)
        cp_form.addRow("Confirm:", self._cp_con)
        t1_lay.addLayout(cp_form)
        self._cp_msg = QLabel(""); self._cp_msg.setWordWrap(True); t1_lay.addWidget(self._cp_msg)
        cp_save_b = btn("Save New Password","primary"); cp_save_b.clicked.connect(self._change_password_action)
        t1_lay.addWidget(cp_save_b)
        t1_lay.addSpacing(16)

        t1_lay.addWidget(QLabel("── Danger Zone ──", font=QFont("Segoe UI",10,QFont.Weight.Bold)))
        clr_b = btn("Clear Local Database","danger"); clr_b.clicked.connect(self._clear_local_db)
        t1_lay.addWidget(clr_b); t1_lay.addStretch()
        self._stt_stack.addWidget(t1)  # index 6

        self._stt_idx = {k: i for i, (k, _) in enumerate(_TAB_ITEMS)}

    def _svc_add_row(self):
        _r = self._svc_table.rowCount(); self._svc_table.insertRow(_r)
        self._svc_table.setItem(_r, 0, QTableWidgetItem("New Service"))
        self._svc_table.setItem(_r, 1, QTableWidgetItem("0.00"))
        self._svc_table.setItem(_r, 2, QTableWidgetItem("0.00"))
        self._svc_table.editItem(self._svc_table.item(_r, 0))

    def _svc_del_row(self):
        _r = self._svc_table.currentRow()
        if _r >= 0: self._svc_table.removeRow(_r)

    def _save_services(self):
        svcs = {}
        for _r in range(self._svc_table.rowCount()):
            name = (self._svc_table.item(_r, 0) or QTableWidgetItem("")).text().strip()
            if not name: continue
            try: price = float((self._svc_table.item(_r, 1) or QTableWidgetItem("0")).text())
            except: price = 0.0
            try: cert_fee = float((self._svc_table.item(_r, 2) or QTableWidgetItem("0")).text())
            except: cert_fee = 0.0
            svcs[name] = {"price": price, "cert_fee": cert_fee}
        set_setting(self.db, "services", json.dumps(svcs))
        self._f_svc.clear(); self._f_svc.addItems(list(svcs.keys()))
        QMessageBox.information(self, "Saved", "Services saved successfully.")

    def _reload_svc_table(self):
        svcs = get_services(self.db)
        self._svc_table.setRowCount(0)
        for svc_name, svc_data in svcs.items():
            _r = self._svc_table.rowCount(); self._svc_table.insertRow(_r)
            self._svc_table.setItem(_r, 0, QTableWidgetItem(svc_name))
            self._svc_table.setItem(_r, 1, QTableWidgetItem(str(svc_data.get("price", 0.0))))
            self._svc_table.setItem(_r, 2, QTableWidgetItem(str(svc_data.get("cert_fee", 0.0))))

    def _on_show_settings(self):
        self._set_page_title("Settings")
        self._reload_svc_table()
        self._stt_show("business")

    def _settings_signin(self):
        u = self._s_user.text().strip(); p = self._s_pass.text().strip()
        if not u or not p: self._s_err.setText("Enter username and password."); return
        self._s_err.setText("Signing in..."); QApplication.processEvents()
        try:
            token,company_id,company_name = api_login(u,p)
            save_creds({"username":u,"password":p,"token":token,"company_id":company_id,"company_name":company_name})
            self._s_err.setText(""); QMessageBox.information(self,"Signed In",f"Signed in as {u}")
        except Exception as e: self._s_err.setText(str(e))

    def _update_email_action(self):
        email = self._acct_email.text().strip().lower()
        self._ae_msg.setStyleSheet("color:red;")
        if not email or '@' not in email:
            self._ae_msg.setText("Enter a valid email address."); return
        if not requests:
            self._ae_msg.setText("Network library not available."); return
        try:
            r = requests.post(f"{API_BASE}/v1/auth/update_email",
                              json={"email": email}, headers=_hdrs(), timeout=10)
            if r.status_code == 200:
                self._ae_msg.setStyleSheet("color:green;")
                self._ae_msg.setText("Email updated.")
            else:
                self._ae_msg.setText(f"Server error: {r.status_code}")
        except Exception as e:
            self._ae_msg.setText(f"Network error: {e}")

    def _change_password_action(self):
        cur  = self._cp_cur.text().strip()
        newp = self._cp_new.text().strip()
        con  = self._cp_con.text().strip()
        self._cp_msg.setStyleSheet("color:red;")
        if not cur or not newp or not con:
            self._cp_msg.setText("All fields are required."); return
        if len(newp) < 8:
            self._cp_msg.setText("New password must be at least 8 characters."); return
        if newp != con:
            self._cp_msg.setText("New passwords do not match."); return
        if not requests:
            self._cp_msg.setText("Network library not available."); return
        try:
            hdrs = {**_hdrs(), "x-current-password": cur, "x-new-password": newp}
            r = requests.post(f"{API_BASE}/v1/auth/change_password", headers=hdrs, timeout=10)
            if r.status_code == 200:
                # Update stored credentials so sync keeps working
                creds = load_creds()
                creds["password"] = newp
                save_creds(creds)
                self._cp_cur.clear(); self._cp_new.clear(); self._cp_con.clear()
                self._cp_msg.setStyleSheet("color:green;")
                self._cp_msg.setText("Password changed successfully.")
            elif r.status_code == 401:
                self._cp_msg.setText("Current password is incorrect.")
            else:
                self._cp_msg.setText(f"Server error: {r.status_code}")
        except Exception as e:
            self._cp_msg.setText(f"Network error: {e}")

    def _settings_signout(self):
        if QMessageBox.question(self,"Sign Out","Sign out and return to login?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            self._do_logout(confirmed=True)

    def _settings_subscribe(self):
        if not requests:
            QMessageBox.warning(self, "Error", "Network library (requests) not installed.")
            return
        try:
            r = requests.post(
                f"{API_BASE}/v1/subscription/checkout",
                json={"plan": "monthly"},
                headers=_hdrs(),
                timeout=15,
            )
            r.raise_for_status()
            url = r.json().get("checkout_url", "")
            if url:
                import webbrowser; webbrowser.open(url)
            else:
                QMessageBox.warning(self, "Subscribe", "No checkout URL returned from server.")
        except Exception as e:
            QMessageBox.warning(self, "Subscribe Error", str(e))

    def _settings_manage_subscription(self):
        if not requests:
            QMessageBox.warning(self, "Error", "Network library (requests) not installed.")
            return
        try:
            r = requests.post(f"{API_BASE}/v1/subscription/portal",
                              headers=_hdrs(), timeout=15)
            r.raise_for_status()
            url = r.json().get("portal_url", "")
            if url:
                import webbrowser
                webbrowser.open(url)
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
        path,_ = QFileDialog.getOpenFileName(self,"Select Logo","","Images (*.png *.jpg *.jpeg *.ico *.bmp *.webp)")
        if not path: return
        self._biz_logo.setText(path)
        pix = QPixmap(path).scaled(156,156,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
        if not pix.isNull():
            self._logo_preview.setPixmap(pix)
        else:
            self._logo_preview.setText("Could not load image")

    def _browse_qr(self):
        path,_ = QFileDialog.getOpenFileName(self,"Select QR Code Image","","Images (*.png *.jpg *.jpeg *.bmp *.webp)")
        if not path: return
        self._biz_qr.setText(path)
        pix = QPixmap(path).scaled(116,116,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
        if not pix.isNull():
            self._qr_preview.setPixmap(pix)
        else:
            self._qr_preview.setText("Could not load image")

    def _save_biz(self):
        biz = {k: e.text() if isinstance(e,QLineEdit) else e.toPlainText() for k,e in self._biz.items()}
        biz["invoice_notice"] = self._biz_notice.toPlainText()
        biz["logo_path"] = self._biz_logo.text()
        biz["qr_path"]   = self._biz_qr.text() if hasattr(self, "_biz_qr") else ""
        try: biz["card_fee"] = float(biz.get("card_fee",5.0))
        except: biz["card_fee"] = 5.0
        # Compose address_line2 for PDF/legacy compatibility
        csz_parts = [biz.get("city","").strip(), biz.get("state","").strip(), biz.get("zip","").strip()]
        biz["address_line2"] = " ".join(p for p in csz_parts if p)
        set_setting(self.db,"business",json.dumps(biz))
        # Sync business info to mobile devices
        enqueue(self.db,"company_settings","upsert",{
            "co_name":              biz.get("name",""),
            "co_addr":              biz.get("address_line1",""),
            "co_city":              biz.get("city",""),
            "co_state":             biz.get("state",""),
            "co_zip":               biz.get("zip",""),
            "co_phone":             biz.get("phone",""),
            "co_email":             biz.get("email",""),
            "co_ard":               biz.get("ard",""),
            "invoice_notice":       biz.get("invoice_notice",""),
            "card_surcharge_value": str(biz.get("card_fee","")),
            "card_surcharge_type":  biz.get("card_surcharge_type","percent"),
        })
        self._refresh_sidebar_name()
        # Refresh logo preview if on logo tab
        logo_path = biz.get("logo_path","")
        if hasattr(self,"_logo_preview"):
            if logo_path and os.path.isfile(logo_path):
                pix = QPixmap(logo_path).scaled(156,156,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
                if not pix.isNull():
                    self._logo_preview.setPixmap(pix)
            else:
                self._logo_preview.setPixmap(QPixmap())
                self._logo_preview.setText("No logo selected")
        QMessageBox.information(self,"Saved","Business info saved.")

    def _save_printer(self):
        ps = {"mode":"printer" if self._pr_printer.isChecked() else "pdf",
              "printer_name":self._pr_name.currentText(),"copies":self._pr_copies.value(),
              "auto_print":self._pr_auto.isChecked()}
        set_setting(self.db,"printer_setting",json.dumps(ps))
        QMessageBox.information(self,"Saved","Printer settings saved.")

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    #  SCREEN: CUSTOMERS
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    _AVATAR_COLORS = ["#4E9EDF","#2A9D8F","#D4776E","#9B72CF","#E9944A","#5E8F6E","#7B68EE","#3A7CA5"]

    def _avatar_color(self, name):
        return self._AVATAR_COLORS[abs(hash(name or "?")) % len(self._AVATAR_COLORS)]

    def _make_avatar_widget(self, initials, name, email, bg):
        w = QWidget(); w.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        h = QHBoxLayout(w); h.setContentsMargins(10,4,8,4); h.setSpacing(10)
        av = QLabel(initials); av.setFixedSize(34,34); av.setAlignment(Qt.AlignmentFlag.AlignCenter)
        av.setStyleSheet(f"background:{bg};border-radius:17px;color:white;font-weight:700;font-size:9pt;")
        h.addWidget(av)
        txt = QWidget(); txt.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        tl = QVBoxLayout(txt); tl.setContentsMargins(0,0,0,0); tl.setSpacing(1)
        nl = QLabel(name); nl.setStyleSheet(f"font-weight:600;font-size:10pt;color:{CLR_TEXT};background:transparent;")
        el = QLabel(email); el.setStyleSheet(f"font-size:8pt;color:{CLR_TSUB};background:transparent;")
        tl.addWidget(nl); tl.addWidget(el); h.addWidget(txt, 1)
        return w

    def _build_customers_screen(self):
        w = QWidget(); self._screens["customers"] = w
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        self._stack.addWidget(w)

        # ── Page header ──────────────────────────────────────────────────
        hdr = QWidget(); hdr.setStyleSheet(f"background:{CLR_CARD};border-bottom:1px solid {CLR_BORDER};")
        hdr_h = QHBoxLayout(hdr); hdr_h.setContentsMargins(20,10,16,10); hdr_h.setSpacing(12)
        _ttl = QLabel("Customers")
        _ttl.setStyleSheet(f"color:{CLR_TEXT};font-size:14pt;font-weight:700;background:transparent;")
        hdr_h.addWidget(_ttl); hdr_h.addSpacing(8)
        self._cust_search = QLineEdit()
        self._cust_search.setPlaceholderText("Search by name, phone...")
        self._cust_search.setStyleSheet(
            f"QLineEdit{{background:{CLR_SURFACE};border:1px solid {CLR_BORDER};border-radius:6px;"
            f"padding:4px 10px;color:{CLR_TEXT};}}")
        self._cust_search.setMinimumWidth(200); self._cust_search.setMaximumWidth(310)
        self._cust_search.textChanged.connect(self._refresh_customers)
        hdr_h.addWidget(self._cust_search); hdr_h.addStretch()
        self._cust_count_lbl = QLabel("")
        self._cust_count_lbl.setStyleSheet(f"color:{CLR_TSUB};font-size:9pt;background:transparent;")
        hdr_h.addWidget(self._cust_count_lbl)
        add_b = QPushButton("+ Add customer")
        add_b.setStyleSheet(
            f"QPushButton{{background:{CLR_BLUE};border:none;border-radius:5px;"
            f"color:white;padding:4px 14px;font-weight:600;}}"
            f"QPushButton:hover{{background:{CLR_NAVY};}}")
        add_b.clicked.connect(self._cust_new)
        hdr_h.addWidget(add_b)
        lay.addWidget(hdr)

        # ── Filter pills ─────────────────────────────────────────────────
        pills_bar = QWidget()
        pills_bar.setStyleSheet(f"background:{CLR_SURFACE};border-bottom:1px solid {CLR_BORDER};")
        pills_h = QHBoxLayout(pills_bar); pills_h.setContentsMargins(20,6,16,6); pills_h.setSpacing(6)
        self._cust_pill = "all"
        self._cust_pill_btns = {}
        for key, label in [("all","All customers"), ("active","Active (30 days)"), ("est","Has estimates")]:
            b = QPushButton(label)
            b.setStyleSheet(self._dl_pill_style(key == "all"))
            self._cust_pill_btns[key] = b
            b.clicked.connect(lambda _, k=key: self._cust_set_pill(k))
            pills_h.addWidget(b)
        pills_h.addStretch()
        lay.addWidget(pills_bar)

        # ── Table ─────────────────────────────────────────────────────────
        cols = ["Customer","Phone","Inspections","Last inspection","Last result","Discount"]
        self._cust_table = QTableWidget(0, len(cols)); self._cust_table.setHorizontalHeaderLabels(cols)
        self._cust_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._cust_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._cust_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._cust_table.verticalHeader().setVisible(False)
        self._cust_table.setShowGrid(False)
        self._cust_table.verticalHeader().setDefaultSectionSize(54)
        self._register_table("customers", self._cust_table)
        self._cust_table.doubleClicked.connect(self._cust_view)
        self._cust_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._cust_table.customContextMenuRequested.connect(self._cust_context_menu)
        for i, w2 in enumerate([0, 110, 90, 120, 100, 90]):
            if w2: self._cust_table.setColumnWidth(i, w2)
        lay.addWidget(self._cust_table)
        self._cust_ids = []

    def _cust_set_pill(self, key):
        self._cust_pill = key
        for k, b in self._cust_pill_btns.items(): b.setStyleSheet(self._dl_pill_style(k == key))
        self._refresh_customers()

    def _on_show_customers(self):
        self._set_page_title("Customers")
        self._refresh_customers()

    def _refresh_customers(self):
        q    = self._cust_search.text().strip().lower()
        pill = getattr(self, '_cust_pill', 'all')
        cutoff30 = (datetime.today() - timedelta(days=30)).strftime("%Y-%m-%d")

        # Three fast flat queries — no correlated subqueries
        all_rows = self.db.execute(
            "SELECT * FROM customers ORDER BY company_name, last_name, first_name"
        ).fetchall()

        # Aggregated stats per customer (single GROUP BY scan)
        stats_map = {}
        for r in self.db.execute(
            "SELECT customer_id,"
            " SUM(CASE WHEN is_estimate=0 THEN 1 ELSE 0 END) as inv_count,"
            " MAX(CASE WHEN is_estimate=0 THEN invoice_date ELSE NULL END) as last_inv,"
            " SUM(CASE WHEN is_estimate=1 THEN 1 ELSE 0 END) as est_count"
            " FROM invoices GROUP BY customer_id"
        ).fetchall():
            stats_map[r["customer_id"]] = r

        # Last test result per customer using MAX(invoice_date) join
        last_res_map = {}
        for r in self.db.execute(
            "SELECT i.customer_id, i.test_result FROM invoices i"
            " INNER JOIN ("
            "  SELECT customer_id, MAX(invoice_date) as md"
            "  FROM invoices WHERE is_estimate=0 GROUP BY customer_id"
            " ) m ON i.customer_id=m.customer_id AND i.invoice_date=m.md"
            " WHERE i.is_estimate=0 GROUP BY i.customer_id"
        ).fetchall():
            last_res_map[r["customer_id"]] = (r["test_result"] or "").upper()
        self._cust_table.setUpdatesEnabled(False)
        self._cust_table.setRowCount(0); self._cust_ids = []; shown = 0
        _zsz = self._zoom.get(self._current_screen, 10)
        for row in all_rows:
            cid  = row["customer_id"]
            st   = stats_map.get(cid)
            inv_count  = st["inv_count"]  if st else 0
            last_inv   = st["last_inv"]   if st else ""
            est_count  = st["est_count"]  if st else 0
            last_res   = last_res_map.get(cid, "")
            display_name = row["company_name"] or f"{row['first_name']} {row['last_name']}".strip()
            email = row["email"] or ""; phone = row["phone"] or ""
            if q and q not in " ".join([display_name, phone, email,
                                        row["first_name"] or "", row["last_name"] or ""]).lower():
                continue
            if pill == "active" and (not last_inv or last_inv < cutoff30):
                continue
            if pill == "est" and not est_count:
                continue
            r = self._cust_table.rowCount(); self._cust_table.insertRow(r)
            self._cust_ids.append(cid)
            words = display_name.split()
            initials = (words[0][0] + (words[-1][0] if len(words)>1 else "")).upper() if words else "?"
            bg = self._avatar_color(display_name)
            self._cust_table.setCellWidget(r, 0, self._make_avatar_widget(initials, display_name, email, bg))
            self._cust_table.setItem(r, 0, QTableWidgetItem(""))  # blank — widget renders on top
            self._cust_table.setItem(r, 1, QTableWidgetItem(phone))
            cnt_item = QTableWidgetItem(str(inv_count or 0))
            cnt_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
            self._cust_table.setItem(r, 2, cnt_item)
            try: last_d = datetime.strptime(last_inv, "%Y-%m-%d").strftime("%m/%d/%y") if last_inv else ""
            except Exception: last_d = last_inv or ""
            dt_item = QTableWidgetItem(last_d)
            dt_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
            self._cust_table.setItem(r, 3, dt_item)
            res_item = QTableWidgetItem(last_res)
            res_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
            if last_res == "PASS":
                res_item.setForeground(QColor(CLR_PASS)); res_item.setBackground(QColor(CLR_PASSBG))
                res_item.setFont(QFont("Segoe UI", _zsz, QFont.Weight.Bold))
            elif last_res in ("FAIL","RETEST"):
                res_item.setForeground(QColor(CLR_FAIL)); res_item.setBackground(QColor(CLR_FAILBG))
                res_item.setFont(QFont("Segoe UI", _zsz, QFont.Weight.Bold))
            self._cust_table.setItem(r, 4, res_item)
            disc = row["discount_percent"] or 0.0
            try: disc_type = (row["discount_type"] or "PERCENT").upper()
            except Exception: disc_type = "PERCENT"
            disc_str = (f"${disc:.2f} off" if disc_type == "FLAT" else f"{disc:.0f}%") if disc else "—"
            disc_item = QTableWidgetItem(disc_str)
            disc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
            if disc: disc_item.setForeground(QColor(CLR_PASS))
            self._cust_table.setItem(r, 5, disc_item)
            shown += 1
        self._cust_table.setUpdatesEnabled(True)
        self._cust_count_lbl.setText(f"{shown} customers")

    def _cust_selected_id(self):
        r = self._cust_table.currentRow()
        return self._cust_ids[r] if 0<=r<len(self._cust_ids) else None

    def _cust_view(self):
        cid = self._cust_selected_id()
        if not cid: return
        cust = self.db.execute("SELECT * FROM customers WHERE customer_id=?",(cid,)).fetchone()
        if not cust: return
        invs = self.db.execute(
            "SELECT invoice_id,invoice_number,invoice_date,amount_cents,payment_method,is_estimate "
            "FROM invoices WHERE customer_id=? ORDER BY invoice_date DESC",(cid,)).fetchall()
        vehs = self.db.execute("""
            SELECT plate, year, make, model, vin,
                   MAX(next_test_due) as next_test_due,
                   test_interval_days
            FROM vehicles
            WHERE customer_id=?
            GROUP BY CASE WHEN vin!='' THEN vin ELSE plate END
            ORDER BY year DESC
        """, (cid,)).fetchall()

        name = f"{cust['first_name']} {cust['last_name']}".strip() or cust['company_name'] or "-"
        dlg = QDialog(self); dlg.setWindowTitle(f"Customer — {name}"); dlg.resize(820, 620)
        lay = QVBoxLayout(dlg); lay.setSpacing(8)

        # Header info
        parts = []
        if cust['phone']: parts.append(f"Phone: {cust['phone']}")
        if cust['email']: parts.append(f"Email: {cust['email']}")
        addr = " ".join(filter(None,[cust['address'],cust['city'],cust['state'],cust['zip']]))
        if addr: parts.append(addr)
        info = QLabel(f"<b>{name}</b>" + (f" — {cust['company_name']}" if cust['company_name'] and cust['company_name'] != name else "") +
                      ("<br>" + "  ·  ".join(parts) if parts else ""))
        info.setTextFormat(Qt.TextFormat.RichText); lay.addWidget(info)

        # — Invoices section —
        inv_lbl = QLabel("Invoices"); inv_lbl.setStyleSheet("font-weight:700;font-size:10pt;")
        lay.addWidget(inv_lbl)
        tbl = QTableWidget(0,5); tbl.setHorizontalHeaderLabels(["#","Date","Type","Amount","Payment"])
        tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        tbl.verticalHeader().setVisible(False)
        total = 0.0; inv_ids = []
        for row in invs:
            r = tbl.rowCount(); tbl.insertRow(r); amt = row["amount_cents"] / 100; total += amt
            inv_ids.append(row["invoice_id"])
            for col, val in enumerate([str(row["invoice_number"] or "-"), row["invoice_date"],
                                        "EST" if row["is_estimate"] else "INV",
                                        f"${amt:,.2f}", row["payment_method"] or ""]):
                tbl.setItem(r, col, QTableWidgetItem(val))
        def _open_invoice(row, _col):
            if row < len(inv_ids):
                self._open_pdf_for_invoice(inv_ids[row])
        tbl.cellDoubleClicked.connect(_open_invoice)
        lay.addWidget(tbl, 2)
        total_lbl = QLabel(f"Total invoiced: ${total:,.2f}")
        total_lbl.setStyleSheet("font-weight:600;"); lay.addWidget(total_lbl)

        # — Vehicles section —
        veh_lbl = QLabel("Vehicles"); veh_lbl.setStyleSheet("font-weight:700;font-size:10pt;margin-top:4px;")
        lay.addWidget(veh_lbl)
        vtbl = QTableWidget(0, 5); vtbl.setHorizontalHeaderLabels(["Plate","VIN","Year","Make / Model","Due By"])
        vtbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        vtbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        vtbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        vtbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        vtbl.verticalHeader().setVisible(False)
        today = datetime.now().date()
        for v in vehs:
            r = vtbl.rowCount(); vtbl.insertRow(r)
            mk_mod = " ".join(filter(None,[v["make"], v["model"]]))
            due_str = (v["next_test_due"] or "").strip()
            due_display = "—"; due_color = None
            if due_str:
                try:
                    due_date = datetime.strptime(due_str, "%Y-%m-%d").date()
                    due_display = due_date.strftime("%m/%d/%Y")
                    days_left = (due_date - today).days
                    if days_left < 0: due_color = QColor("#B91C1C")
                    elif days_left <= 30: due_color = QColor("#D97706")
                    else: due_color = QColor("#1E7E34")
                except ValueError: pass
            for col, val in enumerate([v["plate"] or "—", v["vin"] or "", str(v["year"] or ""), mk_mod, due_display]):
                item = QTableWidgetItem(val)
                if col == 4 and due_color: item.setForeground(due_color)
                vtbl.setItem(r, col, item)
        if not vehs:
            vtbl.insertRow(0)
            item = QTableWidgetItem("No vehicles on file")
            item.setForeground(QColor("#888")); vtbl.setItem(0, 0, item)
            vtbl.setSpan(0, 0, 1, 5)
        veh_list = list(vehs)
        def _edit_vehicle(row, _col):
            if row >= len(veh_list): return
            v = veh_list[row]
            if v["vin"]:
                vrow = self.db.execute("SELECT * FROM vehicles WHERE vin=? AND customer_id=?", (v["vin"], cid)).fetchone()
            else:
                vrow = self.db.execute("SELECT * FROM vehicles WHERE plate=? AND customer_id=? AND (vin='' OR vin IS NULL)", (v["plate"], cid)).fetchone()
            if not vrow: return
            ed = QDialog(dlg); ed.setWindowTitle("Edit Vehicle"); ed.resize(400,320)
            el = QVBoxLayout(ed); ef = QFormLayout()
            ep = QLineEdit(vrow["plate"] or ""); ey = QLineEdit(vrow["year"] or "")
            emk = QLineEdit(vrow["make"] or ""); emd = QLineEdit(vrow["model"] or "")
            evin = QLineEdit(vrow["vin"] or "")
            for w2 in (ep,ey,emk,emd,evin): w2.setMinimumWidth(200)
            ef.addRow("Plate:", ep); ef.addRow("Year:", ey)
            ef.addRow("Make:", emk); ef.addRow("Model:", emd); ef.addRow("VIN:", evin)
            int_cb = QComboBox()
            for lbl_txt, _ in _INTERVAL_OPTS: int_cb.addItem(lbl_txt)
            cur_int = vrow["test_interval_days"] if "test_interval_days" in vrow.keys() else None
            for i,(_, days) in enumerate(_INTERVAL_OPTS):
                if days == cur_int: int_cb.setCurrentIndex(i); break
            ef.addRow("Test Frequency:", int_cb)
            due_edit = QDateEdit()
            due_edit.setCalendarPopup(True)
            due_edit.setDisplayFormat("MM/dd/yyyy")
            cur_due_str = (vrow["next_test_due"] if "next_test_due" in vrow.keys() else None) or ""
            if cur_due_str:
                try:
                    _d = datetime.strptime(cur_due_str, "%Y-%m-%d").date()
                    due_edit.setDate(QDate(_d.year, _d.month, _d.day))
                except ValueError:
                    due_edit.setDate(QDate.currentDate())
            else:
                due_edit.setDate(QDate.currentDate())
            ef.addRow("Next Due Date:", due_edit)

            # Auto-update due date when interval changes
            def _auto_due(idx):
                days = _INTERVAL_OPTS[idx][1]
                if days is None: return
                last_inv = self.db.execute(
                    "SELECT invoice_date FROM invoices WHERE customer_id=? AND (plate=? OR vin=?) "
                    "AND is_estimate=0 ORDER BY invoice_date DESC LIMIT 1",
                    (cid, vrow["plate"] or "", vrow["vin"] or "")).fetchone()
                base = datetime.strptime(last_inv["invoice_date"], "%Y-%m-%d") if last_inv else datetime.now()
                nd = (base + timedelta(days=days)).date()
                due_edit.setDate(QDate(nd.year, nd.month, nd.day))
            int_cb.currentIndexChanged.connect(_auto_due)

            el.addLayout(ef)
            ebb = QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
            ebb.accepted.connect(ed.accept); ebb.rejected.connect(ed.reject); el.addWidget(ebb)
            if ed.exec() != QDialog.DialogCode.Accepted: return
            new_interval = _INTERVAL_OPTS[int_cb.currentIndex()][1]
            if new_interval is None:
                new_due = ""
            else:
                qd = due_edit.date()
                new_due = f"{qd.year():04d}-{qd.month():02d}-{qd.day():02d}"
            self.db.execute(
                "UPDATE vehicles SET plate=?,year=?,make=?,model=?,vin=?,test_interval_days=?,next_test_due=?,updated_at=? WHERE vehicle_id=?",
                (ep.text().strip().upper(), ey.text().strip(), emk.text().strip().upper(),
                 emd.text().strip().upper(), evin.text().strip().upper(),
                 new_interval, new_due, now_iso(), vrow["vehicle_id"]))
            self.db.commit()
            enqueue(self.db,"vehicle","upsert",{"vehicle_id":vrow["vehicle_id"],
                "customer_id":cid,"plate":ep.text().strip().upper(),"vin":evin.text().strip().upper(),
                "make":emk.text().strip().upper(),"model":emd.text().strip().upper(),
                "year":ey.text().strip(),"test_interval_days":new_interval,
                "next_test_due":new_due,"next_due":new_due,"service_interval_days":new_interval})
            dlg.accept(); self._cust_view()
        def _delete_vehicle():
            sel = vtbl.selectionModel().selectedRows()
            if not sel:
                QMessageBox.information(dlg, "No Selection",
                    "Click a vehicle row to select it, then click Delete Vehicle.")
                return
            row = sel[0].row()
            if row >= len(veh_list): return
            v = veh_list[row]
            if v["vin"]:
                vrow = self.db.execute("SELECT * FROM vehicles WHERE vin=? AND customer_id=?", (v["vin"], cid)).fetchone()
            else:
                vrow = self.db.execute("SELECT * FROM vehicles WHERE plate=? AND customer_id=? AND (vin='' OR vin IS NULL)", (v["plate"], cid)).fetchone()
            if not vrow: return
            plate_disp = vrow["plate"] or "(no plate)"
            vin_disp   = vrow["vin"]   or "(no VIN)"
            if QMessageBox.question(dlg, "Delete Vehicle",
                f"Delete vehicle  {plate_disp}  /  {vin_disp}?\n\nThis cannot be undone.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            ) != QMessageBox.StandardButton.Yes: return
            self.db.execute("DELETE FROM vehicles WHERE vehicle_id=?", (vrow["vehicle_id"],))
            self.db.commit()
            enqueue(self.db, "vehicle", "delete", {"vehicle_id": vrow["vehicle_id"]})
            dlg.accept(); self._cust_view()

        vtbl.cellDoubleClicked.connect(_edit_vehicle)
        lay.addWidget(vtbl, 3)

        del_veh_btn = QPushButton("Delete Selected Vehicle")
        del_veh_btn.setStyleSheet(
            f"QPushButton{{color:{CLR_FAIL};border:1px solid {CLR_BORDER};"
            f"border-radius:4px;padding:3px 10px;background:transparent;}}"
            f"QPushButton:hover{{background:{CLR_FAILBG};border-color:{CLR_FAIL};}}")
        del_veh_btn.clicked.connect(_delete_vehicle)
        veh_btn_row = QHBoxLayout()
        veh_btn_row.addWidget(del_veh_btn); veh_btn_row.addStretch()
        lay.addLayout(veh_btn_row)

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


    def _cust_context_menu(self, pos):
        cid = self._cust_selected_id()
        if not cid: return
        menu = QMenu(self)
        menu.addAction("View", self._cust_view)
        menu.addAction("Edit", self._cust_edit)
        menu.addSeparator()
        menu.addAction("Delete...", self._cust_delete)
        menu.exec(self._cust_table.viewport().mapToGlobal(pos))

    def _cust_delete(self):
        cid = self._cust_selected_id()
        if not cid: return
        cust = self.db.execute("SELECT * FROM customers WHERE customer_id=?", (cid,)).fetchone()
        if not cust: return
        name = (cust["first_name"]+" "+cust["last_name"]).strip() or cust["company_name"] or cid
        inv_count = self.db.execute("SELECT COUNT(*) FROM invoices WHERE customer_id=?", (cid,)).fetchone()[0]
        extra = (f"\n\nWarning: {inv_count} invoice(s) will also be deleted." if inv_count else "")
        if QMessageBox.question(self, "Delete Customer",
            f"Permanently delete '{name}'?{extra}\n\nThis cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) != QMessageBox.StandardButton.Yes: return
        if requests:
            try:
                r = requests.delete(f"{API_BASE}/v1/customers/{cid}", headers=_hdrs(), timeout=15); r.raise_for_status()
            except Exception as e:
                QMessageBox.warning(self, "Server", f"Server delete failed: {e}\nDeleted locally only.")
        cust_invoices = self.db.execute("SELECT invoice_id FROM invoices WHERE customer_id=?", (cid,)).fetchall()
        cust_vehicles = self.db.execute("SELECT vehicle_id FROM vehicles WHERE customer_id=?", (cid,)).fetchall()
        self.db.execute("DELETE FROM invoice_lines WHERE invoice_id IN (SELECT invoice_id FROM invoices WHERE customer_id=?)", (cid,))
        self.db.execute("DELETE FROM invoices WHERE customer_id=?", (cid,))
        self.db.execute("DELETE FROM vehicles WHERE customer_id=?", (cid,))
        self.db.execute("DELETE FROM customers WHERE customer_id=?", (cid,))
        self.db.commit()
        for inv_row in cust_invoices:
            enqueue(self.db, "invoice", "delete", {"invoice_id": inv_row["invoice_id"]})
        for veh_row in cust_vehicles:
            enqueue(self.db, "vehicle", "delete", {"vehicle_id": veh_row["vehicle_id"]})
        enqueue(self.db, "customer", "delete", {"customer_id": cid})
        self._refresh_customers()

    def _smog_master_import(self):
        dlg = SmogMasterImportDialog(self.db, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
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

    def _on_show_admin(self):
        self._set_page_title("Admin Dashboard")
        self._admin_refresh()

    def _admin_stat_card(self, parent_layout, label, value, color):
        card = QFrame(); card.setFrameShape(QFrame.Shape.StyledPanel)
        card.setStyleSheet(f"background:{CLR_SURFACE}; border:1px solid {CLR_BORDER}; border-radius:8px;")
        cl = QVBoxLayout(card); cl.setContentsMargins(12,12,12,12)
        val_lbl = QLabel(value)
        val_lbl.setStyleSheet(f"color:{color}; font-size:22pt; font-weight:bold;")
        val_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_lbl = QLabel(label); lbl_lbl.setStyleSheet("color:#374151; font-size:9pt;")
        lbl_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(val_lbl); cl.addWidget(lbl_lbl)
        parent_layout.addWidget(card)
        return val_lbl

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
        ex_count      = sum(1 for e in exempts if e["username"] != "bluesky_master")
        total_inv     = sum(co.get("invoice_count", 0) for co in companies)
        suspended_cnt = sum(1 for co in companies if co.get("is_suspended"))
        self._adm_stat_co.setText(str(len(companies)))
        self._adm_stat_inv.setText(str(total_inv))
        self._adm_stat_ex.setText(str(ex_count))
        self._adm_stat_sus.setText(str(suspended_cnt))

        # Exemptions table
        self._adm_ex_tbl.setRowCount(0)
        for e in exempts:
            r = self._adm_ex_tbl.rowCount(); self._adm_ex_tbl.insertRow(r)
            self._adm_ex_tbl.setItem(r,0,QTableWidgetItem(e["username"]))
            self._adm_ex_tbl.setItem(r,1,QTableWidgetItem((e.get("added_at","") or "")[:10]))

        # Companies table
        self._adm_co_tbl.setRowCount(0)
        for co in companies:
            sus = co.get("is_suspended", False)
            r   = self._adm_co_tbl.rowCount(); self._adm_co_tbl.insertRow(r)
            self._adm_co_tbl.setItem(r,0,QTableWidgetItem(co.get("company_name","")))
            self._adm_co_tbl.setItem(r,1,QTableWidgetItem("@"+co.get("username","")))
            self._adm_co_tbl.setItem(r,2,QTableWidgetItem(str(co.get("invoice_count",0))))
            si = QTableWidgetItem("Suspended" if sus else "Active")
            si.setForeground(QColor(RED if sus else GREEN))
            self._adm_co_tbl.setItem(r,3,si)
            ls = (co.get("last_activity","") or "")[:10] or "Never"
            self._adm_co_tbl.setItem(r,4,QTableWidgetItem(ls))
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


_active_update_workers = []  # keeps update worker objects alive until their thread exits


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
            if ver_tuple(tag) > ver_tuple(APP_VERSION):
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
        dl_btn = msg.addButton("Download Now", QMessageBox.ButtonRole.AcceptRole)
        msg.addButton("Later", QMessageBox.ButtonRole.RejectRole)
        msg.exec()
        if msg.clickedButton() == dl_btn:
            import webbrowser
            webbrowser.open(_DOWNLOAD_URL)

    # Qt automatically queues cross-thread signal delivery to the main thread
    worker.update_found.connect(_show, Qt.ConnectionType.QueuedConnection)
    _active_update_workers.append(worker)

    def _run_and_release():
        try:
            worker.check()
        finally:
            try: _active_update_workers.remove(worker)
            except ValueError: pass

    threading.Thread(target=_run_and_release, daemon=True).start()


def main():
    _hide_console()   # hide & remove from taskbar before anything else
    import traceback
    try:
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        app.setStyleSheet(APP_STYLE)
        _af = QFont()
        _af.setFamilies(["Segoe UI", "Segoe UI Emoji", "Segoe UI Symbol", "Arial"])
        _af.setPointSizeF(10.5)
        _af.setWeight(QFont.Weight.Medium)
        app.setFont(_af)
        init_db()
        migrate_db()

        while True:
            app._show_login = False
            creds = load_creds()
            # Always require explicit sign-in if no valid token
            if not creds.get("token"):
                dlg = LoginDialog()
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    break
            window = App()
            window.showMaximized()
            _hide_console()
            # Check for updates 3 seconds after launch so UI is fully loaded
            QTimer.singleShot(3000, lambda: _check_for_update(window))
            app.exec()
            # If logout was triggered, loop back to show login dialog
            if not app._show_login:
                break

        sys.exit(0)
    except Exception as e:
        try:
            QMessageBox.critical(None, "Startup Error",
                f"Failed to start:\n{e}\n\n{traceback.format_exc()}")
        except Exception:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()
