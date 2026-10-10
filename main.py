# -*- coding: utf-8 -*-
"""
Quest Sell Bot - aiogram 3.x, SQLite, single-file project.

Every reply-keyboard button and every inline-keyboard button is colored using the
Telegram Bot API 9.4 `style` field ("primary" = blue, "success" = green, "danger" = red).
"""

import asyncio
import html
import io
import itertools
import logging
import os
import re
import sqlite3
import time
from contextlib import contextmanager
from datetime import datetime, time as dtime, timezone
from decimal import Decimal
from typing import Optional
from zoneinfo import ZoneInfo

from aiogram import BaseMiddleware, Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import (
    TelegramAPIError,
    TelegramBadRequest,
    TelegramForbiddenError,
    TelegramRetryAfter,
)
from aiogram.filters import Command, CommandObject, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    ErrorEvent,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    User,
)
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
PRIMARY_ADMIN_ID = 8002809016  # hardcoded primary admin (only one allowed to use /on and /off)


def _second_admin() -> Optional[int]:
    raw = os.getenv("ADMIN_ID", "").strip()
    return int(raw) if raw.isdigit() else None


SECOND_ADMIN_ID = _second_admin()
ADMIN_IDS = {PRIMARY_ADMIN_ID} | ({SECOND_ADMIN_ID} if SECOND_ADMIN_ID else set())


def resolve_db_path() -> str:
    explicit = os.getenv("DB_PATH", "").strip()
    if explicit:
        return explicit
    volume = os.getenv("RAILWAY_VOLUME_MOUNT_PATH", "").strip()
    if volume:
        return os.path.join(volume, "quest_sell.db")
    return "quest_sell.db"


DB_PATH = resolve_db_path()

SUPPORT_URL = "https://t.me/suppor9287"
BKASH_NUMBER = "01705930972"
NAGAD_NUMBER = "01995839364"
BINANCE_UID = "1237510670"

MIN_DEPOSIT = 2000  # 20 BDT in poisha
MAX_AMOUNT = 100_000_000  # 1,000,000 BDT in poisha
PENDING_TTL = 600  # pending order lifetime (seconds)
AMOUNT_STATE_TTL = 1800  # deposit amount step lifetime
REFERENCE_STATE_TTL = 21600  # deposit TrxID step lifetime
MAX_UPLOAD_BYTES = 5 * 1024 * 1024
MAX_XLSX_ROWS = 50000

DHAKA = ZoneInfo("Asia/Dhaka")
NAME_PREFIX = " "  # decorative prefix from the spec's profile template; set to "" to remove

OFF_TEXT = "🔴 Bot turned off."

# bKash deposit switch: True = চালু, False = বন্ধ (শুধু bKash বন্ধ হবে, পুরো বট না)
BKASH_ENABLED = False
BKASH_OFF_TEXT = (
    "🔴 bKash এখন বন্ধ আছে (Turned off)।\n"
    "অনুগ্রহ করে Nagad অথবা Binance দিয়ে ডিপোজিট করুন।"
)

REFERRAL_BONUS = 70  # 0.70 BDT in poisha, per successful referral
_bot_username: Optional[str] = None

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("quest_sell_bot")

# --------------------------------------------------------------------------- #
# Catalog
# --------------------------------------------------------------------------- #
PRODUCTS = {
    "meta_ai": {
        "name": "Meta AI Account",
        "label": "𝐌𝐞𝐭𝐚 𝐀𝐈 𝐀𝐜𝐜𝐨𝐮𝐧𝐭",
        "title": "Meta AI ID",
        "admin_label": "Meta Account",
        "price": 50,
    },
    "meta_horizon": {
        "name": "Meta Horizon",
        "label": "𝐌𝐞𝐭𝐚 𝐇𝐨𝐫𝐢𝐳𝐨𝐧",
        "title": "Meta Horizon",
        "admin_label": "Meta Horizon",
        "price": 55,
    },
}

METHODS = {
    "bkash": {"name": "bKash", "mfs": True},
    "nagad": {"name": "Nagad", "mfs": True},
    "binance": {"name": "Binance", "mfs": False},
}

# Reply-keyboard labels (users tap these; handlers match the exact text)
BTN_BUY = "🛒 𝐁𝐮𝐲 𝐏𝐫𝐨𝐝𝐮𝐜𝐭𝐬"
BTN_BAL = "💰 𝐁𝐚𝐥𝐚𝐧𝐜𝐞"
BTN_DEP = "💳 𝐃𝐞𝐩𝐨𝐬𝐢𝐭"
BTN_REF = "👥 𝐑𝐞𝐟𝐞𝐫𝐫𝐚𝐥𝐬"
BTN_SUP = "📞 𝐒𝐮𝐩𝐩𝐨𝐫𝐭"

WELCOME_TEXT = (
    "👋 Welcome to Quest Sell Bot!\n\n"
    "🛍️ আপনার পছন্দের প্রোডাক্ট কিনুন সহজেই।\n"
    "💳 ব্যালেন্স ডিপোজিট করুন এবং অর্ডার সম্পন্ন করুন।\n\n"
    "✨ Trusted Products | Fast Delivery"
)

# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
BN_DIGITS = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")


def esc(value) -> str:
    return html.escape(str(value if value is not None else ""), quote=False)


def now_ts() -> int:
    return int(time.time())


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def fmt_money(poisha: int) -> str:
    return f"{poisha // 100}.{poisha % 100:02d}"


def price_short(poisha: int) -> str:
    return f"{poisha / 100:g}"


def parse_amount(text: str) -> Optional[int]:
    """Parse a BDT amount (max 2 decimals) into integer poisha."""
    t = (text or "").strip().translate(BN_DIGITS).replace(",", "")
    if not re.fullmatch(r"\d{1,9}(\.\d{1,2})?", t):
        return None
    poisha = int(Decimal(t) * 100)
    if poisha <= 0 or poisha > MAX_AMOUNT:
        return None
    return poisha


def parse_positive_int(text: str, max_digits: int = 6) -> Optional[int]:
    t = (text or "").strip().translate(BN_DIGITS)
    if not re.fullmatch(rf"\d{{1,{max_digits}}}", t):
        return None
    value = int(t)
    return value if value > 0 else None


def mfs_is_open() -> bool:
    current = datetime.now(DHAKA).time()
    return dtime(9, 0) <= current < dtime(23, 0)


def is_admin(user_id: Optional[int]) -> bool:
    return user_id in ADMIN_IDS


def is_primary(user_id: Optional[int]) -> bool:
    return user_id == PRIMARY_ADMIN_ID


# --------------------------------------------------------------------------- #
# Colored keyboard builders (Bot API 9.4 `style`)
# --------------------------------------------------------------------------- #
def ibtn(text: str, callback_data: Optional[str] = None, url: Optional[str] = None,
         style: str = "primary") -> InlineKeyboardButton:
    return InlineKeyboardButton(text=text, callback_data=callback_data, url=url, style=style)


def kbtn(text: str, style: str = "primary") -> KeyboardButton:
    return KeyboardButton(text=text, style=style)


def ikb(*rows) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[list(r) for r in rows])


def main_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [kbtn(BTN_BUY, "primary")],
            [kbtn(BTN_BAL, "success"), kbtn(BTN_DEP, "success")],
            [kbtn(BTN_REF, "success"), kbtn(BTN_SUP, "primary")],
        ],
        resize_keyboard=True,
    )


def cancel_flow_kb() -> InlineKeyboardMarkup:
    return ikb([ibtn("❌ 𝐂𝐚𝐧𝐜𝐞𝐥", "flowcancel", style="danger")])


# --------------------------------------------------------------------------- #
# Database
# --------------------------------------------------------------------------- #
_db_dir = os.path.dirname(os.path.abspath(DB_PATH))
os.makedirs(_db_dir, exist_ok=True)

db = sqlite3.connect(DB_PATH, isolation_level=None, check_same_thread=False, timeout=30)
db.row_factory = sqlite3.Row


class PurchaseError(Exception):
    pass


@contextmanager
def tx():
    """Atomic transaction. All DB calls are synchronous, so no await can interleave inside."""
    db.execute("BEGIN IMMEDIATE")
    try:
        yield db
        db.execute("COMMIT")
    except BaseException:
        db.execute("ROLLBACK")
        raise


def init_db() -> None:
    db.execute("PRAGMA journal_mode=WAL")
    db.execute("PRAGMA synchronous=NORMAL")
    db.execute("PRAGMA foreign_keys=ON")
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            full_name TEXT,
            balance INTEGER NOT NULL DEFAULT 0 CHECK (balance >= 0),
            started INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS products (
            code TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            price INTEGER NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_code TEXT NOT NULL,
            email TEXT NOT NULL,
            password TEXT NOT NULL,
            available INTEGER NOT NULL DEFAULT 1,
            order_id INTEGER,
            UNIQUE (product_code, email)
        );
        CREATE INDEX IF NOT EXISTS idx_accounts_stock ON accounts (product_code, available);
        CREATE INDEX IF NOT EXISTS idx_accounts_order ON accounts (order_id);
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            product_code TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            unit_price INTEGER NOT NULL,
            total_price INTEGER NOT NULL,
            status TEXT NOT NULL,
            delivered INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            created_ts INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_orders_user ON orders (user_id, status);
        CREATE TABLE IF NOT EXISTS deposits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            method TEXT NOT NULL,
            amount INTEGER NOT NULL,
            reference TEXT NOT NULL,
            status TEXT NOT NULL,
            processed_by INTEGER,
            created_at TEXT NOT NULL,
            processed_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_deposits_ref ON deposits (method, reference);
        CREATE TABLE IF NOT EXISTS deposit_msgs (
            deposit_id INTEGER NOT NULL,
            chat_id INTEGER NOT NULL,
            message_id INTEGER NOT NULL
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """
    )
    for code, meta in PRODUCTS.items():
        db.execute(
            "INSERT INTO products (code, name, price, active) VALUES (?, ?, ?, 1) "
            "ON CONFLICT(code) DO UPDATE SET price=excluded.price",
            (code, meta["name"], meta["price"]),
        )
    db.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('bot_enabled', '1')")
    cols = [r["name"] for r in db.execute("PRAGMA table_info(users)").fetchall()]
    if "referred_by" not in cols:
        db.execute("ALTER TABLE users ADD COLUMN referred_by INTEGER")


# ---- settings ------------------------------------------------------------- #
def is_enabled() -> bool:
    row = db.execute("SELECT value FROM settings WHERE key='bot_enabled'").fetchone()
    return row is None or row["value"] == "1"


def set_enabled(flag: bool) -> None:
    db.execute(
        "INSERT INTO settings (key, value) VALUES ('bot_enabled', ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        ("1" if flag else "0",),
    )


# ---- users ---------------------------------------------------------------- #
def upsert_user(u: User) -> bool:
    """Insert/update a user. Returns True if this is the user's first time starting the bot."""
    prev = db.execute("SELECT started FROM users WHERE user_id=?", (u.id,)).fetchone()
    is_new = prev is None or not prev["started"]
    db.execute(
        "INSERT INTO users (user_id, username, full_name, balance, started, created_at) "
        "VALUES (?, ?, ?, 0, 1, ?) "
        "ON CONFLICT(user_id) DO UPDATE SET username=excluded.username, "
        "full_name=excluded.full_name, started=1",
        (u.id, u.username, u.full_name, now_iso()),
    )
    return is_new


def get_user(user_id: int):
    return db.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()


def get_balance(user_id: int) -> int:
    row = get_user(user_id)
    return row["balance"] if row else 0


def add_balance(user_id: int, amount: int) -> int:
    with tx():
        db.execute(
            "INSERT OR IGNORE INTO users (user_id, username, full_name, balance, started, created_at) "
            "VALUES (?, NULL, '', 0, 0, ?)",
            (user_id, now_iso()),
        )
        db.execute("UPDATE users SET balance = balance + ? WHERE user_id=?", (amount, user_id))
        return db.execute("SELECT balance FROM users WHERE user_id=?", (user_id,)).fetchone()["balance"]


# ---- referrals ------------------------------------------------------------ #
def register_referral(new_user_id: int, referrer_id: int) -> Optional[int]:
    """Credit the referrer once for a brand-new user. Returns referrer's new balance or None."""
    if new_user_id == referrer_id:
        return None
    with tx():
        ref = db.execute(
            "SELECT started FROM users WHERE user_id=?", (referrer_id,)
        ).fetchone()
        if ref is None or not ref["started"]:
            return None
        cur = db.execute(
            "UPDATE users SET referred_by=? WHERE user_id=? AND referred_by IS NULL",
            (referrer_id, new_user_id),
        )
        if cur.rowcount != 1:
            return None
        db.execute(
            "UPDATE users SET balance = balance + ? WHERE user_id=?",
            (REFERRAL_BONUS, referrer_id),
        )
        return db.execute(
            "SELECT balance FROM users WHERE user_id=?", (referrer_id,)
        ).fetchone()["balance"]


def count_referrals(user_id: int) -> int:
    return db.execute(
        "SELECT COUNT(*) AS c FROM users WHERE referred_by=?", (user_id,)
    ).fetchone()["c"]


# ---- products / stock ----------------------------------------------------- #
def get_product(code: str):
    return db.execute("SELECT * FROM products WHERE code=?", (code,)).fetchone()


def get_stock(code: str) -> int:
    return db.execute(
        "SELECT COUNT(*) AS c FROM accounts WHERE product_code=? AND available=1", (code,)
    ).fetchone()["c"]


def import_accounts(code: str, rows: list) -> int:
    imported = 0
    with tx():
        for email, password in rows:
            cur = db.execute(
                "INSERT OR IGNORE INTO accounts (product_code, email, password, available) "
                "VALUES (?, ?, ?, 1)",
                (code, email, password),
            )
            imported += cur.rowcount
    return imported


def get_order_accounts(order_id: int) -> list:
    rows = db.execute(
        "SELECT email, password FROM accounts WHERE order_id=? ORDER BY id", (order_id,)
    ).fetchall()
    return [(r["email"], r["password"]) for r in rows]


# ---- orders --------------------------------------------------------------- #
def expire_pending() -> None:
    db.execute(
        "UPDATE orders SET status='CANCELLED' WHERE status='PENDING' AND created_ts < ?",
        (now_ts() - PENDING_TTL,),
    )


def get_pending_order(user_id: int):
    expire_pending()
    return db.execute(
        "SELECT * FROM orders WHERE user_id=? AND status='PENDING' ORDER BY id DESC LIMIT 1",
        (user_id,),
    ).fetchone()


def get_order(order_id: int):
    return db.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()


def create_pending_order(user_id: int, code: str, quantity: int) -> int:
    product = get_product(code)
    unit = product["price"]
    with tx():
        cur = db.execute(
            "INSERT INTO orders (user_id, product_code, quantity, unit_price, total_price, "
            "status, delivered, created_at, created_ts) VALUES (?, ?, ?, ?, ?, 'PENDING', 0, ?, ?)",
            (user_id, code, quantity, unit, unit * quantity, now_iso(), now_ts()),
        )
        return cur.lastrowid


def cancel_order(order_id: int, user_id: int) -> bool:
    with tx():
        cur = db.execute(
            "UPDATE orders SET status='CANCELLED' WHERE id=? AND user_id=? AND status='PENDING'",
            (order_id, user_id),
        )
        return cur.rowcount == 1


def do_purchase(order_id: int, user_id: int) -> dict:
    with tx():
        order = db.execute(
            "SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user_id)
        ).fetchone()
        if order is None or order["status"] != "PENDING":
            return {"status": "invalid"}
        if order["created_ts"] < now_ts() - PENDING_TTL:
            db.execute("UPDATE orders SET status='CANCELLED' WHERE id=?", (order_id,))
            return {"status": "expired"}
        product = get_product(order["product_code"])
        if product is None or not product["active"]:
            db.execute("UPDATE orders SET status='CANCELLED' WHERE id=?", (order_id,))
            return {"status": "inactive"}
        total = order["total_price"]
        quantity = order["quantity"]
        balance = db.execute("SELECT balance FROM users WHERE user_id=?", (user_id,)).fetchone()
        balance = balance["balance"] if balance else 0
        if balance < total:
            return {"status": "balance", "need": total, "balance": balance}
        picked = db.execute(
            "SELECT id FROM accounts WHERE product_code=? AND available=1 ORDER BY id LIMIT ?",
            (order["product_code"], quantity),
        ).fetchall()
        if len(picked) < quantity:
            db.execute("UPDATE orders SET status='CANCELLED' WHERE id=?", (order_id,))
            return {"status": "stock", "stock": len(picked)}
        ids = [r["id"] for r in picked]
        marks = ",".join("?" * len(ids))
        cur = db.execute(
            f"UPDATE accounts SET available=0, order_id=? WHERE available=1 AND id IN ({marks})",
            (order_id, *ids),
        )
        if cur.rowcount != quantity:
            raise PurchaseError("account reservation mismatch")
        cur = db.execute(
            "UPDATE users SET balance = balance - ? WHERE user_id=? AND balance >= ?",
            (total, user_id, total),
        )
        if cur.rowcount != 1:
            raise PurchaseError("balance deduction failed")
        cur = db.execute(
            "UPDATE orders SET status='COMPLETED' WHERE id=? AND status='PENDING'", (order_id,)
        )
        if cur.rowcount != 1:
            raise PurchaseError("order state changed")
        return {"status": "ok"}


# ---- deposits ------------------------------------------------------------- #
def reference_exists(method: str, reference: str) -> bool:
    row = db.execute(
        "SELECT 1 FROM deposits WHERE method=? AND reference=? AND status IN ('PENDING','APPROVED')",
        (method, reference),
    ).fetchone()
    return row is not None


def create_deposit(user_id: int, method: str, amount: int, reference: str) -> int:
    with tx():
        cur = db.execute(
            "INSERT INTO deposits (user_id, method, amount, reference, status, created_at) "
            "VALUES (?, ?, ?, ?, 'PENDING', ?)",
            (user_id, method, amount, reference, now_iso()),
        )
        return cur.lastrowid


def get_deposit(dep_id: int):
    return db.execute("SELECT * FROM deposits WHERE id=?", (dep_id,)).fetchone()


def finalize_deposit(dep_id: int, admin_id: int, approve: bool) -> Optional[dict]:
    """Atomically approve/cancel a pending deposit exactly once. Returns None if already processed."""
    with tx():
        dep = db.execute("SELECT * FROM deposits WHERE id=?", (dep_id,)).fetchone()
        if dep is None or dep["status"] != "PENDING":
            return None
        new_status = "APPROVED" if approve else "CANCELLED"
        cur = db.execute(
            "UPDATE deposits SET status=?, processed_by=?, processed_at=? "
            "WHERE id=? AND status='PENDING'",
            (new_status, admin_id, now_iso(), dep_id),
        )
        if cur.rowcount != 1:
            return None
        new_balance = None
        if approve:
            db.execute(
                "INSERT OR IGNORE INTO users (user_id, username, full_name, balance, started, created_at) "
                "VALUES (?, NULL, '', 0, 0, ?)",
                (dep["user_id"], now_iso()),
            )
            db.execute(
                "UPDATE users SET balance = balance + ? WHERE user_id=?",
                (dep["amount"], dep["user_id"]),
            )
            new_balance = db.execute(
                "SELECT balance FROM users WHERE user_id=?", (dep["user_id"],)
            ).fetchone()["balance"]
        return {"user_id": dep["user_id"], "amount": dep["amount"], "balance": new_balance}


def save_deposit_msg(dep_id: int, chat_id: int, message_id: int) -> None:
    db.execute(
        "INSERT INTO deposit_msgs (deposit_id, chat_id, message_id) VALUES (?, ?, ?)",
        (dep_id, chat_id, message_id),
    )


def get_deposit_msgs(dep_id: int) -> list:
    return db.execute(
        "SELECT chat_id, message_id FROM deposit_msgs WHERE deposit_id=?", (dep_id,)
    ).fetchall()


# --------------------------------------------------------------------------- #
# Text builders
# --------------------------------------------------------------------------- #
def order_summary_text(order) -> str:
    meta = PRODUCTS[order["product_code"]]
    return (
        f"🛒 <b>{esc(meta['title'])}</b>\n\n"
        f"💰 প্রাইস: {fmt_money(order['unit_price'])} TK\n"
        f"স্টক: {get_stock(order['product_code'])}\n\n"
        f"পরিমাণ: {order['quantity']}\n"
        f"মোট খরচ: {fmt_money(order['total_price'])} TK"
    )


def order_kb(order_id: int) -> InlineKeyboardMarkup:
    return ikb([
        ibtn("✅ 𝐂𝐨𝐧𝐟𝐢𝐫𝐦 𝐎𝐫𝐝𝐞𝐫", f"cfm:{order_id}", style="success"),
        ibtn("❌ 𝐂𝐚𝐧𝐜𝐞𝐥", f"cnl:{order_id}", style="danger"),
    ])


def products_kb() -> InlineKeyboardMarkup:
    rows = []
    for code, meta in PRODUCTS.items():
        product = get_product(code)
        rows.append([
            ibtn(
                f"{meta['label']} | {price_short(product['price'])} BDT | Stock: {get_stock(code)}",
                f"buy:{code}",
                style="primary",
            )
        ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def deposit_admin_text(dep, user, note: Optional[str] = None) -> str:
    method = METHODS[dep["method"]]
    uname = f"@{esc(user['username'])}" if user is not None and user["username"] else "Not set"
    lines = [
        "💳 <b>ডিপোজিট রিকোয়েস্ট</b>",
        "",
        f"🧾 Request ID: <code>{dep['id']}</code>",
        f"💱 Method: {method['name']}",
        f"👤 Username: {uname}",
        f"🆔 Telegram ID: <code>{dep['user_id']}</code>",
        f"💰 Amount: {fmt_money(dep['amount'])} TK",
    ]
    if dep["method"] == "binance":
        lines.append(f"🏦 Binance Pay UID: <code>{BINANCE_UID}</code>")
        lines.append(f"🔖 Order ID: <code>{esc(dep['reference'])}</code>")
    else:
        lines.append(f"🔖 TrxID: <code>{esc(dep['reference'])}</code>")
    if note:
        lines += ["", note]
    return "\n".join(lines)


def build_xlsx(accounts: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Accounts"
    ws["A1"] = "Email"
    ws["B1"] = "Password"
    for row_no, (email, password) in enumerate(accounts, start=2):
        for col, value in ((1, email), (2, password)):
            cell = ws.cell(row=row_no, column=col, value=value)
            cell.data_type = "s"  # never let uploaded text become a formula
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 26
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Telegram helpers
# --------------------------------------------------------------------------- #
async def safe_edit(message, text: str, kb: Optional[InlineKeyboardMarkup] = None) -> None:
    if not isinstance(message, Message):
        return
    try:
        await message.edit_text(text, reply_markup=kb)
    except TelegramBadRequest:
        pass


async def safe_clear_kb(message) -> None:
    if not isinstance(message, Message):
        return
    try:
        await message.edit_reply_markup(reply_markup=None)
    except TelegramBadRequest:
        pass


async def safe_send(bot: Bot, chat_id: int, text: str, **kwargs) -> Optional[Message]:
    for _ in range(3):
        try:
            return await bot.send_message(chat_id, text, **kwargs)
        except TelegramRetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
        except TelegramAPIError as e:
            log.warning("send_message failed for %s: %s", chat_id, type(e).__name__)
            return None
    return None


def state_expired(data: dict, ttl: int) -> bool:
    return time.time() - data.get("ts", 0) > ttl


# --------------------------------------------------------------------------- #
# FSM states
# --------------------------------------------------------------------------- #
class BuyFlow(StatesGroup):
    quantity = State()


class DepositFlow(StatesGroup):
    amount = State()
    reference = State()


class AdminAdd(StatesGroup):
    user_id = State()
    amount = State()


class AdminUpload(StatesGroup):
    file = State()


class AdminBroadcast(StatesGroup):
    message = State()


# --------------------------------------------------------------------------- #
# Middleware: user registration + bot on/off gate
# --------------------------------------------------------------------------- #
ADMIN_CALLBACK_PREFIXES = ("dep:", "acc:", "flowcancel")


class GateMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, data):
        user: Optional[User] = data.get("event_from_user")
        if user is None or user.is_bot:
            return await handler(event, data)

        data["is_new_user"] = upsert_user(user)

        if is_enabled():
            return await handler(event, data)

        # ---- bot is disabled ----
        if isinstance(event, Message):
            text = event.text or ""
            if text.startswith("/start"):
                await event.answer(OFF_TEXT)
                return None
            if is_admin(user.id):
                if text.startswith("/"):
                    return await handler(event, data)
                state: FSMContext = data.get("state")
                current = await state.get_state() if state else None
                if current and current.startswith("Admin"):
                    return await handler(event, data)
            await event.answer(OFF_TEXT)
            return None

        if isinstance(event, CallbackQuery):
            if is_admin(user.id) and (event.data or "").startswith(ADMIN_CALLBACK_PREFIXES):
                return await handler(event, data)
            await event.answer(OFF_TEXT, show_alert=True)
            return None

        return await handler(event, data)


router = Router()
router.message.filter(F.chat.type == "private")
router.message.outer_middleware(GateMiddleware())
router.callback_query.outer_middleware(GateMiddleware())


def not_command(message: Message) -> bool:
    return not (message.text or "").startswith("/")


# --------------------------------------------------------------------------- #
# Commands
# --------------------------------------------------------------------------- #
@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext, bot: Bot, command: CommandObject,
                    is_new_user: bool = False):
    await state.clear()
    await message.answer(WELCOME_TEXT, reply_markup=main_menu())

    # Referral: only for brand-new users opening a t.me/<bot>?start=ref_<id> link
    m = re.fullmatch(r"ref_(\d{5,15})", (command.args or "").strip())
    if not (m and is_new_user):
        return
    referrer_id = int(m.group(1))
    new_balance = register_referral(message.from_user.id, referrer_id)
    if new_balance is None:
        return
    name = esc(message.from_user.full_name)
    await safe_send(
        bot,
        referrer_id,
        "🎉 <b>নতুন রেফারেল!</b>\n\n"
        f"👤 {name} আপনার লিংকে জয়েন করেছে।\n"
        f"➕ বোনাস: {fmt_money(REFERRAL_BONUS)} TK\n"
        f"👛 নতুন ব্যালেন্স: {fmt_money(new_balance)} TK",
    )


@router.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("❌ ক্যানসেল করা হয়েছে।", reply_markup=main_menu())


@router.message(Command("retry"))
async def cmd_retry(message: Message, bot: Bot):
    rows = db.execute(
        "SELECT id FROM orders WHERE user_id=? AND status='COMPLETED' AND delivered=0 ORDER BY id",
        (message.from_user.id,),
    ).fetchall()
    if not rows:
        await message.answer("✅ কোনো পেন্ডিং ডেলিভারি নেই।")
        return
    for row in rows:
        ok = await deliver_order(bot, row["id"])
        if not ok:
            await message.answer(
                f"⚠️ অর্ডার #{row['id']} এর ফাইল এখনো পাঠানো যায়নি। কিছুক্ষণ পর আবার চেষ্টা করুন।",
                reply_markup=ikb([ibtn("🔄 আবার চেষ্টা করুন", f"rtry:{row['id']}", style="success")]),
            )


@router.message(Command("off"))
async def cmd_off(message: Message):
    if not is_primary(message.from_user.id):
        await message.answer("⛔ এই কমান্ড শুধু প্রাইমারি অ্যাডমিনের জন্য।")
        return
    set_enabled(False)
    await message.answer("🔴 বট বন্ধ করা হয়েছে। চালু করতে /on দিন।")


@router.message(Command("on"))
async def cmd_on(message: Message):
    if not is_primary(message.from_user.id):
        await message.answer("⛔ এই কমান্ড শুধু প্রাইমারি অ্যাডমিনের জন্য।")
        return
    set_enabled(True)
    await message.answer("🟢 বট আবার চালু করা হয়েছে।", reply_markup=main_menu())


@router.message(Command("add"))
async def cmd_add(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ আপনি অ্যাডমিন নন।")
        return
    await state.clear()
    await state.set_state(AdminAdd.user_id)
    await message.answer("👤 টার্গেট ইউজারের Telegram ID পাঠান:", reply_markup=cancel_flow_kb())


@router.message(Command("account", ignore_case=True))
async def cmd_account(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ আপনি অ্যাডমিন নন।")
        return
    await state.clear()
    rows = [[ibtn(meta["admin_label"], f"acc:{code}", style="primary")] for code, meta in PRODUCTS.items()]
    await message.answer("📦 Select Products", reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))


@router.message(Command("broadcast"))
async def cmd_broadcast(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ আপনি অ্যাডমিন নন।")
        return
    await state.clear()
    await state.set_state(AdminBroadcast.message)
    await message.answer("📢 আপনার মেসেজ লিখুন বা পাঠান।", reply_markup=cancel_flow_kb())


# --------------------------------------------------------------------------- #
# Main menu buttons
# --------------------------------------------------------------------------- #
@router.message(F.text == BTN_BUY)
async def menu_buy(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "🛍️ <b>Trusted Products</b> - প্রোডাক্ট সিলেক্ট করুন:", reply_markup=products_kb()
    )


@router.message(F.text == BTN_BAL)
async def menu_balance(message: Message, state: FSMContext):
    await state.clear()
    user = get_user(message.from_user.id)
    username = f"@{esc(user['username'])}" if user and user["username"] else "Not set"
    full_name = esc(message.from_user.full_name)
    balance = fmt_money(user["balance"] if user else 0)
    await message.answer(
        "👤 <b>আমার প্রোফাইল</b>\n\n"
        f"🆔 User Name: {username}\n"
        f"👤 নাম: {NAME_PREFIX}{full_name}\n"
        f"💰 ব্যালেন্স: {balance} TK"
    )


@router.message(F.text == BTN_DEP)
async def menu_deposit(message: Message, state: FSMContext):
    await state.clear()
    kb = ikb(
        [ibtn("🏦 𝐁𝐤𝐚𝐬𝐡", "pay:bkash", style="danger")],
        [ibtn("💱 𝐍𝐚𝐠𝐚𝐝", "pay:nagad", style="primary")],
        [ibtn("💳 𝐁𝐢𝐧𝐚𝐧𝐜𝐞", "pay:binance", style="success")],
    )
    await message.answer(
        "💳 <b>ডিপোজিট মেথড সিলেক্ট করুন:</b>\n\n"
        "💡 নোট: bKash, Nagad ডিপোজিট প্রতিদিন সকাল ০৯:০০ AM থেকে রাত ১১:০০ PM পর্যন্ত সক্রিয় থাকে। "
        "Binance Pay ২৪ ঘণ্টাই সচল।",
        reply_markup=kb,
    )


@router.message(F.text == BTN_REF)
async def menu_referral(message: Message, state: FSMContext, bot: Bot):
    await state.clear()
    global _bot_username
    if _bot_username is None:
        _bot_username = (await bot.get_me()).username
    link = f"https://t.me/{_bot_username}?start=ref_{message.from_user.id}"
    total = count_referrals(message.from_user.id)
    await message.answer(
        "👥 <b>রেফারেল সিস্টেম</b>\n\n"
        f"🎁 প্রতি রেফারে পাবেন <b>{fmt_money(REFERRAL_BONUS)} BDT</b>!\n\n"
        "🔗 আপনার রেফার লিংক:\n"
        f"<code>{link}</code>\n\n"
        f"👤 মোট রেফার: {total} জন\n"
        f"💰 মোট আয়: {fmt_money(total * REFERRAL_BONUS)} TK\n\n"
        "📌 লিংকটি বন্ধুদের শেয়ার করুন। কেউ আপনার লিংকে প্রথমবার বট স্টার্ট করলেই বোনাস ব্যালেন্সে যোগ হবে।"
    )


@router.message(F.text == BTN_SUP)
async def menu_support(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "📞 <b>সাপোর্ট ও সাহায্য</b>\n\n"
        "যেকোনো সমস্যা বা সহায়তার জন্য নিচের বাটনে ক্লিক করে অ্যাডমিনের সাথে যোগাযোগ করুন:",
        reply_markup=ikb([ibtn("👨‍💻 𝐂𝐨𝐧𝐭𝐚𝐜𝐭 𝐀𝐝𝐦𝐢𝐧", url=SUPPORT_URL, style="primary")]),
    )


# --------------------------------------------------------------------------- #
# Generic cancel button
# --------------------------------------------------------------------------- #
@router.callback_query(F.data == "flowcancel")
async def cb_flowcancel(call: CallbackQuery, state: FSMContext, bot: Bot):
    await state.clear()
    await call.answer()
    await safe_edit(call.message, "❌ ক্যানসেল করা হয়েছে।")
    await safe_send(bot, call.from_user.id, "🏠 মেইন মেনু", reply_markup=main_menu())


# --------------------------------------------------------------------------- #
# Buying flow
# --------------------------------------------------------------------------- #
@router.callback_query(F.data.startswith("buy:"))
async def cb_buy(call: CallbackQuery, state: FSMContext, bot: Bot):
    code = call.data.split(":", 1)[1]
    if code not in PRODUCTS:
        await call.answer("❌ প্রোডাক্ট পাওয়া যায়নি।", show_alert=True)
        return
    pending = get_pending_order(call.from_user.id)
    if pending:
        await call.answer("⚠️ আগের অর্ডারটি আগে Confirm বা Cancel করুন।", show_alert=True)
        await safe_send(bot, call.from_user.id, order_summary_text(pending), reply_markup=order_kb(pending["id"]))
        return
    product = get_product(code)
    if product is None or not product["active"]:
        await call.answer("❌ এই প্রোডাক্টটি এখন বন্ধ আছে।", show_alert=True)
        return
    stock = get_stock(code)
    if stock <= 0:
        await call.answer("❌ স্টক শেষ হয়ে গেছে।", show_alert=True)
        return
    await call.answer()
    await state.clear()
    await state.set_state(BuyFlow.quantity)
    await state.update_data(code=code, ts=time.time())
    await safe_send(
        bot,
        call.from_user.id,
        f"🛒 <b>{esc(PRODUCTS[code]['name'])}</b>\n\n"
        f"💰 প্রাইস: {fmt_money(product['price'])} TK\n"
        f"স্টক: {stock}\n\n"
        "পরিমাণ লিখুন:",
        reply_markup=cancel_flow_kb(),
    )


@router.message(BuyFlow.quantity, F.text, not_command)
async def got_quantity(message: Message, state: FSMContext):
    data = await state.get_data()
    code = data.get("code")
    if code not in PRODUCTS or state_expired(data, PENDING_TTL):
        await state.clear()
        await message.answer("⌛ সময় শেষ হয়ে গেছে। আবার শুরু করুন।", reply_markup=main_menu())
        return
    pending = get_pending_order(message.from_user.id)
    if pending:
        await state.clear()
        await message.answer("⚠️ আগের অর্ডারটি আগে Confirm বা Cancel করুন।")
        await message.answer(order_summary_text(pending), reply_markup=order_kb(pending["id"]))
        return
    quantity = parse_positive_int(message.text)
    if quantity is None:
        await message.answer("❌ সঠিক সংখ্যা লিখুন (শুধু ধনাত্মক পূর্ণসংখ্যা, যেমন: 5)।")
        return
    product = get_product(code)
    if product is None or not product["active"]:
        await state.clear()
        await message.answer("❌ এই প্রোডাক্টটি এখন বন্ধ আছে।", reply_markup=main_menu())
        return
    stock = get_stock(code)
    if quantity > stock:
        await message.answer(f"❌ পর্যাপ্ত স্টক নেই। বর্তমান স্টক: {stock}। কম পরিমাণ লিখুন।")
        return
    order_id = create_pending_order(message.from_user.id, code, quantity)
    await state.clear()
    order = get_order(order_id)
    await message.answer(order_summary_text(order), reply_markup=order_kb(order_id))


@router.callback_query(F.data.regexp(r"^cfm:\d+$"))
async def cb_confirm(call: CallbackQuery, bot: Bot):
    order_id = int(call.data.split(":", 1)[1])
    try:
        result = do_purchase(order_id, call.from_user.id)
    except Exception as e:  # never leak details; transaction already rolled back
        log.error("purchase failed for order %s: %s", order_id, type(e).__name__)
        await call.answer("❌ কিছু ভুল হয়েছে। আবার চেষ্টা করুন।", show_alert=True)
        return

    status = result["status"]
    if status == "invalid":
        await call.answer("⚠️ এই অর্ডারটি আর বৈধ নয় বা আগেই প্রসেস হয়েছে।", show_alert=True)
        await safe_clear_kb(call.message)
    elif status == "expired":
        await call.answer("⌛ অর্ডারের সময় শেষ। আবার অর্ডার করুন।", show_alert=True)
        await safe_edit(call.message, "⌛ অর্ডারের সময় শেষ হয়ে গেছে।")
    elif status == "inactive":
        await call.answer("❌ প্রোডাক্টটি এখন বন্ধ আছে।", show_alert=True)
        await safe_edit(call.message, "❌ প্রোডাক্টটি এখন বন্ধ আছে।")
    elif status == "stock":
        await call.answer("❌ স্টক অপর্যাপ্ত।", show_alert=True)
        await safe_edit(
            call.message, f"❌ স্টক অপর্যাপ্ত। বর্তমান স্টক: {result['stock']}। অর্ডার বাতিল হয়েছে।"
        )
    elif status == "balance":
        await call.answer("❌ ব্যালেন্স পর্যাপ্ত নয়।", show_alert=True)
        await safe_send(
            bot,
            call.from_user.id,
            "❌ <b>ব্যালেন্স পর্যাপ্ত নয়।</b>\n\n"
            f"💰 প্রয়োজন: {fmt_money(result['need'])} TK\n"
            f"👛 আপনার ব্যালেন্স: {fmt_money(result['balance'])} TK\n\n"
            "💳 ডিপোজিট করে আবার Confirm করুন।",
        )
    elif status == "ok":
        await call.answer("✅ অর্ডার সম্পন্ন হচ্ছে...")
        await safe_clear_kb(call.message)
        delivered = await deliver_order(bot, order_id)
        if not delivered:
            await safe_send(
                bot,
                call.from_user.id,
                "⚠️ আপনার অর্ডার সম্পন্ন হয়েছে এবং টাকা কাটা হয়েছে, কিন্তু ফাইলটি পাঠানো যায়নি।\n"
                "নিচের বাটনে চাপ দিয়ে আবার ফাইল নিন (আপনার অ্যাকাউন্টগুলো সংরক্ষিত আছে)।",
                reply_markup=ikb([ibtn("🔄 ফাইল আবার পাঠান", f"rtry:{order_id}", style="success")]),
            )


async def deliver_order(bot: Bot, order_id: int) -> bool:
    order = get_order(order_id)
    if order is None or order["status"] != "COMPLETED":
        return False
    accounts = get_order_accounts(order_id)
    if not accounts:
        return False
    meta = PRODUCTS.get(order["product_code"], {"name": order["product_code"]})
    caption = (
        "✅ <b>অর্ডার সফল হয়েছে!</b>\n\n"
        f"🧾 অর্ডার ID: <code>#{order_id}</code>\n"
        f"📦 প্রোডাক্ট: {esc(meta['name'])}\n"
        f"🔢 পরিমাণ: {order['quantity']}\n"
        f"💰 কাটা হয়েছে: {fmt_money(order['total_price'])} TK"
    )
    payload = build_xlsx(accounts)
    filename = f"order_{order_id}_{order['product_code']}.xlsx"
    for _ in range(3):
        try:
            await bot.send_document(
                order["user_id"], BufferedInputFile(payload, filename=filename), caption=caption
            )
            db.execute("UPDATE orders SET delivered=1 WHERE id=?", (order_id,))
            return True
        except TelegramRetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
        except TelegramAPIError as e:
            log.warning("delivery failed for order %s: %s", order_id, type(e).__name__)
            return False
    return False


@router.callback_query(F.data.regexp(r"^rtry:\d+$"))
async def cb_retry(call: CallbackQuery, bot: Bot):
    order_id = int(call.data.split(":", 1)[1])
    order = get_order(order_id)
    if (
        order is None
        or order["user_id"] != call.from_user.id
        or order["status"] != "COMPLETED"
        or order["delivered"]
    ):
        await call.answer("⚠️ এই ডেলিভারি আর পেন্ডিং নেই।", show_alert=True)
        return
    ok = await deliver_order(bot, order_id)
    if ok:
        await call.answer("✅ ফাইল পাঠানো হয়েছে।")
        await safe_clear_kb(call.message)
    else:
        await call.answer("❌ এখনো পাঠানো যায়নি। একটু পর আবার চেষ্টা করুন।", show_alert=True)


@router.callback_query(F.data.regexp(r"^cnl:\d+$"))
async def cb_cancel_order(call: CallbackQuery, state: FSMContext, bot: Bot):
    order_id = int(call.data.split(":", 1)[1])
    ok = cancel_order(order_id, call.from_user.id)
    await state.clear()
    if not ok:
        await call.answer("⚠️ এই অর্ডারটি আগেই প্রসেস বা বাতিল হয়েছে।", show_alert=True)
        await safe_clear_kb(call.message)
        return
    await call.answer("❌ অর্ডার বাতিল হয়েছে।")
    await safe_edit(call.message, "❌ অর্ডার ক্যানসেল করা হয়েছে। কোনো ব্যালেন্স কাটা হয়নি।")
    await safe_send(bot, call.from_user.id, "🏠 মেইন মেনু", reply_markup=main_menu())


# --------------------------------------------------------------------------- #
# Deposit flow
# --------------------------------------------------------------------------- #
def mfs_closed_text(method_name: str) -> str:
    return (
        f"⏰ {method_name} ডিপোজিট এখন বন্ধ আছে।\n"
        "bKash ও Nagad ডিপোজিট প্রতিদিন সকাল ০৯:০০ AM থেকে রাত ১১:০০ PM (বাংলাদেশ সময়) পর্যন্ত সক্রিয় থাকে।\n"
        "আপনি চাইলে Binance ব্যবহার করতে পারেন (২৪ ঘণ্টা সচল)।"
    )


@router.callback_query(F.data.regexp(r"^pay:(bkash|nagad|binance)$"))
async def cb_pay(call: CallbackQuery, state: FSMContext, bot: Bot):
    method = call.data.split(":", 1)[1]
    info = METHODS[method]
    if method == "bkash" and not BKASH_ENABLED:
        await call.answer(BKASH_OFF_TEXT, show_alert=True)
        return
    if info["mfs"] and not mfs_is_open():
        await call.answer(mfs_closed_text(info["name"]), show_alert=True)
        return
    await call.answer()
    await state.clear()
    await state.set_state(DepositFlow.amount)
    await state.update_data(method=method, ts=time.time())
    title = "Binance Pay" if method == "binance" else info["name"]
    await safe_send(
        bot,
        call.from_user.id,
        f"💳 <b>{title} ডিপোজিট</b>\n\n"
        "আপনি কত টাকা ডিপোজিট করতে চান তা লিখে পাঠান:\n"
        "(সর্বনিম্ন: 20 টাকা)",
        reply_markup=cancel_flow_kb(),
    )


def payment_instruction_text(method: str, amount: int) -> str:
    amt = fmt_money(amount)
    if method == "bkash":
        return (
            "🎉 <b>অটো পেমেন্ট সিস্টেম তৈরি করা হয়েছে</b>\n☑️\n\n"
            f"☑️ পরিমাণ: <b>{amt}</b> টাকা (ঠিক এই অ্যামাউন্ট পাঠাবেন)\n\n"
            "💚 <b>BKASH</b>\n\n"
            "📱 Personal (Send Money):\n"
            f"<code>{BKASH_NUMBER}</code>\n\n"
            "👨‍💻 পেমেন্ট করার পর আপনার TrxID এখানে পাঠান 📍\n\n"
            "⚠️ আগে পেমেন্ট করুন, তারপর সঠিক TrxID পাঠান। অ্যাডমিন ম্যানুয়ালি যাচাই করে ব্যালেন্স যোগ করবেন।"
        )
    if method == "nagad":
        return (
            "🎉 <b>অটো পেমেন্ট সিস্টেম তৈরি করা হয়েছে</b>\n☑️\n\n"
            f"☑️ পরিমাণ: <b>{amt}</b> টাকা (ঠিক এই অ্যামাউন্ট পাঠাবেন)\n\n"
            "🧡 <b>NAGAD</b>\n\n"
            "📱 Personal (Send Money):\n"
            f"<code>{NAGAD_NUMBER}</code>\n\n"
            "👨‍💻 পেমেন্ট করার পর আপনার TrxID এখানে পাঠান 📍\n\n"
            "⚠️ আগে পেমেন্ট করুন, তারপর সঠিক TrxID পাঠান। অ্যাডমিন ম্যানুয়ালি যাচাই করে ব্যালেন্স যোগ করবেন।"
        )
    return (
        "🎉 <b>পেমেন্ট ইনস্ট্রাকশন তৈরি করা হয়েছে</b>\n☑️\n\n"
        f"☑️ পরিমাণ: <b>{amt}</b> টাকা (ঠিক এই অ্যামাউন্ট পাঠাবেন)\n\n"
        "💛 <b>BINANCE PAY</b>\n\n"
        "🆔 Binance Pay UID:\n"
        f"<code>{BINANCE_UID}</code>\n\n"
        "👨‍💻 পেমেন্ট করার পর আপনার Binance Order ID এখানে পাঠান 📍\n\n"
        "⚠️ আগে পেমেন্ট করুন, তারপর সঠিক Order ID পাঠান। অ্যাডমিন ম্যানুয়ালি যাচাই করে ব্যালেন্স যোগ করবেন।"
    )


@router.message(DepositFlow.amount, F.text, not_command)
async def got_deposit_amount(message: Message, state: FSMContext):
    data = await state.get_data()
    method = data.get("method")
    if method not in METHODS or state_expired(data, AMOUNT_STATE_TTL):
        await state.clear()
        await message.answer("⌛ সময় শেষ হয়ে গেছে। আবার ডিপোজিট শুরু করুন।", reply_markup=main_menu())
        return
    info = METHODS[method]
    if method == "bkash" and not BKASH_ENABLED:
        await state.clear()
        await message.answer(BKASH_OFF_TEXT, reply_markup=main_menu())
        return
    if info["mfs"] and not mfs_is_open():
        await state.clear()
        await message.answer(mfs_closed_text(info["name"]), reply_markup=main_menu())
        return
    amount = parse_amount(message.text)
    if amount is None:
        await message.answer("❌ সঠিক পরিমাণ লিখুন (শুধু সংখ্যা, যেমন: 100)।")
        return
    if amount < MIN_DEPOSIT:
        await message.answer("❌ সর্বনিম্ন ডিপোজিট 20 টাকা। আবার লিখুন:")
        return
    await state.set_state(DepositFlow.reference)
    await state.update_data(amount=amount, ts=time.time())
    await message.answer(payment_instruction_text(method, amount), reply_markup=cancel_flow_kb())


@router.message(DepositFlow.reference, F.text, not_command)
async def got_deposit_reference(message: Message, state: FSMContext, bot: Bot):
    data = await state.get_data()
    method = data.get("method")
    amount = data.get("amount")
    if method not in METHODS or not isinstance(amount, int) or state_expired(data, REFERENCE_STATE_TTL):
        await state.clear()
        await message.answer("⌛ সময় শেষ হয়ে গেছে। আবার ডিপোজিট শুরু করুন।", reply_markup=main_menu())
        return
    reference = (message.text or "").strip()
    label = "Order ID" if method == "binance" else "TrxID"
    if not re.fullmatch(r"[A-Za-z0-9_\-]{6,64}", reference):
        await message.answer(f"❌ সঠিক {label} পাঠান (শুধু ইংরেজি অক্ষর/সংখ্যা, ৬-৬৪ অক্ষর)।")
        return
    if reference_exists(method, reference):
        await message.answer(f"❌ এই {label} আগেই জমা দেওয়া হয়েছে।")
        return
    dep_id = create_deposit(message.from_user.id, method, amount, reference)
    await state.clear()
    await message.answer(
        "✅ <b>আপনার ডিপোজিট রিকোয়েস্ট জমা হয়েছে।</b>\n\n"
        f"🧾 Request ID: <code>{dep_id}</code>\n"
        f"💰 পরিমাণ: {fmt_money(amount)} TK\n\n"
        "⏳ অ্যাডমিন যাচাই করার পর ব্যালেন্স যোগ হবে।",
        reply_markup=main_menu(),
    )
    await notify_admins_deposit(bot, dep_id)


async def notify_admins_deposit(bot: Bot, dep_id: int) -> None:
    dep = get_deposit(dep_id)
    user = get_user(dep["user_id"])
    text = deposit_admin_text(dep, user)
    kb = ikb([
        ibtn("✅ Confirm Deposit", f"dep:ok:{dep_id}", style="success"),
        ibtn("❌ Cancel", f"dep:no:{dep_id}", style="danger"),
    ])
    for admin_id in ADMIN_IDS:
        sent = await safe_send(bot, admin_id, text, reply_markup=kb)
        if sent:
            save_deposit_msg(dep_id, admin_id, sent.message_id)


@router.callback_query(F.data.regexp(r"^dep:(ok|no):\d+$"))
async def cb_admin_deposit(call: CallbackQuery, bot: Bot):
    if not is_admin(call.from_user.id):
        await call.answer("⛔ আপনি অ্যাডমিন নন।", show_alert=True)
        return
    _, action, raw_id = call.data.split(":")
    dep_id = int(raw_id)
    approve = action == "ok"
    try:
        result = finalize_deposit(dep_id, call.from_user.id, approve)
    except Exception as e:
        log.error("deposit finalize failed for %s: %s", dep_id, type(e).__name__)
        await call.answer("❌ কিছু ভুল হয়েছে।", show_alert=True)
        return

    dep = get_deposit(dep_id)
    user = get_user(dep["user_id"]) if dep else None
    admin_tag = f"@{esc(call.from_user.username)}" if call.from_user.username else esc(call.from_user.full_name)

    if result is None:
        await call.answer("⚠️ এই রিকোয়েস্ট আগেই প্রসেস করা হয়েছে।", show_alert=True)
        if dep is not None:
            note = {
                "APPROVED": "✅ Approved",
                "CANCELLED": "❌ Cancelled",
            }.get(dep["status"], dep["status"])
            await safe_edit(call.message, deposit_admin_text(dep, user, note))
        return

    if approve:
        note = f"✅ <b>Approved</b> by {admin_tag}"
        await call.answer("✅ ডিপোজিট অনুমোদিত।")
        user_text = (
            "✅ <b>আপনার ডিপোজিট অনুমোদিত হয়েছে!</b>\n\n"
            f"🧾 Request ID: <code>{dep_id}</code>\n"
            f"💰 যোগ হয়েছে: {fmt_money(result['amount'])} TK\n"
            f"👛 নতুন ব্যালেন্স: {fmt_money(result['balance'])} TK"
        )
    else:
        note = f"❌ <b>Cancelled</b> by {admin_tag}"
        await call.answer("❌ ডিপোজিট বাতিল।")
        user_text = (
            "❌ <b>আপনার ডিপোজিট রিকোয়েস্ট বাতিল করা হয়েছে।</b>\n\n"
            f"🧾 Request ID: <code>{dep_id}</code>\n"
            "কোনো ব্যালেন্স যোগ হয়নি। সমস্যা থাকলে সাপোর্টে যোগাযোগ করুন।"
        )

    final_text = deposit_admin_text(dep, user, note)
    for row in get_deposit_msgs(dep_id):
        try:
            await bot.edit_message_text(
                final_text, chat_id=row["chat_id"], message_id=row["message_id"], reply_markup=None
            )
        except TelegramAPIError:
            pass
    await safe_send(bot, dep["user_id"], user_text)


# --------------------------------------------------------------------------- #
# Admin: /add
# --------------------------------------------------------------------------- #
@router.message(AdminAdd.user_id, F.text, not_command)
async def add_get_user(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        await state.clear()
        return
    raw = (message.text or "").strip().translate(BN_DIGITS)
    if not re.fullmatch(r"\d{5,15}", raw):
        await message.answer("❌ সঠিক Telegram ID লিখুন (শুধু সংখ্যা)।")
        return
    await state.update_data(target=int(raw))
    await state.set_state(AdminAdd.amount)
    await message.answer("💰 কত টাকা যোগ করবেন? (যেমন: 50 বা 12.50)", reply_markup=cancel_flow_kb())


@router.message(AdminAdd.amount, F.text, not_command)
async def add_get_amount(message: Message, state: FSMContext, bot: Bot):
    if not is_admin(message.from_user.id):
        await state.clear()
        return
    data = await state.get_data()
    target = data.get("target")
    amount = parse_amount(message.text)
    if amount is None:
        await message.answer("❌ সঠিক পরিমাণ লিখুন (ধনাত্মক সংখ্যা, সর্বোচ্চ ২ দশমিক ঘর)।")
        return
    await state.clear()  # clear first so the same input can never be processed twice
    if not isinstance(target, int):
        await message.answer("❌ কিছু ভুল হয়েছে। /add দিয়ে আবার শুরু করুন।")
        return
    new_balance = add_balance(target, amount)
    await message.answer(
        "✅ <b>ব্যালেন্স যোগ করা হয়েছে</b>\n\n"
        f"🆔 ইউজার ID: <code>{target}</code>\n"
        f"💰 যোগ হয়েছে: {fmt_money(amount)} TK\n"
        f"👛 নতুন ব্যালেন্স: {fmt_money(new_balance)} TK"
    )
    user = get_user(target)
    if user is not None and user["started"]:
        sent = await safe_send(
            bot,
            target,
            "💰 <b>অ্যাডমিন আপনার ব্যালেন্সে টাকা যোগ করেছেন।</b>\n\n"
            f"➕ যোগ হয়েছে: {fmt_money(amount)} TK\n"
            f"👛 নতুন ব্যালেন্স: {fmt_money(new_balance)} TK",
        )
        if sent is None:
            await message.answer("⚠️ ইউজারকে নোটিফিকেশন পাঠানো যায়নি।")
    else:
        await message.answer("ℹ️ ইউজার এখনো বট স্টার্ট করেননি, তাই নোটিফিকেশন পাঠানো হয়নি।")


# --------------------------------------------------------------------------- #
# Admin: /Account (XLSX stock import)
# --------------------------------------------------------------------------- #
@router.callback_query(F.data.regexp(r"^acc:(meta_ai|meta_horizon)$"))
async def cb_account_select(call: CallbackQuery, state: FSMContext, bot: Bot):
    if not is_admin(call.from_user.id):
        await call.answer("⛔ আপনি অ্যাডমিন নন।", show_alert=True)
        return
    code = call.data.split(":", 1)[1]
    await call.answer()
    await state.clear()
    await state.set_state(AdminUpload.file)
    await state.update_data(code=code)
    await safe_send(
        bot,
        call.from_user.id,
        f"📤 Send your XLSX file.\n\n<i>{esc(PRODUCTS[code]['name'])} — Column A: Email, Column B: Password (হেডার থাকা ঐচ্ছিক)</i>",
        reply_markup=cancel_flow_kb(),
    )


def _norm(value) -> str:
    return str(value).strip().lower() if value is not None else ""


def _cell_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_workbook(raw: bytes) -> dict:
    wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True, keep_vba=False)
    try:
        ws = wb.worksheets[0]
        rows = ws.iter_rows(min_row=1, max_col=2, values_only=True)
        first = next(rows, None)
        if first is None:
            return {"error": "empty"}
        # Header row is optional: skip it only if A1="Email" and B1="Password"
        has_header = len(first) >= 2 and _norm(first[0]) == "email" and _norm(first[1]) == "password"
        start = 2 if has_header else 1
        if not has_header:
            rows = itertools.chain([first], rows)
        total = 0
        valid, seen, dup, invalid = [], set(), 0, []
        for line_no, row in enumerate(rows, start=start):
            if line_no > MAX_XLSX_ROWS + 1:
                return {"error": "too_many"}
            a = row[0] if len(row) > 0 else None
            b = row[1] if len(row) > 1 else None
            if a is None and b is None:
                continue
            total += 1
            email, password = _cell_str(a), _cell_str(b)
            if not email or not password:
                invalid.append(line_no)
                continue
            if email in seen:
                dup += 1
                continue
            seen.add(email)
            valid.append((email, password))
        return {"total": total, "valid": valid, "dup_file": dup, "invalid": invalid}
    finally:
        wb.close()


@router.message(AdminUpload.file, F.document)
async def admin_upload_file(message: Message, state: FSMContext, bot: Bot):
    if not is_admin(message.from_user.id):
        await state.clear()
        return
    data = await state.get_data()
    code = data.get("code")
    if code not in PRODUCTS:
        await state.clear()
        await message.answer("❌ কিছু ভুল হয়েছে। /Account দিয়ে আবার শুরু করুন।")
        return
    doc = message.document
    if not (doc.file_name or "").lower().endswith(".xlsx"):
        await message.answer("❌ শুধু .xlsx ফাইল পাঠান।")
        return
    if doc.file_size and doc.file_size > MAX_UPLOAD_BYTES:
        await message.answer(f"❌ ফাইল অনেক বড়। সর্বোচ্চ {MAX_UPLOAD_BYTES // (1024 * 1024)} MB।")
        return
    buf = io.BytesIO()
    try:
        await bot.download(doc, destination=buf)
    except TelegramAPIError:
        await message.answer("❌ ফাইল ডাউনলোড করা যায়নি। আবার পাঠান।")
        return
    raw = buf.getvalue()
    if len(raw) > MAX_UPLOAD_BYTES or not raw.startswith(b"PK"):
        await message.answer("❌ এটি বৈধ XLSX ফাইল নয়।")
        return
    try:
        parsed = await asyncio.to_thread(parse_workbook, raw)
    except Exception:
        await message.answer("❌ ফাইলটি পড়া যায়নি। বৈধ XLSX ফাইল পাঠান।")
        return
    if "error" in parsed:
        if parsed["error"] == "empty":
            await message.answer("❌ ফাইলটি খালি। কলাম A তে Email এবং কলাম B তে Password দিন।")
        else:
            await message.answer(f"❌ সর্বোচ্চ {MAX_XLSX_ROWS} সারি অনুমোদিত।")
        return
    imported = await asyncio.to_thread(import_accounts, code, parsed["valid"])
    dup_total = parsed["dup_file"] + (len(parsed["valid"]) - imported)
    await state.clear()
    invalid = parsed["invalid"]
    invalid_note = ""
    if invalid:
        shown = ", ".join(str(n) for n in invalid[:20])
        more = f" ... (+{len(invalid) - 20})" if len(invalid) > 20 else ""
        invalid_note = f"\n   সারি নম্বর: {shown}{more}"
    await message.answer(
        f"📊 <b>ইমপোর্ট সামারি — {esc(PRODUCTS[code]['name'])}</b>\n\n"
        f"📄 মোট সারি প্রসেস: {parsed['total']}\n"
        f"✅ সফলভাবে ইমপোর্ট: {imported}\n"
        f"♻️ ডুপ্লিকেট স্কিপ: {dup_total}\n"
        f"⚠️ ইনভ্যালিড স্কিপ: {len(invalid)}{invalid_note}\n"
        f"📦 আপডেটেড স্টক: {get_stock(code)}"
    )


# --------------------------------------------------------------------------- #
# Admin: /broadcast
# --------------------------------------------------------------------------- #
async def send_one(bot: Bot, user_id: int, from_chat: int, message_id: int) -> bool:
    for _ in range(3):
        try:
            await bot.copy_message(user_id, from_chat, message_id)
            return True
        except TelegramRetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
        except (TelegramForbiddenError, TelegramBadRequest):
            return False
        except TelegramAPIError:
            await asyncio.sleep(1)
    return False


@router.message(AdminBroadcast.message, not_command)
async def admin_broadcast_message(message: Message, state: FSMContext, bot: Bot):
    if not is_admin(message.from_user.id):
        await state.clear()
        return
    await state.clear()
    user_ids = [r["user_id"] for r in db.execute("SELECT user_id FROM users WHERE started=1").fetchall()]
    status = await message.answer(f"📢 ব্রডকাস্ট শুরু হয়েছে... ({len(user_ids)} জন)")
    ok = failed = 0
    for uid in user_ids:
        if await send_one(bot, uid, message.chat.id, message.message_id):
            ok += 1
        else:
            failed += 1
        await asyncio.sleep(0.05)  # stay well under Telegram's broadcast rate limit
    await safe_edit(
        status,
        "📢 <b>ব্রডকাস্ট সম্পন্ন</b>\n\n"
        f"👥 মোট ইউজার: {len(user_ids)}\n"
        f"✅ সফল: {ok}\n"
        f"❌ ব্যর্থ: {failed}",
    )


# --------------------------------------------------------------------------- #
# Fallback + error handling
# --------------------------------------------------------------------------- #
@router.message()
async def fallback(message: Message, state: FSMContext):
    if await state.get_state():
        await message.answer("❌ অবৈধ ইনপুট। সঠিকভাবে লিখুন অথবা /cancel দিয়ে বাতিল করুন।")


dp = Dispatcher(storage=MemoryStorage())
dp.include_router(router)


@dp.errors()
async def on_error(event: ErrorEvent):
    log.error("Unhandled error: %s", type(event.exception).__name__, exc_info=event.exception)
    return True


# --------------------------------------------------------------------------- #
# Entrypoint
# --------------------------------------------------------------------------- #
async def main() -> None:
    if not BOT_TOKEN:
        raise SystemExit("BOT_TOKEN environment variable is not set.")
    init_db()
    log.info("Database: %s | Admins: %s", DB_PATH, sorted(ADMIN_IDS))
    if SECOND_ADMIN_ID is None:
        log.warning("ADMIN_ID is not set - only the primary admin is active.")

    probe = InlineKeyboardButton(text="x", callback_data="x", style="success")
    if "style" not in probe.model_dump(exclude_none=True):
        log.warning("This aiogram version drops the button 'style' field - upgrade aiogram.")

    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    await bot.delete_webhook(drop_pending_updates=False)
    log.info("Quest Sell Bot started (polling).")
    await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())


if __name__ == "__main__":
    asyncio.run(main())
