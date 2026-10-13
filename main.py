# -*- coding: utf-8 -*-
"""
NEXA STORE - automated digital-goods store for Telegram (aiogram 3.x + Firebase Realtime Database).

Every reply-keyboard button and every inline-keyboard button is colored using the
Telegram Bot API 9.4 `style` field ("primary" = blue, "success" = green, "danger" = red).
"""

import asyncio
import hashlib
import html
import io
import itertools
import json
import logging
import os
import re
import threading
import time
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
import firebase_admin
from dotenv import load_dotenv
from firebase_admin import credentials, db as rtdb
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


FIREBASE_CREDENTIALS = os.getenv("FIREBASE_CREDENTIALS", "").strip()  # service-account JSON (or path to it)
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()  # Firebase Realtime Database URL

SUPPORT_URL = "https://t.me/team_passivpay_nexa"
BKASH_NUMBER = "01533235544"
NAGAD_NUMBER = "01533235544"
BINANCE_UID = "1019679566"

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
BKASH_ENABLED = True
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
        "label": "𝗠𝗲𝘁𝗮 𝗔𝗜 𝗔𝗰𝗰𝗼𝘂𝗻𝘁",
        "title": "Meta AI ID",
        "admin_label": "Meta Account",
        "price": 50,
    },
    "meta_horizon": {
        "name": "Meta Horizon",
        "label": "𝗠𝗲𝘁𝗮 𝗛𝗼𝗿𝗶𝘇𝗼𝗻",
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
BTN_BUY = "🛍️ 𝗦𝗵𝗼𝗽 𝗡𝗼𝘄"
BTN_BAL = "👛 𝗠𝘆 𝗪𝗮𝗹𝗹𝗲𝘁"
BTN_DEP = "➕ 𝗔𝗱𝗱 𝗙𝘂𝗻𝗱𝘀"
BTN_REF = "🎁 𝗥𝗲𝗳𝗲𝗿 & 𝗘𝗮𝗿𝗻"
BTN_SUP = "💬 𝗛𝗲𝗹𝗽 𝗗𝗲𝘀𝗸"

WELCOME_TEXT = (
    "💎 <b>NEXA STORE-এ স্বাগতম!</b>\n\n"
    "🛒 পছন্দের প্রোডাক্ট বেছে নিন, মুহূর্তেই ডেলিভারি পান।\n"
    "👛 ওয়ালেটে ফান্ড যোগ করুন আর নিশ্চিন্তে কেনাকাটা করুন।\n\n"
    "⚡ Instant Delivery • 🔒 Secure • 🤝 24/7 Support"
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
    return True  # bKash / Nagad deposits are open 24/7


def is_admin(user_id: Optional[int]) -> bool:
    return user_id in ADMIN_IDS


def is_primary(user_id: Optional[int]) -> bool:
    return user_id == PRIMARY_ADMIN_ID


def is_stock_admin(user_id: Optional[int]) -> bool:
    """Admins allowed to use /vpn, /account, /broadcast (primary admin is limited to /add, /on, /off)."""
    return is_admin(user_id) and not is_primary(user_id)


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
            [kbtn(BTN_SUP, "primary")],
        ],
        resize_keyboard=True,
    )


def cancel_flow_kb() -> InlineKeyboardMarkup:
    return ikb([ibtn("❌ 𝗖𝗮𝗻𝗰𝗲𝗹", "flowcancel", style="danger")])


# --------------------------------------------------------------------------- #
# Database (Firebase Realtime Database)
# --------------------------------------------------------------------------- #
# Layout:
#   users/{user_id}, products/{code}, vpns/{id}, orders/{id}, deposits/{id}
#   settings/bot_enabled, counters/{name}
#   stock/{code}/{account_id}        -> available accounts
#   stock_count/{code}               -> number of available accounts
#   account_keys/{code}/{sha1}       -> duplicate guard (code + email)
#   order_accounts/{order_id}/{id}   -> accounts delivered for an order
#   pending_orders/{order_id}, undelivered/{user_id}/{order_id}
#   referrals/{referrer_id}/{user_id}, deposit_refs/{method}/{sha1}
#   deposit_msgs/{deposit_id}/{push_key}
_lock = threading.RLock()  # serialises read-modify-write sections (all DB calls are synchronous)


class PurchaseError(Exception):
    pass


def _init_firebase() -> None:
    if not FIREBASE_CREDENTIALS:
        raise SystemExit("FIREBASE_CREDENTIALS environment variable is not set.")
    if not DATABASE_URL:
        raise SystemExit("DATABASE_URL environment variable is not set.")
    if FIREBASE_CREDENTIALS.startswith("{"):
        info = json.loads(FIREBASE_CREDENTIALS, strict=False)
        if isinstance(info.get("private_key"), str):
            info["private_key"] = info["private_key"].replace("\\n", "\n")
        cred = credentials.Certificate(info)
    else:  # path to a service-account JSON file
        cred = credentials.Certificate(FIREBASE_CREDENTIALS)
    firebase_admin.initialize_app(cred, {"databaseURL": DATABASE_URL})


def _ref(path: str):
    return rtdb.reference(path)


def _update(updates: dict) -> None:
    """Atomic multi-path write (a None value deletes that path)."""
    rtdb.reference("/").update(updates)


def _items(val) -> list:
    """Firebase returns integer-keyed nodes as lists; normalise to (key, value) pairs."""
    if not val:
        return []
    if isinstance(val, list):
        return [(str(i), v) for i, v in enumerate(val) if v is not None]
    return [(str(k), v) for k, v in val.items()]


def _keys(path: str) -> list:
    ref = _ref(path)
    try:
        val = ref.get(shallow=True)
    except TypeError:  # very old firebase-admin without shallow reads
        val = ref.get()
    return [k for k, _ in _items(val)]


def _next_id(name: str, count: int = 1) -> int:
    """Atomically reserve `count` ids. Returns the last reserved id (ids start at 1)."""
    return _ref(f"counters/{name}").transaction(lambda cur: (cur or 0) + count)


def _hash(text) -> str:
    return hashlib.sha1(str(text).encode("utf-8")).hexdigest()


def _valid_code(code) -> bool:
    return isinstance(code, str) and re.fullmatch(r"[A-Za-z0-9_]{1,64}", code) is not None


def _clean(d: dict) -> dict:
    return {k: v for k, v in d.items() if v is not None}


def init_db() -> None:
    _init_firebase()
    for code, meta in PRODUCTS.items():
        if _ref(f"products/{code}").get():
            _ref(f"products/{code}/name").set(meta["name"])
        else:
            _ref(f"products/{code}").set(
                {"code": code, "name": meta["name"], "price": meta["price"], "active": 1}
            )
    if _ref("settings/bot_enabled").get() is None:
        _ref("settings/bot_enabled").set("1")


# ---- settings ------------------------------------------------------------- #
def is_enabled() -> bool:
    val = _ref("settings/bot_enabled").get()
    return val is None or str(val) == "1"


def set_enabled(flag: bool) -> None:
    _ref("settings/bot_enabled").set("1" if flag else "0")


# ---- users ---------------------------------------------------------------- #
def _user_row(user_id: int, d: Optional[dict]) -> Optional[dict]:
    if not d:
        return None
    return {
        "user_id": user_id, "username": None, "full_name": "", "balance": 0,
        "started": 0, "created_at": "", "referred_by": None, **d,
    }


def upsert_user(u: User) -> bool:
    """Insert/update a user. Returns True if this is the user's first time starting the bot."""
    with _lock:
        prev = _ref(f"users/{u.id}").get()
        is_new = not prev or not prev.get("started")
        if prev:
            _update({
                f"users/{u.id}/username": u.username,
                f"users/{u.id}/full_name": u.full_name,
                f"users/{u.id}/started": 1,
            })
        else:
            _ref(f"users/{u.id}").set(_clean({
                "user_id": u.id, "username": u.username, "full_name": u.full_name,
                "balance": 0, "started": 1, "created_at": now_iso(),
            }))
        return is_new


def get_user(user_id: int):
    return _user_row(user_id, _ref(f"users/{user_id}").get())


def get_balance(user_id: int) -> int:
    row = get_user(user_id)
    return row["balance"] if row else 0


def add_balance(user_id: int, amount: int) -> int:
    with _lock:
        user = _ref(f"users/{user_id}").get()
        if user:
            new_balance = int(user.get("balance", 0)) + amount
            _update({f"users/{user_id}/balance": new_balance})
        else:
            new_balance = amount
            _ref(f"users/{user_id}").set(
                {"user_id": user_id, "full_name": "", "balance": new_balance,
                 "started": 0, "created_at": now_iso()}
            )
        return new_balance


def get_started_user_ids() -> list:
    return [int(k) for k, v in _items(_ref("users").get()) if v.get("started")]


# ---- referrals ------------------------------------------------------------ #
def register_referral(new_user_id: int, referrer_id: int) -> Optional[int]:
    """Credit the referrer once for a brand-new user. Returns referrer's new balance or None."""
    if new_user_id == referrer_id:
        return None
    with _lock:
        ref = _ref(f"users/{referrer_id}").get()
        if not ref or not ref.get("started"):
            return None
        new = _ref(f"users/{new_user_id}").get()
        if not new or new.get("referred_by") is not None:
            return None
        new_balance = int(ref.get("balance", 0)) + REFERRAL_BONUS
        _update({
            f"users/{new_user_id}/referred_by": referrer_id,
            f"users/{referrer_id}/balance": new_balance,
            f"referrals/{referrer_id}/{new_user_id}": True,
        })
        return new_balance


def count_referrals(user_id: int) -> int:
    return len(_keys(f"referrals/{user_id}"))


# ---- products / stock ----------------------------------------------------- #
def get_product(code: str):
    if not _valid_code(code):
        return None
    d = _ref(f"products/{code}").get()
    if not d:
        return None
    return {"code": code, "name": "", "price": 0, "active": 1, **d}


VPN_PREFIX = "vpn_"


def _vpn_row(vpn_id, d: dict) -> dict:
    return {"name": "", "days": 0, "price": 0, **d, "id": int(vpn_id)}


def list_vpns() -> list:
    rows = [_vpn_row(k, v) for k, v in _items(_ref("vpns").get())]
    return sorted(rows, key=lambda r: r["id"])


def get_vpn_by_name(name: str):
    wanted = name.lower()
    for row in list_vpns():
        if row["name"].lower() == wanted:
            return row
    return None


def get_vpn_by_code(code: str):
    if not code.startswith(VPN_PREFIX) or not code[len(VPN_PREFIX):].isdigit():
        return None
    vpn_id = int(code[len(VPN_PREFIX):])
    d = _ref(f"vpns/{vpn_id}").get()
    return _vpn_row(vpn_id, d) if d else None


def create_vpn(name: str, days: int, price: int) -> str:
    """Create a VPN (also registered as a product so the normal buy flow works). Returns its code."""
    with _lock:
        if get_vpn_by_name(name):
            raise ValueError("VPN name already exists")
        vpn_id = _next_id("vpns")
        code = f"{VPN_PREFIX}{vpn_id}"
        _update({
            f"vpns/{vpn_id}": {"id": vpn_id, "name": name, "days": days, "price": price},
            f"products/{code}": {"code": code, "name": name, "price": price, "active": 1},
        })
        return code


def product_meta(code: str) -> Optional[dict]:
    meta = PRODUCTS.get(code)
    if meta:
        return meta
    vpn = get_vpn_by_code(code)
    if vpn:
        return {"name": f"{vpn['name']} ({vpn['days']} Days)", "title": f"{vpn['name']} ({vpn['days']} Days)"}
    return None


def set_product_price(code: str, price: int) -> None:
    if get_product(code) is not None:
        _ref(f"products/{code}/price").set(price)


def _stock_count(code: str) -> int:
    """Available accounts for a product (kept in stock_count; rebuilt from stock if missing)."""
    val = _ref(f"stock_count/{code}").get()
    if val is None:
        val = len(_keys(f"stock/{code}"))
        _ref(f"stock_count/{code}").set(val)
    return int(val)


def get_stock(code: str) -> int:
    if not _valid_code(code):
        return 0
    return _stock_count(code)


def import_accounts(code: str, rows: list) -> int:
    with _lock:
        known = set(_keys(f"account_keys/{code}"))
        fresh = []
        for email, password in rows:
            h = _hash(email)
            if h in known:
                continue
            known.add(h)
            fresh.append((email, password, h))
        if not fresh:
            return 0
        count = _stock_count(code)
        first_id = _next_id("accounts", len(fresh)) - len(fresh) + 1
        for start in range(0, len(fresh), 500):
            chunk = fresh[start:start + 500]
            updates = {}
            for offset, (email, password, h) in enumerate(chunk, start):
                account_id = first_id + offset
                updates[f"stock/{code}/{account_id}"] = {"email": email, "password": password}
                updates[f"account_keys/{code}/{h}"] = account_id
            updates[f"stock_count/{code}"] = count + start + len(chunk)
            _update(updates)
        return len(fresh)


def get_order_accounts(order_id: int) -> list:
    items = sorted(_items(_ref(f"order_accounts/{order_id}").get()), key=lambda kv: int(kv[0]))
    return [(v["email"], v["password"]) for _, v in items]


# ---- orders --------------------------------------------------------------- #
def expire_pending() -> list:
    """Cancel timed-out pending orders. Returns the still-pending (order_id, info) pairs."""
    with _lock:
        cutoff = now_ts() - PENDING_TTL
        updates, remaining = {}, []
        for k, v in _items(_ref("pending_orders").get()):
            if v.get("created_ts", 0) < cutoff:
                updates[f"orders/{k}/status"] = "CANCELLED"
                updates[f"pending_orders/{k}"] = None
            else:
                remaining.append((k, v))
        if updates:
            _update(updates)
        return remaining


def get_pending_order(user_id: int):
    mine = [int(k) for k, v in expire_pending() if v.get("user_id") == user_id]
    return get_order(max(mine)) if mine else None


def get_order(order_id: int):
    d = _ref(f"orders/{order_id}").get()
    return {"id": order_id, **d} if d else None


def create_pending_order(user_id: int, code: str, quantity: int) -> int:
    product = get_product(code)
    unit = product["price"]
    with _lock:
        order_id = _next_id("orders")
        ts = now_ts()
        _update({
            f"orders/{order_id}": {
                "id": order_id, "user_id": user_id, "product_code": code, "quantity": quantity,
                "unit_price": unit, "total_price": unit * quantity, "status": "PENDING",
                "delivered": 0, "created_at": now_iso(), "created_ts": ts,
            },
            f"pending_orders/{order_id}": {"user_id": user_id, "created_ts": ts},
        })
        return order_id


def cancel_order(order_id: int, user_id: int) -> bool:
    with _lock:
        order = get_order(order_id)
        if order is None or order["user_id"] != user_id or order["status"] != "PENDING":
            return False
        _update({f"orders/{order_id}/status": "CANCELLED", f"pending_orders/{order_id}": None})
        return True


def get_undelivered_orders(user_id: int) -> list:
    ids = sorted(int(k) for k in _keys(f"undelivered/{user_id}"))
    return [{"id": i} for i in ids]


def mark_delivered(order_id: int) -> None:
    order = get_order(order_id)
    if order is not None:
        _update({
            f"orders/{order_id}/delivered": 1,
            f"undelivered/{order['user_id']}/{order_id}": None,
        })


def do_purchase(order_id: int, user_id: int) -> dict:
    with _lock:
        order = get_order(order_id)
        if order is None or order["user_id"] != user_id or order["status"] != "PENDING":
            return {"status": "invalid"}
        cancel = {f"orders/{order_id}/status": "CANCELLED", f"pending_orders/{order_id}": None}
        if order["created_ts"] < now_ts() - PENDING_TTL:
            _update(cancel)
            return {"status": "expired"}
        product = get_product(order["product_code"])
        if product is None or not product["active"]:
            _update(cancel)
            return {"status": "inactive"}
        total = order["total_price"]
        quantity = order["quantity"]
        code = order["product_code"]
        user = _ref(f"users/{user_id}").get()
        balance = int(user.get("balance", 0)) if user else 0
        if balance < total:
            return {"status": "balance", "need": total, "balance": balance}
        stock_count = _stock_count(code)
        picked = sorted(
            _items(_ref(f"stock/{code}").order_by_key().limit_to_first(quantity).get()),
            key=lambda kv: int(kv[0]),
        )
        if len(picked) < quantity:
            _update(cancel)
            return {"status": "stock", "stock": len(picked)}
        updates = {
            f"users/{user_id}/balance": balance - total,
            f"orders/{order_id}/status": "COMPLETED",
            f"pending_orders/{order_id}": None,
            f"undelivered/{user_id}/{order_id}": True,
            f"stock_count/{code}": stock_count - quantity,
        }
        for account_id, account in picked:
            updates[f"stock/{code}/{account_id}"] = None
            updates[f"order_accounts/{order_id}/{account_id}"] = {
                "email": account["email"], "password": account["password"],
            }
        _update(updates)  # one atomic write: balance, order, stock and accounts together
        return {"status": "ok"}


# ---- deposits ------------------------------------------------------------- #
def _deposit_ref_path(method: str, reference: str) -> str:
    return f"deposit_refs/{method}/{_hash(reference)}"


def reference_exists(method: str, reference: str) -> bool:
    return _ref(_deposit_ref_path(method, reference)).get() is not None


def create_deposit(user_id: int, method: str, amount: int, reference: str) -> int:
    with _lock:
        dep_id = _next_id("deposits")
        _update({
            f"deposits/{dep_id}": {
                "id": dep_id, "user_id": user_id, "method": method, "amount": amount,
                "reference": reference, "status": "PENDING", "created_at": now_iso(),
            },
            _deposit_ref_path(method, reference): dep_id,
        })
        return dep_id


def get_deposit(dep_id: int):
    d = _ref(f"deposits/{dep_id}").get()
    if not d:
        return None
    return {"processed_by": None, "processed_at": None, **d, "id": dep_id}


def finalize_deposit(dep_id: int, admin_id: int, approve: bool) -> Optional[dict]:
    """Atomically approve/cancel a pending deposit exactly once. Returns None if already processed."""
    with _lock:
        dep = get_deposit(dep_id)
        if dep is None or dep["status"] != "PENDING":
            return None
        updates = {
            f"deposits/{dep_id}/status": "APPROVED" if approve else "CANCELLED",
            f"deposits/{dep_id}/processed_by": admin_id,
            f"deposits/{dep_id}/processed_at": now_iso(),
        }
        new_balance = None
        if approve:
            user = _ref(f"users/{dep['user_id']}").get()
            if user:
                new_balance = int(user.get("balance", 0)) + dep["amount"]
                updates[f"users/{dep['user_id']}/balance"] = new_balance
            else:
                new_balance = dep["amount"]
                updates[f"users/{dep['user_id']}"] = {
                    "user_id": dep["user_id"], "full_name": "", "balance": new_balance,
                    "started": 0, "created_at": now_iso(),
                }
        else:
            updates[_deposit_ref_path(dep["method"], dep["reference"])] = None  # reference can be reused
        _update(updates)
        return {"user_id": dep["user_id"], "amount": dep["amount"], "balance": new_balance}


def save_deposit_msg(dep_id: int, chat_id: int, message_id: int) -> None:
    _ref(f"deposit_msgs/{dep_id}").push({"chat_id": chat_id, "message_id": message_id})


def get_deposit_msgs(dep_id: int) -> list:
    return [v for _, v in _items(_ref(f"deposit_msgs/{dep_id}").get())]


# --------------------------------------------------------------------------- #
# Text builders
# --------------------------------------------------------------------------- #
def order_summary_text(order) -> str:
    meta = product_meta(order["product_code"]) or {"title": order["product_code"]}
    return (
        f"🛒 <b>{esc(meta['title'])}</b>\n\n"
        f"💰 প্রাইস: {fmt_money(order['unit_price'])} TK\n"
        f"স্টক: {get_stock(order['product_code'])}\n\n"
        f"পরিমাণ: {order['quantity']}\n"
        f"মোট খরচ: {fmt_money(order['total_price'])} TK"
    )


def order_kb(order_id: int) -> InlineKeyboardMarkup:
    return ikb([
        ibtn("✅ 𝗖𝗼𝗻𝗳𝗶𝗿𝗺 𝗢𝗿𝗱𝗲𝗿", f"cfm:{order_id}", style="success"),
        ibtn("❌ 𝗖𝗮𝗻𝗰𝗲𝗹", f"cnl:{order_id}", style="danger"),
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
    rows.append([ibtn("VPN 🛡️", "vpnlist", style="primary")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def vpn_label(vpn) -> str:
    code = f"{VPN_PREFIX}{vpn['id']}"
    return f"🛡️ {vpn['name']} | ⌛ {vpn['days']} Days | {price_short(vpn['price'])} BDT | 📦 {get_stock(code)}"


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


class AdminPrice(StatesGroup):
    amount = State()


class AdminVpn(StatesGroup):
    name = State()
    days = State()
    price = State()
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
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(WELCOME_TEXT, reply_markup=main_menu())


@router.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("❌ ক্যানসেল করা হয়েছে।", reply_markup=main_menu())


@router.message(Command("retry"))
async def cmd_retry(message: Message, bot: Bot):
    rows = get_undelivered_orders(message.from_user.id)
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
    if not is_stock_admin(message.from_user.id):
        await message.answer("⛔ আপনি অ্যাডমিন নন।")
        return
    await state.clear()
    rows = [[ibtn(meta["admin_label"], f"acc:{code}", style="primary")] for code, meta in PRODUCTS.items()]
    await message.answer("📦 Select Products", reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))


@router.message(Command("broadcast"))
async def cmd_broadcast(message: Message, state: FSMContext):
    if not is_stock_admin(message.from_user.id):
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
        "🛍️ <b>NEXA STORE</b>\n\nআপনার প্রোডাক্টটি বেছে নিন:", reply_markup=products_kb()
    )


@router.message(F.text == BTN_BAL)
async def menu_balance(message: Message, state: FSMContext):
    await state.clear()
    user = get_user(message.from_user.id)
    username = f"@{esc(user['username'])}" if user and user["username"] else "Not set"
    full_name = esc(message.from_user.full_name)
    balance = fmt_money(user["balance"] if user else 0)
    await message.answer(
        "👛 <b>My Wallet</b>\n\n"
        f"🔖 Username: {username}\n"
        f"🙋 নাম: {NAME_PREFIX}{full_name}\n"
        f"💎 ব্যালেন্স: {balance} TK"
    )


@router.message(F.text == BTN_DEP)
async def menu_deposit(message: Message, state: FSMContext):
    await state.clear()
    kb = ikb(
        [ibtn("🌸 𝗯𝗞𝗮𝘀𝗵", "pay:bkash", style="danger")],
        [ibtn("🟠 𝗡𝗮𝗴𝗮𝗱", "pay:nagad", style="primary")],
        [ibtn("🟡 𝗕𝗶𝗻𝗮𝗻𝗰𝗲", "pay:binance", style="success")],
    )
    await message.answer(
        "➕ <b>Add Funds</b>\n\nপেমেন্ট মেথড বেছে নিন:\n\n"
        "ℹ️ bKash, Nagad ও Binance Pay সব মেথড ২৪/৭ চালু থাকে।",
        reply_markup=kb,
    )


@router.message(F.text == BTN_SUP)
async def menu_support(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "💬 <b>Help Desk</b>\n\n"
        "কোনো সমস্যা বা প্রশ্ন থাকলে নিচের বাটনে ট্যাপ করে আমাদের টিমকে জানান:",
        reply_markup=ikb([ibtn("🎧 𝗖𝗼𝗻𝘁𝗮𝗰𝘁 𝗧𝗲𝗮𝗺", url=SUPPORT_URL, style="primary")]),
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
    meta = product_meta(code)
    if meta is None:
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
        f"🛒 <b>{esc(meta['name'])}</b>\n\n"
        f"💰 প্রাইস: {fmt_money(product['price'])} TK\n"
        f"স্টক: {stock}\n\n"
        "পরিমাণ লিখুন:",
        reply_markup=cancel_flow_kb(),
    )


@router.message(BuyFlow.quantity, F.text, not_command)
async def got_quantity(message: Message, state: FSMContext):
    data = await state.get_data()
    code = data.get("code")
    if not code or product_meta(code) is None or state_expired(data, PENDING_TTL):
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


@router.callback_query(F.data == "vpnlist")
async def cb_vpn_list(call: CallbackQuery, bot: Bot):
    vpns = list_vpns()
    if not vpns:
        await call.answer("❌ এখন কোনো VPN নেই।", show_alert=True)
        return
    await call.answer()
    rows = [[ibtn(vpn_label(v), f"buy:{VPN_PREFIX}{v['id']}", style="primary")] for v in vpns]
    await safe_send(
        bot,
        call.from_user.id,
        "🛡️ <b>VPN Store</b>\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "আপনার VPN বেছে নিন:\n"
        "📦 = Stock",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )


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
    meta = product_meta(order["product_code"]) or {"name": order["product_code"]}
    caption = (
        "🎉 <b>ডেলিভারি সম্পন্ন!</b>\n\n"
        f"🧾 Order No: <code>#{order_id}</code>\n"
        f"🏷️ প্রোডাক্ট: {esc(meta['name'])}\n"
        f"🔢 পরিমাণ: {order['quantity']}\n"
        f"💎 মোট খরচ: {fmt_money(order['total_price'])} TK"
    )
    payload = build_xlsx(accounts)
    filename = f"NexaStore_{order_id}_{order['product_code']}.xlsx"
    for _ in range(3):
        try:
            await bot.send_document(
                order["user_id"], BufferedInputFile(payload, filename=filename), caption=caption
            )
            mark_delivered(order_id)
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
            "🧾 <b>পেমেন্ট ইনভয়েস তৈরি হয়েছে</b>\n\n"
            f"💵 পরিমাণ: <b>{amt}</b> টাকা (ঠিক এই অ্যামাউন্টই পাঠাবেন)\n\n"
            "💚 <b>BKASH</b>\n\n"
            "📱 Personal (Send Money):\n"
            f"<code>{BKASH_NUMBER}</code>\n\n"
            "👨‍💻 পেমেন্ট করার পর আপনার TrxID এখানে পাঠান 📍\n\n"
            "⚠️ আগে পেমেন্ট করুন, তারপর সঠিক TrxID পাঠান। অ্যাডমিন ম্যানুয়ালি যাচাই করে ব্যালেন্স যোগ করবেন।"
        )
    if method == "nagad":
        return (
            "🧾 <b>পেমেন্ট ইনভয়েস তৈরি হয়েছে</b>\n\n"
            f"💵 পরিমাণ: <b>{amt}</b> টাকা (ঠিক এই অ্যামাউন্টই পাঠাবেন)\n\n"
            "🧡 <b>NAGAD</b>\n\n"
            "📱 Personal (Send Money):\n"
            f"<code>{NAGAD_NUMBER}</code>\n\n"
            "👨‍💻 পেমেন্ট করার পর আপনার TrxID এখানে পাঠান 📍\n\n"
            "⚠️ আগে পেমেন্ট করুন, তারপর সঠিক TrxID পাঠান। অ্যাডমিন ম্যানুয়ালি যাচাই করে ব্যালেন্স যোগ করবেন।"
        )
    return (
        "🧾 <b>পেমেন্ট ইনভয়েস তৈরি হয়েছে</b>\n\n"
        f"💵 পরিমাণ: <b>{amt}</b> টাকা (ঠিক এই অ্যামাউন্টই পাঠাবেন)\n\n"
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
    if not is_stock_admin(call.from_user.id):
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
        has_header = len(first) >= 2 and _norm(first[0]) in ("email", "username") and _norm(first[1]) == "password"
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
# Admin: /price
# --------------------------------------------------------------------------- #
@router.message(Command("price"))
async def cmd_price(message: Message, state: FSMContext):
    if not is_stock_admin(message.from_user.id):
        await message.answer("⛔ এই কমান্ড আপনার জন্য নয়।")
        return
    await state.clear()
    rows = [[ibtn(meta["admin_label"], f"prc:{code}", style="primary")] for code, meta in PRODUCTS.items()]
    await message.answer("💰 Select a product to set price:", reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))


@router.callback_query(F.data.regexp(r"^prc:(meta_ai|meta_horizon)$"))
async def cb_price_select(call: CallbackQuery, state: FSMContext, bot: Bot):
    if not is_stock_admin(call.from_user.id):
        await call.answer("⛔ আপনি অ্যাডমিন নন।", show_alert=True)
        return
    code = call.data.split(":", 1)[1]
    product = get_product(code)
    await call.answer()
    await state.clear()
    await state.set_state(AdminPrice.amount)
    await state.update_data(code=code)
    await safe_send(
        bot,
        call.from_user.id,
        f"💰 <b>{esc(PRODUCTS[code]['name'])}</b>\n"
        f"বর্তমান প্রাইস: {fmt_money(product['price'])} TK\n\n"
        "নতুন প্রাইস দিন (BDT, যেমন: 50 বা 12.50):",
        reply_markup=cancel_flow_kb(),
    )


@router.message(AdminPrice.amount, F.text, not_command)
async def price_get_amount(message: Message, state: FSMContext):
    if not is_stock_admin(message.from_user.id):
        await state.clear()
        return
    data = await state.get_data()
    code = data.get("code")
    price = parse_amount(message.text)
    if price is None:
        await message.answer("❌ সঠিক প্রাইস লিখুন (ধনাত্মক সংখ্যা, সর্বোচ্চ ২ দশমিক ঘর)।")
        return
    await state.clear()
    if code not in PRODUCTS:
        await message.answer("❌ কিছু ভুল হয়েছে। /price দিয়ে আবার শুরু করুন।")
        return
    set_product_price(code, price)
    await message.answer(
        f"✅ <b>{esc(PRODUCTS[code]['name'])}</b> এর প্রাইস সেট হয়েছে: {fmt_money(price)} TK"
    )


# --------------------------------------------------------------------------- #
# Admin: /vpn
# --------------------------------------------------------------------------- #
VPN_FILE_PROMPT = (
    "📤 এখন XLSX ফাইল পাঠান।\n\n"
    "<b>Format:</b>\n"
    "A কলাম: Email অথবা Username\n"
    "B কলাম: Password\n"
    "<i>(হেডার থাকা ঐচ্ছিক)</i>"
)


@router.message(Command("vpn"))
async def cmd_vpn(message: Message, state: FSMContext):
    if not is_stock_admin(message.from_user.id):
        await message.answer("⛔ এই কমান্ড আপনার জন্য নয়।")
        return
    await state.clear()
    await state.set_state(AdminVpn.name)
    await message.answer("🛡️ VPN এর নাম দিন:", reply_markup=cancel_flow_kb())


@router.message(AdminVpn.name, F.text, not_command)
async def vpn_get_name(message: Message, state: FSMContext):
    if not is_stock_admin(message.from_user.id):
        await state.clear()
        return
    name = " ".join((message.text or "").split())
    if not name or len(name) > 40:
        await message.answer("❌ সঠিক নাম দিন (সর্বোচ্চ ৪০ অক্ষর)।")
        return
    existing = get_vpn_by_name(name)
    if existing:
        await state.update_data(code=f"{VPN_PREFIX}{existing['id']}", ts=time.time())
        await state.set_state(AdminVpn.file)
        await message.answer(
            f"♻️ <b>{esc(existing['name'])}</b> আগে থেকেই আছে। এর স্টকে নতুন অ্যাকাউন্ট যোগ হবে।\n\n"
            + VPN_FILE_PROMPT,
            reply_markup=cancel_flow_kb(),
        )
        return
    await state.update_data(name=name)
    await state.set_state(AdminVpn.days)
    await message.answer("⌛ এই VPN কত দিনের জন্য থাকবে? (দিন লিখুন, যেমন: 30)", reply_markup=cancel_flow_kb())


@router.message(AdminVpn.days, F.text, not_command)
async def vpn_get_days(message: Message, state: FSMContext):
    if not is_stock_admin(message.from_user.id):
        await state.clear()
        return
    days = parse_positive_int(message.text, 4)
    if days is None:
        await message.answer("❌ সঠিক দিন লিখুন (শুধু সংখ্যা, যেমন: 30)।")
        return
    await state.update_data(days=days)
    await state.set_state(AdminVpn.price)
    await message.answer("💰 এই VPN এর প্রাইস কত? (BDT, যেমন: 50)", reply_markup=cancel_flow_kb())


@router.message(AdminVpn.price, F.text, not_command)
async def vpn_get_price(message: Message, state: FSMContext):
    if not is_stock_admin(message.from_user.id):
        await state.clear()
        return
    price = parse_amount(message.text)
    if price is None:
        await message.answer("❌ সঠিক প্রাইস লিখুন (ধনাত্মক সংখ্যা, সর্বোচ্চ ২ দশমিক ঘর)।")
        return
    await state.update_data(price=price)
    await state.set_state(AdminVpn.file)
    await message.answer(VPN_FILE_PROMPT, reply_markup=cancel_flow_kb())


@router.message(AdminVpn.file, F.document)
async def vpn_upload_file(message: Message, state: FSMContext, bot: Bot):
    if not is_stock_admin(message.from_user.id):
        await state.clear()
        return
    data = await state.get_data()
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
            await message.answer("❌ ফাইলটি খালি। কলাম A তে Email/Username এবং কলাম B তে Password দিন।")
        else:
            await message.answer(f"❌ সর্বোচ্চ {MAX_XLSX_ROWS} সারি অনুমোদিত।")
        return
    if not parsed["valid"]:
        await message.answer("❌ কোনো বৈধ সারি পাওয়া যায়নি। ফাইল ঠিক করে আবার পাঠান।")
        return

    # Create the VPN only now, so a cancelled flow never leaves an empty VPN in the shop.
    code = data.get("code")
    if code:
        vpn = get_vpn_by_code(code)
    else:
        name, days, price = data.get("name"), data.get("days"), data.get("price")
        if not (name and isinstance(days, int) and isinstance(price, int)):
            await state.clear()
            await message.answer("❌ কিছু ভুল হয়েছে। /vpn দিয়ে আবার শুরু করুন।")
            return
        vpn = get_vpn_by_name(name)
        if vpn is None:
            code = create_vpn(name, days, price)
            vpn = get_vpn_by_code(code)
    if vpn is None:
        await state.clear()
        await message.answer("❌ VPN পাওয়া যায়নি। /vpn দিয়ে আবার শুরু করুন।")
        return
    code = f"{VPN_PREFIX}{vpn['id']}"

    imported = await asyncio.to_thread(import_accounts, code, parsed["valid"])
    await state.clear()
    dup_total = parsed["dup_file"] + (len(parsed["valid"]) - imported)
    await message.answer(
        f"📊 <b>VPN স্টক আপডেট — {esc(vpn['name'])}</b>\n\n"
        f"⌛ মেয়াদ: {vpn['days']} Days\n"
        f"💰 প্রাইস: {fmt_money(vpn['price'])} TK\n\n"
        f"📄 মোট সারি প্রসেস: {parsed['total']}\n"
        f"✅ সফলভাবে ইমপোর্ট: {imported}\n"
        f"♻️ ডুপ্লিকেট স্কিপ: {dup_total}\n"
        f"⚠️ ইনভ্যালিড স্কিপ: {len(parsed['invalid'])}\n"
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
    user_ids = get_started_user_ids()
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
    log.info("Database: Firebase | Admins: %s", sorted(ADMIN_IDS))
    if SECOND_ADMIN_ID is None:
        log.warning("ADMIN_ID is not set - only the primary admin is active.")

    probe = InlineKeyboardButton(text="x", callback_data="x", style="success")
    if "style" not in probe.model_dump(exclude_none=True):
        log.warning("This aiogram version drops the button 'style' field - upgrade aiogram.")

    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    await bot.delete_webhook(drop_pending_updates=False)
    try:  # show "NEXA STORE" as the bot's display name in Telegram
        if (await bot.get_my_name()).name != "NEXA STORE":
            await bot.set_my_name("NEXA STORE")
    except TelegramAPIError as e:
        log.warning("Could not set bot name (set it in @BotFather): %s", type(e).__name__)
    log.info("NEXA STORE started (polling).")
    await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())


if __name__ == "__main__":
    asyncio.run(main())
