import os
import io
import re
import asyncio
import logging
import sys
import sqlite3
import time
import traceback
import datetime
from typing import Callable, Dict, Any, Awaitable
from itertools import permutations, zip_longest
from zoneinfo import ZoneInfo
import aiohttp
from aiohttp import web
from aiogram import Bot, Dispatcher, html, types
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramForbiddenError, TelegramBadRequest
from aiogram.filters import CommandStart, Command, CommandObject
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    InlineKeyboardMarkup, InlineKeyboardButton, ErrorEvent, BufferedInputFile,
    InlineQueryResultArticle, InputTextMessageContent, WebAppInfo,
    KeyboardButton, ReplyKeyboardMarkup,
)
from aiogram import BaseMiddleware
from aiogram.types import TelegramObject, Message, CallbackQuery

# 📊 Excel va PDF eksport uchun kutubxonalar (requirements.txt'ga qo'shilgan)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from fpdf import FPDF

# 🤖 AI-yordamchi uchun Anthropic (Claude) SDK
from anthropic import AsyncAnthropic


TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_ID = 7522996937  # ⚠️ Bu yerga o'zingizning Telegram ID raqamingizni yozing
dp = Dispatcher(storage=MemoryStorage())

# 🤖 Bot obyektiga global murojaat qilish uchun (masalan global xato handler ichidan).
# main() ichida to'ldiriladi.
BOT_INSTANCE: Bot | None = None

# "Noma'lum" so'zi (f-string ichida teskari slesh ishlatmaslik uchun alohida o'zgaruvchi —
# Python 3.12 dan eski versiyalarda f-string ichidagi "\" SyntaxError beradi)
UNKNOWN_NAME = "Noma\u02bclum"


def esc(text: Any) -> str:
    """Matnni Telegram HTML rejimi uchun xavfsiz qiladi (<, >, & belgilarini almashtiradi)."""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# 🕓 Server (Render) UTC vaqt zonasida ishlaydi, lekin universitet Toshkentda
# (UTC+5). Shuning uchun "hozirgi vaqt/sana" har doim shu funksiya orqali
# olinadi — now_tashkent() yoki now_tashkent().date() TO'G'RIDAN-TO'G'RI
# ishlatilmasin, aks holda vaqt 5 soatga farq qiladi.
TASHKENT_TZ = ZoneInfo("Asia/Tashkent")


def now_tashkent() -> datetime.datetime:
    return datetime.datetime.now(TASHKENT_TZ)


def fmt_dt(iso_text: str | None) -> str:
    """Bazadagi ISO vaqtni "2026-09-23 14:50" ko'rinishiga o'giradi."""
    if not iso_text:
        return "—"
    try:
        return datetime.datetime.fromisoformat(iso_text).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return str(iso_text)[:16]


# ============================================================
# 👥 FOYDALANUVCHILAR BAZASI (SQLite)
# Har bir botdan foydalanuvchi (ID, username, ism, birinchi/oxirgi marta
# yozgan vaqti, nechta harakat qilgani, bloklangan-bloklanmaganligi) shu
# yerda saqlanadi. SQLite fayli — qo'shimcha xizmat sozlash shart emas.
# ⚠️ Eslatma: Render'ning bepul tarifida disk vaqti-vaqti bilan (masalan
# yangi deploy qilinganda) tozalanishi mumkin — bazani doimiy saqlash uchun
# Render'ning "Persistent Disk" xizmatidan foydalaning.
# ============================================================

DB_PATH = os.environ.get("DB_PATH", "bot_users.db")
USERS_PAGE_SIZE = 8  # admin panelida bir sahifada nechta foydalanuvchi ko'rsatiladi


def _init_db() -> None:
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            full_name TEXT,
            first_seen TEXT,
            last_seen TEXT,
            message_count INTEGER DEFAULT 0,
            is_banned INTEGER DEFAULT 0
        )
        """
    )
    # ⭐ SEVIMLILAR: foydalanuvchi tez-tez qidiradigan xona (blok+qavat) yoki
    # talaba/o'qituvchi/guruhni "sevimli" sifatida saqlab qo'yishi mumkin.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS favorites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            fav_type TEXT NOT NULL,   -- 'room' yoki 'person'
            fav_key TEXT NOT NULL,    -- 'room': "BLOCK:FLOOR", 'person': "code:entity_id"
            label TEXT NOT NULL,
            created_at TEXT,
            UNIQUE(user_id, fav_type, fav_key)
        )
        """
    )
    # 🔔 OBUNALAR: bitta foydalanuvchi bitta talaba/o'qituvchi/guruhning
    # kunlik (soat 7:00) jadvaliga va/yoki "10 daqiqa qoldi" eslatmasiga
    # obuna bo'lishi mumkin. Har bir foydalanuvchida faqat bitta faol obuna.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS subscriptions (
            user_id INTEGER PRIMARY KEY,
            entity_type TEXT NOT NULL,   -- 's' (talaba), 't' (o'qituvchi), 'g' (guruh)
            entity_id TEXT NOT NULL,
            entity_label TEXT NOT NULL,
            daily_enabled INTEGER DEFAULT 1,
            reminder_enabled INTEGER DEFAULT 1,
            created_at TEXT
        )
        """
    )
    # 🤖 AI FOYDALANISH KVOTASI: har bir foydalanuvchi kuniga nechta AI
    # so'rovi yuborganini kuzatib boradi (xarajat va suiiste'moldan himoya).
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS ai_usage (
            user_id INTEGER NOT NULL,
            usage_date TEXT NOT NULL,
            count INTEGER DEFAULT 0,
            PRIMARY KEY (user_id, usage_date)
        )
        """
    )
    conn.commit()
    conn.close()


async def db_touch_user(user_id: int, username: str | None, full_name: str) -> None:
    """Foydalanuvchini bazaga qo'shadi yoki oxirgi faollik vaqti va harakatlar sonini yangilaydi."""

    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        now = now_tashkent().isoformat()
        row = conn.execute("SELECT user_id FROM users WHERE user_id=?", (user_id,)).fetchone()
        if row:
            conn.execute(
                "UPDATE users SET username=?, full_name=?, last_seen=?, message_count=message_count+1 WHERE user_id=?",
                (username, full_name, now, user_id),
            )
        else:
            conn.execute(
                "INSERT INTO users (user_id, username, full_name, first_seen, last_seen, message_count, is_banned) "
                "VALUES (?, ?, ?, ?, ?, 1, 0)",
                (user_id, username, full_name, now, now),
            )
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_ensure_user(user_id: int, username: str | None, full_name: str) -> None:
    """db_touch_user'ning "yengil" varianti: foydalanuvchini ro'yxatga qo'shadi/yangilaydi,
    lekin harakatlar sonini OSHIRMAYDI (inline qidiruvda har harf yozilganda so'rov
    ketgani uchun hisobni sun'iy oshirib yubormaslik uchun)."""

    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        now = now_tashkent().isoformat()
        row = conn.execute("SELECT 1 FROM users WHERE user_id=?", (user_id,)).fetchone()
        if row:
            conn.execute(
                "UPDATE users SET username=?, full_name=?, last_seen=? WHERE user_id=?",
                (username, full_name, now, user_id),
            )
        else:
            conn.execute(
                "INSERT INTO users (user_id, username, full_name, first_seen, last_seen, message_count, is_banned) "
                "VALUES (?, ?, ?, ?, ?, 0, 0)",
                (user_id, username, full_name, now, now),
            )
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_is_banned(user_id: int) -> bool:
    def _work() -> bool:
        conn = sqlite3.connect(DB_PATH)
        row = conn.execute("SELECT is_banned FROM users WHERE user_id=?", (user_id,)).fetchone()
        conn.close()
        return bool(row and row[0])

    return await asyncio.to_thread(_work)


async def db_set_banned(user_id: int, banned: bool) -> None:
    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("UPDATE users SET is_banned=? WHERE user_id=?", (1 if banned else 0, user_id))
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_get_stats() -> tuple[int, int, list[tuple]]:
    def _work():
        conn = sqlite3.connect(DB_PATH)
        total = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        banned = conn.execute("SELECT COUNT(*) FROM users WHERE is_banned=1").fetchone()[0]
        top = conn.execute(
            "SELECT user_id, username, full_name, message_count FROM users "
            "ORDER BY message_count DESC LIMIT 10"
        ).fetchall()
        conn.close()
        return total, banned, top

    return await asyncio.to_thread(_work)


async def db_get_overview() -> dict:
    """Admin panel uchun umumiy ko'rsatkichlar."""

    def _work() -> dict:
        conn = sqlite3.connect(DB_PATH)
        today = now_tashkent().date().isoformat()
        data = {
            "total": conn.execute("SELECT COUNT(*) FROM users").fetchone()[0],
            "banned": conn.execute("SELECT COUNT(*) FROM users WHERE is_banned=1").fetchone()[0],
            "active_today": conn.execute(
                "SELECT COUNT(*) FROM users WHERE substr(last_seen,1,10)=?", (today,)
            ).fetchone()[0],
            "new_today": conn.execute(
                "SELECT COUNT(*) FROM users WHERE substr(first_seen,1,10)=?", (today,)
            ).fetchone()[0],
            "subs": conn.execute("SELECT COUNT(*) FROM subscriptions").fetchone()[0],
            "ai_today": conn.execute(
                "SELECT COALESCE(SUM(count),0) FROM ai_usage WHERE usage_date=?", (today,)
            ).fetchone()[0],
        }
        conn.close()
        return data

    return await asyncio.to_thread(_work)


async def db_list_recent_users(limit: int = 15) -> list[tuple]:
    def _work():
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT user_id, username, full_name, is_banned FROM users ORDER BY last_seen DESC LIMIT ?",
            (limit,),
        ).fetchall()
        conn.close()
        return rows

    return await asyncio.to_thread(_work)


async def db_list_users_page(page: int, page_size: int = USERS_PAGE_SIZE) -> tuple[list[tuple], int, int]:
    """Foydalanuvchilarning bitta sahifasini qaytaradi (eng yangi faol bo'lganlar birinchi).
    Qaytaradi: (qatorlar, jami_soni, tuzatilgan_sahifa_raqami).
    Qator: (user_id, username, full_name, is_banned, message_count, last_seen)"""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        total = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        pages = max((total + page_size - 1) // page_size, 1)
        p = min(max(page, 0), pages - 1)
        rows = conn.execute(
            "SELECT user_id, username, full_name, is_banned, message_count, last_seen FROM users "
            "ORDER BY last_seen DESC LIMIT ? OFFSET ?",
            (page_size, p * page_size),
        ).fetchall()
        conn.close()
        return rows, total, p

    return await asyncio.to_thread(_work)


async def db_search_users(query: str, limit: int = 10) -> list[tuple]:
    """Ism, username yoki ID bo'yicha foydalanuvchi qidiradi (kirill/lotin harflariga
    sezgir bo'lmasligi uchun qidiruv Python tomonida casefold bilan bajariladi)."""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT user_id, username, full_name, is_banned, message_count, last_seen FROM users "
            "ORDER BY last_seen DESC"
        ).fetchall()
        conn.close()
        q = query.strip().lstrip("@").casefold()
        if not q:
            return []
        out = []
        for r in rows:
            if q in str(r[0]) or q in (r[1] or "").casefold() or q in (r[2] or "").casefold():
                out.append(r)
        return out[:limit]

    return await asyncio.to_thread(_work)


async def db_get_user_row(user_id: int) -> tuple | None:
    """Qaytaradi: (user_id, username, full_name, first_seen, last_seen, message_count, is_banned) yoki None"""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT user_id, username, full_name, first_seen, last_seen, message_count, is_banned "
            "FROM users WHERE user_id=?",
            (user_id,),
        ).fetchone()
        conn.close()
        return row

    return await asyncio.to_thread(_work)


async def db_count_favorites(user_id: int) -> int:
    def _work() -> int:
        conn = sqlite3.connect(DB_PATH)
        n = conn.execute("SELECT COUNT(*) FROM favorites WHERE user_id=?", (user_id,)).fetchone()[0]
        conn.close()
        return n

    return await asyncio.to_thread(_work)


async def db_find_user(query: str) -> tuple | None:
    """Berilgan ID raqami yoki @username orqali foydalanuvchini bazadan topadi."""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        q = query.strip()
        if q.lstrip("-").isdigit():
            row = conn.execute(
                "SELECT user_id, username, full_name, is_banned FROM users WHERE user_id=?", (int(q),)
            ).fetchone()
        else:
            uname = q.lstrip("@")
            row = conn.execute(
                "SELECT user_id, username, full_name, is_banned FROM users WHERE username=? COLLATE NOCASE",
                (uname,),
            ).fetchone()
        conn.close()
        return row

    return await asyncio.to_thread(_work)


# ============================================================
# ⭐ SEVIMLILAR (favorites) — DB funksiyalari
# ============================================================

async def db_add_favorite(user_id: int, fav_type: str, fav_key: str, label: str) -> bool:
    """Sevimlilarga qo'shadi. Agar allaqachon bor bo'lsa, False qaytaradi."""

    def _work() -> bool:
        conn = sqlite3.connect(DB_PATH)
        try:
            conn.execute(
                "INSERT INTO favorites (user_id, fav_type, fav_key, label, created_at) VALUES (?, ?, ?, ?, ?)",
                (user_id, fav_type, fav_key, label, now_tashkent().isoformat()),
            )
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    return await asyncio.to_thread(_work)


async def db_remove_favorite(user_id: int, fav_type: str, fav_key: str) -> None:
    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            "DELETE FROM favorites WHERE user_id=? AND fav_type=? AND fav_key=?",
            (user_id, fav_type, fav_key),
        )
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_remove_favorite_by_id(user_id: int, fav_id: int) -> None:
    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("DELETE FROM favorites WHERE user_id=? AND id=?", (user_id, fav_id))
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_is_favorite(user_id: int, fav_type: str, fav_key: str) -> bool:
    def _work() -> bool:
        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT id FROM favorites WHERE user_id=? AND fav_type=? AND fav_key=?",
            (user_id, fav_type, fav_key),
        ).fetchone()
        conn.close()
        return row is not None

    return await asyncio.to_thread(_work)


async def db_list_favorites(user_id: int) -> list[tuple]:
    """Qaytaradi: [(id, fav_type, fav_key, label), ...]"""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT id, fav_type, fav_key, label FROM favorites WHERE user_id=? ORDER BY created_at DESC",
            (user_id,),
        ).fetchall()
        conn.close()
        return rows

    return await asyncio.to_thread(_work)


# ============================================================
# 🔔 OBUNALAR (kunlik jadval + dars eslatmasi) — DB funksiyalari
# ============================================================

async def db_set_subscription(
    user_id: int, entity_type: str, entity_id: str, entity_label: str
) -> None:
    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            "INSERT INTO subscriptions (user_id, entity_type, entity_id, entity_label, "
            "daily_enabled, reminder_enabled, created_at) VALUES (?, ?, ?, ?, 1, 1, ?) "
            "ON CONFLICT(user_id) DO UPDATE SET entity_type=excluded.entity_type, "
            "entity_id=excluded.entity_id, entity_label=excluded.entity_label, "
            "daily_enabled=1, reminder_enabled=1",
            (user_id, entity_type, entity_id, entity_label, now_tashkent().isoformat()),
        )
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_remove_subscription(user_id: int) -> None:
    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("DELETE FROM subscriptions WHERE user_id=?", (user_id,))
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_get_subscription(user_id: int) -> tuple | None:
    """Qaytaradi: (entity_type, entity_id, entity_label, daily_enabled, reminder_enabled) yoki None"""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT entity_type, entity_id, entity_label, daily_enabled, reminder_enabled "
            "FROM subscriptions WHERE user_id=?",
            (user_id,),
        ).fetchone()
        conn.close()
        return row

    return await asyncio.to_thread(_work)


async def db_toggle_subscription_flag(user_id: int, flag: str) -> None:
    """flag: 'daily_enabled' yoki 'reminder_enabled' ni teskarisiga o'zgartiradi."""
    assert flag in ("daily_enabled", "reminder_enabled")

    def _work() -> None:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            f"UPDATE subscriptions SET {flag} = 1 - {flag} WHERE user_id=?", (user_id,)
        )
        conn.commit()
        conn.close()

    await asyncio.to_thread(_work)


async def db_list_daily_subscriptions() -> list[tuple]:
    """Qaytaradi: [(user_id, entity_type, entity_id, entity_label), ...] — kunlik yuborish yoqilganlar."""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT s.user_id, s.entity_type, s.entity_id, s.entity_label FROM subscriptions s "
            "JOIN users u ON u.user_id = s.user_id "
            "WHERE s.daily_enabled=1 AND u.is_banned=0"
        ).fetchall()
        conn.close()
        return rows

    return await asyncio.to_thread(_work)


async def db_list_reminder_subscriptions() -> list[tuple]:
    """Qaytaradi: [(user_id, entity_type, entity_id, entity_label), ...] — eslatma yoqilganlar."""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT s.user_id, s.entity_type, s.entity_id, s.entity_label FROM subscriptions s "
            "JOIN users u ON u.user_id = s.user_id "
            "WHERE s.reminder_enabled=1 AND u.is_banned=0"
        ).fetchall()
        conn.close()
        return rows

    return await asyncio.to_thread(_work)


# ============================================================
# 🤖 AI KVOTASI — DB funksiyalari
# Har bir foydalanuvchi kuniga cheklangan sondagina AI-yordamchiga
# murojaat qila oladi (xarajat va suiiste'moldan himoya uchun).
# ============================================================

async def db_get_ai_usage_today(user_id: int) -> int:
    def _work() -> int:
        conn = sqlite3.connect(DB_PATH)
        today = now_tashkent().date().isoformat()
        row = conn.execute(
            "SELECT count FROM ai_usage WHERE user_id=? AND usage_date=?", (user_id, today)
        ).fetchone()
        conn.close()
        return row[0] if row else 0

    return await asyncio.to_thread(_work)


async def db_increment_ai_usage(user_id: int) -> int:
    """Bugungi hisobni 1taga oshiradi va yangi qiymatni qaytaradi."""

    def _work() -> int:
        conn = sqlite3.connect(DB_PATH)
        today = now_tashkent().date().isoformat()
        conn.execute(
            "INSERT INTO ai_usage (user_id, usage_date, count) VALUES (?, ?, 1) "
            "ON CONFLICT(user_id, usage_date) DO UPDATE SET count = count + 1",
            (user_id, today),
        )
        conn.commit()
        new_count = conn.execute(
            "SELECT count FROM ai_usage WHERE user_id=? AND usage_date=?", (user_id, today)
        ).fetchone()[0]
        conn.close()
        return new_count

    return await asyncio.to_thread(_work)


async def db_get_ai_stats_today() -> list[tuple]:
    """Admin uchun: bugun eng ko'p AI'dan foydalangan foydalanuvchilar.
    Qaytaradi: [(user_id, username, full_name, count), ...]"""

    def _work():
        conn = sqlite3.connect(DB_PATH)
        today = now_tashkent().date().isoformat()
        rows = conn.execute(
            "SELECT a.user_id, u.username, u.full_name, a.count FROM ai_usage a "
            "JOIN users u ON u.user_id = a.user_id "
            "WHERE a.usage_date=? ORDER BY a.count DESC LIMIT 10",
            (today,),
        ).fetchall()
        conn.close()
        return rows

    return await asyncio.to_thread(_work)


class AdminLoggerMiddleware(BaseMiddleware):
    """Har bir foydalanuvchi harakati (matn, rasm, fayl, tugma bosish) haqida
    adminga (ADMIN_ID) xabar yuboradi. Matn bo'lmagan xabarlar (rasm, ovozli
    xabar, video va h.k.) uchun xabarning O'ZI ham (media fayli bilan birga)
    to'g'ridan-to'g'ri adminga nusxalab (copy_message) yuboriladi — shunda
    admin nafaqat "rasm yubordi" degan yozuvni, balki rasmning o'zini ham
    ko'radi."""

    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: dict
    ):
        user = None
        action_text = ""
        message_obj: Message | None = None

        # 1. Agar foydalanuvchi matn yozsa, rasm/fayl/stiker tashlasa yoki /start bossa
        if isinstance(event, Message):
            user = event.from_user
            message_obj = event

            if event.text:
                action_text = f"yozdi: <i>{esc(event.text)}</i>"
            elif event.photo:
                caption = f" (izoh: {esc(event.caption)})" if event.caption else ""
                action_text = f"rasm yubordi{caption}"
            elif event.document:
                fname = event.document.file_name or "noma'lum"
                action_text = f"fayl yubordi: <i>{esc(fname)}</i>"
            elif event.video:
                action_text = "video yubordi"
            elif event.voice:
                action_text = "ovozli xabar yubordi"
            elif event.audio:
                action_text = "audio yubordi"
            elif event.sticker:
                action_text = f"stiker yubordi: {event.sticker.emoji or ''}"
            elif event.video_note:
                action_text = "video-xabar (doira) yubordi"
            elif event.location:
                action_text = "joylashuvini yubordi"
            elif event.contact:
                action_text = "kontakt yubordi"
            else:
                action_text = "boshqa turdagi xabar yubordi"

        # 2. Agar foydalanuvchi botdagi biror inline tugmani bossa
        elif isinstance(event, CallbackQuery):
            user = event.from_user
            action_text = f"tugmani bosdi (data: <code>{esc(event.data)}</code>)"

        if user and user.id != ADMIN_ID:
            # 👥 Foydalanuvchini bazaga qayd etamiz (birinchi/oxirgi faollik, harakatlar soni)
            await db_touch_user(user.id, user.username, user.full_name)

            # 🚫 Bloklangan foydalanuvchi bo'lsa — handler'ga umuman yo'l qo'ymaymiz
            if await db_is_banned(user.id):
                if isinstance(event, Message):
                    try:
                        await event.answer("🚫 Siz botdan foydalanish huquqidan mahrum qilingansiz.")
                    except Exception:
                        pass
                elif isinstance(event, CallbackQuery):
                    try:
                        await event.answer("🚫 Siz bloklangansiz.", show_alert=True)
                    except Exception:
                        pass
                return  # handler chaqirilmaydi

            username_text = user.username if user.username else "yo'q"
            header = (
                f"ℹ️ <b>{esc(user.full_name)}</b> (ID: <code>{user.id}</code>)\n"
                f"@{esc(username_text)} — {action_text}"
            )

            bot = data["bot"]
            try:
                await bot.send_message(chat_id=ADMIN_ID, text=header, parse_mode="HTML")
            except Exception as e:
                print(f"Adminga kuzatuv xabarini yuborishda xato: {e}")

            # Matn bo'lmagan xabarlar uchun (rasm, video, ovozli xabar, fayl,
            # stiker va h.k.) — xabarning O'ZINI ham adminga nusxalaymiz, shunda
            # admin faqat tavsifni emas, aynan yuborilgan faylning o'zini ko'radi.
            if message_obj is not None and not message_obj.text:
                try:
                    await bot.copy_message(
                        chat_id=ADMIN_ID,
                        from_chat_id=message_obj.chat.id,
                        message_id=message_obj.message_id,
                    )
                except Exception as e:
                    print(f"Adminga xabarni nusxalab (copy_message) yuborishda xato: {e}")

        return await handler(event, data)


# ============================================================
# 🛑 GLOBAL XATO KUZATUVI
# Botning istalgan handler'ida (matn, tugma, va h.k.) kutilmagan xato yuz
# bersa — bot yiqilib qolmaydi, xato to'liq traceback bilan ADMIN_ID'ga
# avtomatik yuboriladi, foydalanuvchiga esa tushunarli xabar ko'rsatiladi.
# ============================================================

@dp.error()
async def global_error_handler(event: ErrorEvent) -> None:
    exception = event.exception
    update = event.update

    tb_text = "".join(
        traceback.format_exception(type(exception), exception, exception.__traceback__)
    )
    print(f"[XATO] {tb_text}")

    # 👑 Adminga to'liq xato ma'lumotini yuboramiz (juda uzun bo'lsa qisqartiramiz).
    # ⚠️ Traceback ichida "<module>" kabi matnlar uchraydi — HTML rejimida bular
    # xabarni buzib yuboradi, shuning uchun esc() bilan xavfsiz qilinadi.
    if BOT_INSTANCE is not None:
        update_info = ""
        try:
            update_info = f"\nUpdate: <code>{esc(str(update)[:500])}</code>"
        except Exception:
            pass

        error_text = f"🛑 <b>BOTDA XATO YUZ BERDI</b>\n\n<code>{esc(tb_text[-3000:])}</code>{update_info}"
        chunks = [error_text[i:i + 3900] for i in range(0, len(error_text), 3900)] or [error_text]
        for chunk in chunks[:3]:  # ortiqcha spam bo'lmasligi uchun ko'pi bilan 3 ta xabar
            try:
                await BOT_INSTANCE.send_message(chat_id=ADMIN_ID, text=chunk, parse_mode=ParseMode.HTML)
            except Exception as e:
                print(f"[XATO] Adminga xato xabarini yuborib bo'lmadi: {e}")

    # 🙋 Foydalanuvchiga ham tushunarli xabar ko'rsatishga harakat qilamiz
    try:
        chat_id = None
        if update.message:
            chat_id = update.message.chat.id
        elif update.callback_query and update.callback_query.message:
            chat_id = update.callback_query.message.chat.id

        if chat_id and BOT_INSTANCE is not None:
            await BOT_INSTANCE.send_message(
                chat_id=chat_id,
                text="⚠️ Kutilmagan xatolik yuz berdi. Iltimos, qayta urinib ko'ring yoki /start bosing.\n"
                     "Bu haqda admin allaqachon xabardor qilindi.",
            )
    except Exception as e:
        print(f"[XATO] Foydalanuvchiga xato xabarini yuborib bo'lmadi: {e}")


class SearchStates(StatesGroup):
    """Talaba/o'qituvchi ismini kiritishni kutish holati."""
    waiting_name = State()


class AIStates(StatesGroup):
    """AI-yordamchi bilan suhbat holati."""
    chatting = State()


class AdminStates(StatesGroup):
    """Admin panelidagi matn kutish holatlari."""
    waiting_user_search = State()  # foydalanuvchini ism/username/ID bo'yicha qidirish
    waiting_dm = State()           # tanlangan foydalanuvchiga xabar yozish


# ============================================================
# 🤖 AI-YORDAMCHI (Claude / Anthropic API)
# Foydalanuvchi tabiiy tilda savol yozadi ("A blokda hozir bo'sh xona
# bormi?", "Aliyev Vali ertaga band bo'ladimi?" va h.k.) — AI so'rovni
# tushunib, kerakli tool'ni (funksiyani) chaqiradi va natijani odam
# tiliga o'girib javob beradi.
#
# ⚙️ SOZLASH: Render'da "Environment" bo'limiga ANTHROPIC_API_KEY
# o'zgaruvchisini qo'shing (https://console.anthropic.com dan olinadi).
# Agar u sozlanmagan bo'lsa, AI-yordamchi tugmasi "hozircha ishlamaydi"
# deb xabar beradi, lekin bot o'zi normal ishlashda davom etadi.
# ============================================================

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
AI_MODEL = os.environ.get("AI_MODEL", "claude-haiku-4-5-20251001")
# 🚦 Har bir foydalanuvchi kuniga necha marta AI-yordamchiga murojaat qila
# olishi. Xarajat va suiiste'moldan himoya uchun. Render'da AI_DAILY_LIMIT
# o'zgaruvchisi orqali o'zgartirish mumkin. Admin (ADMIN_ID) bu limitdan ozod.
AI_DAILY_LIMIT = int(os.environ.get("AI_DAILY_LIMIT", "15"))
ai_client: AsyncAnthropic | None = AsyncAnthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else None

AI_SYSTEM_PROMPT = (
    "Siz JIDU (Jizzax davlat pedagogika universiteti) 'Bo'sh xonalar boti'ning "
    "AI-yordamchisisiz. Foydalanuvchilar sizdan xonalar bandligi, talaba, "
    "o'qituvchi yoki guruh dars jadvali haqida so'raydi.\n\n"
    "QOIDALAR:\n"
    "1. Har doim sizga berilgan tool (funksiya)lardan foydalanib ANIQ ma'lumot "
    "oling — hech qachon o'zingizdan taxmin qilib javob bermang.\n"
    "2. Agar foydalanuvchi sana aytmasa, bugungi kunni nazarda tutgan deb hisoblang.\n"
    "3. Sana kerak bo'lganda har doim YYYY-MM-DD formatida yuboring (sizga "
    "xabarning oxirida hozirgi sana va vaqt beriladi).\n"
    "4. Javoblaringiz o'zbek tilida, qisqa, tabiiy va aniq bo'lsin — HTML yoki "
    "Markdown belgilari ishlatmang, faqat oddiy matn yozing.\n"
    "5. Agar so'rov botning vazifasidan tashqarida bo'lsa (masalan umumiy "
    "bilim savoli), muloyimlik bilan botning vazifasi faqat xonalar va "
    "dars jadvali ekanini eslating.\n"
    "6. Agar biror ma'lumot (masalan qidirilgan odam) topilmasa, buni "
    "to'g'ridan-to'g'ri va ochiq ayting, o'ylab topmang.\n"
    "7. Odam ismi, familiyasi yoki guruh nomini tool'ga foydalanuvchi yozganidek, "
    "HECH QANDAY o'zgartirmasdan yuboring (o' va g' harflaridagi apostrofni ham, "
    "kirill harflarini ham aynan saqlang) — bot apostrof turlarini va yozilish "
    "farqlarini o'zi to'g'rilaydi."
)

AI_TOOLS = [
    {
        "name": "find_free_rooms",
        "description": (
            "Berilgan blok, qavat, sana va para (dars soati)da qaysi xonalar bo'sh, "
            "qaysilari band ekanini topadi. Foydalanuvchi 'hozir', 'ayni vaqtda' "
            "desa yoki vaqt aytmasa, para='now' yuboring."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "block": {"type": "string", "description": "Blok nomi: A, B, D yoki F"},
                "floor": {"type": "string", "description": "Qavat raqami (masalan '2') yoki butun blok uchun 'ALL'"},
                "date": {"type": "string", "description": "Sana YYYY-MM-DD formatida"},
                "para": {"type": "string", "description": "Para raqami (1-7) yoki hozirgi vaqt uchun 'now'"},
            },
            "required": ["block", "floor", "date", "para"],
        },
    },
    {
        "name": "find_free_room_streaks",
        "description": (
            "Berilgan blok/qavat/sanada ketma-ket (2 va undan ortiq para) bo'sh "
            "turadigan xonalarni topadi — foydalanuvchi 'uzoq vaqt bo'sh xona', "
            "'ketma-ket bo'sh xona' desa shuni ishlating."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "block": {"type": "string", "description": "Blok nomi: A, B, D, F yoki barcha bloklar uchun 'ALL'"},
                "floor": {"type": "string", "description": "Qavat raqami yoki 'ALL' (block='ALL' bo'lsa buni ham 'ALL' qiling)"},
                "date": {"type": "string", "description": "Sana YYYY-MM-DD formatida"},
            },
            "required": ["block", "floor", "date"],
        },
    },
    {
        "name": "room_day_schedule",
        "description": (
            "BITTA aniq xonaning (masalan A305, D-412) berilgan kundagi barcha "
            "paralar bo'yicha holatini (qaysi para band, qaysi biri bo'sh) ko'rsatadi. "
            "Foydalanuvchi aniq bir xona haqida so'rasa ('A305 bugun bandmi?', "
            "'D412 ertaga qachon bo'sh?') shuni ishlating."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "room": {"type": "string", "description": "Xona nomi, masalan 'A305' yoki 'D-412'"},
                "date": {"type": "string", "description": "Sana YYYY-MM-DD formatida"},
            },
            "required": ["room", "date"],
        },
    },
    {
        "name": "search_person_schedule",
        "description": (
            "Talaba, o'qituvchi yoki guruhning ism/familiya/nomi bo'yicha qidirib, "
            "berilgan kundagi dars jadvalini yoki hozirgi (joriy) holatini topadi."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "entity_type": {
                    "type": "string",
                    "enum": ["students", "teachers", "groups"],
                    "description": "students=talaba, teachers=o'qituvchi, groups=guruh",
                },
                "name": {"type": "string", "description": "Ism-familiya, guruh nomi yoki ID raqami (foydalanuvchi yozganidek, o'zgartirmasdan)"},
                "date": {"type": "string", "description": "Sana YYYY-MM-DD formatida, yoki hozirgi holat uchun 'now'"},
            },
            "required": ["entity_type", "name", "date"],
        },
    },
]


def _ai_resolve_floor(block: str, floor: str) -> str:
    """AI yuborgan qavat qiymatini UNIVERSITY_MAP kalitlariga moslashtiradi
    (masalan '2-qavat' -> '2'), topilmasa 'ALL' qaytaradi."""
    floor = str(floor).strip().upper()
    if floor == "ALL" or floor in UNIVERSITY_MAP.get(block, {}):
        return floor
    digits = "".join(ch for ch in floor if ch.isdigit())
    if digits in UNIVERSITY_MAP.get(block, {}):
        return digits
    return "ALL"


def _ai_resolve_date(date_str: str) -> str:
    """AI yuborgan sanani tekshiradi, noto'g'ri formatda bo'lsa bugungi kunni qaytaradi."""
    try:
        datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return date_str
    except (ValueError, TypeError):
        return now_tashkent().date().strftime("%Y-%m-%d")


async def execute_ai_tool(name: str, tool_input: dict) -> str:
    """AI tanlagan tool'ni haqiqiy bot funksiyalari orqali bajaradi va
    natijani AI tushunadigan oddiy matn ko'rinishida qaytaradi."""
    try:
        if name == "find_free_rooms":
            block = str(tool_input.get("block", "")).strip().upper()
            if block not in UNIVERSITY_MAP:
                return f"Xato: '{block}' nomli blok mavjud emas. Mavjud bloklar: A, B, D, F."
            floor = _ai_resolve_floor(block, tool_input.get("floor", "ALL"))
            date_str = _ai_resolve_date(tool_input.get("date", ""))
            para = str(tool_input.get("para", "now")).strip().lower()

            rooms = get_rooms_for_floor(block, floor)
            if not rooms:
                return f"{block} blok, {floor_label(floor)}da ro'yxatga olingan xona yo'q."

            if para == "now":
                now = now_tashkent()
                current_time = now.strftime("%H:%M")
                start_t = None
                for pid, (st, et, disp) in PARALAR.items():
                    if st <= current_time <= et:
                        start_t = st
                        break
                if not start_t:
                    return (
                        f"Hozir ({now.strftime('%H:%M')}) darsdan tashqari vaqt — "
                        f"{block} blok, {floor_label(floor)}dagi barcha xonalar rasman bo'sh."
                    )
            else:
                para_info = PARALAR.get(para)
                if not para_info:
                    return f"Xato: {para}-para mavjud emas. Para 1 dan 7 gacha bo'lishi kerak."
                start_t = para_info[0]

            busy, invalid = await get_room_statuses(rooms, block, date_str, start_t)
            free = [r for r in rooms if r not in busy and r not in invalid]
            free_display = ", ".join(format_room_display(block, r) for r in free[:40]) or "yo'q"
            busy_parts = []
            for r, info in list(busy.items())[:15]:
                busy_parts.append(f"{format_room_display(block, r)} ({info.get('title', 'fan nomsiz')})")
            busy_display = "; ".join(busy_parts) or "yo'q"

            return (
                f"{date_str} sana, {block} blok, {floor_label(floor)}.\n"
                f"BO'SH XONALAR: {free_display}\n"
                f"BAND XONALAR: {busy_display}"
            )

        elif name == "find_free_room_streaks":
            block = str(tool_input.get("block", "ALL")).strip().upper()
            date_str = _ai_resolve_date(tool_input.get("date", ""))

            if block == "ALL":
                results = await get_room_free_streaks_all_blocks(date_str)
                if not results:
                    return f"{date_str} sanada hech qaysi blokda 2+ ketma-ket bo'sh para topilmadi."
                parts = [
                    f"{format_room_display(b, r)}: {'-'.join(str(p) for p in run)}-paralar"
                    for b, r, run, length in results[:15]
                ]
                return f"{date_str} sana, barcha bloklar bo'yicha eng uzun bo'sh ketma-ketliklar: " + "; ".join(parts)

            if block not in UNIVERSITY_MAP:
                return f"Xato: '{block}' nomli blok mavjud emas. Mavjud bloklar: A, B, D, F, ALL."
            floor = _ai_resolve_floor(block, tool_input.get("floor", "ALL"))
            rooms = get_rooms_for_floor(block, floor)
            results = await get_room_free_streaks(rooms, block, date_str)
            if not results:
                return f"{date_str} sanada {block} blok, {floor_label(floor)}da 2+ ketma-ket bo'sh para topilmadi."
            parts = [
                f"{format_room_display(block, r)}: {'-'.join(str(p) for p in run)}-paralar"
                for r, run, length in results[:15]
            ]
            return f"{date_str} sana, {block} blok, {floor_label(floor)}: " + "; ".join(parts)

        elif name == "room_day_schedule":
            room_text = str(tool_input.get("room", "")).strip()
            date_str = _ai_resolve_date(tool_input.get("date", ""))
            key = re.sub(r"[^A-Za-z0-9]", "", room_text).upper()
            matches = find_rooms_by_text(room_text, limit=8)
            if not matches:
                return f"Xato: '{room_text}' nomli xona ro'yxatda topilmadi."
            exact = [m for m in matches if m[1] == key]
            if exact:
                block, room = exact[0]
            elif len(matches) == 1:
                block, room = matches[0]
            else:
                names = ", ".join(format_room_display(b, r) for b, r in matches)
                return f"'{room_text}' bo'yicha bir nechta xona mos keldi: {names}. Aniq xona nomini so'rang."

            by_para = await get_room_day_events(block, room, date_str)
            display = format_room_display(block, room)
            if by_para is None:
                return f"{display} xonasi UWED tizimida topilmadi."
            lines = []
            for pid, (st, et, disp) in PARALAR.items():
                e = by_para.get(pid)
                if e:
                    lines.append(
                        f"{pid}-para ({st}-{et}): BAND — {e.get('title') or 'fan nomsiz'}, "
                        f"o'qituvchi: {e.get('instructor') or 'nomalum'}"
                    )
                else:
                    lines.append(f"{pid}-para ({st}-{et}): bo'sh")
            return f"{display} xonasi, {date_str}:\n" + "\n".join(lines)

        elif name == "search_person_schedule":
            entity_type = tool_input.get("entity_type", "students")
            query = str(tool_input.get("name", "")).strip()
            date_str = str(tool_input.get("date", "now")).strip()

            if not query:
                return "Xato: qidirish uchun ism yoki nom berilmadi."

            results = await search_entities(entity_type, query)
            if not results:
                type_label = {"students": "talaba", "teachers": "o'qituvchi", "groups": "guruh"}.get(entity_type, "")
                return f"'{query}' nomli {type_label} topilmadi."

            item = results[0]
            entity_id = item.get("id")
            display_name = entity_display_name(item)
            code = {"students": "s", "teachers": "t", "groups": "g"}.get(entity_type, "s")
            event_type = CODE_TO_EVENT_TYPE.get(code, "STUDENT")
            extra_note = ""
            if len(results) > 1:
                others = ", ".join(entity_display_name(r) for r in results[1:4])
                extra_note = (
                    f" (Diqqat: '{query}' bo'yicha {len(results)} ta natija topildi, "
                    f"birinchisi ko'rsatilyapti. Boshqa mos natijalar: {others}.)"
                )

            now = now_tashkent()
            if date_str == "now":
                events = await fetch_week_events(entity_id, now.date(), event_type=event_type)
                today_str = now.date().strftime("%Y-%m-%d")
                current_time = now.strftime("%H:%M")
                current_event = None
                for e in events:
                    if str(e.get("date", "")) == today_str:
                        st, et = str(e.get("start_time", "")), str(e.get("end_time", ""))
                        if st <= current_time <= et:
                            current_event = e
                            break
                if current_event:
                    return (
                        f"{display_name} hozir ({current_time}) band: "
                        f"{current_event.get('title') or 'fan nomsiz'}, "
                        f"xona {get_event_room(current_event)}, "
                        f"tugashi {current_event.get('end_time', '')}.{extra_note}"
                    )
                return f"{display_name} hozir ({current_time}) band emas / darsda emas.{extra_note}"

            date_str = _ai_resolve_date(date_str)
            target_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            events = await fetch_week_events(entity_id, target_date, event_type=event_type)
            day_events = [e for e in events if str(e.get("date", "")) == date_str]
            if not day_events:
                return f"{display_name} uchun {date_str} kunida darslar yo'q.{extra_note}"
            day_events.sort(key=lambda e: str(e.get("start_time", "")))
            parts = [
                f"{e.get('start_time', '')}-{e.get('end_time', '')}: {e.get('title') or 'fan nomsiz'}, "
                f"xona {get_event_room(e)}"
                for e in day_events
            ]
            return f"{display_name} uchun {date_str} kunlik jadvali: " + "; ".join(parts) + extra_note

        return f"Noma'lum funksiya: {name}"
    except Exception as e:
        print(f"[AI TOOL] '{name}' bajarishda xato: {e}")
        return f"Bu funksiyani bajarishda texnik xato yuz berdi: {e}"


async def run_ai_conversation(messages: list[dict], context_note: str) -> str:
    """Claude bilan tool-calling siklini yuritadi: AI tool tanlasa — uni
    bajaramiz, natijani qaytaramiz — AI yakuniy matnli javob bergunicha."""
    msgs = list(messages)
    if msgs and msgs[-1]["role"] == "user":
        msgs[-1] = {"role": "user", "content": msgs[-1]["content"] + context_note}

    for _ in range(5):  # cheksiz aylanishning oldini olish uchun chegara
        response = await ai_client.messages.create(
            model=AI_MODEL,
            max_tokens=800,
            system=AI_SYSTEM_PROMPT,
            tools=AI_TOOLS,
            messages=msgs,
        )

        if response.stop_reason != "tool_use":
            text_parts = [block.text for block in response.content if block.type == "text"]
            result = "\n".join(text_parts).strip()
            return result or "Kechirasiz, javob shakllantira olmadim. Savolni boshqacharoq yozib ko'ring."

        msgs.append({"role": "assistant", "content": response.content})
        tool_results = []
        for block in response.content:
            if block.type == "tool_use":
                result_text = await execute_ai_tool(block.name, block.input)
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": result_text,
                })
        msgs.append({"role": "user", "content": tool_results})

    return "Kechirasiz, so'rovingiz murakkab bo'lib qoldi. Iltimos, aniqroq va qisqaroq yozib ko'ring."


# 🔗 UWED API asosiy manzili (xona qidirish va events uchun)
API_BASE = "https://manage.uwed.uz/api/timetable/public/v1"

# 🔗 Talaba/o'qituvchi qidirish uchun ALOHIDA API manzili (tasdiqlangan, DevTools orqali topilgan)
EDUCATION_API_BASE = "https://manage.uwed.uz/api/education/public/v1"

# ⚠️ GURUH qidiruv manzillari — TASDIQLANMAGAN (taxminiy). Talaba/o'qituvchi
# qidiruvi kabi DevTools orqali tekshirilmagan, shuning uchun bir nechta
# ehtimoliy variant ketma-ket sinaladi. Agar birortasi ham ishlamasa,
# brauzer DevTools > Network tab orqali "guruh" qidirilganda ketayotgan
# haqiqiy so'rov manzilini shu ro'yxatga qo'shing.
GROUP_SEARCH_URLS = [
    f"{EDUCATION_API_BASE}/group/search/for-timetable",
    f"{EDUCATION_API_BASE}/groups/search/for-timetable",
    f"{API_BASE}/search/groups",
]

# 🏢 JIDU barcha binolari, qavatlari va xonalarining ro'yxati.
# Diqqat: bot ishlash vaqtida har bir xonani UWED tizimida tekshiradi va
# topilmagan xonalarni natijalardan avtomatik chiqarib tashlaydi, shuning
# uchun bu yerda ortiqcha/noto'g'ri xona bo'lsa ham natijaga ta'sir qilmaydi.
UNIVERSITY_MAP = {
    "A": {
        "1": ["A110", "A111", "A113", "A114"],
        "2": ["A201", "A205", "A206", "A208", "A211", "A212", "A214", "A216", "A217", "A218", "A219", "A222"],
        "3": ["A301", "A303", "A304", "A305", "A307", "A310", "A311", "A312", "A315", "A317", "A318", "A319", "A320", "A323"],
        "4": ["A400", "A401", "A402", "A404", "A407", "A408", "A410", "A412", "A413", "A414", "A415"],
    },
    "B": {
        "1": [],  # Bo'sh xonalar mavjud emas
        "2": ["B204", "B206", "B208"],
        "3": ["B301", "B303", "B305", "B307", "B309"],
        "4": ["B401", "B403", "B405", "B407", "B408", "B409", "B410"],
        "5": ["B510", "B512", "B513"],
        "6": ["B601", "B603", "B605", "B607", "B608", "B610", "B612", "B613"],
        "7": ["B703", "B705", "B707", "B708", "B710", "B712"],
    },
    "D": {
        "1": [],  # Bo'sh xonalar mavjud emas
        "2": ["D201", "D202", "D203", "D204", "D205", "D208", "D211", "D212", "D213", "D214", "D215", "D216", "D217", "D218", "D222", "D223"],
        "3": ["D302", "D303", "D304", "D305", "D308", "D311", "D312", "D313", "D314", "D315", "D316", "D319", "D321", "D322", "D325", "D316A"],
        "4": ["D401", "D402", "D403", "D404", "D405", "D407", "D410", "D412", "D413", "D414", "D419", "D420", "D421", "D422", "D412A"],
    },
    "F": {
        "1": ["F1", "F3", "F4", "F5", "F6", "F7", "F8", "F102", "F106", "F107"],
    }
}

# 📆 Hafta kunlari o'zbek tilida
HAFTA_KUNLARI = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]


def build_date_label(offset_days: int, date_obj: datetime.date) -> str:
    """Sana tanlash tugmalari uchun chiroyli matn tuzadi. "Bugun" va "Ertaga"
    so'zlari yoniga ham hafta kuni nomi qo'shiladi (masalan
    "📅 Bugun, Chorshanba (2026-09-23)"), qolgan kunlar uchun esa faqat hafta
    kuni va sana ko'rsatiladi (masalan "📅 Juma (2026-09-25)")."""
    weekday_num = date_obj.isoweekday()
    day_name = HAFTA_KUNLARI[weekday_num - 1] if weekday_num != 7 else "Yakshanba"
    date_str = date_obj.strftime('%Y-%m-%d')

    if offset_days == 0:
        return f"📅 Bugun, {day_name} ({date_str})"
    elif offset_days == 1:
        return f"📅 Ertaga, {day_name} ({date_str})"
    return f"📅 {day_name} ({date_str})"

# ⏰ Universitet rasmiy dars soatlari (Paralar)
PARALAR = {
    "1": ("08:30", "09:40", "1-para (08:30 - 09:40)"),
    "2": ("09:55", "11:05", "2-para (09:55 - 11:05)"),
    "3": ("11:20", "12:30", "3-para (11:20 - 12:30)"),
    "4": ("12:45", "13:55", "4-para (12:45 - 13:55)"),
    "5": ("14:30", "15:40", "5-para (14:30 - 15:40)"),
    "6": ("15:55", "17:05", "6-para (15:55 - 17:05)"),
    "7": ("17:20", "18:30", "7-para (17:20 - 18:30)"),
}

# ⏰ start_time -> para raqami (teskari xarita, talaba/o'qituvchi jadvalida kerak bo'ladi)
START_TIME_TO_PARA = {v[0]: k for k, v in PARALAR.items()}

# 🗂 Xona nomi -> UUID keshi (bir marta topilgach qayta so'ralmaydi)
ROOM_ID_CACHE: dict[str, str] = {}
# Xona nomi -> UWED tizimida topilmagani ma'lum bo'lgan xonalar (qayta-qayta so'ramaslik uchun)
INVALID_ROOM_CACHE: set[str] = set()

# Brauzerga o'xshash so'rov headerlari - ba'zi serverlar header'siz so'rovlarni 404/403 bilan bloklaydi
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://manage.uwed.uz/public/pages/inf-tt",
    "Origin": "https://manage.uwed.uz",
    "X-Requested-With": "XMLHttpRequest",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Dest": "empty",
    "X-Public-Key": "public-key",
}

# Umumiy aiohttp sessiyasi (bloklamaydigan so'rovlar uchun)
http_session: aiohttp.ClientSession | None = None


def format_room_display(block: str, room: str) -> str:
    """Xona kodini (masalan 'A400') API qidiruv formatiga ('A-400') o'giradi.
    Bu format ham foydalanuvchiga ko'rsatish uchun, ham /search/rooms so'rovi uchun ishlatiladi."""
    # F blok xonalari ham boshqa bloklar kabi API formatida yoziladi:
    # F1 -> F-1, F102 -> F-102.
    if room == "B701B":
        return "B-701-B"
    if room == "D316A":
        return "D-316-A"
    if room == "D412A":
        return "D-412-A"
    return f"{block}-{room[len(block):]}"


def get_rooms_for_floor(block: str, floor: str) -> list[str]:
    """Berilgan qavatdagi xonalar ro'yxatini qaytaradi.
    floor == "ALL" bo'lsa, blokning BARCHA qavatlaridagi xonalarni qaytaradi."""
    if floor == "ALL":
        rooms: list[str] = []
        for floor_rooms in UNIVERSITY_MAP[block].values():
            rooms.extend(floor_rooms)
        return rooms
    return UNIVERSITY_MAP[block][floor]


def get_room_floor(block: str, room: str) -> str:
    """Berilgan xona qaysi qavatda joylashganini topadi (guruhlab ko'rsatish uchun)."""
    for floor, rooms in UNIVERSITY_MAP[block].items():
        if room in rooms:
            return floor
    return "?"


def floor_label(floor: str) -> str:
    return "barcha qavatlar" if floor == "ALL" else f"{floor}-qavat"


def find_rooms_by_text(text: str, limit: int = 5) -> list[tuple[str, str]]:
    """Foydalanuvchi yozgan xona nomini ("A305", "a-305", "TTJ-B-108", "D41") ro'yxatdagi
    xonalarga moslashtiradi. Aniq moslik birinchi, so'ng boshlanishi mos kelganlar.
    Qaytaradi: [(blok, xona_kodi), ...]"""
    key = re.sub(r"[^A-Za-z0-9]", "", text or "").upper()
    if not key:
        return []
    exact: list[tuple[str, str]] = []
    partial: list[tuple[str, str]] = []
    for block, floors in UNIVERSITY_MAP.items():
        for rooms in floors.values():
            for room in rooms:
                if room == key:
                    exact.append((block, room))
                elif room.startswith(key):
                    partial.append((block, room))
    return (exact + partial)[:limit]


async def get_room_id(block: str, room: str) -> str | None:
    """Xonaning UUID'ini topadi (/search/rooms orqali) va keshlab qo'yadi.
    Topilmasa, buni ham keshlab (qayta-qayta so'ramaslik uchun) None qaytaradi."""
    if room in ROOM_ID_CACHE:
        return ROOM_ID_CACHE[room]
    if room in INVALID_ROOM_CACHE:
        return None

    query = format_room_display(block, room)
    url = f"{API_BASE}/search/rooms"
    try:
        async with http_session.get(
            url,
            params={"search": query},
            headers=REQUEST_HEADERS,
            timeout=aiohttp.ClientTimeout(total=10),
        ) as resp:
            if resp.status != 200:
                print(f"Xona qidirishda xato ({room}): status={resp.status}")
                INVALID_ROOM_CACHE.add(room)
                return None
            data = await resp.json(content_type=None)
    except Exception as e:
        print(f"Xona qidirishda xato ({room}): {e}")
        return None

    if not data.get("success") or not data.get("data"):
        print(f"Xona UWED tizimida topilmadi: {room} (so'rov: {query})")
        INVALID_ROOM_CACHE.add(room)
        return None

    results = data["data"]
    exact = next(
        (r for r in results if r.get("room_name", "").replace("-", "") == room),
        None,
    )
    chosen = exact or results[0]
    room_id = chosen.get("id")
    if room_id:
        ROOM_ID_CACHE[room] = room_id
        return room_id

    INVALID_ROOM_CACHE.add(room)
    return None


# ============================================================
# ⚡ TEZKOR KESH (5 daqiqa)
# "Eng bo'sh xonalar (barcha bloklar)" kabi so'rovlar o'nlab xonalarni
# tekshiradi, ko'p hollarda BIR XIL haftaning ma'lumotini qayta-qayta so'raydi.
# Shu sabab natijalarni qisqa vaqtga (5 daqiqa) xotirada saqlab qo'yamiz —
# bu botni sezilarli tezlashtiradi va UWED serveriga ortiqcha yuk tushmaydi.
# ============================================================
EVENTS_CACHE: dict[tuple, tuple[float, list[dict]]] = {}
EVENTS_CACHE_TTL_SECONDS = 300  # 5 daqiqa


async def fetch_week_events(request_id: str, any_date: datetime.date, event_type: str = "ROOM") -> list[dict]:
    """Berilgan sana joylashgan haftaning barcha darslarini olib keladi (Dushanbadan boshlab).
    event_type: "ROOM" (xona), "STUDENT" (talaba, ✅ tasdiqlangan) yoki "TEACHER" (o'qituvchi, taxmin).
    Natija 5 daqiqaga keshlanadi (tezlik uchun)."""
    monday = any_date - datetime.timedelta(days=any_date.weekday())
    cache_key = (request_id, monday.isoformat(), event_type)

    cached = EVENTS_CACHE.get(cache_key)
    if cached and (time.monotonic() - cached[0]) < EVENTS_CACHE_TTL_SECONDS:
        return cached[1]

    url = f"{API_BASE}/events"
    params = {
        "request_id": request_id,
        "type": event_type,
        "start_time": monday.strftime("%Y-%m-%d"),
    }
    try:
        async with http_session.get(
            url,
            params=params,
            headers=REQUEST_HEADERS,
            timeout=aiohttp.ClientTimeout(total=10),
        ) as resp:
            data = await resp.json(content_type=None)
    except Exception as e:
        print(f"Jadval olishda xato: {e}")
        return []

    if not data.get("success"):
        return []
    result = data.get("data", [])
    EVENTS_CACHE[cache_key] = (time.monotonic(), result)

    # Xotira cheksiz o'smasligi uchun eskirgan yozuvlarni tozalaymiz
    if len(EVENTS_CACHE) > 4000:
        now_m = time.monotonic()
        for k in [k for k, (ts, _) in EVENTS_CACHE.items() if now_m - ts >= EVENTS_CACHE_TTL_SECONDS]:
            EVENTS_CACHE.pop(k, None)
    return result


# ============================================================
# 🔤 O' / G' HARFLARI VA YOZILISH FARQLARINI TO'G'RILASH
#
# Muammo: o' va g' harflaridagi "tutuq belgisi"ni telefon/kompyuter klaviaturalari
# turlicha yuboradi (oddiy ' , egri ’ ‘ , rasmiy ʻ ʼ , ` ´ ...), UWED bazasida esa
# bittasi saqlangan. Bundan tashqari ko'pchilik apostrofni umuman yozmaydi
# ("Oktam" o'rniga "O'ktam") yoki kirill harflarida yozadi ("Ўктам").
#
# ⚠️ Belgilar \u kodlari bilan yoziladi: oddiy matn sifatida nusxalaganda egri
# tirnoqlar (’ ‘) o'zi-o'zidan oddiy ' ga aylanib ketadi va ro'yxat buziladi
# (aynan shu sabab avval o' va g' "tanilmayotgan" edi).
#
# Qidiruv uch bosqichda ishlaydi (qarang: search_entities):
#   1) aniq ibora, barcha apostrof turlari bilan (serverning o'zi to'g'ri ism+familiya
#      juftligini toraytiradi);
#   2) "yumshoq": o/g va apostroflarsiz eng uzun bo'lak bo'yicha qidirib, natijani
#      ismlarni apostrofsiz solishtirib o'zimiz saralaymiz;
#   3) o/g dan keyin apostrof qo'shib ko'rish ("Ozod" -> "O'zod").
# ============================================================
APOSTROPHE_VARIANTS = [
    "\u02bb",  # ʻ — o‘ / g‘ uchun rasmiy belgi
    "\u0027",  # ' — klaviaturadagi oddiy apostrof
    "\u2019",  # ’ — o'ng egri tirnoq (iPhone / ko'p klaviaturalar shuni yuboradi)
    "\u2018",  # ‘ — chap egri tirnoq
    "\u02bc",  # ʼ — modifier apostrof
    "\u0060",  # ` — backtick
    "\u00b4",  # ´ — acute
    "\u2032",  # ′ — prime
    "\u02b9",  # ʹ — modifier prime
    "\u201b",  # ‛ — teskari 9 shaklidagi tirnoq
]
INSERT_CHARS = ("\u0027", "\u02bb", "\u2019")  # apostrof qo'shib sinaladigan asosiy turlar

_APOS_TABLE = {ord(ch): None for ch in APOSTROPHE_VARIANTS}
_AMBIG_SPLIT_RE = re.compile("[" + re.escape("".join(APOSTROPHE_VARIANTS)) + "oOgG]+")

# Kirill -> lotin (o'zbek) harflari xaritasi
_CYR_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo", "ж": "j",
    "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "x", "ц": "ts",
    "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "'", "ы": "i", "ь": "", "э": "e", "ю": "yu",
    "я": "ya", "ў": "o'", "қ": "q", "ғ": "g'", "ҳ": "h",
}


def has_cyrillic(text: str) -> bool:
    return any("\u0400" <= ch <= "\u04ff" for ch in text)


def cyr_to_lat(text: str) -> str:
    """Kirill harflarini o'zbek lotin yozuviga o'giradi (ў -> o', ғ -> g' va h.k.)."""
    out: list[str] = []
    for ch in text:
        low = ch.lower()
        rep = _CYR_MAP.get(low)
        if rep is None:
            out.append(ch)
            continue
        if ch != low and rep:  # bosh harf bo'lsa, almashtirishning birinchi harfi ham bosh bo'lsin
            rep = rep[0].upper() + rep[1:]
        out.append(rep)
    return "".join(out)


def strip_apostrophes(text: str) -> str:
    """Matndagi barcha apostrof turlarini olib tashlaydi."""
    return text.translate(_APOS_TABLE)


def _name_key(text: str) -> str:
    """Solishtirish uchun kalit: kirill bo'lsa lotinga o'giriladi, apostroflar olib tashlanadi,
    kichik harfga o'tkaziladi, ortiqcha bo'shliqlar yig'iladi. Shu sabab "O'ktam",
    "Oʻktam", "Oktam" va "Ўктам" bir xil kalit beradi."""
    if has_cyrillic(text):
        text = cyr_to_lat(text)
    return " ".join(strip_apostrophes(text).split()).casefold()


def _has_letters(text: str) -> bool:
    """Matnda kamida bitta harf bor-yo'qligini tekshiradi. Talabaning ismi
    maxfiylik sababli faqat raqam (ID) bilan almashtirilgan bo'lishi mumkin —
    bunday holatni "haqiqiy ism emas" deb aniqlash uchun ishlatiladi."""
    return any(ch.isalpha() for ch in text)


def generate_query_variants(query: str) -> list[str]:
    """Qidiruv matnidagi apostrofni mumkin bo'lgan barcha belgilarga almashtirib
    variantlar ro'yxatini tuzadi (aynan yozilgan holat birinchi, apostrofsiz
    variant oxirida). Agar matnda apostrofga o'xshash belgi bo'lmasa, faqat asl
    matnni qaytaradi (ortiqcha so'rov yubormaslik uchun)."""
    if not any(ch in query for ch in APOSTROPHE_VARIANTS):
        return [query]

    variants: list[str] = []
    seen: set[str] = set()

    def add(v: str) -> None:
        if v not in seen:
            seen.add(v)
            variants.append(v)

    add(query)
    for candidate_char in APOSTROPHE_VARIANTS:
        variant = query
        for ch in APOSTROPHE_VARIANTS:
            variant = variant.replace(ch, candidate_char)
        add(variant)
    add(strip_apostrophes(query))
    return variants


def _ambiguity_insertions(word: str) -> list[str]:
    """Foydalanuvchi o' / g' dagi apostrofni tushirib qoldirgan bo'lishi mumkin
    ("Ozod" -> "O'zod"). Har bir o/g dan keyin apostrof qo'yilgan variantlarni qaytaradi."""
    out: list[str] = []
    for i, ch in enumerate(word):
        if ch.lower() in ("o", "g") and (i + 1 >= len(word) or word[i + 1] not in APOSTROPHE_VARIANTS):
            for a in INSERT_CHARS:
                out.append(word[:i + 1] + a + word[i + 1:])
    return out[:12]


def _phrase_candidates(query: str) -> list[str]:
    """Bitta yozuv (lotin yoki kirill) uchun so'z tartibi variantlarini tuzadi."""
    words = query.split()
    if not words:
        return [query]

    candidates: list[str] = []

    def add(phrase: str) -> None:
        if phrase and phrase not in candidates:
            candidates.append(phrase)

    # 1) Foydalanuvchi aynan yozgan tartibda (ko'pincha "Familiya Ism" tartibida
    #    yoziladi — bu bazadagi tartibga to'g'ridan-to'g'ri mos keladi)
    add(" ".join(words))

    n = len(words)
    if n == 2:
        # "Ism Familiya" tartibida yozilgan bo'lishi ham mumkin
        add(f"{words[1]} {words[0]}")
    elif n == 3:
        # Familiya + Ism + Otasining ismi — barcha tartiblarni va ikkitalik
        # kombinatsiyalarni (masalan faqat Ism+Otasi yoki Familiya+Otasi) sinaymiz
        for perm in permutations(words):
            add(" ".join(perm))
        for i in range(3):
            for j in range(3):
                if i != j:
                    add(f"{words[i]} {words[j]}")
    elif n > 3:
        add(" ".join(reversed(words)))
        for i in range(n - 1):
            add(f"{words[i]} {words[i + 1]}")

    # Eng oxirgi, eng "kengroq" variant: faqat eng uzun (odatda familiya) so'z
    add(max(words, key=len))
    return candidates


def build_search_candidates(query: str) -> list[str]:
    """Foydalanuvchi yozgan so'zlardan bir nechta so'rov variantini tuzadi.

    🔎 UWED bazasida ism "Familiya Ism Otasining ismi" tartibida saqlanadi va
    qidiruv faqat oddiy substring (aynan mos matn) ishlaydi. Foydalanuvchi esa
    "Familiya Ism", "Ism Familiya", "Ism Otasi Familiya" va hokazo har xil
    tartibda yozishi mumkin — shu sabab bir nechta so'z tartibini navbat bilan
    sinaymiz, ENG ANIQ (ko'p so'zli) variantlardan boshlab. Server o'zi HAQIQIY
    (maxfiy bo'lmagan) ism ustida qidirgani uchun talaba ismi javobda ID bilan
    almashtirilgan bo'lsa ham bu qidiruvga xalaqit bermaydi.

    Agar so'rov kirill harflarida bo'lsa, uning lotincha varianti ham (navbat
    bilan) sinaladi."""
    original = _phrase_candidates(query)
    if not has_cyrillic(query):
        return original

    latin = _phrase_candidates(cyr_to_lat(query))
    merged: list[str] = []
    for a, b in zip_longest(original, latin):
        for c in (a, b):
            if c and c not in merged:
                merged.append(c)
    return merged


def entity_base_name(item: dict) -> str:
    """API javobidagi asosiy ism maydonini qaytaradi (qo'shimcha ID/guruhsiz)."""
    for key in ("full_name", "fio", "name", "student_name", "teacher_name", "fish", "title"):
        val = item.get(key)
        if val:
            return str(val)
    return UNKNOWN_NAME


def entity_display_name(item: dict) -> str:
    """Tugmada ko'rsatiladigan nom. Agar ID/guruh maydoni ham bo'lsa, bir xil
    ismli talabalarni ajratish uchun qavs ichida qo'shib ko'rsatadi."""
    name = entity_base_name(item)

    for key in ("hemis_id", "student_id", "id_number", "student_number", "group", "group_name", "user_id"):
        val = item.get(key)
        if val:
            name = f"{name} ({val})"
            break

    return name


def _dedupe_by_id(items: list[dict]) -> list[dict]:
    seen: set = set()
    out: list[dict] = []
    for it in items:
        key = it.get("id")
        if key is not None:
            if key in seen:
                continue
            seen.add(key)
        out.append(it)
    return out


async def _fetch_search(url: str, term: str) -> list[dict]:
    """Bitta qidiruv so'rovini yuboradi. Xato bo'lsa bo'sh ro'yxat qaytaradi."""
    try:
        async with http_session.get(
            url,
            params={"search": term},
            headers=REQUEST_HEADERS,
            timeout=aiohttp.ClientTimeout(total=10),
        ) as resp:
            if resp.status != 200:
                return []
            data = await resp.json(content_type=None)
    except Exception as e:
        print(f"Qidirishda xato: {e} (term={term!r})")
        return []

    if data.get("success") and data.get("data"):
        return list(data["data"])
    return []


async def _try_api_search(url: str, phrase: str) -> list[dict]:
    """Bitta so'rov matnining barcha apostrof variantlarini BIR VAQTDA (parallel)
    API'ga yuboradi va topilgan natijalarni birlashtirib qaytaradi."""
    variants = generate_query_variants(phrase)
    batches = await asyncio.gather(*(_fetch_search(url, v) for v in variants))
    merged: list[dict] = []
    for b in batches:
        merged.extend(b)
    return _dedupe_by_id(merged)


def _refine_results(found: list[dict], words: list[str], matched_phrase: str) -> list[dict]:
    """1-bosqichda topilgan natijalarni qolgan so'zlar bo'yicha saralaydi.
    Server ikki va undan ortiq so'z bo'yicha aniq moslashtirgan bo'lsa — ishonamiz.
    Faqat bitta so'z (masalan familiya) bilan topilgan bo'lsa, qolgan so'zlar
    ism ichida (apostrofsiz solishtirib) bor-yo'qligi tekshiriladi. Ismi ID bilan
    almashtirilgan (maxfiy) yozuvlarni tekshirib bo'lmaydi — saqlab qolinadi."""
    matched_keys = {_name_key(w) for w in matched_phrase.split()}
    if len(words) <= 1 or len(matched_keys) >= 2:
        return found

    other_keys = [k for k in (_name_key(w) for w in words) if k not in matched_keys]
    if not other_keys:
        return found

    out: list[dict] = []
    for item in found:
        base = entity_base_name(item)
        if not _has_letters(base):
            out.append(item)
            continue
        key = _name_key(base)
        if all(k in key for k in other_keys):
            out.append(item)
    return out


async def _loose_search(urls: list[str], words: list[str], tried: set[str]) -> list[dict]:
    """2- va 3-bosqich: aniq ibora topilmaganda "yumshoq" qidiruv.
    a) har bir so'zdan o/g harflari va apostroflarsiz eng uzun bo'lakni olib qidiramiz
       ("Oktam" -> "ktam": bu "O'ktam" ni ham, "Oktam" ni ham topadi);
    b) hech narsa topilmasa — o/g dan keyin apostrof qo'shib ko'ramiz ("Ozod" -> "O'zod").
    Natija ismlarni apostrofsiz solishtirib saralanadi."""
    latin_words = [cyr_to_lat(w) if has_cyrillic(w) else w for w in words]
    keys = [_name_key(w) for w in words]

    collected: list[dict] = []

    terms: list[str] = []
    for w in latin_words:
        pieces = [p for p in _AMBIG_SPLIT_RE.split(w) if len(p) >= 3]
        if not pieces:
            continue
        term = max(pieces, key=len)
        if term.casefold() not in tried and term not in terms:
            terms.append(term)
    for term in terms[:3]:
        for url in urls:
            res = await _try_api_search(url, term)
            if res:
                collected.extend(res)
                break

    if not collected:
        insert_terms: list[str] = []
        for w in latin_words[:2]:
            for v in _ambiguity_insertions(w):
                if v not in insert_terms:
                    insert_terms.append(v)
        if insert_terms:
            for url in urls:
                batches = await asyncio.gather(*(_fetch_search(url, t) for t in insert_terms))
                for b in batches:
                    collected.extend(b)
                if collected:
                    break

    collected = _dedupe_by_id(collected)
    good: list[dict] = []
    masked: list[dict] = []
    for item in collected:
        base = entity_base_name(item)
        if not _has_letters(base):
            masked.append(item)
            continue
        key = _name_key(base)
        if all(k in key for k in keys):
            good.append(item)

    if good:
        return good
    # Ismi yashirilgan (ID) yozuvlarni faqat bitta so'zli so'rovda va cheklangan miqdorda qaytaramiz
    return masked[:8] if len(words) == 1 else []


async def search_entities(entity_type: str, query: str) -> list[dict]:
    """Talaba ('students'), o'qituvchi ('teachers') yoki guruh ('groups') qidiradi.

    ✅ Talaba manzili DevTools orqali tasdiqlangan:
       https://manage.uwed.uz/api/education/public/v1/student/search/for-timetable?search=...
       Bu javobda ("full_name" maydonida) haqiqiy ism o'rniga talabaning ID raqami
       chiqishi mumkin — bu maxfiylik uchun UWED tomonidan shunday qilingan, xato emas.

    ✅ O'qituvchi manzili ham DevTools orqali tasdiqlangan:
       https://manage.uwed.uz/api/timetable/public/v1/search/teachers?search=...
       — bu xona qidirish (/search/rooms) bilan bir xil API_BASE ostida, EDUCATION_API_BASE'da emas.

    ⚠️ GURUH manzili DevTools orqali TASDIQLANMAGAN (taxminiy). Bir nechta
    ehtimoliy manzil ketma-ket sinaladi — birinchi natija bergani ishlatiladi.
    Agar guruh qidiruvi ishlamasa, brauzer DevTools (Network tab) orqali haqiqiy
    manzilni topib, GROUP_SEARCH_URLS ro'yxatiga qo'shish kerak.

    🎯 Qidiruv bosqichlari (o' / g' harflari, so'z tartibi va kirill yozuviga chidamli):
       1) so'z tartibining har bir varianti, barcha apostrof turlari bilan (parallel);
       2) "yumshoq" qidiruv: apostrof/o/g harflarisiz eng uzun bo'lak bo'yicha;
       3) o/g dan keyin apostrof qo'shib ko'rish.
    Natijalar ismlarni apostrofsiz solishtirib saralanadi, shuning uchun
    "O'ktam", "Oʻktam", "Oktam" va "Ўктам" bir xil odamni topadi."""
    if entity_type == "students":
        urls = [f"{EDUCATION_API_BASE}/student/search/for-timetable"]
    elif entity_type == "teachers":
        urls = [f"{API_BASE}/search/teachers"]
    else:  # "groups" — ⚠️ tasdiqlanmagan, bir nechta variant sinaladi
        urls = GROUP_SEARCH_URLS

    query = " ".join((query or "").split())
    words = query.split()
    if not words:
        return []

    tried: set[str] = set()

    # --- 1-bosqich: aniq ibora variantlari ---
    for phrase in build_search_candidates(query):
        tried.add(phrase.casefold())
        found: list[dict] = []
        for url in urls:
            found = await _try_api_search(url, phrase)
            if found:
                break
        if not found:
            continue
        refined = _refine_results(found, words, phrase)
        if refined:
            return refined

    # --- 2- va 3-bosqich: yumshoq qidiruv ---
    return await _loose_search(urls, words, tried)


def get_event_room(event: dict) -> str:
    """Dars voqeasidan xona nomini chiqaradi. ✅ Asosiy maydon "room_number" (DevTools
    orqali tasdiqlangan), qolganlari ehtiyot uchun zaxira variantlar."""
    for key in ("room_number", "room_name", "room", "auditorium", "auditorium_name", "cabinet", "classroom"):
        val = event.get(key)
        if val:
            return str(val).replace('"', "")
    return "Noma'lum"


# 🔤 Qisqa "code" harfi ('s'/'t'/'g') <-> to'liq entity_type/label/API event turi
# xaritasi. Talaba, o'qituvchi va guruh bilan bog'liq barcha callback'lar shu
# uchta harfdan birini ishlatadi — kodni bir joyda saqlash xatoni kamaytiradi.
CODE_TO_ENTITY_TYPE = {"s": "students", "t": "teachers", "g": "groups"}
CODE_TO_LABEL = {"s": "Talaba", "t": "O'qituvchi", "g": "Guruh"}
CODE_TO_EVENT_TYPE = {"s": "STUDENT", "t": "TEACHER", "g": "GROUP"}


def build_person_schedule_text(events: list[dict], date_str: str, day_name: str, entity_type: str) -> str:
    """Talaba yoki o'qituvchining bir kunlik to'liq jadvalini matn qilib tuzadi."""
    day_events = [e for e in events if str(e.get("date", "")) == date_str]
    day_events.sort(key=lambda e: str(e.get("start_time", "")))

    header = f"📅 <b>{date_str} ({day_name})</b> uchun dars jadvali\n\n"
    if not day_events:
        return header + "✅ Bu kunda darslar yo'q."

    text = header
    for e in day_events:
        para_id = START_TIME_TO_PARA.get(str(e.get("start_time", "")), "?")
        text += f"⏱ <b>{para_id}-para</b> ({e.get('start_time', '')} - {e.get('end_time', '')})\n"
        text += f"📚 Fan: {e.get('title') or 'Nomsiz'}\n"
        if e.get("lesson_name"):
            text += f"📝 Mavzu: {e['lesson_name']}\n"
        if entity_type in ("students", "groups"):
            instructor = e.get('instructor') or "Noma'lum"
            text += f"👨‍🏫 O'qituvchi: {instructor}\n"
        groups = (e.get("groups") or "").replace('"', "")
        if groups:
            text += f"👥 Guruh: {groups}\n"
        text += f"🚪 Xona: {get_event_room(e)}\n"
        if e.get("type"):
            text += f"📖 Turi: {e['type']}\n"
        if e.get("lesson_format"):
            text += f"🖥 Format: {e['lesson_format']}\n"
        text += "\n"
    return text


async def get_room_day_events(block: str, room: str, date_str: str) -> dict[str, dict] | None:
    """Bitta xonaning berilgan kundagi darslarini para raqami bo'yicha qaytaradi:
    {"1": voqea, "3": voqea, ...}. Xona UWED'da topilmasa None qaytaradi."""
    room_id = await get_room_id(block, room)
    if not room_id:
        return None
    target = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    events = await fetch_week_events(room_id, target)
    by_para: dict[str, dict] = {}
    for e in events:
        if str(e.get("date", "")) == date_str:
            p = START_TIME_TO_PARA.get(str(e.get("start_time", "")))
            if p and p not in by_para:
                by_para[p] = e
    return by_para


async def build_room_day_text(block: str, room: str, date_str: str) -> str:
    """Xonaning kunlik holatini (HTML) matn qilib tuzadi — inline qidiruv uchun."""
    display = format_room_display(block, room)
    by_para = await get_room_day_events(block, room, date_str)
    if by_para is None:
        return f"🚪 <b>{display}</b>\n❔ Bu xona UWED tizimida topilmadi."

    target = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    wd = target.isoweekday()
    day_name = HAFTA_KUNLARI[wd - 1] if wd != 7 else "Yakshanba"

    now = now_tashkent()
    current_para = None
    if date_str == now.date().isoformat():
        hm = now.strftime("%H:%M")
        for pid, (st, et, _disp) in PARALAR.items():
            if st <= hm <= et:
                current_para = pid
                break

    lines = [f"🚪 <b>{display}</b> — {date_str} ({day_name})", ""]
    for pid, (st, et, _disp) in PARALAR.items():
        mark = "👉 " if pid == current_para else ""
        e = by_para.get(pid)
        if e:
            who = f" — {esc(e['instructor'])}" if e.get("instructor") else ""
            lines.append(f"{mark}🔴 {pid}-para ({st}-{et}): {esc(e.get('title') or 'fan nomsiz')}{who}")
        else:
            lines.append(f"{mark}🟢 {pid}-para ({st}-{et}): bo'sh")
    return "\n".join(lines)


async def get_room_free_streaks(
    rooms: list[str], block: str, date_str: str
) -> list[tuple[str, list[int], int]]:
    """Har bir xona uchun bir kun davomida ketma-ket bo'sh paralar ketma-ketligini hisoblaydi.
    Qaytaradi: [(xona, [para_raqamlari], uzunlik), ...] eng uzunidan boshlab saralangan,
    faqat kamida 2 ta ketma-ket bo'sh parasi bor xonalar."""
    target_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    all_para_ids = sorted(int(p) for p in PARALAR)
    results: list[tuple[str, list[int], int]] = []

    for room in rooms:
        room_id = await get_room_id(block, room)
        if not room_id:
            continue

        events = await fetch_week_events(room_id, target_date)
        busy_paras: set[int] = set()
        for e in events:
            if str(e.get("date", "")) == date_str:
                p = START_TIME_TO_PARA.get(str(e.get("start_time", "")))
                if p:
                    busy_paras.add(int(p))

        free_paras = [p for p in all_para_ids if p not in busy_paras]
        if not free_paras:
            continue

        best_run: list[int] = []
        current_run = [free_paras[0]]
        for p in free_paras[1:]:
            if p == current_run[-1] + 1:
                current_run.append(p)
            else:
                if len(current_run) > len(best_run):
                    best_run = current_run
                current_run = [p]
        if len(current_run) > len(best_run):
            best_run = current_run

        if len(best_run) >= 2:
            results.append((room, best_run, len(best_run)))

    results.sort(key=lambda x: -x[2])
    return results


def build_streak_text(
    results: list[tuple[str, list[int], int]], block: str, floor: str, date_str: str
) -> str:
    header = (
        f"⏳ <b>KETMA-KET BO'SH XONALAR</b>\n"
        f"🏢 {block} blok, {floor_label(floor)}\n"
        f"📅 {date_str}\n\n"
    )
    if not results:
        return header + "❌ Hech qaysi xonada 2 va undan ortiq ketma-ket bo'sh para topilmadi."

    text = header
    for room, run, length in results:
        display = format_room_display(block, room)
        paras_str = "-".join(str(p) for p in run)
        text += f"🟢 <b>{display}</b>: {paras_str}-paralar bo'sh ({length} ta ketma-ket)\n"
    return text


async def get_room_free_streaks_all_blocks(date_str: str) -> list[tuple[str, str, list[int], int]]:
    """Universitetning BARCHA bloklaridagi (A, B, D, F) barcha xonalarni tekshirib,
    har biri uchun ketma-ket bo'sh paralarni hisoblaydi.
    Qaytaradi: [(blok, xona, [paralar], uzunlik), ...] eng uzunidan boshlab saralangan."""
    all_results: list[tuple[str, str, list[int], int]] = []
    for block in UNIVERSITY_MAP:
        rooms = get_rooms_for_floor(block, "ALL")
        block_results = await get_room_free_streaks(rooms, block, date_str)
        for room, run, length in block_results:
            all_results.append((block, room, run, length))
    all_results.sort(key=lambda x: -x[3])
    return all_results


def build_streak_all_text(results: list[tuple[str, str, list[int], int]], date_str: str) -> str:
    header = f"🌟 <b>ENG UZUN KETMA-KET BO'SH XONALAR — BARCHA BLOKLAR</b>\n📅 {date_str}\n\n"
    if not results:
        return header + "❌ Hech qaysi blokda 2 va undan ortiq ketma-ket bo'sh para topilmadi."

    MAX_PER_FLOOR = 6
    medal_map = {(b, r): m for (b, r, _, _), m in zip(results[:3], ["🥇", "🥈", "🥉"])}

    grouped: dict[str, dict[str, list[tuple[str, list[int], int]]]] = {}
    for block, room, run, length in results:
        floor = get_room_floor(block, room)
        grouped.setdefault(block, {}).setdefault(floor, []).append((room, run, length))

    text = header
    for block in UNIVERSITY_MAP:
        if block not in grouped:
            continue
        text += f"🏢 <b>{block} blok</b>\n"
        floors = grouped[block]
        for floor in sorted(floors.keys(), key=lambda f: int(f) if f.isdigit() else 999):
            entries = sorted(floors[floor], key=lambda x: -x[2])
            text += f"\n<u>{floor}-qavat:</u>\n"
            for room, run, length in entries[:MAX_PER_FLOOR]:
                display = format_room_display(block, room)
                paras_str = "-".join(str(p) for p in run)
                prefix = medal_map.get((block, room), "🟢")
                text += f"{prefix} <b>{display}</b>: {paras_str}-paralar bo'sh ({length} ta)\n"
            extra = len(entries) - MAX_PER_FLOOR
            if extra > 0:
                text += f"<i>...va yana {extra} ta xona</i>\n"
        text += "\n"
    return text


# ============================================================
# 📏 UZUN XABARLARNI XAVFSIZ YUBORISH
# Telegram bitta xabarga 4096 belgidan ortiq matn qo'ymaydi
# (MESSAGE_TOO_LONG xatosi shundan). Shu sabab natija matni
# uzun bo'lsa, bir nechta xabarga bo'lib yuboramiz.
# ============================================================

TELEGRAM_SAFE_LIMIT = 3900


def split_html_text(text: str, limit: int = TELEGRAM_SAFE_LIMIT) -> list[str]:
    """Matnni qator chegaralaridan foydalanib bo'laklarga bo'ladi
    (shu bilan HTML teglari yoki bitta xona ma'lumoti o'rtadan uzilib qolmaydi)."""
    if len(text) <= limit:
        return [text]

    lines = text.split("\n")
    chunks: list[str] = []
    current = ""
    for line in lines:
        candidate = f"{current}\n{line}" if current else line
        if len(candidate) > limit:
            if current:
                chunks.append(current)
            # Agar bitta qatorning o'zi limitdan uzun bo'lsa (kamdan-kam holat), majburan kesamiz
            if len(line) > limit:
                for i in range(0, len(line), limit):
                    chunks.append(line[i:i + limit])
                current = ""
            else:
                current = line
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks


async def send_result(
    callback_query: types.CallbackQuery, text: str, reply_markup: InlineKeyboardMarkup
) -> None:
    """Natija matnini yuboradi. Agar 4096 belgidan uzun bo'lsa, avtomatik bir nechta
    xabarga bo'lib yuboradi (MESSAGE_TOO_LONG xatosining oldini olish uchun).
    Tugmalar faqat oxirgi bo'lakka biriktiriladi."""
    chunks = split_html_text(text)

    if len(chunks) == 1:
        await callback_query.message.edit_text(
            text=chunks[0], reply_markup=reply_markup, parse_mode=ParseMode.HTML
        )
        return

    await callback_query.message.edit_text(text=chunks[0], parse_mode=ParseMode.HTML)
    for chunk in chunks[1:-1]:
        await callback_query.message.answer(text=chunk, parse_mode=ParseMode.HTML)
    await callback_query.message.answer(
        text=chunks[-1], reply_markup=reply_markup, parse_mode=ParseMode.HTML
    )


async def get_room_statuses(
    rooms_in_floor: list[str], block: str, date_str: str, start_t: str
) -> tuple[dict[str, dict], set[str]]:
    """Berilgan qavatdagi xonalarni tekshiradi.
    Qaytaradi: (band_xonalar -> dars ma'lumoti, mavjud_bolmagan_xonalar to'plami).
    Qolgan xonalar - bo'sh xonalar hisoblanadi."""
    target_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    busy: dict[str, dict] = {}
    invalid: set[str] = set()

    for room in rooms_in_floor:
        room_id = await get_room_id(block, room)
        if not room_id:
            invalid.add(room)
            continue

        events = await fetch_week_events(room_id, target_date)
        for event in events:
            if str(event.get("date", "")) == date_str and str(event.get("start_time", "")) == start_t:
                busy[room] = {
                    "title": event.get("title") or "Fan nomi ko'rsatilmagan",
                    "lesson_name": event.get("lesson_name") or "",
                    "instructor": event.get("instructor") or "Noma'lum",
                    "groups": (event.get("groups") or "").replace('"', "") or None,
                    "type": event.get("type") or "",
                    "lesson_format": event.get("lesson_format") or "",
                    "end_time": event.get("end_time") or "",
                }
                break

    return busy, invalid


def build_result_text(
    rooms_in_floor: list[str],
    busy_rooms: dict[str, dict],
    invalid_rooms: set[str],
    block: str,
    header: str,
    floor: str = "1",
) -> str:
    if not rooms_in_floor:
        return header + "\n\nℹ️ Bu qavatda ro'yxatga olingan xona yo'q."

    free_rooms_raw = [
        r for r in rooms_in_floor if r not in busy_rooms and r not in invalid_rooms
    ]

    text = header + "\n\n"

    if free_rooms_raw:
        text += "🟢 <b>BO'SH XONALAR:</b>\n\n"
        if floor == "ALL":
            # Butun blok bo'yicha qidirilganda, bo'sh xonalarni qavat bo'yicha guruhlab ko'rsatamiz
            by_floor: dict[str, list[str]] = {}
            for r in free_rooms_raw:
                fl = get_room_floor(block, r)
                by_floor.setdefault(fl, []).append(format_room_display(block, r))
            for fl in sorted(by_floor.keys()):
                text += f"<b>{fl}-qavat:</b> " + ", ".join(by_floor[fl]) + "\n"
            text += "\n"
        else:
            for idx, r in enumerate(free_rooms_raw, 1):
                text += f"{idx}. 🟢 <b>{format_room_display(block, r)}</b>\n"
    else:
        text += "❌ <b>Barcha xonalar band!</b>\n"

    if busy_rooms:
        text += "\n🔴 <b>BAND XONALAR:</b>\n\n"
        for room, info in busy_rooms.items():
            display_room = format_room_display(block, room)
            text += f"🔴 <b>{display_room}</b>\n"
            text += f"📚 Fan: {info['title']}\n"
            if info["lesson_name"]:
                text += f"📝 Mavzu: {info['lesson_name']}\n"
            text += f"👨‍🏫 O'qituvchi: {info['instructor']}\n"
            if info["groups"]:
                text += f"👥 Guruh: {info['groups']}\n"
            if info["type"]:
                text += f"📖 Turi: {info['type']}\n"
            if info["end_time"]:
                text += f"⏰ Tugash vaqti: {info['end_time']}\n"
            text += "\n"

    if invalid_rooms:
        text += f"<i>ℹ️ {len(invalid_rooms)} ta xona UWED tizimida topilmadi va ro'yxatdan chiqarib tashlandi.</i>"

    return text


# ============================================================
# 📥 EXCEL / PDF EKSPORT
# Talaba/o'qituvchi/guruh kunlik jadvalini yoki xona holatini fayl
# sifatida yuklab olish imkoniyati.
# ============================================================

def _pdf_safe(text: str) -> str:
    """PDF asosiy (core) shriftlari faqat Latin-1 belgilarni qo'llab-quvvatlaydi.
    O'zbekcha tutuq belgisi (ʻ, ʼ, ’ va h.k.) kabi belgilarni oddiy apostrofga
    almashtiradi, kirill harflarini esa lotinga o'giradi — shunda PDF'da
    "?" belgilari chiqib qolmaydi."""
    if has_cyrillic(text):
        text = cyr_to_lat(text)
    for ch in APOSTROPHE_VARIANTS:
        text = text.replace(ch, "'")
    return text.encode("latin-1", errors="replace").decode("latin-1")


def build_person_schedule_excel(events: list[dict], date_str: str, entity_label: str, entity_type: str) -> bytes:
    day_events = [e for e in events if str(e.get("date", "")) == date_str]
    day_events.sort(key=lambda e: str(e.get("start_time", "")))

    wb = Workbook()
    ws = wb.active
    ws.title = "Jadval"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    ws.append([f"{entity_label} — {date_str} kunlik jadvali"])
    ws.merge_cells("A1:F1")
    ws["A1"].font = Font(bold=True, size=13)

    columns = ["Para", "Boshlanish", "Tugash", "Fan", "Xona"]
    if entity_type in ("students", "groups"):
        columns.append("O'qituvchi")
    columns.append("Guruh")
    ws.append(columns)
    for col_idx in range(1, len(columns) + 1):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for e in day_events:
        para_id = START_TIME_TO_PARA.get(str(e.get("start_time", "")), "?")
        row = [
            para_id,
            e.get("start_time", ""),
            e.get("end_time", ""),
            e.get("title") or "Nomsiz",
            get_event_room(e),
        ]
        if entity_type in ("students", "groups"):
            row.append(e.get("instructor") or "Noma'lum")
        row.append((e.get("groups") or "").replace('"', ""))
        ws.append(row)

    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 40)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_person_schedule_pdf(events: list[dict], date_str: str, entity_label: str, entity_type: str) -> bytes:
    day_events = [e for e in events if str(e.get("date", "")) == date_str]
    day_events.sort(key=lambda e: str(e.get("start_time", "")))

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.multi_cell(0, 8, _pdf_safe(f"{entity_label} - {date_str} kunlik jadvali"))
    pdf.ln(2)

    if not day_events:
        pdf.set_font("Helvetica", "", 12)
        pdf.multi_cell(0, 8, "Bu kunda darslar yoq.")
    else:
        for e in day_events:
            para_id = START_TIME_TO_PARA.get(str(e.get("start_time", "")), "?")
            pdf.set_font("Helvetica", "B", 11)
            pdf.multi_cell(0, 7, _pdf_safe(f"{para_id}-para  {e.get('start_time','')} - {e.get('end_time','')}"))
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 6, _pdf_safe(f"Fan: {e.get('title') or 'Nomsiz'}"))
            pdf.multi_cell(0, 6, _pdf_safe(f"Xona: {get_event_room(e)}"))
            if entity_type in ("students", "groups") and e.get("instructor"):
                pdf.multi_cell(0, 6, _pdf_safe(f"O'qituvchi: {e['instructor']}"))
            groups = (e.get("groups") or "").replace('"', "")
            if groups:
                pdf.multi_cell(0, 6, _pdf_safe(f"Guruh: {groups}"))
            pdf.ln(2)

    return bytes(pdf.output())


def build_room_status_excel(
    rooms_in_floor: list[str], busy_rooms: dict[str, dict], invalid_rooms: set[str],
    block: str, floor: str, title: str,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Xonalar"

    ws.append([title])
    ws.merge_cells("A1:E1")
    ws["A1"].font = Font(bold=True, size=13)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    columns = ["Xona", "Holati", "Fan", "O'qituvchi", "Guruh"]
    ws.append(columns)
    for col_idx in range(1, len(columns) + 1):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font

    for room in rooms_in_floor:
        display = format_room_display(block, room)
        if room in busy_rooms:
            info = busy_rooms[room]
            ws.append([display, "🔴 Band", info.get("title", ""), info.get("instructor", ""), info.get("groups") or ""])
        elif room in invalid_rooms:
            ws.append([display, "❔ Topilmadi", "", "", ""])
        else:
            ws.append([display, "🟢 Bo'sh", "", "", ""])

    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 40)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_room_status_pdf(
    rooms_in_floor: list[str], busy_rooms: dict[str, dict], invalid_rooms: set[str],
    block: str, floor: str, title: str,
) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.multi_cell(0, 8, _pdf_safe(title))
    pdf.ln(2)

    for room in rooms_in_floor:
        display = format_room_display(block, room)
        if room in busy_rooms:
            info = busy_rooms[room]
            pdf.set_font("Helvetica", "B", 11)
            pdf.multi_cell(0, 7, _pdf_safe(f"{display} - BAND"))
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 6, _pdf_safe(f"Fan: {info.get('title','')}"))
            if info.get("instructor"):
                pdf.multi_cell(0, 6, _pdf_safe(f"O'qituvchi: {info['instructor']}"))
        elif room in invalid_rooms:
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 6, _pdf_safe(f"{display} - topilmadi"))
        else:
            pdf.set_font("Helvetica", "B", 11)
            pdf.multi_cell(0, 7, _pdf_safe(f"{display} - BO'SH"))
        pdf.ln(1)

    return bytes(pdf.output())


def build_main_menu(is_admin: bool = False) -> InlineKeyboardMarkup:
    buttons = []
    webapp_url = (os.environ.get("WEBAPP_URL") or os.environ.get("RENDER_EXTERNAL_URL") or "").strip().rstrip("/")
    if webapp_url.startswith("https://"):
        buttons.append([InlineKeyboardButton(text="✨ UWED SPACE — Ilovani ochish", web_app=WebAppInfo(url=f"{webapp_url}/app"))])
    buttons += [
        [InlineKeyboardButton(text="🏢 A blok", callback_data="block:A")],
        [InlineKeyboardButton(text="🏢 B blok", callback_data="block:B")],
        [InlineKeyboardButton(text="🏢 D blok", callback_data="block:D")],
        [InlineKeyboardButton(text="🏢 F blok", callback_data="block:F")],
        [InlineKeyboardButton(text="🌟 Eng bo'sh xonalar (barcha bloklar)", callback_data="streak_all")],
        [InlineKeyboardButton(text="🎓 Talaba qidirish", callback_data="search:students")],
        [InlineKeyboardButton(text="👨‍🏫 O'qituvchi qidirish", callback_data="search:teachers")],
        [InlineKeyboardButton(text="👥 Guruh bo'yicha qidirish", callback_data="search:groups")],
        [InlineKeyboardButton(text="🤖 AI-yordamchi (savol bering)", callback_data="ai_start")],
        [InlineKeyboardButton(text="⭐ Sevimlilar", callback_data="favorites")],
        [InlineKeyboardButton(text="🔔 Mening obunam", callback_data="mysub")],
    ]
    if is_admin:
        buttons.append([InlineKeyboardButton(text="👑 Admin panel", callback_data="admin_panel")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def build_persistent_keyboard() -> ReplyKeyboardMarkup:
    """Chat pastida doim turadigan Mini App va bot menyusi tugmalari."""
    webapp_url = (os.environ.get("WEBAPP_URL") or os.environ.get("RENDER_EXTERNAL_URL") or "").strip().rstrip("/")
    buttons = []
    if webapp_url.startswith("https://"):
        buttons.append(KeyboardButton(text="📱 UWED SPACE", web_app=WebAppInfo(url=f"{webapp_url}/app")))
    buttons.append(KeyboardButton(text="🤖 Bot menyusi"))
    return ReplyKeyboardMarkup(keyboard=[buttons], resize_keyboard=True, is_persistent=True, input_field_placeholder="Botdan yoki ilovadan foydalaning…")


@dp.message(CommandStart())
async def command_start_handler(message: types.Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(
        "📱 Ilovani pastdagi <b>UWED SPACE</b> tugmasidan oching. Bot xizmatlari uchun <b>Bot menyusi</b>ni bosing.",
        reply_markup=build_persistent_keyboard(),
        parse_mode=ParseMode.HTML
    )
    await message.answer(
        text=f"Salom, {html.bold(message.from_user.full_name)}! 🏫 <b>JIDU Bo'sh xonalar boti</b>ga xush kelibsiz.\n\nJadvalni ko'rish uchun bino/blokni tanlang, yoki talaba/o'qituvchi ismini qidiring:",
        reply_markup=build_main_menu(message.from_user.id == ADMIN_ID),
        parse_mode=ParseMode.HTML
    )


@dp.message(lambda message: message.text == "🤖 Bot menyusi")
async def persistent_bot_menu_handler(message: types.Message) -> None:
    await message.answer(
        "🤖 <b>Bot xizmatlari</b> — kerakli bo‘limni tanlang:",
        reply_markup=build_main_menu(message.from_user.id == ADMIN_ID),
        parse_mode=ParseMode.HTML
    )


# ============================================================
# ❓ YORDAM VA 👑 ADMIN BUYRUQLARI
# ============================================================

@dp.message(Command("help"))
async def command_help_handler(message: types.Message) -> None:
    me = await message.bot.me()
    text = (
        "🤖 <b>JIDU Bo'sh xonalar boti — yordam</b>\n\n"
        "/start — botni qayta ishga tushirish (bosh menyu)\n"
        "/help — shu yordam matni\n\n"
        "Botdan foydalanish uchun bosh menyudagi tugmalarni bosing:\n"
        "🏢 xona qidirish uchun blok va qavatni tanlang\n"
        "🎓👨‍🏫👥 talaba, o'qituvchi yoki guruh ismini qidiring "
        "(o' va g' harflari, kirill yozuvi va so'z tartibi farqi muhim emas)\n"
        "🤖 AI-yordamchi orqali oddiy tilda savol bering (masalan \"A blokda "
        "hozir bo'sh xona bormi?\", \"A305 bugun bandmi?\")\n"
        "⭐ tez-tez qidiriladigan xona yoki odamni sevimlilarga saqlang\n"
        "🔔 o'z jadvalingizga obuna bo'lsangiz, har kuni soat 7:00 da "
        "kunlik jadvalingizni va har darsdan 10 daqiqa oldin eslatma olasiz\n"
        "📥 talaba/o'qituvchi jadvali va xona holatini Excel yoki PDF "
        "fayl sifatida ham yuklab olishingiz mumkin\n"
        f"🔎 istalgan chatda <code>@{me.username} A305</code> deb yozsangiz, "
        "xonaning bugungi holati chiqadi"
    )
    if message.from_user.id == ADMIN_ID:
        text += (
            "\n\n👑 <b>Admin buyruqlari:</b>\n"
            "Bosh menyudagi \"👑 Admin panel\" tugmasi — foydalanuvchilar ro'yxati, qidiruv, "
            "profil kartochkasi, bloklash va xabar yuborish\n"
            "/users — foydalanuvchilar ro'yxati\n"
            "/stats — umumiy statistika\n"
            "/aistats — AI-yordamchi kunlik statistikasi\n"
            "/ban [ID yoki @username] — foydalanuvchini bloklash\n"
            "/unban [ID yoki @username] — foydalanuvchini blokdan chiqarish\n"
            "/broadcast [matn] — barcha foydalanuvchilarga xabar yuborish"
        )
    await message.answer(text, parse_mode=ParseMode.HTML)


@dp.message(Command("stats"))
async def command_stats_handler(message: types.Message) -> None:
    if message.from_user.id != ADMIN_ID:
        return

    total, banned, top = await db_get_stats()
    text = (
        f"📊 <b>Statistika</b>\n\n"
        f"👥 Jami foydalanuvchilar: <b>{total}</b>\n"
        f"🚫 Bloklangan: <b>{banned}</b>\n"
    )
    if top:
        text += "\n🏆 <b>Eng faol foydalanuvchilar:</b>\n"
        for uid, uname, fname, count in top:
            uname_text = f"@{esc(uname)}" if uname else "yo'q"
            text += f"• {esc(fname or UNKNOWN_NAME)} ({uname_text}, ID: <code>{uid}</code>) — {count} ta harakat\n"

    await message.answer(text, parse_mode=ParseMode.HTML)


async def build_ai_stats_text() -> str:
    rows = await db_get_ai_stats_today()
    total_today = sum(r[3] for r in rows)
    text = (
        f"🤖 <b>AI-yordamchi statistikasi (bugun)</b>\n\n"
        f"📈 Eng faol 10 ta foydalanuvchi so'rovlari yig'indisi: <b>{total_today}</b>\n"
        f"🚦 Har bir foydalanuvchi uchun kunlik limit: <b>{AI_DAILY_LIMIT}</b>\n"
    )
    if rows:
        text += "\n🏆 <b>Eng ko'p foydalanganlar:</b>\n"
        for uid, uname, fname, count in rows:
            uname_text = f"@{esc(uname)}" if uname else "yo'q"
            text += f"• {esc(fname or UNKNOWN_NAME)} ({uname_text}, ID: <code>{uid}</code>) — {count} ta so'rov\n"
    else:
        text += "\nBugun hali hech kim AI-yordamchidan foydalanmagan."
    return text


@dp.message(Command("aistats"))
async def command_ai_stats_handler(message: types.Message) -> None:
    if message.from_user.id != ADMIN_ID:
        return
    await message.answer(await build_ai_stats_text(), parse_mode=ParseMode.HTML)


@dp.message(Command("ban"))
async def command_ban_handler(message: types.Message, command: CommandObject) -> None:
    if message.from_user.id != ADMIN_ID:
        return
    if not command.args:
        await message.answer("Foydalanish: <code>/ban [ID yoki @username]</code>", parse_mode=ParseMode.HTML)
        return

    row = await db_find_user(command.args)
    if not row:
        await message.answer("❌ Bunday foydalanuvchi topilmadi (u hali botga murojaat qilmagan bo'lishi mumkin).")
        return

    uid = row[0]
    if uid == ADMIN_ID:
        await message.answer("O'zingizni bloklay olmaysiz 🙂")
        return

    await db_set_banned(uid, True)
    await message.answer(f"🚫 Foydalanuvchi (ID: <code>{uid}</code>) bloklandi.", parse_mode=ParseMode.HTML)


@dp.message(Command("unban"))
async def command_unban_handler(message: types.Message, command: CommandObject) -> None:
    if message.from_user.id != ADMIN_ID:
        return
    if not command.args:
        await message.answer("Foydalanish: <code>/unban [ID yoki @username]</code>", parse_mode=ParseMode.HTML)
        return

    row = await db_find_user(command.args)
    if not row:
        await message.answer("❌ Bunday foydalanuvchi topilmadi.")
        return

    uid = row[0]
    await db_set_banned(uid, False)
    await message.answer(f"✅ Foydalanuvchi (ID: <code>{uid}</code>) blokdan chiqarildi.", parse_mode=ParseMode.HTML)


@dp.message(Command("broadcast"))
async def command_broadcast_handler(message: types.Message, command: CommandObject) -> None:
    if message.from_user.id != ADMIN_ID:
        return
    if not command.args:
        await message.answer("Foydalanish: <code>/broadcast [xabar matni]</code>", parse_mode=ParseMode.HTML)
        return

    text_to_send = command.args
    rows = await db_list_recent_users(100000)  # deyarli barcha foydalanuvchilar
    total = len(rows)
    if total == 0:
        await message.answer("Hali hech kim botdan foydalanmagan, yuboradigan hech kim yo'q.")
        return

    sent = 0
    failed = 0
    status_msg = await message.answer(f"📢 Yuborilmoqda... 0/{total}")

    for idx, (uid, uname, fname, is_banned) in enumerate(rows, 1):
        if is_banned:
            continue
        try:
            await message.bot.send_message(chat_id=uid, text=text_to_send, parse_mode=None)
            sent += 1
        except Exception:
            failed += 1
        if idx % 25 == 0:
            try:
                await status_msg.edit_text(f"📢 Yuborilmoqda... {idx}/{total}")
            except Exception:
                pass
        # Telegram flood-limit'iga tushib qolmaslik uchun kichik tanaffus
        await asyncio.sleep(0.05)

    await status_msg.edit_text(
        f"✅ Yuborish yakunlandi.\n✔️ Muvaffaqiyatli: {sent}\n❌ Yetkazilmadi: {failed}"
    )


# ============================================================
# 👑 ADMIN PANEL — foydalanuvchilar guruh/kanal a'zolari kabi ro'yxat ko'rinishida
#   • umumiy ko'rsatkichlar (jami, bugun yangi, bugun faol, bloklangan, AI...)
#   • sahifalangan foydalanuvchilar ro'yxati (eng so'nggi faol bo'lganlar tepada)
#   • har bir foydalanuvchi uchun profil kartochkasi (ID, username, faollik,
#     obuna, sevimlilar, AI hisobi), profilni ochish, bloklash va unga xabar yozish
#   • ism / @username / ID bo'yicha qidirish
# ============================================================

def _short(text: str, n: int) -> str:
    text = text or ""
    return text if len(text) <= n else text[: n - 1] + "…"


async def _safe_edit(callback_query: types.CallbackQuery, text: str, reply_markup: InlineKeyboardMarkup | None = None) -> None:
    """edit_text, lekin "message is not modified" xatosini e'tiborsiz qoldiradi
    (masalan "Yangilash" tugmasi hech narsa o'zgarmaganda bosilganda)."""
    try:
        await callback_query.message.edit_text(
            text=text, reply_markup=reply_markup, parse_mode=ParseMode.HTML
        )
    except TelegramBadRequest as e:
        if "message is not modified" in str(e).lower():
            return
        raise


async def _require_admin(callback_query: types.CallbackQuery) -> bool:
    if callback_query.from_user.id != ADMIN_ID:
        await callback_query.answer("Ruxsat yo'q.", show_alert=True)
        return False
    return True


async def build_users_page(page: int) -> tuple[str, InlineKeyboardMarkup]:
    rows, total, page = await db_list_users_page(page)
    pages = max((total + USERS_PAGE_SIZE - 1) // USERS_PAGE_SIZE, 1)

    text = (
        f"👥 <b>Foydalanuvchilar</b> — jami <b>{total}</b> ta\n"
        f"📄 Sahifa: <b>{page + 1}/{pages}</b> (eng so'nggi faol bo'lganlar tepada)\n\n"
        "Foydalanuvchi haqida to'liq ma'lumot olish uchun uning ustiga bosing:"
    )
    if total == 0:
        text = "👥 <b>Foydalanuvchilar</b>\n\nHali hech kim botdan foydalanmagan."

    buttons: list[list[InlineKeyboardButton]] = []
    for uid, uname, fname, is_banned, count, last_seen in rows:
        icon = "🚫" if is_banned else "🟢"
        name = _short(fname or str(uid), 22)
        uname_part = f" @{uname}" if uname else ""
        label = f"{icon} {name}{uname_part} · {count}"
        buttons.append([InlineKeyboardButton(text=label, callback_data=f"uinfo:{uid}:{page}")])

    if pages > 1:
        nav: list[InlineKeyboardButton] = []
        if page > 0:
            nav.append(InlineKeyboardButton(text="⬅️ Oldingi", callback_data=f"ulist:{page - 1}"))
        nav.append(InlineKeyboardButton(text=f"{page + 1}/{pages}", callback_data="noop"))
        if page < pages - 1:
            nav.append(InlineKeyboardButton(text="Keyingi ➡️", callback_data=f"ulist:{page + 1}"))
        buttons.append(nav)

    buttons.append([InlineKeyboardButton(text="🔎 Qidirish", callback_data="usearch")])
    buttons.append([InlineKeyboardButton(text="⬅️ Admin panel", callback_data="admin_panel")])
    return text, InlineKeyboardMarkup(inline_keyboard=buttons)


async def build_user_card(uid: int, page: int, with_profile_link: bool = True) -> tuple[str, InlineKeyboardMarkup] | None:
    row = await db_get_user_row(uid)
    if not row:
        return None
    _uid, username, full_name, first_seen, last_seen, count, is_banned = row

    ai_used = await db_get_ai_usage_today(uid)
    favs = await db_count_favorites(uid)
    sub = await db_get_subscription(uid)
    sub_text = esc(sub[2]) if sub else "yo'q"
    uname_text = f"@{esc(username)}" if username else "yo'q"

    text = (
        f"👤 <b>{esc(full_name or UNKNOWN_NAME)}</b>\n\n"
        f"🆔 ID: <code>{uid}</code>\n"
        f"🔗 Username: {uname_text}\n"
        f"📅 Birinchi marta: {fmt_dt(first_seen)}\n"
        f"🕒 Oxirgi faollik: {fmt_dt(last_seen)}\n"
        f"💬 Jami harakatlar: <b>{count}</b>\n"
        f"🤖 Bugungi AI so'rovlari: {ai_used}\n"
        f"⭐ Sevimlilar soni: {favs}\n"
        f"🔔 Obuna: {sub_text}\n"
        f"📌 Holati: {'🚫 Bloklangan' if is_banned else '✅ Faol'}"
    )

    buttons: list[list[InlineKeyboardButton]] = []
    if with_profile_link:
        url = f"https://t.me/{username}" if username else f"tg://user?id={uid}"
        buttons.append([InlineKeyboardButton(text="👤 Profilni ochish", url=url)])
    buttons.append([InlineKeyboardButton(text="💬 Xabar yuborish", callback_data=f"udm:{uid}:{page}")])
    if is_banned:
        buttons.append([InlineKeyboardButton(text="✅ Blokdan chiqarish", callback_data=f"uunban:{uid}:{page}")])
    else:
        buttons.append([InlineKeyboardButton(text="🚫 Bloklash", callback_data=f"uban:{uid}:{page}")])
    buttons.append([InlineKeyboardButton(text="⬅️ Ro'yxatga qaytish", callback_data=f"ulist:{page}")])
    return text, InlineKeyboardMarkup(inline_keyboard=buttons)


async def show_user_card(callback_query: types.CallbackQuery, uid: int, page: int) -> None:
    built = await build_user_card(uid, page, with_profile_link=True)
    if built is None:
        await callback_query.answer("Foydalanuvchi topilmadi.", show_alert=True)
        return
    text, kb = built
    try:
        await _safe_edit(callback_query, text, kb)
    except TelegramBadRequest:
        # tg://user?id=... tugmasi foydalanuvchi maxfiylik sozlamalari tufayli qabul qilinmasligi
        # mumkin (BUTTON_USER_PRIVACY_RESTRICTED) — shunda profil tugmasisiz ko'rsatamiz
        built2 = await build_user_card(uid, page, with_profile_link=False)
        if built2:
            await _safe_edit(callback_query, built2[0], built2[1])


async def build_admin_panel() -> tuple[str, InlineKeyboardMarkup]:
    ov = await db_get_overview()
    text = (
        "👑 <b>Admin panel</b>\n\n"
        f"👥 Jami foydalanuvchilar: <b>{ov['total']}</b>\n"
        f"🆕 Bugun qo'shilganlar: <b>{ov['new_today']}</b>\n"
        f"🟢 Bugun faol: <b>{ov['active_today']}</b>\n"
        f"🚫 Bloklanganlar: <b>{ov['banned']}</b>\n"
        f"🔔 Obunachilar: <b>{ov['subs']}</b>\n"
        f"🤖 Bugungi AI so'rovlari: <b>{ov['ai_today']}</b>"
    )
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="👥 Foydalanuvchilar ro'yxati", callback_data="ulist:0")],
        [InlineKeyboardButton(text="🔎 Foydalanuvchi qidirish", callback_data="usearch")],
        [InlineKeyboardButton(text="🤖 AI statistikasi", callback_data="adm_ai")],
        [InlineKeyboardButton(text="📢 Hammaga xabar yuborish", callback_data="adm_bc")],
        [InlineKeyboardButton(text="🔄 Yangilash", callback_data="admin_panel")],
        [InlineKeyboardButton(text="⬅️ Bosh sahifa", callback_data="restart_search")],
    ])
    return text, kb


@dp.message(Command("users"))
async def command_users_handler(message: types.Message) -> None:
    if message.from_user.id != ADMIN_ID:
        return
    text, kb = await build_users_page(0)
    await message.answer(text, reply_markup=kb, parse_mode=ParseMode.HTML)


@dp.callback_query(lambda c: c.data == "noop")
async def process_noop(callback_query: types.CallbackQuery) -> None:
    await callback_query.answer()


@dp.callback_query(lambda c: c.data == "admin_panel")
async def process_admin_panel(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    if not await _require_admin(callback_query):
        return
    await state.clear()
    text, kb = await build_admin_panel()
    await _safe_edit(callback_query, text, kb)
    await callback_query.answer()


@dp.callback_query(lambda c: c.data == "adm_ai")
async def process_admin_ai(callback_query: types.CallbackQuery) -> None:
    if not await _require_admin(callback_query):
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Admin panel", callback_data="admin_panel")]
    ])
    await _safe_edit(callback_query, await build_ai_stats_text(), kb)
    await callback_query.answer()


@dp.callback_query(lambda c: c.data == "adm_bc")
async def process_admin_broadcast_info(callback_query: types.CallbackQuery) -> None:
    if not await _require_admin(callback_query):
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Admin panel", callback_data="admin_panel")]
    ])
    await _safe_edit(
        callback_query,
        "📢 <b>Hammaga xabar yuborish</b>\n\n"
        "Quyidagi buyruq bilan yuboring:\n"
        "<code>/broadcast Xabar matni</code>\n\n"
        "Xabar bloklanmagan barcha foydalanuvchilarga boradi.\n"
        "Bitta foydalanuvchiga xabar yozish uchun uning kartochkasida "
        "\"💬 Xabar yuborish\" tugmasidan foydalaning (matn, rasm, ovoz, fayl — hammasi yuboriladi).",
        kb,
    )
    await callback_query.answer()


@dp.callback_query(lambda c: c.data.startswith("ulist:"))
async def process_users_list(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    if not await _require_admin(callback_query):
        return
    await state.clear()
    page = int(callback_query.data.split(":")[1])
    text, kb = await build_users_page(page)
    await _safe_edit(callback_query, text, kb)
    await callback_query.answer()


@dp.callback_query(lambda c: c.data.startswith("uinfo:"))
async def process_user_info(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    if not await _require_admin(callback_query):
        return
    await state.clear()
    _, uid_str, page_str = callback_query.data.split(":")
    await show_user_card(callback_query, int(uid_str), int(page_str))
    try:
        await callback_query.answer()
    except Exception:
        pass


@dp.callback_query(lambda c: c.data.startswith("uban:") or c.data.startswith("uunban:"))
async def process_user_ban_toggle(callback_query: types.CallbackQuery) -> None:
    if not await _require_admin(callback_query):
        return
    action, uid_str, page_str = callback_query.data.split(":")
    uid = int(uid_str)
    if action == "uban":
        if uid == ADMIN_ID:
            await callback_query.answer("O'zingizni bloklay olmaysiz 🙂", show_alert=True)
            return
        await db_set_banned(uid, True)
        await callback_query.answer("🚫 Foydalanuvchi bloklandi.")
    else:
        await db_set_banned(uid, False)
        await callback_query.answer("✅ Foydalanuvchi blokdan chiqarildi.")
    await show_user_card(callback_query, uid, int(page_str))


@dp.callback_query(lambda c: c.data == "usearch")
async def process_user_search_start(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    if not await _require_admin(callback_query):
        return
    await state.set_state(AdminStates.waiting_user_search)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Bekor qilish", callback_data="ulist:0")]
    ])
    await _safe_edit(
        callback_query,
        "🔎 <b>Foydalanuvchi qidirish</b>\n\nIsm, @username yoki ID raqamini yozib yuboring:",
        kb,
    )
    await callback_query.answer()


@dp.callback_query(lambda c: c.data.startswith("udm:"))
async def process_user_dm_start(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    if not await _require_admin(callback_query):
        return
    _, uid_str, page_str = callback_query.data.split(":")
    await state.set_state(AdminStates.waiting_dm)
    await state.update_data(target_uid=int(uid_str), back_page=int(page_str))
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Bekor qilish", callback_data=f"uinfo:{uid_str}:{page_str}")]
    ])
    await _safe_edit(
        callback_query,
        f"💬 <b>Xabar yuborish</b>\n\nFoydalanuvchi ID: <code>{uid_str}</code>\n\n"
        "Yubormoqchi bo'lgan xabaringizni yozing (matn, rasm, ovozli xabar, fayl — "
        "istalgan turdagi xabar aynan o'zi yuboriladi):",
        kb,
    )
    await callback_query.answer()


@dp.message(AdminStates.waiting_user_search)
async def process_user_search_query(message: types.Message, state: FSMContext) -> None:
    if message.from_user.id != ADMIN_ID:
        return
    query = (message.text or "").strip()
    if not query:
        await message.answer("Iltimos, ism, @username yoki ID yozing.")
        return

    rows = await db_search_users(query, limit=10)
    await state.clear()

    if not rows:
        await message.answer(
            "❌ Hech kim topilmadi.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🔎 Qayta qidirish", callback_data="usearch")],
                [InlineKeyboardButton(text="⬅️ Admin panel", callback_data="admin_panel")],
            ]),
        )
        return

    buttons = []
    for uid, uname, fname, is_banned, count, last_seen in rows:
        icon = "🚫" if is_banned else "🟢"
        uname_part = f" @{uname}" if uname else ""
        buttons.append([InlineKeyboardButton(
            text=f"{icon} {_short(fname or str(uid), 22)}{uname_part} · {count}",
            callback_data=f"uinfo:{uid}:0",
        )])
    buttons.append([InlineKeyboardButton(text="⬅️ Admin panel", callback_data="admin_panel")])
    await message.answer(
        f"🔎 Topildi: <b>{len(rows)}</b> ta",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode=ParseMode.HTML,
    )


@dp.message(AdminStates.waiting_dm)
async def process_admin_dm_send(message: types.Message, state: FSMContext) -> None:
    if message.from_user.id != ADMIN_ID:
        return
    data = await state.get_data()
    target_uid = data.get("target_uid")
    back_page = data.get("back_page", 0)
    await state.clear()

    if target_uid is None:
        await message.answer("❌ Qabul qiluvchi aniqlanmadi, qaytadan urinib ko'ring.")
        return

    try:
        # copy_message — matn, rasm, ovoz, video, fayl va h.k. ni aynan o'zini yuboradi
        await message.bot.copy_message(
            chat_id=target_uid,
            from_chat_id=message.chat.id,
            message_id=message.message_id,
        )
        result = f"✅ Xabar foydalanuvchiga (ID: <code>{target_uid}</code>) yuborildi."
    except Exception as e:
        result = f"❌ Yuborib bo'lmadi: <code>{esc(e)}</code>\n(Foydalanuvchi botni bloklagan bo'lishi mumkin.)"

    await message.answer(
        result,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="⬅️ Foydalanuvchi kartochkasi", callback_data=f"uinfo:{target_uid}:{back_page}")],
            [InlineKeyboardButton(text="👥 Ro'yxat", callback_data=f"ulist:{back_page}")],
        ]),
        parse_mode=ParseMode.HTML,
    )


# ============================================================
# 🔎 INLINE QIDIRUV — istalgan chatda "@bot_username A305" deb yozilganda
# xonaning bugungi holati chiqadi.
# ⚙️ SOZLASH: @BotFather -> /setinline buyrug'i bilan botingiz uchun inline
# rejimni yoqing (placeholder sifatida masalan "Xona nomi: A305" yozing).
# ============================================================

@dp.inline_query()
async def inline_room_search(inline_query: types.InlineQuery) -> None:
    user = inline_query.from_user
    try:
        if user.id != ADMIN_ID:
            await db_ensure_user(user.id, user.username, user.full_name)
            if await db_is_banned(user.id):
                await inline_query.answer([], cache_time=5, is_personal=True)
                return

        q = (inline_query.query or "").strip()
        if not q:
            hint = InlineQueryResultArticle(
                id="hint",
                title="🔎 Xona nomini yozing",
                description="Masalan: A305, D-412, B607",
                input_message_content=InputTextMessageContent(
                    message_text="🔎 Xona holatini bilish uchun @bot_nomi dan keyin xona nomini yozing, masalan: A305"
                ),
            )
            await inline_query.answer([hint], cache_time=5, is_personal=True)
            return

        matches = find_rooms_by_text(q, limit=5)
        if not matches:
            empty = InlineQueryResultArticle(
                id="none",
                title="❌ Bunday xona topilmadi",
                description="Xona nomini tekshirib qayta yozing (masalan A305)",
                input_message_content=InputTextMessageContent(message_text=f"❌ '{q}' nomli xona topilmadi."),
            )
            await inline_query.answer([empty], cache_time=5, is_personal=True)
            return

        date_str = now_tashkent().date().isoformat()
        texts = await asyncio.gather(*(build_room_day_text(b, r, date_str) for b, r in matches))
        results = []
        for (b, r), text in zip(matches, texts):
            results.append(InlineQueryResultArticle(
                id=f"{b}:{r}:{date_str}",
                title=f"🚪 {format_room_display(b, r)} — bugungi holat",
                description="Bosing — xonaning bugungi jadvali yuboriladi",
                input_message_content=InputTextMessageContent(message_text=text, parse_mode=ParseMode.HTML),
            ))
        await inline_query.answer(results, cache_time=30, is_personal=False)
    except Exception as e:
        print(f"[INLINE] xato: {e}")
        try:
            await inline_query.answer([], cache_time=5, is_personal=True)
        except Exception:
            pass


@dp.callback_query(lambda c: c.data.startswith("block:"))
async def process_block_selection(callback_query: types.CallbackQuery) -> None:
    _, selected_block = callback_query.data.split(":")
    floors = UNIVERSITY_MAP[selected_block]

    floor_buttons = [
        [InlineKeyboardButton(text="📚 Butun blok (barcha qavatlar)", callback_data=f"floor:{selected_block}:ALL")]
    ]
    row = []
    for floor in floors.keys():
        row.append(InlineKeyboardButton(text=f"{floor}-qavat", callback_data=f"floor:{selected_block}:{floor}"))
        if len(row) == 2:
            floor_buttons.append(row)
            row = []
    if row:
        floor_buttons.append(row)

    floor_buttons.append([InlineKeyboardButton(text="⬅️ Ortga", callback_data="restart_search")])
    keyboard = InlineKeyboardMarkup(inline_keyboard=floor_buttons)

    await callback_query.message.edit_text(
        text=f"📍 Siz <b>{selected_block} blokni</b> tanladingiz.\nEndi kerakli qavatni tanlang:",
        reply_markup=keyboard,
        parse_mode=ParseMode.HTML
    )
    await callback_query.answer()


@dp.callback_query(lambda c: c.data.startswith("floor:"))
async def process_floor_selection(callback_query: types.CallbackQuery) -> None:
    _, block, floor = callback_query.data.split(":")
    today = now_tashkent().date()
    user_id = callback_query.from_user.id

    is_fav = await db_is_favorite(user_id, "room", f"{block}:{floor}")
    fav_text = "⭐ Sevimlilardan olib tashlash" if is_fav else "☆ Sevimlilarga saqlash"

    day_buttons = [
        [InlineKeyboardButton(text="⚡️ Hozirgi ayni vaqtdagi bo'sh xonalar", callback_data=f"quick:{block}:{floor}")],
        [InlineKeyboardButton(text="⏳ Ketma-ket bo'sh xonalar (bugun)", callback_data=f"streak:{block}:{floor}")],
        [InlineKeyboardButton(text=fav_text, callback_data=f"fav:room:{block}:{floor}")],
    ]

    for i in range(0, 6):
        next_day = today + datetime.timedelta(days=i)
        weekday_num = next_day.isoweekday()

        if weekday_num != 7:
            day_name = HAFTA_KUNLARI[weekday_num - 1]
            date_str = next_day.strftime('%Y-%m-%d')
            button_text = build_date_label(i, next_day)

            day_buttons.append([InlineKeyboardButton(text=button_text, callback_data=f"day:{block}:{floor}:{date_str}:{day_name}")])

    day_buttons.append([InlineKeyboardButton(text="⬅️ Ortga", callback_data=f"block:{block}")])
    keyboard = InlineKeyboardMarkup(inline_keyboard=day_buttons)

    await callback_query.message.edit_text(
        text=f"🏢 <b>{block} blok, {floor_label(floor)}</b> tanlandi.\n\nSizga hozirgi soniyadagi bo'sh xonalar kerakmi yoki kelajak sanalarimi?",
        reply_markup=keyboard,
        parse_mode=ParseMode.HTML
    )
    await callback_query.answer()


@dp.callback_query(lambda c: c.data.startswith("streak:"))
async def process_streak_search(callback_query: types.CallbackQuery) -> None:
    _, block, floor = callback_query.data.split(":")
    rooms_in_floor = get_rooms_for_floor(block, floor)

    await callback_query.answer("Hisoblanmoqda... bu biroz vaqt olishi mumkin")

    today_str = now_tashkent().date().strftime("%Y-%m-%d")
    results = await get_room_free_streaks(rooms_in_floor, block, today_str)
    result_text = build_streak_text(results, block, floor, today_str)

    restart_markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Ortga", callback_data=f"floor:{block}:{floor}")],
        [InlineKeyboardButton(text="Bosh sahifa 🔄", callback_data="restart_search")]
    ])
    await send_result(callback_query, result_text, restart_markup)


@dp.callback_query(lambda c: c.data == "streak_all")
async def process_streak_all(callback_query: types.CallbackQuery) -> None:
    await callback_query.answer("Hisoblanmoqda... barcha bloklar tekshirilyapti, biroz vaqt oladi")

    today_str = now_tashkent().date().strftime("%Y-%m-%d")
    results = await get_room_free_streaks_all_blocks(today_str)
    result_text = build_streak_all_text(results, today_str)

    restart_markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Bosh sahifa 🔄", callback_data="restart_search")]
    ])
    await send_result(callback_query, result_text, restart_markup)


@dp.callback_query(lambda c: c.data.startswith("quick:"))
async def process_quick_search(callback_query: types.CallbackQuery) -> None:
    _, block, floor = callback_query.data.split(":")
    rooms_in_floor = get_rooms_for_floor(block, floor)

    await callback_query.answer("Qidirilmoqda...")

    now = now_tashkent()
    current_date = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M")

    active_para_start = None
    para_name = "Darsdan tashqari vaqt"

    for para_id, (start_t, end_t, display) in PARALAR.items():
        if start_t <= current_time <= end_t:
            active_para_start = start_t
            para_name = display
            break

    if not active_para_start:
        busy_rooms, invalid_rooms = {}, set()
    else:
        busy_rooms, invalid_rooms = await get_room_statuses(rooms_in_floor, block, current_date, active_para_start)

    header = (
        f"⚡️ <b>TEZKOR QIDIRUV NATIJASI</b>\n"
        f"🏢 <b>{block} blok, {floor_label(floor)}</b>\n"
        f"⏰ Hozirgi vaqt: <b>{current_time}</b> ({para_name})"
    )
    result_text = build_result_text(rooms_in_floor, busy_rooms, invalid_rooms, block, header, floor)

    restart_markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Ortga", callback_data=f"floor:{block}:{floor}")],
        [InlineKeyboardButton(text="Bosh sahifa 🔄", callback_data="restart_search")]
    ])
    await send_result(callback_query, result_text, restart_markup)


@dp.callback_query(lambda c: c.data.startswith("day:"))
async def process_day_selection(callback_query: types.CallbackQuery) -> None:
    _, block, floor, selected_date, day_name = callback_query.data.split(":")
    para_buttons = []
    row = []
    for para_id, para_info in PARALAR.items():
        row.append(InlineKeyboardButton(text=f"⏱ {para_id}-para", callback_data=f"para:{block}:{floor}:{selected_date}:{day_name}:{para_id}"))
        if len(row) == 2:
            para_buttons.append(row)
            row = []
    if row:
        para_buttons.append(row)

    para_buttons.append([InlineKeyboardButton(text="⬅️ Ortga", callback_data=f"floor:{block}:{floor}")])
    keyboard = InlineKeyboardMarkup(inline_keyboard=para_buttons)

    await callback_query.message.edit_text(
        text=f"📅 Sana: <b>{selected_date} ({day_name})</b>\n🏢 Joy: <b>{block} blok, {floor_label(floor)}</b>\n\nQaysi dars soatidagi bo'sh xonalarni ko'rmoqchisiz?",
        reply_markup=keyboard,
        parse_mode=ParseMode.HTML
    )
    await callback_query.answer()


@dp.callback_query(lambda c: c.data.startswith("para:"))
async def process_para_selection(callback_query: types.CallbackQuery) -> None:
    _, block, floor, selected_date, day_name, para_id = callback_query.data.split(":")
    rooms_in_floor = get_rooms_for_floor(block, floor)

    await callback_query.answer("Qidirilmoqda...")

    start_t, end_t, display = PARALAR[para_id]

    busy_rooms, invalid_rooms = await get_room_statuses(rooms_in_floor, block, selected_date, start_t)

    header = (
        f"📋 <b>QIDIRUV NATIJASI</b>\n"
        f"📅 Sana: <b>{selected_date} ({day_name})</b>\n"
        f"🏢 Joy: <b>{block} blok, {floor_label(floor)}</b>\n"
        f"⏰ Dars vaqti: <b>{display}</b>"
    )
    result_text = build_result_text(rooms_in_floor, busy_rooms, invalid_rooms, block, header, floor)

    restart_markup = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="📥 Excel", callback_data=f"xls:r:{block}:{floor}:{selected_date}:{para_id}"),
            InlineKeyboardButton(text="📄 PDF", callback_data=f"pdf:r:{block}:{floor}:{selected_date}:{para_id}"),
        ],
        [InlineKeyboardButton(text="⬅️ Ortga", callback_data=f"day:{block}:{floor}:{selected_date}:{day_name}")],
        [InlineKeyboardButton(text="Bosh sahifa 🔄", callback_data="restart_search")]
    ])

    await send_result(callback_query, result_text, restart_markup)


@dp.callback_query(lambda c: c.data == "restart_search")
async def process_restart(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await callback_query.message.edit_text(
        text="🏫 <b>JIDU Bo'sh xonalar boti</b>\n\nJadvalni ko'rish uchun bino/blokni tanlang, yoki talaba/o'qituvchi ismini qidiring:",
        reply_markup=build_main_menu(callback_query.from_user.id == ADMIN_ID),
        parse_mode=ParseMode.HTML
    )
    await callback_query.answer()


# ============================================================
# 🎓 TALABA / 👨‍🏫 O'QITUVCHI / 👥 GURUH QIDIRUVI
# ============================================================

# 🗂 Talaba/o'qituvchi/guruh ID'sidan to'liq ismini topish uchun tezkor kesh
# (qidiruv natijalarida bir marta ko'rilgach xotirada saqlanadi — ⭐ Sevimlilar
# va 🔔 Obuna funksiyalari uchun chiroyli nom ko'rsatish maqsadida ishlatiladi).
ENTITY_NAME_CACHE: dict[str, str] = {}

SEARCH_LABELS = {
    "students": "talaba ism-familiyasi yoki ID raqami",
    "teachers": "o'qituvchi ism-familiyasi",
    "groups": "guruh nomi yoki raqami",
}


@dp.callback_query(lambda c: c.data.startswith("search:"))
async def process_search_start(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    _, entity_type = callback_query.data.split(":")  # "students" | "teachers" | "groups"
    await state.update_data(entity_type=entity_type)
    await state.set_state(SearchStates.waiting_name)

    label = SEARCH_LABELS.get(entity_type, "ism")
    cancel_markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Bekor qilish", callback_data="restart_search")]
    ])
    await callback_query.message.edit_text(
        text=f"🔎 Qidirish uchun {label}ni yozib yuboring:",
        reply_markup=cancel_markup,
        parse_mode=ParseMode.HTML
    )
    await callback_query.answer()


@dp.message(SearchStates.waiting_name)
async def process_search_query(message: types.Message, state: FSMContext) -> None:
    data = await state.get_data()
    entity_type = data.get("entity_type", "students")
    query = (message.text or "").strip()

    if not query:
        await message.answer("Iltimos, matn kiriting.")
        return

    wait_msg = await message.answer("Qidirilmoqda... ⏳")
    results = await search_entities(entity_type, query)

    if not results:
        extra = ""
        if entity_type == "groups":
            # ⚠️ Guruh qidiruv manzili UWED'da tasdiqlanmagan (qarang: GROUP_SEARCH_URLS)
            extra = (
                "\n\nℹ️ Guruh bo'yicha qidiruv hozircha taxminiy manzil orqali "
                "ishlaydi va UWED tizimida hali tasdiqlanmagan. Agar bu doim "
                "natija bermasa, admin bilan bog'laning."
            )
        await wait_msg.edit_text(
            "❌ Hech narsa topilmadi. Boshqa nom bilan urinib ko'ring yoki /start bosing." + extra
        )
        return

    results = results[:8]
    code = {"students": "s", "teachers": "t", "groups": "g"}.get(entity_type, "s")
    buttons = []
    for item in results:
        name = entity_display_name(item)
        item_id = item.get("id")
        if not item_id:
            continue
        ENTITY_NAME_CACHE[str(item_id)] = name
        buttons.append([InlineKeyboardButton(text=name, callback_data=f"p:{code}:{item_id}")])
    buttons.append([InlineKeyboardButton(text="⬅️ Bosh sahifa", callback_data="restart_search")])

    await state.clear()
    await wait_msg.edit_text(
        "Natijalardan birini tanlang 👇",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )


@dp.callback_query(lambda c: c.data.startswith("p:"))
async def process_person_selection(callback_query: types.CallbackQuery) -> None:
    _, code, entity_id = callback_query.data.split(":", 2)
    label = CODE_TO_LABEL.get(code, "Talaba")
    user_id = callback_query.from_user.id

    fav_key = f"{code}:{entity_id}"
    is_fav = await db_is_favorite(user_id, "person", fav_key)
    sub = await db_get_subscription(user_id)
    is_subscribed = bool(sub and sub[0] == code and sub[1] == entity_id)

    today = now_tashkent().date()
    day_buttons = [
        [InlineKeyboardButton(text="📍 Hozir qayerda? (joriy holat)", callback_data=f"now:{code}:{entity_id}")],
    ]
    for i in range(0, 6):
        next_day = today + datetime.timedelta(days=i)
        if next_day.isoweekday() == 7:
            continue
        date_str = next_day.strftime('%Y-%m-%d')
        btn_text = build_date_label(i, next_day)
        day_buttons.append([InlineKeyboardButton(text=btn_text, callback_data=f"pday:{code}:{entity_id}:{date_str}")])

    fav_text = "⭐ Sevimlilardan olib tashlash" if is_fav else "☆ Sevimlilarga saqlash"
    day_buttons.append([InlineKeyboardButton(text=fav_text, callback_data=f"fav:{code}:{entity_id}")])

    sub_text = "🔕 Obunani bekor qilish" if is_subscribed else "🔔 Kunlik jadval + eslatmaga obuna bo'lish"
    day_buttons.append([InlineKeyboardButton(text=sub_text, callback_data=f"sub:{code}:{entity_id}")])

    day_buttons.append([InlineKeyboardButton(text="⬅️ Bosh sahifa", callback_data="restart_search")])
    keyboard = InlineKeyboardMarkup(inline_keyboard=day_buttons)

    await callback_query.message.edit_text(
        text=f"📅 {label} uchun qaysi sanadagi jadvalni ko'rmoqchisiz?",
        reply_markup=keyboard,
        parse_mode=ParseMode.HTML
    )
    # ℹ️ Bu funksiya boshqa handler'lardan ham (masalan obuna almashtirilgach)
    # qayta chaqirilishi mumkin — o'sha holatlarda callback_query allaqachon
    # answer() qilingan bo'lishi mumkin, shuning uchun xatoni e'tiborsiz qoldiramiz.
    try:
        await callback_query.answer()
    except Exception:
        pass


@dp.callback_query(lambda c: c.data.startswith("now:"))
async def process_now_status(callback_query: types.CallbackQuery) -> None:
    _, code, entity_id = callback_query.data.split(":", 2)
    label = CODE_TO_LABEL.get(code, "Talaba")

    await callback_query.answer("Tekshirilmoqda...")

    now = now_tashkent()
    today_str = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M")

    # ✅ DevTools orqali tasdiqlangan: talaba uchun type=STUDENT, o'qituvchi uchun
    # type=TEACHER. Guruh uchun type=GROUP — ⚠️ tasdiqlanmagan (taxminiy).
    event_type = CODE_TO_EVENT_TYPE.get(code, "STUDENT")
    events = await fetch_week_events(entity_id, now.date(), event_type=event_type)
    current_event = None
    for e in events:
        if str(e.get("date", "")) == today_str:
            start_t = str(e.get("start_time", ""))
            end_t = str(e.get("end_time", ""))
            if start_t <= current_time <= end_t:
                current_event = e
                break

    header = f"📍 <b>HOZIRGI HOLAT</b>\n⏰ Vaqt: <b>{current_time}</b>\n\n"
    if current_event:
        room = get_event_room(current_event)
        text = (
            header
            + f"🔴 {label} hozir band!\n"
            + f"🚪 Xona: <b>{room}</b>\n"
            + f"📚 Fan: {current_event.get('title') or 'Nomsiz'}\n"
        )
        if code in ("s", "g") and current_event.get("instructor"):
            text += f"👨‍🏫 O'qituvchi: {current_event['instructor']}\n"
        if current_event.get("end_time"):
            text += f"⏰ Tugash vaqti: {current_event['end_time']}\n"
    else:
        text = header + f"🟢 {label} hozir band emas / darsda emas."

    restart_markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔄 Yangilash", callback_data=callback_query.data)],
        [InlineKeyboardButton(text="⬅️ Ortga", callback_data=f"p:{code}:{entity_id}")],
        [InlineKeyboardButton(text="Bosh sahifa 🔄", callback_data="restart_search")]
    ])
    await _safe_edit(callback_query, text, restart_markup)


@dp.callback_query(lambda c: c.data.startswith("pday:"))
async def process_person_day_selection(callback_query: types.CallbackQuery) -> None:
    _, code, entity_id, date_str = callback_query.data.split(":", 3)
    entity_type = CODE_TO_ENTITY_TYPE.get(code, "students")

    await callback_query.answer("Yuklanmoqda...")

    target_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    weekday_num = target_date.isoweekday()
    day_name = HAFTA_KUNLARI[weekday_num - 1] if weekday_num != 7 else "Yakshanba"

    # ✅ DevTools orqali tasdiqlangan: talaba uchun type=STUDENT, o'qituvchi uchun
    # type=TEACHER. Guruh uchun type=GROUP — ⚠️ tasdiqlanmagan (taxminiy).
    event_type = CODE_TO_EVENT_TYPE.get(code, "STUDENT")
    events = await fetch_week_events(entity_id, target_date, event_type=event_type)
    result_text = build_person_schedule_text(events, date_str, day_name, entity_type)

    restart_markup = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="📥 Excel", callback_data=f"xls:p:{code}:{entity_id}:{date_str}"),
            InlineKeyboardButton(text="📄 PDF", callback_data=f"pdf:p:{code}:{entity_id}:{date_str}"),
        ],
        [InlineKeyboardButton(text="⬅️ Ortga", callback_data=f"p:{code}:{entity_id}")],
        [InlineKeyboardButton(text="Bosh sahifa 🔄", callback_data="restart_search")]
    ])
    await send_result(callback_query, result_text, restart_markup)


# ============================================================
# ⭐ SEVIMLILAR — saqlash/olib tashlash va ro'yxatni ko'rish
# ============================================================

@dp.callback_query(lambda c: c.data.startswith("fav:"))
async def process_favorite_toggle(callback_query: types.CallbackQuery) -> None:
    """Xona (fav:room:BLOCK:FLOOR) yoki talaba/o'qituvchi/guruhni
    (fav:code:entity_id) sevimlilarga qo'shadi yoki olib tashlaydi."""
    parts = callback_query.data.split(":")
    user_id = callback_query.from_user.id

    if parts[1] == "room":
        block, floor = parts[2], parts[3]
        fav_key = f"{block}:{floor}"
        label = f"🏢 {block} blok, {floor_label(floor)}"
        already = await db_is_favorite(user_id, "room", fav_key)
        if already:
            await db_remove_favorite(user_id, "room", fav_key)
            await callback_query.answer("⭐ Sevimlilardan olib tashlandi.")
        else:
            await db_add_favorite(user_id, "room", fav_key, label)
            await callback_query.answer("⭐ Sevimlilarga saqlandi!")
        return

    code, entity_id = parts[1], parts[2]
    fav_key = f"{code}:{entity_id}"
    label = f"{CODE_TO_LABEL.get(code, '')}: {ENTITY_NAME_CACHE.get(entity_id, entity_id)}"
    already = await db_is_favorite(user_id, "person", fav_key)

    if already:
        await db_remove_favorite(user_id, "person", fav_key)
        await callback_query.answer("⭐ Sevimlilardan olib tashlandi.")
        new_text = "☆ Sevimlilarga saqlash"
    else:
        await db_add_favorite(user_id, "person", fav_key, label)
        await callback_query.answer("⭐ Sevimlilarga saqlandi!")
        new_text = "⭐ Sevimlilardan olib tashlash"

    # Tugma matnini joyida yangilaymiz (butun sahifani qayta yuklamasdan)
    try:
        kb = callback_query.message.reply_markup
        new_rows = []
        for row in kb.inline_keyboard:
            new_row = []
            for btn in row:
                if btn.callback_data == callback_query.data:
                    new_row.append(InlineKeyboardButton(text=new_text, callback_data=btn.callback_data))
                else:
                    new_row.append(btn)
            new_rows.append(new_row)
        await callback_query.message.edit_reply_markup(reply_markup=InlineKeyboardMarkup(inline_keyboard=new_rows))
    except Exception:
        pass


@dp.callback_query(lambda c: c.data == "favorites")
async def process_favorites_list(callback_query: types.CallbackQuery) -> None:
    user_id = callback_query.from_user.id
    rows = await db_list_favorites(user_id)

    if not rows:
        await callback_query.message.edit_text(
            text="⭐ <b>Sevimlilar</b>\n\nHali hech narsa saqlamagansiz.\n\n"
                 "Xona qidirganda yoki talaba/o'qituvchi/guruh sahifasida "
                 "\"☆ Sevimlilarga saqlash\" tugmasini bosing.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="⬅️ Bosh sahifa", callback_data="restart_search")]
            ]),
            parse_mode=ParseMode.HTML,
        )
        try:
            await callback_query.answer()
        except Exception:
            pass
        return

    buttons = []
    for fav_id, fav_type, fav_key, label in rows:
        if fav_type == "room":
            block, floor = fav_key.split(":")
            open_cb = f"floor:{block}:{floor}"
        else:
            code, entity_id = fav_key.split(":", 1)
            open_cb = f"p:{code}:{entity_id}"
        buttons.append([
            InlineKeyboardButton(text=label, callback_data=open_cb),
            InlineKeyboardButton(text="🗑", callback_data=f"favdel:{fav_id}"),
        ])
    buttons.append([InlineKeyboardButton(text="⬅️ Bosh sahifa", callback_data="restart_search")])

    await callback_query.message.edit_text(
        text="⭐ <b>Sevimlilar</b>\n\nOchish uchun nomini, o'chirish uchun 🗑 tugmasini bosing:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode=ParseMode.HTML,
    )
    try:
        await callback_query.answer()
    except Exception:
        pass


@dp.callback_query(lambda c: c.data.startswith("favdel:"))
async def process_favorite_delete(callback_query: types.CallbackQuery) -> None:
    fav_id = int(callback_query.data.split(":")[1])
    await db_remove_favorite_by_id(callback_query.from_user.id, fav_id)
    await callback_query.answer("🗑 O'chirildi.")
    await process_favorites_list(callback_query)


# ============================================================
# 🔔 OBUNA — kunlik jadval (07:00) va "10 daqiqa qoldi" eslatmasi
# ============================================================

@dp.callback_query(lambda c: c.data.startswith("sub:"))
async def process_subscription_toggle(callback_query: types.CallbackQuery) -> None:
    _, code, entity_id = callback_query.data.split(":", 2)
    user_id = callback_query.from_user.id
    label = f"{CODE_TO_LABEL.get(code, '')}: {ENTITY_NAME_CACHE.get(entity_id, entity_id)}"

    current = await db_get_subscription(user_id)
    if current and current[0] == code and current[1] == entity_id:
        await db_remove_subscription(user_id)
        await callback_query.answer("🔕 Obuna bekor qilindi.", show_alert=True)
    else:
        await db_set_subscription(user_id, code, entity_id, label)
        await callback_query.answer(
            "🔔 Obuna faollashtirildi!\nHar kuni soat 7:00 da kunlik jadval, "
            "har darsdan 10 daqiqa oldin eslatma keladi.",
            show_alert=True,
        )
    # Sahifani yangilaymiz, shunda tugma matni to'g'ri ko'rinadi
    await process_person_selection(callback_query)


@dp.callback_query(lambda c: c.data == "mysub")
async def process_my_subscription(callback_query: types.CallbackQuery) -> None:
    user_id = callback_query.from_user.id
    sub = await db_get_subscription(user_id)

    if not sub:
        await callback_query.message.edit_text(
            text="🔔 <b>Mening obunam</b>\n\nHozircha hech qanday obunangiz yo'q.\n\n"
                 "Talaba/o'qituvchi/guruh sahifasida \"🔔 Obuna bo'lish\" tugmasini bosing — "
                 "shunda har kuni soat 7:00 da kunlik jadvalingiz va har darsdan 10 daqiqa "
                 "oldin eslatma keladi.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="⬅️ Bosh sahifa", callback_data="restart_search")]
            ]),
            parse_mode=ParseMode.HTML,
        )
        try:
            await callback_query.answer()
        except Exception:
            pass
        return

    entity_type, entity_id, entity_label, daily_enabled, reminder_enabled = sub
    daily_text = "✅ Kunlik jadval (7:00): yoqilgan" if daily_enabled else "❌ Kunlik jadval (7:00): o'chirilgan"
    reminder_text = "✅ Dars eslatmasi (10 daq oldin): yoqilgan" if reminder_enabled else "❌ Dars eslatmasi (10 daq oldin): o'chirilgan"

    text = f"🔔 <b>Mening obunam</b>\n\n👤 {esc(entity_label)}\n\n{daily_text}\n{reminder_text}"
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔄 Kunlik jadvalni yoqish/o'chirish", callback_data="subtoggle:daily")],
        [InlineKeyboardButton(text="🔄 Eslatmani yoqish/o'chirish", callback_data="subtoggle:reminder")],
        [InlineKeyboardButton(text="🔕 Obunani butunlay bekor qilish", callback_data="subcancel")],
        [InlineKeyboardButton(text="⬅️ Bosh sahifa", callback_data="restart_search")],
    ])
    await callback_query.message.edit_text(text=text, reply_markup=keyboard, parse_mode=ParseMode.HTML)
    try:
        await callback_query.answer()
    except Exception:
        pass


@dp.callback_query(lambda c: c.data.startswith("subtoggle:"))
async def process_subscription_flag_toggle(callback_query: types.CallbackQuery) -> None:
    _, flag = callback_query.data.split(":")
    flag_column = "daily_enabled" if flag == "daily" else "reminder_enabled"
    await db_toggle_subscription_flag(callback_query.from_user.id, flag_column)
    await callback_query.answer("✅ Yangilandi.")
    await process_my_subscription(callback_query)


@dp.callback_query(lambda c: c.data == "subcancel")
async def process_subscription_cancel(callback_query: types.CallbackQuery) -> None:
    await db_remove_subscription(callback_query.from_user.id)
    await callback_query.answer("🔕 Obuna butunlay bekor qilindi.", show_alert=True)
    await process_my_subscription(callback_query)


# ============================================================
# 📥 EXCEL / PDF EKSPORT — tugma bosilganda faylni tayyorlab yuborish
# ============================================================

@dp.callback_query(lambda c: c.data.startswith("xls:") or c.data.startswith("pdf:"))
async def process_export(callback_query: types.CallbackQuery) -> None:
    fmt, kind, *rest = callback_query.data.split(":")
    await callback_query.answer("📥 Fayl tayyorlanmoqda...")

    try:
        if kind == "p":
            code, entity_id, date_str = rest
            entity_type = CODE_TO_ENTITY_TYPE.get(code, "students")
            label = f"{CODE_TO_LABEL.get(code, '')}: {ENTITY_NAME_CACHE.get(entity_id, entity_id)}"
            event_type = CODE_TO_EVENT_TYPE.get(code, "STUDENT")
            target_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            events = await fetch_week_events(entity_id, target_date, event_type=event_type)

            if fmt == "xls":
                content = build_person_schedule_excel(events, date_str, label, entity_type)
                filename = f"jadval_{date_str}.xlsx"
            else:
                content = build_person_schedule_pdf(events, date_str, label, entity_type)
                filename = f"jadval_{date_str}.pdf"

        elif kind == "r":
            block, floor, date_str, para_id = rest
            rooms_in_floor = get_rooms_for_floor(block, floor)
            start_t, end_t, display = PARALAR[para_id]
            busy_rooms, invalid_rooms = await get_room_statuses(rooms_in_floor, block, date_str, start_t)
            title = f"{block} blok, {floor_label(floor)} - {date_str} ({display})"

            if fmt == "xls":
                content = build_room_status_excel(rooms_in_floor, busy_rooms, invalid_rooms, block, floor, title)
                filename = f"xonalar_{block}_{date_str}.xlsx"
            else:
                content = build_room_status_pdf(rooms_in_floor, busy_rooms, invalid_rooms, block, floor, title)
                filename = f"xonalar_{block}_{date_str}.pdf"
        else:
            return

        input_file = BufferedInputFile(content, filename=filename)
        await callback_query.message.answer_document(input_file)
    except Exception as e:
        print(f"[EXPORT] xato: {e}")
        await callback_query.message.answer("❌ Faylni tayyorlashda xato yuz berdi. Qayta urinib ko'ring.")


# ============================================================
# 🤖 AI-YORDAMCHI — foydalanuvchi bilan tabiiy tilda suhbat
# ============================================================

@dp.callback_query(lambda c: c.data == "ai_start")
async def process_ai_start(callback_query: types.CallbackQuery, state: FSMContext) -> None:
    if ai_client is None:
        await callback_query.answer(
            "🤖 AI-yordamchi hozircha sozlanmagan (ANTHROPIC_API_KEY yo'q).",
            show_alert=True,
        )
        return

    await state.set_state(AIStates.chatting)
    await state.update_data(ai_history=[])

    user_id = callback_query.from_user.id
    limit_line = ""
    if user_id != ADMIN_ID:
        used = await db_get_ai_usage_today(user_id)
        remaining = max(AI_DAILY_LIMIT - used, 0)
        limit_line = f"\n\n🚦 Bugungi limitingiz: {remaining}/{AI_DAILY_LIMIT} so'rov qoldi."

    await callback_query.message.edit_text(
        text=(
            "🤖 <b>AI-yordamchi</b>\n\n"
            "Menga oddiy tilda savolingizni yozing, masalan:\n"
            "• <i>\"A blokda hozir bo'sh xona bormi?\"</i>\n"
            "• <i>\"A305 xona bugun qaysi paralarda band?\"</i>\n"
            "• <i>\"Aliyev Vali ertaga band bo'ladimi?\"</i>\n"
            "• <i>\"D blok 3-qavatda uzoq vaqt bo'sh turadigan xona bormi?\"</i>\n"
            "• <i>\"512-guruh bugun soat nechada darsi bor?\"</i>"
            + limit_line
        ),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="⬅️ AI suhbatni tugatish", callback_data="restart_search")]
        ]),
        parse_mode=ParseMode.HTML,
    )
    await callback_query.answer()


@dp.message(AIStates.chatting)
async def process_ai_message(message: types.Message, state: FSMContext) -> None:
    if ai_client is None:
        await message.answer("🤖 AI-yordamchi hozircha sozlanmagan.")
        await state.clear()
        return

    user_text = (message.text or "").strip()
    if not user_text:
        await message.answer("Iltimos, savolingizni matn ko'rinishida yozing.")
        return

    user_id = message.from_user.id

    # 🚦 Kunlik limitni tekshiramiz (admin bundan ozod)
    if user_id != ADMIN_ID:
        used = await db_get_ai_usage_today(user_id)
        if used >= AI_DAILY_LIMIT:
            await message.answer(
                f"🚦 Siz bugun AI-yordamchidan {AI_DAILY_LIMIT} martalik limitni "
                f"to'liq ishlatib bo'ldingiz.\n\nErtaga soat 00:00 (Toshkent vaqti) da "
                f"limit yangilanadi. Bu vaqtgacha botning oddiy tugmalari (xona qidirish, "
                f"talaba/o'qituvchi/guruh qidiruvi va h.k.) cheklovsiz ishlayveradi."
            )
            return

    data = await state.get_data()
    history = data.get("ai_history", [])

    wait_msg = await message.answer("🤖 O'ylanmoqda... ⏳")

    now = now_tashkent()
    weekday_num = now.isoweekday()
    day_name = HAFTA_KUNLARI[weekday_num - 1] if weekday_num != 7 else "Yakshanba"
    context_note = (
        f"\n\n(Tizim ma'lumoti — foydalanuvchiga ko'rsatmang: hozirgi sana "
        f"{now.strftime('%Y-%m-%d')}, {day_name}, vaqt {now.strftime('%H:%M')}, Toshkent.)"
    )

    messages = history + [{"role": "user", "content": user_text}]

    try:
        final_text = await run_ai_conversation(messages, context_note)
    except Exception as e:
        print(f"[AI] Suhbatda xato: {e}")
        try:
            await wait_msg.edit_text("❌ AI bilan bog'lanishda xato yuz berdi. Birozdan so'ng qayta urinib ko'ring.")
        except Exception:
            pass
        return

    # ✅ Muvaffaqiyatli javob berilgach hisoblagichni oshiramiz (xatolar
    # uchun foydalanuvchi jazolanmasligi kerak — shuning uchun bu yerda,
    # API chaqiruvidan KEYIN qilinadi)
    if user_id != ADMIN_ID:
        new_used = await db_increment_ai_usage(user_id)
        remaining = max(AI_DAILY_LIMIT - new_used, 0)
        if remaining <= 3:
            final_text += f"\n\n(ℹ️ Bugun {remaining} ta AI so'rovingiz qoldi.)"

    # Suhbat tarixini saqlaymiz (juda uzun bo'lib ketmasligi uchun oxirgi
    # 3 ta savol-javob juftini — 6 ta xabarni — saqlaymiz)
    new_history = (history + [
        {"role": "user", "content": user_text},
        {"role": "assistant", "content": final_text},
    ])[-6:]
    await state.update_data(ai_history=new_history)

    # ⚠️ AI javobi oddiy matn: bot standart HTML rejimida ishlagani uchun javobdagi
    # "<" yoki "&" belgilari xabarni buzmasligi uchun parse_mode=None qilib yuboramiz.
    try:
        await wait_msg.edit_text(final_text[:4000], parse_mode=None)
    except Exception:
        await message.answer(final_text[:4000], parse_mode=None)


# 🌐 Render kabi hostinglar uchun oddiy "tirik" javob beruvchi endpoint
async def handle(request):
    return web.Response(text="Bot is running!")


# ============================================================
# 💤 RENDER "UXLAB QOLISH" MUAMMOSINING OLDINI OLISH
# Render'ning bepul tarifi 15 daqiqa davomida hech qanday tashqi HTTP so'rov
# kelmasa, xizmatni to'xtatib (spin down) qo'yadi. Bot esa Telegram bilan faqat
# OUTBOUND (o'zi so'rov yuborish - polling) orqali ishlagani uchun Render buni
# "faollik" deb hisoblamaydi. Shuning uchun bot o'zining ochiq URL'iga vaqti-
# vaqti bilan o'zi so'rov yuborib, "men tirikman" deb bildirib turadi.
#
# ⚠️ Bu 100% kafolat bermaydi — eng ishonchli yechim tashqi monitoring xizmati
# (masalan bepul https://uptimerobot.com) orqali RENDER_EXTERNAL_URL manzilini
# har 5-10 daqiqada "ping" qilib turishdir.
# ============================================================

SELF_PING_INTERVAL_SECONDS = 600  # 10 daqiqa


async def self_ping_loop() -> None:
    external_url = os.environ.get("RENDER_EXTERNAL_URL")
    if not external_url:
        print("[SELF-PING] RENDER_EXTERNAL_URL topilmadi, self-ping o'chirilgan.")
        return

    print(f"[SELF-PING] Yoqildi: {external_url} har {SELF_PING_INTERVAL_SECONDS} soniyada so'raladi.")
    while True:
        await asyncio.sleep(SELF_PING_INTERVAL_SECONDS)
        try:
            async with http_session.get(
                external_url, timeout=aiohttp.ClientTimeout(total=20)
            ) as resp:
                print(f"[SELF-PING] status={resp.status}")
        except Exception as e:
            print(f"[SELF-PING] xato: {e}")


# ============================================================
# 🔔📆 OBUNACHILAR UCHUN FON VAZIFALARI
# 1) daily_schedule_loop — har kuni soat 07:00 (Toshkent vaqti) da obuna
#    bo'lgan har bir foydalanuvchiga o'sha kunlik jadvalini yuboradi.
# 2) reminder_loop — har daqiqada tekshiradi: agar kimningdir darsi 10
#    daqiqadan keyin boshlansa, unga eslatma yuboradi.
# ============================================================

# (foydalanuvchi_id, sana, boshlanish_vaqti) — bugun allaqachon eslatma
# yuborilgan darslarni belgilab qo'yish uchun (qayta-qayta yubormaslik).
# Bot qayta ishga tushganda tozalanadi — bu xavfsiz, chunki eng yomon holatda
# faqat bitta eslatma qayta yuborilishi mumkin.
NOTIFIED_REMINDERS: set[tuple] = set()


async def send_daily_schedules(bot: Bot) -> None:
    rows = await db_list_daily_subscriptions()
    if not rows:
        return

    today = now_tashkent().date()
    today_str = today.strftime("%Y-%m-%d")
    weekday_num = today.isoweekday()
    day_name = HAFTA_KUNLARI[weekday_num - 1] if weekday_num != 7 else "Yakshanba"

    for user_id, entity_type_code, entity_id, entity_label in rows:
        event_type = CODE_TO_EVENT_TYPE.get(entity_type_code, "STUDENT")
        entity_type = CODE_TO_ENTITY_TYPE.get(entity_type_code, "students")
        try:
            events = await fetch_week_events(entity_id, today, event_type=event_type)
            schedule_text = build_person_schedule_text(events, today_str, day_name, entity_type)
            text = f"☀️ <b>Xayrli tong! Bugungi jadvalingiz</b>\n👤 {esc(entity_label)}\n\n{schedule_text}"
            await bot.send_message(chat_id=user_id, text=text, parse_mode=ParseMode.HTML)
        except TelegramForbiddenError:
            # Foydalanuvchi botni bloklagan — obunasini avtomatik bekor qilamiz
            await db_remove_subscription(user_id)
        except Exception as e:
            print(f"[KUNLIK JADVAL] {user_id} ga yuborishda xato: {e}")
        await asyncio.sleep(0.05)  # flood-limitga tushmaslik uchun


async def daily_schedule_loop(bot: Bot) -> None:
    while True:
        now = now_tashkent()
        target = now.replace(hour=7, minute=0, second=0, microsecond=0)
        if now >= target:
            target += datetime.timedelta(days=1)
        wait_seconds = (target - now).total_seconds()
        await asyncio.sleep(wait_seconds)
        try:
            await send_daily_schedules(bot)
        except Exception as e:
            print(f"[KUNLIK JADVAL] Umumiy xato: {e}")


async def reminder_loop(bot: Bot) -> None:
    while True:
        await asyncio.sleep(60)
        try:
            now = now_tashkent()
            today_str = now.date().isoformat()
            current_hm = now.strftime("%H:%M")

            rows = await db_list_reminder_subscriptions()
            for user_id, entity_type_code, entity_id, entity_label in rows:
                event_type = CODE_TO_EVENT_TYPE.get(entity_type_code, "STUDENT")
                try:
                    events = await fetch_week_events(entity_id, now.date(), event_type=event_type)
                except Exception as e:
                    print(f"[ESLATMA] Jadval olishda xato ({user_id}): {e}")
                    continue

                for e in events:
                    if str(e.get("date", "")) != today_str:
                        continue
                    start_t = str(e.get("start_time", ""))
                    if not start_t or ":" not in start_t:
                        continue
                    try:
                        hh, mm = start_t.split(":")[:2]
                        start_dt = now.replace(hour=int(hh), minute=int(mm), second=0, microsecond=0)
                    except (ValueError, TypeError):
                        continue

                    reminder_dt = start_dt - datetime.timedelta(minutes=10)
                    key = (user_id, today_str, start_t)
                    if reminder_dt.strftime("%H:%M") == current_hm and key not in NOTIFIED_REMINDERS:
                        NOTIFIED_REMINDERS.add(key)
                        text = (
                            f"⏰ <b>10 daqiqadan keyin darsingiz boshlanadi!</b>\n\n"
                            f"👤 {esc(entity_label)}\n"
                            f"📚 Fan: {esc(e.get('title') or 'Nomsiz')}\n"
                            f"🚪 Xona: {esc(get_event_room(e))}\n"
                            f"⏱ {start_t} - {e.get('end_time', '')}"
                        )
                        try:
                            await bot.send_message(chat_id=user_id, text=text, parse_mode=ParseMode.HTML)
                        except TelegramForbiddenError:
                            await db_remove_subscription(user_id)
                        except Exception as e2:
                            print(f"[ESLATMA] {user_id} ga yuborishda xato: {e2}")

            # Xotira cheksiz o'smasligi uchun vaqti-vaqti bilan tozalaymiz
            if len(NOTIFIED_REMINDERS) > 5000:
                NOTIFIED_REMINDERS.clear()
        except Exception as e:
            print(f"[ESLATMA] Umumiy xato: {e}")


async def run_polling_forever(bot: Bot) -> None:
    """dp.start_polling ichida kutilmagan xato yuz bersa ham (masalan tarmoq
    uzilishi), bot butunlay to'xtab qolmasligi uchun avtomatik qayta
    ishga tushiriladi."""
    while True:
        try:
            await dp.start_polling(bot)
        except Exception as e:
            print(f"[POLLING] Kutilmagan xato, 5 soniyadan keyin qayta urinilyapti: {e}")
            await asyncio.sleep(5)
        else:
            # start_polling odatda cheksiz ishlaydi; agar u shunchaki tugasa ham,
            # qayta ishga tushiramiz (masalan aiogram'ning o'z sabablariga ko'ra)
            print("[POLLING] start_polling to'xtadi, qayta ishga tushirilyapti...")
            await asyncio.sleep(2)


# 🚀 BOTNI ISHGA TUSHIRUVCHI ASOSIY FUNKSIYA
async def main() -> None:
    # Bazani ishga tushirishning eng boshida tayyorlaymiz (aks holda birinchi
    # xabardayoq "no such table: users" xatosi chiqadi).
    await asyncio.to_thread(_init_db)

    global http_session
    http_session = aiohttp.ClientSession()

    try:
        async with http_session.get(
            "https://manage.uwed.uz/public/pages/inf-tt",
            headers={"User-Agent": REQUEST_HEADERS["User-Agent"]},
            timeout=aiohttp.ClientTimeout(total=15),
        ) as resp:
            print(f"[DEBUG] Bosh sahifa ochildi: status={resp.status}")
    except Exception as e:
        print(f"[DEBUG] Bosh sahifani ochishda xato: {e}")

    # Render "web service" turi tirik HTTP portni talab qiladi, shuning
    # uchun oddiy aiohttp serverni fon rejimida ko'taramiz.
    app = web.Application()
    app.router.add_get('/', handle)
    from webapp import register_webapp
    register_webapp(app, sys.modules[__name__])
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    global BOT_INSTANCE
    bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    BOT_INSTANCE = bot

    # Admin kuzatuv middleware'ini ro'yxatdan o'tkazamiz. OUTER middleware sifatida
    # (oddiy .middleware() emas!) — chunki oddiy "inner" middleware faqat mos
    # handler (masalan /start yoki qidiruv holati) topilganda ishga tushadi.
    # Bo'sh matn, stiker yoki hech qaysi handler'ga mos kelmaydigan xabarlar esa
    # "inner" middleware'ni umuman ishga tushirmaydi. Outer middleware esa HAR
    # QANDAY kiruvchi xabar/tugma uchun ishlaydi, handler topilgan-topilmaganidan
    # qat'iy nazar — shu sabab admin barcha harakatlarni ko'ra oladi.
    dp.message.outer_middleware(AdminLoggerMiddleware())
    dp.callback_query.outer_middleware(AdminLoggerMiddleware())

    print("Bot muvaffaqiyatli ishga tushdi va xabarlarni kutyapti...")

    # Self-ping va obunachilar uchun fon vazifalarini ishga tushiramiz
    # (polling'ni to'sib qo'ymasligi uchun alohida task sifatida)
    ping_task = asyncio.create_task(self_ping_loop())
    daily_task = asyncio.create_task(daily_schedule_loop(bot))
    reminder_task = asyncio.create_task(reminder_loop(bot))

    try:
        await run_polling_forever(bot)
    finally:
        ping_task.cancel()
        daily_task.cancel()
        reminder_task.cancel()
        await http_session.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    asyncio.run(main())