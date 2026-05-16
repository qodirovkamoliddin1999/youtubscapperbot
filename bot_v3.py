#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YouTube Comments Scraper Telegram Bot v3
- API quota monitoring
- Xatolikka bardoshlilik (retry)
- Dublikat izohlarni olib tashlash
- TXT/CSV/Excel eksport
- Statistika
- Playlist qo'llab-quvvatlash
- Admin tasdiqlash tizimi (xabarnoma bilan)
"""

import os
import re
import html
import sqlite3
import asyncio
import csv
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
from io import BytesIO

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, Bot
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

try:
    import emoji
    EMOJI_INSTALLED = True
except ImportError:
    EMOJI_INSTALLED = False

try:
    import openpyxl
    EXCEL_INSTALLED = True
except ImportError:
    EXCEL_INSTALLED = False

try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    MATPLOTLIB_INSTALLED = True
except ImportError:
    MATPLOTLIB_INSTALLED = False

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import httplib2
import socket

# Configuration
BOT_TOKEN = "7281638441:AAHYcT2v825jxlz0Z_ET3TeXuz__cj_tUnI"
ADMIN_IDS = [656337840]  # Adminlar ID raqamlari
SESSION_HOURS = 12
MIN_COMMENT_LENGTH = 3
MAX_RETRIES = 3
RETRY_DELAY = 2
DB_PATH = Path(__file__).parent / "users.db"
TEMP_DIR = Path(__file__).parent / "temp"

# Conversation states
WAITING_API_KEY = 1
WAITING_CHANNEL = 2
WAITING_PLAYLIST = 3

# Progress animation
LOADING_FRAMES = ["◐", "◓", "◑", "◒"]

TEMP_DIR.mkdir(exist_ok=True)


# ==================== DATABASE ====================

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            api_key TEXT,
            session_start TIMESTAMP,
            username TEXT,
            quota_used INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            is_approved INTEGER DEFAULT 0,
            is_admin INTEGER DEFAULT 0
        )
    """)
    # Eski database uchun ustunlarni qo'shish
    for col_def in [
        "ALTER TABLE users ADD COLUMN quota_used INTEGER DEFAULT 0",
        "ALTER TABLE users ADD COLUMN is_approved INTEGER DEFAULT 0",
        "ALTER TABLE users ADD COLUMN is_admin INTEGER DEFAULT 0",
    ]:
        try:
            c.execute(col_def)
        except sqlite3.OperationalError:
            pass

    # Oldindan belgilangan adminlarni qo'shish
    for admin_id in ADMIN_IDS:
        c.execute("""
            INSERT OR IGNORE INTO users (user_id, api_key, is_approved, is_admin)
            VALUES (?, '', 1, 1)
        """, (admin_id,))

    conn.commit()
    conn.close()


def get_user(user_id: int) -> dict | None:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "SELECT api_key, session_start, username, quota_used, is_approved, is_admin "
        "FROM users WHERE user_id = ?",
        (user_id,),
    )
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "api_key": row[0],
            "session_start": datetime.fromisoformat(row[1]) if row[1] else None,
            "username": row[2],
            "quota_used": row[3] or 0,
            "is_approved": row[4] or 0,
            "is_admin": row[5] or 0,
        }
    return None


def save_user(user_id: int, api_key: str, username: str = None):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Foydalanuvchi allaqachon bazada bormi?
    c.execute("SELECT is_admin, is_approved FROM users WHERE user_id = ?", (user_id,))
    existing = c.fetchone()

    if existing:
        # Faqat api_key, session_start, username ni yangilash — is_admin/is_approved ni o'zgartirmaslik!
        c.execute("""
            UPDATE users SET api_key = ?, session_start = ?, username = ?
            WHERE user_id = ?
        """, (api_key, datetime.now().isoformat(), username, user_id))
        conn.commit()
        conn.close()
        return existing[0]  # is_admin ni qaytarish

    # Yangi foydalanuvchi — admin IDlar ro'yxatida bormi?
    is_admin = 1 if user_id in ADMIN_IDS else 0
    is_approved = 1 if user_id in ADMIN_IDS else 0

    c.execute("""
        INSERT INTO users (user_id, api_key, session_start, username, quota_used, is_approved, is_admin)
        VALUES (?, ?, ?, ?, 0, ?, ?)
    """, (user_id, api_key, datetime.now().isoformat(), username, is_approved, is_admin))
    conn.commit()
    conn.close()
    return is_admin


def is_new_user(user_id: int) -> bool:
    """Foydalanuvchi bazada yo'qmi?"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT 1 FROM users WHERE user_id = ?", (user_id,))
    exists = c.fetchone() is not None
    conn.close()
    return not exists


def update_quota(user_id: int, units: int):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE users SET quota_used = quota_used + ? WHERE user_id = ?", (units, user_id))
    conn.commit()
    conn.close()


def get_quota_used(user_id: int) -> int:
    user = get_user(user_id)
    return user["quota_used"] if user else 0


def is_session_valid(user_id: int) -> bool:
    user = get_user(user_id)
    if not user or not user["session_start"]:
        return False
    expires_at = user["session_start"] + timedelta(hours=SESSION_HOURS)
    return datetime.now() < expires_at


def get_session_remaining(user_id: int) -> str:
    user = get_user(user_id)
    if not user or not user["session_start"]:
        return "Sessiya topilmadi"
    expires_at = user["session_start"] + timedelta(hours=SESSION_HOURS)
    remaining = expires_at - datetime.now()
    if remaining.total_seconds() <= 0:
        return "Sessiya tugagan"
    hours, remainder = divmod(int(remaining.total_seconds()), 3600)
    minutes, _ = divmod(remainder, 60)
    return f"{hours} soat {minutes} daqiqa"


def approve_user(user_id: int) -> bool:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE users SET is_approved = 1 WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    return True


def reject_user(user_id: int) -> bool:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM users WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    return True


def make_admin(user_id: int) -> bool:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE users SET is_admin = 1, is_approved = 1 WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    return True


def remove_admin(user_id: int) -> bool:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE users SET is_admin = 0 WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    return True


def get_all_users() -> list:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "SELECT user_id, username, is_approved, is_admin, created_at "
        "FROM users ORDER BY created_at DESC"
    )
    rows = c.fetchall()
    conn.close()
    return rows


def get_all_admin_ids() -> list[int]:
    """Barcha adminlar ID sini qaytaradi."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT user_id FROM users WHERE is_admin = 1")
    rows = c.fetchall()
    conn.close()
    return [r[0] for r in rows]


# ==================== RETRY DECORATOR ====================

async def retry_on_error(func, *args, max_retries=MAX_RETRIES, **kwargs):
    last_error = None
    for attempt in range(max_retries):
        try:
            return await asyncio.get_event_loop().run_in_executor(None, lambda: func(*args, **kwargs))
        except (httplib2.ServerNotFoundError, socket.timeout, ConnectionError, OSError) as e:
            last_error = e
            if attempt < max_retries - 1:
                await asyncio.sleep(RETRY_DELAY * (attempt + 1))
            continue
        except HttpError as e:
            raise e
    raise last_error


def sync_retry(func, *args, max_retries=MAX_RETRIES, **kwargs):
    last_error = None
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except (httplib2.ServerNotFoundError, socket.timeout, ConnectionError, OSError) as e:
            last_error = e
            if attempt < max_retries - 1:
                import time
                time.sleep(RETRY_DELAY * (attempt + 1))
            continue
        except HttpError as e:
            raise e
    raise last_error


# ==================== ADMIN NOTIFICATION ====================

async def notify_admins_new_user(bot: Bot, new_user_id: int, username: str | None, full_name: str):
    """Barcha adminlarga yangi foydalanuvchi haqida xabar yuborish."""
    admin_ids = get_all_admin_ids()
    if not admin_ids:
        return

    display = f"@{username}" if username else full_name
    text = (
        f"🔔 *Yangi foydalanuvchi kirdi!*\n\n"
        f"👤 Ism: {full_name}\n"
        f"🔗 Username: {'@' + username if username else '—'}\n"
        f"🆔 ID: `{new_user_id}`\n\n"
        f"Botdan foydalanishga ruxsat berasizmi?"
    )

    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"approve_{new_user_id}"),
            InlineKeyboardButton("❌ Rad etish", callback_data=f"reject_{new_user_id}"),
        ]
    ])

    for admin_id in admin_ids:
        try:
            await bot.send_message(
                chat_id=admin_id,
                text=text,
                parse_mode="Markdown",
                reply_markup=keyboard,
            )
        except Exception as e:
            print(f"Admin {admin_id} ga xabar yuborishda xato: {e}")


async def notify_user_approved(bot: Bot, user_id: int):
    """Foydalanuvchiga tasdiqlanganligi haqida xabar yuborish."""
    try:
        await bot.send_message(
            chat_id=user_id,
            text=(
                "✅ *Tabriklaymiz! So'rovingiz tasdiqlandi.*\n\n"
                "Endi botdan to'liq foydalanishingiz mumkin.\n"
                "/start buyrug'ini yuboring."
            ),
            parse_mode="Markdown",
        )
    except Exception as e:
        print(f"Foydalanuvchi {user_id} ga xabar yuborishda xato: {e}")


async def notify_user_rejected(bot: Bot, user_id: int):
    """Foydalanuvchiga rad etilganligi haqida xabar yuborish."""
    try:
        await bot.send_message(
            chat_id=user_id,
            text=(
                "❌ *Afsuski, so'rovingiz rad etildi.*\n\n"
                "Qo'shimcha ma'lumot uchun admin bilan bog'laning."
            ),
            parse_mode="Markdown",
        )
    except Exception as e:
        print(f"Foydalanuvchi {user_id} ga xabar yuborishda xato: {e}")


# ==================== PROGRESS HELPER ====================

async def update_progress(message, text: str, current: int, total: int, extra_info: str = "", frame_idx: int = 0):
    percent = int((current / total) * 100) if total > 0 else 0
    bar_filled = int(percent / 10)
    bar_empty = 10 - bar_filled
    progress_bar = "█" * bar_filled + "░" * bar_empty
    frame = LOADING_FRAMES[frame_idx % len(LOADING_FRAMES)]

    progress_text = f"{frame} *{text}*\n\n[{progress_bar}] {percent}%\n📊 {current}/{total}"
    if extra_info:
        progress_text += f"\n{extra_info}"

    try:
        await message.edit_text(progress_text, parse_mode="Markdown")
    except:
        pass


# ==================== YOUTUBE API ====================

def validate_api_key(api_key: str) -> tuple[bool, str]:
    try:
        youtube = build("youtube", "v3", developerKey=api_key)
        youtube.videos().list(part="snippet", id="dQw4w9WgXcQ").execute()
        return True, "API kalit to'g'ri ✅"
    except HttpError as e:
        if "API key not valid" in str(e) or "badRequest" in str(e):
            return False, "API kalit noto'g'ri ❌"
        elif "quotaExceeded" in str(e):
            return False, "API quota tugagan ❌"
        return False, f"Xato: {str(e)[:100]}"
    except Exception as e:
        return False, f"Xato: {str(e)[:100]}"


def get_channel_info(api_key: str, query: str) -> tuple[dict | None, str]:
    try:
        youtube = build("youtube", "v3", developerKey=api_key)

        if query.startswith("@"):
            handle = query[1:]
            res = sync_retry(youtube.channels().list(
                part="snippet,contentDetails,statistics",
                forHandle=handle
            ).execute)
            if res.get("items"):
                channel = res["items"][0]
                return {
                    "id": channel["id"],
                    "title": channel["snippet"]["title"],
                    "video_count": channel["statistics"].get("videoCount", "0"),
                    "uploads_playlist": channel["contentDetails"]["relatedPlaylists"]["uploads"]
                }, None

        if "youtube.com" in query:
            if "/@" in query:
                handle = query.split("/@")[1].split("/")[0].split("?")[0]
                return get_channel_info(api_key, f"@{handle}")
            if "channel/" in query:
                channel_id = query.split("channel/")[1].split("/")[0].split("?")[0]
                res = sync_retry(youtube.channels().list(
                    part="snippet,contentDetails,statistics",
                    id=channel_id
                ).execute)
                if res.get("items"):
                    channel = res["items"][0]
                    return {
                        "id": channel["id"],
                        "title": channel["snippet"]["title"],
                        "video_count": channel["statistics"].get("videoCount", "0"),
                        "uploads_playlist": channel["contentDetails"]["relatedPlaylists"]["uploads"]
                    }, None

        req = sync_retry(youtube.search().list(
            part="snippet",
            q=query,
            type="channel",
            maxResults=1
        ).execute)

        if req.get("items"):
            channel_id = req["items"][0]["snippet"]["channelId"]
            res = sync_retry(youtube.channels().list(
                part="snippet,contentDetails,statistics",
                id=channel_id
            ).execute)
            if res.get("items"):
                channel = res["items"][0]
                return {
                    "id": channel["id"],
                    "title": channel["snippet"]["title"],
                    "video_count": channel["statistics"].get("videoCount", "0"),
                    "uploads_playlist": channel["contentDetails"]["relatedPlaylists"]["uploads"]
                }, None

        return None, "Kanal topilmadi"
    except HttpError as e:
        if "quotaExceeded" in str(e):
            return None, "API quota tugadi! Ertaga qayta urinib ko'ring."
        return None, f"API xatosi: {str(e)[:100]}"
    except Exception as e:
        return None, f"Xato: {str(e)[:100]}"


def get_playlist_info(api_key: str, playlist_url: str) -> tuple[dict | None, str]:
    try:
        playlist_id = None
        if "list=" in playlist_url:
            playlist_id = playlist_url.split("list=")[1].split("&")[0]
        elif playlist_url.startswith("PL"):
            playlist_id = playlist_url

        if not playlist_id:
            return None, "Playlist ID topilmadi"

        youtube = build("youtube", "v3", developerKey=api_key)
        res = sync_retry(youtube.playlists().list(
            part="snippet,contentDetails",
            id=playlist_id
        ).execute)

        if res.get("items"):
            playlist = res["items"][0]
            return {
                "id": playlist_id,
                "title": playlist["snippet"]["title"],
                "video_count": playlist["contentDetails"]["itemCount"]
            }, None

        return None, "Playlist topilmadi"
    except HttpError as e:
        return None, f"API xatosi: {str(e)[:100]}"
    except Exception as e:
        return None, f"Xato: {str(e)[:100]}"


def get_all_video_links_sync(api_key: str, playlist_id: str) -> list:
    youtube = build("youtube", "v3", developerKey=api_key)
    links = []
    next_page = None

    while True:
        res = sync_retry(youtube.playlistItems().list(
            part="snippet",
            playlistId=playlist_id,
            maxResults=50,
            pageToken=next_page
        ).execute)

        for item in res["items"]:
            video_id = item["snippet"]["resourceId"]["videoId"]
            links.append(f"https://www.youtube.com/watch?v={video_id}")

        next_page = res.get("nextPageToken")
        if not next_page:
            break

    return links


def is_meaningless(text: str) -> bool:
    if re.match(r'^(.)\1{3,}$', text.lower()):
        return True
    if re.match(r'^.(.)\1{4,}$', text.lower()):
        return True
    if len(set(text.lower().replace(' ', ''))) <= 2 and len(text) > 3:
        return True
    return False


def clean_comment(text: str) -> str | None:
    if not text:
        return None

    text = html.unescape(text)
    text = re.sub(r'https?://\S+', '', text)
    text = re.sub(r'www\.\S+', '', text)
    text = re.sub(r'@\w+', '', text)
    text = re.sub(r'#\w+', '', text)

    if EMOJI_INSTALLED:
        text = emoji.replace_emoji(text, '')
    else:
        emoji_pattern = re.compile(
            "[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF"
            "\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF"
            "\U00002702-\U000027B0\U000024C2-\U0001F251]+",
            flags=re.UNICODE,
        )
        text = emoji_pattern.sub('', text)

    text = re.sub(r'<[^>]+>', '', text)
    text = re.sub(r"[^\w\s\u0400-\u04FF\u0600-\u06FF']", ' ', text)
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()

    if text.isdigit():
        return None
    if is_meaningless(text):
        return None

    words = text.split()
    if len(words) < 3:
        return None
    valid_words = [w for w in words if len(w) >= 2]
    if len(valid_words) < 3:
        return None

    return text


def get_video_comments(api_key: str, video_id: str) -> tuple[list, str | None, int]:
    quota_units = 1
    try:
        youtube = build("youtube", "v3", developerKey=api_key)
        comments = []
        next_page = None

        while True:
            res = sync_retry(youtube.commentThreads().list(
                part="snippet,replies",
                videoId=video_id,
                maxResults=100,
                pageToken=next_page,
                textFormat="plainText"
            ).execute)

            for item in res.get("items", []):
                top_comment = item["snippet"]["topLevelComment"]["snippet"]["textDisplay"]
                cleaned = clean_comment(top_comment)
                if cleaned:
                    comments.append(cleaned)
                if "replies" in item:
                    for reply in item["replies"]["comments"]:
                        reply_text = reply["snippet"]["textDisplay"]
                        cleaned = clean_comment(reply_text)
                        if cleaned:
                            comments.append(cleaned)

            next_page = res.get("nextPageToken")
            if next_page:
                quota_units += 1
            else:
                break

        return comments, None, quota_units
    except HttpError as e:
        error_str = str(e)
        if "commentsDisabled" in error_str:
            return [], "Izohlar yopiq", quota_units
        elif "videoNotFound" in error_str:
            return [], "Video topilmadi", quota_units
        elif "quotaExceeded" in error_str:
            return [], "API quota tugadi", quota_units
        return [], "API xatosi", quota_units
    except Exception:
        return [], "Xato", quota_units


# ==================== STATISTICS ====================

def generate_statistics(comments: list, title: str) -> dict:
    all_words = []
    for comment in comments:
        words = comment.lower().split()
        all_words.extend([w for w in words if len(w) >= 3])

    word_counts = Counter(all_words)
    top_words = word_counts.most_common(20)

    return {
        "total_comments": len(comments),
        "total_words": len(all_words),
        "unique_words": len(word_counts),
        "top_words": top_words,
        "avg_words_per_comment": round(len(all_words) / len(comments), 1) if comments else 0,
    }


def create_stats_image(stats: dict, title: str) -> BytesIO | None:
    if not MATPLOTLIB_INSTALLED or not stats["top_words"]:
        return None

    try:
        plt.figure(figsize=(12, 8))
        plt.style.use('seaborn-v0_8-darkgrid')

        words = [w[0] for w in stats["top_words"][:15]]
        counts = [w[1] for w in stats["top_words"][:15]]

        colors = plt.cm.viridis([i / len(words) for i in range(len(words))])
        plt.barh(words[::-1], counts[::-1], color=colors[::-1])
        plt.xlabel('Soni', fontsize=12)
        plt.ylabel("So'zlar", fontsize=12)
        plt.title(f"📊 Eng ko'p ishlatiladigan so'zlar\n{title}", fontsize=14, fontweight='bold')

        for i, (word, count) in enumerate(zip(words[::-1], counts[::-1])):
            plt.text(count + 0.5, i, str(count), va='center', fontsize=10)

        plt.tight_layout()
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        buf.seek(0)
        plt.close()
        return buf
    except Exception as e:
        print(f"Grafik xatosi: {e}")
        return None


# ==================== EXPORT FUNCTIONS ====================

def export_to_csv(comments: list, title: str) -> Path:
    file_path = TEMP_DIR / f"{title}_comments.csv"
    with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["№", "Izoh"])
        for i, comment in enumerate(comments, 1):
            writer.writerow([i, comment])
    return file_path


def export_to_excel(comments: list, stats: dict, title: str) -> Path | None:
    if not EXCEL_INSTALLED:
        return None

    file_path = TEMP_DIR / f"{title}_comments.xlsx"
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Izohlar"
    ws1.append(["№", "Izoh"])
    for i, comment in enumerate(comments, 1):
        ws1.append([i, comment])

    ws2 = wb.create_sheet("Statistika")
    ws2.append(["Ko'rsatkich", "Qiymat"])
    ws2.append(["Jami izohlar", stats["total_comments"]])
    ws2.append(["Jami so'zlar", stats["total_words"]])
    ws2.append(["Noyob so'zlar", stats["unique_words"]])
    ws2.append(["O'rtacha so'z/izoh", stats["avg_words_per_comment"]])

    ws3 = wb.create_sheet("Top so'zlar")
    ws3.append(["So'z", "Soni"])
    for word, count in stats["top_words"]:
        ws3.append([word, count])

    wb.save(file_path)
    return file_path


# ==================== KEYBOARDS ====================

def get_main_menu_keyboard(has_api: bool = False, is_admin: bool = False):
    if has_api:
        keyboard = [
            [InlineKeyboardButton("📺 Kanal videolari", callback_data="action_channel")],
            [InlineKeyboardButton("📋 Playlist", callback_data="action_playlist")],
            [InlineKeyboardButton("💬 Izohlarni yig'ish", callback_data="action_comments")],
            [InlineKeyboardButton("📊 Sessiya/Quota holati", callback_data="action_status")],
            [InlineKeyboardButton("🔄 API kalitni yangilash", callback_data="action_setapi")],
        ]
        if is_admin:
            keyboard.append([InlineKeyboardButton("👥 Foydalanuvchilar (Admin)", callback_data="action_users")])
        keyboard.append([InlineKeyboardButton("❓ Yordam", callback_data="action_help")])
    else:
        keyboard = [
            [InlineKeyboardButton("🔑 API kalitni kiritish", callback_data="action_setapi")],
            [InlineKeyboardButton("📚 API olish yo'riqnomasi", callback_data="action_guide")],
        ]
        if is_admin:
            keyboard.append([InlineKeyboardButton("👥 Foydalanuvchilar (Admin)", callback_data="action_users")])
        keyboard.append([InlineKeyboardButton("❓ Yordam", callback_data="action_help")])
    return InlineKeyboardMarkup(keyboard)


def get_export_keyboard():
    keyboard = [
        [InlineKeyboardButton("📄 TXT", callback_data="export_txt")],
        [InlineKeyboardButton("📊 CSV", callback_data="export_csv")],
    ]
    if EXCEL_INSTALLED:
        keyboard.append([InlineKeyboardButton("📗 Excel", callback_data="export_excel")])
    keyboard.append([InlineKeyboardButton("📦 Barchasi", callback_data="export_all")])
    keyboard.append([InlineKeyboardButton("🔙 Orqaga", callback_data="action_back")])
    return InlineKeyboardMarkup(keyboard)


def get_back_keyboard():
    return InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga", callback_data="action_back")]])


def get_cancel_keyboard():
    return InlineKeyboardMarkup([[InlineKeyboardButton("❌ Bekor qilish", callback_data="action_cancel")]])


def get_user_management_keyboard(users: list):
    keyboard = []
    for user_id, username, is_approved, is_admin, created_at in users[:10]:
        status = "✅" if is_approved else "⏳"
        admin_badge = "👑" if is_admin else ""
        label = f"{status} {username or user_id} {admin_badge}"
        keyboard.append([InlineKeyboardButton(label, callback_data=f"user_{user_id}")])
    keyboard.append([InlineKeyboardButton("🔙 Orqaga", callback_data="action_back")])
    return InlineKeyboardMarkup(keyboard)


def get_user_action_keyboard(user_id: int, is_admin_target: bool):
    keyboard = [
        [InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"approve_{user_id}")],
        [InlineKeyboardButton("❌ Rad etish", callback_data=f"reject_{user_id}")],
    ]
    if not is_admin_target:
        keyboard.append([InlineKeyboardButton("👑 Admin qilish", callback_data=f"makeadmin_{user_id}")])
    else:
        keyboard.append([InlineKeyboardButton("👤 Adminlikdan olish", callback_data=f"removeadmin_{user_id}")])
    keyboard.append([InlineKeyboardButton("🔙 Orqaga", callback_data="action_users")])
    return InlineKeyboardMarkup(keyboard)


def get_comments_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Ha, izohlarni yig'ish", callback_data="start_comments")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="action_back")]
    ])


# ==================== TEXTS ====================

WELCOME_TEXT = """
🎬 *YouTube Comments Scraper Bot v3*

✨ *Imkoniyatlar:*
• API quota monitoring
• Xatolikka bardoshlilik
• Dublikat izohlarni olib tashlash
• TXT/CSV/Excel eksport
• Statistika va grafiklar
• Playlist qo'llab-quvvatlash

⏰ Sessiya muddati: *12 soat*
"""

PENDING_APPROVAL_TEXT = """
⏳ *Tasdiqlash kutilmoqda*

Sizning so'rovingiz admin tomonidan ko'rib chiqilmoqda.

Admin tasdiqlagandan so'ng botdan foydalanishingiz mumkin bo'ladi.
Tasdiqlangach, sizga xabar keladi. 🔔
"""

API_GUIDE_TEXT = """
📚 *Google YouTube API kalitini olish*

*1.* https://console.cloud.google.com/ ga kiring
*2.* Yangi loyiha yarating
*3.* APIs & Services → Library
*4.* "YouTube Data API v3" ni yoqing
*5.* APIs & Services → Credentials
*6.* Create Credentials → API Key
*7.* Kalitni botga yuboring

⚠️ Kunlik limit: *10,000* so'rov
"""

HELP_TEXT = """
❓ *Yordam*

*Qanday ishlaydi:*
1. API kalitingizni kiriting
2. Kanal yoki Playlist tanlang
3. Video havolalar yuklanadi
4. Izohlarni yig'ing
5. Format tanlang (TXT/CSV/Excel)

*Kanal formatlari:*
• @username
• https://youtube.com/@username

*Playlist formatlari:*
• https://youtube.com/playlist?list=PLxxxxx
• PLxxxxx (faqat ID)

*Statistika:*
• Eng ko'p ishlatiladigan so'zlar
• Izohlar soni grafigi
"""


# ==================== HANDLERS ====================
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    
    # AVVAL tekshir, KEYIN saqlash
    new_registration = is_new_user(user.id)

    if new_registration:
        save_user(user.id, "", user.username)
        user_data = get_user(user.id)

        if user_data["is_admin"]:
            await update.message.reply_text(
                "👑 *Xush kelibsiz, Admin!*\n\nSiz birinchi foydalanuvchisiz.",
                parse_mode="Markdown",
                reply_markup=get_main_menu_keyboard(False, True),
            )
        else:
            # Xabar yuborilishini tekshirish uchun print qo'shing
            print(f"Yangi foydalanuvchi: {user.id}, adminlarga xabar yuborilmoqda...")
            await notify_admins_new_user(
                context.bot, user.id, user.username, user.full_name
            )
            print(f"Xabar yuborildi!")
            await update.message.reply_text(
                PENDING_APPROVAL_TEXT,
                parse_mode="Markdown",
            )
        return ConversationHandler.END

    # Mavjud foydalanuvchi
    user_data = get_user(user.id)

    if not user_data["is_approved"]:
        await update.message.reply_text(PENDING_APPROVAL_TEXT, parse_mode="Markdown")
        return ConversationHandler.END

    has_valid_session = bool(user_data["api_key"]) and is_session_valid(user.id)
    await update.message.reply_text(
        WELCOME_TEXT,
        parse_mode="Markdown",
        reply_markup=get_main_menu_keyboard(has_valid_session, bool(user_data["is_admin"])),
    )
    return ConversationHandler.END


async def callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user = update.effective_user
    user_data = get_user(user.id)
    has_valid_session = bool(user_data and user_data["api_key"] and is_session_valid(user.id))
    is_admin = bool(user_data["is_admin"]) if user_data else False
    action = query.data

    # ── Umumiy navigatsiya ──────────────────────────────────────────────────
    if action == "action_back":
        await query.message.edit_text(
            WELCOME_TEXT,
            parse_mode="Markdown",
            reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
        )
        return ConversationHandler.END

    elif action == "action_cancel":
        context.user_data.clear()
        await query.message.edit_text(
            "❌ Bekor qilindi\n\n" + WELCOME_TEXT,
            parse_mode="Markdown",
            reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
        )
        return ConversationHandler.END

    elif action == "action_guide":
        await query.message.edit_text(
            API_GUIDE_TEXT, parse_mode="Markdown", reply_markup=get_back_keyboard()
        )

    elif action == "action_help":
        await query.message.edit_text(
            HELP_TEXT, parse_mode="Markdown", reply_markup=get_back_keyboard()
        )

    elif action == "action_setapi":
        await query.message.edit_text(
            "🔑 *API kalitni kiriting*\n\n"
            "Google Cloud Console'dan olgan API kalitingizni yuboring:",
            parse_mode="Markdown",
            reply_markup=get_cancel_keyboard(),
        )
        return WAITING_API_KEY

    elif action == "action_status":
        if has_valid_session:
            remaining = get_session_remaining(user.id)
            quota_used = get_quota_used(user.id)
            quota_remaining = max(0, 10000 - quota_used)
            await query.message.edit_text(
                f"📊 *Sessiya va Quota holati*\n\n"
                f"✅ Sessiya: faol\n"
                f"⏰ Qolgan vaqt: {remaining}\n\n"
                f"📈 *API Quota:*\n"
                f"• Ishlatilgan: {quota_used} / 10,000\n"
                f"• Qolgan: {quota_remaining}\n"
                f"• Foiz: {int((quota_used / 10000) * 100)}%\n\n"
                f"🔑 API: ...{user_data['api_key'][-8:]}",
                parse_mode="Markdown",
                reply_markup=get_back_keyboard(),
            )
        else:
            await query.message.edit_text(
                "❌ Sessiya tugagan yoki API kalit kiritilmagan",
                reply_markup=get_main_menu_keyboard(False, is_admin),
            )

    # ── Kanal / Playlist / Izohlar ──────────────────────────────────────────
    elif action == "action_channel":
        if not has_valid_session:
            await query.message.edit_text(
                "❌ Avval API kalitni kiriting!",
                reply_markup=get_main_menu_keyboard(False, is_admin),
            )
            return ConversationHandler.END
        await query.message.edit_text(
            "📺 *Kanal nomini kiriting*\n\n"
            "Formatlar:\n• `@username`\n• `https://youtube.com/@username`",
            parse_mode="Markdown",
            reply_markup=get_cancel_keyboard(),
        )
        return WAITING_CHANNEL

    elif action == "action_playlist":
        if not has_valid_session:
            await query.message.edit_text(
                "❌ Avval API kalitni kiriting!",
                reply_markup=get_main_menu_keyboard(False, is_admin),
            )
            return ConversationHandler.END
        await query.message.edit_text(
            "📋 *Playlist havolasini kiriting*\n\n"
            "Formatlar:\n"
            "• `https://youtube.com/playlist?list=PLxxxxx`\n"
            "• `PLxxxxx` (faqat ID)",
            parse_mode="Markdown",
            reply_markup=get_cancel_keyboard(),
        )
        return WAITING_PLAYLIST

    elif action == "action_comments":
        if not has_valid_session:
            await query.message.edit_text(
                "❌ Avval API kalitni kiriting!",
                reply_markup=get_main_menu_keyboard(False, is_admin),
            )
            return ConversationHandler.END
        video_links = context.user_data.get("video_links", [])
        if not video_links:
            await query.message.edit_text(
                "❌ Video havolalar topilmadi.\n\nAvval kanal yoki playlist tanlang.",
                reply_markup=get_main_menu_keyboard(True, is_admin),
            )
            return ConversationHandler.END
        title = context.user_data.get("title", "Kanal")
        await query.message.edit_text(
            f"💬 *Izohlarni yig'ish*\n\n"
            f"📺 {title}\n"
            f"🎬 Videolar: {len(video_links)} ta\n\n"
            f"Davom etasizmi?",
            parse_mode="Markdown",
            reply_markup=get_comments_keyboard(),
        )

    elif action == "start_comments":
        await process_all_comments(query, context, user.id)

    elif action.startswith("export_"):
        export_format = action.replace("export_", "")
        await export_comments(query, context, user.id, export_format)

    # ── Admin: foydalanuvchilar ro'yxati ────────────────────────────────────
    elif action == "action_users":
        if not is_admin:
            await query.message.edit_text(
                "❌ Sizda admin huquqi yo'q",
                reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
            )
            return ConversationHandler.END
        users = get_all_users()
        if not users:
            await query.message.edit_text(
                "📭 Foydalanuvchilar yo'q",
                reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
            )
            return ConversationHandler.END
        await query.message.edit_text(
            f"👥 *Foydalanuvchilar ({len(users)} ta)*\n\nBoshqarish uchun foydalanuvchini tanlang:",
            parse_mode="Markdown",
            reply_markup=get_user_management_keyboard(users),
        )

    elif action.startswith("user_"):
        if not is_admin:
            return ConversationHandler.END
        target_id = int(action.split("_")[1])
        target = get_user(target_id)
        if not target:
            await query.message.edit_text(
                "❌ Foydalanuvchi topilmadi",
                reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
            )
            return ConversationHandler.END
        await query.message.edit_text(
            f"👤 *Foydalanuvchi: {target['username'] or target_id}*\n\n"
            f"ID: {target_id}\n"
            f"API: {'✅' if target['api_key'] else '❌'}\n"
            f"Tasdiqlangan: {'✅' if target['is_approved'] else '⏳'}\n"
            f"Admin: {'✅' if target['is_admin'] else '❌'}\n\n"
            f"Harakatni tanlang:",
            parse_mode="Markdown",
            reply_markup=get_user_action_keyboard(target_id, bool(target["is_admin"])),
        )

    # ── Admin: tasdiqlash / rad etish / admin o'zgartirish ─────────────────
    elif action.startswith("approve_"):
        if not is_admin:
            return ConversationHandler.END
        target_id = int(action.split("_")[1])
        approve_user(target_id)
        # Foydalanuvchiga xabar yuborish
        await notify_user_approved(context.bot, target_id)
        await query.message.edit_text(
            f"✅ Foydalanuvchi {target_id} tasdiqlandi va xabardor qilindi.",
            reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
        )

    elif action.startswith("reject_"):
        if not is_admin:
            return ConversationHandler.END
        target_id = int(action.split("_")[1])
        await notify_user_rejected(context.bot, target_id)
        reject_user(target_id)
        await query.message.edit_text(
            f"❌ Foydalanuvchi {target_id} rad etildi va xabardor qilindi.",
            reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
        )

    elif action.startswith("makeadmin_"):
        if not is_admin:
            return ConversationHandler.END
        target_id = int(action.split("_")[1])
        make_admin(target_id)
        await query.message.edit_text(
            f"👑 Foydalanuvchi {target_id} admin bo'ldi.",
            reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
        )

    elif action.startswith("removeadmin_"):
        if not is_admin:
            return ConversationHandler.END
        target_id = int(action.split("_")[1])
        remove_admin(target_id)
        await query.message.edit_text(
            f"👤 Foydalanuvchi {target_id} adminlikdan olindi.",
            reply_markup=get_main_menu_keyboard(has_valid_session, is_admin),
        )

    return ConversationHandler.END


async def handle_api_key(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    api_key = update.message.text.strip()

    msg = await update.message.reply_text("◐ API kalit tekshirilmoqda...")

    valid, message = validate_api_key(api_key)

    if valid:
        save_user(user.id, api_key, user.username)
        user_data = get_user(user.id)
        _is_admin = bool(user_data["is_admin"]) if user_data else False

        await msg.edit_text(
            f"✅ *API kalit saqlandi!*\n\n"
            f"⏰ Sessiya: {SESSION_HOURS} soat\n"
            f"📊 Quota: 10,000 so'rov\n\n"
            f"Endi quyidagilardan birini tanlang:",
            parse_mode="Markdown",
            reply_markup=get_main_menu_keyboard(True, _is_admin),
        )
    else:
        await msg.edit_text(
            f"❌ {message}\n\nQaytadan urinib ko'ring:",
            reply_markup=get_cancel_keyboard(),
        )
        return WAITING_API_KEY

    return ConversationHandler.END


async def handle_channel_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_data = get_user(user.id)
    is_admin = bool(user_data["is_admin"]) if user_data else False
    q = update.message.text.strip()

    msg = await update.message.reply_text("◐ Kanal qidirilmoqda...")
    channel_info, error = get_channel_info(user_data["api_key"], q)

    if error:
        await msg.edit_text(
            f"❌ {error}\n\nQaytadan urinib ko'ring:", reply_markup=get_cancel_keyboard()
        )
        return WAITING_CHANNEL

    await msg.edit_text(
        f"📺 *{channel_info['title']}*\n"
        f"🎬 Videolar: {channel_info['video_count']} ta\n\n"
        f"◐ Video havolalar yuklanmoqda...",
        parse_mode="Markdown",
    )

    try:
        links = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: get_all_video_links_sync(user_data["api_key"], channel_info["uploads_playlist"]),
        )
        update_quota(user.id, len(links) // 50 + 1)
        context.user_data["video_links"] = links
        context.user_data["title"] = channel_info["title"]

        file_path = TEMP_DIR / f"{user.id}_videos.txt"
        with open(file_path, "w", encoding="utf-8") as f:
            for link in links:
                f.write(link + "\n")

        await msg.edit_text(
            f"✅ *Yuklab olindi!*\n\n📺 {channel_info['title']}\n🎬 Videolar: {len(links)} ta",
            parse_mode="Markdown",
        )
        await update.message.reply_document(
            document=open(file_path, "rb"),
            filename=f"{channel_info['title']}_videos.txt",
            caption=f"📺 {len(links)} ta video",
        )
        await update.message.reply_text(
            "💬 Izohlarni yig'ish uchun tugmani bosing:",
            reply_markup=get_comments_keyboard(),
        )
    except Exception as e:
        await msg.edit_text(
            f"❌ Xato: {str(e)[:100]}", reply_markup=get_main_menu_keyboard(True, is_admin)
        )

    return ConversationHandler.END


async def handle_playlist_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_data = get_user(user.id)
    is_admin = bool(user_data["is_admin"]) if user_data else False
    q = update.message.text.strip()

    msg = await update.message.reply_text("◐ Playlist qidirilmoqda...")
    playlist_info, error = get_playlist_info(user_data["api_key"], q)

    if error:
        await msg.edit_text(
            f"❌ {error}\n\nQaytadan urinib ko'ring:", reply_markup=get_cancel_keyboard()
        )
        return WAITING_PLAYLIST

    await msg.edit_text(
        f"📋 *{playlist_info['title']}*\n"
        f"🎬 Videolar: {playlist_info['video_count']} ta\n\n"
        f"◐ Video havolalar yuklanmoqda...",
        parse_mode="Markdown",
    )

    try:
        links = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: get_all_video_links_sync(user_data["api_key"], playlist_info["id"]),
        )
        update_quota(user.id, len(links) // 50 + 1)
        context.user_data["video_links"] = links
        context.user_data["title"] = playlist_info["title"]

        file_path = TEMP_DIR / f"{user.id}_videos.txt"
        with open(file_path, "w", encoding="utf-8") as f:
            for link in links:
                f.write(link + "\n")

        await msg.edit_text(
            f"✅ *Yuklab olindi!*\n\n📋 {playlist_info['title']}\n🎬 Videolar: {len(links)} ta",
            parse_mode="Markdown",
        )
        await update.message.reply_document(
            document=open(file_path, "rb"),
            filename=f"{playlist_info['title']}_videos.txt",
            caption=f"📋 {len(links)} ta video",
        )
        await update.message.reply_text(
            "💬 Izohlarni yig'ish uchun tugmani bosing:",
            reply_markup=get_comments_keyboard(),
        )
    except Exception as e:
        await msg.edit_text(
            f"❌ Xato: {str(e)[:100]}", reply_markup=get_main_menu_keyboard(True, is_admin)
        )

    return ConversationHandler.END


async def process_all_comments(query, context, user_id: int):
    user_data = get_user(user_id)
    video_links = context.user_data.get("video_links", [])
    title = context.user_data.get("title", "videos")
    total = len(video_links)
    is_admin = bool(user_data["is_admin"]) if user_data else False

    msg = await query.message.edit_text(
        f"💬 *Izohlar yig'ilmoqda...*\n\n"
        f"📺 {title}\n[░░░░░░░░░░] 0%\n📊 0/{total}",
        parse_mode="Markdown",
    )

    all_comments = []
    seen_comments = set()
    errors = []
    total_quota = 0

    for i, link in enumerate(video_links, 1):
        match = re.search(r'(?:youtube\.com/watch\?v=|youtu\.be/)([a-zA-Z0-9_-]{11})', link)
        if not match:
            errors.append(f"{link} - noto'g'ri format")
            continue

        video_id = match.group(1)
        comments, error, quota_units = get_video_comments(user_data["api_key"], video_id)
        total_quota += quota_units

        if error:
            errors.append(f"{video_id} - {error}")
        else:
            for comment in comments:
                h = hash(comment.lower().strip())
                if h not in seen_comments:
                    seen_comments.add(h)
                    all_comments.append(comment)

        if i % 3 == 0 or i == total:
            percent = int((i / total) * 100)
            bar_filled = int(percent / 10)
            progress_bar = "█" * bar_filled + "░" * (10 - bar_filled)
            frame = LOADING_FRAMES[i % len(LOADING_FRAMES)]
            try:
                await msg.edit_text(
                    f"{frame} *Izohlar yig'ilmoqda...*\n\n"
                    f"📺 {title}\n"
                    f"[{progress_bar}] {percent}%\n"
                    f"📊 Videolar: {i}/{total}\n"
                    f"💬 Izohlar: {len(all_comments)}\n"
                    f"🔄 Dublikatlar olib tashlandi\n"
                    f"⚠️ Xatolar: {len(errors)}",
                    parse_mode="Markdown",
                )
            except:
                pass

        await asyncio.sleep(0.1)

    update_quota(user_id, total_quota)

    if all_comments:
        stats = generate_statistics(all_comments, title)
        context.user_data["comments"] = all_comments
        context.user_data["stats"] = stats
        context.user_data["errors"] = errors

        await msg.edit_text(
            f"✅ *Yakunlandi!*\n\n"
            f"📺 {title}\n"
            f"🎬 Videolar: {total} ta\n"
            f"💬 Izohlar: {len(all_comments)} ta\n"
            f"🔄 Dublikatlar olib tashlandi\n"
            f"⚠️ Xatolar: {len(errors)} ta\n\n"
            f"📊 *Statistika:*\n"
            f"• Jami so'zlar: {stats['total_words']}\n"
            f"• Noyob so'zlar: {stats['unique_words']}\n"
            f"• O'rtacha: {stats['avg_words_per_comment']} so'z/izoh\n\n"
            f"📤 *Eksport formatini tanlang:*",
            parse_mode="Markdown",
            reply_markup=get_export_keyboard(),
        )
    else:
        await msg.edit_text(
            "📭 Hech qanday izoh topilmadi",
            reply_markup=get_main_menu_keyboard(True, is_admin),
        )


async def export_comments(query, context, user_id: int, format_type: str):
    comments = context.user_data.get("comments", [])
    stats = context.user_data.get("stats", {})
    errors = context.user_data.get("errors", [])
    title = context.user_data.get("title", "comments")
    user_data = get_user(user_id)
    is_admin = bool(user_data["is_admin"]) if user_data else False

    if not comments:
        await query.message.edit_text(
            "❌ Izohlar topilmadi",
            reply_markup=get_main_menu_keyboard(True, is_admin),
        )
        return

    safe_title = re.sub(r'[^\w\s-]', '', title)[:30]
    msg = await query.message.edit_text("📤 Fayllar tayyorlanmoqda...")

    files_to_send = []

    if format_type in ["txt", "all"]:
        file_path = TEMP_DIR / f"{safe_title}_comments.txt"
        with open(file_path, "w", encoding="utf-8") as f:
            for comment in comments:
                f.write(f"{comment}\n")
        files_to_send.append((file_path, f"{safe_title}_comments.txt", "📄 TXT"))

    if format_type in ["csv", "all"]:
        file_path = export_to_csv(comments, safe_title)
        files_to_send.append((file_path, f"{safe_title}_comments.csv", "📊 CSV"))

    if format_type in ["excel", "all"] and EXCEL_INSTALLED:
        file_path = export_to_excel(comments, stats, safe_title)
        if file_path:
            files_to_send.append((file_path, f"{safe_title}_comments.xlsx", "📗 Excel"))

    if MATPLOTLIB_INSTALLED and stats:
        stats_img = create_stats_image(stats, title)
        if stats_img:
            await query.message.reply_photo(
                photo=stats_img, caption=f"📊 Statistika: {title}"
            )

    for file_path, filename, caption in files_to_send:
        await query.message.reply_document(
            document=open(file_path, "rb"),
            filename=filename,
            caption=f"{caption}: {len(comments)} ta izoh",
        )

    if errors:
        error_path = TEMP_DIR / f"{safe_title}_errors.txt"
        with open(error_path, "w", encoding="utf-8") as f:
            for error in errors:
                f.write(f"{error}\n")
        await query.message.reply_document(
            document=open(error_path, "rb"),
            filename="errors.txt",
            caption=f"⚠️ {len(errors)} ta xato",
        )

    await msg.edit_text(
        f"✅ Eksport yakunlandi!\n\n📁 Yuborilgan fayllar: {len(files_to_send)}",
        reply_markup=get_main_menu_keyboard(True, is_admin),
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_data = get_user(user.id)
    is_admin = bool(user_data["is_admin"]) if user_data else False

    if not user_data or not user_data["api_key"] or not is_session_valid(user.id):
        await update.message.reply_text(
            "❌ Avval API kalitni kiriting!",
            reply_markup=get_main_menu_keyboard(False, is_admin),
        )
        return

    document = update.message.document
    if not document.file_name.endswith(".txt"):
        await update.message.reply_text("❌ Faqat .txt fayl qabul qilinadi")
        return

    file = await document.get_file()
    file_path = TEMP_DIR / f"{user.id}_input.txt"
    await file.download_to_drive(file_path)

    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    video_links = [
        line.strip() for line in lines if line.strip() and "youtube" in line.lower()
    ]

    if not video_links:
        await update.message.reply_text("❌ Faylda YouTube havolalar topilmadi")
        return

    context.user_data["video_links"] = video_links
    context.user_data["title"] = "uploaded"

    await update.message.reply_text(
        f"✅ *{len(video_links)}* ta video havolasi topildi.\n\nIzohlarni yig'ishni boshlaysizmi?",
        parse_mode="Markdown",
        reply_markup=get_comments_keyboard(),
    )


# ==================== MAIN ====================

def main():
    init_db()
    print("🤖 Bot v3 ishga tushmoqda...")

    if not EXCEL_INSTALLED:
        print("⚠️ openpyxl o'rnatilmagan — Excel eksport ishlamaydi")
        print("   O'rnatish: pip install openpyxl")

    if not MATPLOTLIB_INSTALLED:
        print("⚠️ matplotlib o'rnatilmagan — grafiklar ishlamaydi")
        print("   O'rnatish: pip install matplotlib")

    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", start_command),
            CallbackQueryHandler(callback_handler),
        ],
        states={
            WAITING_API_KEY: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_api_key),
                CallbackQueryHandler(callback_handler),
            ],
            WAITING_CHANNEL: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_channel_input),
                CallbackQueryHandler(callback_handler),
            ],
            WAITING_PLAYLIST: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_playlist_input),
                CallbackQueryHandler(callback_handler),
            ],
        },
        fallbacks=[
            CommandHandler("start", start_command),
            CallbackQueryHandler(callback_handler),
        ],
        per_message=False,
    )

    app.add_handler(conv_handler)
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    print("✅ Bot v3 tayyor!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()