import telebot
from telebot import types
import json
import os
import re
import time
import threading
import requests
import phonenumbers
import random
import csv
import io
import tempfile
import openpyxl
import xlrd
from bs4 import BeautifulSoup
from phonenumbers import region_code_for_number, geocoder

_PID_FILE = "/tmp/ar_otp_bot.pid"
_my_pid = os.getpid()
if os.path.exists(_PID_FILE):
    try:
        _old_pid = int(open(_PID_FILE).read().strip())
        if _old_pid != _my_pid:
            try:
                os.kill(_old_pid, 9)
                time.sleep(1)
                print(f"[START] Killed old instance PID {_old_pid}")
            except ProcessLookupError:
                pass
    except Exception:
        pass
open(_PID_FILE, "w").write(str(_my_pid))

API_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
ADMIN_IDS = [6664150885]
CHANNEL_2 = "https://t.me/mailbotnewsofficial"

# â”€â”€ Panel 1 (Mahofuza) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
P1_BASE_URL = "http://91.232.105.47/ints"
P1_LOGIN_PAGE = P1_BASE_URL + "/login"
P1_SIGNIN_URL = P1_BASE_URL + "/signin"
P1_CDR_PAGE = P1_BASE_URL + "/agent/SMSCDRStats"
P1_CDR_DATA_URL = P1_BASE_URL + "/agent/res/data_smscdr.php"
P1_USER_NAME = "Mahofuza"
P1_PASSWORD = "Mahofuza"

# â”€â”€ Panel 2 (Sagardas50 / XISORA) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
P2_BASE_URL = "http://94.23.31.29/sms"
P2_SIGNIN_URL = P2_BASE_URL + "/signmein"
P2_REPORTS_PAGE = P2_BASE_URL + "/client/Reports"
P2_DATA_URL = P2_BASE_URL + "/client/ajax/dt_reports.php"
P2_USER_NAME = "Sagardas50"
P2_PASSWORD = "Sagardas50"

# â”€â”€ Panel 3 (Rabbi1_FD) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
P3_BASE_URL = "http://168.119.13.175/ints"
P3_LOGIN_PAGE = P3_BASE_URL + "/login"
P3_SIGNIN_URL = P3_BASE_URL + "/signin"
P3_CDR_PAGE = P3_BASE_URL + "/agent/SMSCDRStats"
P3_CDR_DATA_URL = P3_BASE_URL + "/agent/res/data_smscdr.php"
P3_USER_NAME = "Rabbi1_FD"
P3_PASSWORD = "Rabbi1_FD"

# â”€â”€ Panel 4 (Rabbi12) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
P4_BASE_URL = "http://144.217.71.192/ints"
P4_LOGIN_PAGE = P4_BASE_URL + "/login"
P4_SIGNIN_URL = P4_BASE_URL + "/signin"
P4_CDR_PAGE = P4_BASE_URL + "/agent/SMSCDRStats"
P4_CDR_DATA_URL = P4_BASE_URL + "/agent/res/data_smscdr.php"
P4_USER_NAME = "Rabbi12"
P4_PASSWORD = "Rabbi12"

# â”€â”€ Panel 5 (Rabbi12_v2 / 51.75.144.178) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
P5_BASE_URL = "http://51.75.144.178/ints"
P5_LOGIN_PAGE = P5_BASE_URL + "/login"
P5_SIGNIN_URL = P5_BASE_URL + "/signin"
P5_CDR_PAGE = P5_BASE_URL + "/agent/SMSCDRStats"
P5_CDR_DATA_URL = P5_BASE_URL + "/agent/res/data_smscdr.php"
P5_USER_NAME = "Rabbi12"
P5_PASSWORD = "Rabbi12@"

# â”€â”€ Panel 6 (Sagardas50 / TrueSMS.net â€” SMSRanges) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
P6_BASE_URL = "https://truesms.net"
P6_LOGIN_PAGE = P6_BASE_URL + "/login"
P6_SIGNIN_URL = P6_BASE_URL + "/signin"
P6_CDR_PAGE = P6_BASE_URL + "/agent/SMSRanges"
P6_CDR_DATA_URL = P6_BASE_URL + "/agent/res/data_smsranges.php"
P6_USER_NAME = "Sagardas50"
P6_PASSWORD = "Sagardas50"


POLL_INTERVAL = 3  # seconds â€” real-time as possible
DATA_FILE = "stock_data.json"
USERS_FILE = "users.json"
SEEN_FILE = "seen_otps.json"

bot = telebot.TeleBot(API_TOKEN, threaded=True, num_threads=40)

# â”€â”€ Persistent helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path) as f:
                return json.load(f)
        except Exception:
            pass
    return default


def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


stock = load_json(
    DATA_FILE,
    {
        "whatsapp": {},
        "facebook": {},
        "telegram": {},
        "instagram": {},
        "pc clone": {},
        "binance": {},
    },
)
users = load_json(USERS_FILE, [])
seen_otps = load_json(SEEN_FILE, {})
USER_NAMES_FILE = "user_names.json"
user_names = load_json(USER_NAMES_FILE, {})

SUPER_ADMIN_ID = 6664150885
ADMINS_FILE = "admins.json"
_extra_admins = load_json(ADMINS_FILE, [])
for _aid in _extra_admins:
    if _aid not in ADMIN_IDS:
        ADMIN_IDS.append(_aid)


def save_admins():
    save_json(ADMINS_FILE, [a for a in ADMIN_IDS if a != SUPER_ADMIN_ID])


def add_admin(uid):
    if uid not in ADMIN_IDS:
        ADMIN_IDS.append(uid)
        save_admins()
        return True
    return False


def remove_admin(uid):
    if uid == SUPER_ADMIN_ID:
        return False
    if uid in ADMIN_IDS:
        ADMIN_IDS.remove(uid)
        save_admins()
        return True
    return False

GROUP_SETTINGS_FILE = "group_settings.json"
_group_settings = load_json(GROUP_SETTINGS_FILE, {
    "otp_group_id": -1003738666960,
    "otp_group_link": "https://t.me/aR_OTP_rcv",
    "auto_delete": True,
    "auto_delete_seconds": 3600,
    "channel2": "https://t.me/mailbotnewsofficial",
    "bot_link": "https://t.me/ar_otp_bot",
})

CHANNEL_1 = _group_settings["otp_group_link"]
OTP_GROUP_ID = _group_settings["otp_group_id"]


def save_group_settings():
    save_json(GROUP_SETTINGS_FILE, _group_settings)


def get_otp_group_id():
    return _group_settings.get("otp_group_id")


def get_otp_group_link():
    return _group_settings.get("otp_group_link", "")


def _extract_username(link):
    """Extract @username from a t.me link for use with get_chat_member."""
    if not link:
        return None
    link = link.strip().rstrip("/")
    if "joinchat" in link or "/+" in link:
        return None
    if "t.me/" in link:
        uname = link.split("t.me/")[-1].split("/")[0]
        if uname:
            return "@" + uname
    return None


def _check_member(chat_ref, user_id):
    """Returns True if member, False if not, None if cannot check."""
    if not chat_ref:
        return None
    try:
        m = bot.get_chat_member(chat_ref, user_id)
        return m.status not in ("left", "kicked")
    except Exception:
        return None


def get_channel2():
    return _group_settings.get("channel2", "https://t.me/mailbotnewsofficial")


def get_bot_link():
    return _group_settings.get("bot_link", "https://t.me/ar_otp_bot")


def is_auto_delete():
    return _group_settings.get("auto_delete", True)


def _schedule_delete(chat_id, msg_id):
    delay = _group_settings.get("auto_delete_seconds", 3600)
    def _do_delete():
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
    threading.Timer(delay, _do_delete).start()

SERVICES_FILE = "services.json"
_DEFAULT_SERVICES = [
    {"label": "Instagram â†’", "key": "instagram"},
    {"label": "Facebook ðŸ’Ž", "key": "facebook"},
    {"label": "WhatsApp", "key": "whatsapp"},
    {"label": "PC Clone ðŸ’Ž", "key": "pc clone"},
]
_services = load_json(SERVICES_FILE, list(_DEFAULT_SERVICES))
_addservice_state = {}
_countdowns = {}

user_map = {}
user_map_lock = threading.Lock()
assigned_time = {}


def register_number(user_id, number):
    clean = re.sub(r"\D", "", str(number))
    with user_map_lock:
        user_map[clean] = user_id
        assigned_time[clean] = time.time()


ADMIN_CONFIGS_DIR = "admin_configs"
os.makedirs(ADMIN_CONFIGS_DIR, exist_ok=True)

_ADMIN_CFG_DEFAULTS = {
    "brand": "RABBI",
    "numChannel": "",
    "mainChannel": "",
    "botLink": "",
    "group_id": None,
    "group_link": "",
}

def _admin_cfg_path(uid):
    return os.path.join(ADMIN_CONFIGS_DIR, f"{uid}.json")

def get_admin_config(uid):
    path = _admin_cfg_path(uid)
    cfg = load_json(path, {})
    result = dict(_ADMIN_CFG_DEFAULTS)
    result.update(cfg)
    return result

def save_admin_config(uid, cfg):
    save_json(_admin_cfg_path(uid), cfg)

def get_brand(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    return get_admin_config(uid).get("brand", "RABBI")

def get_num_channel(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    cfg = get_admin_config(uid)
    return cfg.get("numChannel") or _group_settings.get("bot_link", "")

def get_main_channel_url(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    cfg = get_admin_config(uid)
    return cfg.get("mainChannel") or get_channel2()

def get_admin_group_id(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    cfg = get_admin_config(uid)
    gid = cfg.get("group_id")
    if gid:
        return int(gid)
    return get_otp_group_id()

def get_admin_group_link(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    cfg = get_admin_config(uid)
    return cfg.get("group_link") or get_otp_group_link()

def get_admin_bot_link(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    cfg = get_admin_config(uid)
    return cfg.get("botLink") or get_bot_link()

def detect_service_from_sms(sms_text):
    t = sms_text.lower()
    if "telegram" in t: return "Telegram"
    if "facebook" in t or " fb " in t: return "Facebook"
    if "whatsapp" in t: return "WhatsApp"
    if "instagram" in t: return "Instagram"
    if "twitter" in t or "x.com" in t: return "Twitter/X"
    if "google" in t: return "Google"
    if "apple" in t or "icloud" in t: return "Apple"
    if "tiktok" in t: return "TikTok"
    if "discord" in t: return "Discord"
    if "twverify" in t: return "TWVerify"
    if "amazon" in t: return "Amazon"
    if "netflix" in t: return "Netflix"
    if "uber" in t: return "Uber"
    if "paypal" in t: return "PayPal"
    if "microsoft" in t: return "Microsoft"
    if "snapchat" in t: return "Snapchat"
    if "viber" in t: return "Viber"
    if "signal" in t: return "Signal"
    if "linkedin" in t: return "LinkedIn"
    if "yahoo" in t: return "Yahoo"
    if "wechat" in t: return "WeChat"
    return "SMS"

def mask_number(number, owner_uid=None):
    brand = get_brand(owner_uid)
    d = re.sub(r"\D", "", str(number))
    if len(d) <= 6:
        return d
    return f"{d[:3]}â˜…{brand}â˜…{d[-3:]}"


# â”€â”€ OTP Messages â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def send_otp_message(chat_id, otp, number, seconds, cli="", sms_txt="", owner_uid=None):
    if owner_uid is None:
        owner_uid = SUPER_ADMIN_ID
    brand = get_brand(owner_uid)
    service = detect_service_from_sms(sms_txt) if sms_txt else (cli.upper() if cli else "SMS")
    c_name, flag = get_country_details(number)
    cli_line = f"\nðŸ“Ÿ CLI: <code>{cli}</code>" if cli and cli.strip() else ""

    message = (
        f"ðŸ” <b>{service.upper()} OTP RECEIVED</b> ðŸ”\n\n"
        f"{flag} {c_name}\n"
        f"ðŸ“ž <code>{mask_number(number, owner_uid)}</code>\n"
        f"ðŸ“± Service: <b>{service}</b>"
        f"{cli_line}\n\n"
        f"ðŸ’¬ OTP: <code>{otp}</code>\n"
        f"ðŸ‘‘{brand}"
    )

    num_ch = get_num_channel(owner_uid)
    main_ch = get_main_channel_url(owner_uid)
    markup = types.InlineKeyboardMarkup()
    row = []
    if num_ch:
        row.append(types.InlineKeyboardButton("ðŸ“² Numbers", url=num_ch))
    if main_ch:
        row.append(types.InlineKeyboardButton("ðŸ“¢ Main CH", url=main_ch))
    if row:
        markup.row(*row)

    try:
        sent = bot.send_message(
            chat_id=chat_id, text=message, parse_mode="HTML", reply_markup=markup if row else None
        )
        admin_gid = get_admin_group_id(owner_uid)
        if chat_id == admin_gid and is_auto_delete():
            _schedule_delete(chat_id, sent.message_id)
    except Exception as e:
        print(f"[MONITOR] Send error to {chat_id}: {e}")


def _dispatch_otp(otp, number, seconds, cli="", sms_txt="", owner_uid=None):
    if owner_uid is None:
        owner_uid = SUPER_ADMIN_ID
    group_id = get_admin_group_id(owner_uid)
    send_otp_message(group_id, otp, number, seconds, cli, sms_txt, owner_uid)
    clean = re.sub(r"\D", "", str(number))
    with user_map_lock:
        uid = user_map.pop(clean, None)
        assigned_time.pop(clean, None)
    if uid:
        send_otp_message(uid, otp, number, seconds, cli, sms_txt, owner_uid)


def send_status_message(chat_id, status_text):
    message = (
        "âš™ï¸ <b>ð—¦ð—§ð—”ð—§ð—¨ð—¦ ð—”ð—Ÿð—˜ð—¥ð—§</b> âš™ï¸\n"
        "ðŸ”¥â”â”â”â”â”â”â”â”â”â”â”â”â”â”ðŸ”¥\n\n"
        f"ðŸ“› {status_text} ðŸ“›\n\n"
        "ðŸ”¥â”â”â”â”â”â”â”â”â”â”â”â”â”â”ðŸ”¥\n"
        "ðŸ¤–âš¡ <b>ð—”ð—¥ ð—¢ð—§ð—£ ð—•ð—¢ð—§ â€” ð—”ð—–ð—§ð—œð—©ð—˜</b> âš¡ðŸ¤–"
    )
    try:
        bot.send_message(chat_id=chat_id, text=message, parse_mode="HTML")
    except Exception as e:
        print(f"[MONITOR] Status send error: {e}")


# â”€â”€ Country helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def get_country_details(num_str):
    try:
        num_str = str(num_str).strip()
        if not num_str.startswith("+"):
            num_str = "+" + num_str
        parsed = phonenumbers.parse(num_str)
        country_code = region_code_for_number(parsed)
        country_name = geocoder.description_for_number(parsed, "en")
        flag = "".join(chr(ord(c.upper()) + 127397) for c in country_code)
        return country_name, flag
    except Exception:
        return "Unknown", "ðŸŒ"


# â”€â”€ Stock helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def save_stock():
    save_json(DATA_FILE, stock)


def register_user(chat_id, first_name="", last_name="", username=""):
    if chat_id not in users:
        users.append(chat_id)
        save_json(USERS_FILE, users)
    full = f"{first_name} {last_name}".strip()
    if full and username:
        display = f"{full} (@{username})"
    elif full:
        display = full
    elif username:
        display = f"@{username}"
    else:
        display = None
    if display:
        user_names[str(chat_id)] = display
        save_json(USER_NAMES_FILE, user_names)


# â”€â”€ Panel sessions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

_p1_session = None
_p1_sesskey = None
_p1_lock = threading.Lock()

_p2_session = None
_p2_lock = threading.Lock()

_p3_session = None
_p3_csstr = None
_p3_lock = threading.Lock()

_p4_session = None
_p4_sesskey = None
_p4_lock = threading.Lock()

_p5_session = None
_p5_sesskey = None
_p5_lock = threading.Lock()

_p6_session = None
_p6_sesskey = None
_p6_lock = threading.Lock()


# â”€â”€ Panel stats (for /panels command) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
_panel_stats = {
    "p1": {
        "name": "Mahofuza",
        "host": "91.232.105.47",
        "status": "â³",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p2": {
        "name": "Sagardas50",
        "host": "94.23.31.29",
        "status": "â³",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p3": {
        "name": "Rabbi1_FD",
        "host": "168.119.13.175",
        "status": "â³",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p4": {
        "name": "Rabbi12",
        "host": "144.217.71.192",
        "status": "â³",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p5": {
        "name": "Rabbi12_v2",
        "host": "51.75.144.178",
        "status": "â³",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p6": {
        "name": "TrueSMS/Ranges",
        "host": "truesms.net",
        "status": "â³",
        "count": 0,
        "last": None,
        "errors": 0,
    },
}
_stats_lock = threading.Lock()


def _record_fetch(pid, count):
    with _stats_lock:
        _panel_stats[pid]["status"] = "ðŸŸ¢"
        _panel_stats[pid]["count"] = count
        _panel_stats[pid]["last"] = time.time()
        _panel_stats[pid]["errors"] = 0


def _record_error(pid):
    with _stats_lock:
        _panel_stats[pid]["status"] = "ðŸ”´"
        _panel_stats[pid]["errors"] += 1


# â”€â”€ Demo OTP state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
_demo_active = False
_demo_lock = threading.Lock()
_demo_config = {
    "numbers": ["8801700000000"],
    "digits": 6,
    "service": "Facebook",
    "interval": 30,
}

seen_lock = threading.Lock()

# â”€â”€ Dynamic panel system â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
DYNAMIC_PANELS_FILE = "dynamic_panels.json"
_dynamic_panels = load_json(DYNAMIC_PANELS_FILE, [])
_dynamic_sessions = {}
_dynamic_locks = {}
_addpanel_state = {}
_pending_excel = {}  # uid â†’ {'numbers': [...], 'filename': str}


def save_dynamic_panels():
    save_json(DYNAMIC_PANELS_FILE, _dynamic_panels)


def _get_dp_lock(pid):
    if pid not in _dynamic_locks:
        _dynamic_locks[pid] = threading.Lock()
    return _dynamic_locks[pid]


def _ints_login(panel):
    pid = panel["id"]
    base = panel["base_url"]
    panel_type = panel.get("panel_type", "smscdr")
    cdr_endpoint = "/agent/SMSRanges" if panel_type == "smsranges" else "/agent/SMSCDRStats"
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    owner_uid = panel.get("owner_uid", SUPER_ADMIN_ID)
    brand = get_brand(owner_uid)
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] ðŸš€ SMS Forwarder à¦¶à§à¦°à§ à¦¹à¦šà§à¦›à§‡... Brand: {brand}  [{pid} | {panel.get('host','')}]")
    try:
        r = sess.get(base + "/login", timeout=15, verify=False)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if m:
            answer = int(m.group(1)) + int(m.group(2))
            ts2 = time.strftime("%H:%M:%S")
            print(f"[{ts2}] ðŸ”‘ Captcha = {answer}  [{pid}]")
            r2 = sess.post(
                base + "/signin",
                data={
                    "username": panel["username"],
                    "password": panel["password"],
                    "capt": answer,
                },
                timeout=15,
                allow_redirects=True,
                verify=False,
            )
        else:
            r2 = sess.post(
                base + "/signin",
                data={"username": panel["username"], "password": panel["password"]},
                timeout=15,
                allow_redirects=True,
                verify=False,
            )
        if "login" in r2.url.lower() and "agent" not in r2.url.lower():
            ts3 = time.strftime("%H:%M:%S")
            print(f"[{ts3}] âŒ Login failed [{pid}]: {r2.url}")
            return None, None
        cdr_page = base + cdr_endpoint
        r3 = sess.get(cdr_page, timeout=15, headers={"Referer": base + "/agent/"}, verify=False)
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        cs = re.search(r"csstr=([a-f0-9]+)", r3.text)
        token = sk.group(1) if sk else (cs.group(1) if cs else "")
        ts4 = time.strftime("%H:%M:%S")
        print(f"[{ts4}] âœ… Login successful  [{pid} | user={panel['username']}]")
        return sess, token
    except Exception as e:
        ts5 = time.strftime("%H:%M:%S")
        print(f"[{ts5}] âŒ Login error [{pid}]: {e}")
        return None, None


def _ints_fetch(panel):
    pid = panel["id"]
    base = panel["base_url"]
    panel_type = panel.get("panel_type", "smscdr")
    if panel_type == "smsranges":
        data_url = base + "/agent/res/data_smsranges.php"
        cdr_page = base + "/agent/SMSRanges"
    else:
        data_url = base + "/agent/res/data_smscdr.php"
        cdr_page = base + "/agent/SMSCDRStats"
    found = {}
    with _get_dp_lock(pid):
        sd = _dynamic_sessions.get(pid, {})
        if not sd.get("session"):
            s, tok = _ints_login(panel)
            if not s:
                _record_error(pid)
                return found
            _dynamic_sessions[pid] = {"session": s, "token": tok}
            sd = _dynamic_sessions[pid]
        sess = sd["session"]
        token = sd.get("token", "")
        today = time.strftime("%Y-%m-%d")

        def build_url():
            return (
                f"{data_url}"
                f"?fdate1={today}%2000:00:00&fdate2={today}%2023:59:59"
                f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                f"&sesskey={token}"
            )

        headers = {"Referer": cdr_page, "X-Requested-With": "XMLHttpRequest"}
        try:
            r = sess.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[{pid}] Bad response, re-logging in.")
                _dynamic_sessions[pid] = {}
                s, tok = _ints_login(panel)
                if not s:
                    _record_error(pid)
                    return found
                _dynamic_sessions[pid] = {"session": s, "token": tok}
                r = s.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch(pid, len(rows))
            if found:
                print(f"[{pid}] âœ… Fetched {len(found)} records.")
        except Exception as e:
            print(f"[{pid}] Fetch error: {e}")
            _record_error(pid)
            _dynamic_sessions[pid] = {}
    return found


def _start_dynamic_panel(panel):
    pid = panel["id"]
    owner_uid = panel.get("owner_uid", SUPER_ADMIN_ID)
    with _stats_lock:
        _panel_stats[pid] = {
            "name": panel.get("username", pid),
            "host": panel.get("host", ""),
            "status": "â³",
            "count": 0,
            "last": None,
            "errors": 0,
            "owner_uid": owner_uid,
        }

    def monitor():
        global seen_otps
        existing = _ints_fetch(panel)
        with seen_lock:
            for key in existing:
                seen_otps[key] = True
            save_json(SEEN_FILE, seen_otps)
        ts = time.strftime("%H:%M:%S")
        print(f"[{ts}] ðŸ“¦ {len(existing)} à¦Ÿà¦¿ à¦ªà§à¦°à¦¾à¦¨à§‹ SMS à¦²à§‹à¦¡ à¦¹à¦¯à¦¼à§‡à¦›à§‡ (forward à¦¹à¦¬à§‡ à¦¨à¦¾) [{pid}]")
        while True:
            try:
                process_new_otps(_ints_fetch(panel), owner_uid)
            except Exception as e:
                print(f"[{pid}] Loop error: {e}")
            time.sleep(POLL_INTERVAL)

    threading.Thread(target=monitor, daemon=True).start()


def extract_otp_from_sms(sms_text):
    cleaned = re.sub(r"(?<=\d) (?=\d)", "", sms_text)
    m = re.search(r"\b(\d{4,8})\b", cleaned)
    return m.group(1) if m else None


# â”€â”€ Panel 1 login & fetch (Mahofuza) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def p1_login():
    global _p1_session, _p1_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P1_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P1] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P1_SIGNIN_URL,
            data={"username": P1_USER_NAME, "password": P1_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "login" in r2.url.lower() or "login" in r2.text.lower()[:500]:
            print("[P1] Login failed â€” still on login page")
            return False
        r3 = sess.get(
            P1_CDR_PAGE, timeout=15, headers={"Referer": P1_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p1_sesskey = sk.group(1) if sk else ""
        _p1_session = sess
        print(f"[P1] Logged in. sesskey={_p1_sesskey}")
        return True
    except Exception as e:
        print(f"[P1] Login error: {e}")
        return False


def fetch_panel1():
    global _p1_session, _p1_sesskey
    found = {}
    with _p1_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P1_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&sesskey={_p1_sesskey or ''}"
                )

            headers = {"Referer": P1_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p1_session is None:
                if not p1_login():
                    return found
            r = _p1_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P1] Bad response ({r.status_code}), re-logging in.")
                _p1_session = None
                if not p1_login():
                    return found
                r = _p1_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p1", len(rows))
            if found:
                print(f"[P1] âœ… Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P1] Fetch error: {e}")
            _record_error("p1")
            _p1_session = None
    return found


# â”€â”€ Panel 2 login & fetch (Sagardas50 / XISORA) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def p2_login():
    global _p2_session
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.post(
            P2_SIGNIN_URL,
            data={"username": P2_USER_NAME, "password": P2_PASSWORD},
            timeout=15,
            allow_redirects=True,
        )
        if "signin" in r.url.lower() or "login" in r.url.lower():
            print("[P2] Login failed â€” still on login page")
            return False
        _p2_session = sess
        print(f"[P2] Logged in. URL={r.url}")
        return True
    except Exception as e:
        print(f"[P2] Login error: {e}")
        return False


def fetch_panel2():
    global _p2_session
    found = {}
    with _p2_lock:
        try:
            today = time.strftime("%Y-%m-%d")
            url = (
                f"{P2_DATA_URL}"
                f"?fdate1={today}%2000:00:00"
                f"&fdate2={today}%2023:59:59"
                f"&ftermination=&fclient=&fnum=&fcli="
                f"&fgdate=0&fgtermination=0&fgclient=0&fgnumber=0&fgcli=0&fg=0"
            )
            headers = {"Referer": P2_REPORTS_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p2_session is None:
                if not p2_login():
                    return found
            r = _p2_session.get(url, headers=headers, timeout=15)
            body = r.text.strip()
            if r.status_code != 200 or not body or body.startswith("<"):
                print(f"[P2] Bad response ({r.status_code}), re-logging in.")
                _p2_session = None
                if not p2_login():
                    return found
                r = _p2_session.get(url, headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[10]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p2", len(rows))
            if found:
                print(f"[P2] âœ… Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P2] Fetch error: {e}")
            _record_error("p2")
            _p2_session = None
    return found


# â”€â”€ Shared OTP processor â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def process_new_otps(current, owner_uid=None):
    global seen_otps
    if owner_uid is None:
        owner_uid = SUPER_ADMIN_ID
    for key, (number, otp, sms_txt, cli) in current.items():
        with seen_lock:
            if key in seen_otps:
                continue
            seen_otps[key] = True
            save_json(SEEN_FILE, seen_otps)
        clean = re.sub(r"\D", "", str(number))
        with user_map_lock:
            t_start = assigned_time.get(clean)
        seconds = int(time.time() - t_start) if t_start else 0
        _dispatch_otp(otp, number, seconds, cli, sms_txt, owner_uid)
        service = detect_service_from_sms(sms_txt)
        ts = time.strftime("%H:%M:%S")
        brand = get_brand(owner_uid)
        print(
            f"[{ts}] ðŸ’¬ OTP FORWARD | {service} | {number} | {otp} | admin={owner_uid} | brand={brand}"
        )


# â”€â”€ Global OTP monitors â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def _panel_startup(label, fetch_fn, owner_uid):
    global seen_otps
    brand = get_brand(owner_uid)
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] ðŸš€ SMS Forwarder à¦¶à§à¦°à§ à¦¹à¦šà§à¦›à§‡... Brand: {brand}  [{label}]")
    existing = fetch_fn()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    ts2 = time.strftime("%H:%M:%S")
    print(f"[{ts2}] ðŸ“¦ {len(existing)} à¦Ÿà¦¿ à¦ªà§à¦°à¦¾à¦¨à§‹ SMS à¦²à§‹à¦¡ à¦¹à¦¯à¦¼à§‡à¦›à§‡ (forward à¦¹à¦¬à§‡ à¦¨à¦¾) [{label}]")
    return existing


def panel1_monitor():
    _panel_startup("P1-Mahofuza", fetch_panel1, SUPER_ADMIN_ID)
    while True:
        try:
            process_new_otps(fetch_panel1(), SUPER_ADMIN_ID)
        except Exception as e:
            print(f"[P1] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


def panel2_monitor():
    _panel_startup("P2-Sagardas50", fetch_panel2, SUPER_ADMIN_ID)
    while True:
        try:
            process_new_otps(fetch_panel2(), SUPER_ADMIN_ID)
        except Exception as e:
            print(f"[P2] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# â”€â”€ Panel 3 login & fetch (Rabbi1_FD) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def p3_login():
    global _p3_session, _p3_csstr
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P3_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P3] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P3_SIGNIN_URL,
            data={"username": P3_USER_NAME, "password": P3_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "login" in r2.url.lower() or "signin" in r2.url.lower():
            print("[P3] Login failed â€” still on login page")
            return False
        r3 = sess.get(
            P3_CDR_PAGE, timeout=15, headers={"Referer": P3_BASE_URL + "/agent/"}
        )
        cs = re.search(r"csstr=([a-f0-9]+)", r3.text)
        _p3_csstr = cs.group(1) if cs else ""
        _p3_session = sess
        print(f"[P3] Logged in. csstr={_p3_csstr}")
        return True
    except Exception as e:
        print(f"[P3] Login error: {e}")
        return False


def fetch_panel3():
    global _p3_session, _p3_csstr
    found = {}
    with _p3_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P3_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&csstr={_p3_csstr or ''}"
                )

            headers = {"Referer": P3_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p3_session is None:
                if not p3_login():
                    return found
            r = _p3_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P3] Bad response ({r.status_code}), re-logging in.")
                _p3_session = None
                if not p3_login():
                    return found
                r = _p3_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p3", len(rows))
            if found:
                print(f"[P3] âœ… Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P3] Fetch error: {e}")
            _record_error("p3")
            _p3_session = None
    return found


def panel3_monitor():
    _panel_startup("P3-Rabbi1_FD", fetch_panel3, SUPER_ADMIN_ID)
    while True:
        try:
            process_new_otps(fetch_panel3(), SUPER_ADMIN_ID)
        except Exception as e:
            print(f"[P3] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# â”€â”€ Panel 4 login & fetch (Rabbi12 / 144.217.71.192) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def p4_login():
    global _p4_session, _p4_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P4_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P4] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P4_SIGNIN_URL,
            data={"username": P4_USER_NAME, "password": P4_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "SMSDashboard" not in r2.url and "agent" not in r2.url:
            print(f"[P4] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P4_CDR_PAGE, timeout=15, headers={"Referer": P4_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p4_sesskey = sk.group(1) if sk else ""
        _p4_session = sess
        print(f"[P4] Logged in. sesskey={_p4_sesskey}")
        return True
    except Exception as e:
        print(f"[P4] Login error: {e}")
        return False


def fetch_panel4():
    global _p4_session, _p4_sesskey
    found = {}
    with _p4_lock:
        if not _p4_session and not p4_login():
            return found
        today = time.strftime("%Y-%m-%d")

        def build_url():
            return (
                f"{P4_CDR_DATA_URL}"
                f"?fdate1={today}%2000:00:00&fdate2={today}%2023:59:59"
                f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                f"&sesskey={_p4_sesskey}"
            )

        headers = {"Referer": P4_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
        try:
            r = _p4_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P4] Bad response ({r.status_code}), re-logging in.")
                _p4_session = None
                if not p4_login():
                    return found
                r = _p4_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p4", len(rows))
            if found:
                print(f"[P4] âœ… Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P4] Fetch error: {e}")
            _record_error("p4")
            _p4_session = None
    return found


def panel4_monitor():
    _panel_startup("P4-Rabbi12", fetch_panel4, SUPER_ADMIN_ID)
    while True:
        try:
            process_new_otps(fetch_panel4(), SUPER_ADMIN_ID)
        except Exception as e:
            print(f"[P4] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# â”€â”€ Panel 5 login & fetch (Rabbi12_v2 / 51.75.144.178) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def p5_login():
    global _p5_session, _p5_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P5_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P5] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P5_SIGNIN_URL,
            data={"username": P5_USER_NAME, "password": P5_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "SMSDashboard" not in r2.url and "agent" not in r2.url:
            print(f"[P5] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P5_CDR_PAGE, timeout=15, headers={"Referer": P5_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p5_sesskey = sk.group(1) if sk else ""
        _p5_session = sess
        print(f"[P5] Logged in. sesskey={_p5_sesskey}")
        return True
    except Exception as e:
        print(f"[P5] Login error: {e}")
        return False


def fetch_panel5():
    global _p5_session, _p5_sesskey
    found = {}
    with _p5_lock:
        if not _p5_session and not p5_login():
            return found
        today = time.strftime("%Y-%m-%d")

        def build_url():
            return (
                f"{P5_CDR_DATA_URL}"
                f"?fdate1={today}%2000:00:00&fdate2={today}%2023:59:59"
                f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                f"&sesskey={_p5_sesskey}"
            )

        headers = {"Referer": P5_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
        try:
            r = _p5_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P5] Bad response ({r.status_code}), re-logging in.")
                _p5_session = None
                if not p5_login():
                    return found
                r = _p5_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p5", len(rows))
            if found:
                print(f"[P5] âœ… Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P5] Fetch error: {e}")
            _record_error("p5")
            _p5_session = None
    return found


def panel5_monitor():
    _panel_startup("P5-Rabbi12_v2", fetch_panel5, SUPER_ADMIN_ID)
    while True:
        try:
            process_new_otps(fetch_panel5(), SUPER_ADMIN_ID)
        except Exception as e:
            print(f"[P5] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# â”€â”€ Panel 6 login & fetch (TrueSMS.net / SMSRanges) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def p6_login():
    global _p6_session, _p6_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P6_LOGIN_PAGE, timeout=20, verify=False)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if m:
            answer = int(m.group(1)) + int(m.group(2))
            r2 = sess.post(
                P6_SIGNIN_URL,
                data={
                    "username": P6_USER_NAME,
                    "password": P6_PASSWORD,
                    "capt": answer,
                },
                timeout=20,
                allow_redirects=True,
                verify=False,
            )
        else:
            r2 = sess.post(
                P6_SIGNIN_URL,
                data={"username": P6_USER_NAME, "password": P6_PASSWORD},
                timeout=20,
                allow_redirects=True,
                verify=False,
            )
        if "login" in r2.url.lower() and "agent" not in r2.url.lower():
            print(f"[P6] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P6_CDR_PAGE,
            timeout=20,
            headers={"Referer": P6_BASE_URL + "/agent/"},
            verify=False,
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        cs = re.search(r"csstr=([a-f0-9]+)", r3.text)
        _p6_sesskey = sk.group(1) if sk else (cs.group(1) if cs else "")
        _p6_session = sess
        print(f"[P6] Logged in. token={_p6_sesskey[:10] if _p6_sesskey else 'none'}")
        return True
    except Exception as e:
        print(f"[P6] Login error: {e}")
        return False


def fetch_panel6():
    global _p6_session, _p6_sesskey
    found = {}
    with _p6_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P6_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&sesskey={_p6_sesskey or ''}"
                )

            headers = {"Referer": P6_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p6_session is None:
                if not p6_login():
                    return found
            r = _p6_session.get(build_url(), headers=headers, timeout=20, verify=False)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P6] Bad response ({r.status_code}), re-logging in.")
                _p6_session = None
                if not p6_login():
                    return found
                r = _p6_session.get(
                    build_url(), headers=headers, timeout=20, verify=False
                )
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip() if len(row) > 3 else "TrueSMS"
                sms_txt = str(row[5]).strip() if len(row) > 5 else ""
                if not sms_txt and len(row) > 4:
                    sms_txt = str(row[4]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p6", len(rows))
            if found:
                print(f"[P6] âœ… Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P6] Fetch error: {e}")
            _record_error("p6")
            _p6_session = None
    return found


def panel6_monitor():
    _panel_startup("P6-TrueSMS", fetch_panel6, SUPER_ADMIN_ID)
    while True:
        try:
            process_new_otps(fetch_panel6(), SUPER_ADMIN_ID)
        except Exception as e:
            print(f"[P6] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# â”€â”€ Demo OTP monitor â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def demo_monitor():
    print("[DEMO] Thread started.")
    while True:
        with _demo_lock:
            active = _demo_active
            cfg = dict(_demo_config)
        if active:
            otp = "".join([str(random.randint(0, 9)) for _ in range(cfg["digits"])])
            number = random.choice(cfg["numbers"])
            try:
                send_otp_message(get_otp_group_id(), otp, number, 0, "", f"Your {cfg['service']} code is {otp}")
            except Exception as e:
                print(f"[DEMO] Send error: {e}")
        time.sleep(cfg["interval"])


def demo_status_text():
    with _demo_lock:
        active = _demo_active
        cfg = dict(_demo_config)
    status = "ðŸŸ¢ <b>RUNNING</b>" if active else "ðŸ”´ <b>STOPPED</b>"
    nums = cfg["numbers"]
    SHOW_MAX = 10
    num_lines = ""
    for n in nums[:SHOW_MAX]:
        c_name, flag = get_country_details(n)
        num_lines += f"  â€¢ <code>{n}</code>  {flag} {c_name}\n"
    if len(nums) > SHOW_MAX:
        num_lines += f"  ... +{len(nums) - SHOW_MAX} more\n"
    return (
        f"ðŸŽ­ðŸ”¥ <b>DEMO OTP PANEL</b> ðŸ”¥ðŸŽ­\n"
        f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
        f"ðŸ“¡ <b>Status   â–¸â–¸</b>  {status}\n"
        f"ðŸ“± <b>Numbers ({len(nums)}):</b>\n{num_lines}"
        f"ðŸ”¢ <b>Digits   â–¸â–¸</b>  {cfg['digits']}\n"
        f"ðŸ’¬ <b>Service  â–¸â–¸</b>  {cfg['service']}\n"
        f"â±ï¸ <b>Interval â–¸â–¸</b>  every {cfg['interval']}s\n\n"
        f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡"
    )


def demo_menu_markup():
    m = types.ReplyKeyboardMarkup(resize_keyboard=True)
    with _demo_lock:
        active = _demo_active
    if active:
        m.add("â¹ï¸ ð——ð—˜ð— ð—¢ ð—¦ð—§ð—¢ð—£")
    else:
        m.add("â–¶ï¸ ð——ð—˜ð— ð—¢ ð—¦ð—§ð—”ð—¥ð—§")
    m.add("âš™ï¸ ð——ð—˜ð— ð—¢ ð—–ð—¢ð—¡ð—™ð—œð—š")
    m.add("ðŸ”™ ð—”ð——ð— ð—œð—¡ ð—£ð—”ð—¡ð—˜ð—Ÿ")
    return m


# â”€â”€ Menus â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def main_menu(user_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    markup.add(types.KeyboardButton("â˜Žï¸ ð—¡ð—¨ð— ð—•ð—”ð—¥ â˜Žï¸"))
    markup.add(types.KeyboardButton("ðŸ“Š ð—¦ð—§ð—¢ð—–ð—ž"), types.KeyboardButton("ðŸ“ž ð—¦ð—”ð—£ð—¢ð—¥ð—§"))
    if user_id in ADMIN_IDS:
        markup.add(types.KeyboardButton("âš™ï¸ ð—”ð——ð— ð—œð—¡ ð—£ð—”ð—¡ð—˜ð—Ÿ âš™ï¸"))
    return markup


def save_services():
    save_json(SERVICES_FILE, _services)


def _get_svc_map():
    return {s["label"]: s["key"] for s in _services}


SERVICE_BUTTON_MAP = {}


def show_services(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btns = [types.KeyboardButton(s["label"]) for s in _services]
    for i in range(0, len(btns), 2):
        markup.add(*btns[i:i + 2])
    markup.add(types.KeyboardButton("ðŸ”™ Main Menu"))
    bot.send_message(
        message.chat.id,
        "ðŸ›  <b>Select Service:</b>",
        reply_markup=markup,
        parse_mode="HTML",
    )


def show_countries(chat_id, svc):
    markup = types.InlineKeyboardMarkup(row_width=2)
    btns = []
    if svc in stock:
        for cnt, nums in stock[svc].items():
            if nums:
                _, flag = get_country_details(nums[0])
                btns.append(
                    types.InlineKeyboardButton(
                        f"{flag} {cnt}", callback_data=f"n:{svc}:{cnt}"
                    )
                )
    if btns:
        markup.add(*btns)
    markup.add(
        types.InlineKeyboardButton("â¬…ï¸ ð—•ð—®ð—°ð—¸", callback_data="back_to_services")
    )
    bot.send_message(
        chat_id,
        f"ðŸ”¥ <b>{svc.upper()} â€” COUNTRY SELECT</b> ðŸ”¥",
        reply_markup=markup,
        parse_mode="HTML",
    )


# â”€â”€ Handlers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


@bot.message_handler(commands=["start"])
def start_cmd(message):
    u = message.from_user
    register_user(
        message.chat.id,
        first_name=u.first_name or "",
        last_name=u.last_name or "",
        username=u.username or "",
    )
    uname = f"@{u.username}" if u.username else (u.first_name or "User")
    uid_str = u.id
    is_admin = u.id in ADMIN_IDS
    # Per-admin group/channel links â€” admins see THEIR OWN links
    if is_admin:
        grp_link = get_admin_group_link(u.id) or ""
        main_ch  = get_main_channel_url(u.id) or ""
    else:
        grp_link = get_otp_group_link() or ""
        main_ch  = get_channel2() or ""

    markup = types.InlineKeyboardMarkup()
    if grp_link:
        markup.add(types.InlineKeyboardButton("ðŸ”¥ ð—¢ð—§ð—£ ð—šð—¿ð˜‚ð—½ ð—ð—¢ð—œð—¡ ðŸ”¥", url=grp_link))
    if main_ch:
        markup.add(types.InlineKeyboardButton("ðŸ“¢ ð— ð—®ð—¶ð—» ð—–ð—µð—®ð—»ð—»ð—²ð—¹ ð—ð—¢ð—œð—¡", url=main_ch))
    markup.add(types.InlineKeyboardButton("âœ… ð—©ð—˜ð—¥ð—œð—™ð—¬ ð—žð—¢ð—¥ð—¢ âœ…", callback_data="v"))
    bot.send_message(
        message.chat.id,
        f"ðŸ”¥ <b>ð—¥ð—”ð—•ð—•ð—œ ð—¢ð—§ð—£ ð—•ð—¢ð—§-ð—² ð—¦ð—”ð—šð—¢ð—§ð—¢ð— !</b> ðŸ”¥\n\n"
        f"â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—\n"
        f"   ðŸ§¾ <b>USER DASHBOARD</b>\n"
        f"â• â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•£\n"
        f"  ðŸ‘¤ <b>User:</b> {uname}\n"
        f"  ðŸ†” <b>ID:</b> <code>{uid_str}</code>\n"
        f"  ðŸ“Š <b>Status:</b> ðŸ’Ž Premium\n"
        f"  ðŸš€ <b>Workers:</b> 0\n"
        f"â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•\n\n"
        f"â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—\n"
        f" ð—¡ð—¶ð—°ð—µð—²ð—¿ ð—°ð—µð—®ð—»ð—»ð—²ð—¹ð—² <b>ð—ð—¢ð—œð—¡</b> ð—µð—¼ð˜†ð—²\n"
        f" <b>ð—©ð—˜ð—¥ð—œð—™ð—¬</b> ð—¯ð—®ð˜ð—®ð—»ð—² ð—°ð—¹ð—¶ð—°ð—¸ ð—¸ð—¼ð—¿ð—¼!\n"
        f"â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•\n\n"
        f"ðŸ¤– <b>POWERED BY RABBI ð—¢ð—§ð—£ ð—•ð—¢ð—§</b> ðŸ”¥",
        reply_markup=markup,
        parse_mode="HTML",
    )


@bot.message_handler(commands=["test"])
def test_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    fake_otp = str(random.randint(100000, 999999))
    fake_number = "8801712345678"
    fake_cli = "880TWVERIFY001"
    fake_sms = "Your Instagram verification code is " + fake_otp
    fake_secs = 12
    send_otp_message(message.chat.id, fake_otp, fake_number, fake_secs, fake_cli, fake_sms)
    try:
        send_otp_message(get_otp_group_id(), fake_otp, fake_number, fake_secs, fake_cli, fake_sms)
        bot.send_message(message.chat.id, "âœ… Test OTP group-eà¦“ pathano hoyeche!", parse_mode="HTML")
    except Exception as e:
        bot.send_message(message.chat.id, f"âš ï¸ Group-e pathate parina: <code>{e}</code>", parse_mode="HTML")


@bot.message_handler(commands=["setbrand"])
def setbrand_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    parts = message.text.strip().split(None, 1)
    if len(parts) < 2:
        bot.send_message(message.chat.id, "âš ï¸ à¦¬à§à¦¯à¦¬à¦¹à¦¾à¦°: <code>/setbrand RABBI</code>", parse_mode="HTML")
        return
    brand = parts[1].strip()
    cfg = get_admin_config(uid)
    cfg["brand"] = brand
    save_admin_config(uid, cfg)
    bot.send_message(message.chat.id, f"âœ… à¦†à¦ªà¦¨à¦¾à¦° Brand à¦¸à§‡à¦Ÿ à¦¹à¦¯à¦¼à§‡à¦›à§‡: <b>{brand}</b>", parse_mode="HTML")


@bot.message_handler(commands=["setnumch"])
def setnumch_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    parts = message.text.strip().split(None, 1)
    if len(parts) < 2:
        bot.send_message(message.chat.id, "âš ï¸ à¦¬à§à¦¯à¦¬à¦¹à¦¾à¦°: <code>/setnumch https://t.me/channel</code>", parse_mode="HTML")
        return
    url = parts[1].strip()
    cfg = get_admin_config(uid)
    cfg["numChannel"] = url
    save_admin_config(uid, cfg)
    bot.send_message(message.chat.id, f"âœ… à¦†à¦ªà¦¨à¦¾à¦° Number Channel à¦¸à§‡à¦Ÿ: {url}", parse_mode="HTML")


@bot.message_handler(commands=["setmainch"])
def setmainch_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    parts = message.text.strip().split(None, 1)
    if len(parts) < 2:
        bot.send_message(message.chat.id, "âš ï¸ à¦¬à§à¦¯à¦¬à¦¹à¦¾à¦°: <code>/setmainch https://t.me/channel</code>", parse_mode="HTML")
        return
    url = parts[1].strip()
    cfg = get_admin_config(uid)
    cfg["mainChannel"] = url
    save_admin_config(uid, cfg)
    bot.send_message(message.chat.id, f"âœ… à¦†à¦ªà¦¨à¦¾à¦° Main Channel à¦¸à§‡à¦Ÿ: {url}", parse_mode="HTML")


@bot.message_handler(commands=["setbotlink"])
def setbotlink_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    parts = message.text.strip().split(None, 1)
    if len(parts) < 2:
        bot.send_message(message.chat.id, "âš ï¸ à¦¬à§à¦¯à¦¬à¦¹à¦¾à¦°: <code>/setbotlink https://t.me/mybot</code>", parse_mode="HTML")
        return
    url = parts[1].strip()
    cfg = get_admin_config(uid)
    cfg["botLink"] = url
    save_admin_config(uid, cfg)
    bot.send_message(message.chat.id, f"âœ… à¦†à¦ªà¦¨à¦¾à¦° Bot Link à¦¸à§‡à¦Ÿ: {url}", parse_mode="HTML")


@bot.message_handler(commands=["setgroup"])
def setgroup_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    parts = message.text.strip().split(None, 1)
    if len(parts) < 2:
        bot.send_message(
            message.chat.id,
            "âš ï¸ à¦¬à§à¦¯à¦¬à¦¹à¦¾à¦°:\n"
            "<code>/setgroup -1001234567890</code> â€” Group ID à¦¦à¦¿à¦¯à¦¼à§‡\n\n"
            "ðŸ“Œ Group ID à¦ªà§‡à¦¤à§‡: à¦¬à¦Ÿà¦•à§‡ group-à¦ add à¦•à¦°à§à¦¨, à¦¤à¦¾à¦°à¦ªà¦° group-à¦ /getid à¦¦à¦¿à¦¨\n"
            "à¦…à¦¥à¦¬à¦¾ group ID + link à¦¦à§à¦Ÿà§‹ à¦¦à¦¿à¦¨:\n"
            "<code>/setgroup -1001234567890 https://t.me/mygroup</code>",
            parse_mode="HTML"
        )
        return
    args = parts[1].strip().split()
    group_id_str = args[0]
    group_link = args[1] if len(args) > 1 else ""
    try:
        group_id = int(group_id_str)
    except ValueError:
        bot.send_message(message.chat.id, "âŒ Valid group ID à¦¦à¦¾à¦“ (à¦¯à§‡à¦®à¦¨: -1001234567890)", parse_mode="HTML")
        return
    cfg = get_admin_config(uid)
    cfg["group_id"] = group_id
    if group_link:
        cfg["group_link"] = group_link
    save_admin_config(uid, cfg)
    bot.send_message(
        message.chat.id,
        f"âœ… à¦†à¦ªà¦¨à¦¾à¦° OTP Group à¦¸à§‡à¦Ÿ à¦¹à¦¯à¦¼à§‡à¦›à§‡!\n"
        f"ðŸ†” Group ID: <code>{group_id}</code>\n"
        f"ðŸ”— Link: {group_link or '(à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡)'}",
        parse_mode="HTML"
    )


@bot.message_handler(commands=["getid"])
def getid_cmd(message):
    cid = message.chat.id
    ctype = message.chat.type
    title = getattr(message.chat, "title", "") or ""
    bot.send_message(
        message.chat.id,
        f"ðŸ†” <b>Chat ID:</b> <code>{cid}</code>\n"
        f"ðŸ“‹ Type: {ctype}\n"
        f"ðŸ“Œ Title: {title or 'â€”'}",
        parse_mode="HTML"
    )


@bot.message_handler(commands=["myconfig"])
def myconfig_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    cfg = get_admin_config(uid)
    gid = cfg.get("group_id") or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    glink = cfg.get("group_link") or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    brand = cfg.get("brand") or "RABBI"
    num_ch = cfg.get("numChannel") or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    main_ch = cfg.get("mainChannel") or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    bot_lnk = cfg.get("botLink") or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    bot.send_message(
        message.chat.id,
        f"âš™ï¸ <b>à¦†à¦ªà¦¨à¦¾à¦° Config (Admin {uid})</b>\n\n"
        f"ðŸ· Brand: <b>{brand}</b>\n"
        f"ðŸ†” OTP Group ID: <code>{gid}</code>\n"
        f"ðŸ”— Group Link: {glink}\n"
        f"ðŸ“² Number Ch: {num_ch}\n"
        f"ðŸ“¢ Main Ch: {main_ch}\n"
        f"ðŸ¤– Bot Link: {bot_lnk}\n\n"
        f"<b>à¦¸à§‡à¦Ÿ à¦•à¦°à¦¤à§‡:</b>\n"
        f"/setbrand &lt;à¦¨à¦¾à¦®&gt;\n"
        f"/setgroup &lt;group_id&gt; [link]\n"
        f"/setnumch &lt;link&gt;\n"
        f"/setmainch &lt;link&gt;\n"
        f"/setbotlink &lt;link&gt;",
        parse_mode="HTML"
    )


@bot.message_handler(commands=["panels"])
def panels_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    with _stats_lock:
        stats = {k: dict(v) for k, v in _panel_stats.items()}
    lines = ""
    for pid in ["p1", "p2", "p3", "p4", "p5", "p6"]:
        s = stats.get(pid, {})
        if s.get("last"):
            ago = int(time.time() - s["last"])
            last_str = f"{ago}s ago"
        else:
            last_str = "never"
        err_str = f"  âš ï¸ {s['errors']} err" if s.get("errors") else ""
        lines += (
            f"{s.get('status', 'â³')} <b>{s.get('name', '?')}</b>\n"
            f"   ðŸŒ <code>{s.get('host', '?')}</code>\n"
            f"   ðŸ“Š {s.get('count', 0)} records  â€¢  ðŸ• {last_str}{err_str}\n\n"
        )
    with _demo_lock:
        demo_on = _demo_active
    demo_str = "ðŸŸ¢ Running" if demo_on else "ðŸ”´ Stopped"
    bot.send_message(
        message.chat.id,
        f"ðŸ“¡ <b>PANEL STATUS</b>\n"
        f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
        f"{lines}"
        f"ðŸŽ­ <b>Demo OTP:</b>  {demo_str}\n\n"
        f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n"
        f"ðŸ”„ <i>Updates every {POLL_INTERVAL}s</i>",
        parse_mode="HTML",
    )
    dp_copy = list(_dynamic_panels)
    if dp_copy:
        dp_lines = ""
        for p in dp_copy:
            pid = p["id"]
            with _stats_lock:
                s = _panel_stats.get(pid, {})
            st = s.get("status", "â³")
            cnt = s.get("count", 0)
            err = s.get("errors", 0)
            t = s.get("last")
            last_str = f"{int(time.time() - t)}s ago" if t else "never"
            err_str = f"  âš ï¸ {err} err" if err else ""
            dp_lines += (
                f"{st} <b>{p.get('username', '?')}</b> <code>[{pid}]</code>\n"
                f"   ðŸŒ <code>{p.get('host', '?')}</code>\n"
                f"   ðŸ“Š {cnt} records  â€¢  ðŸ• {last_str}{err_str}\n\n"
            )
        bot.send_message(
            message.chat.id,
            f"ðŸ“¡ <b>DYNAMIC PANELS</b>\n"
            f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
            f"{dp_lines}"
            f"ðŸ’¡ <i>/addpanel diye naya panel add koro</i>",
            parse_mode="HTML",
        )


@bot.message_handler(commands=["broadcast"])
def broadcast_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    msg = bot.send_message(
        message.chat.id,
        "âœï¸ <b>Broadcast content à¦ªà¦¾à¦ à¦¾à¦“:</b> \n\n"
        "ðŸ“ Text\nðŸ–¼ï¸ Photo\nðŸŽ¥ Video\nðŸŽ­ Sticker\n"
        "ðŸŽžï¸ GIF / Animation\nðŸŽµ Audio / Music\nðŸŽ¤ Voice message\nðŸ“Ž Document / APK / ZIP / PDF\n\n"
        "<i>Caption support ache â€” sob kichute!</i>",
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, do_broadcast)


def _clr_service_markup():
    markup = types.InlineKeyboardMarkup(row_width=2)
    services = [
        ("facebook", "ðŸ’¬"),
        ("instagram", "ðŸ“¸"),
        ("whatsapp", "ðŸ“±"),
        ("telegram", "âœˆï¸"),
        ("binance", "ðŸª™"),
        ("pc clone", "ðŸ’»"),
    ]
    for svc, icon in services:
        total = sum(len(v) for v in stock.get(svc, {}).values())
        markup.add(
            types.InlineKeyboardButton(
                f"{icon} {svc.upper()} ({total})", callback_data=f"clr_s:{svc}"
            )
        )
    markup.add(types.InlineKeyboardButton(" Clear ALL Stock", callback_data="clr_all"))
    return markup


@bot.message_handler(commands=["addpanel"])
def addpanel_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    _addpanel_state[message.from_user.id] = {"step": "url", "data": {}}
    msg = bot.send_message(
        message.chat.id,
        "ðŸ”§ðŸ”¥ <b>ADD NEW PANEL</b> ðŸ”¥ðŸ”§\n\n"
        "ðŸ“¡ <b>Step 1/3:</b> Panel URL pathao\n"
        "<i>Example: http://1.2.3.4/ints/agent/SMSCDRStats</i>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _ap_get_url)


def _ap_get_url(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    url = (message.text or "").strip()
    base_url = None
    panel_type = "smscdr"
    m_ints = re.match(r"(https?://[^/]+/(?:ints|sms))(?:/|$)", url)
    m_agent = re.match(r"(https?://[^/]+)/agent/(SMSCDRStats|SMSRanges)", url, re.IGNORECASE)
    m_domain = re.match(r"(https?://[^/?#]+)/?$", url)
    if m_ints:
        base_url = m_ints.group(1)
        panel_type = "smscdr"
    elif m_agent:
        base_url = m_agent.group(1)
        panel_type = "smsranges" if m_agent.group(2).lower() == "smsranges" else "smscdr"
    elif m_domain:
        base_url = m_domain.group(1)
        panel_type = "smscdr"
    if not base_url:
        msg = bot.send_message(
            message.chat.id,
            "âŒ Valid URL dao:\n"
            "â€¢ <code>http://1.2.3.4/ints/agent/SMSCDRStats</code>\n"
            "â€¢ <code>https://truesms.net/agent/SMSCDRStats</code>\n"
            "â€¢ <code>https://truesms.net/agent/SMSRanges</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _ap_get_url)
        return
    host_m = re.search(r"//([^/]+)", base_url)
    uid = message.from_user.id
    _addpanel_state[uid]["data"]["base_url"] = base_url
    _addpanel_state[uid]["data"]["host"] = host_m.group(1) if host_m else base_url
    _addpanel_state[uid]["data"]["panel_type"] = panel_type
    type_label = "SMSRanges" if panel_type == "smsranges" else "SMSCDRStats"
    msg = bot.send_message(
        message.chat.id,
        f"âœ… URL: <code>{base_url}</code>\n"
        f"ðŸ“Š Type: <b>{type_label}</b>\n\n"
        f"ðŸ‘¤ <b>Step 2/3:</b> Username pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _ap_get_user)


def _ap_get_user(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    username = (message.text or "").strip()
    if not username:
        msg = bot.send_message(message.chat.id, "âŒ Username dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _ap_get_user)
        return
    _addpanel_state[message.from_user.id]["data"]["username"] = username
    msg = bot.send_message(
        message.chat.id,
        f"âœ… Username: <code>{username}</code>\n\nðŸ”‘ <b>Step 3/3:</b> Password pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _ap_get_pass)


def _ap_get_pass(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    if _is_back(message.text):
        _addpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    password = (message.text or "").strip()
    if not password:
        msg = bot.send_message(message.chat.id, "âŒ Password dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _ap_get_pass)
        return
    data = _addpanel_state.get(uid, {}).get("data", {})
    data["password"] = password
    wait_msg = bot.send_message(
        message.chat.id,
        "â³ðŸ”¥ <b>Connection test korchi...</b>\n<i>Ektu wait koro!</i>",
        parse_mode="HTML",
    )
    panel_id = f"d{int(time.time()) % 100000}"
    panel = {
        "id": panel_id,
        "host": data.get("host", ""),
        "base_url": data.get("base_url", ""),
        "username": data.get("username", ""),
        "password": password,
        "panel_type": data.get("panel_type", "smscdr"),
        "owner_uid": uid,
    }
    sess, token = _ints_login(panel)
    try:
        bot.delete_message(message.chat.id, wait_msg.message_id)
    except Exception:
        pass
    if not sess:
        bot.send_message(
            message.chat.id,
            "âŒ <b>Connection FAILED!</b> âŒ\n\n"
            "âš ï¸ URL, username ba password check koro.\n"
            "Aro try korte /addpanel pathao.",
            parse_mode="HTML",
        )
        _addpanel_state.pop(uid, None)
        return
    _dynamic_sessions[panel_id] = {"session": sess, "token": token}
    _dynamic_panels.append(panel)
    save_dynamic_panels()
    _start_dynamic_panel(panel)
    bot.send_message(
        message.chat.id,
        f"âœ…ðŸ”¥ <b>PANEL ADDED & STARTED!</b> ðŸ”¥âœ…\n"
        f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
        f"ðŸ†” <b>ID     â–¸â–¸</b> <code>{panel_id}</code>\n"
        f"ðŸŒ <b>Host   â–¸â–¸</b> <code>{data['host']}</code>\n"
        f"ðŸ‘¤ <b>User   â–¸â–¸</b> <code>{data['username']}</code>\n\n"
        f"ðŸ“¡ Monitor thread started! /panels diye check koro.",
        parse_mode="HTML",
    )
    _addpanel_state.pop(uid, None)


def _svc_get_label(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addservice_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    label = (message.text or "").strip()
    if not label:
        msg = bot.send_message(message.chat.id, "âŒ Label dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _svc_get_label)
        return
    _addservice_state[message.from_user.id]["label"] = label
    msg = bot.send_message(
        message.chat.id,
        f"âœ… Label: <b>{label}</b>\n\n"
        "ðŸ”‘ <b>Step 2/2:</b> Internal key dao (lowercase, no space)\n"
        "<i>Example: telegram, binance, tiktok</i>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _svc_get_key)


def _svc_get_key(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addservice_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    key = (message.text or "").strip().lower()
    if not key:
        msg = bot.send_message(message.chat.id, "âŒ Key dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _svc_get_key)
        return
    label = _addservice_state.get(message.from_user.id, {}).get("label", "")
    existing_keys = [s["key"] for s in _services]
    if key in existing_keys:
        msg = bot.send_message(
            message.chat.id,
            f"âŒ Key <code>{key}</code> already ache! Onnyo key dao:",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _svc_get_key)
        return
    _services.append({"label": label, "key": key})
    save_services()
    _addservice_state.pop(message.from_user.id, None)
    _go_admin_panel(
        message,
        f"âœ…ðŸ”¥ <b>Service Added!</b>\n\n"
        f"ðŸ·ï¸ Label: <b>{label}</b>\n"
        f"ðŸ”‘ Key: <code>{key}</code>\n\n"
        f"<i>Service menu-te dekha jabe!</i>",
    )


@bot.message_handler(commands=["listpanels"])
def listpanels_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if not _dynamic_panels:
        bot.send_message(
            message.chat.id,
            "ðŸ“‹ No dynamic panels yet.\nðŸ’¡ /addpanel diye add koro.",
            parse_mode="HTML",
        )
        return
    uid = message.from_user.id
    is_super = (uid == SUPER_ADMIN_ID)
    visible = [p for p in _dynamic_panels if is_super or p.get("owner_uid") == uid]
    if not visible:
        bot.send_message(message.chat.id, "ðŸ“‹ à¦†à¦ªà¦¨à¦¾à¦° à¦•à§‹à¦¨à§‹ panel à¦¨à§‡à¦‡à¥¤\nðŸ’¡ /addpanel à¦¦à¦¿à¦¯à¦¼à§‡ add à¦•à¦°à§à¦¨à¥¤", parse_mode="HTML")
        return
    lines = "ðŸ“‹ðŸ”¥ <b>YOUR PANELS</b> ðŸ”¥ðŸ“‹\nâš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
    for p in visible:
        pid = p["id"]
        with _stats_lock:
            s = _panel_stats.get(pid, {})
        st = s.get("status", "â³")
        owner_id = p.get("owner_uid", "?")
        owner_tag = f" (admin {owner_id})" if is_super else ""
        lines += (
            f"{st} ðŸ†” <code>{pid}</code>{owner_tag}\n"
            f"   ðŸŒ <code>{p.get('host', '?')}</code>\n"
            f"   ðŸ‘¤ {p.get('username', '?')}\n\n"
        )
    lines += "ðŸ—‘ï¸ Remove: <code>/removepanel [ID]</code>"
    bot.send_message(message.chat.id, lines, parse_mode="HTML")


@bot.message_handler(commands=["removepanel"])
def removepanel_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        bot.send_message(
            message.chat.id,
            "âŒ Panel ID dao:\n<code>/removepanel d12345</code>\n\n"
            "ðŸ’¡ /listpanels diye ID dekho.",
            parse_mode="HTML",
        )
        return
    pid = args[1].strip()
    before = len(_dynamic_panels)
    _dynamic_panels[:] = [p for p in _dynamic_panels if p["id"] != pid]
    if len(_dynamic_panels) < before:
        save_dynamic_panels()
        with _stats_lock:
            _panel_stats.pop(pid, None)
        _dynamic_sessions.pop(pid, None)
        _dynamic_locks.pop(pid, None)
        bot.send_message(
            message.chat.id,
            f"âœ…ðŸ”¥ Panel <code>{pid}</code> removed!\n"
            f"<i>Monitor thread will stop naturally.</i>",
            parse_mode="HTML",
        )
    else:
        bot.send_message(
            message.chat.id,
            f"âŒ Panel <code>{pid}</code> not found.\n"
            f"ðŸ’¡ /listpanels diye ID check koro.",
            parse_mode="HTML",
        )


@bot.callback_query_handler(func=lambda call: True)
def callback_handler(call):
    global stock
    try:
        data = call.data

        if data == "v":
            uid = call.from_user.id

            grp_id = get_otp_group_id()
            grp_link = get_otp_group_link()
            ch2_link = get_channel2()
            ch2_ref = _extract_username(ch2_link)

            not_joined = []

            grp_ok = _check_member(grp_id, uid) if grp_id else None
            if grp_ok is False:
                not_joined.append(("ðŸ”¥ OTP Group", grp_link))

            ch2_ok = _check_member(ch2_ref, uid) if ch2_ref else None
            if ch2_ok is False:
                not_joined.append(("ðŸ“¢ Main Channel", ch2_link))

            if not_joined:
                bot.answer_callback_query(call.id, "âŒ Sob jagay join hao nai!", show_alert=False)
                lines = "âŒ <b>Verify hote parcho na!</b>\n\n"
                lines += "â›” Tumi ekhono nicher jagay join hao nai:\n\n"
                for name, _ in not_joined:
                    lines += f"  ðŸš« <b>{name}</b>\n"
                lines += "\nðŸ‘‡ Join kore <b>Verify Koro</b> te click koro:"
                err_markup = types.InlineKeyboardMarkup(row_width=1)
                for name, lnk in not_joined:
                    err_markup.add(types.InlineKeyboardButton(
                        f"ðŸ‘‰ {name}-e JOIN KORO", url=lnk
                    ))
                err_markup.add(types.InlineKeyboardButton(
                    "ðŸ”„ Verify Koro", callback_data="v"
                ))
                try:
                    bot.edit_message_text(
                        lines,
                        call.message.chat.id,
                        call.message.message_id,
                        reply_markup=err_markup,
                        parse_mode="HTML",
                    )
                except Exception:
                    bot.send_message(
                        call.message.chat.id,
                        lines,
                        reply_markup=err_markup,
                        parse_mode="HTML",
                    )
            else:
                bot.delete_message(call.message.chat.id, call.message.message_id)
                vname = call.from_user.first_name or call.from_user.username or "User"
                bot.send_message(
                    call.message.chat.id,
                    f"ðŸ”¥ <b>VERIFICATION COMPLETE!</b> ðŸ”¥\n\n"
                    f"â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—\n"
                    f"   âœ… <b>ACCESS GRANTED</b>\n"
                    f"â• â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•£\n"
                    f"  ðŸ‘‹ <b>Welcome, {vname}!</b>\n"
                    f"  ðŸ†” <b>ID:</b> <code>{uid}</code>\n"
                    f"  ðŸ“Š <b>Status:</b> ðŸ’Ž Premium\n"
                    f"â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•\n\n"
                    f"âš¡ <b>ð—˜ð—¸ð—¸ð—µð—¼ð—» ð—»ð˜‚ð—ºð—¯ð—®ð—¿ ð—»ð—¶ð˜ð—² ð—½ð—®ð—¿ð—¯ð—²!</b> âš¡",
                    reply_markup=main_menu(call.from_user.id),
                    parse_mode="HTML",
                )

        elif data == "back_to_services":
            show_services(call.message)

        elif data.startswith("s:"):
            svc = data.split(":")[1]
            markup = types.InlineKeyboardMarkup(row_width=2)
            btns = []
            if svc in stock:
                for cnt, nums in stock[svc].items():
                    if nums:
                        _, flag = get_country_details(nums[0])
                        btns.append(
                            types.InlineKeyboardButton(
                                f" {flag} {cnt}", callback_data=f"n:{svc}:{cnt}"
                            )
                        )
            if btns:
                markup.add(*btns)
            markup.add(
                types.InlineKeyboardButton("â¬…ï¸ ð—•ð—®ð—°ð—¸", callback_data="back_to_services")
            )
            bot.edit_message_text(
                f"ðŸ”¥ <b>{svc.upper()} â€” COUNTRY</b> ðŸ”¥",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data.startswith("n:"):
            _, svc, scnt = data.split(":")
            if scnt in stock.get(svc, {}) and stock[svc][scnt]:
                num = stock[svc][scnt].pop(0)
                save_stock()
                c_name, flag = get_country_details(num)
                register_number(call.message.chat.id, num)
                display_num = num if num.startswith("+") else "+" + num
                init_kb = types.InlineKeyboardMarkup(row_width=2)
                init_kb.add(
                    types.InlineKeyboardButton("ðŸ”„ New Number", callback_data=f"n:{svc}:{scnt}"),
                    types.InlineKeyboardButton("ðŸŒ Change Country", callback_data=f"s:{svc}"),
                )
                init_kb.add(
                    types.InlineKeyboardButton("ðŸ“¢ OTP Group", url=get_otp_group_link()),
                )
                res = (
                    f"âœ… <b>Number Assigned Successfully !</b>\n\n"
                    f"ðŸ”§ <b>Platform :</b> {svc.capitalize()}\n"
                    f"ðŸŒ <b>Country :</b> {flag} {c_name}\n\n"
                    f"ðŸ“ž <b>Number :</b> <code>{display_num}</code>\n\n"
                    f"â± <b>Auto code fetch :</b> 10:00s"
                )
                bot.edit_message_text(
                    res,
                    call.message.chat.id,
                    call.message.message_id,
                    reply_markup=init_kb,
                    parse_mode="HTML",
                )
                _start_countdown(
                    call.message.chat.id,
                    call.message.message_id,
                    svc, flag, c_name, display_num, scnt,
                )
            else:
                bot.answer_callback_query(call.id, " STOCK SHESH! ", show_alert=True)

        elif data == "clr_menu":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.edit_message_text(
                "ðŸ—‘ï¸ðŸ”¥ <b>STOCK CLEAR PANEL</b> ðŸ”¥ðŸ—‘ï¸\n\n"
                " <b>Kon service-er stock clear korbe?</b>\n"
                "â¬‡ï¸ Service choose koro:",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=_clr_service_markup(),
                parse_mode="HTML",
            )

        elif data.startswith("clr_s:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            svc = data[6:]
            markup = types.InlineKeyboardMarkup(row_width=1)
            svc_stock = stock.get(svc, {})
            has_any = False
            for cnt, nums in svc_stock.items():
                if nums:
                    has_any = True
                    _, flag = get_country_details(nums[0])
                    cb = f"clr_c:{svc}:{cnt}"
                    if len(cb.encode()) <= 64:
                        markup.add(
                            types.InlineKeyboardButton(
                                f"ðŸ—‘ï¸ {flag} {cnt}  ({len(nums)} à¦Ÿà¦¿)", callback_data=cb
                            )
                        )
            if not has_any:
                markup.add(
                    types.InlineKeyboardButton("âš ï¸ Stock nai!", callback_data="clr_menu")
                )
            markup.add(types.InlineKeyboardButton("â¬…ï¸ Back", callback_data="clr_menu"))
            bot.edit_message_text(
                f"ðŸ”¥ <b>{svc.upper()} â€” Kon desh clear korbe?</b> ðŸ”¥\n\n"
                f"â¬‡ï¸ Country choose koro:",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data.startswith("clr_c:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            _, svc, cnt = data.split(":", 2)
            count = len(stock.get(svc, {}).get(cnt, []))
            _, flag = get_country_details(stock[svc][cnt][0]) if count else ("", "ðŸŒ")
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton(
                    "âœ… Haa, Delete Koro", callback_data=f"clr_y:{svc}:{cnt}"
                ),
                types.InlineKeyboardButton("âŒ Cancel", callback_data=f"clr_s:{svc}"),
            )
            bot.edit_message_text(
                f"âš ï¸ <b>CONFIRM DELETE</b> âš ï¸\n\n"
                f"ðŸ’¬ <b>Service â–¸â–¸</b>  {svc.upper()}\n"
                f"ðŸŒ <b>Country â–¸â–¸</b>  {flag} {cnt}\n"
                f"ðŸ“± <b>Numbers â–¸â–¸</b>  {count} à¦Ÿà¦¿\n\n"
                f" Sure? Ei {count} à¦Ÿà¦¿ number delete hoye jabe!",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data.startswith("clr_y:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            _, svc, cnt = data.split(":", 2)
            removed = len(stock.get(svc, {}).get(cnt, []))
            if svc in stock and cnt in stock[svc]:
                del stock[svc][cnt]
                save_stock()
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton("ðŸ—‘ï¸ Aro Clear", callback_data=f"clr_s:{svc}"),
                types.InlineKeyboardButton("ðŸ”™ Services", callback_data="clr_menu"),
            )
            bot.edit_message_text(
                f"âœ…ðŸ”¥ <b>DELETE COMPLETE!</b> ðŸ”¥âœ…\n\n"
                f"ðŸ’¬ <b>Service â–¸â–¸</b>  {svc.upper()}\n"
                f"ðŸŒ <b>Country â–¸â–¸</b>  {cnt}\n"
                f"ðŸ“± <b>Deleted  â–¸â–¸</b>  {removed} à¦Ÿà¦¿ number\n\n"
                f"âš¡ <i>Stock update hoyeche!</i>",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data == "clr_all":
            if call.from_user.id not in ADMIN_IDS:
                return
            total = sum(
                len(nums) for svc_d in stock.values() for nums in svc_d.values()
            )
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton(
                    " Haa, SOB Clear", callback_data="clr_allok"
                ),
                types.InlineKeyboardButton("âŒ Cancel", callback_data="clr_menu"),
            )
            bot.edit_message_text(
                f"â˜ ï¸âš ï¸ <b>CLEAR ALL CONFIRM</b> âš ï¸â˜ ï¸\n\n"
                f" Total <b>{total} à¦Ÿà¦¿</b> number delete hobe!\n"
                f"âš¡ Sob service-er sob country mochhe jabe!\n\n"
                f"ðŸ”¥ Sure? Eta undo kora jabe na!",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data == "clr_allok":
            if call.from_user.id not in ADMIN_IDS:
                return
            stock = {
                "whatsapp": {},
                "facebook": {},
                "telegram": {},
                "instagram": {},
                "pc clone": {},
                "binance": {},
            }
            save_stock()
            bot.edit_message_text(
                "ðŸ”¥ <b>SOB STOCK CLEAR HOYECHE!</b> ðŸ”¥\n <i>Ekhon naya number add koro!</i> ",
                call.message.chat.id,
                call.message.message_id,
                parse_mode="HTML",
            )

        elif data.startswith("rmpanel:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            pid = data.split(":", 1)[1]
            before = len(_dynamic_panels)
            _dynamic_panels[:] = [p for p in _dynamic_panels if p["id"] != pid]
            if len(_dynamic_panels) < before:
                save_dynamic_panels()
                with _stats_lock:
                    _panel_stats.pop(pid, None)
                _dynamic_sessions.pop(pid, None)
                _dynamic_locks.pop(pid, None)
                bot.edit_message_text(
                    f"âœ…ðŸ”¥ <b>Panel <code>{pid}</code> removed!</b>\n"
                    f"<i>Monitor thread will stop naturally.</i>",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            else:
                bot.answer_callback_query(call.id, "âŒ Panel pawa jaini!", show_alert=True)

        elif data.startswith("rmsvc:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            key = data.split(":", 1)[1]
            before = len(_services)
            _services[:] = [s for s in _services if s["key"] != key]
            if len(_services) < before:
                save_services()
                bot.edit_message_text(
                    f"âœ…ðŸ”¥ <b>Service <code>{key}</code> removed!</b>\n"
                    f"<i>Service menu theke hatiye dewa hoyeche.</i>",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            else:
                bot.answer_callback_query(call.id, "âŒ Service pawa jaini!", show_alert=True)

        elif data.startswith("rmadmin:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            target = int(data.split(":")[1])
            if remove_admin(target):
                name = user_names.get(str(target), {}).get("first_name", "") or str(target)
                bot.answer_callback_query(call.id, f"âœ… {name} removed!", show_alert=False)
                try:
                    bot.edit_message_text(
                        f"âœ… <b>ADMIN REMOVED!</b>\n\n"
                        f"ðŸ—‘ï¸ <b>Removed:</b> {name} [<code>{target}</code>]\n\n"
                        f"<i>Ekhon theke ei user admin access harabe.</i>",
                        call.message.chat.id,
                        call.message.message_id,
                        parse_mode="HTML",
                    )
                except Exception:
                    pass
            else:
                bot.answer_callback_query(call.id, "âŒ Remove kora gelo na (Super Admin)!", show_alert=True)

        elif data == "grp_info":
            if call.from_user.id not in ADMIN_IDS:
                return
            _show_settings_inline(call)

        elif data == "set_autodel":
            if call.from_user.id not in ADMIN_IDS:
                return
            cur = _group_settings.get("auto_delete", True)
            _group_settings["auto_delete"] = not cur
            save_group_settings()
            bot.answer_callback_query(
                call.id,
                "âœ… Auto Delete: " + ("ðŸŸ¢ ON" if not cur else "ðŸ”´ OFF"),
                show_alert=False,
            )
            _show_settings_inline(call)

        elif data == "set_channel2":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ“¢ <b>Notun Join Channel link dao:</b>\n\n"
                "<i>Example: https://t.me/aR_OTP_rcv</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _sett_get_channel2)

        elif data == "set_botlink":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ¤– <b>Notun Bot link dao:</b>\n\n"
                "<i>Example: https://t.me/ar_otp_bot</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _sett_get_botlink)

        elif data == "grp_setlink":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ”— <b>Notun OTP Group Link dao:</b>\n\n"
                "<i>Example: https://t.me/aR_OTP_rcv</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _grp_get_link)

        elif data == "grp_setid":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ†” <b>Notun OTP Group Chat ID dao:</b>\n\n"
                "<i>Example: -1001234567890</i>\n"
                "âš ï¸ Negative number dite hobe (group ID always negative)",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _grp_get_id)

        elif data == "grp_remove":
            if call.from_user.id not in ADMIN_IDS:
                return
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton("âœ… Haa, Remove", callback_data="grp_removeok"),
                types.InlineKeyboardButton("âŒ Cancel", callback_data="grp_info"),
            )
            bot.answer_callback_query(call.id)
            bot.edit_message_text(
                "âš ï¸ <b>CONFIRM GROUP REMOVE</b> âš ï¸\n\n"
                "OTP Group setting reset hobe!\n"
                "Group-e aro OTP pathano bondho hobe.\n\n"
                "Sure?",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data == "grp_removeok":
            if call.from_user.id not in ADMIN_IDS:
                return
            _group_settings["otp_group_id"] = None
            _group_settings["otp_group_link"] = ""
            save_group_settings()
            bot.answer_callback_query(call.id, "âœ… Group removed!")
            _show_settings_inline(call)

        # â”€â”€ Per-admin Settings inline buttons â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        elif data.startswith("my_setgroup:"):
            admin_uid = int(data.split(":")[1])
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ†” <b>OTP Group Chat ID dao:</b>\n\n"
                "1ï¸âƒ£ Bot-ke group-e add koro\n"
                "2ï¸âƒ£ Group-e /getid send koro â€” ID pabe\n"
                "3ï¸âƒ£ Sei ID ekhane paste koro:\n"
                "<i>Example: -1001234567890</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            def _step_setgroup(m, _uid=admin_uid):
                if _is_back(m.text):
                    _show_settings(m)
                    return
                raw = (m.text or "").strip()
                try:
                    gid = int(raw)
                except ValueError:
                    msg2 = bot.send_message(m.chat.id, "âŒ Number daw (e.g. -1001234567890):", reply_markup=_back_admin_kb(), parse_mode="HTML")
                    bot.register_next_step_handler(msg2, _step_setgroup)
                    return
                cfg = get_admin_config(_uid)
                cfg["group_id"] = gid
                save_admin_config(_uid, cfg)
                _go_admin_panel(m, f"âœ… <b>OTP Group ID set!</b>\n\nðŸ†” <code>{gid}</code>")
            bot.register_next_step_handler(msg, _step_setgroup)

        elif data.startswith("my_setbrand:"):
            admin_uid = int(data.split(":")[1])
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ‘‘ <b>Brand name dao:</b>\n\n<i>Example: RABBI</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            def _step_setbrand(m, _uid=admin_uid):
                if _is_back(m.text):
                    _show_settings(m)
                    return
                brand = (m.text or "").strip()
                if not brand:
                    msg2 = bot.send_message(m.chat.id, "âŒ Brand name dao:", reply_markup=_back_admin_kb(), parse_mode="HTML")
                    bot.register_next_step_handler(msg2, _step_setbrand)
                    return
                cfg = get_admin_config(_uid)
                cfg["brand"] = brand
                save_admin_config(_uid, cfg)
                _go_admin_panel(m, f"âœ… <b>Brand set:</b> <b>{brand}</b>")
            bot.register_next_step_handler(msg, _step_setbrand)

        elif data.startswith("my_setnumch:"):
            admin_uid = int(data.split(":")[1])
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ“² <b>Number Channel link dao:</b>\n\n<i>Example: https://t.me/gsjggj98</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            def _step_setnumch(m, _uid=admin_uid):
                if _is_back(m.text):
                    _show_settings(m)
                    return
                link = (m.text or "").strip()
                if not link.startswith("https://"):
                    msg2 = bot.send_message(m.chat.id, "âŒ Valid https:// link dao:", reply_markup=_back_admin_kb(), parse_mode="HTML")
                    bot.register_next_step_handler(msg2, _step_setnumch)
                    return
                cfg = get_admin_config(_uid)
                cfg["numChannel"] = link
                save_admin_config(_uid, cfg)
                _go_admin_panel(m, f"âœ… <b>Number Channel set:</b>\n{link}")
            bot.register_next_step_handler(msg, _step_setnumch)

        elif data.startswith("my_setmainch:"):
            admin_uid = int(data.split(":")[1])
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ“¢ <b>Main Channel link dao:</b>\n\n<i>Example: https://t.me/facboo578</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            def _step_setmainch(m, _uid=admin_uid):
                if _is_back(m.text):
                    _show_settings(m)
                    return
                link = (m.text or "").strip()
                if not link.startswith("https://"):
                    msg2 = bot.send_message(m.chat.id, "âŒ Valid https:// link dao:", reply_markup=_back_admin_kb(), parse_mode="HTML")
                    bot.register_next_step_handler(msg2, _step_setmainch)
                    return
                cfg = get_admin_config(_uid)
                cfg["mainChannel"] = link
                save_admin_config(_uid, cfg)
                _go_admin_panel(m, f"âœ… <b>Main Channel set:</b>\n{link}")
            bot.register_next_step_handler(msg, _step_setmainch)

        elif data.startswith("my_setbotlink:"):
            admin_uid = int(data.split(":")[1])
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "ðŸ¤– <b>Bot link dao:</b>\n\n<i>Example: https://t.me/rabbi_otp_bot</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            def _step_setbotlink(m, _uid=admin_uid):
                if _is_back(m.text):
                    _show_settings(m)
                    return
                link = (m.text or "").strip()
                if not link.startswith("https://"):
                    msg2 = bot.send_message(m.chat.id, "âŒ Valid https:// link dao:", reply_markup=_back_admin_kb(), parse_mode="HTML")
                    bot.register_next_step_handler(msg2, _step_setbotlink)
                    return
                cfg = get_admin_config(_uid)
                cfg["botLink"] = link
                save_admin_config(_uid, cfg)
                _go_admin_panel(m, f"âœ… <b>Bot Link set:</b>\n{link}")
            bot.register_next_step_handler(msg, _step_setbotlink)

    except Exception as e:
        print(f"Callback Error: {e}")


# â”€â”€ Excel / CSV helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

VALID_SERVICES = [
    "facebook",
    "instagram",
    "whatsapp",
    "telegram",
    "binance",
    "pc clone",
]


def _parse_spreadsheet(data: bytes, filename: str):
    """
    Parse Excel (.xlsx / .xls) or CSV file.
    Returns:
      - (rows, mode)
        mode='two_col' â†’ rows = list of (service, number)
        mode='one_col' â†’ rows = list of number strings
    Accepts header rows with 'service'/'number' labels.
    Falls back: 2-column files = service+number, 1-column = numbers only.
    """
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    raw_rows = []

    if ext == "csv":
        text = data.decode("utf-8", errors="replace")
        reader = csv.reader(io.StringIO(text))
        for row in reader:
            cleaned = [c.strip() for c in row if c.strip()]
            if cleaned:
                raw_rows.append(cleaned)
    elif ext == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            cleaned = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cleaned:
                raw_rows.append(cleaned)
    elif ext == "xls":
        wb = xlrd.open_workbook(file_contents=data)
        ws = wb.sheet_by_index(0)
        for ri in range(ws.nrows):
            cleaned = [
                str(ws.cell_value(ri, ci)).strip()
                for ci in range(ws.ncols)
                if str(ws.cell_value(ri, ci)).strip()
            ]
            if cleaned:
                raw_rows.append(cleaned)
    else:
        return [], "unknown"

    if not raw_rows:
        return [], "empty"

    # Detect header row
    start = 0
    first = [c.lower() for c in raw_rows[0]]
    if any(h in first for h in ("service", "number", "phone", "mobile")):
        start = 1

    data_rows = raw_rows[start:]
    if not data_rows:
        return [], "empty"

    # Detect mode by column count of the majority of rows
    two_col_count = sum(1 for r in data_rows if len(r) >= 2)
    one_col_count = len(data_rows) - two_col_count

    if two_col_count > one_col_count:
        result = []
        for r in data_rows:
            if len(r) < 2:
                continue
            col0, col1 = r[0], r[1]
            # Determine which column is service and which is number
            col0_is_num = re.match(r"^\+?\d{6,15}$", re.sub(r"\s", "", col0))
            col1_is_num = re.match(r"^\+?\d{6,15}$", re.sub(r"\s", "", col1))
            if col0_is_num and not col1_is_num:
                svc = col1.lower().strip()
                num = re.sub(r"\D", "", col0)
            elif col1_is_num and not col0_is_num:
                svc = col0.lower().strip()
                num = re.sub(r"\D", "", col1)
            else:
                svc = col0.lower().strip()
                num = re.sub(r"\D", "", col1)
            if num and len(num) >= 7:
                result.append((svc, num))
        return result, "two_col"
    else:
        result = []
        for r in data_rows:
            num = re.sub(r"\D", "", r[0])
            if len(num) >= 7:
                result.append(num)
        return result, "one_col"


def _add_numbers_bulk(svc: str, numbers: list):
    """Add a list of number strings to stock[svc]. Returns (added, skipped)."""
    added, skipped = 0, 0
    svc = svc.lower().strip()
    if svc not in stock:
        return 0, len(numbers)
    for num in numbers:
        num = re.sub(r"\D", "", str(num))
        if not num:
            skipped += 1
            continue
        c_name, _ = get_country_details(num)
        if c_name == "Unknown":
            skipped += 1
            continue
        if c_name not in stock[svc]:
            stock[svc][c_name] = []
        stock[svc][c_name].append(num)
        added += 1
    if added:
        save_stock()
    return added, skipped


def _service_select_markup():
    m = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    m.add("Facebook", "Instagram", "WhatsApp", "Telegram", "Binance", "PC Clone")
    return m


@bot.message_handler(content_types=["document"])
def document_handler(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    register_user(message.chat.id)

    doc = message.document
    name = doc.file_name or ""
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""

    if ext not in ("xlsx", "xls", "csv"):
        bot.send_message(
            message.chat.id,
            "âŒ <b>Unsupported file!</b>\n\n"
            "ðŸ“Ž Supported formats:\n"
            "  â€¢ <b>.xlsx</b> â€” Excel (new)\n"
            "  â€¢ <b>.xls</b>  â€” Excel (old)\n"
            "  â€¢ <b>.csv</b>  â€” CSV\n\n"
            "ðŸ’¡ File pathao abar!",
            parse_mode="HTML",
        )
        return

    wait = bot.send_message(
        message.chat.id, f"â³ðŸ”¥ <b>{name}</b> parse korchi...", parse_mode="HTML"
    )

    try:
        file_info = bot.get_file(doc.file_id)
        raw = bot.download_file(file_info.file_path)
    except Exception as e:
        bot.edit_message_text(
            f"âŒ File download hoyni: {e}",
            message.chat.id,
            wait.message_id,
            parse_mode="HTML",
        )
        return

    rows, mode = _parse_spreadsheet(raw, name)

    try:
        bot.delete_message(message.chat.id, wait.message_id)
    except Exception:
        pass

    if mode in ("unknown", "empty") or not rows:
        bot.send_message(
            message.chat.id,
            "âš ï¸ <b>File-e kono data paini!</b> âš ï¸\n\n"
            "ðŸ“‹ <b>Supported formats:</b>\n"
            "  â€¢ <b>2-column:</b>  Service | Number\n"
            "  â€¢ <b>1-column:</b>  Number only (service pore dao)\n\n"
            "ðŸ’¡ Sample format:\n"
            "<code>facebook  | 8801700123456\n"
            "whatsapp  | 8801800234567\n"
            "telegram  | 251912345678</code>",
            parse_mode="HTML",
        )
        return

    if mode == "two_col":
        # Group by service and add directly
        service_map = {}
        for svc, num in rows:
            service_map.setdefault(svc, []).append(num)

        total_added, total_skipped = 0, 0
        report_lines = ""
        for svc, nums in service_map.items():
            added, skipped = _add_numbers_bulk(svc, nums)
            total_added += added
            total_skipped += skipped
            icon = "âœ…" if added else "âš ï¸"
            report_lines += f"{icon} <b>{svc.upper()}</b>: +{added} added"
            if skipped:
                report_lines += f"  (âš ï¸ {skipped} skip)"
            report_lines += "\n"

        bot.send_message(
            message.chat.id,
            f"ðŸ“ŠðŸ”¥ <b>EXCEL IMPORT DONE!</b> ðŸ”¥ðŸ“Š\n"
            f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
            f"ðŸ“Ž <b>File:</b> <code>{name}</code>\n"
            f"ðŸ“‹ <b>Rows parsed:</b> {len(rows)}\n\n"
            f"{report_lines}\n"
            f"âœ… <b>Total added:</b> {total_added}\n"
            f"âš ï¸ <b>Skipped:</b> {total_skipped}\n\n"
            f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n"
            f"ðŸ’¡ /panels diye stock check koro.",
            reply_markup=main_menu(uid),
            parse_mode="HTML",
        )

    else:
        # one_col: ask which service
        _pending_excel[uid] = {"numbers": rows, "filename": name}
        bot.send_message(
            message.chat.id,
            f"ðŸ“‚ðŸ”¥ <b>FILE LOADED!</b> ðŸ”¥ðŸ“‚\n"
            f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
            f"ðŸ“Ž <b>File:</b> <code>{name}</code>\n"
            f"ðŸ“± <b>Numbers found:</b> {len(rows)}\n\n"
            f" <b>Kon service-e add korbo?</b>\n"
            f"â¬‡ï¸ Choose koro:",
            reply_markup=_service_select_markup(),
            parse_mode="HTML",
        )
        msg = bot.send_message(
            message.chat.id, "â¬‡ï¸ Service type koro:", parse_mode="HTML"
        )
        bot.register_next_step_handler(msg, _excel_pick_service)


def _excel_pick_service(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    svc_raw = (message.text or "").strip().lower()
    # normalise common aliases
    svc_map = {
        "facebook": "facebook",
        "fb": "facebook",
        "instagram": "instagram",
        "ig": "instagram",
        "whatsapp": "whatsapp",
        "wa": "whatsapp",
        "telegram": "telegram",
        "tg": "telegram",
        "binance": "binance",
        "bnb": "binance",
        "pc clone": "pc clone",
        "pc": "pc clone",
        "clone": "pc clone",
    }
    svc = svc_map.get(svc_raw)
    if svc is None:
        # try direct match
        for s in VALID_SERVICES:
            if svc_raw == s:
                svc = s
                break
    if svc is None:
        msg = bot.send_message(
            message.chat.id,
            "âŒ Valid service choose koro:\n"
            "<code>Facebook / Instagram / WhatsApp / Telegram / Binance / PC Clone</code>",
            reply_markup=_service_select_markup(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _excel_pick_service)
        return

    pending = _pending_excel.pop(uid, None)
    if not pending:
        bot.send_message(
            message.chat.id,
            "âš ï¸ Session expired. File abar pathao.",
            reply_markup=main_menu(uid),
        )
        return

    numbers = pending["numbers"]
    filename = pending["filename"]
    added, skipped = _add_numbers_bulk(svc, numbers)

    bot.send_message(
        message.chat.id,
        f"ðŸ“ŠðŸ”¥ <b>EXCEL IMPORT DONE!</b> ðŸ”¥ðŸ“Š\n"
        f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
        f"ðŸ“Ž <b>File:</b>     <code>{filename}</code>\n"
        f"ðŸ’¬ <b>Service:</b>  <b>{svc.upper()}</b>\n"
        f"ðŸ“± <b>Parsed:</b>   {len(numbers)}\n\n"
        f"âœ… <b>Added:</b>    {added}\n"
        f"âš ï¸ <b>Skipped:</b>  {skipped}\n\n"
        f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n"
        f"ðŸ’¡ /panels diye stock check koro.",
        reply_markup=main_menu(uid),
        parse_mode="HTML",
    )


@bot.message_handler(func=lambda m: True)
def text_handler(message):
    global stock
    uid = message.from_user.id
    txt = message.text
    register_user(message.chat.id)

    if txt == "â˜Žï¸ ð—¡ð—¨ð— ð—•ð—”ð—¥ â˜Žï¸":
        show_services(message)

    elif txt in _get_svc_map():
        svc = _get_svc_map()[txt]
        show_countries(message.chat.id, svc)

    elif txt == "ðŸ”™ Main Menu":
        mname = message.from_user.first_name or message.from_user.username or "User"
        bot.send_message(
            message.chat.id,
            f"â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—\n"
            f"      USER MENU-te WELCOME!\n"
            f"   ðŸ‘‹ <b>{mname}</b>, ki korte chao?\n"
            f"â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•",
            reply_markup=main_menu(uid),
            parse_mode="HTML",
        )

    elif txt == "ðŸ“ž ð—¦ð—”ð—£ð—¢ð—¥ð—§":
        markup = types.InlineKeyboardMarkup()
        markup.add(
            types.InlineKeyboardButton("ðŸ‘‘ RABBI â€” Support", url="https://t.me/Rabbi122q")
        )
        bot.send_message(
            message.chat.id,
            "ðŸ“ž <b>SUPPORT</b> ðŸ“ž\n"
            "âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
            "ðŸ‘‘ <b>RABBI</b>\n"
            "ðŸ“© à¦¯à§‡à¦•à§‹à¦¨à§‹ à¦¸à¦®à¦¸à§à¦¯à¦¾à¦¯à¦¼ à¦¨à¦¿à¦šà§‡à¦° à¦¬à¦¾à¦Ÿà¦¨à§‡ à¦•à§à¦²à¦¿à¦• à¦•à¦°à§‹!\n\n"
            "âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡",
            reply_markup=markup,
            parse_mode="HTML",
        )

    elif txt == "ðŸ“Š ð—¦ð—§ð—¢ð—–ð—ž":
        report = "ðŸ”¥ <b>LIVE STOCK REPORT</b> ðŸ”¥\nâš¡â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
        for s, d in stock.items():
            total = sum(len(v) for v in d.values())
            report += f" <b>{s.upper()}</b>: {total} à¦Ÿà¦¿ \n"
        report += "\nâš¡â”â”â”â”â”â”â”â”â”â”â”â”âš¡\nðŸ¤– <b>RABBI OTP BOT</b> ðŸ”¥"
        bot.send_message(message.chat.id, report, parse_mode="HTML")

    elif txt == "âš™ï¸ ð—”ð——ð— ð—œð—¡ ð—£ð—”ð—¡ð—˜ð—Ÿ âš™ï¸" and uid in ADMIN_IDS:
        _go_admin_panel(message)

    elif txt == "ðŸ”¥ðŸ“¢ ð—•ð—¿ð—¼ð—®ð—±ð—°ð—®ð˜€ð˜" and uid in ADMIN_IDS:
        msg = bot.send_message(
            message.chat.id,
            "âœï¸ <b>Broadcast content à¦ªà¦¾à¦ à¦¾à¦“:</b> \n\n"
            "ðŸ“ Text, ðŸ–¼ï¸ Photo, ðŸŽ¥ Video, or ðŸŽ­ Sticker (with optional caption) â€” à¦¸à¦¬ accept à¦¹à¦¬à§‡!\n\n"
            "ðŸ”™ Back jete <b>Admin Panel</b> button press koro.",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, do_broadcast)

    elif txt == "âš¡ðŸ‘¥ ð—¨ð˜€ð—²ð—¿ ð—–ð—¼ð˜‚ð—»ð˜" and uid in ADMIN_IDS:
        bot.send_message(
            message.chat.id,
            f" <b>TOTAL USERS</b> \n\nâš¡ <b>{len(users)}</b> à¦œà¦¨ à¦†à¦›à§‡! ðŸ”¥",
            parse_mode="HTML",
        )

    elif txt == "ðŸ“‹ðŸ‘¥ ð—¨ð˜€ð—²ð—¿ ð—Ÿð—¶ð˜€ð˜" and uid in ADMIN_IDS:
        all_ids = list(users)
        total = len(all_ids)
        if total == 0:
            bot.send_message(message.chat.id, "ðŸ“‹ No users yet.", parse_mode="HTML")
        else:
            bot.send_message(
                message.chat.id, "â³ Loading user names...", parse_mode="HTML"
            )
            updated = False
            for user_id in all_ids:
                key = str(user_id)
                existing = user_names.get(key, "")
                if existing and not existing.strip().lstrip("-").isdigit():
                    continue
                try:
                    chat_info = bot.get_chat(user_id)
                    full = f"{chat_info.first_name or ''} {chat_info.last_name or ''}".strip()
                    uname = chat_info.username or ""
                    if full and uname:
                        display = f"{full} (@{uname})"
                    elif full:
                        display = full
                    elif uname:
                        display = f"@{uname}"
                    else:
                        display = None
                    if display:
                        user_names[key] = display
                        updated = True
                except Exception:
                    pass
            if updated:
                save_json(USER_NAMES_FILE, user_names)

            PAGE = 50
            chunks = [all_ids[i : i + PAGE] for i in range(0, total, PAGE)]
            for idx, chunk in enumerate(chunks):
                lines = (
                    f"ðŸ“‹ðŸ‘¥ <b>USER LIST</b> ðŸ‘¥ðŸ“‹\n"
                    f"âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n"
                    f"ðŸ“Š Total: <b>{total}</b> users"
                    + (f"  |  Page {idx + 1}/{len(chunks)}" if len(chunks) > 1 else "")
                    + "\nâš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
                )
                for i, user_id in enumerate(chunk, start=idx * PAGE + 1):
                    name = user_names.get(str(user_id), "â€”")
                    lines += f"{i}. ðŸ†” <code>{user_id}</code>\n    ðŸ‘¤ {name}\n\n"
                bot.send_message(message.chat.id, lines, parse_mode="HTML")

    elif txt == "âž• ð—¡ð˜‚ð—ºð—¯ð—®ð—¿ ð—”ð—±ð—±" and uid in ADMIN_IDS:
        m = types.ReplyKeyboardMarkup(resize_keyboard=True)
        m.add("facebook", "instagram", "whatsapp", "telegram", "binance", "pc clone")
        m.add("âŒ Cancel")
        msg = bot.send_message(
            message.chat.id,
            "ðŸ”¥ <b>Service choose koro:</b> ðŸ”¥",
            reply_markup=m,
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, process_auto_add)

    elif txt == "ðŸ—‘ï¸ ð—¦ð—¼ð—¯ ð—–ð—¹ð—²ð—®ð—¿" and uid in ADMIN_IDS:
        bot.send_message(
            message.chat.id,
            "ðŸ—‘ï¸ðŸ”¥ <b>STOCK CLEAR PANEL</b> ðŸ”¥ðŸ—‘ï¸\n\n"
            " <b>Kon service-er stock clear korbe?</b>\n"
            "â¬‡ï¸ Service choose koro:",
            reply_markup=_clr_service_markup(),
            parse_mode="HTML",
        )

    elif txt == "ðŸŽ­ ð——ð—˜ð— ð—¢ ð—¢ð—§ð—£" and uid in ADMIN_IDS:
        bot.send_message(
            message.chat.id,
            demo_status_text(),
            reply_markup=demo_menu_markup(),
            parse_mode="HTML",
        )

    elif txt == "âž• ð—”ð—±ð—± ð—£ð—®ð—»ð—²ð—¹" and uid in ADMIN_IDS:
        _addpanel_state[uid] = {"step": "url", "data": {}}
        msg = bot.send_message(
            message.chat.id,
            "ðŸ”§ðŸ”¥ <b>ADD NEW PANEL</b> ðŸ”¥ðŸ”§\n\n"
            "ðŸ“¡ <b>Step 1/3:</b> Panel URL pathao\n\n"
            "Supported formats:\n"
            "â€¢ <code>http://1.2.3.4/ints/agent/SMSCDRStats</code>\n"
            "â€¢ <code>https://truesms.net/agent/SMSCDRStats</code>\n"
            "â€¢ <code>https://truesms.net/agent/SMSRanges</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _ap_get_url)

    elif txt == "âž• ð—”ð—±ð—± ð—¦ð—²ð—¿ð˜ƒð—¶ð—°ð—²" and uid in ADMIN_IDS:
        _addservice_state[uid] = {}
        msg = bot.send_message(
            message.chat.id,
            "ðŸ“‹ðŸ”¥ <b>ADD NEW SERVICE</b> ðŸ”¥ðŸ“‹\n\n"
            "ðŸ·ï¸ <b>Step 1/2:</b> Button-e ki lekha thakbe?\n"
            "<i>Example: Telegram ðŸ”µ, Binance ðŸ’›, TikTok ðŸŽµ</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _svc_get_label)

    elif txt == "ðŸ—‘ï¸ ð—¥ð—²ð—ºð—¼ð˜ƒð—² ð—¦ð—²ð—¿ð˜ƒð—¶ð—°ð—²" and uid in ADMIN_IDS:
        if not _services:
            bot.send_message(message.chat.id, "ðŸ“‹ Kono service nai!", parse_mode="HTML")
        else:
            markup = types.InlineKeyboardMarkup(row_width=1)
            for s in _services:
                markup.add(types.InlineKeyboardButton(
                    f"ðŸ—‘ï¸ {s['label']}  [{s['key']}]",
                    callback_data=f"rmsvc:{s['key']}",
                ))
            bot.send_message(
                message.chat.id,
                "ðŸ—‘ï¸ðŸ”¥ <b>REMOVE SERVICE</b>\n\nKon service remove korbe?",
                reply_markup=markup,
                parse_mode="HTML",
            )

    elif txt == "ðŸ—‘ï¸ ð—¥ð—²ð—ºð—¼ð˜ƒð—² ð—£ð—®ð—»ð—²ð—¹" and uid in ADMIN_IDS:
        if not _dynamic_panels:
            bot.send_message(
                message.chat.id,
                "ðŸ“‹ <b>Kono dynamic panel nai!</b>\nðŸ’¡ Add Panel button diye add koro.",
                parse_mode="HTML",
            )
        else:
            markup = types.InlineKeyboardMarkup(row_width=1)
            for p in _dynamic_panels:
                pid = p["id"]
                with _stats_lock:
                    s = _panel_stats.get(pid, {})
                st = s.get("status", "â³")
                markup.add(
                    types.InlineKeyboardButton(
                        f"{st} {p.get('username','?')} â€” {p.get('host','?')}",
                        callback_data=f"rmpanel:{pid}",
                    )
                )
            bot.send_message(
                message.chat.id,
                "ðŸ—‘ï¸ðŸ”¥ <b>REMOVE PANEL</b>\n\nKon panel remove korbe?",
                reply_markup=markup,
                parse_mode="HTML",
            )

    elif txt == "â–¶ï¸ ð——ð—˜ð— ð—¢ ð—¦ð—§ð—”ð—¥ð—§" and uid in ADMIN_IDS:
        global _demo_active
        with _demo_lock:
            _demo_active = True
        bot.send_message(
            message.chat.id,
            "ðŸŸ¢ðŸ”¥ <b>DEMO OTP STARTED!</b> ðŸ”¥ðŸŸ¢\nâš¡ Group-e fake OTP pathano shuru hoyeche!",
            reply_markup=demo_menu_markup(),
            parse_mode="HTML",
        )

    elif txt == "â¹ï¸ ð——ð—˜ð— ð—¢ ð—¦ð—§ð—¢ð—£" and uid in ADMIN_IDS:
        with _demo_lock:
            _demo_active = False
        bot.send_message(
            message.chat.id,
            "ðŸ”´ <b>DEMO OTP STOPPED!</b> ðŸ”´\nâš¡ Fake OTP pathano bondho hoyeche.",
            reply_markup=demo_menu_markup(),
            parse_mode="HTML",
        )

    elif txt == "âš™ï¸ ð——ð—˜ð— ð—¢ ð—–ð—¢ð—¡ð—™ð—œð—š" and uid in ADMIN_IDS:
        msg = bot.send_message(
            message.chat.id,
            "ðŸ“± <b>Phone number(s) dao:</b>\n\n"
            "â€¢ Ekta number: <code>8801700123456</code>\n"
            "â€¢ Multiple (newline or comma):\n"
            "<code>8801700123456\n251912345678\n2348012345678</code>\n\n"
            "âš ï¸ Full country code including number lagbe!",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _demo_cfg_number)

    elif txt == "ðŸ“Š ð—£ð—®ð—»ð—²ð—¹ð˜€" and uid in ADMIN_IDS:
        panels_cmd(message)

    elif txt == "ðŸ‘‘ ð—”ð—±ð—± ð—”ð—±ð—ºð—¶ð—»" and uid in ADMIN_IDS:
        msg = bot.send_message(
            message.chat.id,
            "ðŸ‘‘ <b>New Admin add</b>\n\n"
            "Notun admin-er Telegram <b>User ID</b> dao:\n"
            "<i>Example: 123456789</i>\n\n"
            "ðŸ’¡ User ID jante hole sei user-ke @userinfobot-e forward koro.",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _admin_add_get_id)

    elif txt == "ðŸ—‘ï¸ ð—¥ð—²ð—ºð—¼ð˜ƒð—² ð—”ð—±ð—ºð—¶ð—»" and uid in ADMIN_IDS:
        _show_remove_admin(message)

    elif txt == "âš™ï¸ ð—¦ð—²ð˜ð˜ð—¶ð—»ð—´ð˜€" and uid in ADMIN_IDS:
        _show_settings(message)

    elif txt in ("ðŸ”™ ð—”ð——ð— ð—œð—¡ ð—£ð—”ð—¡ð—˜ð—Ÿ", "ðŸ”™ Admin Panel") and uid in ADMIN_IDS:
        _go_admin_panel(message)

    elif txt == "â¬…ï¸ðŸ”™ ð—¨ð˜€ð—²ð—¿ ð— ð—²ð—»ð˜‚":
        mname = message.from_user.first_name or message.from_user.username or "User"
        bot.send_message(
            message.chat.id,
            f"â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—\n"
            f"      USER MENU-te WELCOME!\n"
            f"   ðŸ‘‹ <b>{mname}</b>, ki korte chao?\n"
            f"â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•",
            reply_markup=main_menu(uid),
            parse_mode="HTML",
        )


# â”€â”€ Demo OTP config step handlers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


def _demo_cfg_number(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    raw_lines = re.split(r"[\n,]+", message.text or "")
    candidates = [re.sub(r"\D", "", ln) for ln in raw_lines if re.sub(r"\D", "", ln)]
    if not candidates:
        msg = bot.send_message(
            message.chat.id,
            "âŒ Kono number paini. Ekta ba multiple number dao:",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _demo_cfg_number)
        return
    valid, invalid = [], []
    result_lines = ""
    for num in candidates:
        if len(num) < 7:
            invalid.append(num)
            continue
        c_name, flag = get_country_details(num)
        if c_name == "Unknown":
            invalid.append(num)
        else:
            valid.append(num)
            result_lines += f"  âœ… <code>{num}</code>  {flag} {c_name}\n"
    if not valid:
        msg = bot.send_message(
            message.chat.id,
            f"âš ï¸ <b>Kono valid number paini!</b>\n\n"
            f"Full international number dao (country code including):\n"
            f"ðŸ‡§ðŸ‡© Bangladesh â†’ <code>8801700123456</code>\n"
            f"ðŸ‡ªðŸ‡¹ Ethiopia   â†’ <code>251912345678</code>\n"
            f"ðŸ‡³ðŸ‡¬ Nigeria    â†’ <code>2348012345678</code>\n\n"
            f"Aro ekbar try koro:",
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _demo_cfg_number)
        return
    with _demo_lock:
        _demo_config["numbers"] = valid
    SHOW_MAX = 10
    shown = result_lines.split("\n")[:SHOW_MAX]
    preview = "\n".join(shown)
    if len(valid) > SHOW_MAX:
        preview += f"\n  ... +{len(valid) - SHOW_MAX} more"
    feedback = f"âœ… <b>{len(valid)} à¦Ÿà¦¿ number set hoiche:</b>\n{preview}\n"
    if invalid:
        inv_preview = invalid[:5]
        feedback += (
            f"\nâš ï¸ Skip (invalid): {', '.join(f'<code>{x}</code>' for x in inv_preview)}"
        )
        if len(invalid) > 5:
            feedback += f" +{len(invalid) - 5} more"
        feedback += "\n"
    svc_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    svc_markup.add("4", "5", "6", "7", "8")
    svc_markup.add("ðŸ”™ Admin Panel")
    msg = bot.send_message(
        message.chat.id,
        feedback + "\nðŸ”¢ <b>OTP digit count choose koro:</b>",
        reply_markup=svc_markup,
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _demo_cfg_digits)


def _demo_cfg_digits(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    try:
        d = int(message.text.strip())
        if d < 4 or d > 8:
            raise ValueError
    except ValueError:
        svc_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        svc_markup.add("4", "5", "6", "7", "8")
        svc_markup.add("ðŸ”™ Admin Panel")
        msg = bot.send_message(message.chat.id, "âŒ 4 theke 8 er modhye number dao:", reply_markup=svc_markup)
        bot.register_next_step_handler(msg, _demo_cfg_digits)
        return
    with _demo_lock:
        _demo_config["digits"] = d
    svc_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    svc_markup.add("Facebook", "Instagram", "WhatsApp", "Telegram", "PC Clone")
    svc_markup.add("ðŸ”™ Admin Panel")
    msg = bot.send_message(
        message.chat.id,
        f"âœ… Digits set: <b>{d}</b>\n\nðŸ’¬ <b>Service choose koro:</b>",
        reply_markup=svc_markup,
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _demo_cfg_service)


def _demo_cfg_service(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    svc = (message.text or "").strip()
    if not svc:
        svc_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        svc_markup.add("Facebook", "Instagram", "WhatsApp", "Telegram", "PC Clone")
        svc_markup.add("ðŸ”™ Admin Panel")
        msg = bot.send_message(message.chat.id, "âŒ Service name dao:", reply_markup=svc_markup)
        bot.register_next_step_handler(msg, _demo_cfg_service)
        return
    with _demo_lock:
        _demo_config["service"] = svc
    intvl_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    intvl_markup.add("15", "30", "60", "120", "300")
    intvl_markup.add("ðŸ”™ Admin Panel")
    msg = bot.send_message(
        message.chat.id,
        f"âœ… Service set: <b>{svc}</b>\n\nâ±ï¸ <b>Interval (seconds) dao:</b>",
        reply_markup=intvl_markup,
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _demo_cfg_interval)


def _demo_cfg_interval(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    try:
        iv = int(message.text.strip())
        if iv < 5:
            raise ValueError
    except ValueError:
        intvl_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        intvl_markup.add("15", "30", "60", "120", "300")
        intvl_markup.add("ðŸ”™ Admin Panel")
        msg = bot.send_message(message.chat.id, "âŒ Minimum 5 second. Aro dao:", reply_markup=intvl_markup)
        bot.register_next_step_handler(msg, _demo_cfg_interval)
        return
    with _demo_lock:
        _demo_config["interval"] = iv
    bot.send_message(
        message.chat.id,
        f"âœ… Interval set: <b>{iv}s</b>\n\n" + demo_status_text(),
        reply_markup=demo_menu_markup(),
        parse_mode="HTML",
    )


def make_broadcast_msg(text):
    return (
        "ðŸ”¥ <b>ð—”ð—¥ ð—¢ð—§ð—£ ð—•ð—¢ð—§ â€” ð—•ð—¥ð—¢ð—”ð——ð—–ð—”ð—¦ð—§!</b> ðŸ”¥\n"
        "âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
        f"ðŸ“¢ {text} ðŸ“¢\n\n"
        "âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n"
        "ðŸ¤–ðŸ”¥ <i>ð™‹ð™¤ð™¬ð™šð™§ð™šð™™ ð™—ð™®</i>  <b>ð—”ð—¥ ð—¢ð—§ð—£ ð—•ð—¢ð—§</b>  ðŸ”¥ðŸ¤–"
    )


def do_broadcast(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    has_text = bool(message.text)
    has_photo = bool(message.photo)
    has_video = bool(message.video)
    has_sticker = bool(message.sticker)
    has_animation = bool(message.animation)
    has_audio = bool(message.audio)
    has_voice = bool(message.voice)
    has_document = bool(message.document)
    has_video_note = bool(message.video_note)

    if not any(
        [
            has_text,
            has_photo,
            has_video,
            has_sticker,
            has_animation,
            has_audio,
            has_voice,
            has_document,
            has_video_note,
        ]
    ):
        bot.send_message(
            message.chat.id,
            "âš ï¸ <b>Kono content paoa jaini!</b> âš ï¸\n"
            "Text, Photo, Video, GIF, Audio, Voice, Document ba Sticker pathao.",
            parse_mode="HTML",
        )
        return

    cap = (
        lambda m: make_broadcast_msg(m.caption) if m.caption else make_broadcast_msg("")
    )

    bot.send_message(
        message.chat.id,
        f"â³ðŸ”¥ <b>{len(users)} à¦œà¦¨à¦•à§‡ à¦ªà¦¾à¦ à¦¾à¦¨à§‹ à¦¹à¦šà§à¦›à§‡...</b> ðŸ”¥â³",
        parse_mode="HTML",
    )

    success, fail = 0, 0
    for uid in list(users):
        try:
            if has_photo:
                bot.send_photo(
                    uid,
                    message.photo[-1].file_id,
                    caption=cap(message),
                    parse_mode="HTML",
                )
            elif has_animation:
                bot.send_animation(
                    uid,
                    message.animation.file_id,
                    caption=cap(message),
                    parse_mode="HTML",
                )
            elif has_video:
                bot.send_video(
                    uid, message.video.file_id, caption=cap(message), parse_mode="HTML"
                )
            elif has_video_note:
                bot.send_video_note(uid, message.video_note.file_id)
            elif has_sticker:
                bot.send_sticker(uid, message.sticker.file_id)
            elif has_audio:
                bot.send_audio(
                    uid, message.audio.file_id, caption=cap(message), parse_mode="HTML"
                )
            elif has_voice:
                bot.send_voice(
                    uid, message.voice.file_id, caption=cap(message), parse_mode="HTML"
                )
            elif has_document:
                bot.send_document(
                    uid,
                    message.document.file_id,
                    caption=cap(message),
                    parse_mode="HTML",
                )
            else:
                bot.send_message(
                    uid, make_broadcast_msg(message.text), parse_mode="HTML"
                )
            success += 1
        except Exception:
            fail += 1

    bot.send_message(
        message.chat.id,
        f" <b>BROADCAST COMPLETE!</b> \n\n"
        f"âœ… <b>ð—¦ð—¼ð—³ð—¼ð—¹:</b> {success} à¦œà¦¨ ðŸ”¥\n"
        f"âŒ <b>ð—•ð—®ð—¿ð˜ð—µð—¼:</b> {fail} à¦œà¦¨ ",
        reply_markup=main_menu(message.from_user.id),
        parse_mode="HTML",
    )


_pending_add = {}


def _start_countdown(chat_id, msg_id, svc, flag, c_name, display_num, scnt):
    if chat_id in _countdowns:
        _countdowns[chat_id].set()
    cancel = threading.Event()
    _countdowns[chat_id] = cancel

    def run():
        total = 600
        while not cancel.is_set():
            mins = total // 60
            secs = total % 60
            text = (
                f"âœ… <b>Number Assigned Successfully !</b>\n\n"
                f"ðŸ”§ <b>Platform :</b> {svc.capitalize()}\n"
                f"ðŸŒ <b>Country :</b> {flag} {c_name}\n\n"
                f"ðŸ“ž <b>Number :</b> <code>{display_num}</code>\n\n"
                f"â± <b>Auto code fetch :</b> {mins:02d}:{secs:02d}s"
            )
            kb = types.InlineKeyboardMarkup(row_width=2)
            kb.add(
                types.InlineKeyboardButton("ðŸ”„ New Number", callback_data=f"n:{svc}:{scnt}"),
                types.InlineKeyboardButton("ðŸŒ Change Country", callback_data=f"s:{svc}"),
            )
            kb.add(
                types.InlineKeyboardButton("ðŸ“¢ OTP Group", url=get_otp_group_link()),
            )
            try:
                bot.edit_message_text(
                    text, chat_id, msg_id,
                    reply_markup=kb, parse_mode="HTML",
                )
            except Exception:
                pass
            cancel.wait(5)
            if cancel.is_set():
                break
            total -= 5
            if total < 0:
                total = 600

    threading.Thread(target=run, daemon=True).start()


def _settings_text(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    cfg = get_admin_config(uid)
    grp_id   = cfg.get("group_id") or _group_settings.get("otp_group_id")
    grp_link = cfg.get("group_link") or _group_settings.get("otp_group_link", "")
    brand    = cfg.get("brand") or "RABBI"
    num_ch   = cfg.get("numChannel") or ""
    main_ch  = cfg.get("mainChannel") or ""
    bot_lnk  = cfg.get("botLink") or ""
    auto_del = _group_settings.get("auto_delete", True)
    del_secs = _group_settings.get("auto_delete_seconds", 3600)

    id_str   = f"<code>{grp_id}</code>" if grp_id else "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    link_str = grp_link or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    brand_str = brand
    num_str  = num_ch or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    main_str = main_ch or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    bot_str  = bot_lnk or "âŒ à¦¸à§‡à¦Ÿ à¦¨à§‡à¦‡"
    auto_str = f"ðŸŸ¢ ON ({del_secs // 60} min)" if auto_del else "ðŸ”´ OFF"
    return (
        f"âš™ï¸ <b>à¦†à¦ªà¦¨à¦¾à¦° SETTINGS</b> (Admin {uid})\n"
        "âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n\n"
        "ðŸ“¡ <b>OTP GROUP</b>\n"
        f"  ðŸ”— Link: {link_str}\n"
        f"  ðŸ†” Chat ID: {id_str}\n"
        f"  â±ï¸ Auto Delete: {auto_str}\n\n"
        "ðŸŽ¨ <b>BRAND & CHANNELS</b>\n"
        f"  ðŸ‘‘ Brand: <b>{brand_str}</b>\n"
        f"  ðŸ“² Number Ch: {num_str}\n"
        f"  ðŸ“¢ Main Ch: {main_str}\n"
        f"  ðŸ¤– Bot Link: {bot_str}\n\n"
        "âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n"
        "â¬‡ï¸ à¦•à§€ change à¦•à¦°à¦¤à§‡ à¦šà¦¾à¦“?"
    )


def _settings_markup(uid=None):
    if uid is None:
        uid = SUPER_ADMIN_ID
    auto_del = _group_settings.get("auto_delete", True)
    auto_label = "â±ï¸ Auto Delete: ðŸŸ¢ ON" if auto_del else "â±ï¸ Auto Delete: ðŸ”´ OFF"
    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
        types.InlineKeyboardButton("ðŸ†” OTP Group Set", callback_data=f"my_setgroup:{uid}"),
        types.InlineKeyboardButton("ðŸ‘‘ Brand Set", callback_data=f"my_setbrand:{uid}"),
    )
    markup.add(
        types.InlineKeyboardButton("ðŸ“² Number Ch", callback_data=f"my_setnumch:{uid}"),
        types.InlineKeyboardButton("ðŸ“¢ Main Ch", callback_data=f"my_setmainch:{uid}"),
    )
    markup.add(
        types.InlineKeyboardButton("ðŸ¤– Bot Link", callback_data=f"my_setbotlink:{uid}"),
        types.InlineKeyboardButton(auto_label, callback_data="set_autodel"),
    )
    return markup


def _show_settings(message):
    uid = message.from_user.id
    bot.send_message(
        message.chat.id,
        _settings_text(uid),
        reply_markup=_settings_markup(uid),
        parse_mode="HTML",
    )


def _show_settings_inline(call):
    uid = call.from_user.id
    try:
        bot.edit_message_text(
            _settings_text(uid),
            call.message.chat.id,
            call.message.message_id,
            reply_markup=_settings_markup(uid),
            parse_mode="HTML",
        )
    except Exception:
        pass


def _show_group_settings(message):
    _show_settings(message)


def _show_group_settings_inline(call):
    _show_settings_inline(call)


def _grp_get_link(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    link = (message.text or "").strip()
    if not link.startswith("https://t.me/") and not link.startswith("http://"):
        msg = bot.send_message(
            message.chat.id,
            "âŒ Valid Telegram link dao:\n<i>Example: https://t.me/aR_OTP_rcv</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _grp_get_link)
        return
    _group_settings["otp_group_link"] = link
    save_group_settings()
    _go_admin_panel(
        message,
        f"âœ…ðŸ”¥ <b>GROUP LINK UPDATED!</b>\n\n"
        f"ðŸ”— <b>Notun Link:</b> {link}\n\n"
        f"<i>Ekhon theke number-er nichor OTP Group button-e ei link thakbe.</i>",
    )


def _grp_get_id(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    raw = (message.text or "").strip()
    try:
        gid = int(raw)
    except ValueError:
        msg = bot.send_message(
            message.chat.id,
            "âŒ Valid Chat ID dao (number):\n<i>Example: -1001234567890</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _grp_get_id)
        return
    _group_settings["otp_group_id"] = gid
    save_group_settings()
    _go_admin_panel(
        message,
        f"âœ…ðŸ”¥ <b>GROUP CHAT ID UPDATED!</b>\n\n"
        f"ðŸ†” <b>Notun Chat ID:</b> <code>{gid}</code>\n\n"
        f"<i>Ekhon theke OTP ei group-e pathano hobe.</i>",
    )


def _sett_get_channel2(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    link = (message.text or "").strip()
    if not link.startswith("https://") and not link.startswith("http://"):
        msg = bot.send_message(
            message.chat.id,
            "âŒ Valid link dao:\n<i>Example: https://t.me/aR_OTP_rcv</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _sett_get_channel2)
        return
    _group_settings["channel2"] = link
    save_group_settings()
    _go_admin_panel(
        message,
        f"âœ… <b>JOIN CHANNEL UPDATED!</b>\n\n"
        f"ðŸ“¢ <b>Notun Link:</b> {link}",
    )


def _sett_get_botlink(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    link = (message.text or "").strip()
    if not link.startswith("https://") and not link.startswith("http://"):
        msg = bot.send_message(
            message.chat.id,
            "âŒ Valid link dao:\n<i>Example: https://t.me/ar_otp_bot</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _sett_get_botlink)
        return
    _group_settings["bot_link"] = link
    save_group_settings()
    _go_admin_panel(
        message,
        f"âœ… <b>BOT LINK UPDATED!</b>\n\n"
        f"ðŸ¤– <b>Notun Link:</b> {link}",
    )


def _admin_add_get_id(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    raw = (message.text or "").strip()
    try:
        new_uid = int(raw)
    except ValueError:
        msg = bot.send_message(
            message.chat.id,
            "âŒ Valid Telegram User ID dao (shudhu number):\n<i>Example: 123456789</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _admin_add_get_id)
        return
    if add_admin(new_uid):
        _go_admin_panel(
            message,
            f"âœ… <b>ADMIN ADDED!</b>\n\n"
            f"ðŸ‘‘ <b>New Admin ID:</b> <code>{new_uid}</code>\n\n"
            f"<i>Ekhon theke ei user admin panel access pabe.</i>",
        )
    else:
        _go_admin_panel(
            message,
            f"âš ï¸ <b>User <code>{new_uid}</code> already admin ache!</b>",
        )


def _show_remove_admin(message):
    removable = [a for a in ADMIN_IDS if a != SUPER_ADMIN_ID]
    if not removable:
        bot.send_message(
            message.chat.id,
            "â„¹ï¸ <b>Remove korar moto kono extra admin nei.</b>\n\n"
            "<i>Super Admin remove kora jabe na.</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        return
    markup = types.InlineKeyboardMarkup(row_width=1)
    for aid in removable:
        name = user_names.get(str(aid), {}).get("first_name", "") or str(aid)
        markup.add(types.InlineKeyboardButton(
            f"ðŸ—‘ï¸ {name} [{aid}]", callback_data=f"rmadmin:{aid}"
        ))
    bot.send_message(
        message.chat.id,
        "ðŸ—‘ï¸ <b>Remove Admin</b>\n\n"
        "âš¡â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”âš¡\n"
        "Niche theke admin select koro:\n\n"
        "<i>âš ï¸ Super Admin remove kora jabe na.</i>",
        reply_markup=markup,
        parse_mode="HTML",
    )


def _go_admin_panel(message, text="ðŸ”¥ <b>ADMIN PANEL</b>"):
    m_admin = types.ReplyKeyboardMarkup(resize_keyboard=True)
    m_admin.add("âž• ð—¡ð˜‚ð—ºð—¯ð—®ð—¿ ð—”ð—±ð—±", "ðŸ—‘ï¸ ð—¦ð—¼ð—¯ ð—–ð—¹ð—²ð—®ð—¿")
    m_admin.add("ðŸ”¥ðŸ“¢ ð—•ð—¿ð—¼ð—®ð—±ð—°ð—®ð˜€ð˜", "âš¡ðŸ‘¥ ð—¨ð˜€ð—²ð—¿ ð—–ð—¼ð˜‚ð—»ð˜")
    m_admin.add("ðŸ“‹ðŸ‘¥ ð—¨ð˜€ð—²ð—¿ ð—Ÿð—¶ð˜€ð˜")
    m_admin.add("ðŸŽ­ ð——ð—˜ð— ð—¢ ð—¢ð—§ð—£")
    m_admin.add("âž• ð—”ð—±ð—± ð—£ð—®ð—»ð—²ð—¹", "ðŸ—‘ï¸ ð—¥ð—²ð—ºð—¼ð˜ƒð—² ð—£ð—®ð—»ð—²ð—¹")
    m_admin.add("âž• ð—”ð—±ð—± ð—¦ð—²ð—¿ð˜ƒð—¶ð—°ð—²", "ðŸ—‘ï¸ ð—¥ð—²ð—ºð—¼ð˜ƒð—² ð—¦ð—²ð—¿ð˜ƒð—¶ð—°ð—²")
    m_admin.add("ðŸ“Š ð—£ð—®ð—»ð—²ð—¹ð˜€")
    m_admin.add("ðŸ‘‘ ð—”ð—±ð—± ð—”ð—±ð—ºð—¶ð—»", "ðŸ—‘ï¸ ð—¥ð—²ð—ºð—¼ð˜ƒð—² ð—”ð—±ð—ºð—¶ð—»")
    m_admin.add("âš™ï¸ ð—¦ð—²ð˜ð˜ð—¶ð—»ð—´ð˜€")
    m_admin.add("â¬…ï¸ðŸ”™ ð—¨ð˜€ð—²ð—¿ ð— ð—²ð—»ð˜‚")
    bot.send_message(
        message.chat.id,
        text,
        reply_markup=m_admin,
        parse_mode="HTML",
    )


def _cancel_kb():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("âŒ Cancel")
    return kb


def _back_admin_kb():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("ðŸ”™ Admin Panel")
    return kb


def _is_back(txt):
    return (txt or "").strip() in ("ðŸ”™ Admin Panel", "âŒ Cancel")


def process_auto_add(message):
    svc = (message.text or "").strip().lower()
    if svc == "âŒ cancel":
        _go_admin_panel(message)
        return
    if svc not in stock:
        m = types.ReplyKeyboardMarkup(resize_keyboard=True)
        m.add("facebook", "instagram", "whatsapp", "telegram", "binance", "pc clone")
        m.add("âŒ Cancel")
        msg = bot.send_message(
            message.chat.id,
            " <b>Vul service! Abar choose koro:</b>",
            reply_markup=m,
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, process_auto_add)
        return
    msg = bot.send_message(
        message.chat.id,
        f"ðŸ”¥ <b>{svc.upper()}</b>\n\n"
        f"ðŸ“ <b>Slot name dao:</b>\n"
        f"<i>Udharan: Mali 1, Germany 2, India 3</i>",
        reply_markup=_cancel_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, lambda m: ask_numbers_for_slot(m, svc))


def ask_numbers_for_slot(message, svc):
    slot_name = (message.text or "").strip()
    if slot_name == "âŒ Cancel":
        _go_admin_panel(message)
        return
    if not slot_name:
        msg = bot.send_message(
            message.chat.id,
            "âŒ Slot name dao:",
            reply_markup=_cancel_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, lambda m: ask_numbers_for_slot(m, svc))
        return
    msg = bot.send_message(
        message.chat.id,
        f"âœ… Slot: <b>{slot_name}</b>\n\n"
        f"ðŸ“± Ekhon <b>{svc.upper()}</b> er number gulo pathao:\n"
        f"<i>(Newline ba comma diye alag koro)</i>",
        reply_markup=_cancel_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, lambda m: finalize_auto_add(m, svc, slot_name))


def finalize_auto_add(message, svc, slot_name=None):
    global stock
    uid = message.from_user.id
    if (message.text or "").strip() == "âŒ Cancel":
        _go_admin_panel(message)
        return
    nums = [n.strip() for n in re.split(r"[,\n\r]", message.text) if n.strip()]
    if slot_name:
        if slot_name not in stock[svc]:
            stock[svc][slot_name] = []
        for num in nums:
            stock[svc][slot_name].append(num)
        added_count = len(nums)
    else:
        added_count = 0
        for num in nums:
            c_name, _ = get_country_details(num)
            if c_name == "Unknown":
                continue
            if c_name not in stock[svc]:
                stock[svc][c_name] = []
            stock[svc][c_name].append(num)
            added_count += 1
    save_stock()
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add("âž• Aro Add koro", "ðŸ”™ Admin Menu")
    bot.send_message(
        message.chat.id,
        f"âœ…ðŸ”¥ <b>DONE!</b>\n\n"
        f"ðŸ—‚ <b>Slot:</b> {slot_name or 'Auto'}\n"
        f"ðŸ“± <b>Added:</b> {added_count} à¦Ÿà¦¿ number",
        reply_markup=markup,
        parse_mode="HTML",
    )
    bot.register_next_step_handler(
        bot.send_message(message.chat.id, "â¬‡ï¸ Ki korbe?", parse_mode="HTML"),
        lambda m: _after_add_handler(m, svc),
    )


def _after_add_handler(message, last_svc):
    txt = (message.text or "").strip()
    if txt == "âž• Aro Add koro":
        msg = bot.send_message(
            message.chat.id,
            f"ðŸ“ <b>Notun slot name dao:</b>\n<i>Udharan: Mali 2, Germany 3</i>",
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, lambda m: ask_numbers_for_slot(m, last_svc))
    else:
        bot.send_message(
            message.chat.id,
            "ðŸ”™ Admin Menu",
            reply_markup=main_menu(message.from_user.id),
            parse_mode="HTML",
        )


# â”€â”€ Heartbeat / watchdog â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€



# â”€â”€ Start â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

try:
    requests.get(
        f"https://api.telegram.org/bot{API_TOKEN}/deleteWebhook?drop_pending_updates=true",
        timeout=10,
    )
    print("[START] Webhook cleared.")
except Exception as e:
    print(f"[START] Webhook clear failed: {e}")

time.sleep(3)

threading.Thread(target=panel1_monitor, daemon=True).start()
threading.Thread(target=panel2_monitor, daemon=True).start()
threading.Thread(target=panel3_monitor, daemon=True).start()
threading.Thread(target=panel4_monitor, daemon=True).start()
threading.Thread(target=panel5_monitor, daemon=True).start()
threading.Thread(target=panel6_monitor, daemon=True).start()
threading.Thread(target=demo_monitor, daemon=True).start()

for _dp in _dynamic_panels:
    _start_dynamic_panel(_dp)
    print(f"[DYN] Loaded saved panel: {_dp['id']} ({_dp['host']})")

_ts_boot = time.strftime("%H:%M:%S")
brand_main = get_brand(SUPER_ADMIN_ID)
print(f"[{_ts_boot}] â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—")
print(f"[{_ts_boot}] â•‘  ðŸ”¥ RABBI OTP BOT â€” MULTI-TENANT v2.0  â•‘")
print(f"[{_ts_boot}] â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•")
print(f"[{_ts_boot}] ðŸš€ Starting 6 hardcoded panels + dynamic panels...")
print(f"[{_ts_boot}] ðŸ‘‘ Main Brand: {brand_main}")
print(f"[{_ts_boot}] âš™ï¸  Poll interval: {POLL_INTERVAL}s")
print(f"[{_ts_boot}] ðŸ“ Admin configs dir: {ADMIN_CONFIGS_DIR}/")
print(f"[{_ts_boot}] ðŸ“‹ Dynamic panels loaded: {len(_dynamic_panels)}")


def _clear_webhook():
    try:
        requests.get(
            f"https://api.telegram.org/bot{API_TOKEN}/deleteWebhook?drop_pending_updates=true",
            timeout=10,
        )
    except Exception:
        pass


while True:
    try:
        _clear_webhook()
        time.sleep(2)
        bot.infinity_polling(
            timeout=60,
            long_polling_timeout=60,
            allowed_updates=["message", "callback_query"],
        )
    except requests.exceptions.ReadTimeout:
        print("[POLLING] ReadTimeout â€” restarting in 3s...")
        time.sleep(3)
    except requests.exceptions.ConnectionError:
        print("[POLLING] ConnectionError â€” restarting in 5s...")
        time.sleep(5)
    except Exception as e:
        print(f"[POLLING] Error: {e} â€” restarting in 5s...")
        time.sleep(5)
