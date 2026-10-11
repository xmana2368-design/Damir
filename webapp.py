"""
JIDU Mini App — backend qismi.

Bu fayl asosiy bot faylingiz (bot.py) bilan BIR PAPKADA turishi kerak.
Asosiy bot funksiyalaringiz (search_entities, fetch_week_events, get_room_statuses
va h.k.) shu yerdan qayta ishlatiladi — kodni takrorlash shart emas.

Ulash: bot.py ichidagi main() da
    from webapp import register_webapp
    register_webapp(app, sys.modules[__name__])
"""
import asyncio
import hmac
import json
import time
import hashlib
import datetime
from pathlib import Path
from urllib.parse import parse_qsl

from aiohttp import web

PAGE_PATH = Path(__file__).with_name("webapp.html")
CODES = {"s", "t", "g"}


# ---------------- Telegram initData tekshiruvi ----------------
def _validate_init_data(init_data: str, token: str, max_age: int = 86400 * 3):
    """Telegram Mini App yuborgan initData haqiqiyligini tekshiradi.
    To'g'ri bo'lsa foydalanuvchi dict'ini, aks holda None qaytaradi."""
    try:
        pairs = dict(parse_qsl(init_data, keep_blank_values=True))
        received = pairs.pop("hash", None)
        if not received:
            return None
        check = "\n".join(f"{k}={v}" for k, v in sorted(pairs.items()))
        secret = hmac.new(b"WebAppData", token.encode(), hashlib.sha256).digest()
        calc = hmac.new(secret, check.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(calc, received):
            return None
        if time.time() - int(pairs.get("auth_date", "0")) > max_age:
            return None
        return json.loads(pairs["user"])
    except Exception:
        return None


@web.middleware
async def auth_middleware(request: web.Request, handler):
    if request.path.startswith("/api/"):
        m = request.app["bot_mod"]
        user = _validate_init_data(request.headers.get("X-Init-Data", ""), m.TOKEN)
        if not user:
            return web.json_response({"error": "unauthorized"}, status=401)
        uid = user.get("id")
        if uid != m.ADMIN_ID and await m.db_is_banned(uid):
            return web.json_response({"error": "banned"}, status=403)
        request["tg_user"] = user
    return await handler(request)


def _bad(msg: str = "bad request") -> web.Response:
    return web.json_response({"error": msg}, status=400)


def _valid_date(value: str | None, m) -> str:
    try:
        datetime.datetime.strptime(value or "", "%Y-%m-%d")
        return value
    except (ValueError, TypeError):
        return m.now_tashkent().date().isoformat()


# ---------------- Sahifa ----------------
async def page(request: web.Request) -> web.Response:
    html = PAGE_PATH.read_text(encoding="utf-8")
    return web.Response(
        text=html,
        content_type="text/html",
        headers={"Cache-Control": "no-store"},
    )


# ---------------- API ----------------
async def api_structure(request: web.Request) -> web.Response:
    m = request.app["bot_mod"]
    data = {
        block: [f for f, rooms in floors.items() if rooms]
        for block, floors in m.UNIVERSITY_MAP.items()
    }
    return web.json_response({"blocks": data})


async def api_search(request: web.Request) -> web.Response:
    m = request.app["bot_mod"]
    code = request.query.get("code", "s")
    q = " ".join(request.query.get("q", "").split())
    if code not in CODES or len(q) < 2 or len(q) > 60:
        return _bad()

    results = await m.search_entities(m.CODE_TO_ENTITY_TYPE[code], q)
    out = []
    for item in results[:12]:
        item_id = item.get("id")
        if not item_id:
            continue
        name = m.entity_display_name(item)
        m.ENTITY_NAME_CACHE[str(item_id)] = name
        out.append({"id": str(item_id), "name": name})
    return web.json_response({"results": out})


async def api_schedule(request: web.Request) -> web.Response:
    m = request.app["bot_mod"]
    code = request.query.get("code", "s")
    entity_id = request.query.get("id", "")
    if code not in CODES or not entity_id or len(entity_id) > 64:
        return _bad()

    date_str = _valid_date(request.query.get("date"), m)
    target = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    events = await m.fetch_week_events(
        entity_id, target, event_type=m.CODE_TO_EVENT_TYPE[code]
    )
    day = [e for e in events if str(e.get("date", "")) == date_str]
    day.sort(key=lambda e: str(e.get("start_time", "")))

    lessons = []
    for e in day:
        lessons.append({
            "para": m.START_TIME_TO_PARA.get(str(e.get("start_time", "")), "?"),
            "start": str(e.get("start_time", "")),
            "end": str(e.get("end_time", "")),
            "title": e.get("title") or "Nomsiz",
            "topic": e.get("lesson_name") or "",
            "instructor": e.get("instructor") or "",
            "groups": (e.get("groups") or "").replace('"', ""),
            "room": m.get_event_room(e),
            "type": e.get("type") or "",
            "format": e.get("lesson_format") or "",
        })
    return web.json_response({"date": date_str, "lessons": lessons})


async def api_rooms(request: web.Request) -> web.Response:
    m = request.app["bot_mod"]
    block = request.query.get("block", "").upper()
    if block not in m.UNIVERSITY_MAP:
        return _bad("blok topilmadi")
    floor = str(request.query.get("floor", "ALL")).upper()
    if floor != "ALL" and floor not in m.UNIVERSITY_MAP[block]:
        return _bad("qavat topilmadi")

    rooms = m.get_rooms_for_floor(block, floor)
    now = m.now_tashkent()
    para = str(request.query.get("para", "now")).lower()
    date_str = _valid_date(request.query.get("date"), m)

    if para == "now":
        date_str = now.date().isoformat()
        hm = now.strftime("%H:%M")
        start_t, label = None, None
        for pid, (st, et, disp) in m.PARALAR.items():
            if st <= hm <= et:
                start_t, label = st, disp
                break
        if not start_t:
            return web.json_response({"break": True, "time": hm, "total": len(rooms)})
    else:
        info = m.PARALAR.get(para)
        if not info:
            return _bad("para topilmadi")
        start_t, label = info[0], info[2]

    busy, invalid = await m.get_room_statuses(rooms, block, date_str, start_t)
    free = [
        {"room": m.format_room_display(block, r), "floor": m.get_room_floor(block, r)}
        for r in rooms if r not in busy and r not in invalid
    ]
    busy_list = [
        {
            "room": m.format_room_display(block, r),
            "floor": m.get_room_floor(block, r),
            "title": info.get("title", ""),
            "instructor": info.get("instructor", ""),
            "groups": info.get("groups") or "",
            "end": info.get("end_time", ""),
        }
        for r, info in busy.items()
    ]
    return web.json_response({
        "break": False,
        "label": label,
        "date": date_str,
        "free": free,
        "busy": busy_list,
        "unknown": len(invalid),
    })


async def api_room_day(request: web.Request) -> web.Response:
    m = request.app["bot_mod"]
    q = request.query.get("q", "")
    if not q or len(q) > 20:
        return _bad()
    date_str = _valid_date(request.query.get("date"), m)

    matches = m.find_rooms_by_text(q, limit=5)
    if not matches:
        return web.json_response({"found": False})

    key = "".join(ch for ch in q if ch.isalnum()).upper()
    exact = [x for x in matches if x[1] == key]
    block, room = (exact or matches)[0]

    by_para = await m.get_room_day_events(block, room, date_str)
    display = m.format_room_display(block, room)
    if by_para is None:
        return web.json_response({"found": True, "room": display, "known": False})

    paras = []
    for pid, (st, et, _disp) in m.PARALAR.items():
        e = by_para.get(pid)
        paras.append({
            "para": pid, "start": st, "end": et, "busy": bool(e),
            "title": (e.get("title") if e else "") or "",
            "instructor": (e.get("instructor") if e else "") or "",
        })
    others = [m.format_room_display(b, r) for b, r in matches if (b, r) != (block, room)]
    return web.json_response({
        "found": True, "known": True, "room": display,
        "date": date_str, "paras": paras, "others": others,
    })


# Bloklar bandligi (bosh sahifadagi "jonli" halqalar). Hisoblash og'ir bo'lgani uchun
# 60 soniya keshlanadi va bir vaqtdagi ko'p so'rov bitta hisob-kitobni kutadi.
_OV: dict = {"key": None, "ts": 0.0, "data": None}
_OV_LOCK = asyncio.Lock()


async def api_overview(request: web.Request) -> web.Response:
    m = request.app["bot_mod"]
    now = m.now_tashkent()
    hm = now.strftime("%H:%M")
    date_str = now.date().isoformat()

    start_t, label = None, None
    if now.isoweekday() != 7:
        for pid, (st, et, disp) in m.PARALAR.items():
            if st <= hm <= et:
                start_t, label = st, disp
                break

    if not start_t:
        blocks = []
        for b in m.UNIVERSITY_MAP:
            total = len(m.get_rooms_for_floor(b, "ALL"))
            blocks.append({"block": b, "total": total, "free": total, "busy": 0})
        return web.json_response({"break": True, "time": hm, "blocks": blocks})

    key = (date_str, start_t)
    async with _OV_LOCK:
        if _OV["key"] == key and time.monotonic() - _OV["ts"] < 60:
            return web.json_response(_OV["data"])

        async def one(b: str) -> dict:
            rooms = m.get_rooms_for_floor(b, "ALL")
            busy, invalid = await m.get_room_statuses(rooms, b, date_str, start_t)
            total = len(rooms) - len(invalid)
            return {"block": b, "total": total, "free": max(total - len(busy), 0), "busy": len(busy)}

        blocks = await asyncio.gather(*(one(b) for b in m.UNIVERSITY_MAP))
        data = {"break": False, "time": hm, "label": label, "blocks": list(blocks)}
        _OV.update(key=key, ts=time.monotonic(), data=data)
    return web.json_response(data)


# ---------------- Adminga murojaat ----------------
CONTACT_KINDS = {"savol": "💬 Savol", "xato": "🐞 Xatolik", "taklif": "💡 Taklif"}
CONTACT_MIN, CONTACT_MAX = 5, 1000
CONTACT_GAP, CONTACT_PER_HOUR = 30, 5          # sekund / soatiga nechta
_CONTACT_LOG: dict[int, list[float]] = {}      # user_id -> yuborilgan vaqtlar (monotonic)


async def api_contact(request: web.Request) -> web.Response:
    """Foydalanuvchi xabarini bot orqali adminga yuboradi. Foydalanuvchi ma'lumoti
    (ID, ism, username) mijozdan emas, Telegram imzosi bilan tasdiqlangan initData'dan olinadi."""
    m = request.app["bot_mod"]
    user = request["tg_user"]
    if request.content_length and request.content_length > 8192:
        return _bad("Matn juda uzun")
    try:
        body = await request.json()
    except Exception:
        return _bad("Noto'g'ri so'rov")
    if not isinstance(body, dict):
        return _bad("Noto'g'ri so'rov")

    text = str(body.get("text", "")).replace("\r", "").strip()
    while "\n\n\n" in text:
        text = text.replace("\n\n\n", "\n\n")
    kind = body.get("kind") if body.get("kind") in CONTACT_KINDS else "savol"
    if len(text) < CONTACT_MIN:
        return _bad(f"Kamida {CONTACT_MIN} ta belgi yozing")
    if len(text) > CONTACT_MAX:
        return _bad(f"Matn {CONTACT_MAX} belgidan oshmasin")

    uid = int(user["id"])
    now = time.monotonic()
    log = [t for t in _CONTACT_LOG.get(uid, []) if now - t < 3600]
    if log and now - log[-1] < CONTACT_GAP:
        wait = int(CONTACT_GAP - (now - log[-1])) + 1
        return web.json_response({"error": f"Biroz kuting — {wait} soniyadan keyin yuboring", "retry": wait}, status=429)
    if len(log) >= CONTACT_PER_HOUR:
        return web.json_response({"error": "Soatiga 5 tagacha murojaat yuborish mumkin. Keyinroq urinib ko'ring."}, status=429)

    bot = getattr(m, "BOT_INSTANCE", None)
    if bot is None:
        return web.json_response({"error": "Bot hali tayyor emas, birozdan keyin urinib ko'ring"}, status=503)

    name = " ".join(x for x in (user.get("first_name"), user.get("last_name")) if x) or "Foydalanuvchi"
    uname = user.get("username")
    msg = (
        f"📩 <b>Ilovadan murojaat</b> · {CONTACT_KINDS[kind]}\n\n"
        f"👤 <a href=\"tg://user?id={uid}\">{m.esc(name)}</a>"
        + (f" · @{m.esc(uname)}" if uname else "")
        + f"\n🆔 <code>{uid}</code>\n\n{m.esc(text)}\n\n"
        f"↩️ Javob berish: <code>/javob {uid} </code>matn"
    )
    try:
        await bot.send_message(chat_id=m.ADMIN_ID, text=msg, parse_mode="HTML")
    except Exception as e:
        print(f"[WEBAPP] Adminga murojaatni yuborib bo'lmadi: {e}")
        return web.json_response({"error": "Xabarni yuborib bo'lmadi. Keyinroq urinib ko'ring."}, status=502)

    log.append(now)
    _CONTACT_LOG[uid] = log
    if len(_CONTACT_LOG) > 5000:                # xotira cheksiz o'smasligi uchun
        for k in [k for k, v in _CONTACT_LOG.items() if not v or now - v[-1] > 3600]:
            _CONTACT_LOG.pop(k, None)
    return web.json_response({"ok": True})


def register_webapp(app: web.Application, bot_module) -> None:
    """Mini App sahifasi va API yo'llarini mavjud aiohttp serverga qo'shadi."""
    app["bot_mod"] = bot_module
    app.middlewares.append(auth_middleware)
    app.router.add_get("/app", page)
    app.router.add_get("/api/structure", api_structure)
    app.router.add_get("/api/search", api_search)
    app.router.add_get("/api/schedule", api_schedule)
    app.router.add_get("/api/rooms", api_rooms)
    app.router.add_get("/api/room", api_room_day)
    app.router.add_get("/api/overview", api_overview)
    app.router.add_post("/api/contact", api_contact)
