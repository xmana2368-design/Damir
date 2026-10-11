UWED SPACE — TELEGRAM MINI APP
================================

PAPKA ICHIDA:
- bot.py        — bot.py ning Mini App bilan integratsiya qilingan nusxasi
- webapp.py     — Mini App API va Telegram initData tekshiruvi
- webapp.html   — mobil ilovaga o'xshash UI
- requirements.txt

MUHIM:
1. Hozirgi bot.py faylingizning zaxira nusxasini oling.
2. Ushbu papkadagi bot.py ni loyihangizdagi bot.py o'rniga qo'ying. Bu nusxa siz yuborgan
   bot.py asosida tayyorlangan; mavjud handlerlar va asosiy funksiyalar saqlangan.
3. webapp.py va webapp.html fayllarini bot.py bilan BIR PAPKAGA joylang.
4. requirements.txt dagi paketlar Render build uchun ishlatiladi. Agar loyihangizda qo'shimcha
   kutubxonalar bo'lsa, eski requirements.txt dagilarni ham saqlang.
5. Render → Environment bo'limida quyidagilar bo'lsin:
   BOT_TOKEN = Telegram bot tokeni
   WEBAPP_URL = https://SIZNING-RENDER-SERVICE-URL
   (Masalan: https://jidu-bot-xxxx.onrender.com — oxirida /app yozmang.)
   ADMIN_ID va bot ishlatadigan boshqa environment o'zgaruvchilarini ham avvalgidek saqlang.
6. Render'da redeploy qiling.
7. Telegram'da botga /start yuboring va “✨ UWED SPACE — Ilovani ochish” tugmasini bosing.

WEBAPP_URL:
- URL albatta HTTPS bo'lishi kerak.
- WEBAPP_URL ni Render service'ning aniq public URL manzili bilan to'ldiring.
- RENDER_EXTERNAL_URL avtomatik berilgan bo'lsa, kod undan ham foydalanadi.

XAVFSIZLIK:
- BOT_TOKEN ni hech kimga yubormang va GitHub repository ichiga yozmang.
- Mini App API so'rovlari Telegram initData HMAC bilan tekshiriladi.
- Web App'ni Telegram ichidagi bot tugmasidan oching; oddiy brauzerda API 401 qaytarishi normal.

ESLATMA:
- Mini App backend sizning bot.py ichidagi search_entities, fetch_week_events,
  find_rooms_by_text va get_room_day_events funksiyalaridan foydalanadi.
- UWED API guruh qidiruvi yoki ayrim maydonlar javob bermasa, bu universitet API javobiga
  bog'liq bo'lishi mumkin. Bunday vaziyatda Render Logs'dagi xabarni tekshiring.
- Xona holati jadvalda mavjud darslar asosida ko'rsatiladi; API'da yo'q ma'lumotni
  ilova o'zi taxmin qilib to'ldirmaydi.

YANGI QO'SHILGANLAR:
- Mini App ichida "Bloklar" bo'limi: A/B/D/F blokini, qavatni va sanani tanlang.
- 1–7-paralar bo'yicha bo'sh/band auditoriyalar soni hamda xona ro'yxati ko'rsatiladi.
- Telegram chatining pastki reply keyboard'ida "📱 UWED SPACE" va "🤖 Bot menyusi" tugmalari chiqadi.
- "Bot menyusi" tugmasi eski inline bot xizmatlarini qayta ochadi.

ESLATMA:
- Blok holatini tekshirishda tanlangan qavat uchun barcha 7 para so'rovi bajariladi. Birinchi tekshiruvda UWED API sabab biroz vaqt olishi mumkin.
- Bo'sh xona degani UWED jadvalida shu para uchun dars topilmaganini anglatadi; bu jonli xona sensori emas.
