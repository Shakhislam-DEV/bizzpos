# 🏪 МАГАЗИН ЕСАБЫ — ОРНАТЫЎ НУСҚАЎЛАМАСЫ

## 📋 ЖӘМИ ФАЙЛЛАР ДҮЗИМИ

```
magazin-esaby/
├── index.html
├── package.json
├── vite.config.js
├── .env.local          ← сизниң мағлыўматларыңыз
├── public/
│   └── manifest.json
└── src/
    ├── main.jsx
    └── App.jsx         ← баслапқы программа
```

---

## 🔧 1-ҚӘДЕМ: Supabase орнатыў

1. https://supabase.com → кіриң
2. "New Project" басыңыз → атын, паролін қойыңыз
3. Проект ашылғаннан кейин:
   - Сол жақтан → **SQL Editor** → "New Query"
   - `supabase_schema.sql` файлындағы ҳәмме кодты копиялаң → **Run** басыңыз Error: Failed to run sql query: ERROR: 42P07: relation "profiles" already exists
4. **Settings → API** бетине кіриңиз:
   - `Project URL` → `.env.local`-ға жазыңыз
   - `anon public key` → `.env.local`-ға жазыңыз

---

## 🤖 2-ҚӘДЕМ: Telegram Bot жасаў

1. Telegram-да **@BotFather**-ге жазыңыз
2. `/newbot` → бот атын жазыңыз (мысалы: MagazinEsabyBot)
3. **Token** беріледі → `.env.local`-ға жазыңыз
4. Бот менен чат ашыңыз ямаса группаға қосыңыз
5. Chat ID алыў ушын: https://api.telegram.org/bot**TOKEN**/getUpdates
   - `"chat":{"id": -100XXXXXXXXX}` — усы Chat ID
6. Chat ID → `.env.local`-ға жазыңыз

---

## 💻 3-ҚӘДЕМ: GitHub-қа жүклеў

1. https://github.com → кіриңиз
2. "New repository" → атын жазыңыз (мысалы: `magazin-esaby`)
3. Компьютерде **VS Code** ашыңыз (жоқ болса орнатыңыз)
4. Terminal ашыңыз (Ctrl+`) ҳәм мына командаларды орындаңыз:

```bash
# Node.js орнатылғанын тексериң (жоқ болса nodejs.org-дан орнатыңыз)
node --version

# Папка жасаңыз
mkdir magazin-esaby
cd magazin-esaby

# Файлларды жасаңыз (мен жасаған файлларды копиялаңыз)
# App.jsx → src/App.jsx ишине
# package.json, vite.config.js, index.html → тамырға

# .env.local файл жасаңыз
cp .env.example .env.local
# Ишине Supabase ҳәм Telegram мағлыўматларыңызды жазыңыз

# Пакетлерди орнатыңыз
npm install

# Жергиликли тексериңиз
npm run dev
# http://localhost:5173 — браузерде ашыңыз
```

---

## 🚀 4-ҚӘДЕМ: Vercel-ге шығарыў

1. https://vercel.com → GitHub аккаунт менен кіриңиз
2. "Add New Project" → GitHub репозиторийиңизди таңлаңыз
3. **Environment Variables** бөлимине кіриңиз:
   - `VITE_SUPABASE_URL` → мәнін жазыңыз
   - `VITE_SUPABASE_ANON_KEY` → мәнін жазыңыз
   - `VITE_TELEGRAM_BOT_TOKEN` → мәнін жазыңыз
   - `VITE_TELEGRAM_CHAT_ID` → мәнін жазыңыз
   - `VITE_SHOP_NAME` → магазин атыңыз
4. **Deploy** басыңыз!
5. 2-3 минуттан кейин адресиңиз tayyor: `https://magazin-esaby.vercel.app`

---

## 👤 5-ҚӘДЕМ: Директор аккаунтын жасаў

1. Supabase → **Authentication → Users → Invite User**
2. Директордың email-ін жазыңыз
3. Email келеді → парол қойады
4. Суперbase → **SQL Editor** → мына кодты орындаңыз:

```sql
-- Директор ролін қолдан бериңиз (email-ни өзгертиңиз)
UPDATE profiles
SET role = 'director'
WHERE id = (SELECT id FROM auth.users WHERE email = 'director@email.com');
```

5. Директор кіргеннен кейин → **Параметр** бөлиминен сатыўшы ҳәм снабженец қосады

---

## 📱 6-ҚӘДЕМ: Телефонға орнатыў (PWA)

**Android/Chrome:**
1. Сайтты ашыңыз
2. Chrome мәзири (3 нүкте) → **"Экранға қосыў"**
3. Орнатыңыз → қосымша сыяқлы ислейди ✅

**iPhone/Safari:**
1. Сайтты ашыңыз
2. Бөлисиў белгиси → **"Баслапқы экранға қосыў"**
3. Орнатыңыз ✅

---

## 🖨️ Xprinter XP-58IIH ОРНАТЫЎ

1. USB менен компьютерге қосыңыз
2. Драйверді жүклеңиз: https://www.xprinter.net/download
3. Windows → Принтер орнатыңыз
4. Браузерде чек басқанда → принтериңизди таңлаңыз
5. Қағаз өлшемін **58мм** деп қойыңыз

---

## 🔑 РОЛЛЕР

| Рол | Кіре алатын бөлимлер |
|-----|---------------------|
| Сатыўшы | Басты бет, Сатыў, Сораныс |
| Снабженец | Басты бет, Кириш, Товарлар, Сораныс, Есабат |
| Директор | Ҳәммеси |

---

## ❓ КӨП УШЫРАСАТУҒЫН МӘСЕЛЕЛЕР

**Камера ислемейди:**
- Тек HTTPS-та ислейди (Vercel-де автоматик)
- Chrome браузеринде жақсы ислейди

**Суперbase қосылмайды:**
- `.env.local` дурыс жазылғанын тексериңиз
- Vercel-де Environment Variables қойылғанын тексериңиз

**Telegram хабар келмейди:**
- Bot Token ҳәм Chat ID дурыслығын тексериңиз
- Ботты чатқа admin қылып қосыңыз





supabase.com
BizzPOS
ZQfh+Px%@LK.4qC



VITE_SUPABASE_URL=https://vctmkekpnfslscugengd.supabase.co
VITE_SUPABASE_ANON_KEY=sb_publishable_b2GJKK7rWEp-sMsdnvy2Qw_fiToR9Lj
VITE_TELEGRAM_BOT_TOKEN=мунда_bot_token_жазыңыз
VITE_TELEGRAM_CHAT_ID=мунда_chat_id_жазыңыз
VITE_SHOP_NAME=Магазин атыңыз