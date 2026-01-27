# Brandshop Scanner Telegram Bot

Telegram бот для переобліку товарів з інтеграцією Google Sheets.

## Функціонал

- Перевірка всієї категорії
- Перевірка по артикулах з файлу
- Перевірка одного артикулу
- Генерація Excel файлів з результатами перевірки

## Встановлення

1. Клонуйте репозиторій:
```bash
git clone https://github.com/romchhh/brandshop-scanner-tg-bot.git
cd brandshop-scanner-tg-bot
```

2. Встановіть залежності:
```bash
pip install -r requirements.txt
```

3. Створіть файл `.env` на основі `.env.example`:
```env
BOT_TOKEN=your_bot_token_here
GOOGLE_CREDENTIALS={"type":"service_account",...}
```

## Локальний запуск

```bash
python main.py
```

## Розгортання на Railway

### Крок 1: Підготовка GitHub репозиторію

1. Створіть новий репозиторій на GitHub (якщо ще не створено)
2. Ініціалізуйте git (якщо ще не зроблено):
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/romchhh/brandshop-scanner-tg-bot.git
git push -u origin main
```

### Крок 2: Налаштування Railway

1. Зайдіть на [railway.app](https://railway.app) та увійдіть через GitHub
2. Натисніть **"New Project"** → **"Deploy from GitHub repo"**
3. Виберіть репозиторій `romchhh/brandshop-scanner-tg-bot`
4. Railway автоматично визначить проект як Python

### Крок 3: Додавання змінних середовища

У налаштуваннях проекту Railway додайте наступні змінні:

1. **BOT_TOKEN** - токен вашого Telegram бота (отримайте у [@BotFather](https://t.me/BotFather))
2. **GOOGLE_CREDENTIALS** - весь JSON з credentials як один рядок (без переносів)

Для додавання `GOOGLE_CREDENTIALS`:
- Відкрийте ваш `credentials.json` файл
- Скопіюйте весь вміст
- Вставте в змінну `GOOGLE_CREDENTIALS` в Railway (як один рядок, без переносів)

### Крок 4: Деплой

Railway автоматично почне деплой після пушу в GitHub. Перевірте логи в Railway dashboard, щоб переконатися, що бот запустився успішно.

## Структура проекту

```
.
├── main.py                 # Головний файл бота
├── config.py              # Конфігурація та налаштування
├── handlers/              # Обробники повідомлень
│   └── inventory_handlers.py
├── keyboards/            # Клавіатури
│   └── inventory_keyboards.py
├── states/               # FSM стани
│   └── inventory_states.py
├── utils/                # Утиліти
│   ├── sheets_utils.py
│   ├── excel_generator.py
│   ├── admin_utils.py
│   └── category_translations.py
├── requirements.txt      # Залежності Python
├── Procfile             # Конфігурація для Railway
└── .env                 # Змінні середовища (не комітиться)
```

## Важливо

- Файл `.env` не комітиться в git (додано в `.gitignore`)
- `credentials.json` також не комітиться
- На Railway всі credentials передаються через змінні середовища

## Підтримка

Якщо виникли проблеми з деплоєм, перевірте:
1. Логи в Railway dashboard
2. Чи правильно додані всі змінні середовища
3. Чи правильно відформатований `GOOGLE_CREDENTIALS` (має бути один рядок JSON)
