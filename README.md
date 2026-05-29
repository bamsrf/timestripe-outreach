# TimeStripe Outreach — поиск инфлюенсеров

YouTube парсер для поиска каналов по тирам (A/B/C/D) для аутрича TimeStripe.
Использует официальный YouTube Data API v3 — **бесплатно**, 10 000 запросов в день.

## Тиры

| Tier | Тематика | Доля | Subs |
|---|---|---|---|
| A | Dream Life Planner — лайфстайл, plan with me, vision boards | 40% | 500–10K |
| B | Self-Improvement — привычки, дисциплина, journaling | 30% | 500–10K |
| C | Creative Professional — фрилансеры, дизайнеры, writers | 20% | 500–10K |
| D | Productivity Tech — обзорщики приложений | 10% | 10K–300K |

Тиры и ключевики настраиваются в [`config.yaml`](config.yaml).

## Быстрый старт

### 1. Получи YouTube API key (5 минут, бесплатно)
1. https://console.cloud.google.com/ → создай новый проект (любое имя)
2. APIs & Services → Library → "YouTube Data API v3" → **Enable**
3. APIs & Services → Credentials → **Create Credentials → API key**
4. Скопируй ключ

### 2. Установи зависимости
```bash
cd /Users/vladislavrumancev/Desktop/Cursor/Timestripe
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# вставь свой ключ в .env
```

### 3. Запусти
```bash
# все 4 тира
python youtube_finder.py

# один конкретный тир
python youtube_finder.py --tier A_dream_life_planner

# только русскоязычные каналы
python youtube_finder.py --lang ru --region RU

# свежие каналы (постят за последние 30 дней)
python youtube_finder.py --days 30
```

Результат — два CSV в `output/`:
- **`youtube_YYYYMMDD_HHMM.csv`** — основной, готов к импорту в твою таблицу. Только каналы с публичным email. Колонки 1-в-1 как в Google Sheet:
  `Name, Email, Platform, URL, Followers, Tier, Niche, Status, First Contact Date, Last Follow-up, Notes, Program Type, Videos Published, Total Views, Payout Owed, Payout Sent, ROI`
- **`youtube_skipped_no_email_YYYYMMDD_HHMM.csv`** — каналы без email (для ручной доразведки: на канал → About → "View email address").

В `Notes` парсер докладывает: alt-email'ы, ссылки IG/TT/TG/X/Linktree, топ-3 ключевика по которым нашли.

## Импорт в Google Sheets

В твоей CRM (`Vlad` вкладка) сделай `File → Import → Upload → CSV → Append to current sheet` или открой CSV в новом листе и копируй колонки в нужные.

## Лимиты YouTube API

- 10 000 единиц в день бесплатно
- `search.list` = 100 единиц за запрос (50 каналов)
- `channels.list` = 1 единица за запрос (до 50 каналов)
- Грубо: ~80 ключевиков в день. Конфиг сейчас ~70 ключей суммарно → влезает в одну сессию.

Если квота кончилась — жди следующего дня (UTC midnight) или сделай второй проект в Google Cloud (бесплатно).

## Дальше — TikTok / Instagram

Официального бесплатного API нет. Варианты:

1. **TikTok Creator Marketplace** (creatormarketplace.tiktok.com) — официально, бесплатно, но нужен бизнес-аккаунт TikTok Ads. Фильтры по нише, гео, кол-ву подписчиков. **Рекомендую как первый шаг для TikTok.**
2. **Apify** — free tier ~$5 кредитов в месяц, есть готовые scraper'ы под TikTok / Instagram. Можно подключить и собрать профили по хэштегам типа `#planwithme`, `#productivity`.
3. **Modash / Heepsy / Collabstr** — free trial. Есть discovery с фильтрами.
4. Скажи мне — допишу TikTok модуль через Apify API (нужен будет токен Apify).
