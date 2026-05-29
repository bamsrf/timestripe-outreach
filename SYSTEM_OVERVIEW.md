# TimeStripe Outreach Automation — техническое описание системы

> Документ для команды Timestripe. Описывает архитектуру, бизнес-логику и
> текущее состояние системы автоматизированного аутрича к микро-инфлюенсерам.

---

## 1. Executive Summary

**Что:** автоматизированный pipeline для поиска, оценки и контакта с
микро-инфлюенсерами (YouTube → TikTok → Instagram) под программу
сотрудничества Timestripe.

**Зачем:** вручную перебирать TikTok/YouTube под нужные тиры — это десятки
часов рутины в неделю. Система делает ту же работу за 5-10 минут (на пайплайн
+ один клик "send batch") и масштабируется до тысяч контактов.

**Стек:** Python + YouTube Data API + Yandex 360 SMTP/IMAP + Google Sheets API
+ Streamlit (веб-админка). Локально, single-machine, бесплатно (исключения см.
раздел "Стоимость").

**Состояние на сегодня:**
- YouTube парсер ✅ работает, 119 каналов в базе
- Outreach по email ✅ работает, отправлены тестовые письма
- Google Sheets интеграция ✅ готова, дедуп против ручной CRM
- Авто-детект ответов через IMAP ✅ работает (cron в 00:00)
- Веб-админка ✅ работает (KPI, funnel, таблица контактов, очередь отправок)
- TikTok парсер ⏳ следующий этап
- Instagram парсер ⏳ за TikTok'ом

---

## 2. Бизнес-логика

### Целевая аудитория — 4 тира

| Tier | Описание | Subs | Доля бюджета внимания |
|---|---|---|---|
| A — Life Planner | plan-with-me, vision boards, study vlogs, утренние рутины | 500–10K | 40% |
| B — Self-Improvement | habits, discipline, journaling, self-growth | 500–10K | 30% |
| C — Creative Professional | freelance designers, writers, illustrators | 500–10K | 20% |
| D — Productivity Tech | обзорщики приложений, productivity setups | 10K–300K | 10% |

Критерий "хороший лид" — не просто канал в нише, а **визуально пользуется
планировщиком** (показывает реальную продуктивную рутину). Парсер
оценивает каждый канал по fit-score (см. ниже).

### Воронка outreach

```
Найдено парсером → С email → Прошло fit-фильтр → Welcome sent →
→ Ответили → Offer sent → Договорились → Запостили контент
```

Текущий шаблон писем — две стадии:
- **Welcome** (cold intro) — рассказывает про Timestripe, спрашивает
  заинтересован ли блогер
- **Offer** (после ответа) — детали программы: Revenue Share 20% / Views-Based
  Payment $50–500 за 100K–1M просмотров

---

## 3. Архитектура

```
┌────────────────────────────────────────────────────────────────┐
│  YouTube Data API (бесплатно, 10K квоты/день)                  │
└────────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌────────────────────────────────────────────────────────────────┐
│  youtube_finder.py                                             │
│  • 88 ключевиков × 4 тира                                      │
│  • Дедуп между запусками через seen_channels.csv               │
│  • Fit-scoring (productivity terms / spam ratio / recency)     │
│  • Фильтр кириллицы (не русская аудитория)                     │
│  • Экстракция email + соцссылок из описаний                    │
└────────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌────────────────────────────────────────────────────────────────┐
│  MASTER_youtube_outreach.xlsx (локальный CRM)                  │
│  17 колонок: Name, Email, Platform, URL, Followers, Tier,      │
│  Niche, Status, First Contact Date, Last Follow-up, Notes...   │
└────────────────────────────────────────────────────────────────┘
                ▲             │            │
                │             │            │
                │             ▼            ▼
       ┌────────────┐  ┌────────────┐  ┌────────────────────┐
       │ sheets_    │  │ outreach_  │  │ admin.py           │
       │ sync.py    │  │ sender.py  │  │ (Streamlit web UI) │
       └────────────┘  └────────────┘  └────────────────────┘
              │             │                  │
              ▼             ▼                  │
   ┌────────────────┐ ┌────────────────┐       │
   │ Google Sheet   │ │ Yandex SMTP    │       │
   │ "Vlad" tab     │ │ (vlad@         │       │
   │ (видна команде)│ │  timestripe.   │       │
   │                │ │  com)          │       │
   └────────────────┘ └────────────────┘       │
                            │                  │
                            ▼                  │
                  ┌─────────────────────┐      │
                  │ Inbox + Bloggers    │◀─────┘
                  │ folder (IMAP)       │
                  └─────────────────────┘
                            │
                            ▼
                  ┌─────────────────────┐
                  │ reply_detector.py   │
                  │ (cron в 00:00)      │
                  │ Status="Replied"    │
                  │ если пришёл ответ   │
                  └─────────────────────┘
```

---

## 4. Компоненты — детально

### 4.1 `youtube_finder.py` — поиск каналов

- **Источник данных:** YouTube Data API v3 (официальный, бесплатный)
- **Логика поиска:**
  1. Для каждого тира — список ключевиков (`config.yaml`)
  2. Для каждого ключевика — `search.list` → 50 свежих видео → каналы авторов
  3. Для каждого канала — `channels.list` → подписчики, видео, описание
- **Фильтры (в порядке применения):**
  - Подписчики в диапазоне тира
  - ≥3 видео опубликовано
  - Имя без кириллицы (не русская аудитория)
  - Fit-score ≥ -2
  - Есть email в описании канала → в основную выгрузку
  - Нет email → в `youtube_skipped_no_email_*.csv` (для ручной доразведки)
- **Fit-score (от -17 до +14):**
  - `+`: слова productivity/planning/journal/goals/etc. в описании
  - `+`: недавний upload (<30 дней = +3, <90 = +1)
  - `−`: spammy videos/subs ratio (>0.5 = -6)
  - `−`: мемориальный канал (RIP/tribute/in heaven) = -10
  - `−`: эзотерика/новости/астрология = -4
- **Дедуп между запусками:** `seen_channels.csv` хранит channel_id всех уже
  виденных каналов; следующий запуск в delta-режиме исключает их

### 4.2 `outreach_sender.py` — отправщик писем

- **SMTP:** smtp.yandex.ru:465 (SSL) с App Password
- **Защита от случайной отправки:** по умолчанию dry-run; нужен явный `--send`
- **Дедуп:** пропускает контакты со статусом "Sent intro" / "Replied" / etc.
- **Дедуп по Google Sheet:** читает `crm_emails.csv` (синхронизированный из
  таблицы вручную ведённых контактов) → не пишет тем, кого ты уже добавил
  руками
- **Rate-limit:** 30–60 секунд между отправками (Yandex лимит 150/час)
- **Ротация subject:** 4 варианта тем рандомно, чтобы не выглядеть как
  массовая рассылка
- **Smart name extraction:** имя вытягивается из канала ("LifeCraft with Sree"
  → "Sree") или email ("queennjasmine@..." → "Jasmine")
- **Шаблоны:** `outreach_config.yaml`
  - Welcome: один body + 4 тир-специфичных intro
  - Offer: текст программы с условиями
- **Дублирование в Bloggers:** после успешной отправки скрипт через IMAP
  кладёт копию письма в папку `Bloggers` в почтовом ящике (чтобы все аутрич-
  письма были в одном месте)

### 4.3 `reply_detector.py` — авто-детект ответов

- **Запуск:** автоматически в 00:00 каждый день через launchd, или вручную
  из админки кнопкой "📥 Sync replies"
- **Что делает:**
  1. INBOX scan за последние 2 дня → email-адреса отправителей
  2. Sent folder scan → email-адреса получателей (для случая когда ты
     написал блогеру вручную из Яндекса)
  3. Сверяет с MASTER. Если совпадение — обновляет Status:
     - Был "Sent intro" + новый reply → "Replied"
     - Был пустой + найдено в Sent → "Sent intro" + First Contact Date

### 4.4 `sheets_sync.py` — двусторонняя синхронизация с Google Sheet

- **Push mode (по умолчанию):** новые контакты из MASTER → Google Sheet
  (Vlad tab, append после последней заполненной строки)
- **Pull mode:** существующие emails из Google Sheet → `crm_emails.csv`
  (для дедупа в outreach_sender)
- **Не трогает существующие строки** — только добавляет новые в конец
- **Сервис-аккаунт:** `timestripe-sheets-sync@tiemstripe-parcing.iam.gserviceaccount.com`

### 4.5 `admin.py` — Streamlit веб-админка

5 страниц:

1. **📊 Overview** — KPI tiles, funnel chart, breakdowns по тиру/статусу,
   последние отправки
2. **👥 Contacts** — таблица с фильтрами + редактируемый Status (dropdown)
   прямо в UI
3. **📬 Send Queue** — превью следующих N для отправки, ротация subject,
   кнопка "🚀 Send batch" с подтверждением
4. **📝 Templates** — все шаблоны welcome/offer, subject-вариации, intro по
   тирам (read-only, редактирование через YAML)
5. **📜 Send Log** — полный лог отправок с фильтрами и экспортом

Запуск: `streamlit run admin.py` → http://localhost:8501

---

## 5. Pipeline / Workflow

### Еженедельный цикл

| День | Действие | Команда / место |
|---|---|---|
| Пн | Запустить YouTube парсер (delta-режим) | админка → кнопка "🔍 Run parser" |
| Пн | Push новых контактов в Google Sheet | `python sheets_sync.py` |
| Пн | Просмотреть Send Queue в админке | http://localhost:8501 |
| Пн–Пт | Отправлять welcome пачками по 20–40 в день | админка → Send Queue → Send batch |
| Каждый день | Reply sync (автоматом в 00:00) | launchd |
| По мере появления | Offer-рассылка тем кто ответил | админка → Send Queue → Offer stage |
| По мере появления | Заполнять Payout / Program Type вручную в Google Sheet | руками |

### Что происходит при одной отправке

1. Sender читает MASTER + crm_emails.csv (исключений)
2. Фильтрует по fit_score ≥ 3, по Status, по тиру
3. Берёт первого по приоритету (Tier A → fit_score desc → followers desc)
4. Smart-extract имя из канала/email
5. Подставляет {Name}, {Platform}, intro_tier в шаблон
6. Ротирует subject (4 варианта по кругу)
7. SMTP отправка через Yandex 360
8. IMAP append копии в папку Bloggers
9. Update MASTER: Status="Sent intro", First Contact Date=today,
   Last Follow-up=today
10. Append в outreach_log.csv
11. Sleep 30-60 секунд → следующий

---

## 6. Метрики и KPI

Все доступны на странице 📊 Overview в админке.

| Метрика | Откуда |
|---|---|
| Total contacts | MASTER, кол-во строк |
| Pending | Status пустой или "Not contacted" |
| Welcome sent | Status="Sent intro" |
| Replied | Status="Replied" (заполняется автоматически из IMAP) |
| Offer sent | Status="Sent program info" |
| Negotiating | Status="Wants money for posts" |
| Won (live collab) | Status="Live collab" (заполняется вручную) |
| Sent today | Из outreach_log.csv, фильтр по дате |

### Целевые конверсии для cold outreach (бенчмарк индустрии)

- **Welcome → Replied:** 8-15% (реалистично 5-10% на первой итерации, до 20%
  при отточенных шаблонах)
- **Replied → Offer sent → Live collab:** 30-50%
- **Итого Welcome → Live collab:** 3-7%

При 100 welcome в неделю это ~3-7 коллабораций в неделю → ~12-28 в месяц.

---

## 7. Стоимость

| Сервис | Цена | Лимит | Что делаем при достижении |
|---|---|---|---|
| YouTube Data API | $0 | 10K квоты/день (~80-100 ключевиков) | Сбрасывается в 00:00 UTC |
| Yandex 360 (уже есть) | (корп подписка) | 3000-5000 писем/день | С запасом |
| Google Sheets API | $0 | 60 запросов/минуту | С запасом |
| Hardware | $0 | один Mac | Можно мигрировать на VPS за $5/мес |

**Итого: $0 в месяц** (пока работаем только с YouTube + email).

### Что добавит стоимость дальше

- **TikTok парсер через Apify**: free tier $5/мес кредитов хватит на ~10K
  профилей. Дальше — $49/мес стартовый тариф.
- **Smartlead/Instantly для warmup и follow-up**: $37-97/мес. Не обязательно
  если корп Яндекс хватает.
- **HikerAPI для Instagram**: $50-200/мес. Только если бесплатные пути
  (Apify Instagram Scraper) не справляются.

**Реалистичный максимум при полной автоматизации YouTube + TikTok + IG:**
≈ $100-200/мес. При 28 коллабах/мес стоимость лида = $4-7. Сравни со среднерыночными платными discovery-инструментами Modash / Heepsy ($200-600/мес только за поиск, без отправки и автоматизации).

---

## 8. Технологический стек

| Язык / технология | Где используется |
|---|---|
| Python 3.12 | весь backend |
| Streamlit | веб-админка |
| pandas + openpyxl | работа с XLSX |
| Plotly | графики и funnel |
| gspread | Google Sheets API |
| google-api-python-client | YouTube Data API |
| python-dotenv | secrets management |
| smtplib + imaplib (stdlib) | Yandex SMTP/IMAP |
| launchd (macOS) | cron для авто-задач |

**Все секреты** (YouTube API key, Yandex App Password, Google Service Account
key) в `.env` файле и `~/.config/timestripe-gsheets.json`. В git не
коммитятся (`.gitignore`).

---

## 9. Текущий статус и метрики

(на момент составления документа)

- **Парсер YouTube:** 119 каналов в MASTER, 87 с fit_score ≥+3
- **Шаблоны:** Welcome + Offer полностью прописаны
- **Тестовые отправки:** 4 письма прошли через всю систему (на тестовый
  адрес), все 4 успешно
- **Google Sheet интеграция:** работает, готово к синку 119 контактов
- **Авто-детект ответов:** настроен, ждёт первый реальный реплай
- **TikTok / Instagram:** запланированы

---

## 10. Дорожная карта

### Phase 2 (следующий месяц): TikTok parser
- Apify Free Tier → собирать creators по хэштегам
- Извлечение email из bio + резолв Linktree/Beacons/Stan.store
- Те же fit-фильтры
- TikTok даст ~60-70% объёма (в текущей ручной CRM 16/17 = 94% TikTok)

### Phase 3: Instagram parser
- Apify Instagram Hashtag Scraper или HikerAPI
- Извлечение `business_email` из публичных Business Profiles
- Самый сложный из-за anti-bot Meta

### Phase 4: Smart automation
- Follow-up sequences (если не ответили через 5 дней → шлём напоминание)
- A/B тестирование шаблонов (3 варианта welcome, статистика open/reply rate)
- Auto-pause при низкой deliverability (отслеживаем bounce-rate)

### Phase 5: Reporting
- Еженедельный отчёт в Slack/Telegram (новые контакты, конверсии, ROI)
- Дашборд с историческими данными (по неделям/месяцам)

---

## 11. Безопасность

- Все секреты в `.env` / `~/.config/` — не в git
- Yandex App Password можно отозвать в id.yandex.ru → Безопасность одной
  кнопкой
- YouTube API ключ ограничен только YouTube Data API (если утечёт — нельзя
  использовать ни для чего другого в Google Cloud)
- Google Service Account имеет доступ ТОЛЬКО к расшаренным с ним таблицам
  (Editor role, без других прав)
- Skipped CSV не содержит секретов — можно безопасно отдавать команде

---

## 12. Что НЕ автоматизировано

Эти шаги остаются ручными (намеренно — требуют человеческого решения):

- Финальная проверка списка перед отправкой
- Заполнение Status = "Live collab" после реальной коллаборации
- Payout Owed / Payout Sent — финансовые поля
- Program Type — какую программу предложили (Revenue Share / Views-Based)
- Редактирование шаблонов под фидбэк (через `outreach_config.yaml`)
- Расширение ключевиков под новые ниши (через `config.yaml`)
- Реакция на ответы — система пометит "Replied", но ОТВЕЧАТЬ нужно вручную

---

## Контакты

Поддерживает: Vlad Rumantsev — vlad@timestripe.com

Репозиторий: `~/Desktop/Cursor/Timestripe/`
