# TimeStripe Outreach — план парсинга микроинфлюенсеров

> 🎯 **Цель:** автоматизированно наполнять CRM-таблицу ([Google Sheets, вкладка Vlad](https://docs.google.com/spreadsheets/d/1jjr-Au1nqOZGpUX8FMQdUlHp35myWXjLM_AwLdZk_tw/edit?gid=244477095)) валидными контактами инфлюенсеров под 4 тира.
>
> 🚫 **Жёсткое правило:** строка БЕЗ email и БЕЗ ссылки на аккаунт — в таблицу не попадает. Это не "nice to have", это hard filter на этапе записи.

---

## Целевая структура CRM (17 колонок)

| # | Колонка | Источник | Обязательная? |
|---|---|---|---|
| 1 | Name | парсер | ✅ |
| 2 | **Email** | парсер (из описания канала/био) | ✅ **БЛОКЕР** |
| 3 | Platform | парсер (YouTube/TikTok/Instagram) | ✅ |
| 4 | **URL** | парсер (ссылка на аккаунт) | ✅ **БЛОКЕР** |
| 5 | Followers | парсер | ✅ |
| 6 | Tier | парсер (по конфигу) | ✅ |
| 7 | Niche | парсер (ключевик, по которому нашли) | ✅ |
| 8 | Status | вручную (Sent intro / Replied / Wants money / Dead) | ❌ |
| 9 | First Contact Date | вручную | ❌ |
| 10 | Last Follow-up | вручную | ❌ |
| 11 | Notes | вручную + парсер докидывает доп.контакты (IG/TG/TT) | ⚠️ |
| 12 | Program Type | вручную | ❌ |
| 13 | Videos Published | парсер (videoCount) | ✅ |
| 14 | Total Views | парсер (viewCount) | ✅ |
| 15 | Payout Owed | вручную | ❌ |
| 16 | Payout Sent | вручную | ❌ |
| 17 | ROI | формула | ❌ |

Парсер закрывает: **1, 2, 3, 4, 5, 6, 7, 11 (доп.соцсети), 13, 14**. Остальное — outreach-pipeline, заполняешь руками по мере работы.

---

## Тиры (фиксируем здесь, чтобы не путаться)

| Tier | Niche | Доля | Subs | Платформы |
|---|---|---|---|---|
| A | Dream Life Planner — plan with me, vision boards, утренние рутины | 40% | 500–10K | TikTok / Instagram / YouTube |
| B | Self-Improvement — habits, discipline, journaling | 30% | 500–10K | TikTok / Instagram / YouTube |
| C | Creative Professional — freelancers, designers, writers | 20% | 500–10K | YouTube / Instagram |
| D | Productivity Tech — обзорщики приложений | 10% | 10K–300K | YouTube |

---

## Фазы (по приоритету)

### ✅ Фаза 0 — что уже сделано
- [x] `youtube_finder.py` — поиск по тирам через YouTube Data API v3 (бесплатно)
- [x] `config.yaml` — 4 тира × 69 ключевиков (RU + EN)
- [x] Извлечение email из описаний каналов (regex)
- [x] Извлечение Instagram / TikTok / Telegram / Twitter из описаний
- [x] CSV-выгрузка с тиром, подписчиками, видео, просмотрами

### ✅ Фаза 1 — выровнять YouTube под CRM (СДЕЛАНО)
- [x] Hard-filter: каналы без email уходят в отдельный `youtube_skipped_no_email_*.csv`, в основной CSV не попадают
- [x] Колонки 1-в-1 как в Google Sheet (17 колонок: Name, Email, Platform, URL, Followers, Tier, Niche, Status, ...)
- [x] `Platform = "YouTube"`, заполняем поля парсера, остальные пустые для ручного outreach
- [x] `Niche` — `A - Life Planner` / `B - Self-Improvement` / `C - Creative` / `D - Productivity Tech` (через `niche_label` в config.yaml)
- [x] `Notes` собирает: alt-email, IG/TT/TG/X/Linktree, топ-3 keyword'а
- [x] Dedup по email — один email = одна строка
- [x] URL формируется в виде `youtube.com/@handle` если у канала есть customUrl, иначе `/channel/UCxxx`
- [x] Linktree/Beacons/Stan.store/bio.link добавлены в regex (для Фазы 2+ это +30% к email-conversion)

### 🎯 Фаза 2 — TikTok parser (главный по объёму)
**Зачем:** Tier A и B живут в основном на TikTok. Это даст ~60% контактов.

**Бесплатный путь** (хардкор):
- [ ] Scraper через [TikTok Apify Actor](https://apify.com/clockworks/free-tiktok-scraper) (free tier ~$5/мес, хватит на ~10K профилей).
- [ ] Альтернатива: [tiktokapi-python](https://github.com/davidteather/TikTok-Api) — unofficial, периодически ломается, но 0$.
- [ ] **Самый надёжный free путь — TikTok Creator Marketplace** (creatormarketplace.tiktok.com): нужен TikTok Ads business аккаунт, дальше есть discovery с фильтрами по нише/гео/подписчикам и **email отдаётся API-через**.

**Что должен делать `tiktok_finder.py`:**
- [ ] Поиск по хэштегам из конфига (`#planwithme`, `#productivityhacks`, etc.) → собирает video creators
- [ ] Для каждого creator: имя, ссылка, followers, bio
- [ ] Из bio парсит email regex'ом + ссылки на IG/YouTube/Linktree
- [ ] **Опционально:** если в bio Linktree → fetch Linktree → ищет email там (часто email прячут именно в линктри)
- [ ] Hard-filter: нет email → в skipped, не в основной CSV
- [ ] Записывает в тот же формат CRM-колонок, `Platform = TikTok`

**Эстимейт:** 1–2 сессии разработки.

### 📸 Фаза 3 — Instagram parser
**Зачем:** Tier A/B/C активно в Instagram. Сложнее всего из-за антибота.

**Варианты:**
- [ ] Apify Instagram Hashtag Scraper / Profile Scraper — тот же free tier
- [ ] [instaloader](https://github.com/instaloader/instaloader) — Python библиотека, требует логин, легко получить ban → нужен дамми-аккаунт
- [ ] [HikerAPI / RocketAPI](https://hikerapi.com/) — платные, $50+/мес, но самые надёжные. Только если бесплатные пути не справляются.

**Что должен делать `instagram_finder.py`:**
- [ ] Поиск по хэштегам → top posts → creators
- [ ] Для каждого: bio, followers, public email (в Instagram Business profiles есть отдельное поле `business_email` — золото)
- [ ] Из bio парсит email + Linktree
- [ ] Hard-filter без email

**Эстимейт:** 2–3 сессии. Сложнее всего из-за Meta anti-scraping.

### 🧹 Фаза 4 — дедупликация и слияние
**Зачем:** один и тот же блогер может найтись и в YouTube, и в TikTok. Не хотим спамить.

- [ ] Создать `merge.py`, который читает все `output/*.csv` и собирает один мастер-файл.
- [ ] **Ключ дедупликации:** email (primary), затем имя+платформа.
- [ ] Если один email у двух платформ → одна строка, в `Platform` пишем `TikTok + YouTube`, в `Notes` обе ссылки и суммарные подписчики.
- [ ] Отдельный CSV `master_crm_ready.csv` — это финальная выгрузка для импорта в Google Sheet.

### 🔁 Фаза 5 — Google Sheets push (опционально)
**Зачем:** убрать ручной импорт CSV каждый раз.

- [ ] `sheets_push.py` через `gspread` + service account
- [ ] Читает master CSV → дописывает в твою таблицу `Append` (не перезаписывая ручные правки в Status / Notes / даты)
- [ ] **Дедуп на стороне sheets:** перед append читает существующие email из таблицы, не пушит уже существующие
- [ ] Запуск раз в неделю по расписанию (cron или scheduled task)

### 📅 Фаза 6 — расписание + мониторинг
- [ ] `launchd` / `cron` запускает все парсеры раз в неделю
- [ ] Отчёт в файл: `output/run_log_YYYYMMDD.txt` — сколько найдено, сколько отфильтровано по email, сколько новых после дедупа
- [ ] Опционально: уведомление в Telegram bot (если хочешь)

---

## Острые углы / что нужно решить до запуска Фазы 2

1. **Apify token** — без него TikTok/Instagram через Apify не запустить. Создашь бесплатный аккаунт на apify.com → нужен будет `APIFY_TOKEN` в `.env`.
2. **TikTok Creator Marketplace** — есть ли у тебя бизнес-аккаунт TikTok Ads? Если да — это лучший free путь и почта приходит официально (без скрапинга).
3. **Языковая фокусировка** — Tier A/B/C/D — это в основном англоязычный рынок или RU тоже? Сейчас в конфиге 50/50. Если англо — уберу RU ключи, увеличу EN покрытие.
4. **Linktree / Beacons / Stan Store** — много блогеров не пишут email в bio, а только Linktree. Готов написать резолвер этих ссылок и достать email оттуда — это +30% к conversion. Заложу в Фазу 2.

---

## Команды (как делаешь сейчас → как будет)

```bash
# Сейчас (только YouTube, без фильтра по email):
python youtube_finder.py

# После Фазы 1:
python youtube_finder.py
# → output/youtube_YYYYMMDD.csv (только с email, формат CRM-колонок)
# → output/youtube_skipped_no_email_YYYYMMDD.csv (для ручной доразведки)

# После Фазы 2-3:
python tiktok_finder.py
python instagram_finder.py

# После Фазы 4:
python merge.py
# → output/master_crm_ready_YYYYMMDD.csv ← это импортируешь в Google Sheet

# После Фазы 5:
python sheets_push.py
# → автоматически append в таблицу
```

---

## Что от тебя нужно прямо сейчас, чтобы начать Фазу 1

Ничего. Фаза 1 — это правки в текущем `youtube_finder.py`, делаю сам. Готов начать в следующем сообщении.

Для Фазы 2 (TikTok) потребуется:
- Apify аккаунт + токен → положишь в `.env`
- ИЛИ TikTok Ads business аккаунт → дашь доступ через TikTok Creator Marketplace

Скажи "поехали Фаза 1" — и стартую.

---

## 🚀 Дополнительные направления роста (на будущее)

Эти каналы пока не реализованы в коде, но **обязательны** к проработке когда основной outreach pipeline разгонится. Источник идей — кейс Cedric Roberge ($125K за 80 дней на peptide-tracking app, см. наш чат).

### 🎯 Spark Ads / Meta Ads на лучших постах креаторов
**Что:** когда креатор постит органику про Timestripe и пост залетает (например 50K+ views) — мы запрашиваем у него TikTok Spark Code (или Meta-эквивалент), заливаем в свой Ads Manager и крутим как платную рекламу. Реклама показывается с аккаунта креатора, выглядит органически. Усиливает лучшие посты в 10-100x.

**Что нужно:**
- TikTok Ads / Meta Ads бизнес-аккаунт Timestripe
- Бюджет на ads ($5-50K/мес зависит от амбиций)
- В offer-шаблоне просить Spark Code у креатора заранее ("If the post performs well, we'd love to amplify it via Spark Ads")
- Скрипт для отслеживания performance постов креаторов (новая колонка `Post URL` в CRM + ручной мониторинг views)

**Почему это критично:** Cedric делал $125K не от органики, а от Spark Ads на best-performing креаторском контенте. Без этого канала ROI от outreach в 10x ниже потенциала.

**Когда внедрять:** после первых 5-10 живых коллабораций — когда есть готовый материал для усиления.

### 📱 Reddit organic strategy
**Что:** не классический outreach, а контент-стратегия. В subreddit'ах нашей ниши делать **косвенные** посты, отвечать в комментариях про Timestripe.

**Релевантные subreddits для Timestripe:**
- r/productivity (1.5M)
- r/getdisciplined (1.5M)
- r/notion (380K)
- r/selfimprovement (2M)
- r/bujo (bullet journal, 180K)
- r/studyblr (студенты, 100K)
- r/zettelkasten (50K)
- r/PKMS (личные базы знаний, 30K)

**Подход:**
- НЕ прямая реклама ("посмотрите наш Timestripe") — это спам, бан
- А индирект: "Ребят, последний месяц пытаюсь визуализировать свои цели на 5 лет вперёд. Notion слишком сухой, бумажный планер не масштабируется. Кто чем пользуется?" → в комментах отвечаешь "о, попробовал Timestripe, у них прикольная Horizons-фича, работает у меня"
- Или AMA от Sergey (CEO) — на r/productivity об основах планирования + продукт упоминается органически

**Кто делает:** контент-человек из команды Timestripe (не код), 1-2 поста/неделю.

**Метрика:** клики на timestripe.com из Reddit-источника (отслеживается через UTM).

**Когда внедрять:** параллельно с creator outreach. Дополняет, не конкурирует.

### 👥 Hiring VAs для масштабирования outreach
**Что:** Cedric хвалится 1000+ messages/неделю силами команды. Это не один человек — это нанятые VA.

**Когда внедрять:** когда у нас стабильный pipeline и pivot'ы шаблонов завершены — нанимаем 1-2 VA из Восточной Европы / Филиппин ($400-800/мес), они через нашу же систему отправляют по 100/день. Все процессы уже автоматизированы, VA только делает review + push send-button в админке.

**Метрика готовности:** когда у тебя >50% твоего рабочего времени занимает review текущих контактов в админке, а не написание кода.

### 💰 Третья платёжная модель — Flat Fee для micro-creators
**Что:** наш текущий оффер ($50/100k views) **в 10x менее привлекателен** для creators <10K subs чем то что платит Cedric ($100/8 posts flat). Это снижает conversion на Tier A/B/C.

**Предложение:** добавить третью модель — `$80 for 4 posts over 30 days` для under-10K. Tier D остаётся на views-based.

**Решение:** на руководстве. Если согласны — добавлю в offer-шаблон одной правкой.
