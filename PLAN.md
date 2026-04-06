# MCP Server for On-Premise MS Exchange — Подробная спецификация

---

## 1. Что это и зачем

**outlook-mcp** — MCP-сервер, который даёт LLM (Claude и другим) полный доступ к корпоративному
почтовому серверу MS Exchange, развёрнутому on-premise.

Без этого сервера Claude не знает ничего о вашей почте, календаре и коллегах.
С ним можно в диалоге:

- «Покажи непрочитанные письма от Иванова за эту неделю»
- «Создай встречу с командой в пятницу в 15:00, найди всем свободное время»
- «Перешли это письмо в отдел кадров с комментарием»
- «Напомни, что у меня на следующей неделе»
- «Найди контакт главного бухгалтера»

Протокол MCP (Model Context Protocol) — стандарт Anthropic для подключения внешних инструментов
к LLM. Сервер декларирует список **tools** (функции) и **resources** (данные), LLM вызывает их
по необходимости.

### 1.1 Цели и границы MVP

**Цели MVP**
- Надёжно подключаться к on-premise Exchange через EWS
- Дать LLM безопасный доступ к базовым сценариям: читать почту, отправлять письма, смотреть календарь, создавать встречи
- Возвращать предсказуемые JSON-ответы, удобные для LLM и для ручной отладки
- Работать локально через `stdio` без обязательной внешней инфраструктуры

**Что не входит в MVP**
- Полноценная синхронизация состояния или локальный кеш почтового ящика
- Поддержка Microsoft 365 / Exchange Online через Graph API
- Потоковая загрузка больших вложений и сложная обработка MIME
- Массовые операции по тысячам писем или фоновая индексация почты
- Полноценная multi-tenant архитектура

---

## 2. Как это работает технически

### 2.1 Протокол MCP

```
Пользователь → Claude Desktop / Claude Code
                    │
                    │  JSON-RPC 2.0 (stdio или SSE)
                    ▼
             ┌──────────────┐
             │  MCP Server  │  ← outlook-mcp (этот проект)
             │   (Python)   │
             └──────┬───────┘
                    │  HTTPS
                    │  EWS (SOAP XML)
                    ▼
           Exchange Server (on-premise)
           Exchange 2010 / 2013 / 2016 / 2019
```

Транспортов два:
- **stdio** — процесс запускается Claude Desktop, общение через stdin/stdout. Основной режим.
- **SSE** — HTTP-сервер (uvicorn), позволяет подключаться удалённо или из нескольких клиентов.

### 2.2 EWS (Exchange Web Services)

Exchange предоставляет SOAP API по адресу `https://mail.company.com/EWS/Exchange.asmx`.
EWS — зрелый, хорошо документированный протокол, работает со всеми версиями Exchange начиная с 2007.

Библиотека **exchangelib** инкапсулирует всю SOAP-работу и предоставляет Python-объекты:
`Account`, `Message`, `CalendarItem`, `Contact`, `Task`.

### 2.3 Аутентификация

| Метод | Когда использовать |
|---|---|
| **NTLM** | Стандарт для большинства on-premise Exchange в домене AD |
| **Basic** | Если NTLM недоступен (нужен HTTPS!) |
| **OAuth2 / ADFS** | Exchange 2016+ с настроенным ADFS |

Учётные данные передаются через переменные окружения, не хранятся на диске.

### 2.4 Жизненный цикл запроса

```
1. Claude решает вызвать tool, например list_emails
2. MCP SDK вызывает соответствующий Python-обработчик
3. Обработчик через ExchangeClient делает EWS-запрос (FindItem / GetItem)
4. Exchange возвращает SOAP-ответ
5. exchangelib десериализует в Python-объекты
6. Обработчик конвертирует в упрощённый dict / Pydantic-модель
7. MCP SDK сериализует в JSON и возвращает Claude
8. Claude формирует ответ пользователю
```

### 2.5 Ключевые инженерные решения

- Все datetime на входе принимаются в ISO 8601; если timezone не указан, используется `EXCHANGE_TIMEZONE`
- Все datetime на выходе возвращаются в ISO 8601 с timezone
- Update/delete-операции должны быть максимально идемпотентны; create/send-операции обязаны явно документировать, что повторный вызов может создать дубликат, если клиент не реализует защиту на своей стороне
- По умолчанию сервер отдаёт компактные ответы; полные тела писем, HTML и вложения загружаются только по явному запросу
- Все операции с сетью проходят через единый `ExchangeClient`, чтобы централизованно управлять таймаутами, retry, логированием и маппингом ошибок

---

## 3. Полная функциональность

### 3.0 Scope по фазам

**Входит в MVP**
- `ping_exchange`
- `list_emails`
- `get_email`
- `send_email`
- `list_events`
- `get_event`
- `create_event`
- `find_free_slots`

**Post-MVP / расширения**
- почтовые операции сверх базового чтения/отправки: `search_emails`, `reply_email`, `forward_email`, `move_email`, `copy_email`, `create_draft`, `send_draft`, `mark_email`, `list_folders`, `create_folder`, `get_attachment`
- календарные расширения: `update_event`, `delete_event`, `respond_to_invite`, `get_my_availability`, `list_calendars`
- контакты целиком

### 3.1 Почта (Email)

#### Tools

**`list_emails`**
```
Параметры:
  folder        string   "inbox" | "sent" | "drafts" | "deleted" | путь папки
  limit         int      макс. кол-во писем (default 20, max 100)
  offset        int      пагинация
  from_address  string?  фильтр по отправителю
  subject       string?  фильтр по теме (contains)
  since         date?    письма не старше даты
  before        date?    письма не новее даты
  unread_only   bool     только непрочитанные
  has_attachments bool   только с вложениями

Возвращает: список { id, subject, from, to, date, is_read, has_attachments, preview }
```

**`get_email`**
```
Параметры:
  id  string  идентификатор письма (EWS ItemId)

Возвращает: полное письмо {
  id, subject, from, to, cc, bcc, date, is_read,
  body_text, body_html,
  attachments: [{ name, size, content_type, id }],
  headers,
  conversation_id
}
```

**`search_emails`**
```
Параметры:
  query         string   поисковый запрос (полный текст, тема, отправитель)
  folder        string?  ограничить папкой (default: все папки)
  limit         int      default 20

Возвращает: список писем (как list_emails)

Примечание: использует EWS FindItem с AQS (Advanced Query Syntax)
```

**`send_email`**
```
Параметры:
  to            string[]  список адресов
  subject       string
  body          string    текст или HTML
  body_type     string    "text" | "html" (default "text")
  cc            string[]?
  bcc           string[]?
  reply_to      string?
  attachments   path[]?   пути к локальным файлам для прикрепления
  importance    string?   "low" | "normal" | "high"

Возвращает: { id, status: "sent", warning? }

Примечание: перед отправкой валидируются адреса получателей и существование файлов из `attachments`.
```

**`reply_email`**
```
Параметры:
  id            string   ID исходного письма
  body          string   текст ответа
  reply_all     bool     ответить всем (default false)
  attachments   path[]?

Возвращает: { id, status: "sent", warning? }
```

**`forward_email`**
```
Параметры:
  id            string    ID письма
  to            string[]  получатели
  comment       string?   комментарий перед телом письма
  attachments   path[]?   доп. вложения

Возвращает: { id, status: "sent", warning? }
```

**`move_email`**
```
Параметры:
  id            string   ID письма
  folder        string   целевая папка ("inbox", "sent", или путь)

Возвращает: { id, status: "moved", new_folder }
```

**`delete_email`**
```
Параметры:
  id            string   ID письма
  hard_delete   bool     безвозвратно (default false — в корзину)

Возвращает: { id, status: "deleted" }

Примечание: `hard_delete=true` относится к опасным операциям и должен быть явно указан LLM, не подставляться сервером по умолчанию.
```

**`mark_email`**
```
Параметры:
  id            string   ID письма
  read          bool?    пометить прочитанным/непрочитанным
  flag          string?  "flagged" | "complete" | "none"
  importance    string?  "low" | "normal" | "high"

Возвращает: { id, status: "updated", updated_fields }
```

**`list_folders`**
```
Параметры:
  parent  string?  начать с папки (default: root)
  depth   int      глубина вложенности (default 2)

Возвращает: дерево { name, path, unread_count, total_count, children[] }
```

**`create_folder`**
```
Параметры:
  name    string   имя папки
  parent  string?  родительская папка (default: inbox)

Возвращает: { id, status: "created", path }
```

**`copy_email`**
```
Параметры:
  id            string   ID письма
  folder        string   целевая папка ("inbox", "archive", или путь)

Возвращает: { source_id, new_id, status: "copied", new_folder }

Примечание: использует EWS CopyItem; исходное письмо остаётся на месте.
```

**`create_draft`**
```
Параметры:
  to            string[]  список адресов
  subject       string
  body          string
  body_type     string    "text" | "html" (default "text")
  cc            string[]?
  bcc           string[]?
  attachments   path[]?

Возвращает: { id, status: "draft", warning? }
```

**`send_draft`**
```
Параметры:
  id  string  ID черновика

Возвращает: { id, status: "sent" }

Примечание: эквивалентен отправке сохранённого черновика из папки Drafts.
Перед вызовом рекомендуется показать пользователю содержимое через get_email.
```

**`get_attachment`**
```
Параметры:
  email_id      string   ID письма
  attachment_id string   ID вложения
  save_path     string?  куда сохранить (default: temp)

Возвращает: { filename, size, saved_path, content_type }

Примечание: если `save_path` не указан, вложение сохраняется во временную директорию сервера.
Если файл с таким именем уже существует, сервер должен выбрать безопасное уникальное имя и вернуть фактический `saved_path`.
В MVP возврат бинарного содержимого в ответе tool не поддерживается.
```

#### Resources (read-only данные, без вызова)

- `mailbox://folders` — полное дерево папок
- `mailbox://email/{id}` — письмо по ID
- `mailbox://inbox?limit=10` — последние письма входящих
- `mailbox://drafts` — черновики

---

### 3.2 Календарь (Calendar)

#### Tools

**`list_events`**
```
Параметры:
  start         datetime  начало диапазона (default: сегодня)
  end           datetime  конец диапазона (default: +7 дней)
  calendar_id   string?   ID календаря (default: основной)
  include_recurring bool  разворачивать повторяющиеся (default true)

Возвращает: список {
  id, subject, start, end, location, organizer,
  attendees[], is_all_day, is_recurring, my_response,
  online_meeting_url
}
```

**`get_event`**
```
Параметры:
  id  string

Возвращает: полное событие {
  id, subject, start, end, location, body,
  organizer, attendees: [{ email, name, response_type }],
  is_all_day, is_recurring, recurrence_pattern,
  reminder_minutes, online_meeting_url,
  categories[], importance
}
```

**`create_event`**
```
Параметры:
  subject           string
  start             datetime
  end               datetime
  calendar_id       string?   ID календаря (default: основной)
  location          string?
  body              string?
  attendees         string[]?  email-адреса
  is_all_day        bool       default false
  reminder_minutes  int?       default 15
  recurrence        object?    {
                                 type: "daily"|"weekly"|"monthly"|"yearly",
                                 interval: int,
                                 end_date: date?,
                                 occurrences: int?,
                                 days_of_week: string[]?
                               }
  categories        string[]?
  importance        string?    "low"|"normal"|"high"
  online_meeting    bool       создать Teams/Skype-ссылку (если настроено)

Возвращает: { id, status: "created", subject, start, end, invite_sent: bool, warning? }

Примечание: сервер должен валидировать `end > start`; для all-day событий границы нормализуются к началу/концу дня в целевой timezone.
```

**`update_event`**
```
Параметры:
  id              string
  subject         string?
  start           datetime?
  end             datetime?
  location        string?
  body            string?
  add_attendees   string[]?
  remove_attendees string[]?
  reminder_minutes int?
  send_updates    string    "none"|"all"|"modified" (default "all")

Возвращает: { id, status: "updated", updated_fields }
```

**`delete_event`**
```
Параметры:
  id              string
  notify_attendees bool   default true
  cancel_message  string? текст в уведомлении об отмене

Возвращает: { id, status: "deleted" }

Примечание: удаление встречи с участниками считается опасной операцией; в логах фиксируется только факт вызова без текста `cancel_message`.
```

**`respond_to_invite`**
```
Параметры:
  id        string   ID события
  response  string   "accept" | "tentative" | "decline"
  message   string?  сопроводительный текст

Возвращает: { id, status }
```

**`find_free_slots`**
```
Параметры:
  attendees   string[]  email-адреса участников
  duration    int       длительность встречи в минутах
  start       datetime  начало поиска (default: завтра)
  end         datetime  конец поиска (default: +5 рабочих дней)
  work_hours  object?   { start: "09:00", end: "18:00" }

Возвращает: список свободных слотов [{
  start, end,
  all_available: bool,
  busy_attendees: []  (кто занят в альтернативных слотах)
}]

Примечание: использует EWS GetUserAvailability / FreeBusy API
```

**`get_my_availability`**
```
Параметры:
  start  datetime
  end    datetime

Возвращает: { free_slots[], busy_slots[{ start, end, subject? }] }
```

**`list_calendars`**
```
Параметры: нет

Возвращает: список {
  id, name, is_default, color?, owner_email
}

Примечание: `id` — канонический идентификатор для подстановки в `calendar_id` у create_event / list_events.
Возвращает личные календари; общие/shared-календари перечисляются отдельно если есть права.
```

#### Resources

- `calendar://events?start=&end=` — события в диапазоне
- `calendar://today` — события сегодня

---

### 3.3 Контакты (Contacts)

Два источника: **личные контакты** (папка Contacts) и **GAL** (Global Address List — корпоративный каталог AD).

#### Tools

**`search_contacts`**
```
Параметры:
  query   string   имя, email, организация, должность
  source  string   "personal" | "gal" | "all" (default "all")
  limit   int      default 10

Возвращает: список {
  id, display_name, email_addresses[], phone_numbers[],
  company, job_title, department, source
}
```

**`get_contact`**
```
Параметры:
  id  string

Возвращает: полная карточка {
  id, display_name, first_name, last_name,
  email_addresses: [{ type, address }],
  phone_numbers: [{ type, number }],
  addresses: [{ type, street, city, ... }],
  company, job_title, department, manager,
  notes, photo_url, birthday,
  source: "personal" | "gal"
}
```

**`create_contact`** (только личные контакты)
```
Параметры:
  display_name    string
  first_name      string?
  last_name       string?
  email           string?
  phone           string?
  company         string?
  job_title       string?
  notes           string?

Возвращает: { id, status: "created" }
```

**`update_contact`** (только личные контакты)
```
Параметры:
  id  string
  ...те же поля что create_contact

Возвращает: { id, updated_fields }
```

**`delete_contact`** (только личные контакты)
```
Параметры:
  id  string

Возвращает: { id, status: "deleted" }
```

---

### 3.4 Системные / служебные tools

**`ping_exchange`**
```
Проверяет соединение с сервером Exchange.
Возвращает: { status: "ok"|"error", server, version, latency_ms, error? }
```

**`get_mailbox_info`**
```
Возвращает: {
  email_address, display_name, timezone,
  mailbox_size_mb, quota_mb,
  exchange_version
}
```

---

## 4. Модели данных

### Email
```python
class EmailSummary(BaseModel):
    id: str
    subject: str
    from_: EmailAddress
    to: list[EmailAddress]
    date: datetime
    is_read: bool
    has_attachments: bool
    preview: str              # первые 200 символов тела
    importance: str
    categories: list[str]

class EmailFull(EmailSummary):
    cc: list[EmailAddress]
    bcc: list[EmailAddress]
    body_text: str
    body_html: str | None
    attachments: list[Attachment]
    conversation_id: str
    headers: dict[str, str]
```

Правила возврата тела письма:
- `body_text` обязателен и содержит текстовую версию письма, пригодную для LLM
- `body_html` возвращается только если HTML реально присутствует у исходного письма
- если тело письма слишком большое, сервер может усечь `body_text` и/или `body_html`, добавив `truncated: true`
- inline-изображения не встраиваются в ответ как бинарные данные; они остаются вложениями/ссылками Exchange

### CalendarEvent
```python
class CalendarEvent(BaseModel):
    id: str
    subject: str
    start: datetime
    end: datetime
    location: str | None
    organizer: EmailAddress
    attendees: list[Attendee]  # { email, name, response: accept/tentative/decline/unknown }
    is_all_day: bool
    is_recurring: bool
    my_response: str           # accept/tentative/decline/unknown
    online_meeting_url: str | None
    body: str | None
    reminder_minutes: int | None
    categories: list[str]
```

---

## 5. Обработка ошибок

| Ситуация | Поведение |
|---|---|
| Exchange недоступен | `{ error: "exchange_unavailable", message: "..." }` |
| Неверные учётные данные | `{ error: "auth_failed" }` (не раскрывать детали) |
| Элемент не найден | `{ error: "not_found", id: "..." }` |
| Нет прав | `{ error: "permission_denied" }` |
| Throttling (EWS 503) | Автоматический retry с exponential backoff (3 попытки) |
| Большой ответ | Усечение + `{ truncated: true, total: N }` |
| Таймаут | `{ error: "timeout", timeout_seconds: 30 }` |
| Неверные параметры | `{ error: "validation_error", details: [...] }` |
| Конфликт обновления | `{ error: "conflict", message: "item changed remotely" }` |

Все ошибки возвращаются как валидный JSON, не бросают исключения на уровне MCP —
LLM должен получить осмысленный ответ в любом случае.

### Формат ошибок

Используется `CallToolResult(isError=True)` — официальный MCP-механизм. LLM получает полные структурированные данные и понимает, что вызов завершился ошибкой. На успешных путях данные возвращаются плоско, без обёртки `ok/data`.

```python
# Ошибка
return CallToolResult(
    content=[TextContent(type="text", text=json.dumps({
        "error": "validation_error",
        "message": "end must be greater than start",
        "details": [{ "field": "end", "reason": "must be greater than start" }]
    }))],
    isError=True
)

# Успех
return CallToolResult(
    content=[TextContent(type="text", text=json.dumps({
        "id": "AAMk...",
        "status": "sent"
    }))]
)
```

`McpError` (protocol-level JSON-RPC error) используется только для инфраструктурных сбоев (сервер не смог обработать запрос вообще), не для бизнес-ошибок.

### Наблюдаемость и диагностика

- Каждый tool call должен логироваться с `tool_name`, `request_id`, длительностью и итоговым статусом
- `request_id` должен проходить через весь путь обработки: MCP handler -> `ExchangeClient` -> EWS вызов -> результат/ошибка
- В логи не попадают тела писем, HTML, вложения, пароли, токены и `cancel_message`
- Для retry полезно логировать номер попытки и тип ошибки (`timeout`, `throttling`, `auth_failed` и т.д.)

---

## 6. Безопасность

### Что защищается
- Учётные данные передаются только через env-переменные, никогда не логируются
- Тела писем не кэшируются на диске
- Tool `delete_email` по умолчанию мягкое удаление (в корзину)
- Tool `delete_event` с `notify_attendees=true` по умолчанию
- Логи по умолчанию не содержат `body_text`, `body_html`, заголовки писем, токены и пароли
- Для HTML-писем в ответы наружу возвращается сырой HTML без попытки "исполнить" или рендерить его на стороне сервера

### Что НЕ делает сервер
- Не хранит состояние между сессиями
- Не имеет собственной БД
- Не отправляет данные куда-либо кроме Exchange

### Подтверждение опасных операций

Сервер не ведёт интерактивный диалог сам, но спецификация должна считать следующие операции
опасными и требующими явного намерения со стороны LLM/пользователя:

- `delete_email(hard_delete=true)`
- `delete_event(...)`
- `send_email(...)` внешним адресатам вне корпоративного домена
- `forward_email(...)` с вложениями
- `move_email(...)` в нестандартную папку, если путь был сгенерирован, а не выбран из `list_folders`

На уровне реализации это означает:
- никаких "умных" дефолтов, которые усиливают разрушительное действие
- максимально подробная валидация аргументов до сетевого вызова
- audit-friendly лог события без утечки чувствительного содержимого

### Impersonation
Опционально: если учётная запись имеет права `ApplicationImpersonation`, сервер
может работать от имени другого пользователя (указывается через `EXCHANGE_IMPERSONATE_AS`).
Используется для корпоративных сценариев (ИТ-отдел, CRM-интеграции).

---

## 7. Конфигурация

### Переменные окружения

```env
# Обязательные
EXCHANGE_SERVER=https://mail.company.com/EWS/Exchange.asmx
EXCHANGE_USERNAME=DOMAIN\username
EXCHANGE_PASSWORD=secret
EXCHANGE_AUTH_TYPE=NTLM        # NTLM | Basic | OAuth2

# Версия Exchange (auto-detect если не задано)
EXCHANGE_VERSION=EXCHANGE_2016  # Exchange2010_SP2 | Exchange2013 | Exchange2016 | Exchange2019

# OAuth2 / ADFS (если AUTH_TYPE=OAuth2)
OAUTH2_CLIENT_ID=
OAUTH2_CLIENT_SECRET=
OAUTH2_TOKEN_URL=https://adfs.company.com/adfs/oauth2/token
OAUTH2_SCOPE=

# Опциональные
EXCHANGE_IMPERSONATE_AS=       # email для impersonation
EXCHANGE_TIMEOUT=30            # таймаут запросов в секундах
EXCHANGE_MAX_RETRIES=3
EXCHANGE_TIMEZONE=Europe/Moscow
ATTACHMENT_MAX_SIZE_MB=10      # лимит на загрузку/отправку вложений

# MCP транспорт
MCP_TRANSPORT=stdio            # stdio | sse
MCP_SSE_HOST=127.0.0.1
MCP_SSE_PORT=8080

# Логирование
LOG_LEVEL=INFO                 # DEBUG | INFO | WARNING | ERROR
LOG_FILE=                      # путь к файлу (по умолчанию stderr)
```

---

## 8. Структура проекта

```
outlook-mcp/
├── src/
│   └── outlook_mcp/
│       ├── __init__.py
│       ├── server.py              # MCP-сервер, регистрация tools и resources
│       ├── config.py              # Pydantic Settings, загрузка .env
│       ├── exchange_client.py     # Singleton-обёртка над exchangelib.Account
│       ├── auth.py                # Логика выбора credentials (NTLM/OAuth2/...)
│       ├── models.py              # Pydantic-модели EmailFull, CalendarEvent, ...
│       ├── errors.py              # Кастомные исключения, конвертация EWS-ошибок
│       └── tools/
│           ├── __init__.py
│           ├── email.py           # list_emails, get_email, send_email, ...
│           ├── calendar.py        # list_events, create_event, find_free_slots, ...
│           ├── contacts.py        # search_contacts, get_contact, ...
│           └── system.py          # ping_exchange, get_mailbox_info
├── tests/
│   ├── conftest.py                # фикстуры с mock Exchange
│   ├── test_email.py
│   ├── test_calendar.py
│   ├── test_contacts.py
│   ├── test_system.py
│   └── test_errors.py
├── .env.example
├── pyproject.toml
├── Dockerfile
└── README.md
```

---

## 9. Зависимости

```toml
[project]
name = "outlook-mcp"
requires-python = ">=3.11"
dependencies = [
    "mcp>=1.0",              # Anthropic MCP SDK
    "exchangelib>=5.0",      # EWS-клиент
    "pydantic>=2.0",         # модели данных и валидация
    "pydantic-settings>=2.0",# загрузка конфига из env
    "requests-ntlm>=1.3",    # NTLM аутентификация
    "uvicorn>=0.30",         # для SSE-транспорта
    "python-dateutil>=2.9",  # парсинг дат
]

[project.scripts]
outlook-mcp = "outlook_mcp.server:main"
```

**Дополнительно для dev/test**
- `pytest`
- `pytest-asyncio` (если SDK или transport потребуют async-тестов)
- `respx` или `responses` только если понадобится мокать HTTP вне `exchangelib`
- `ruff` для линтинга и единообразного стиля

---

## 10. Примеры диалогов

**Организация встречи** — типичный сценарий с несколькими tool calls:
```
Пользователь: Запланируй встречу с иванов@ и сидоров@ на следующей неделе на час

→ find_free_slots([иванов, сидоров], duration=60, start=..., end=...)
→ Предлагает 3 слота, пользователь выбирает
→ create_event(subject, start, end, attendees=[...])
```

**Draft-workflow** — безопасная отправка:
```
Пользователь: Напиши ответ Петрову и дай посмотреть перед отправкой

→ get_email(id) — читает исходное
→ create_draft(to, subject, body) — сохраняет черновик
→ get_email(draft_id) — показывает пользователю
→ send_draft(draft_id) — после подтверждения
```

---

## 11. Ограничения и известные нюансы

| Ограничение | Детали |
|---|---|
| Exchange Online (O365) | **Не поддерживается** — там используется Graph API, это отдельный проект |
| EWS может быть отключён | В Exchange 2019+ Microsoft рекомендует REST API; EWS всё ещё работает, но надо проверить |
| Поиск по GAL | EWS не даёт полный список GAL; поиск работает, полный список — нет |
| Повторяющиеся события | Создание сложных recurrence-паттернов может работать нестабильно |
| Вложения > 10 МБ | Потоковая передача не реализована в MVP; лимит задаётся через `ATTACHMENT_MAX_SIZE_MB` |
| `search_emails` (AQS) | AQS-поиск требует включённого Managed Search на сервере; Exchange 2010 может не поддерживать |
| Часовые пояса | Exchange хранит в UTC, конвертация настраивается через `EXCHANGE_TIMEZONE` |

---

## 12. Критерии готовности

### MVP считается готовым, если:

- сервер поднимается в `stdio`-режиме и проходит `ping_exchange`
- `list_emails`, `get_email`, `send_email`, `list_events`, `create_event`, `find_free_slots` покрыты тестами
- ошибки аутентификации, таймаута и `not_found` возвращаются в согласованном формате
- `.env.example` достаточно, чтобы новый разработчик понял все обязательные настройки
- README содержит локальный сценарий запуска и минимальный пример конфигурации Claude Desktop

### Минимальный план тестирования

- unit-тесты на маппинг моделей EWS -> Pydantic
- unit-тесты на валидацию входных параметров
- unit-тесты на конвертацию ошибок EWS -> API-ошибки
- интеграционный smoke-test с тестовым Exchange или стендом заказчика
- ручная проверка опасных операций: отправка письма, удаление письма, отмена встречи

---

## 13. Приоритет реализации

```
MVP (рабочий продукт):
  Фаза 1: Скелет + подключение + ping
  Фаза 2: Почта (чтение + отправка)
  Фаза 3: Календарь (чтение + создание + find_free_slots)

Полная версия:
  Фаза 4: Контакты
  Фаза 5: OAuth2/ADFS, SSE-транспорт, вложения, Impersonation, Docker
```
