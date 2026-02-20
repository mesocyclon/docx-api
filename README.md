# docx-api

REST API для тестирования packaging-слоя библиотеки **docx-go**. Позволяет загрузить существующий `.docx`-файл, открыть его через `packaging.Open`, сохранить обратно через `packaging.Save` и получить результат.

## API Endpoints

| Метод  | Путь                            | Описание                                                                 |
|--------|---------------------------------|--------------------------------------------------------------------------|
| GET    | `/health`                       | Health check                                                             |
| GET    | `/ready`                        | Readiness probe                                                          |
| POST   | `/api/v1/documents/open`        | Загрузить .docx → получить JSON с метаданными документа                  |
| POST   | `/api/v1/documents/roundtrip`   | Загрузить .docx → Open → Save → скачать результирующий .docx             |
| POST   | `/api/v1/documents/validate`    | Загрузить .docx → Open → Save → Re-Open → JSON-отчёт о валидности       |

Все `POST`-эндпоинты принимают `multipart/form-data` с полем **`file`**.

## Быстрый старт с Docker

```bash
# Сборка и запуск
docker compose up -d

# Проверка здоровья
curl http://localhost:8080/health

# Тестирование round-trip
curl -X POST http://localhost:8080/api/v1/documents/roundtrip \
  -F file=@my-document.docx \
  -o roundtrip.docx

# Получение метаданных
curl -X POST http://localhost:8080/api/v1/documents/open \
  -F file=@my-document.docx

# Валидация packaging
curl -X POST http://localhost:8080/api/v1/documents/validate \
  -F file=@my-document.docx
```

## Локальная разработка

```bash
# Требования: Go 1.25+
make run

# Тесты
make test

# Тесты с покрытием
make test-cover
```

## Структура проекта

```
docx-api/
├── cmd/server/           # Точка входа приложения
│   └── main.go
├── internal/
│   ├── config/           # Конфигурация из ENV
│   ├── handler/          # HTTP-обработчики (transport layer)
│   ├── middleware/        # Logging, Recovery, CORS, MaxBody
│   └── service/          # Бизнес-логика (packaging operations)
├── pkg/response/         # JSON response helpers
├── docx-go/              # Библиотека docx-go (подключена через replace)
├── test/testdata/        # Тестовые .docx файлы
├── Dockerfile            # Multi-stage build
├── docker-compose.yml
├── Makefile
└── go.mod
```

## Конфигурация (ENV)

| Переменная          | По умолчанию | Описание                     |
|---------------------|-------------|-------------------------------|
| `PORT`              | `8080`      | Порт HTTP-сервера             |
| `MAX_UPLOAD_SIZE_MB`| `50`        | Максимальный размер файла (МБ)|
| `READ_TIMEOUT`      | `30s`       | Таймаут чтения запроса        |
| `WRITE_TIMEOUT`     | `60s`       | Таймаут записи ответа         |
| `SHUTDOWN_TIMEOUT`  | `10s`       | Таймаут graceful shutdown     |

## Примеры ответов

### POST /api/v1/documents/open

```json
{
  "title": "My Document",
  "creator": "John Doe",
  "application": "Microsoft Word",
  "parts_count": 12,
  "header_count": 1,
  "footer_count": 1,
  "media_files": ["image1.png"],
  "has_styles": true,
  "has_numbering": true,
  "has_comments": false,
  "has_footnotes": false,
  "has_endnotes": false
}
```

### POST /api/v1/documents/validate

```json
{
  "info": { ... },
  "original_size_bytes": 45230,
  "output_size_bytes": 44890,
  "success": true
}
```
