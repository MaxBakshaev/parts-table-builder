# Установка рядом с n8n на Docker Hardened Image

Если ваш n8n собран на Docker Hardened Image, внутри него нет ни `apk`, ни `apt`,
ни `bash` — поставить туда poppler и Python нельзя. Решение: рядом запускается
отдельный маленький контейнер `parts-builder`, а n8n обращается к нему по HTTP.

Сам n8n при этом не меняется: ни пересборки, ни установки пакетов, ни перезапуска.

---

## 1. Загрузить проект на сервер

Склонируйте репозиторий (или скопируйте файлы `Dockerfile`, `app.py`,
`build_parts_table.py`) на сервер, например через WinSCP или `git clone`,
в `/opt/parts-builder`.

Проверьте, что файлы на месте:

```bash
ls /opt/parts-builder
```

Должно быть: `Dockerfile  app.py  build_parts_table.py`

---

## 2. Собрать образ

```bash
cd /opt/parts-builder
docker build -t parts-builder .
```

Занимает 1–2 минуты.

---

## 3. Узнать сеть, в которой работает n8n

```bash
docker inspect <имя_контейнера_n8n> --format '{{range $k,$v := .NetworkSettings.Networks}}{{$k}}{{end}}'
```

Команда выведет имя сети, например `n8n_default`. Подставьте его в следующий шаг.

---

## 4. Запустить сервис

```bash
docker run -d --name parts-builder --restart always --network ИМЯ_СЕТИ parts-builder
```

Проверка, что сервис отвечает (изнутри сети):

```bash
docker exec parts-builder curl -s http://localhost:8000/health
```

Ожидаемый ответ — строка со `"status": "ok"` и версией poppler.

> Порт наружу не открывается: сервис доступен только контейнерам в той же сети.
> Это и безопаснее, и достаточно — обращаться к нему будет только n8n.

---

## 5. Вызывать сервис из своего workflow

Готового workflow в этом репозитории нет — только HTTP API самого сервиса
(см. основной [README](../README.md#api)). В n8n добавьте свои узлы **HTTP Request**,
обращающиеся к `http://parts-builder:8000` (имя контейнера — как к обычному хосту
внутри общей docker-сети):

1. `POST /reset` — очистить рабочую папку перед новым заказом.
2. `POST /upload` (`multipart/form-data`, поле `file`, опционально `filename`) —
   по одному узлу на каждый файл: сначала `Details.txt`, затем все PDF-чертежи.
3. `POST /build?currency=USD` — получить готовый `.xlsx` в ответе (`responseFormat: file`).

Дальше — что угодно: сохранить файл в Google Drive/Dropbox, отправить письмом,
прикрепить к карточке сделки в вашей CRM и т.д. Это уже специфично для вашего
процесса, поэтому в репозиторий такой workflow не включён.

---

## Как это работает

```
n8n  --HTTP-->  parts-builder
 |                    |
 |  1. /reset         очистить рабочую папку
 |  2. /upload        принять Details.txt и все PDF (по одному)
 |  3. /build         собрать таблицу и вернуть .xlsx
 |
 └-> дальше — по вашему сценарию (Drive, email, CRM...)
```

---

## Обслуживание

```bash
docker logs parts-builder --tail 50      # логи сервиса
docker restart parts-builder             # перезапуск
```

Если правили `build_parts_table.py` (например, добавили валюту) — пересоберите:

```bash
cd /opt/parts-builder
docker build -t parts-builder . && docker rm -f parts-builder
docker run -d --name parts-builder --restart always --network ИМЯ_СЕТИ parts-builder
```

Обновления n8n на этот сервис никак не влияют — он живёт отдельно.

---

## Если что-то не так

| Симптом | Решение |
|---|---|
| n8n пишет `ECONNREFUSED` / не находит хост | контейнеры в разных сетях — проверьте шаг 3 и пересоздайте с нужной `--network` |
| `Details.txt не найден среди загруженных файлов` | узел `/upload` не отработал для `.txt`-файла, либо файл не дошёл — проверьте предыдущий шаг workflow |
| В таблице «PDF not found» | имя файла не в формате `N. Название.pdf` — переименуйте перед загрузкой |
| Пустой результат / 500 на `/build` | посмотрите `docker logs parts-builder --tail 50` |
