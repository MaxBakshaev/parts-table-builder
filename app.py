#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — микросервис сборки таблицы деталей.

Эндпоинты:
  GET  /health   -> проверка, что сервис жив
  POST /reset    -> очистить рабочую папку
  POST /upload   -> принять один файл (form-data, поле "file")
  POST /build    -> собрать таблицу и вернуть готовый .xlsx

Работает рядом с n8n в отдельном контейнере; n8n обращается по HTTP.
"""

import os, shutil, traceback, io, zipfile
from flask import Flask, request, send_file, jsonify

from build_parts_table import build, CURRENCIES, DEFAULT_CURRENCY

WORK = "/work"
IN = os.path.join(WORK, "input")
OUT = os.path.join(WORK, "output")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # до 200 МБ на файл


def ensure_dirs():
    os.makedirs(IN, exist_ok=True)
    os.makedirs(OUT, exist_ok=True)


@app.get("/health")
def health():
    import subprocess
    try:
        v = subprocess.run(["pdftoppm", "-v"], capture_output=True, text=True)
        poppler = (v.stderr or v.stdout).strip().splitlines()[0]
    except Exception as e:
        poppler = f"ERROR: {e}"
    return jsonify(status="ok", poppler=poppler,
                   currencies=list(CURRENCIES.keys()))


@app.post("/reset")
def reset():
    shutil.rmtree(WORK, ignore_errors=True)
    ensure_dirs()
    return jsonify(status="ok", message="workspace cleared")


@app.post("/upload")
def upload():
    ensure_dirs()
    f = request.files.get("file")
    if f is None:
        return jsonify(status="error", message="no file field"), 400
    # имя можно переопределить полем filename (важно: в нём номер детали)
    name = request.form.get("filename") or f.filename or "file.bin"
    name = os.path.basename(name)
    f.save(os.path.join(IN, name))
    return jsonify(status="ok", saved=name,
                   files=sorted(os.listdir(IN)))


@app.post("/zip")
def make_zip():
    """
    Собирает zip с сохранением структуры папок из ОДНОГО JSON-запроса.
    Тело: { "name": "Заказ.zip",
            "files": [ { "path": "Чертежи/деталь.pdf", "content": "<base64>" }, ... ] }
    Возвращает готовый zip (бинарно).
    """
    import base64
    payload = request.get_json(silent=True) or {}
    zip_name = payload.get("name") or "archive.zip"
    if not str(zip_name).lower().endswith(".zip"):
        zip_name = f"{zip_name}.zip"
    files = payload.get("files") or []
    if not files:
        return jsonify(status="error", message="files пуст"), 400

    buf = io.BytesIO()
    added, skipped = [], []
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, f in enumerate(files):
            rel = (f.get("path") or f.get("name") or f"file_{i+1}")
            rel = str(rel).lstrip("/").replace("\\", "/")
            content = f.get("content")
            if content is None:
                skipped.append(rel)
                continue
            try:
                data = base64.b64decode(content)
            except Exception:
                skipped.append(rel)
                continue
            zf.writestr(rel, data)
            added.append(rel)

    buf.seek(0)
    resp = send_file(buf, as_attachment=True, download_name=zip_name,
                     mimetype="application/zip")
    resp.headers["X-Zip-Entries"] = str(len(added))
    resp.headers["X-Zip-Skipped"] = str(len(skipped))
    return resp


@app.post("/build")
def build_table():
    ensure_dirs()
    body = request.get_json(silent=True) or {}
    currency = request.args.get("currency") or body.get("currency") or DEFAULT_CURRENCY
    if currency not in CURRENCIES:
        currency = DEFAULT_CURRENCY

    txts = [f for f in os.listdir(IN) if f.lower().endswith(".txt")]
    if not txts:
        return jsonify(status="error",
                       message="Details.txt не найден среди загруженных файлов",
                       files=sorted(os.listdir(IN))), 400

    out_path = os.path.join(OUT, "Parts_table_EN.xlsx")
    try:
        res = build(os.path.join(IN, txts[0]), IN, out_path, currency=currency)
    except Exception:
        return jsonify(status="error", traceback=traceback.format_exc()), 500

    # заголовки с диагностикой — их видно в n8n
    resp = send_file(out_path, as_attachment=True,
                     download_name="Parts_table_EN.xlsx")
    resp.headers["X-Parts-Count"] = str(res["parts"])
    resp.headers["X-Unmatched"] = ",".join(map(str, res["unmatched"])) or "none"
    resp.headers["X-Currency"] = res["currency"]
    return resp


if __name__ == "__main__":
    ensure_dirs()
    app.run(host="0.0.0.0", port=8000)
