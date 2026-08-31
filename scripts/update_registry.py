#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт автообновления снимка реестра российской промышленной продукции.

Что делает:
1. Скачивает свежий CSV с портала открытых данных Минпромторга
   (ссылка на файл каждый день меняется — в неё зашита сегодняшняя дата).
2. Разбирает CSV и раскладывает записи по кусочкам (чанкам) на основе кода
   ОКПД2 — так же, как это было сделано вручную при первой сборке инструмента.
3. Перезаписывает файлы в папке data/ (кроме registry_index.js, который тоже
   пересоздаётся) — только КУСКИ РЕЕСТРА, ничего в самом index.html не трогает.

Запускается автоматически (см. .github/workflows/update-registry.yml),
но можно запустить и вручную: python scripts/update_registry.py
"""

import csv
import io
import json
import os
import re
import sys
import shutil
from datetime import datetime, timedelta, timezone

import urllib.request
import urllib.error

# ---------- настройки ----------

BASE_URL = "https://minpromtorg.gov.ru/opendata/1000000012-ReestrProducts/data-{date}-structure-20210405.csv"
DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
THRESHOLD = 30000  # максимум строк в одном файле-чанке, иначе дробим дальше

# Реалистичные заголовки браузера — некоторые сайты блокируют запросы без них
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "text/csv,*/*;q=0.9",
    "Accept-Language": "ru-RU,ru;q=0.9",
}

csv.field_size_limit(10**7)


def log(msg):
    print(msg, flush=True)


def download_csv():
    """Пробует скачать сегодняшний файл, при неудаче — вчерашний (на случай
    расхождения часовых поясов между сервером МПТ и раннером GitHub)."""
    now = datetime.now(timezone.utc)
    for days_back in (0, 1, 2):
        d = now - timedelta(days=days_back)
        date_str = d.strftime("%Y%m%d")
        url = BASE_URL.format(date=date_str)
        log(f"Пробую скачать: {url}")
        req = urllib.request.Request(url, headers=HEADERS)
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                data = resp.read()
                log(f"  успех, {len(data)} байт")
                return data, d.strftime("%d.%m.%Y")
        except urllib.error.HTTPError as e:
            log(f"  HTTP-ошибка: {e.code} {e.reason}")
        except urllib.error.URLError as e:
            log(f"  ошибка сети: {e.reason}")
    raise RuntimeError(
        "Не удалось скачать файл реестра ни за одну из последних дат. "
        "Возможно, сайт минпромторга заблокировал автоматический запрос, "
        "либо изменился формат ссылки/структуры файла."
    )


def excel_like_date(iso_date):
    """'2027-04-09' -> '09.04.2027'; '-' или пусто -> ''."""
    if not iso_date or iso_date.strip() in ("-", ""):
        return ""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", iso_date.strip())
    if not m:
        return iso_date.strip()
    y, mo, da = m.groups()
    return f"{da}.{mo}.{y}"


def clean(v):
    if v is None:
        return ""
    v = v.strip()
    return "" if v == "-" else v


def parse_rows(csv_bytes):
    """Читает CSV и группирует строки по разделу ОКПД2 (первый сегмент кода).
    Возвращает dict: { "27": [row, row, ...], "_none": [...] }"""
    text = csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    header = next(reader)
    idx = {h: i for i, h in enumerate(header)}

    required = ["Nameoforg", "INN", "Registernumber", "Docdate", "Docvalidtill",
                "Enddate", "Productname", "OKPD2", "Score", "Percentage", "Scoredesc"]
    missing = [f for f in required if f not in idx]
    if missing:
        raise RuntimeError(
            f"В CSV не хватает ожидаемых столбцов: {missing}. "
            f"Похоже, минпромторг поменял структуру файла — нужно смотреть заново."
        )

    buckets = {}
    total = 0
    for row in reader:
        if not row:
            continue
        total += 1
        okpd = clean(row[idx["OKPD2"]])
        trimmed = [
            clean(row[idx["Nameoforg"]]),
            clean(row[idx["INN"]]),
            clean(row[idx["Registernumber"]]),
            excel_like_date(row[idx["Docdate"]]),
            excel_like_date(row[idx["Docvalidtill"]]),
            excel_like_date(row[idx["Enddate"]]),
            clean(row[idx["Productname"]]),
            okpd,
            clean(row[idx["Score"]]),
            clean(row[idx["Percentage"]]),
            clean(row[idx["Scoredesc"]]),
        ]
        section = okpd.split(".")[0] if okpd else "_none"
        # защита от «мусорных» кодов (пробелы/переводы строк/прочее)
        if not re.match(r"^[0-9_]+$", section):
            section = "_anomalies"
        buckets.setdefault(section, []).append(trimmed)

    log(f"Всего прочитано строк: {total}, разделов: {len(buckets)}")
    return buckets


def split_oversized(buckets):
    """Рекурсивно дробит слишком большие разделы на подгруппы по коду,
    пока каждый файл не станет меньше THRESHOLD строк."""
    final = {}
    work = list(buckets.items())
    round_no = 0
    while work:
        round_no += 1
        next_round = []
        for key, rows in work:
            if len(rows) <= THRESHOLD or round_no > 8:
                final[key] = rows
                continue
            depth = len(key.split("."))
            sub = {}
            for r in rows:
                okpd = r[7]
                parts = okpd.split(".")
                if len(parts) > depth:
                    subkey = ".".join(parts[: depth + 1])
                else:
                    subkey = okpd + "_x"
                sub.setdefault(subkey, []).append(r)
            if len(sub) <= 1:
                # дальше делить нечем (все строки — один и тот же код) —
                # дробим просто по количеству, кусками по 15000
                base = key
                CH = 15000
                for i in range(0, len(rows), CH):
                    final[f"{base}__p{i // CH}"] = rows[i : i + CH]
            else:
                for k, v in sub.items():
                    next_round.append((k, v))
        work = next_round
    return final


def write_output(buckets, updated_label):
    if os.path.isdir(DATA_DIR):
        for f in os.listdir(DATA_DIR):
            if f.endswith(".js") and f != "registry_index.js":
                os.remove(os.path.join(DATA_DIR, f))
    else:
        os.makedirs(DATA_DIR)

    index = []
    for key, rows in buckets.items():
        safe_name = re.sub(r"[^0-9A-Za-z_.]", "_", key) + ".js"
        path = os.path.join(DATA_DIR, safe_name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"window.REGISTRY_CHUNKS[{json.dumps(key)}]=")
            f.write(json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
            f.write(";")
        index.append({"key": key, "file": "data/" + safe_name, "rows": len(rows)})

    with open(os.path.join(DATA_DIR, "registry_index.js"), "w", encoding="utf-8") as f:
        f.write(f"window.REGISTRY_UPDATED = {json.dumps(updated_label)};\n")
        f.write("window.REGISTRY_INDEX = ")
        f.write(json.dumps(index, ensure_ascii=False, separators=(",", ":")))
        f.write(";\n")

    log(f"Записано файлов-чанков: {len(index)}")


def main():
    csv_bytes, updated_label = download_csv()
    buckets = parse_rows(csv_bytes)
    buckets = split_oversized(buckets)
    write_output(buckets, updated_label)
    log("Готово.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"ОШИБКА: {e}")
        sys.exit(1)
