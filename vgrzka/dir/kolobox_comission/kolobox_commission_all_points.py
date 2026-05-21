# -*- coding: utf-8 -*-
import time
from pathlib import Path

import pandas as pd
import requests

BASE = "https://okno2.kolobox.ru"

# Твои два cookie (как уже делали)
COOKIE_STRING = (
    "XSRF-TOKEN=eyJpdiI6ImdYcUw2R3VMMDRXcGdDVDBnZHhsTHc9PSIsInZhbHVlIjoiOGI4akhnbU9NaFBJbGVmcXNkc3EwbkJ4SE9sYVhSczc4MkU3ZE1DdTh2NDl5S09sWVNjdnVcL3Y4ZFZXTFFRUlwvIiwibWFjIjoiNjk4Mzc3ZmQ3YmMyOWM3ZTZiN2ZhYjdkNGUwYTBhNWJjOWQyNDMzNmQ2ODJmYjNiOTViNWY0NjI5NmE1ZjllZSJ9; "
    "laravel_session=eyJpdiI6Imx4azRQc3ZBcFdydU1MR0lUeHlzQlE9PSIsInZhbHVlIjoiSFczbDh5NFlSQjQrZDhWdU5LOFI5YlFDdUdoU0REd0Jza1VZMFRCclVFTktiYVQyVzIwVXplaW9KdDBuXC9mc0doQlNxMlpXOUVUXC9STERnVDlzRTZ1STNCK3QreTAxNEpnZzFVa25UNGZnNzk4RmFHMWRmaFV0VTJiZDQ0ZFZXcSIsIm1hYyI6IjhlN2ViOTE4ZjY2N2JmYWY3YjJjOTE2MzVkMjk2NTNlYTY3MWU3MWY5ODUzMWNkYTMxYmI4MjkwYTJhMzAyYjcifQ%3D%3D"
)

# Торговые точки 
ADDRESS_IDS = [229, 2139, 2138, 4344]

FIRST_PAGE    = 0
MAX_PAGES     = 500    # предохранитель
REQUEST_PAUSE = 0.2    # сек, чтобы не спамить

OUT_FILENAME = "kolobox_commission_all_points.xlsx"


def get_commission_page(page: int):
    """Запрос одной страницы ajax /commission/{page} по cookie."""
    url = f"{BASE}/ajax/commission/{page}"
    headers = {
        "Cookie": COOKIE_STRING,
        "Accept": "application/json",
        "X-Requested-With": "XMLHttpRequest",
    }
    r = requests.get(url, headers=headers, timeout=45)
    r.raise_for_status()
    data = r.json()
    if not isinstance(data, list):
        raise RuntimeError(f"Неожиданный формат ответа на странице {page}: {type(data)} {data!r}")
    return data


def fetch_all_commission():
    all_rows = []
    page = FIRST_PAGE

    while page < MAX_PAGES:
        print(f"Страница {page}...", end="", flush=True)
        rows = get_commission_page(page)
        if not rows:
            print(" пусто, стоп.")
            break
        print(f" {len(rows)} записей.")
        all_rows.extend(rows)
        page += 1
        time.sleep(REQUEST_PAUSE)

    print(f"Всего записей (по всем точкам): {len(all_rows)}")
    return all_rows


def to_dataframe(rows):
    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # Оставляем только нужные точки
    if "address_id" in df.columns:
        df = df[df["address_id"].isin(ADDRESS_IDS)]

    # Приводим цену и количество к числам
    if "price" in df.columns:
        df["price"] = (
            df["price"]
            .astype(str)
            .str.replace(" ", "")
            .str.replace(",", ".", regex=False)
        )
        df["price"] = pd.to_numeric(df["price"], errors="coerce")

    if "quantity" in df.columns:
        df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")

    # Удобный порядок колонок (остальное — в хвост)
    preferred = [
        "address_id",
        "articul",
        "product_id",
        "tyres_mark",
        "tyres_model",
        "tyres_other",
        "tyres_tread_width",
        "tyres_profile_height",
        "tyres_diameter",
        "tyres_load_index",
        "tyres_speed_index",
        "price",
        "quantity",
        "document",
    ]
    cols = [c for c in preferred if c in df.columns] + [
        c for c in df.columns if c not in preferred
    ]
    df = df[cols]

    return df


def main():
    rows = fetch_all_commission()
    df = to_dataframe(rows)

    if df.empty:
        print("После фильтрации данных нет.")
        return

    out_path = Path(OUT_FILENAME).resolve()
    df.to_excel(out_path, index=False)
    print(f"Сохранено строк: {len(df)} → {out_path}")


if __name__ == "__main__":
    main()
