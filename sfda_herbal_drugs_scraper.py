# -*- coding: utf-8 -*-
"""
سكريبت سحب قائمة المستحضرات العشبية والصحية من الموقع القديم لـ SFDA
==================================================================
منفصل تمامًا عن باقي السكريبتات (الأدوية البشرية القديمة والجديدة) - محدش بيأثر على التاني.

الفكرة: نفس أسلوب سكريبت الأدوية البشرية القديمة، بس الـ endpoint هنا GET
بسيط (GetHerbalDrugs.php?page=N) من غير Payload، والرد شكله شوية مختلف
(data.results مباشرة، مش data.result.results).

طريقة التشغيل:
    pip install requests pandas openpyxl --break-system-packages
    python sfda_herbal_drugs_scraper.py
"""

import json
import time
import sys
import os
from datetime import datetime

import requests
import pandas as pd

BASE_URL = "https://oldsfda.sfda.gov.sa"
LIST_PAGE_URL = f"{BASE_URL}/ar/herbal-drugs"
API_URL = f"{BASE_URL}/GetHerbalDrugs.php"

COMMON_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ar-SA,ar;q=0.9,en-US;q=0.8,en;q=0.7",
}

API_HEADERS = {
    **COMMON_HEADERS,
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Referer": LIST_PAGE_URL,
    "X-Requested-With": "XMLHttpRequest",
}

REQUEST_DELAY_SECONDS = 0.6
MAX_PAGES = 3000
MAX_EMPTY_RETRIES = 2


def start_session() -> requests.Session:
    """يفتح جلسة، يزور صفحة العشبية الأول عشان ياخد كوكيز الحماية (WAF)."""
    session = requests.Session()
    resp = session.get(LIST_PAGE_URL, headers=COMMON_HEADERS, timeout=30)
    resp.raise_for_status()
    print(f"[✓] فتحنا صفحة المستحضرات العشبية بنجاح (status={resp.status_code}).")
    return session


def extract_rows(payload) -> list:
    """المسار المؤكد: payload.data.results (مباشرة، من غير مستوى result إضافي)."""
    try:
        results = payload["data"]["results"]
        if isinstance(results, list):
            return results
    except (KeyError, TypeError):
        pass

    # fallback عام لو الشكل اتغيّر لأي سبب
    if isinstance(payload, dict):
        common_keys = ["data", "items", "result", "results", "records", "rows"]
        for key in common_keys:
            if key in payload and isinstance(payload[key], list):
                return payload[key]
        list_values = [v for v in payload.values() if isinstance(v, list)]
        if list_values:
            return max(list_values, key=len)

    return []


def extract_page_count(payload):
    try:
        return payload["data"]["pageCount"]
    except (KeyError, TypeError):
        return None


def fetch_page(session: requests.Session, page: int):
    resp = session.get(API_URL, headers=API_HEADERS, params={"page": page}, timeout=30)
    resp.raise_for_status()

    try:
        payload = resp.json()
    except json.JSONDecodeError:
        print(f"[!] الرد في صفحة {page} مش JSON صحيح. أول 300 حرف من الرد:")
        print(resp.text[:300])
        return [], None

    return extract_rows(payload), extract_page_count(payload)


def scrape_all_herbal_drugs() -> pd.DataFrame:
    session = start_session()

    all_rows = []
    empty_retries = 0
    total_pages = MAX_PAGES

    for page in range(1, MAX_PAGES + 1):
        if page > total_pages:
            break
        try:
            rows, page_count = fetch_page(session, page)
        except requests.RequestException as e:
            print(f"[!] خطأ في الاتصال بصفحة {page}: {e}")
            print("    هنجرب تاني بعد ثانيتين...")
            time.sleep(2)
            try:
                rows, page_count = fetch_page(session, page)
            except requests.RequestException as e2:
                print(f"[x] فشلنا تاني في صفحة {page}: {e2}. هنوقف هنا.")
                break

        if page == 1 and page_count:
            total_pages = page_count
            print(f"[i] إجمالي عدد الصفحات المتوقع: {total_pages}")

        if not rows:
            empty_retries += 1
            print(f"[i] صفحة {page} رجعت فاضية ({empty_retries}/{MAX_EMPTY_RETRIES}).")
            if empty_retries >= MAX_EMPTY_RETRIES:
                print("[✓] يبدو إننا وصلنا لآخر صفحة، هنوقف السحب.")
                break
            time.sleep(REQUEST_DELAY_SECONDS)
            continue

        empty_retries = 0
        all_rows.extend(rows)
        print(f"[+] صفحة {page}/{total_pages}: جبنا {len(rows)} صف (الإجمالي لحد دلوقتي: {len(all_rows)})")

        time.sleep(REQUEST_DELAY_SECONDS)

    if not all_rows:
        print("[x] معندناش أي بيانات خالص. راجع الـ endpoint أو الكوكيز.")
        sys.exit(1)

    df = pd.DataFrame(all_rows)

    reg_col = next((c for c in df.columns if "register" in c.lower()), None)
    before = len(df)
    if reg_col:
        df = df.drop_duplicates(subset=[reg_col])
    else:
        df = df.drop_duplicates()
    after = len(df)
    if before != after:
        print(f"[i] شلنا {before - after} صف مكرر.")

    return df


def main():
    print("بدأنا نسحب بيانات المستحضرات العشبية والصحية من SFDA...\n")
    df = scrape_all_herbal_drugs()

    output_dir = "data"
    os.makedirs(output_dir, exist_ok=True)

    today_str = datetime.now().strftime("%Y-%m-%d")
    dated_filename = os.path.join(output_dir, f"SFDA_Herbal_{today_str}.xlsx")
    latest_filename = os.path.join(output_dir, "SFDA_Herbal_latest.xlsx")

    df.to_excel(dated_filename, index=False, engine="openpyxl")
    df.to_excel(latest_filename, index=False, engine="openpyxl")

    metadata = {
        "date": today_str,
        "count": len(df),
        "datedFileName": f"SFDA_Herbal_{today_str}.xlsx",
    }
    with open(os.path.join(output_dir, "last_updated_herbal.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False)

    print(f"\n[✓] تم! اتحفظ ملفين:")
    print(f"    - نسخة بتاريخ اليوم: {dated_filename}")
    print(f"    - نسخة ثابتة: {latest_filename}")
    print(f"    عدد المستحضرات: {len(df)}")
    print(f"    عدد الأعمدة: {len(df.columns)}")
    print(f"    الأعمدة: {list(df.columns)}")


if __name__ == "__main__":
    main()
