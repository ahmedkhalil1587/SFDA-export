# -*- coding: utf-8 -*-
"""
سكريبت سحب قائمة الأدوية من الموقع الجديد لـ SFDA (sfda.gov.sa) - نسخة التفاصيل الكاملة
==================================================================
منفصل تمامًا عن سكريبت الموقع القديم (sfda_drugs_scraper.py) - محدش بيأثر على التاني.

الفكرة:
  1) نسحب جدول القائمة (كل الصفحات) - بيدينا: الاسم العلمي، التجاري، التركيز،
     الشكل الصيدلاني، السعر، ورابط تفاصيل كل دواء.
  2) نفتح صفحة تفاصيل كل دواء لوحده (details_data?nid=...&id=...) ونجيب
     الـ29 حقل الكاملة (رقم التسجيل، طريقة الاستخدام، حجم العبوة، الشركة
     المصنّعة، الوكلاء، ATC codes، حالة الترخيص... إلخ) ونربطها بنفس الصف.

تحذير: الخطوة التانية (التفاصيل) بتضيف حوالي 15 ألف طلب إضافي، فالسكريبت
هياخد وقت أطول بكتير من نسخة القائمة البسيطة (ممكن يوصل لساعة - ساعتين).

طريقة التشغيل:
    pip install requests beautifulsoup4 pandas openpyxl lxml --break-system-packages
    python sfda_new_drugs_scraper.py
"""

import time
import sys
import os
import json
import re
import threading
from datetime import datetime
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://www.sfda.gov.sa"
LIST_URL_TEMPLATE = f"{BASE_URL}/en/drugs-list?page={{page}}"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
}

LIST_DELAY_SECONDS = 0.6
DETAILS_WORKERS = 25
DETAILS_REQUEST_TIMEOUT = 15
MAX_EMPTY_RETRIES = 2
FETCH_DETAILS = True  # لو حبيت توقف مرحلة التفاصيل مؤقتًا، خليها False

_thread_local = threading.local()


def get_thread_session() -> requests.Session:
    """كل Thread بياخد نسخة requests.Session خاصة بيه (أأمن من مشاركة نفس الـ session بين كذا Thread)."""
    if not hasattr(_thread_local, "session"):
        _thread_local.session = requests.Session()
    return _thread_local.session


def get_total_pages(session: requests.Session):
    resp = session.get(LIST_URL_TEMPLATE.format(page=0), headers=HEADERS, timeout=30)
    resp.raise_for_status()
    page_numbers = [int(n) for n in re.findall(r"[?&]page=(\d+)", resp.text)]
    total = max(page_numbers) + 1 if page_numbers else 1
    return total, resp.text


def parse_list_page(html: str) -> list:
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return []

    rows = []
    body_rows = table.find_all("tr")[1:]
    for tr in body_rows:
        cells = tr.find_all("td")
        if len(cells) < 6:
            continue

        details_link_tag = cells[5].find("a")
        details_url = details_link_tag["href"] if details_link_tag else ""

        rows.append({
            "ScientificName": cells[0].get_text(strip=True),
            "TradeName": cells[1].get_text(strip=True),
            "Strength": cells[2].get_text(strip=True),
            "DosageForm": cells[3].get_text(strip=True),
            "Price": cells[4].get_text(strip=True),
            "DetailsURL": details_url,
        })

    return rows


def scrape_list() -> list:
    session = requests.Session()

    print("بنحدد إجمالي عدد الصفحات...")
    total_pages, first_page_html = get_total_pages(session)
    print(f"[i] إجمالي عدد الصفحات: {total_pages}")

    all_rows = []
    empty_retries = 0

    for page in range(0, total_pages):
        if page == 0:
            html = first_page_html
        else:
            try:
                resp = session.get(LIST_URL_TEMPLATE.format(page=page), headers=HEADERS, timeout=30)
                resp.raise_for_status()
                html = resp.text
            except requests.RequestException as e:
                print(f"[!] خطأ في صفحة {page}: {e} - هنجرب تاني بعد ثانيتين")
                time.sleep(2)
                try:
                    resp = session.get(LIST_URL_TEMPLATE.format(page=page), headers=HEADERS, timeout=30)
                    resp.raise_for_status()
                    html = resp.text
                except requests.RequestException as e2:
                    print(f"[x] فشلنا تاني في صفحة {page}: {e2}. هنوقف هنا.")
                    break

        rows = parse_list_page(html)

        if not rows:
            empty_retries += 1
            print(f"[i] صفحة {page} رجعت فاضية ({empty_retries}/{MAX_EMPTY_RETRIES})")
            if empty_retries >= MAX_EMPTY_RETRIES:
                print("[✓] يبدو إننا وصلنا لآخر صفحة فعلية، هنوقف السحب.")
                break
            time.sleep(LIST_DELAY_SECONDS)
            continue

        empty_retries = 0
        all_rows.extend(rows)

        if page % 50 == 0 or page == total_pages - 1:
            print(f"[+] صفحة {page + 1}/{total_pages}: الإجمالي لحد دلوقتي {len(all_rows)} دواء")

        time.sleep(LIST_DELAY_SECONDS)

    return all_rows


def parse_details_page(html: str) -> dict:
    """يستخرج كل أزواج (اسم الحقل، القيمة) من جدول صفحة التفاصيل."""
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    details = {}
    if not table:
        return details

    for tr in table.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if len(cells) == 2:
            label = cells[0].get_text(strip=True)
            value = cells[1].get_text(strip=True)
            if label:
                details[label] = value

    return details


def fetch_one_detail(row: dict) -> tuple:
    """يجيب تفاصيل دواء واحد. بيرجّع (نجح؟, الصف بعد التحديث)."""
    details_url = row.get("DetailsURL", "")
    if not details_url:
        return True, row

    full_url = urljoin(BASE_URL, details_url)
    session = get_thread_session()
    try:
        resp = session.get(full_url, headers=HEADERS, timeout=DETAILS_REQUEST_TIMEOUT)
        resp.raise_for_status()
        details = parse_details_page(resp.text)
        row.update(details)
        return True, row
    except requests.RequestException:
        return False, row


def enrich_with_details(rows: list) -> list:
    total = len(rows)
    print(f"\nبنجيب تفاصيل كل دواء لوحده ({total} دواء) - {DETAILS_WORKERS} طلبات متوازية...")

    fail_count = 0
    done_count = 0

    with ThreadPoolExecutor(max_workers=DETAILS_WORKERS) as executor:
        futures = [executor.submit(fetch_one_detail, row) for row in rows]

        for future in as_completed(futures):
            ok, _ = future.result()
            done_count += 1
            if not ok:
                fail_count += 1

            if done_count % 500 == 0 or done_count == total:
                print(f"[+] {done_count}/{total} (فشل حتى الآن: {fail_count})")

    if fail_count:
        print(f"[!] فشلنا في جلب تفاصيل {fail_count} دواء من إجمالي {total} (سيبناهم بالبيانات الأساسية بس).")

    return rows


def main():
    print("بدأنا نسحب بيانات الأدوية من الموقع الجديد لـ SFDA...\n")

    rows = scrape_list()
    if not rows:
        print("[x] معندناش أي بيانات من القائمة خالص. راجع شكل الصفحة أو الحماية.")
        sys.exit(1)

    # إزالة التكرار قبل مرحلة التفاصيل (عشان ما نفتحش نفس الصفحة مرتين)
    seen = set()
    unique_rows = []
    for r in rows:
        key = r.get("DetailsURL") or (r.get("ScientificName"), r.get("TradeName"), r.get("Strength"))
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(r)

    if len(unique_rows) != len(rows):
        print(f"[i] شلنا {len(rows) - len(unique_rows)} صف مكرر من القائمة.")

    if FETCH_DETAILS:
        unique_rows = enrich_with_details(unique_rows)

    df = pd.DataFrame(unique_rows)

    output_dir = "data"
    os.makedirs(output_dir, exist_ok=True)

    today_str = datetime.now().strftime("%Y-%m-%d")
    dated_filename = os.path.join(output_dir, f"SFDA_Drugs_New_{today_str}.xlsx")
    latest_filename = os.path.join(output_dir, "SFDA_Drugs_New_latest.xlsx")

    df.to_excel(dated_filename, index=False, engine="openpyxl")
    df.to_excel(latest_filename, index=False, engine="openpyxl")

    metadata = {
        "date": today_str,
        "count": len(df),
        "datedFileName": f"SFDA_Drugs_New_{today_str}.xlsx",
    }
    with open(os.path.join(output_dir, "last_updated_new.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False)

    print(f"\n[✓] تم! اتحفظ ملفين:")
    print(f"    - نسخة بتاريخ اليوم: {dated_filename}")
    print(f"    - نسخة ثابتة: {latest_filename}")
    print(f"    عدد الأدوية: {len(df)}")
    print(f"    عدد الأعمدة: {len(df.columns)}")
    print(f"    الأعمدة: {list(df.columns)}")


if __name__ == "__main__":
    main()
