# -*- coding: utf-8 -*-
"""
سكريبت سحب قائمة الأدوية من الموقع الجديد لـ SFDA (sfda.gov.sa)
==================================================================
منفصل تمامًا عن سكريبت الموقع القديم (sfda_drugs_scraper.py) - محدش بيأثر على التاني.

الفكرة:
  الموقع الجديد بيعرض جدول الأدوية جاهز جوه صفحة HTML عادية (SSR)،
  فمش محتاجين نداءات AJAX زي القديم - بس نفتح كل صفحة (?page=0,1,2...)
  ونستخرج الجدول منها.

طريقة التشغيل:
    pip install requests beautifulsoup4 pandas openpyxl lxml --break-system-packages
    python sfda_new_drugs_scraper.py
"""

import time
import sys
import os
import json
import re
from datetime import datetime

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

REQUEST_DELAY_SECONDS = 0.6
MAX_EMPTY_RETRIES = 2


def get_total_pages(session: requests.Session) -> int:
    """يقرأ رقم آخر صفحة من روابط الـ Pagination في أول صفحة."""
    resp = session.get(LIST_URL_TEMPLATE.format(page=0), headers=HEADERS, timeout=30)
    resp.raise_for_status()

    page_numbers = [int(n) for n in re.findall(r"[?&]page=(\d+)", resp.text)]
    total = max(page_numbers) + 1 if page_numbers else 1
    return total, resp.text


def parse_page(html: str) -> list:
    """يستخرج صفوف جدول الأدوية من محتوى الصفحة."""
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return []

    rows = []
    body_rows = table.find_all("tr")[1:]  # نتخطى صف العناوين
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


def scrape_all_drugs() -> pd.DataFrame:
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

        rows = parse_page(html)

        if not rows:
            empty_retries += 1
            print(f"[i] صفحة {page} رجعت فاضية ({empty_retries}/{MAX_EMPTY_RETRIES})")
            if empty_retries >= MAX_EMPTY_RETRIES:
                print("[✓] يبدو إننا وصلنا لآخر صفحة فعلية، هنوقف السحب.")
                break
            time.sleep(REQUEST_DELAY_SECONDS)
            continue

        empty_retries = 0
        all_rows.extend(rows)

        if page % 20 == 0 or page == total_pages - 1:
            print(f"[+] صفحة {page + 1}/{total_pages}: الإجمالي لحد دلوقتي {len(all_rows)} دواء")

        time.sleep(REQUEST_DELAY_SECONDS)

    if not all_rows:
        print("[x] معندناش أي بيانات خالص. راجع شكل الصفحة أو الحماية.")
        sys.exit(1)

    df = pd.DataFrame(all_rows)
    before = len(df)
    df = df.drop_duplicates(subset=["DetailsURL"]) if "DetailsURL" in df.columns else df.drop_duplicates()
    after = len(df)
    if before != after:
        print(f"[i] شلنا {before - after} صف مكرر.")

    return df


def main():
    print("بدأنا نسحب بيانات الأدوية من الموقع الجديد لـ SFDA...\n")
    df = scrape_all_drugs()

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
    print(f"    الأعمدة: {list(df.columns)}")


if __name__ == "__main__":
    main()
