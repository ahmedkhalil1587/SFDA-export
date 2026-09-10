# -*- coding: utf-8 -*-
"""
سكريبت سحب قائمة الأدوية البشرية من موقع الهيئة العامة للغذاء والدواء (SFDA)
=================================================================
الفكرة:
  1) نفتح صفحة الأدوية الأول عشان ناخد الكوكيز اللي محتاجها السيرفر (حماية WAF).
  2) نستخدم نفس الكوكيز عشان نعمل نداء لـ GetDrugs.php صفحة صفحة (page=1,2,3...).
  3) نجمع كل الصفوف في جدول واحد، ونشيل أي تكرار (بالاعتماد على رقم التسجيل).
  4) نصدّر النتيجة في ملف Excel باسم فيه تاريخ اليوم.

طريقة التشغيل:
    pip install requests pandas openpyxl --break-system-packages   (لو لينكس)
    أو  pip install requests pandas openpyxl                        (ويندوز/ماك)
    python sfda_drugs_scraper.py

ملاحظات مهمة:
  - السكريبت لازم يتشغل من جهاز عنده اتصال إنترنت عادي (المتصفح بتاعك يوصل
    للموقع، فجهازك أصلاً هيوصله برضه).
  - لو الموقع غيّر شكل الحماية أو شكل الـ endpoint، السكريبت هيطبع رسالة خطأ
    واضحة تقولك المشكلة فين.
  - شغّل السكريبت مرة كل شهر وهيطلعلك ملف جديد بتاريخ اليوم، وتقدر تقارنه
    بالشهر اللي فات بسهولة.
"""

import json
import time
import sys
import os
from datetime import datetime

import requests
import pandas as pd

BASE_URL = "https://oldsfda.sfda.gov.sa"
LIST_PAGE_URL = f"{BASE_URL}/ar/drugs-list"
API_URL = f"{BASE_URL}/GetDrugs.php"

# نفس الـ headers اللي شفناها في المتصفح، عشان السيرفر يعاملنا كطلب طبيعي
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
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Origin": BASE_URL,
    "Referer": LIST_PAGE_URL,
    "X-Requested-With": "XMLHttpRequest",
}

REQUEST_DELAY_SECONDS = 0.7  # تأخير بسيط بين كل طلب والتاني عشان مانتحظرش
MAX_PAGES = 3000  # سقف أمان بس، الحلقة هتوقف لوحدها قبل كده لو خلصت البيانات
MAX_EMPTY_RETRIES = 2  # لو صفحة رجعت فاضية، نجرب تاني قبل ما نعتبرها آخر صفحة


def start_session() -> requests.Session:
    """يفتح جلسة، يزور صفحة الأدوية الأول عشان ياخد كوكيز الحماية (WAF)."""
    session = requests.Session()
    resp = session.get(LIST_PAGE_URL, headers=COMMON_HEADERS, timeout=30)
    resp.raise_for_status()
    print(f"[✓] فتحنا صفحة الأدوية بنجاح (status={resp.status_code})، الكوكيز اتحطت.")
    return session


def extract_rows(payload) -> list:
    """
    المسار الحقيقي اللي اكتشفناه من فحص الرد الفعلي:
    payload["data"]["result"]["results"]
    مع fallback عام لو الشكل اتغيّر لأي سبب.
    """
    try:
        results = payload["data"]["result"]["results"]
        if isinstance(results, list):
            return results
    except (KeyError, TypeError):
        pass

    if isinstance(payload, list):
        return payload

    if isinstance(payload, dict):
        common_keys = ["data", "items", "result", "results", "records", "aaData", "rows", "Data"]
        for key in common_keys:
            if key in payload and isinstance(payload[key], list):
                return payload[key]

        list_values = [v for v in payload.values() if isinstance(v, list)]
        if list_values:
            return max(list_values, key=len)

    return []


def extract_page_count(payload):
    try:
        return payload["data"]["result"]["pageCount"]
    except (KeyError, TypeError):
        return None


def fetch_page(session: requests.Session, page: int):
    data = {
        "TradeName": "",
        "scientificName": "",
        "Agent": "",
        "ManufacturerName": "",
        "RegNo": "",
        "page": page,
    }
    resp = session.post(API_URL, headers=API_HEADERS, data=data, timeout=30)
    resp.raise_for_status()

    try:
        payload = resp.json()
    except json.JSONDecodeError:
        print(f"[!] الرد في صفحة {page} مش JSON صحيح. أول 300 حرف من الرد:")
        print(resp.text[:300])
        return [], None

    return extract_rows(payload), extract_page_count(payload)


def scrape_all_drugs() -> pd.DataFrame:
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

    df = pd.json_normalize(all_rows)

    # نحاول نشيل التكرار بالاعتماد على أي عمود اسمه فيه "reg" أو "RegNo"
    reg_col = next((c for c in df.columns if "reg" in c.lower()), None)
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
    print("بدأنا نسحب بيانات الأدوية من موقع SFDA...\n")
    df = scrape_all_drugs()

    output_dir = "data"
    os.makedirs(output_dir, exist_ok=True)

    today_str = datetime.now().strftime("%Y-%m-%d")
    dated_filename = os.path.join(output_dir, f"SFDA_Drugs_{today_str}.xlsx")
    latest_filename = os.path.join(output_dir, "SFDA_Drugs_latest.xlsx")

    df.to_excel(dated_filename, index=False, engine="openpyxl")
    df.to_excel(latest_filename, index=False, engine="openpyxl")

    metadata = {
        "date": today_str,
        "count": len(df),
        "datedFileName": f"SFDA_Drugs_{today_str}.xlsx",
    }
    with open(os.path.join(output_dir, "last_updated.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False)

    print(f"\n[✓] تم! اتحفظ ملفين:")
    print(f"    - نسخة بتاريخ اليوم: {dated_filename}")
    print(f"    - نسخة ثابتة (للرابط في الموقع): {latest_filename}")
    print(f"    عدد الأدوية: {len(df)}")
    print(f"    عدد الأعمدة: {len(df.columns)}")
    print(f"    أسماء الأعمدة: {list(df.columns)}")


if __name__ == "__main__":
    main()
