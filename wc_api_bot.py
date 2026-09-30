# -*- coding: utf-8 -*-
"""
ربات ساخت محصول جدید در ووکامرس بر اساس یک محصول الگو + داده‌های اکسل
-----------------------------------------------------------------------
روش کار (با REST API، بدون نیاز به باز کردن مرورگر):

  1. اطلاعات کامل محصول الگو (partdigi.com, post=15217) را با GET می‌گیریم:
     قیمت، دسته‌بندی، تصویر شاخص، وضعیت انتشار و ... را از همان کپی می‌کنیم.
  2. برای هر ردیف اکسل که هنوز «ثبت‌شده» نیست:
       - عنوان جدید را می‌سازیم.
       - توضیحات (description) و توضیح کوتاه (short_description) را
         از روی یک قالب ثابت (که از متن محصول الگو استخراج شده) با مقادیر
         جدید (برند/مدل/شماره‌برد/REV/ظرفیت/فرمت) پر می‌کنیم.
       - با POST یک محصول جدید در وضعیت "publish" می‌سازیم.
       - ستون وضعیت اکسل را به‌روزرسانی می‌کنیم.

نیازی به Playwright نیست؛ فقط `requests` و `openpyxl`.

نصب:
    pip install requests openpyxl
"""

import re
import sys
import time

import openpyxl
import requests

# ---------------------------------------------------------------------------
# تنظیمات — این‌ها را پر کن
# ---------------------------------------------------------------------------

SITE_URL = "https://partdigi.com"          # بدون اسلش انتهایی
CONSUMER_KEY = "ck_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"     # TODO: از تنظیمات ووکامرس
CONSUMER_SECRET = "cs_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"  # TODO

TEMPLATE_PRODUCT_ID = 15217   # همان محصول نمونه‌ای که فرستادی

EXCEL_PATH = "/home/claude/wp-bot/products.xlsx"   # TODO: مسیر واقعی فایل اکسل
STATUS_COLUMN = "H"
DONE_VALUE = "ثبت شده"
HEADER_ROW = 1   # ردیف ۱ = هدر، داده از ردیف ۲ شروع می‌شود

PUBLISH_STATUS = "publish"   # اگر اول می‌خوای پیش‌نویس بسازی برای تست، بذار "draft"
DELAY_BETWEEN_REQUESTS_SEC = 1.5   # برای فشار نیاوردن به هاست

API_BASE = f"{SITE_URL}/wp-json/wc/v3"


# ---------------------------------------------------------------------------
# قالب توضیحات — دقیقاً بر اساس متن محصول نمونه‌ای که فرستادی
# با کامنت‌های بلوکی گوتنبرگ (wp:paragraph / wp:heading / wp:list) نوشته شده
# تا وقتی محصول جدید را در ویرایشگر باز کنی، مثل محصول الگو به‌صورت بلوک‌های
# جدا (نه یک متن خام) نمایش داده شود.
# ---------------------------------------------------------------------------

DESCRIPTION_TEMPLATE = """<!-- wp:heading -->
<h2>فایل BIOS لپ‌تاپ {brand} {model} {board_no} {capacity}</h2>
<!-- /wp:heading -->

<!-- wp:paragraph -->
<p>فایل BIOS لپ‌تاپ <strong>{brand} {model}</strong> مخصوص مادربرد <strong>{board_no} Rev:{rev}</strong> و چیپ BIOS با ظرفیت <strong>{capacity}</strong> ارائه شده است. این فایل با فرمت {format} تهیه شده و برای پروگرام مستقیم آی‌سی BIOS توسط پروگرامرهای سازگار مورد استفاده قرار می‌گیرد.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>این فایل BIOS برای تعمیر و رفع مشکلات مرتبط با Firmware در لپ‌تاپ {brand} {model} قابل استفاده است؛ از جمله مشکلاتی مانند عدم بوت شدن دستگاه، خرابی یا اختلال در BIOS، مشکلات ایجادشده پس از پروگرام اشتباه و مواردی که نیاز به بازنویسی آی‌سی BIOS دارند.</p>
<!-- /wp:paragraph -->

<!-- wp:heading -->
<h2>مشخصات فایل BIOS</h2>
<!-- /wp:heading -->

<!-- wp:list -->
<ul>
<li>برند: {brand}</li>
<li>مدل: {model}</li>
<li>شماره برد: {board_no}</li>
<li>REV: {rev}</li>
<li>ظرفیت : {capacity}</li>
<li>فرمت: {format}</li>
<li>وضعیت: تست‌شده</li>
</ul>
<!-- /wp:list -->

<!-- wp:heading -->
<h2>نکته مهم قبل از خرید و پروگرام</h2>
<!-- /wp:heading -->

<!-- wp:paragraph -->
<p>قبل از خرید و پروگرام فایل BIOS {brand} {model} حتماً شماره برد {board_no}، نسخه REV {rev} و ظرفیت آی‌سی BIOS {capacity} را با مشخصات دستگاه خود مطابقت دهید. صرفاً یکسان بودن مدل لپ‌تاپ برای انتخاب فایل BIOS کافی نیست و استفاده از فایل نامناسب می‌تواند باعث عدم بوت شدن دستگاه یا ایجاد مشکلات دیگر شود.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>این BIOS {brand} {model} – {board_no} – {capacity} برای استفاده تعمیرکاران و تکنسین‌های تعمیرات لپ‌تاپ تهیه شده و فایل پس از خرید به‌صورت دانلود فوری در اختیار شما قرار می‌گیرد.</p>
<!-- /wp:paragraph -->"""

SHORT_DESCRIPTION_TEMPLATE = """<!-- wp:list -->
<ul>
<li>✓ شماره برد: {board_no}</li>
<li>✓ REV: {rev}</li>
<li>✓ حجم فایل: {capacity}</li>
<li>✓ نوع فایل: {format}</li>
<li>✓ تست‌شده</li>
<li>✓ دانلود فوری</li>
</ul>
<!-- /wp:list -->

<!-- wp:paragraph -->
<p>توجه: قبل از خرید فایل ، شماره برد، حجم چیپ BIOS و نسخه/REV برد دستگاه خود را با مشخصات فایل مطابقت دهید.</p>
<!-- /wp:paragraph -->"""

TITLE_TEMPLATE = "فایل BIOS لپ تاپ {brand} {model} {board_no} {capacity}"


def slugify(text: str) -> str:
    """
    اسلاگ لاتین ساده از روی برند/مدل/شماره‌برد می‌سازد، چون در اکسل معمولاً
    این مقادیر لاتین هستند (مثل Dell, INSPIRON 5520, LA-8241P).
    اگر می‌خوای اسلاگ فارسی هم قبول باشه (وردپرس اجازه می‌ده)، این تابع را
    ساده‌تر کن (فقط lower + جایگزینی فاصله با خط تیره).
    """
    text = text.strip().lower()
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"[^a-z0-9\-]", "", text)
    text = re.sub(r"-+", "-", text).strip("-")
    return text


# ---------------------------------------------------------------------------
# توابع کمکی اکسل
# ---------------------------------------------------------------------------

def load_rows(path: str):
    """
    ستون‌ها طبق فایل تو:
        A: نام محصول   B: برند   C: مدل   D: فرمت
        E: شماره برد   F: ظرفیت  G: REV   H: وضعیت
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = []
    for row_idx in range(HEADER_ROW + 1, ws.max_row + 1):
        name = ws[f"A{row_idx}"].value
        if not name:
            continue
        status = ws[f"{STATUS_COLUMN}{row_idx}"].value
        if status == DONE_VALUE:
            continue
        rows.append({
            "row_idx": row_idx,
            "name": str(name).strip(),
            "brand": str(ws[f"B{row_idx}"].value or "").strip(),
            "model": str(ws[f"C{row_idx}"].value or "").strip(),
            "format": str(ws[f"D{row_idx}"].value or "").strip(),
            "board_no": str(ws[f"E{row_idx}"].value or "").strip(),
            "capacity": str(ws[f"F{row_idx}"].value or "").strip(),
            "rev": str(ws[f"G{row_idx}"].value or "").strip(),
            # TODO: اگه ستون جدیدی برای لینک فایل دانلودی هر محصول اضافه کردی
            # (مثلاً ستون I)، این خط را باز کن:
            # "download_url": str(ws[f"I{row_idx}"].value or "").strip(),
        })
    return wb, ws, rows


def mark_status(wb, ws, row_idx: int, path: str, text: str):
    ws[f"{STATUS_COLUMN}{row_idx}"] = text
    wb.save(path)


# ---------------------------------------------------------------------------
# توابع WooCommerce API
# ---------------------------------------------------------------------------

def get_template_product():
    url = f"{API_BASE}/products/{TEMPLATE_PRODUCT_ID}"
    resp = requests.get(url, auth=(CONSUMER_KEY, CONSUMER_SECRET), timeout=30)
    resp.raise_for_status()
    return resp.json()


def build_new_product_payload(template: dict, data: dict) -> dict:
    """
    از محصول الگو، فیلدهای عمومی (قیمت، دسته‌بندی، تصویر و ...) را کپی می‌کند
    و فیلدهای متنی (عنوان/توضیحات) را با مقادیر تازه پر می‌کند.
    """
    fmt_args = dict(
        brand=data["brand"],
        model=data["model"],
        format=data["format"],
        board_no=data["board_no"],
        capacity=data["capacity"],
        rev=data["rev"],
    )

    new_title = data["name"] or TITLE_TEMPLATE.format(**fmt_args)

    payload = {
        "name": new_title,
        "slug": slugify(f"{data['brand']}-{data['model']}-{data['board_no']}-{data['capacity']}-file-bios"),
        "description": DESCRIPTION_TEMPLATE.format(**fmt_args),
        "short_description": SHORT_DESCRIPTION_TEMPLATE.format(**fmt_args),
        "status": PUBLISH_STATUS,
        "type": template.get("type", "simple"),
        "downloadable": template.get("downloadable", False),
        "virtual": template.get("virtual", False),
        "regular_price": template.get("regular_price", ""),
        "sale_price": template.get("sale_price", ""),
        "manage_stock": template.get("manage_stock", False),
        "stock_status": template.get("stock_status", "instock"),
        # Rank Math SEO: کلمه کلیدی اصلی را هم‌نام عنوان جدید می‌کنیم.
        # اگه می‌خوای عنوان/توضیحات متای سئو هم جدا تنظیم بشه، همین‌جا اضافه کن:
        # {"key": "rank_math_title", "value": new_title},
        # {"key": "rank_math_description", "value": ...}
        "meta_data": [
            {"key": "rank_math_focus_keyword", "value": new_title},
        ],
    }

    # دسته‌بندی‌ها را عیناً از محصول الگو کپی کن (فقط id لازم است)
    if template.get("categories"):
        payload["categories"] = [{"id": c["id"]} for c in template["categories"]]

    # تصویر شاخص را هم کپی کن (در صورت وجود)
    if template.get("images"):
        payload["images"] = [{"id": img["id"]} for img in template["images"] if img.get("id")]

    # --- فایل دانلودی محصول ---
    # فعلاً فایل دانلودی محصول الگو عیناً کپی می‌شود (یعنی محصول جدید موقتاً
    # فایل *غلط* خواهد داشت) تا وقتی مشخص کنی فایل هر محصول از کجا می‌آید.
    # وقتی تصمیم گرفتی (ستون اکسل جدید یا هر روش دیگر)، این بخش را با
    # data["download_url"] جایگزین کن، مثلاً:
    #
    # if data.get("download_url"):
    #     payload["downloads"] = [{"name": new_title, "file": data["download_url"]}]
    # else:
    #     payload["downloads"] = template.get("downloads", [])
    #
    if template.get("downloads"):
        payload["downloads"] = template["downloads"]  # TODO: جایگزین با فایل واقعی هر محصول

    return payload


def create_product(payload: dict) -> dict:
    url = f"{API_BASE}/products"
    resp = requests.post(url, json=payload, auth=(CONSUMER_KEY, CONSUMER_SECRET), timeout=30)
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# اجرای اصلی
# ---------------------------------------------------------------------------

def main():
    wb, ws, rows = load_rows(EXCEL_PATH)
    if not rows:
        print("همه‌ی ردیف‌ها قبلاً ثبت شده‌اند یا اکسل خالی است.")
        return

    print(f"{len(rows)} محصول برای ساخت پیدا شد.")
    template = get_template_product()
    print(f"محصول الگو با موفقیت خوانده شد: {template.get('name')}")

    for data in rows:
        print(f"در حال ساخت: {data['name']} ...")
        try:
            payload = build_new_product_payload(template, data)
            created = create_product(payload)
            link = created.get("permalink", "")
            print(f"✅ ساخته شد: {created.get('name')}  →  {link}")
            mark_status(wb, ws, data["row_idx"], EXCEL_PATH, DONE_VALUE)
        except Exception as e:
            print(f"❌ خطا در ردیف {data['row_idx']} ({data['name']}): {e}")
            mark_status(wb, ws, data["row_idx"], EXCEL_PATH, "خطا")

        time.sleep(DELAY_BETWEEN_REQUESTS_SEC)

    print("پایان کار.")


if __name__ == "__main__":
    main()
