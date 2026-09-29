# -*- coding: utf-8 -*-
"""
ربات کپی و ویرایش محصول در ووکامرس با Playwright
--------------------------------------------------
این اسکریپت:
  1. وارد پنل وردپرس می‌شود.
  2. برای هر ردیف در فایل اکسل که هنوز «ثبت‌شده» نیست:
     - وارد صفحه‌ی محصولِ الگو (نمونه/قبلی) می‌شود.
     - روی دکمه‌ی «رونوشت» کلیک می‌کند (یک تب جدید باز می‌شود).
     - در تب جدید، عنوان/اسلاگ/توضیحات را با داده‌های اکسل جایگزین می‌کند.
     - روی «انتشار» کلیک می‌کند.
     - ستون وضعیت را در اکسل به‌روزرسانی و فایل را ذخیره می‌کند.

⚠️ جاهایی که باید عصر، بعد از دیدن سایت واقعی، تنظیم کنی با «# TODO» مشخص شده‌اند.
سلکتورهای فعلی (متن دکمه‌ها و ...) حدس منطقی بر اساس ظاهر معمول وردپرس/ووکامرس هستند،
ولی تا وقتی HTML واقعی صفحه را نبینم قطعی نیستند.
"""

import asyncio
import re
import sys

import openpyxl
from playwright.async_api import async_playwright, Page, BrowserContext

# ---------------------------------------------------------------------------
# تنظیمات — این‌ها را با اطلاعات واقعی خودت پر کن
# ---------------------------------------------------------------------------

WP_LOGIN_URL = "https://example.com/wp-login.php"   # TODO: آدرس واقعی صفحه ورود
WP_USERNAME = "YOUR_USERNAME"                        # TODO
WP_PASSWORD = "YOUR_PASSWORD"                        # TODO

# لینک صفحه‌ی ویرایش محصول الگو (همانی که هر بار از رویش «رونوشت» می‌گیریم)
TEMPLATE_PRODUCT_EDIT_URL = "https://example.com/wp-admin/post.php?post=123&action=edit"  # TODO

EXCEL_PATH = "/home/claude/wp-bot/products.xlsx"     # TODO: مسیر فایل اکسل واقعی
STATUS_COLUMN = "H"      # ستون «وضعیت / ثبت‌شده» طبق عکسی که فرستادی
DONE_VALUE = "ثبت شده"
HEADER_ROW = 1            # اولین ردیف داده بعد از هدر (طبق عکس: ردیف ۱ هدر، ردیف ۲ اولین محصول)

HEADLESS = False   # اول False بذار تا با چشم ببینی چی می‌شه. بعد از تست، True کن تا سریع‌تر و بی‌صدا اجرا شود.
SLOW_MO_MS = 150   # هر اکشن کمی کند اجرا شود تا هم قابل دیدن باشد و هم روی سرور فشار نیاورد.


# ---------------------------------------------------------------------------
# توابع کمکی روی اکسل
# ---------------------------------------------------------------------------

def load_rows(path: str):
    """
    ردیف‌هایی که هنوز ستون وضعیت‌شان پر نشده را برمی‌گرداند.
    هر ردیف را به‌صورت دیکشنری برمی‌گرداند تا در کد اصلی راحت‌تر استفاده شود.
    ستون‌ها طبق عکسی که فرستادی:
        A: نام محصول   B: برند   C: مدل   D: فرمت
        E: شماره برد   F: ظرفیت  G: REV   H: وضعیت
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = []
    for row_idx in range(HEADER_ROW + 1, ws.max_row + 1):
        name = ws[f"A{row_idx}"].value
        if not name:
            continue  # ردیف خالی، رد شو
        status = ws[f"{STATUS_COLUMN}{row_idx}"].value
        if status == DONE_VALUE:
            continue  # قبلاً انجام شده
        rows.append({
            "row_idx": row_idx,
            "name": str(name).strip(),
            "brand": str(ws[f"B{row_idx}"].value or "").strip(),
            "model": str(ws[f"C{row_idx}"].value or "").strip(),
            "format": str(ws[f"D{row_idx}"].value or "").strip(),
            "board_no": str(ws[f"E{row_idx}"].value or "").strip(),
            "capacity": str(ws[f"F{row_idx}"].value or "").strip(),
            "rev": str(ws[f"G{row_idx}"].value or "").strip(),
        })
    return wb, ws, rows


def mark_done(wb, ws, row_idx: int, path: str, ok: bool = True):
    ws[f"{STATUS_COLUMN}{row_idx}"] = DONE_VALUE if ok else "خطا"
    wb.save(path)


def slugify(text: str) -> str:
    """تبدیل نام فارسی/انگلیسی به یک اسلاگ ساده. در صورت نیاز اصلاحش کن."""
    text = text.strip().lower()
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"[^\w\-]", "", text, flags=re.UNICODE)
    return text


# ---------------------------------------------------------------------------
# مراحل مرورگر
# ---------------------------------------------------------------------------

async def login(page: Page):
    await page.goto(WP_LOGIN_URL)
    # TODO: اگر کپچا یا تایید دومرحله‌ای دیدی، اینجا باید متوقف بشیم
    # و منتظر بمونیم تا خودت دستی حلش کنی (input() یا page.pause()).
    await page.fill("#user_login", WP_USERNAME)
    await page.fill("#user_pass", WP_PASSWORD)
    await page.click("#wp-submit")
    await page.wait_for_load_state("networkidle")

    # بررسی ساده اینکه لاگین موفق بوده
    if "wp-admin" not in page.url:
        print("⚠️ لاگین موفق نبود، بررسی کن که آیا کپچا/۲FA هست.")
        # برای دیباگ دستی، این خط را از کامنت خارج کن:
        # await page.pause()


async def open_duplicated_product(context: BrowserContext, page: Page) -> Page:
    """
    به صفحه‌ی محصول الگو می‌رود، روی «رونوشت» کلیک می‌کند
    و صفحه‌ی جدیدی که باز می‌شود (محصول کپی‌شده) را برمی‌گرداند.
    """
    await page.goto(TEMPLATE_PRODUCT_EDIT_URL)
    await page.wait_for_load_state("networkidle")

    # TODO: متن دقیق دکمه را با آنچه واقعاً می‌بینی جایگزین کن.
    # اگر لینک متنی ساده است:
    duplicate_link = page.get_by_text("رونوشت", exact=False)
    # اگر به‌جای متن، عنوان/تایتل انگلیسی دارد (مثلاً "Duplicate")، این خط را هم امتحان کن:
    # duplicate_link = page.locator("a[title*='Duplicate'], a:has-text('رونوشت')")

    async with context.expect_page() as new_page_info:
        await duplicate_link.click()
    new_page = await new_page_info.value
    await new_page.wait_for_load_state("networkidle")
    return new_page


async def fill_product_fields(new_page: Page, data: dict):
    """
    فیلدهای عنوان، اسلاگ و توضیحات را در صفحه‌ی محصولِ کپی‌شده پر می‌کند.
    """
    # --- عنوان محصول ---
    title_selector = "#title"  # TODO: در ادیتور کلاسیک این آی‌دی معمولاً درست است.
    # اگر Gutenberg (ادیتور بلوکی) است، احتمالاً باید از این استفاده کنی:
    # title_selector = "h1.editor-post-title__input, .wp-block-post-title"
    await new_page.fill(title_selector, data["name"])

    # --- اسلاگ (Permalink) ---
    # روال معمول وردپرس کلاسیک: کلیک روی دکمه‌ی "Edit" کنار پرمالینک، تایپ در اینپوت، بعد OK.
    try:
        await new_page.click("#edit-slug-box .edit-slug")  # TODO: سلکتور دقیق را چک کن
        await new_page.fill("#new-post-slug", slugify(data["name"]))
        await new_page.click("button.save")  # دکمه‌ی OK
    except Exception:
        print("ℹ️ نتوانستم اسلاگ را ویرایش کنم؛ شاید ساختار صفحه فرق دارد — بعداً دستی چک کن.")

    # --- توضیحات محصول ---
    # اینجا مهم‌ترین بخش برای تنظیم توسط خودت است.
    # اگر ادیتور کلاسیک (TinyMCE) است:
    try:
        # TinyMCE داخل یک iframe است
        frame = new_page.frame_locator("#content_ifr")
        body = frame.locator("#tinymce")
        old_html = await body.inner_html()
        new_html = build_new_description(old_html, data)
        await body.evaluate("(el, html) => { el.innerHTML = html; }", new_html)
    except Exception:
        print("ℹ️ ادیتور کلاسیک پیدا نشد؛ احتمالاً Gutenberg است — این بخش باید جداگانه نوشته شود.")
        # TODO: اگر Gutenberg بود، باید بلوک‌های پاراگراف را پیدا کنیم و متن‌شان را عوض کنیم،
        # مثلاً با page.locator(".wp-block-paragraph") و replace متن هر کدام.


def build_new_description(old_html: str, data: dict) -> str:
    """
    مقادیر قدیمی را در توضیحات با مقادیر جدید جایگزین می‌کند.
    فعلاً یک نمونه‌ی ساده‌ی find/replace نوشته‌ام — چون هنوز نمونه‌ی واقعی متن توضیحات
    (قبل/بعد) را ندارم. وقتی نمونه را فرستادی، دقیقش می‌کنم.

    فرض فعلی: توضیحات شامل خطوطی مثل زیر است:
        برند: DELL
        مدل: LATITUDE E5550
        فرمت: BIN
        شماره برد: LA-A911P
        ظرفیت: 8
        REV: 1
    و ما مقدار سمت راست هر ":" را با مقدار جدید عوض می‌کنیم.
    """
    replacements = {
        "برند": data["brand"],
        "مدل": data["model"],
        "فرمت": data["format"],
        "شماره برد": data["board_no"],
        "ظرفیت": data["capacity"],
        "REV": data["rev"],
    }
    new_html = old_html
    for label, value in replacements.items():
        # جایگزینی الگوی "label: هرچیزی تا آخر خط/تگ"
        pattern = re.compile(rf"({re.escape(label)}\s*:\s*)([^<\n]*)")
        new_html = pattern.sub(lambda m: m.group(1) + value, new_html)
    return new_html


async def publish(new_page: Page):
    # TODO: متن/آی‌دی دکمه‌ی انتشار را چک کن.
    publish_button = new_page.locator("#publish")
    await publish_button.click()
    await new_page.wait_for_load_state("networkidle")


# ---------------------------------------------------------------------------
# اجرای اصلی
# ---------------------------------------------------------------------------

async def main():
    wb, ws, rows = load_rows(EXCEL_PATH)
    if not rows:
        print("همه‌ی ردیف‌ها قبلاً ثبت شده‌اند یا اکسل خالی است.")
        return

    print(f"{len(rows)} محصول برای ثبت پیدا شد.")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO_MS)
        context = await browser.new_context()
        page = await context.new_page()

        await login(page)

        for data in rows:
            print(f"در حال پردازش: {data['name']} ...")
            try:
                new_page = await open_duplicated_product(context, page)
                await fill_product_fields(new_page, data)
                await publish(new_page)
                await new_page.close()
                mark_done(wb, ws, data["row_idx"], EXCEL_PATH, ok=True)
                print(f"✅ ثبت شد: {data['name']}")
            except Exception as e:
                print(f"❌ خطا در ردیف {data['row_idx']} ({data['name']}): {e}")
                mark_done(wb, ws, data["row_idx"], EXCEL_PATH, ok=False)

        await browser.close()

    print("پایان کار.")


if __name__ == "__main__":
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    asyncio.run(main())
