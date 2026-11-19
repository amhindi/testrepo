#!/usr/bin/env python3
"""
Boshamlan scraper v2 -> Excel

Rewritten against REAL markup confirmed on 2026-09-17:
  - Each result is <article class="post-card ..." data-post-id="277884">,
    not a real <a href>. Navigation happens via a JS click handler.
  - The card already carries the FULL title, price (as a numeric `value`
    attribute on a <data> tag), and the FULL untruncated description
    (the "2 line" look is CSS line-clamp, not missing text).
  - The one thing NOT on the card is the site's verified contact number,
    which only appears as a real <a href="tel:+965XXXXXXXX"> on the ad's
    own detail page. Getting that is the only reason we open each ad.

STRATEGY
  Phase 1 (fast): open the category page (e.g. /sale/houses), scroll and
    click "Load More" until the site says there's nothing left, harvesting
    every card's id / title / price / description / posted time directly
    from the list -- no per-ad page visits needed for this data.

  Phase 2 (slower, resumable): for every id collected in Phase 1, click
    that exact card (by its data-post-id, so position doesn't matter),
    read the phone number off the resulting ad page, then click back.
    Every result is written to a cache file as it completes, so a crash
    or Ctrl+C only costs you the ad in progress, not the whole run.

SETUP (once)
    pip3 install playwright pandas openpyxl
    python3 -m playwright install chromium

RUN
    python3 boshamlan_scraper.py --headful --limit 5     # sanity check first
    python3 boshamlan_scraper.py                          # full run, both types
    python3 boshamlan_scraper.py --types houses            # just houses
    python3 boshamlan_scraper.py --skip-phones             # phase 1 only, fast

If phase 2 (phone numbers) ever behaves oddly -- e.g. going "back" seems to
lose the loaded list -- the script detects that itself and reloads before
continuing, at a one-time cost, so it should self-heal rather than crash.
"""

import argparse
import asyncio
import json
import re
import sys
from pathlib import Path

try:
    from playwright.async_api import async_playwright
except ImportError:
    sys.exit("Missing playwright.  pip3 install playwright && python3 -m playwright install chromium")

import pandas as pd

BASE = "https://www.boshamlan.com"
TYPE_PATHS = {"houses": "/sale/houses", "lands": "/sale/lands"}

SEL = {
    "card": 'article.post-card[data-post-id]',
    "load_more": 'button:has-text("عرض المزيد"), button:has-text("Load More"), '
                 'a:has-text("عرض المزيد"), a:has-text("Load More")',
    "end_marker": 'text=/إنتهت نتائج البحث|No more search results/',
    "price_tag": 'data[value]',
    "title_tag": 'h3',
    "desc_tag": 'p.post-description, p[class*="post-description"], p[class*="description"]',
    "time_tag": 'time',
    "phone_link": 'a[href^="tel:"]',
}

AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
AR_LETTERS = str.maketrans({
    "أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا",
    "ة": "ه",
    "ى": "ي",
})


def norm(s):
    s = (s or "").translate(AR_DIGITS).translate(AR_LETTERS)
    s = s.replace("\u200f", "").replace("\u200e", "")
    s = re.sub(r"[\u064b-\u065f]", "", s)  # strip Arabic diacritics
    return s.strip()


def find_bare_area(blob, window=20):
    """Catch area written as a bare number like '٤٠٠ متر' with no
    'المساحة' keyword nearby, while skipping numbers that belong to a
    frontage/setback mention instead (e.g. 'الواجهة 21 متر')."""
    for m in re.finditer(r"([\d.,]+)\s*(?:متر|م)\b", blob):
        start, end = max(0, m.start() - window), min(len(blob), m.end() + window)
        context = blob[start:end]
        if re.search(r"ارتداد|واجهه|frontage|setback", context, re.I):
            continue
        try:
            return float(m.group(1).replace(",", ""))
        except ValueError:
            continue
    return None


def first(pattern, text, group=1, cast=None):
    m = re.search(pattern, text, re.I)
    if not m:
        return None
    v = m.group(group).replace(",", "").strip()
    if cast:
        try:
            return cast(v)
        except ValueError:
            return None
    return v


def has(pattern, text):
    return "نعم" if re.search(pattern, text, re.I) else ""


def parse_description(desc, title=""):
    """كل الأعمدة والقيم الناتجة هنا بالعربي، بما يطابق طلب المستخدم."""
    t = norm(desc)
    blob = f"{norm(title)} {t}"
    out = {}

    out["المساحة (م٢)"] = (
        first(r"(?:المساحه|مساحه|area)\D{0,12}?([\d.,]+)\s*(?:م|متر|m|sq)", blob, cast=float)
        or first(r"([\d.,]+)\s*(?:م٢|م2|متر مربع|sq\s?m|sqm)", blob, cast=float)
        or find_bare_area(blob)
    )
    out["الواجهة (م)"] = first(r"(?:الواجهه|واجهه|frontage)\D{0,10}?([\d.,]+)", blob, cast=float) \
        or first(r"([\d.,]+)\s*(?:م|متر|m)\w*\s*(?:واجهه|frontage)", blob, cast=float)
    out["الارتداد (م)"] = first(r"(?:ارتداد|frontage|setback)\D{0,10}?([\d.,]+)", blob, cast=float) \
        or first(r"([\d.,]+)\s*(?:م|متر|m)\w*\s*(?:ارتداد|setback)", blob, cast=float)
    out["رقم القطعة"] = first(r"(?:قطعه|ق|block|plot)\s*[:\-]?\s*(\d{1,3})\b", blob)

    out["عدد الأدوار"] = first(r"(\d+(?:\.\d+)?)\s*(?:ادوار|دور|floors?|storey)", blob, cast=float)
    out["عدد الغرف"] = first(r"(\d+)\s*(?:غرف|غرفه|bedrooms?|rooms?)", blob, cast=int)
    out["عدد الحمامات"] = first(r"(\d+)\s*(?:حمام|حمامات|bathrooms?)", blob, cast=int)
    out["عدد الشقق"] = first(r"(\d+)\s*(?:شقه|شقق|apartments?|flats?)", blob, cast=int)
    out["عدد المواقف"] = first(r"(\d+)\s*(?:سياره|سيارات|مصفط|مواقف|parking|car)", blob, cast=int)
    out["سنة البناء"] = first(r"(?:بنيان|بناء|built)\D{0,8}(19\d{2}|20\d{2})", blob, cast=int)

    out["السعر المذكور بالنص (د.ك)"] = first(r"(?:السعر|سعر البيع|البيع|price)\D{0,15}?([\d.,]{3,})", blob, cast=float)
    out["المدخول الشهري (د.ك)"] = first(r"(?:المدخول|الدخل|مؤجر[ه]?\s*ب?|income|rent)\D{0,15}?([\d.,]{3,})", blob, cast=float)
    out["قابل للتفاوض"] = has(r"قابل للتفاوض|قابل للمساوم|علي السوم|negotiab", blob)

    out["رقم الترخيص"] = first(r"(?:ترخيص|رخصه|license)\D{0,12}?(\d{3,4}\s*/\s*\d{3,4})", blob) \
        or first(r"(?:ترخيص|license)\D{0,12}?(\d{4,6})", blob)
    out["نوع الوثيقة"] = "حرة" if re.search(r"وثيقه حره|freehold", blob, re.I) else (
        "حكومية" if re.search(r"حكوم", blob) else "")

    street = []
    if re.search(r"زاويه|corner", blob, re.I): street.append("زاوية")
    if re.search(r"شارعين|two streets", blob, re.I): street.append("شارعين")
    if re.search(r"شارع واحد|one street|single street", blob, re.I): street.append("شارع واحد")
    if re.search(r"بطن وظهر", blob): street.append("بطن وظهر")
    if re.search(r"سكه|alley", blob, re.I): street.append("سكة")
    if re.search(r"شارع رئيسي|main street", blob, re.I): street.append("شارع رئيسي")
    out["موقع الشارع"] = " / ".join(street)
    out["إطلالة بحرية"] = has(r"بحر|اطلاله|سي فيو|sea view|seafront", blob)

    out["جديد / أول ساكن"] = has(r"جديد|اول ساكن|first occupancy|brand new", blob)
    out["مجدد"] = has(r"مجدد|تجديد|renovat", blob)
    out["مفروش"] = has(r"مفروش|مؤثث|furnish", blob)
    out["تكييف مركزي"] = has(r"تكييف مركزي|مركزي|central a", blob)
    out["يوجد مصعد"] = has(r"مصعد|elevator|lift", blob)
    out["يوجد سرداب"] = has(r"سرداب|basement", blob)
    out["من المالك مباشرة"] = has(r"من المالك|سكن المالك|from the owner|direct from owner", blob)
    out["يمنع المكاتب/الوسطاء"] = has(r"يمنع المكاتب|يمنع الوسطاء|ممنوع الوسطاء|no brokers", blob)
    out["نوع الإعلان"] = "مطلوب" if re.search(r"^مطلوب|مطلوب شراء|wanted|needed", blob, re.I) else "عرض للبيع"

    nums = re.findall(r"(?<!\d)([569]\d{7})(?!\d)", t)
    out["أرقام هواتف مذكورة بالوصف"] = " / ".join(dict.fromkeys(nums))
    return out


async def harvest_cards(page):
    """Pull id/title/price/description/posted straight from every loaded card."""
    return await page.eval_on_selector_all(
        "article.post-card[data-post-id]",
        """els => els.map(el => {
            const id = el.getAttribute('data-post-id');
            const h3 = el.querySelector('h3');
            const priceEl = el.querySelector('data[value]');
            const descEl = el.querySelector('p.post-description, p[class*="post-description"], p[class*="description"]');
            const timeEl = el.querySelector('time');
            return {
                id: id,
                title: h3 ? h3.textContent.trim() : '',
                price_value: priceEl ? priceEl.getAttribute('value') : '',
                price_text: priceEl ? priceEl.textContent.trim() : '',
                description: descEl ? descEl.textContent.trim() : '',
                posted_iso: timeEl ? timeEl.getAttribute('datetime') : '',
                posted_text: timeEl ? timeEl.textContent.trim() : '',
            };
        })"""
    )


async def load_all(page, url, limit, max_rounds=3000):
    print(f"\nOpening {url}")
    await page.goto(url, wait_until="domcontentloaded", timeout=60000)
    await page.wait_for_timeout(2500)

    seen = {}
    stagnant = 0
    for rnd in range(max_rounds):
        cards = await harvest_cards(page)
        before = len(seen)
        for c in cards:
            if c["id"]:
                seen[c["id"]] = c

        if limit and len(seen) >= limit:
            break
        if await page.locator(SEL["end_marker"]).count():
            print(f"  end-of-results marker reached ({len(seen)} ads)")
            break

        btn = page.locator(SEL["load_more"]).first
        clicked = False
        if await btn.count() and await btn.is_visible():
            try:
                await btn.click(timeout=5000)
                clicked = True
            except Exception:
                pass
        if not clicked:
            await page.mouse.wheel(0, 6000)

        await page.wait_for_timeout(1000)

        stagnant = stagnant + 1 if len(seen) == before else 0
        if stagnant >= 8:
            print(f"  no new ads in 8 rounds -- assuming list exhausted ({len(seen)} ads)")
            break
        if rnd % 10 == 0:
            print(f"  round {rnd}: {len(seen)} ads loaded")

    items = list(seen.values())
    return items[:limit] if limit else items


async def ensure_card_present(page, post_id, url):
    """After a back-navigation, make sure the target card is still there.
    If the SPA lost its scroll state, reload the list until it reappears."""
    sel = f'article.post-card[data-post-id="{post_id}"]'
    if await page.locator(sel).count():
        return True
    for _ in range(400):
        if await page.locator(sel).count():
            return True
        if await page.locator(SEL["end_marker"]).count():
            return False
        btn = page.locator(SEL["load_more"]).first
        if await btn.count() and await btn.is_visible():
            try:
                await btn.click(timeout=5000)
            except Exception:
                await page.mouse.wheel(0, 6000)
        else:
            await page.mouse.wheel(0, 6000)
        await page.wait_for_timeout(900)
    return False


async def fetch_phone(page, list_url, post_id, delay):
    sel = f'article.post-card[data-post-id="{post_id}"]'
    if not await ensure_card_present(page, post_id, list_url):
        return {"phone": "", "ad_url": "", "ok": False}

    before_url = page.url
    try:
        await page.locator(sel).first.click(timeout=8000)
        await page.wait_for_function(
            "u => location.href !== u", arg=before_url, timeout=8000
        )
    except Exception:
        return {"phone": "", "ad_url": "", "ok": False}

    await page.wait_for_timeout(400)
    ad_url = page.url
    hrefs = await page.eval_on_selector_all(
        SEL["phone_link"], "els => els.map(e => e.getAttribute('href'))"
    )
    nums = []
    for h in hrefs or []:
        m = re.findall(r"(\d{8,})", h or "")
        nums += [n[-8:] for n in m]
    phone = " / ".join(dict.fromkeys(nums))

    try:
        await page.go_back(wait_until="domcontentloaded", timeout=15000)
    except Exception:
        await page.goto(list_url, wait_until="domcontentloaded", timeout=30000)
    await page.wait_for_timeout(max(delay, 0.3) * 1000)

    return {"phone": phone, "ad_url": ad_url, "ok": True}


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--types", nargs="+", default=["houses", "lands"], choices=TYPE_PATHS.keys())
    ap.add_argument("--out", default="boshamlan_listings.xlsx")
    ap.add_argument("--cache", default="boshamlan_cache.jsonl")
    ap.add_argument("--limit", type=int, default=0, help="cap per property type (0 = all)")
    ap.add_argument("--delay", type=float, default=0.5)
    ap.add_argument("--skip-phones", action="store_true", help="phase 1 only, no per-ad visits")
    ap.add_argument("--headful", action="store_true")
    a = ap.parse_args()

    cache = Path(a.cache)
    done = {}
    if cache.exists():
        for line in cache.read_text(encoding="utf-8").splitlines():
            try:
                r = json.loads(line)
                done[r["id"]] = r
            except Exception:
                pass
        print(f"Cache: {len(done)} ads already have a phone lookup")

    all_rows = {}

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=not a.headful)
        ctx = await browser.new_context(
            locale="ar-KW",
            viewport={"width": 1400, "height": 1000},
            user_agent=("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"),
        )
        page = await ctx.new_page()

        for ptype in a.types:
            url = BASE + TYPE_PATHS[ptype]
            items = await load_all(page, url, a.limit)
            print(f"{ptype}: {len(items)} ads collected in phase 1")
            for it in items:
                it["property_type"] = ptype
                it["list_url"] = url
                all_rows[it["id"]] = it

            if a.skip_phones:
                continue

            todo = [it for it in items if it["id"] not in done]
            print(f"{ptype}: {len(todo)} need a phone lookup")
            with cache.open("a", encoding="utf-8") as fh:
                for i, it in enumerate(todo, 1):
                    res = await fetch_phone(page, url, it["id"], a.delay)
                    rec = {"id": it["id"], **res}
                    done[it["id"]] = rec
                    fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
                    fh.flush()
                    if i % 25 == 0:
                        print(f"  {ptype}: {i}/{len(todo)} phone lookups done")

        await browser.close()

    TYPE_AR = {"houses": "بيت", "lands": "أرض"}

    rows = []
    for post_id, it in all_rows.items():
        price_val = it.get("price_value") or ""
        price_val = re.sub(r"[^\d.]", "", price_val)
        rec = {
            "نوع العقار": TYPE_AR.get(it["property_type"], it["property_type"]),
            "الاسم / العنوان": it["title"],
            "السعر (د.ك)": float(price_val) if price_val else None,
            "السعر كما يظهر بالموقع": it["price_text"],
            "رقم التواصل": (done.get(post_id) or {}).get("phone", ""),
            "تاريخ النشر": it.get("posted_text", ""),
            "تاريخ النشر (ISO)": it.get("posted_iso", ""),
            "رقم الإعلان": post_id,
            "رابط الإعلان": (done.get(post_id) or {}).get("ad_url", ""),
            "رابط صفحة القائمة": it["list_url"],
        }
        rec.update(parse_description(it.get("description", ""), it.get("title", "")))
        rec["الوصف الكامل"] = it.get("description", "")
        rows.append(rec)

    df = pd.DataFrame(rows)
    lead = ["نوع العقار", "الاسم / العنوان", "السعر (د.ك)", "السعر كما يظهر بالموقع",
            "رقم التواصل", "تاريخ النشر", "نوع الإعلان"]
    tail = ["الوصف الكامل", "رقم الإعلان", "رابط الإعلان", "رابط صفحة القائمة", "تاريخ النشر (ISO)"]
    mid = [c for c in df.columns if c not in lead + tail]
    df = df[[c for c in lead + mid + tail if c in df.columns]]

    with pd.ExcelWriter(a.out, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="Listings", index=False)
        ws = xl.sheets["Listings"]
        ws.sheet_view.rightToLeft = True
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            letter = col[0].column_letter
            ws.column_dimensions[letter].width = 16
        for c in ws[1]:
            if c.value == "الوصف الكامل":
                ws.column_dimensions[c.column_letter].width = 100
            if c.value == "الاسم / العنوان":
                ws.column_dimensions[c.column_letter].width = 32
            if c.value in ("رابط الإعلان", "رابط صفحة القائمة"):
                ws.column_dimensions[c.column_letter].width = 40

    filled = df.notna().sum()
    print(f"\nWrote {len(df)} rows -> {a.out}")
    print("\nField fill rates:")
    for k in df.columns:
        n = int(filled.get(k, 0))
        print(f"  {k:<28} {n:>5} / {len(df)}")


if __name__ == "__main__":
    asyncio.run(main())
