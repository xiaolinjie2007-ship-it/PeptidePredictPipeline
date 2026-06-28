#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import nodriver as uc
import asyncio, re, time, os, openpyxl
from config import ALLER_USER, ALLER_EMAIL, ALLER_PASS

EXCEL_FILE = r"D:\kanmao\桌面\无毒.xlsx"
LOGIN_URL = "https://www.ddg-pharmfac.net/allertop_test/accounts/login/"
PREDICT_URL = "https://www.ddg-pharmfac.net/allertop_test/"
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "allertop_cache.txt")
U, E, P = ALLER_USER, ALLER_EMAIL, ALLER_PASS

def log(msg):
    print(msg, flush=True)

def load_cache():
    c = {}
    if os.path.exists(CACHE_FILE):
        for line in open(CACHE_FILE):
            p = line.strip().split("\t")
            if len(p) == 2: c[p[0]] = p[1]
    return c

def save_cache(r):
    with open(CACHE_FILE, "w") as f:
        for s, v in sorted(r.items()):
            f.write(f"{s}\t{v}\n")
    log(f"  [CACHE] {len(r)} results saved")

async def wait_for_page(page, expect_title, timeout=60):
    """Wait until page title matches expected title (Cloudflare bypassed)"""
    for i in range(timeout // 3):
        await asyncio.sleep(3)
        try:
            t = await page.evaluate('document.title')
            if t and expect_title in t:
                return t
        except:
            pass
    return await page.evaluate('document.title')

async def do_login(page):
    """Login and return True if successful"""
    log("  Navigating to login...")
    await page.get(LOGIN_URL)
    t = await wait_for_page(page, "AllerTOP", timeout=60)
    log(f"  Login page: {t}")

    if "AllerTOP" not in t:
        log("  Cloudflare still blocking, waiting more...")
        await asyncio.sleep(30)
        t = await page.evaluate('document.title')
        log(f"  Login page (retry): {t}")
    if "AllerTOP" not in t:
        return False

    # Fill form
    await page.evaluate(f"""
    (function(){{
        document.getElementById('id_username').value='{U}';
        document.getElementById('id_email').value='{E}';
        document.getElementById('id_password').value='{P}';
    }})()
    """)
    await asyncio.sleep(1)

    u = await page.evaluate("document.getElementById('id_username').value")
    log(f"  Filled: u={u}")

    # Click submit
    await page.evaluate("document.querySelector('button[type=submit]').click()")
    await asyncio.sleep(12)

    c = await page.get_content()
    return "Log Out" in c

async def main():
    log("=" * 60)
    log("AllerTOP Batch Prediction")
    log("=" * 60)

    # Read Excel
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["digestion_results"]
    sc = None
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=c).value == "fragment_sequence":
            sc = c; break

    peptides = sorted(set(
        str(ws.cell(row=r, column=sc).value).strip().upper()
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=sc).value
        and len(str(ws.cell(row=r, column=sc).value).strip()) >= 6
    ))
    wb.close()

    log(f"Peptides >=6: {len(peptides)}")
    cache = load_cache()
    log(f"Cached: {len(cache)}")

    pending = [p for p in peptides if p not in cache]
    log(f"Pending: {len(pending)}")

    if not pending:
        log("All done, writing results...")
        _write_excel(load_cache())
        return

    # Browser
    log("\nStarting browser (headed)...")
    browser = await uc.start(headless=False)
    page = await browser.get(PREDICT_URL)

    # Wait for Cloudflare once
    t = await wait_for_page(page, "AllerTOP", timeout=90)
    log(f"Main page: {t}")
    if "AllerTOP" not in t:
        log("CRITICAL: Cloudflare not passing, aborting")
        browser.stop()
        return

    # Login
    if not await do_login(page):
        log("Login FAILED, retrying...")
        if not await do_login(page):
            log("Login FAILED again, aborting")
            browser.stop()
            return
    log("Login OK")

    # ---- Main prediction loop ----
    results = dict(cache)
    t0 = time.time()

    for i, pep in enumerate(pending):
        # Navigate to prediction form
        # Use page.get() first time, then go back after first
        if i == 0:
            await page.get(PREDICT_URL)
            await asyncio.sleep(5)
            # Verify we're on prediction page
            t = await page.evaluate('document.title')
            if "AllerTOP" not in t:
                await asyncio.sleep(15)
        else:
            # Click browser back to return to form from result page
            await page.evaluate("history.back()")
            await asyncio.sleep(3)

        # Check we're on prediction form
        content = await page.get_content()
        if 'textarea' not in content:
            log("  Form not found, reloading prediction page...")
            await page.get(PREDICT_URL)
            await asyncio.sleep(5)

        # Verify still logged in
        if 'Log In' in content and 'Log Out' not in content:
            log("  Session expired, re-logging in...")
            if await do_login(page):
                await page.get(PREDICT_URL)
                await asyncio.sleep(5)

        # Fill and submit
        try:
            await page.evaluate(
                f"document.querySelector('textarea[name=protein]').value = '{pep}'")
            await asyncio.sleep(0.3)
            await page.evaluate(
                "document.querySelector('button[type=submit]').click()")
            await asyncio.sleep(6)

            # Parse result
            content = await page.get_content()
            clean = re.sub(r'<(script|style|nav)[^>]*>.*?</\1>', '',
                           content, flags=re.DOTALL)
            bm = re.search(r'<body>(.*?)</body>', clean, re.DOTALL)
            text = re.sub(r'<[^>]+>', ' ', bm.group(1)) if bm else clean

            if re.search(r'Probable\s+NON[\s-]*ALLERGEN', text, re.IGNORECASE):
                result = "Non-allergen"
            elif re.search(r'Probable\s+ALLERGEN', text, re.IGNORECASE):
                result = "Allergen"
            else:
                result = "???"
                # Debug: print page text
                lines = [l.strip() for l in text.split('\n') if l.strip()]
                for l in lines[:30]:
                    log(f"    DEBUG: {l[:120]}")

        except Exception as ex:
            result = f"ERR:{ex}"

        results[pep] = result

        # Progress
        elapsed = time.time() - t0
        done = i + 1
        rate = done / elapsed if elapsed > 0 else 0
        eta = (len(pending) - done) / rate if rate > 0 else 0

        log(f"  [{done}/{len(pending)}] {pep[:22]:<22} -> {result:<14} "
            f"({rate*60:.0f}/min, ETA {eta/60:.0f}min)")

        # Save cache periodically
        if done % 1 == 0:
            save_cache(results)

    browser.stop()
    save_cache(results)
    _write_excel(results)
    log("DONE")


def _write_excel(results):
    log("\nWriting to Excel...")
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["digestion_results"]

    sc = None
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h == "fragment_sequence": sc = c
        if h == "Allergen_Prediction":
            ac = c

    if 'ac' not in dir():
        ac = None
    ac_local = None
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=c).value == "Allergen_Prediction":
            ac_local = c; break
    if not ac_local:
        ac_local = ws.max_column + 1
        ws.cell(row=1, column=ac_local, value="Allergen_Prediction")

    filled = 0
    for r in range(2, ws.max_row + 1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if len(s) >= 6 and s in results:
                ws.cell(row=r, column=ac_local, value=results[s])
                filled += 1

    wb.save(EXCEL_FILE)
    a = sum(1 for v in results.values() if v == "Allergen")
    na = sum(1 for v in results.values() if v == "Non-allergen")
    log(f"Rows: {filled}, Allergen: {a}, Non-allergen: {na}")
    log(f"Saved: {EXCEL_FILE}")
    wb.close()


if __name__ == "__main__":
    asyncio.run(main())
