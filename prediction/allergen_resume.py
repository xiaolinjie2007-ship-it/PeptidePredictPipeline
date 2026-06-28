#!/usr/bin/env python3
"""AllerTOP continuation - retry ERROR entries then finish remaining"""
import nodriver as uc
import asyncio, re, time, os, shutil, openpyxl
from config import ALLER_USER, ALLER_EMAIL, ALLER_PASS

WORK_DIR = r"D:\kanmao\桌面\数据库"
ALLERGEN_FILE = os.path.join(WORK_DIR, "过敏性预测.xlsx")
NONALLER_FILE = os.path.join(WORK_DIR, "无过敏性.xlsx")
CACHE_A = os.path.join(WORK_DIR, "allertop_cache.txt")
ALL_URL = "https://www.ddg-pharmfac.net/allertop_test/"
ALL_LOGIN = "https://www.ddg-pharmfac.net/allertop_test/accounts/login/"
U, E, P = ALLER_USER, ALLER_EMAIL, ALLER_PASS

def log(msg):
    print(msg, flush=True)

def load_cache():
    c = {}
    if os.path.exists(CACHE_A):
        for line in open(CACHE_A):
            p = line.strip().split("\t")
            if len(p) == 2: c[p[0]] = p[1]
    return c

def save_cache(r):
    with open(CACHE_A, "w") as f:
        for s, v in sorted(r.items()):
            f.write("%s\t%s\n" % (s, v))

async def main():
    # Load data
    cache = load_cache()
    log("Current cache: %d (Allergen=%d, Non-allergen=%d, ERROR=%d)" % (
        len(cache),
        sum(1 for v in cache.values() if v == "Allergen"),
        sum(1 for v in cache.values() if v == "Non-allergen"),
        sum(1 for v in cache.values() if v == "ERROR")))

    # Get list of peptides >=6 from 过敏性预测.xlsx
    wb = openpyxl.load_workbook(ALLERGEN_FILE)
    ws = wb[wb.sheetnames[0]]
    sc = next(c for c in range(1, ws.max_column+1) if ws.cell(row=1, column=c).value == "sequence")
    peptides_ge6 = sorted(set(
        str(ws.cell(row=r, column=sc).value).strip().upper()
        for r in range(2, ws.max_row+1)
        if ws.cell(row=r, column=sc).value
        and len(str(ws.cell(row=r, column=sc).value).strip()) >= 6
    ))
    wb.close()
    log("Total peptides >=6: %d" % len(peptides_ge6))

    # Determine pending: not in cache OR cached as ERROR
    pending = [p for p in peptides_ge6
               if p not in cache or cache[p] == "ERROR"]
    log("Pending (missing + ERROR): %d" % len(pending))

    if not pending:
        log("All done, writing final files...")
        _write_final()
        return

    # Browser
    browser = await uc.start(headless=False)
    page = await browser.get(ALL_URL)

    # Cloudflare
    for _ in range(30):
        await asyncio.sleep(3)
        if 'AllerTOP' in await page.evaluate('document.title'): break

    # Login
    await page.get(ALL_LOGIN)
    await asyncio.sleep(15)
    await page.evaluate("document.getElementById('id_username').value='%s'" % U)
    await page.evaluate("document.getElementById('id_email').value='%s'" % E)
    await page.evaluate("document.getElementById('id_password').value='%s'" % P)
    await page.evaluate("document.querySelector('button[type=submit]').click()")
    await asyncio.sleep(12)
    log("Login OK")

    # Predict - use page.get() each time (more reliable than history.back)
    results = dict(cache)
    for i, pep in enumerate(pending):
        try:
            # Always navigate directly to prediction page
            await page.get(ALL_URL)
            await asyncio.sleep(4)

            await page.evaluate(
                "document.querySelector('textarea[name=protein]').value = '%s'" % pep)
            await asyncio.sleep(0.3)
            await page.evaluate(
                "document.querySelector('button[type=submit]').click()")
            await asyncio.sleep(6)

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
                result = "ERROR"

        except Exception as ex:
            log("  EXCEPTION: %s" % ex)
            result = "ERROR"

        results[pep] = result
        cv = "OK" if result != "ERROR" else "ERR"
        log("  [%d/%d] [%s] %-28s -> %s" % (i+1, len(pending), cv, pep[:28], result))
        save_cache(results)

    browser.stop()

    # Write final
    _write_final()


def _write_final():
    cache = load_cache()
    log("\nWriting allergen results to 过敏性预测.xlsx...")
    wb = openpyxl.load_workbook(ALLERGEN_FILE)
    ws = wb[wb.sheetnames[0]]
    sc = next(c for c in range(1, ws.max_column+1) if ws.cell(row=1, column=c).value == "sequence")

    ac = ws.max_column + 1
    ws.cell(row=1, column=ac, value="Allergen_Prediction")

    filled = 0
    for r in range(2, ws.max_row+1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if s in cache:
                ws.cell(row=r, column=ac, value=cache[s])
                filled += 1
    wb.save(ALLERGEN_FILE)

    a = sum(1 for v in cache.values() if v == "Allergen")
    na = sum(1 for v in cache.values() if v == "Non-allergen")
    err = sum(1 for v in cache.values() if v == "ERROR")
    log("Allergen=%d, Non-allergen=%d, ERROR=%d, Filled=%d" % (a, na, err, filled))

    # Create 无过敏性.xlsx
    log("Creating 无过敏性.xlsx...")
    rows_del = [r for r in range(2, ws.max_row+1)
                if ws.cell(row=r, column=ac).value
                and str(ws.cell(row=r, column=ac).value).strip() in ("Allergen", "ERROR")]
    for r in reversed(rows_del):
        ws.delete_rows(r)
    wb.save(NONALLER_FILE)
    wb.close()
    log("Deleted %d rows, saved %s" % (len(rows_del), os.path.basename(NONALLER_FILE)))
    log("DONE")


if __name__ == "__main__":
    asyncio.run(main())
