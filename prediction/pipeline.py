#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
四步完整流水线：抗氧化 → 过敏原 → 毒性
SRC: D:\kanmao\桌面\酶解.xlsx
OUT: D:\kanmao\桌面\ (7个溯源文件)
"""
import subprocess, requests, re, os, sys, time, tempfile, shutil, asyncio, openpyxl
import nodriver as uc
from config import ALLER_USER, ALLER_EMAIL, ALLER_PASS

DESKTOP = r"D:\kanmao\桌面"
SRC = os.path.join(DESKTOP, "酶解.xlsx")

# Step 1 files
F1 = os.path.join(DESKTOP, "活性大于等于0.5.xlsx")
# Step 2 files (AO)
F2_1 = os.path.join(DESKTOP, "抗氧化预测.xlsx")
F2_2 = os.path.join(DESKTOP, "抗氧化活性大于等于0.5.xlsx")
# Step 3 files (Allergen)
F3_1 = os.path.join(DESKTOP, "过敏性预测.xlsx")
F3_2 = os.path.join(DESKTOP, "无过敏性.xlsx")
# Step 4 files (Toxicity)
F4_1 = os.path.join(DESKTOP, "毒性预测.xlsx")
F4_2 = os.path.join(DESKTOP, "无毒.xlsx")

TOX_URL = "https://webs.iiitd.edu.in/raghava/toxinpred3/prediction_action.php"
TOX_DISP = "https://webs.iiitd.edu.in/raghava/toxinpred3/disp1.php"
AO_URL = "https://services.healthtech.dtu.dk/cgi-bin/webface2.cgi"
ALL_URL = "https://www.ddg-pharmfac.net/allertop_test/"
ALL_LOGIN = "https://www.ddg-pharmfac.net/allertop_test/accounts/login/"
ALL_U, ALL_E, ALL_P = ALLER_USER, ALLER_EMAIL, ALLER_PASS

CACHE_AO = os.path.join(DESKTOP, "cache_ao.txt")
CACHE_ALL = os.path.join(DESKTOP, "cache_allergen.txt")
CACHE_TOX = os.path.join(DESKTOP, "cache_tox.txt")

def log(msg): print(msg, flush=True)

# ============ CACHE HELPERS ============

def load_cache(path, n_cols=2):
    c = {}
    if os.path.exists(path):
        for l in open(path):
            p = l.strip().split("\t")
            if len(p) == n_cols: c[p[0]] = p[1:]
    return c

def save_cache(path, d, fmt_func):
    with open(path, "w") as f:
        for k in sorted(d):
            f.write(fmt_func(k, d[k]) + "\n")

# ============ STEP 1: Filter >=0.5 ============

def step1_filter():
    log("=" * 60)
    log("STEP 1: Filter activity >= 0.5")
    log("=" * 60)
    wb_src = openpyxl.load_workbook(SRC, read_only=True)
    ws_src = wb_src["Peptide_Details"]
    sc, pc = 5, 9

    # Create new workbook (much faster than deleting rows)
    wb_dst = openpyxl.Workbook()
    ws_dst = wb_dst.active
    ws_dst.title = "Peptide_Details"

    keep = 0
    for r, row in enumerate(ws_src.iter_rows(values_only=True), 1):
        if r == 1:
            # Copy header
            ws_dst.append(list(row))
        else:
            v = row[pc - 1] if len(row) >= pc else None
            if v is not None and float(v) >= 0.5:
                ws_dst.append(list(row))
                keep += 1
    wb_src.close()
    wb_dst.save(F1)
    wb_dst.close()
    log("  Kept %d rows, saved to %s" % (keep, os.path.basename(F1)))
    return keep

# ============ STEP 2: AnOxPePred ============

def step2_ao():
    log("\n" + "=" * 60)
    log("STEP 2: AnOxPePred Antioxidant Prediction")
    log("=" * 60)

    wb = openpyxl.load_workbook(F1)
    ws = wb["Peptide_Details"]
    sc = 5  # peptide
    peptides = sorted(set(
        str(ws.cell(row=r, column=sc).value).strip().upper()
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=sc).value
        and 2 <= len(str(ws.cell(row=r, column=sc).value).strip()) <= 30
    ))
    log("  Peptides 2-30: %d" % len(peptides))

    cache = load_cache(CACHE_AO, 3)
    pending = [p for p in peptides if p not in cache]
    log("  Cached: %d, Pending: %d" % (len(cache), len(pending)))

    if pending:
        BATCH = 50
        batches = [pending[i:i+BATCH] for i in range(0, len(pending), BATCH)]
        log("  Submitting %d batches..." % len(batches))
        for bi, batch in enumerate(batches, 1):
            fasta = "\n".join(">%d\n%s" % (j+1, s) for j, s in enumerate(batch))
            resp = requests.post(AO_URL, files={
                'configfile': (None, '/var/www/services/services/AnOxPePred-1.0/webface.cf'),
                'SEQSUB': ('p.fasta', fasta, 'text/plain'),
                'P_LEN_MIN': (None, '2'), 'P_LEN_MAX': (None, '30'),
                'TYPE': (None, 'peptide'),
            }, timeout=120)

            m = re.search(r'jobid=(\w+)', resp.text)
            if not m: log("    No job ID!"); continue
            jid = m.group(1)

            for _ in range(90):
                time.sleep(10)
                r2 = requests.get("%s?jobid=%s" % (AO_URL, jid), timeout=60)
                if 'FRS score' in r2.text:
                    for sect in ['FRS', 'CHEL']:
                        sm = re.search(r'# %s score\s*\n(.*?)(?:\n\s*\n|\n#|\Z)' % sect, r2.text, re.DOTALL)
                        if sm:
                            for line in sm.group(1).strip().split('\n'):
                                parts = line.strip().split('\t')
                                if len(parts) >= 2:
                                    try:
                                        s_val = float(parts[0].strip())
                                        seq = parts[1].strip().upper()
                                        key = 'frs' if sect == 'FRS' else 'chel'
                                        cache.setdefault(seq, (0, 0))
                                        frs, chel = cache[seq]
                                        if key == 'frs': frs = s_val
                                        else: chel = s_val
                                        cache[seq] = (frs, chel)
                                    except ValueError: pass
                    break
                if 'failed' in r2.text.lower(): break

            log("  Batch %d/%d: cached %d" % (bi, len(batches), len(cache)))
            save_cache(CACHE_AO, cache, lambda k, v: "%s\t%.6f\t%.6f" % (k, v[0], v[1]))
            if bi < len(batches): time.sleep(3)

        # Retry missing
        retry_missing = [p for p in pending if p not in cache]
        if retry_missing:
            log("  Retrying %d missing..." % len(retry_missing))
            for bi in range(0, len(retry_missing), BATCH):
                batch = retry_missing[bi:bi+BATCH]
                fasta = "\n".join(">%d\n%s" % (j+1, s) for j, s in enumerate(batch))
                resp = requests.post(AO_URL, files={
                    'configfile': (None, '/var/www/services/services/AnOxPePred-1.0/webface.cf'),
                    'SEQSUB': ('p.fasta', fasta, 'text/plain'),
                    'P_LEN_MIN': (None, '2'), 'P_LEN_MAX': (None, '30'),
                    'TYPE': (None, 'peptide'),
                }, timeout=120)
                m = re.search(r'jobid=(\w+)', resp.text)
                if not m: continue
                for _ in range(90):
                    time.sleep(10)
                    r2 = requests.get("%s?jobid=%s" % (AO_URL, m.group(1)), timeout=60)
                    if 'FRS score' in r2.text:
                        for sect in ['FRS', 'CHEL']:
                            sm = re.search(r'# %s score\s*\n(.*?)(?:\n\s*\n|\n#|\Z)' % sect, r2.text, re.DOTALL)
                            if sm:
                                for line in sm.group(1).strip().split('\n'):
                                    parts = line.strip().split('\t')
                                    if len(parts) >= 2:
                                        try:
                                            s_val = float(parts[0].strip())
                                            seq = parts[1].strip().upper()
                                            key = 'frs' if sect == 'FRS' else 'chel'
                                            cache.setdefault(seq, (0, 0))
                                            frs, chel = cache[seq]
                                            if key == 'frs': frs = s_val
                                            else: chel = s_val
                                            cache[seq] = (frs, chel)
                                        except ValueError: pass
                        break
                    if 'failed' in r2.text.lower(): break
                save_cache(CACHE_AO, cache, lambda k, v: "%s\t%.6f\t%.6f" % (k, v[0], v[1]))
                time.sleep(3)

    # Write F2_1: 抗氧化预测.xlsx
    log("\n  Writing %s..." % os.path.basename(F2_1))
    shutil.copy2(F1, F2_1)
    wb = openpyxl.load_workbook(F2_1)
    ws = wb["Peptide_Details"]
    bc = ws.max_column + 1
    for i, h in enumerate(["AO_FRS_Score", "AO_CHEL_Score", "AO_Max_Score"]):
        ws.cell(row=1, column=bc + i, value=h)
    filled = 0
    for r in range(2, ws.max_row + 1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if s in cache:
                frs, chel = cache[s]
                ws.cell(row=r, column=bc, value=frs)
                ws.cell(row=r, column=bc + 1, value=chel)
                ws.cell(row=r, column=bc + 2, value=max(frs, chel))
                filled += 1
    remaining = sum(1 for r in range(2, ws.max_row+1)
                    if ws.cell(row=r, column=sc).value
                    and len(str(ws.cell(row=r, column=sc).value)) <= 30
                    and ws.cell(row=r, column=bc).value is None)
    wb.save(F2_1)
    wb.close()
    log("  Filled: %d, Missing: %d" % (filled, remaining))

    # Create F2_2: 抗氧化活性大于等于0.5.xlsx
    log("  Creating %s..." % os.path.basename(F2_2))
    shutil.copy2(F2_1, F2_2)
    wb = openpyxl.load_workbook(F2_2)
    ws = wb["Peptide_Details"]
    mc = bc + 2
    rows_del = [r for r in range(2, ws.max_row+1)
                if ws.cell(row=r, column=mc).value is None
                or float(ws.cell(row=r, column=mc).value) < 0.5]
    for r in reversed(rows_del):
        ws.delete_rows(r)
    ao_ge05_count = ws.max_row - 1
    wb.save(F2_2)
    wb.close()
    log("  AO >=0.5 rows: %d" % ao_ge05_count)

    # Return peptides >=6 from AO >=0.5 for step 3
    wb = openpyxl.load_workbook(F2_2)
    ws = wb["Peptide_Details"]
    peptides_ge6 = sorted(set(
        str(ws.cell(row=r, column=sc).value).strip().upper()
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=sc).value
        and len(str(ws.cell(row=r, column=sc).value).strip()) >= 6
    ))
    wb.close()
    return peptides_ge6

# ============ STEP 3: AllerTOP ============

async def step3_allergen(peptides_ge6):
    log("\n" + "=" * 60)
    log("STEP 3: AllerTOP Allergenicity Prediction")
    log("=" * 60)
    log("  Peptides >=6 from AO>=0.5: %d" % len(peptides_ge6))

    if not peptides_ge6:
        log("  No peptides >=6, copying directly")
        shutil.copy2(F2_2, F3_1)
        shutil.copy2(F3_1, F3_2)
        wb = openpyxl.load_workbook(F2_2)
        ws = wb["Peptide_Details"]
        sc = 5
        all_peps = set(
            str(ws.cell(row=r, column=sc).value).strip().upper()
            for r in range(2, ws.max_row+1)
            if ws.cell(row=r, column=sc).value
            and len(str(ws.cell(row=r, column=sc).value).strip()) >= 6
        )
        wb.close()
        return list(all_peps)

    cache = load_cache(CACHE_ALL, 2)
    pending = [p for p in peptides_ge6 if p not in cache]
    log("  Cached: %d, Pending: %d" % (len(cache), len(pending)))

    results = dict(cache)
    if pending:
        log("  Starting AllerTOP browser...")
        browser = await uc.start(headless=False)
        page = await browser.get(ALL_URL)
        for _ in range(30):
            await asyncio.sleep(3)
            if 'AllerTOP' in await page.evaluate('document.title'): break

        # Login
        await page.get(ALL_LOGIN)
        await asyncio.sleep(15)
        await page.evaluate("document.getElementById('id_username').value='%s'" % ALL_U)
        await page.evaluate("document.getElementById('id_email').value='%s'" % ALL_E)
        await page.evaluate("document.getElementById('id_password').value='%s'" % ALL_P)
        await page.evaluate("document.querySelector('button[type=submit]').click()")
        await asyncio.sleep(12)
        log("  Login OK")

        t0 = time.time()
        for i, pep in enumerate(pending):
            if i == 0:
                await page.get(ALL_URL)
                await asyncio.sleep(5)
            else:
                await page.evaluate("history.back()")
                await asyncio.sleep(3)

            await page.evaluate("document.querySelector('textarea[name=protein]').value = '%s'" % pep)
            await page.evaluate("document.querySelector('button[type=submit]').click()")
            await asyncio.sleep(6)

            content = await page.get_content()
            clean = re.sub(r'<(script|style|nav)[^>]*>.*?</\1>', '', content, flags=re.DOTALL)
            bm = re.search(r'<body>(.*?)</body>', clean, re.DOTALL)
            text = re.sub(r'<[^>]+>', ' ', bm.group(1)) if bm else clean
            if re.search(r'Probable\s+NON[\s-]*ALLERGEN', text, re.IGNORECASE):
                result = "Non-allergen"
            elif re.search(r'Probable\s+ALLERGEN', text, re.IGNORECASE):
                result = "Allergen"
            else:
                result = "ERROR"
            results[pep] = result

            if (i + 1) % 5 == 0 or i < 3:
                elapsed = time.time() - t0
                done = i + 1
                rate = done / elapsed if elapsed > 0 else 0
                log("  [%d/%d] %-22s -> %-14s (%.1f/min)" % (done, len(pending), pep[:22], result, rate * 60))

            save_cache(CACHE_ALL, results, lambda k, v: "%s\t%s" % (k, v))

        browser.stop()

    # Write F3_1
    log("\n  Writing %s..." % os.path.basename(F3_1))
    shutil.copy2(F2_2, F3_1)
    wb = openpyxl.load_workbook(F3_1)
    ws = wb["Peptide_Details"]
    sc = 5
    ac = ws.max_column + 1
    ws.cell(row=1, column=ac, value="Allergen_Prediction")
    for r in range(2, ws.max_row + 1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if s in results:
                ws.cell(row=r, column=ac, value=results[s])
    wb.save(F3_1)
    na = sum(1 for v in results.values() if v == "Non-allergen")
    al = sum(1 for v in results.values() if v == "Allergen")
    log("  Allergen: %d, Non-allergen: %d" % (al, na))

    # Create F3_2 (non-allergen only)
    log("  Creating %s..." % os.path.basename(F3_2))
    rows_del = [r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=ac).value
                and str(ws.cell(row=r, column=ac).value).strip() == "Allergen"]
    for r in reversed(rows_del):
        ws.delete_rows(r)
    wb.save(F3_2)
    log("  Deleted %d allergen rows, %d rows remain" % (len(rows_del), ws.max_row - 1))

    # Return all peptides in F3_2 for toxicity step
    nonallergen_peps = set(
        str(ws.cell(row=r, column=sc).value).strip().upper()
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=sc).value
    )
    wb.close()
    return list(nonallergen_peps)

# ============ STEP 4: ToxinPred3 ============

def step4_toxicity(nonallergen_peps):
    log("\n" + "=" * 60)
    log("STEP 4: ToxinPred3 Toxicity Prediction")
    log("=" * 60)
    log("  Non-allergen peptides to predict: %d" % len(nonallergen_peps))

    if not nonallergen_peps:
        log("  No peptides to predict")
        return

    cache = load_cache(CACHE_TOX, 5)
    pending = [p for p in nonallergen_peps if p not in cache]
    log("  Cached: %d, Pending: %d" % (len(cache), len(pending)))

    if pending:
        BATCH = 500
        batches = [pending[i:i+BATCH] for i in range(0, len(pending), BATCH)]
        log("  Submitting %d batches..." % len(batches))
        for bi, batch in enumerate(batches, 1):
            fasta = "\n".join(">%d\n%s" % (j+1, s) for j, s in enumerate(batch, 1))
            tf = tempfile.mktemp(suffix=".fasta")
            with open(tf, "w") as f: f.write(fasta)

            r = subprocess.run([
                "curl", "-s", "-X", "POST", TOX_URL,
                "--form", "seq=<%s" % tf,
                "--form", "terminus=1", "--form", "th=0.38", "--form", "submit=Submit",
            ], capture_output=True, text=True, timeout=120)
            m = re.search(r"ran=(\d+)", r.stdout)
            if not m: log("    No Job ID"); continue
            ran = m.group(1)

            for _ in range(60):
                time.sleep(5)
                r2 = subprocess.run(
                    ["curl", "-s", "%s?ran=%s" % (TOX_DISP, ran)],
                    capture_output=True, text=True, timeout=60)
                if "<td>" in r2.stdout: break

            tbody = re.search(r"<tbody>(.*?)</tbody>", r2.stdout, re.DOTALL)
            if tbody:
                rows = re.findall(r"<tr>(.*?)</tr>", tbody.group(1), re.DOTALL)
                for row in rows:
                    cells = re.findall(r"<td>(.*?)</td>", row, re.DOTALL)
                    if len(cells) >= 7:
                        cache[cells[1].strip()] = (
                            cells[1].strip(), cells[2].strip(), cells[5].strip(),
                            cells[6].strip(),
                            cells[7].strip() if len(cells) >= 8 else "0")

            log("  Batch %d/%d: cached %d" % (bi, len(batches), len(cache)))
            save_cache(CACHE_TOX, cache, lambda k, v: "%s\t%s\t%s\t%s\t%s" % (v[0], v[1], v[2], v[3], v[4]))
            if bi < len(batches): time.sleep(3)

    # Write F4_1
    log("\n  Writing %s..." % os.path.basename(F4_1))
    shutil.copy2(F3_2, F4_1)
    wb = openpyxl.load_workbook(F4_1)
    ws = wb["Peptide_Details"]
    sc = 5
    tc = ws.max_column + 1
    for i, h in enumerate(["Toxicity_Prediction", "Toxicity_Hybrid_Score",
                            "Toxicity_ML_Score", "Toxicity_PPV"]):
        ws.cell(row=1, column=tc + i, value=h)

    for r in range(2, ws.max_row + 1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if s in cache:
                v = cache[s]
                ws.cell(row=r, column=tc, value=v[3])
                ws.cell(row=r, column=tc + 1, value=float(v[2]))
                ws.cell(row=r, column=tc + 2, value=float(v[1]))
                ws.cell(row=r, column=tc + 3, value=float(v[4]))
    wb.save(F4_1)

    toxin = sum(1 for v in cache.values() if v[3] == "Toxin")
    nontoxin = sum(1 for v in cache.values() if v[3] == "Non-Toxin")
    log("  Toxin: %d, Non-Toxin: %d" % (toxin, nontoxin))

    # Create F4_2
    log("  Creating %s..." % os.path.basename(F4_2))
    rows_del = [r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=tc).value
                and str(ws.cell(row=r, column=tc).value).strip() == "Toxin"]
    for r in reversed(rows_del):
        ws.delete_rows(r)
    wb.save(F4_2)
    log("  Deleted %d toxin rows, %d rows remain" % (len(rows_del), ws.max_row - 1))
    wb.close()


# ============ MAIN ============

async def main():
    # Step 1
    step1_filter()

    # Step 2
    peptides_for_allergen = step2_ao()

    # Step 3
    peptides_for_tox = await step3_allergen(peptides_for_allergen)

    # Step 4
    step4_toxicity(peptides_for_tox)

    log("\n" + "=" * 60)
    log("ALL DONE - Files created on Desktop:")
    for f in [F1, F2_1, F2_2, F3_1, F3_2, F4_1, F4_2]:
        log("  %s" % os.path.basename(f))
    log("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
