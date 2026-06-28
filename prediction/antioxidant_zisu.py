#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AnOxPePred antioxidant prediction for zisu peptides"""
import requests, re, time, os, shutil, openpyxl

SRC = r"D:\kanmao\桌面\桌面\紫苏肽预防认知障碍虚拟筛选及其作用机制研究\筛选\无过敏性肽.xlsx"
WORK_DIR = os.path.dirname(SRC)
AO_FILE = os.path.join(WORK_DIR, "抗氧化预测.xlsx")
AO_05_FILE = os.path.join(WORK_DIR, "抗氧化活性大于等于0.5.xlsx")

URL = "https://services.healthtech.dtu.dk/cgi-bin/webface2.cgi"
BATCH_SIZE = 50

def log(msg):
    print(msg, flush=True)

def submit_batch(peptides):
    """Submit <=50 peptides to AnOxPePred, return dict {seq: {frs, chel}}"""
    fasta = "\n".join(">%d\n%s" % (j+1, s) for j, s in enumerate(peptides))
    resp = requests.post(URL, files={
        'configfile': (None, '/var/www/services/services/AnOxPePred-1.0/webface.cf'),
        'SEQSUB': ('p.fasta', fasta, 'text/plain'),
        'P_LEN_MIN': (None, '2'), 'P_LEN_MAX': (None, '30'),
        'TYPE': (None, 'peptide'),
    }, timeout=120)

    m = re.search(r'jobid=(\w+)', resp.text)
    if not m:
        return {}
    jid = m.group(1)

    for _ in range(90):
        time.sleep(10)
        r2 = requests.get("%s?jobid=%s" % (URL, jid), timeout=60)
        if 'FRS score' in r2.text:
            results = {}
            for sect in ['FRS', 'CHEL']:
                sm = re.search(r'# %s score\s*\n(.*?)(?:\n\s*\n|\n#|\Z)' % sect, r2.text, re.DOTALL)
                if sm:
                    for line in sm.group(1).strip().split('\n'):
                        parts = line.strip().split('\t')
                        if len(parts) >= 2:
                            try:
                                score = float(parts[0].strip())
                                seq = parts[1].strip().upper()
                                if sect == 'FRS':
                                    results.setdefault(seq, {})['frs'] = score
                                else:
                                    results.setdefault(seq, {})['chel'] = score
                            except ValueError:
                                pass
            return results
        if 'failed' in r2.text.lower():
            return {}
    return {}


def main():
    log("=" * 60)
    log("AnOxPePred - Zisu Peptides")
    log("=" * 60)

    # Read peptides
    wb = openpyxl.load_workbook(SRC)
    ws = wb[wb.sheetnames[0]]
    sc = next(c for c in range(1, ws.max_column + 1)
              if ws.cell(row=1, column=c).value == "fragment_sequence")
    log("Total rows: %d" % (ws.max_row - 1))

    peptides_set = sorted(set(
        str(ws.cell(row=r, column=sc).value).strip().upper()
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=sc).value
        and 2 <= len(str(ws.cell(row=r, column=sc).value).strip()) <= 30
    ))
    wb.close()
    log("Unique peptides 2-30: %d" % len(peptides_set))

    # Submit in batches
    batches = [peptides_set[i:i + BATCH_SIZE] for i in range(0, len(peptides_set), BATCH_SIZE)]
    log("Batches: %d" % len(batches))

    results = {}
    for bi, batch in enumerate(batches, 1):
        log("Batch %d/%d (%d peptides)..." % (bi, len(batches), len(batch)))
        r = submit_batch(batch)
        results.update(r)
        log("  Got %d results (total: %d)" % (len(r), len(results)))
        if bi < len(batches):
            time.sleep(3)

    # Retry missing
    missing = [p for p in peptides_set if p not in results]
    if missing:
        log("Retrying %d missing..." % len(missing))
        for bi in range(0, len(missing), BATCH_SIZE):
            batch = missing[bi:bi + BATCH_SIZE]
            r = submit_batch(batch)
            results.update(r)
            log("  Got %d more" % len(r))
            time.sleep(3)

    missing2 = [p for p in peptides_set if p not in results]
    if missing2:
        log("Retry 2: %d still missing..." % len(missing2))
        for bi in range(0, len(missing2), BATCH_SIZE):
            batch = missing2[bi:bi + BATCH_SIZE]
            r = submit_batch(batch)
            results.update(r)
            log("  Got %d more" % len(r))
            time.sleep(3)

    log("Final results: %d / %d" % (len(results), len(peptides_set)))

    # Write to 抗氧化预测.xlsx
    log("\nWriting to %s..." % os.path.basename(AO_FILE))
    shutil.copy2(SRC, AO_FILE)
    wb = openpyxl.load_workbook(AO_FILE)
    ws = wb[wb.sheetnames[0]]

    bc = ws.max_column + 1
    for i, h in enumerate(["AO_FRS_Score", "AO_CHEL_Score", "AO_Max_Score"]):
        ws.cell(row=1, column=bc + i, value=h)

    filled = 0
    for r in range(2, ws.max_row + 1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if s in results:
                res = results[s]
                frs = res.get('frs', 0)
                chel = res.get('chel', 0)
                ws.cell(row=r, column=bc, value=frs)
                ws.cell(row=r, column=bc + 1, value=chel)
                ws.cell(row=r, column=bc + 2, value=max(frs, chel))
                filled += 1

    wb.save(AO_FILE)
    ge05 = sum(1 for v in results.values() if max(v.get('frs', 0), v.get('chel', 0)) >= 0.5)
    lt05 = sum(1 for v in results.values() if max(v.get('frs', 0), v.get('chel', 0)) < 0.5)
    log("  Filled: %d, >=0.5: %d, <0.5: %d" % (filled, ge05, lt05))

    # Create 抗氧化活性大于等于0.5.xlsx
    log("\nCreating %s..." % os.path.basename(AO_05_FILE))
    mc = bc + 2
    rows_del = [r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=mc).value is None
                or float(ws.cell(row=r, column=mc).value) < 0.5]
    for r in reversed(rows_del):
        ws.delete_rows(r)
    wb.save(AO_05_FILE)
    wb.close()
    log("  Deleted %d rows, %d rows remain" % (len(rows_del), ws.max_row - 1))

    # Final missing check
    wbf = openpyxl.load_workbook(AO_FILE)
    wsf = wbf[wbf.sheetnames[0]]
    remaining = sum(1 for r in range(2, wsf.max_row + 1)
                    if wsf.cell(row=r, column=sc).value
                    and len(str(wsf.cell(row=r, column=sc).value)) <= 30
                    and wsf.cell(row=r, column=bc).value is None)
    wbf.close()
    log("\nFinal missing: %d" % remaining)

    log("\n" + "=" * 60)
    log("DONE")
    log("  %s" % AO_FILE)
    log("  %s" % AO_05_FILE)
    log("=" * 60)


if __name__ == "__main__":
    main()
