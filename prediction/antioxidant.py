#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AnOxPePred 1.0 Antioxidant Activity Prediction Pipeline
========================================================
Submit peptides in FASTA format via file upload (Peptide Mode)
Parse FRS and CHEL scores, write results to Excel
"""
import requests, re, time, os, sys, shutil, openpyxl

WORK_DIR = r"D:\kanmao\桌面\数据库"
SRC_FILE = os.path.join(WORK_DIR, "无过敏性.xlsx")
AO_PRED_FILE = os.path.join(WORK_DIR, "抗氧化预测.xlsx")
AO_05_FILE = os.path.join(WORK_DIR, "抗氧化活性大于等于0.5.xlsx")

URL = "https://services.healthtech.dtu.dk/cgi-bin/webface2.cgi"
MAX_PEPTIDE_LEN = 30
MIN_PEPTIDE_LEN = 2

def log(msg):
    print(msg, flush=True)


def submit_fasta(fasta_text, max_retries=3):
    """Submit FASTA to AnOxPePred, return dict of {peptide: {frs, chel}}"""
    for attempt in range(max_retries):
        try:
            resp = requests.post(URL, files={
                'configfile': (None,
                    '/var/www/services/services/AnOxPePred-1.0/webface.cf'),
                'SEQSUB': ('peptides.fasta', fasta_text, 'text/plain'),
                'P_LEN_MIN': (None, str(MIN_PEPTIDE_LEN)),
                'P_LEN_MAX': (None, str(MAX_PEPTIDE_LEN)),
                'TYPE': (None, 'peptide'),
            }, timeout=120)

            m = re.search(r'jobid=(\w+)', resp.text)
            if not m:
                log("  ERROR: No job ID, attempt %d" % (attempt + 1))
                continue
            jid = m.group(1)
            log("  Job ID: %s" % jid)

            # Poll for results (max ~15 min)
            for i in range(90):
                time.sleep(10)
                r2 = requests.get("%s?jobid=%s" % (URL, jid), timeout=60)

                # Check for failure
                if 'failed' in r2.text.lower():
                    log("  Job failed, retrying...")
                    break

                # Check for results
                if 'FRS score' in r2.text:
                    return _parse_results(r2.text)

                if i % 6 == 0:
                    log("  Waiting... (%ds)" % ((i+1)*10))

            log("  Timeout on attempt %d" % (attempt + 1))

        except Exception as e:
            log("  Exception on attempt %d: %s" % (attempt + 1, e))
            time.sleep(10)

    return {}


def _parse_results(html):
    """Parse AnOxPePred output: extract FRS and CHEL scores per peptide"""
    results = {}

    # Extract FRS section
    frs_match = re.search(
        r'# FRS score\s*\n(.*?)(?:\n\s*\n|\n#|\Z)',
        html, re.DOTALL)
    if frs_match:
        for line in frs_match.group(1).strip().split('\n'):
            parts = line.strip().split('\t')
            if len(parts) >= 2:
                try:
                    score = float(parts[0].strip())
                    seq = parts[1].strip().upper()
                    if seq and score >= 0:
                        results[seq] = {'frs': score, 'chel': 0}
                except ValueError:
                    pass

    # Extract CHEL section
    chel_match = re.search(
        r'# CHEL score\s*\n(.*?)(?:\n\s*\n|\n#|\Z)',
        html, re.DOTALL)
    if chel_match:
        for line in chel_match.group(1).strip().split('\n'):
            parts = line.strip().split('\t')
            if len(parts) >= 2:
                try:
                    score = float(parts[0].strip())
                    seq = parts[1].strip().upper()
                    if seq and score >= 0:
                        if seq not in results:
                            results[seq] = {'frs': 0, 'chel': 0}
                        results[seq]['chel'] = score
                except ValueError:
                    pass

    return results


def main():
    log("=" * 60)
    log("AnOxPePred 1.0 Antioxidant Activity Prediction")
    log("=" * 60)

    # Step 1: Read peptides
    log("\n[1] Reading %s..." % os.path.basename(SRC_FILE))
    wb = openpyxl.load_workbook(SRC_FILE)
    ws = wb[wb.sheetnames[0]]
    sc = next(c for c in range(1, ws.max_column + 1)
              if ws.cell(row=1, column=c).value == "sequence")
    log("  Peptide column: %d" % sc)

    # Get all unique peptides within length limits
    peptides_ok = set()
    peptides_too_long = set()
    for r in range(2, ws.max_row + 1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if len(s) > MAX_PEPTIDE_LEN:
                peptides_too_long.add(s)
            elif len(s) >= MIN_PEPTIDE_LEN:
                peptides_ok.add(s)

    log("  Peptides 2-%d: %d, >%d: %d" %
        (MAX_PEPTIDE_LEN, len(peptides_ok), MAX_PEPTIDE_LEN, len(peptides_too_long)))

    peptides_list = sorted(peptides_ok)
    log("  Total to predict: %d" % len(peptides_list))

    # Step 2: Submit in batches of 50 (server limit)
    import math
    BATCH_SIZE = 50
    batches = [peptides_list[i:i+BATCH_SIZE] for i in range(0, len(peptides_list), BATCH_SIZE)]
    log("\n[2] Submitting %d batches (up to %d peptides each)..." %
        (len(batches), BATCH_SIZE))

    results = {}
    for bi, batch in enumerate(batches, 1):
        fasta_lines = [""">%d\n%s""" % (j+1, s) for j, s in enumerate(batch)]
        fasta_text = "\n".join(fasta_lines)
        log("  Batch %d/%d: %d peptides..." % (bi, len(batches), len(batch)))
        r = submit_fasta(fasta_text)
        if r:
            results.update(r)
            log("    Got %d results (total: %d)" % (len(r), len(results)))
        else:
            log("    WARNING: Batch %d returned no results!" % bi)
        if bi < len(batches):
            time.sleep(3)

    # Check for missing
    missing = [p for p in peptides_list if p not in results]
    if missing:
        log("\n  %d peptides missing, retrying..." % len(missing))
        for bi in range(0, len(missing), BATCH_SIZE):
            batch = missing[bi:bi+BATCH_SIZE]
            fasta = "\n".join(""">%d\n%s""" % (j+1, s) for j, s in enumerate(batch))
            r = submit_fasta(fasta)
            if r:
                results.update(r)
            time.sleep(3)

    # Step 3: Copy source -> 抗氧化预测.xlsx and write results
    log("\n[3] Creating %s..." % os.path.basename(AO_PRED_FILE))
    shutil.copy2(SRC_FILE, AO_PRED_FILE)
    wb = openpyxl.load_workbook(AO_PRED_FILE)
    ws = wb[wb.sheetnames[0]]

    # Add columns
    bc = ws.max_column + 1
    for i, h in enumerate(["AO_FRS_Score", "AO_CHEL_Score", "AO_Max_Score"]):
        ws.cell(row=1, column=bc + i, value=h)

    filled = 0
    for r in range(2, ws.max_row + 1):
        s = ws.cell(row=r, column=sc).value
        if s:
            s = str(s).strip().upper()
            if s in results:
                frs = results[s]['frs']
                chel = results[s]['chel']
                ws.cell(row=r, column=bc, value=frs)
                ws.cell(row=r, column=bc + 1, value=chel)
                ws.cell(row=r, column=bc + 2, value=max(frs, chel))
                filled += 1

    wb.save(AO_PRED_FILE)
    log("  Filled %d rows" % filled)

    # Statistics
    max_scores = {p: max(v['frs'], v['chel']) for p, v in results.items()}
    ge05 = sum(1 for s in max_scores.values() if s >= 0.5)
    lt05 = sum(1 for s in max_scores.values() if s < 0.5)
    log("  Score >= 0.5: %d, < 0.5: %d" % (ge05, lt05))

    # Step 4: Create 抗氧化活性大于等于0.5.xlsx
    log("\n[4] Creating %s (only score >= 0.5)..." %
        os.path.basename(AO_05_FILE))
    mc = bc + 2  # AO_Max_Score column
    rows_del = [r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=mc).value is None
                or float(ws.cell(row=r, column=mc).value) < 0.5]
    for r in reversed(rows_del):
        ws.delete_rows(r)
    wb.save(AO_05_FILE)
    wb.close()
    log("  Deleted %d rows, saved %s" % (len(rows_del),
        os.path.basename(AO_05_FILE)))

    log("\n" + "=" * 60)
    log("DONE - Files created:")
    log("  %s" % AO_PRED_FILE)
    log("  %s" % AO_05_FILE)
    log("=" * 60)


if __name__ == "__main__":
    main()
