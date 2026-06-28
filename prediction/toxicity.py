#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ToxinPred3 peptide toxicity prediction pipeline
Usage: python toxinpred_pipeline.py PEPTIDE1 PEPTIDE2 ...
       python toxinpred_pipeline.py peptides.fasta
"""
import subprocess
import re
import sys
import os
import time
import tempfile

TOXINPRED_URL = "https://webs.iiitd.edu.in/raghava/toxinpred3/prediction_action.php"
DISP_URL = "https://webs.iiitd.edu.in/raghava/toxinpred3/disp1.php"
THRESHOLD = "0.38"
MODEL = "1"  # Hybrid (ET+MERCI)
MAX_WAIT = 120
POLL_INTERVAL = 5

DEFAULT_PEPTIDES = ["AAAAAGGGGGGGGGGA"]


def make_fasta(peptides):
    lines = []
    for i, seq in enumerate(peptides, 1):
        lines.append(">Seq_%d" % i)
        lines.append(seq)
    return "\n".join(lines)


def submit(fasta_text):
    print("=" * 60)
    print("Step 1: Submitting peptides to ToxinPred3...")
    print("Model: Hybrid (ET+MERCI), Threshold: %s" % THRESHOLD)
    print("=" * 60)

    tmpfile = os.path.join(tempfile.gettempdir(), "toxinpred_input.fasta")
    with open(tmpfile, "w", encoding="utf-8") as f:
        f.write(fasta_text)

    cmd = [
        "curl", "-s",
        "-X", "POST", TOXINPRED_URL,
        "--form", "seq=<%s" % tmpfile,
        "--form", "terminus=%s" % MODEL,
        "--form", "th=%s" % THRESHOLD,
        "--form", "submit=Submit",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    if result.returncode != 0:
        print("curl failed: %s" % result.stderr)
        return None

    match = re.search(r"ran=(\d+)", result.stdout)
    if not match:
        print("Failed to extract Job ID")
        print("Response:", result.stdout[:500])
        return None

    ran = match.group(1)
    print("Job ID: %s" % ran)
    return ran


def fetch_result(ran):
    print("\nStep 2: Waiting for server to process (max %ds)..." % MAX_WAIT)

    elapsed = 0
    while elapsed < MAX_WAIT:
        time.sleep(POLL_INTERVAL)
        elapsed += POLL_INTERVAL

        cmd = ["curl", "-s", "%s?ran=%s" % (DISP_URL, ran)]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        if "<td>" in result.stdout:
            print("  [OK] Got results after %ds" % elapsed)
            return result.stdout

        bars = "|" * (elapsed // POLL_INTERVAL)
        print("  [%ds] %s" % (elapsed, bars), end="\r")

    print("\n  [FAIL] Timeout after %ds, no table data in page" % MAX_WAIT)
    return None


def parse_results(html):
    tbody_match = re.search(r"<tbody>(.*?)</tbody>", html, re.DOTALL)
    if not tbody_match:
        print("No <tbody> found in HTML")
        return []

    rows = re.findall(r"<tr>(.*?)</tr>", tbody_match.group(1), re.DOTALL)
    results = []
    for row in rows:
        cells = re.findall(r"<td>(.*?)</td>", row, re.DOTALL)
        if len(cells) >= 7:
            results.append({
                "subject": cells[0].strip(),
                "sequence": cells[1].strip(),
                "ml_score": float(cells[2].strip()),
                "merci_pos": float(cells[3].strip()),
                "merci_neg": float(cells[4].strip()),
                "hybrid_score": float(cells[5].strip()),
                "prediction": cells[6].strip(),
                "ppv": float(cells[7].strip()) if len(cells) >= 8 else 0,
            })
    return results


def output_results(all_results):
    toxins = [r for r in all_results if r["prediction"] == "Toxin"]
    non_toxins = [r for r in all_results if r["prediction"] == "Non-Toxin"]

    print("\n" + "=" * 70)
    print("                    ALL PREDICTION RESULTS")
    print("=" * 70)
    header = "%-5s %-24s %-14s %-12s %-8s" % ("#", "Sequence", "Hybrid Score", "Prediction", "PPV")
    print(header)
    print("-" * 70)
    for i, r in enumerate(all_results, 1):
        line = "%-5s %-24s %-14.3f %-12s %-8.3f" % (
            str(i), r["sequence"], r["hybrid_score"], r["prediction"], r["ppv"]
        )
        print(line)

    print("\n" + "=" * 70)
    print("  Non-Toxin peptides: %d" % len(non_toxins))
    print("=" * 70)
    if non_toxins:
        print("%-5s %-24s %-14s %-8s" % ("#", "Sequence", "Hybrid Score", "PPV"))
        print("-" * 70)
        for i, r in enumerate(non_toxins, 1):
            print("%-5s %-24s %-14.3f %-8.3f" % (str(i), r["sequence"], r["hybrid_score"], r["ppv"]))
    else:
        print("(none)")

    print("\nToxin peptides (%d): %s" % (len(toxins), ", ".join(r["sequence"] for r in toxins)))
    return non_toxins


def main():
    if len(sys.argv) > 1:
        peptides = []
        for arg in sys.argv[1:]:
            arg = arg.strip().upper()
            if os.path.isfile(arg):
                with open(arg) as f:
                    for line in f:
                        line = line.strip().upper()
                        if line and not line.startswith(">"):
                            peptides.append(line)
            else:
                peptides.append(arg)
    else:
        peptides = DEFAULT_PEPTIDES

    print("Input peptides (%d): %s" % (len(peptides), ", ".join(peptides)))
    fasta = make_fasta(peptides)
    print("\nFASTA input:\n%s\n" % fasta)

    ran = submit(fasta)
    if not ran:
        sys.exit(1)

    html = fetch_result(ran)
    if not html:
        sys.exit(1)

    results = parse_results(html)
    if not results:
        print("Parse failed, saving raw HTML to prediction_debug.html")
        with open("prediction_debug.html", "w", encoding="utf-8") as f:
            f.write(html)
        sys.exit(1)

    non_toxins = output_results(results)

    print("\n" + "=" * 70)
    print("  Non-Toxin sequences (copy-ready)")
    print("=" * 70)
    for r in non_toxins:
        print(r["sequence"])


if __name__ == "__main__":
    main()
