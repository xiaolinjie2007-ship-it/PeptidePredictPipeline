"""
从 SATPdb 网站抓取所有细胞穿透肽 (Cell Penetrating Peptides)
URL: https://webs.iiitd.edu.in/raghava/satpdb/browse_sub1.php?token=brsub&type=penetrating
共 17 页，831 条记录
表格列: S-ID | Seq | NF | FC | SF | AI | PD (共 7 列)
"""
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

BASE_URL = "https://webs.iiitd.edu.in/raghava/satpdb/browse_sub1.php"
PARAMS = {"token": "brsub", "type": "penetrating"}
TOTAL_PAGES = 17
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_CSV = os.path.join(SCRIPT_DIR, "cpp_satpdb_all.csv")
OUTPUT_EXCEL = os.path.join(SCRIPT_DIR, "cpp_satpdb_all.xlsx")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

all_peptides = []

for page in range(1, TOTAL_PAGES + 1):
    params = {**PARAMS, "page": str(page)}
    print(f"Fetching page {page}/{TOTAL_PAGES}...", end=" ")

    try:
        resp = requests.get(BASE_URL, params=params, headers=HEADERS, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        print(f"FAILED: {e}")
        continue

    soup = BeautifulSoup(resp.text, "html.parser")
    table = soup.find("table")
    if not table:
        print("no table found")
        continue

    rows = table.find_all("tr")
    data_rows = 0
    for row in rows:
        cols = row.find_all("td")
        if len(cols) < 7:
            continue
        seq = cols[1].get_text(strip=True)
        peptide = {
            "S-ID":     cols[0].get_text(strip=True),
            "Sequence": seq,
            "Length":   len(seq),
            "NF":       cols[2].get_text(strip=True),
            "FC":       cols[3].get_text(strip=True),
            "SF":       cols[4].get_text(strip=True),
            "AI":       cols[5].get_text(strip=True),
            "PD":       cols[6].get_text(strip=True),
        }
        all_peptides.append(peptide)
        data_rows += 1

    print(f"{data_rows} records")
    time.sleep(0.3)

print(f"\nTotal: {len(all_peptides)} peptides")

df = pd.DataFrame(all_peptides)
df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
print(f"Saved CSV: {OUTPUT_CSV}")
df.to_excel(OUTPUT_EXCEL, index=False)
print(f"Saved Excel: {OUTPUT_EXCEL}")

print(f"\nFirst 5 preview:")
print(df.head().to_string())
print(f"\nSequence length distribution:")
print(df["Length"].describe())