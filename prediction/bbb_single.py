#!/usr/bin/env python3
"""
B3Pred 单条提交方案：逐条提交到 B3Pred 获取真实 Score
因为 B3Pred 批量提交时 Score 列会丢失（bug）
"""
import os, json, subprocess, zipfile, csv, glob, time, re

CONFIG = {
    "input_csv": r"D:\kanmao\桌面\6.23问了师兄\9概率校准模型\08_贮藏蛋白_活性肽清单.csv",
    "output_xlsx": r"D:\kanmao\桌面\6.23问了师兄\9概率校准模型\08_贮藏蛋白_BBB预测结果.xlsx",
    "node_script": r"c:\Users\kanmao\.claude\skills\b3pred_single_submit.js",
    "tmp_dir": r"c:\Users\kanmao\.claude\skills\_b3pred_tmp",
    "threshold": 0.5,
}

def read_peptides(csv_path):
    peptides = []
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            seq = row.get('Peptide', '').strip()
            if seq:
                peptides.append(seq)
    return peptides

def esc(t):
    return (str(t) if t is not None else '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def make_xlsx(headers, rows, path):
    NSW = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    NSCT = 'http://schemas.openxmlformats.org/package/2006/content-types'
    NSREL = 'http://schemas.openxmlformats.org/package/2006/relationships'
    NSOFF = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    row_xml = []
    row_xml.append(f'<row r="1">' + ''.join(
        f'<c r="{chr(65+ci)}1" t="inlineStr"><is><t xml:space="preserve">{esc(h)}</t></is></c>'
        for ci, h in enumerate(headers)) + '</row>')
    for ri, row in enumerate(rows):
        rn = ri + 2
        row_xml.append(f'<row r="{rn}">' + ''.join(
            f'<c r="{chr(65+ci)}{rn}" t="inlineStr"><is><t xml:space="preserve">{esc(v)}</t></is></c>'
            for ci, v in enumerate(row)) + '</row>')

    sheet = f'<?xml version="1.0"?><worksheet xmlns="{NSW}"><sheetData>{"".join(row_xml)}</sheetData></worksheet>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml',
            f'<?xml version="1.0"?><Types xmlns="{NSCT}">'
            f'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            f'<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            f'</Types>')
        zf.writestr('_rels/.rels',
            f'<?xml version="1.0"?><Relationships xmlns="{NSREL}">'
            f'<Relationship Id="rId1" Type="{NSOFF}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        zf.writestr('xl/workbook.xml',
            f'<?xml version="1.0"?><workbook xmlns="{NSW}" xmlns:r="{NSOFF}">'
            f'<sheets><sheet name="BBB_Prediction" sheetId="1" r:id="rId1"/></sheets></workbook>')
        zf.writestr('xl/_rels/workbook.xml.rels',
            f'<?xml version="1.0"?><Relationships xmlns="{NSREL}">'
            f'<Relationship Id="rId1" Type="{NSOFF}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{NSOFF}/styles" Target="styles.xml"/></Relationships>')
        zf.writestr('xl/styles.xml',
            f'<?xml version="1.0"?><styleSheet xmlns="{NSW}">'
            f'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            f'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            f'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            f'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>')
        zf.writestr('xl/worksheets/sheet1.xml', sheet)

def main():
    os.makedirs(CONFIG["tmp_dir"], exist_ok=True)
    peptides = read_peptides(CONFIG["input_csv"])
    print(f"总计 {len(peptides)} 条肽段")

    results = {}
    for i, seq in enumerate(peptides):
        fasta = f">{i+1}\n{seq}"
        tmp = os.path.join(CONFIG["tmp_dir"], f"pep_{i+1}.txt")
        with open(tmp, 'w') as f:
            f.write(fasta)

        print(f"  [{i+1}/{len(peptides)}] {seq} ...", end=" ", flush=True)
        try:
            r = subprocess.run(
                ["node", CONFIG["node_script"], tmp, str(CONFIG["threshold"])],
                capture_output=True, text=True, timeout=120,
            )
            out = r.stdout.strip()
            if out:
                data = json.loads(out)
                if data.get("success"):
                    results[seq] = data
                    score = data.get("score", "?")
                    pred = data.get("prediction", "?")
                    print(f"Score={score} {pred}")
                else:
                    print(f"FAIL: {data.get('error','')}")
            else:
                print(f"NO OUTPUT")
        except subprocess.TimeoutExpired:
            print("TIMEOUT")
        except Exception as e:
            print(f"ERROR: {e}")

        time.sleep(1)

    # 生成 Excel
    print(f"\n生成 Excel...")
    headers = ["编号", "Peptide", "BBB_Score", "Prediction",
               "Hydrophobicity", "Hydropathicity", "Hydrophilicity", "Charge", "Mol_Wt"]
    rows = []
    for i, seq in enumerate(peptides):
        r = results.get(seq, {})
        score = r.get("score", "0")
        try:
            pred = "B3P peptide" if float(score) >= CONFIG["threshold"] else "Non-B3P peptide"
        except:
            pred = ""
        rows.append([
            str(i+1), seq, score, pred,
            r.get("hydrophobicity", ""), r.get("hydropathicity", ""),
            r.get("hydrophilicity", ""), r.get("charge", ""),
            r.get("mol_wt", ""),
        ])

    make_xlsx(headers, rows, CONFIG["output_xlsx"])
    print(f"✓ 保存到: {CONFIG['output_xlsx']}")

if __name__ == "__main__":
    main()
