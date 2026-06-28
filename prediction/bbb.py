#!/usr/bin/env python3
"""
B3Pred BBB 穿透预测
读取 CSV 中的肽段序列，提交到 B3Pred 网站进行血脑屏障穿透预测，
结果输出到新的 Excel 文件。
"""
import os, json, csv, subprocess, zipfile

CONFIG = {
    "input_csv": r"D:\kanmao\桌面\6.23问了师兄\9概率校准模型\08_贮藏蛋白_活性肽清单.csv",
    "output_xlsx": r"D:\kanmao\桌面\6.23问了师兄\9概率校准模型\08_贮藏蛋白_BBB预测结果.xlsx",
    "node_script": r"c:\Users\kanmao\.claude\skills\b3pred_batch.js",
    "threshold": 0.5,
}

def read_peptides(csv_path):
    """从 CSV 读取肽段序列"""
    peptides = []
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            seq = row.get('Peptide', '').strip()
            if seq:
                peptides.append(seq)
    return peptides

def build_fasta(peptides):
    lines = []
    for i, seq in enumerate(peptides, 1):
        lines.append(f">{i}")
        lines.append(seq)
    return "\n".join(lines)

def run_b3pred(fasta_str, threshold):
    """调用 Node.js 脚本提交到 B3Pred"""
    tmp = os.path.join(os.path.dirname(CONFIG["node_script"]), "_tmp_b3pred_input.txt")
    with open(tmp, 'w', encoding='utf-8') as f:
        f.write(fasta_str)

    result = subprocess.run(
        ["node", CONFIG["node_script"], tmp, str(threshold)],
        capture_output=True, text=True, timeout=600,
        cwd=os.path.dirname(CONFIG["node_script"]),
    )
    out = result.stdout.strip()
    if out:
        data = json.loads(out)
        if data.get("success"):
            return data.get("results", [])
        raise Exception(data.get("error", "未知错误"))
    raise Exception(f"无输出: {result.stderr[:300]}")

def esc(text):
    return (str(text) if text is not None else '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

def col_letter(i):
    return chr(65 + i)

def make_xlsx(headers, rows, output_path):
    """内联字符串方式生成 xlsx"""
    NSW = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    NSCT = 'http://schemas.openxmlformats.org/package/2006/content-types'
    NSREL = 'http://schemas.openxmlformats.org/package/2006/relationships'
    NSOFF = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    row_xml = []
    row_xml.append(f'<row r="1">' + ''.join(
        f'<c r="{col_letter(ci)}1" t="inlineStr"><is><t xml:space="preserve">{esc(h)}</t></is></c>'
        for ci, h in enumerate(headers)) + '</row>')
    for ri, row in enumerate(rows):
        rn = ri + 2
        row_xml.append(f'<row r="{rn}">' + ''.join(
            f'<c r="{col_letter(ci)}{rn}" t="inlineStr"><is><t xml:space="preserve">{esc(val)}</t></is></c>'
            for ci, val in enumerate(row)) + '</row>')

    sheet_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="{NSW}">\n<sheetData>{"".join(row_xml)}</sheetData>\n</worksheet>'

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
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
            f'<Relationship Id="rId1" Type="{NSOFF}/officeDocument" Target="xl/workbook.xml"/>'
            f'</Relationships>')
        zf.writestr('xl/workbook.xml',
            f'<?xml version="1.0"?><workbook xmlns="{NSW}" xmlns:r="{NSOFF}">'
            f'<sheets><sheet name="BBB_Prediction" sheetId="1" r:id="rId1"/></sheets></workbook>')
        zf.writestr('xl/_rels/workbook.xml.rels',
            f'<?xml version="1.0"?><Relationships xmlns="{NSREL}">'
            f'<Relationship Id="rId1" Type="{NSOFF}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{NSOFF}/styles" Target="styles.xml"/>'
            f'</Relationships>')
        zf.writestr('xl/styles.xml',
            f'<?xml version="1.0"?><styleSheet xmlns="{NSW}">'
            f'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            f'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            f'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            f'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            f'</styleSheet>')
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)

def main():
    print("=" * 60)
    print("B3Pred BBB 穿透预测")
    print("=" * 60)

    peptides = read_peptides(CONFIG["input_csv"])
    print(f"\n[1] 读取肽段: {len(peptides)} 条")

    fasta = build_fasta(peptides)
    print(f"[2] FASTA: {len(fasta)} 字符")

    print(f"[3] 提交 B3Pred（阈值={CONFIG['threshold']}）...")
    results = run_b3pred(fasta, CONFIG["threshold"])
    print(f"  收到 {len(results)} 条结果")

    # 把 B3Pred 返回的 ID 映射回原始肽段序列（网站返回的 Seq 列为空）
    id_to_seq = {str(i+1): peptides[i] for i in range(len(peptides))}
    threshold = CONFIG["threshold"]

    headers = ["编号", "Peptide", "BBB_Score", "Prediction",
               "Hydrophobicity", "Hydropathicity", "Hydrophilicity", "Charge", "Mol_Wt"]
    rows = []
    for r in results:
        pid = r.get("id", "")
        seq = id_to_seq.get(pid, r.get("seq", ""))
        score_val = r.get("score", "0")
        try:
            pred = "B3P peptide" if float(score_val) >= threshold else "Non-B3P peptide"
        except (ValueError, TypeError):
            pred = ""
        rows.append([
            pid, seq, score_val, pred,
            r.get("hydrophobicity", ""), r.get("hydropathicity", ""),
            r.get("hydrophilicity", ""), r.get("charge", ""),
            r.get("mol_wt", ""),
        ])

    make_xlsx(headers, rows, CONFIG["output_xlsx"])
    print(f"[4] ✓ 保存到: {CONFIG['output_xlsx']}")
    print(f"    {len(rows)} 行数据")

if __name__ == "__main__":
    main()
