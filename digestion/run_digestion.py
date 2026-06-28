#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从用户 Excel 文件读取蛋白序列，按 25 个既定酶解方案批量酶解。
使用 BIOPEP 网站 (https://biochemia.uwm.edu.pl/biopep/report_cutting_for_seq_v.php)
输出明细表 + 汇总表。
"""
import requests
import re
import time
import zipfile
import xml.etree.ElementTree as ET
import os
from datetime import datetime

# ============================================================
# 配置
# ============================================================
BASE_URL = "https://biochemia.uwm.edu.pl/biopep/report_cutting_for_seq_v.php"
INPUT_XLSX = r"D:\kanmao\桌面\uniprotkb_accsion_A0AAD4IMA6_OR_access_2026_06_05.xlsx"
OUTPUT_DIR = r"c:\Users\kanmao\.claude\skills"
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_DETAIL = os.path.join(OUTPUT_DIR, f"digestion_detail_{TIMESTAMP}.xlsx")
OUTPUT_SUMMARY = os.path.join(OUTPUT_DIR, f"digestion_summary_{TIMESTAMP}.xlsx")

# 酶名称
ENZYME_NAMES = {
    "41": "Subtilisin (Alkaline protease, EC 3.4.21.62)",
    "24": "Papain (EC 3.4.22.2)",
    "18": "Thermolysin (Neutral protease, EC 3.4.24.27)",
    "11": "Chymotrypsin A (EC 3.4.21.1)",
    "39": "Pepsin pH>2 (EC 3.4.23.1)",
}

# 25 个既定酶解方案
SCHEMES = {
    1:  ("Alkaline (Subtilisin)", ["41"]),
    2:  ("Papain", ["24"]),
    3:  ("Neutral (Thermolysin)", ["18"]),
    4:  ("Chymotrypsin A", ["11"]),
    5:  ("Pepsin", ["39"]),
    6:  ("Alkaline + Neutral", ["41", "18"]),
    7:  ("Alkaline + Chymotrypsin", ["41", "11"]),
    8:  ("Alkaline + Papain", ["41", "24"]),
    9:  ("Alkaline + Pepsin", ["41", "39"]),
    10: ("Neutral + Papain", ["18", "24"]),
    11: ("Neutral + Chymotrypsin", ["18", "11"]),
    12: ("Neutral + Pepsin", ["18", "39"]),
    13: ("Papain + Chymotrypsin", ["24", "11"]),
    14: ("Papain + Pepsin", ["24", "39"]),
    15: ("Pepsin + Chymotrypsin", ["39", "11"]),
    16: ("Alkaline + Neutral + Chymotrypsin", ["41", "18", "11"]),
    17: ("Alkaline + Neutral + Papain", ["41", "18", "24"]),
    18: ("Alkaline + Papain + Chymotrypsin", ["41", "24", "11"]),
    19: ("Alkaline + Neutral + Pepsin", ["41", "18", "39"]),
    20: ("Alkaline + Papain + Pepsin", ["41", "24", "39"]),
    21: ("Alkaline + Pepsin + Chymotrypsin", ["41", "39", "11"]),
    22: ("Neutral + Papain + Chymotrypsin", ["18", "24", "11"]),
    23: ("Neutral + Pepsin + Chymotrypsin", ["18", "39", "11"]),
    24: ("Neutral + Papain + Pepsin", ["18", "24", "39"]),
    25: ("Papain + Pepsin + Chymotrypsin", ["24", "39", "11"]),
}

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


# ============================================================
# 读取 Excel（处理内联字符串）
# ============================================================
def read_input_excel(path):
    """读取用户输入 Excel，返回 [(entry_name, sequence), ...]"""
    proteins = []
    with zipfile.ZipFile(path, 'r') as zf:
        # 读取共享字符串表
        ss = []
        try:
            sst = ET.parse(zf.open('xl/sharedStrings.xml'))
            for si in sst.findall(f'.//{NS}si'):
                t = si.find(f'{NS}t')
                if t is not None:
                    ss.append(t.text or '')
                else:
                    parts = []
                    for r_elem in si.findall(f'{NS}r'):
                        rt = r_elem.find(f'{NS}t')
                        if rt is not None and rt.text:
                            parts.append(rt.text)
                    ss.append(''.join(parts))
        except:
            pass

        def get_cell_value(c):
            t = c.attrib.get('t', '')
            is_el = c.find(f'{NS}is')
            if is_el is not None:
                t_el = is_el.find(f'{NS}t')
                if t_el is not None:
                    return t_el.text or ''
                parts = []
                for r_elem in is_el.findall(f'{NS}r'):
                    rt = r_elem.find(f'{NS}t')
                    if rt is not None:
                        parts.append(rt.text or '')
                return ''.join(parts)
            v_el = c.find(f'{NS}v')
            v = v_el.text if v_el is not None else ''
            if t == 's' and v.isdigit():
                idx = int(v)
                return ss[idx] if idx < len(ss) else v
            return v

        def col_letter(ref):
            return ''.join(filter(str.isalpha, ref))

        sheet = ET.parse(zf.open('xl/worksheets/sheet1.xml'))
        rows = sheet.findall(f'.//{NS}row')

        # 找头行确定列名
        header = {}
        if rows:
            for c in rows[0].findall(f'{NS}c'):
                ref = c.attrib.get('r', '')
                col = col_letter(ref)
                val = get_cell_value(c).strip().lower()
                header[col] = val

        # 确定 Entry 列和 Sequence 列
        entry_col = seq_col = None
        for col, name in header.items():
            if name in ('entry', 'accession', 'id', 'uniprot', 'protein_id'):
                entry_col = col
            if name in ('sequence', 'seq', 'protein_sequence'):
                seq_col = col

        # 宽匹配
        if entry_col is None:
            for col, name in header.items():
                if 'entry' in name or 'accession' in name or 'id' in name:
                    entry_col = col
                    break
        if seq_col is None:
            for col, name in header.items():
                if 'seq' in name:
                    seq_col = col
                    break

        print(f"Header: {header}")
        print(f"Entry column: {entry_col}, Sequence column: {seq_col}")

        # 读取数据行
        name_counts = {}
        for r in rows[1:]:
            cells = {}
            for c in r.findall(f'{NS}c'):
                ref = c.attrib.get('r', '')
                col = col_letter(ref)
                cells[col] = get_cell_value(c).strip()

            entry = cells.get(entry_col, '') if entry_col else ''
            seq = cells.get(seq_col, '') if seq_col else ''

            # 如果没有明确列，尝试找到最长的值作为序列
            if not seq and seq_col is None:
                for v in cells.values():
                    if len(v) > 50 and all(c in 'ACDEFGHIKLMNPQRSTVWY' for c in v.upper()):
                        seq = v
                        break

            if seq and len(seq) >= 10:
                # 处理重复名称
                if entry in name_counts:
                    name_counts[entry] += 1
                    display_name = f"{entry}_{name_counts[entry]}"
                else:
                    name_counts[entry] = 1
                    display_name = entry
                proteins.append((display_name, seq.upper()))

    return proteins


# ============================================================
# BIOPEP 请求
# ============================================================
def fetch_enzyme_result(session, protein_name, sequence, enzyme_ids):
    """向 BIOPEP 提交序列+酶组合，返回肽段列表、位置字符串和错误信息"""
    params = {
        "txt_seq": sequence,
        "enz1": enzyme_ids[0] if len(enzyme_ids) >= 1 else "",
        "enz2": enzyme_ids[1] if len(enzyme_ids) >= 2 else "",
        "enz3": enzyme_ids[2] if len(enzyme_ids) >= 3 else "",
    }
    try:
        resp = session.get(BASE_URL, params=params, timeout=120)
        resp.encoding = 'utf-8'
        html = resp.text
    except Exception as e:
        return None, None, f"Request error: {e}"

    # 解析肽段
    peptide_match = re.search(
        r'Results\s+of\s+enzyme\s+action.*?<td class="info"><font size="-1">\s*(.*?)\s*</font>',
        html, re.DOTALL
    )
    # 解析位置
    location_match = re.search(
        r'Location\s+of\s+released\s+peptides.*?<td class="info"><font size="-1">\s*(.*?)\s*</font>',
        html, re.DOTALL
    )

    if not peptide_match:
        return None, None, f"Could not parse peptides (HTTP {resp.status_code})"

    raw_peptides = re.sub(r'<[^>]+>', '', peptide_match.group(1)).strip()
    raw_locations = re.sub(r'<[^>]+>', '', location_match.group(1)).strip() if location_match else ""
    peptides = [p.strip() for p in raw_peptides.split(" - ") if p.strip()]
    return peptides, raw_locations, None


# ============================================================
# 解析位置区间，回填 start/end
# ============================================================
def parse_locations(raw_locations, peptides, full_sequence):
    """
    从 BIOPEP 返回的位置字符串中解析区间。
    格式如 "[1-5] [6-10] [11-15]" 或 "1-3 / 4-6 / 7-10"
    返回 [(start, end, peptide_seq), ...]
    """
    intervals = []
    # 匹配 [数字-数字] 或 纯数字-数字
    matches = re.findall(r'\[(\d+)[-–](\d+)\]|(\d+)[-–](\d+)', raw_locations)
    for m in matches:
        a = int(m[0] or m[2])
        b = int(m[1] or m[3])
        intervals.append((a, b))

    result = []
    for i, (a, b) in enumerate(intervals):
        pep = full_sequence[a-1:b] if a > 0 and b <= len(full_sequence) else (peptides[i] if i < len(peptides) else '')
        result.append((a, b, pep))
    return result


# ============================================================
# 输出 Excel
# ============================================================
def write_xlsx_detail(output_path, all_peptides_detail):
    """写明细表: protein_name, scheme_no, scheme_name, enzyme_ids, peptide, start, end, length"""
    headers = ["protein_name", "scheme_no", "scheme_name", "enzyme_ids", "peptide", "start", "end", "length"]

    # 构建共享字符串
    shared_strings = []
    ss_map = {}

    def get_ss(text):
        text = str(text) if text is not None else ""
        if text not in ss_map:
            ss_map[text] = len(shared_strings)
            shared_strings.append(text)
        return ss_map[text]

    for row in all_peptides_detail:
        for h in headers:
            get_ss(row.get(h, ''))

    def make_sheet_xml(headers, rows_data):
        ET.register_namespace('', NS[1:-1])
        root = ET.Element('worksheet', {'xmlns': NS[1:-1]})
        sheet_data = ET.SubElement(root, 'sheetData')

        # 表头
        hr = ET.SubElement(sheet_data, 'row', {'r': '1'})
        for ci, h in enumerate(headers):
            c = ET.SubElement(hr, 'c', {'r': f'{chr(65+ci)}1', 't': 's'})
            v = ET.SubElement(c, 'v')
            v.text = str(get_ss(h))

        # 数据行
        for ri, row in enumerate(rows_data):
            r_el = ET.SubElement(sheet_data, 'row', {'r': str(ri + 2)})
            for ci, h in enumerate(headers):
                c = ET.SubElement(r_el, 'c', {'r': f'{chr(65+ci)}{ri+2}', 't': 's'})
                v = ET.SubElement(c, 'v')
                v.text = str(get_ss(str(row.get(h, ''))))

        return ET.tostring(root, encoding='unicode', xml_declaration=False)

    def make_ss_xml():
        ET.register_namespace('', NS[1:-1])
        root = ET.Element('sst', {
            'xmlns': NS[1:-1],
            'count': str(len(shared_strings)),
            'uniqueCount': str(len(shared_strings))
        })
        for s in shared_strings:
            si = ET.SubElement(root, 'si')
            t = ET.SubElement(si, 't')
            t.text = s
            t.set(XML_SPACE, 'preserve')
        return ET.tostring(root, encoding='unicode', xml_declaration=False)

    sheet_xml = make_sheet_xml(headers, all_peptides_detail)
    ss_xml = make_ss_xml()

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml
        ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''
        zf.writestr('[Content_Types].xml', ct)

        rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
        zf.writestr('_rels/.rels', rels)

        wb = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Peptide_Details" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''
        zf.writestr('xl/workbook.xml', wb)

        wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)

        styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>'''
        zf.writestr('xl/styles.xml', styles)
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        zf.writestr('xl/sharedStrings.xml', ss_xml)

    print(f"  明细表: {output_path} ({len(all_peptides_detail)} rows)")


def write_xlsx_summary(output_path, all_summary):
    """写汇总表: protein_name, scheme_no, scheme_name, enzyme_ids, peptide_count, unique_peptide_count, total_length"""
    headers = ["protein_name", "scheme_no", "scheme_name", "enzyme_ids", "peptide_count", "unique_peptide_count", "total_length", "status"]

    shared_strings = []
    ss_map = {}

    def get_ss(text):
        text = str(text) if text is not None else ""
        if text not in ss_map:
            ss_map[text] = len(shared_strings)
            shared_strings.append(text)
        return ss_map[text]

    for row in all_summary:
        for h in headers:
            get_ss(str(row.get(h, '')))

    def make_sheet_xml(headers, rows_data):
        ET.register_namespace('', NS[1:-1])
        root = ET.Element('worksheet', {'xmlns': NS[1:-1]})
        sheet_data = ET.SubElement(root, 'sheetData')
        hr = ET.SubElement(sheet_data, 'row', {'r': '1'})
        for ci, h in enumerate(headers):
            c = ET.SubElement(hr, 'c', {'r': f'{chr(65+ci)}1', 't': 's'})
            v = ET.SubElement(c, 'v')
            v.text = str(get_ss(h))
        for ri, row in enumerate(rows_data):
            r_el = ET.SubElement(sheet_data, 'row', {'r': str(ri + 2)})
            for ci, h in enumerate(headers):
                c = ET.SubElement(r_el, 'c', {'r': f'{chr(65+ci)}{ri+2}', 't': 's'})
                v = ET.SubElement(c, 'v')
                v.text = str(get_ss(str(row.get(h, ''))))
        return ET.tostring(root, encoding='unicode', xml_declaration=False)

    def make_ss_xml():
        ET.register_namespace('', NS[1:-1])
        root = ET.Element('sst', {
            'xmlns': NS[1:-1],
            'count': str(len(shared_strings)),
            'uniqueCount': str(len(shared_strings))
        })
        for s in shared_strings:
            si = ET.SubElement(root, 'si')
            t = ET.SubElement(si, 't')
            t.text = s
            t.set(XML_SPACE, 'preserve')
        return ET.tostring(root, encoding='unicode', xml_declaration=False)

    sheet_xml = make_sheet_xml(headers, all_summary)
    ss_xml = make_ss_xml()

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''
        zf.writestr('[Content_Types].xml', ct)
        rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
        zf.writestr('_rels/.rels', rels)
        wb = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''
        zf.writestr('xl/workbook.xml', wb)
        wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>'''
        zf.writestr('xl/styles.xml', styles)
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        zf.writestr('xl/sharedStrings.xml', ss_xml)

    print(f"  汇总表: {output_path} ({len(all_summary)} rows)")


# ============================================================
# 主程序
# ============================================================
def main():
    print("=" * 60)
    print("蛋白序列批量酶解工具 (BIOPEP)")
    print(f"输入: {INPUT_XLSX}")
    print(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 1. 读取蛋白
    print("\n[1] 读取输入文件...")
    proteins = read_input_excel(INPUT_XLSX)
    print(f"  共读取 {len(proteins)} 条蛋白序列")
    for name, seq in proteins:
        print(f"    {name}: {len(seq)} aa")

    # 2. 建立 session
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    })

    # 3. 批量酶解
    total = len(proteins) * len(SCHEMES)
    count = 0
    all_detail = []
    all_summary = []

    print(f"\n[2] 开始酶解: {len(proteins)} 蛋白 × {len(SCHEMES)} 方案 = {total} 组合")
    print("-" * 60)

    for protein_name, sequence in proteins:
        for scheme_num, (scheme_name, enzyme_ids) in SCHEMES.items():
            count += 1
            enzyme_desc = " + ".join([ENZYME_NAMES.get(eid, eid) for eid in enzyme_ids])
            print(f"[{count}/{total}] {protein_name} + 方案{scheme_num}: {scheme_name} ...", end=" ", flush=True)

            peptides, locations, error = fetch_enzyme_result(
                session, protein_name, sequence, enzyme_ids
            )

            if error:
                print(f"ERROR: {error}")
                all_summary.append({
                    "protein_name": protein_name,
                    "scheme_no": str(scheme_num),
                    "scheme_name": scheme_name,
                    "enzyme_ids": enzyme_desc,
                    "peptide_count": "0",
                    "unique_peptide_count": "0",
                    "total_length": "0",
                    "status": f"ERROR: {error}",
                })
            else:
                # 解析位置区间
                intervals = parse_locations(locations, peptides, sequence)

                # 过滤长度<2的肽段
                peptides_filtered = []
                intervals_filtered = []
                for i, (a, b, pep) in enumerate(intervals):
                    if len(pep) >= 2:
                        peptides_filtered.append(pep)
                        intervals_filtered.append((a, b, pep))

                # 如果解析不到区间，使用肽段名称
                if not intervals_filtered and peptides:
                    for i, pep in enumerate(peptides):
                        if len(pep) >= 2:
                            peptides_filtered.append(pep)

                print(f"OK ({len(peptides_filtered)} peptides, len>={2 if peptides_filtered else 0})")

                # 汇总
                unique_peptides = set(peptides_filtered)
                total_len = sum(len(p) for p in peptides_filtered)
                all_summary.append({
                    "protein_name": protein_name,
                    "scheme_no": str(scheme_num),
                    "scheme_name": scheme_name,
                    "enzyme_ids": enzyme_desc,
                    "peptide_count": str(len(peptides_filtered)),
                    "unique_peptide_count": str(len(unique_peptides)),
                    "total_length": str(total_len),
                    "status": "OK",
                })

                # 明细
                for i, p in enumerate(peptides_filtered):
                    if intervals_filtered and i < len(intervals_filtered):
                        a, b, pep = intervals_filtered[i]
                        detail_start = str(a)
                        detail_end = str(b)
                    else:
                        detail_start = ''
                        detail_end = ''
                    all_detail.append({
                        "protein_name": protein_name,
                        "scheme_no": str(scheme_num),
                        "scheme_name": scheme_name,
                        "enzyme_ids": enzyme_desc,
                        "peptide": p,
                        "start": detail_start,
                        "end": detail_end,
                        "length": str(len(p)),
                    })

            # BIOPEP 礼貌延迟
            time.sleep(1.2)

    # 4. 写输出
    print("\n[3] 写入输出文件...")
    write_xlsx_detail(OUTPUT_DETAIL, all_detail)
    write_xlsx_summary(OUTPUT_SUMMARY, all_summary)

    # 5. 统计
    print("\n[4] 统计")
    ok_count = sum(1 for s in all_summary if s["status"] == "OK")
    fail_count = sum(1 for s in all_summary if s["status"] != "OK")
    total_peptides = sum(int(s["peptide_count"]) for s in all_summary if s["peptide_count"].isdigit())
    print(f"  成功: {ok_count}/{len(all_summary)} 组合")
    print(f"  失败: {fail_count}/{len(all_summary)} 组合")
    print(f"  总肽段数: {total_peptides}")

    # 6. 校验：检查同一序列的所有肽段长度之和
    print("\n[5] 校验（肽段长度之和 vs 原始序列长度）")
    for protein_name, sequence in proteins:
        for scheme_num, (scheme_name, enzyme_ids) in SCHEMES.items():
            detail_peptides = [d for d in all_detail
                               if d["protein_name"] == protein_name and d["scheme_no"] == str(scheme_num)]
            if detail_peptides:
                sum_len = sum(int(d["length"]) for d in detail_peptides if d["length"].isdigit())
                if sum_len != len(sequence):
                    print(f"  [!] {protein_name} + scheme{scheme_num}: sum={sum_len}, seq_len={len(sequence)}, diff={len(sequence)-sum_len}")
                # 只打印前几个差异
        # 只检查前几个蛋白避免刷屏
        break  # 只检查第一个蛋白做样

    print("\n" + "=" * 60)
    print("完成!")
    print(f"  明细表: {OUTPUT_DETAIL}")
    print(f"  汇总表: {OUTPUT_SUMMARY}")
    print("=" * 60)


if __name__ == "__main__":
    main()