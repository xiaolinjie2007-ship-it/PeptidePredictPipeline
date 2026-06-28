#!/usr/bin/env python3
"""
Merge two xlsx files (detail + summary) into one desktop xlsx with 2 sheets.
Both source files use shared strings - need to merge SS tables and remap indices.
"""
import zipfile, xml.etree.ElementTree as ET, os, re

detail_src = r'c:\Users\kanmao\.claude\skills\digestion_detail_20260605_115017.xlsx'
summary_src = r'c:\Users\kanmao\.claude\skills\digestion_summary_20260605_115017.xlsx'
output = r'C:\Users\kanmao\Desktop\enzyme_digestion_results.xlsx'

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def read_shared_strings(zf):
    """Read shared strings from a zipfile, return list of strings"""
    strings = []
    try:
        sst = ET.parse(zf.open('xl/sharedStrings.xml'))
        for si in sst.findall(f'.//{{{NS}}}si'):
            t = si.find(f'{{{NS}}}t')
            if t is not None and t.text:
                strings.append(t.text)
            else:
                # Rich text: concatenate all <r><t> parts
                parts = []
                for r_elem in si.findall(f'{{{NS}}}r'):
                    rt = r_elem.find(f'{{{NS}}}t')
                    if rt is not None and rt.text:
                        parts.append(rt.text)
                strings.append(''.join(parts))
    except Exception as e:
        print(f"  Warning reading SS: {e}")
    return strings

def remap_sheet(xml_bytes, old_to_new):
    """Replace shared string indices in sheet XML. xml_bytes is raw bytes."""
    # Parse, replace, serialize back
    root = ET.fromstring(xml_bytes)
    for c in root.iter(f'{{{NS}}}c'):
        if c.get('t') == 's':
            v_el = c.find(f'{{{NS}}}v')
            if v_el is not None and v_el.text:
                try:
                    old_idx = int(v_el.text)
                    new_idx = old_to_new.get(old_idx)
                    if new_idx is not None:
                        v_el.text = str(new_idx)
                except ValueError:
                    pass
    return ET.tostring(root, encoding='unicode', xml_declaration=False)

# ---- Step 1: Read both SS tables ----
print("[1] Reading shared strings...")
with zipfile.ZipFile(detail_src, 'r') as zf:
    ss1 = read_shared_strings(zf)
    detail_sheet_raw = zf.read('xl/worksheets/sheet1.xml')
print(f"  Detail: {len(ss1)} shared strings, sheet {len(detail_sheet_raw)} bytes")

with zipfile.ZipFile(summary_src, 'r') as zf:
    ss2 = read_shared_strings(zf)
    summary_sheet_raw = zf.read('xl/worksheets/sheet1.xml')
print(f"  Summary: {len(ss2)} shared strings, sheet {len(summary_sheet_raw)} bytes")

# ---- Step 2: Merge SS (dedup) ----
print("[2] Merging shared strings...")
merged_ss = []           # final list
ss1_map = {}             # old index in ss1 -> new index in merged
ss2_map = {}             # old index in ss2 -> new index in merged
value_to_idx = {}        # string value -> new index

for i, val in enumerate(ss1):
    if val in value_to_idx:
        ss1_map[i] = value_to_idx[val]
    else:
        new_idx = len(merged_ss)
        merged_ss.append(val)
        value_to_idx[val] = new_idx
        ss1_map[i] = new_idx

for i, val in enumerate(ss2):
    if val in value_to_idx:
        ss2_map[i] = value_to_idx[val]
    else:
        new_idx = len(merged_ss)
        merged_ss.append(val)
        value_to_idx[val] = new_idx
        ss2_map[i] = new_idx

print(f"  SS1: {len(ss1)} -> merged {len(set(ss1_map.values()))} unique")
print(f"  SS2: {len(ss2)} -> merged {len(set(ss2_map.values()))} unique")
print(f"  Total merged: {len(merged_ss)}")

# ---- Step 3: Remap sheet XML ----
print("[3] Remapping sheet indices...")
detail_sheet_new = remap_sheet(detail_sheet_raw, ss1_map)
summary_sheet_new = remap_sheet(summary_sheet_raw, ss2_map)

# Verify by counting rows
d_root = ET.fromstring(detail_sheet_new)
s_root = ET.fromstring(summary_sheet_new)
d_rows = len(d_root.findall(f'.//{{{NS}}}row'))
s_rows = len(s_root.findall(f'.//{{{NS}}}row'))
print(f"  Detail sheet rows: {d_rows}, Summary sheet rows: {s_rows}")

# ---- Step 4: Build merged SS XML ----
print("[4] Building merged SS XML...")
sst_root = ET.Element(f'{{{NS}}}sst', {
    'xmlns': NS,
    'count': str(len(merged_ss)),
    'uniqueCount': str(len(merged_ss))
})
for s in merged_ss:
    si = ET.SubElement(sst_root, f'{{{NS}}}si')
    t = ET.SubElement(si, f'{{{NS}}}t')
    t.text = s
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
merged_ss_xml = ET.tostring(sst_root, encoding='unicode', xml_declaration=True)

# ---- Step 5: Write final xlsx ----
print("[5] Writing output...")
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr('[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
        '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        '  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n'
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n'
        '</Types>')
    zf.writestr('_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
        '</Relationships>')
    zf.writestr('xl/workbook.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        '  <sheets>\n'
        '    <sheet name="Peptide_Details" sheetId="1" r:id="rId1"/>\n'
        '    <sheet name="Summary" sheetId="2" r:id="rId2"/>\n'
        '  </sheets>\n'
        '</workbook>')
    zf.writestr('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\n'
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>\n'
        '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n'
        '  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        '</Relationships>')
    zf.writestr('xl/styles.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        '  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>\n'
        '  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>\n'
        '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
        '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
        '  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>\n'
        '</styleSheet>')
    zf.writestr('xl/worksheets/sheet1.xml', detail_sheet_new)
    zf.writestr('xl/worksheets/sheet2.xml', summary_sheet_new)
    zf.writestr('xl/sharedStrings.xml', merged_ss_xml)

size_kb = os.path.getsize(output) / 1024
print(f"\nDone! {output} ({size_kb:.1f} KB)")
print("  Sheet 1: Peptide_Details (23035 rows)")
print("  Sheet 2: Summary (350 rows)")