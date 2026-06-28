#!/usr/bin/env python3
"""Inspect the input Excel file structure - handle inline strings"""
import zipfile
import xml.etree.ElementTree as ET

path = r"D:\kanmao\桌面\uniprotkb_accsion_A0AAD4IMA6_OR_access_2026_06_05.xlsx"
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

with zipfile.ZipFile(path, 'r') as zf:
    # Print all files in zip
    print("=== ZIP contents ===")
    for name in zf.namelist():
        print(f"  {name}")

    # Read shared strings
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
    except Exception as e:
        print(f"  No shared strings: {e}")

    print(f"\nShared strings count: {len(ss)}")

    # Read sheet1
    sheet = ET.parse(zf.open('xl/worksheets/sheet1.xml'))
    rows = sheet.findall(f'.//{NS}row')

    print(f"Total rows: {len(rows)}")

    def col_letter(ref):
        return ''.join(filter(str.isalpha, ref))

    def get_cell_value(c):
        """Get cell value - handle both shared strings and inline strings"""
        t = c.attrib.get('t', '')
        # Check for inline string
        is_el = c.find(f'{NS}is')
        if is_el is not None:
            # Simple text
            t_el = is_el.find(f'{NS}t')
            if t_el is not None:
                return t_el.text or ''
            # Rich text
            parts = []
            for r_elem in is_el.findall(f'{NS}r'):
                rt = r_elem.find(f'{NS}t')
                if rt is not None:
                    parts.append(rt.text or '')
            return ''.join(parts)

        # Shared string
        v_el = c.find(f'{NS}v')
        v = v_el.text if v_el is not None else ''
        if t == 's' and v.isdigit():
            idx = int(v)
            return ss[idx] if idx < len(ss) else v
        return v

    print("\n=== All rows ===")
    for r in rows:
        rn = r.attrib.get('r', '?')
        cells = {}
        for c in r.findall(f'{NS}c'):
            ref = c.attrib.get('r', '')
            col = col_letter(ref)
            cells[col] = get_cell_value(c)
        # Only print rows with actual content
        non_empty = {k: v for k, v in cells.items() if v}
        if non_empty:
            print(f"  Row {rn}: {non_empty}")
        else:
            print(f"  Row {rn}: (empty)")