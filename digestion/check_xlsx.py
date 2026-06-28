import zipfile, xml.etree.ElementTree as ET

path = r'c:\Users\kanmao\.claude\skills\digestion_detail_20260605_115017.xlsx'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

with zipfile.ZipFile(path, 'r') as zf:
    sst = ET.parse(zf.open('xl/sharedStrings.xml'))
    si_count = len(sst.findall(f'.//{NS}si'))
    print(f'Shared strings: {si_count}')

    sheet = ET.parse(zf.open('xl/worksheets/sheet1.xml'))
    rows = sheet.findall(f'.//{NS}row')
    print(f'Sheet rows: {len(rows)}')

    if len(rows) >= 2:
        r = rows[1]
        for c in r.findall(f'{NS}c'):
            ref = c.attrib.get('r', '')
            t_val = c.attrib.get('t', '')
            v_el = c.find(f'{NS}v')
            v_txt = v_el.text if v_el is not None else ''
            is_el = c.find(f'{NS}is')
            if is_el is not None:
                it = is_el.find(f'{NS}t')
                v_txt = it.text if it is not None else ''
                t_val = 'inlineStr'
            v_short = v_txt[:30] if v_txt else ''
            print('  Cell {}: type={}, v={}'.format(ref, t_val, v_short))