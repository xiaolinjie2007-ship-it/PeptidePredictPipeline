import zipfile, xml.etree.ElementTree as ET

path = r'C:\Users\kanmao\Desktop\enzyme_digestion_results.xlsx'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def read_ss(zf):
    ss = []
    sst = ET.parse(zf.open('xl/sharedStrings.xml'))
    for si in sst.findall(f'.//{NS}si'):
        t = si.find(f'{NS}t')
        if t is not None and t.text:
            ss.append(t.text)
        else:
            ss.append('')
    return ss

with zipfile.ZipFile(path, 'r') as zf:
    ss = read_ss(zf)
    print('SS count: {}'.format(len(ss)))

    # Sheet 1
    s1 = ET.parse(zf.open('xl/worksheets/sheet1.xml'))
    r1 = s1.findall(f'.//{NS}row')
    print('Sheet1 (Peptide_Details): {} rows'.format(len(r1)))

    # Header
    hdr = {}
    for c in r1[0].findall(f'{NS}c'):
        ref = c.attrib.get('r', '')
        col = ''.join(filter(str.isalpha, ref))
        if c.get('t') == 's':
            v = c.find(f'{NS}v')
            idx = int(v.text) if v is not None and v.text else 0
            hdr[col] = ss[idx] if idx < len(ss) else ''
    print('  Headers: {}'.format(hdr))

    # First 3 data rows
    for row in r1[1:4]:
        cells = {}
        for c in row.findall(f'{NS}c'):
            ref = c.attrib.get('r', '')
            col = ''.join(filter(str.isalpha, ref))
            if c.get('t') == 's':
                v = c.find(f'{NS}v')
                idx = int(v.text) if v is not None and v.text else 0
                cells[col] = ss[idx][:30] if idx < len(ss) else '?'
        rn = row.attrib.get('r', '?')
        print('  Row {}: {}'.format(rn, cells))

    # Last row
    last = r1[-1]
    cells = {}
    for c in last.findall(f'{NS}c'):
        ref = c.attrib.get('r', '')
        col = ''.join(filter(str.isalpha, ref))
        if c.get('t') == 's':
            v = c.find(f'{NS}v')
            idx = int(v.text) if v is not None and v.text else 0
            cells[col] = ss[idx][:30] if idx < len(ss) else '?'
    print('  Last row {}: {}'.format(last.attrib.get('r'), cells))

    # Sheet 2
    s2 = ET.parse(zf.open('xl/worksheets/sheet2.xml'))
    r2 = s2.findall(f'.//{NS}row')
    print('\nSheet2 (Summary): {} rows'.format(len(r2)))

    hdr2 = {}
    for c in r2[0].findall(f'{NS}c'):
        ref = c.attrib.get('r', '')
        col = ''.join(filter(str.isalpha, ref))
        if c.get('t') == 's':
            v = c.find(f'{NS}v')
            idx = int(v.text) if v is not None and v.text else 0
            hdr2[col] = ss[idx] if idx < len(ss) else ''
    print('  Headers: {}'.format(hdr2))

    for row in r2[1:4]:
        cells = {}
        for c in row.findall(f'{NS}c'):
            ref = c.attrib.get('r', '')
            col = ''.join(filter(str.isalpha, ref))
            if c.get('t') == 's':
                v = c.find(f'{NS}v')
                idx = int(v.text) if v is not None and v.text else 0
                cells[col] = ss[idx][:30] if idx < len(ss) else '?'
        rn = row.attrib.get('r', '?')
        print('  Row {}: {}'.format(rn, cells))