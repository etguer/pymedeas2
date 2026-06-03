import zipfile, openpyxl, re
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
PATH = 'models/16sectors_qc/economy.xlsx'

with zipfile.ZipFile(PATH) as z:
    tree = etree.fromstring(z.read('xl/workbook.xml'))
    sheets = [s.get('name') for s in tree.findall(f'.//{{{NS}}}sheet')]

wb = openpyxl.load_workbook(PATH, read_only=True, data_only=True)

# Check energy intensity named ranges: do they have non-zero values?
with zipfile.ZipFile(PATH) as z:
    tree = etree.fromstring(z.read('xl/workbook.xml'))
    for dn in tree.findall(f'.//{{{NS}}}definedName'):
        name = dn.get('name', '')
        if 'final_energy_intensity' not in name: continue
        sid = dn.get('localSheetId')
        sheet = sheets[int(sid)] if sid else 'GLOBAL'
        ref = dn.text or ''
        # parse ref
        m = re.match(r"'?([^!']+)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", ref)
        if not m: continue
        def cn(c):
            n=0
            for ch in c: n=n*26+(ord(ch)-64)
            return n
        c1,r1 = cn(m.group(2)), int(m.group(3))
        c2 = cn(m.group(4)) if m.group(4) else c1
        r2 = int(m.group(5)) if m.group(5) else r1
        if sheet not in wb.sheetnames: continue
        ws = wb[sheet]
        vals = [ws.cell(r,c).value for r in range(r1,min(r1+3,r2+1))
                                    for c in range(c1,min(c1+5,c2+1))]
        non_zero = sum(1 for v in vals if v is not None and v != 'na'
                       and v != 0 and v == v)
        print(f'[{sheet}] {name}')
        print(f'  ref={ref}  first vals: {vals[:8]}')
        print(f'  non-zero in sample: {non_zero}/{len(vals)}')
