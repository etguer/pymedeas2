"""
Update models/16sectors_qc/climate.xlsx to support 16-sector scheme.

Updates World (sheet id=1) and Europe (sheet id=2) — not Catalonia (id=3).

For each sheet:
- Inserts 6 rows after the last 14-sector CCS block for 2 new zero-data sectors
- Adds named ranges ccs_tech_share_{SectorName} for all 17 sectors,
  positionally mapped to existing rows + 2 new zero rows
- Shifts all named ranges in that sheet that reference rows >= insert point by +6
- All original 14-sector named ranges are preserved
"""
import json, openpyxl, zipfile, shutil, os, re
from lxml import etree

SRC    = 'models/16sectors_qc/climate_14sec_backup.xlsx'  # read from backup
PATH   = 'models/16sectors_qc/climate.xlsx'              # final write target
OUTPATH = 'models/16sectors_qc/climate_16sec_new.xlsx'   # intermediate output
SUBS_PATH = 'models/16sectors_qc/_subscripts_pymedeas_w.json'
NS        = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

with open(SUBS_PATH) as f:
    sectors = json.load(f)['SECTORS_and_HOUSEHOLDS']  # 17 entries
print(f'Sectors ({len(sectors)}): {sectors}')
assert len(sectors) == 17

N_EXISTING = 15   # Households + 14 sectors already in both sheets

# Sheet config: name, localSheetId, first CCS data row, insert-after row
SHEETS = {
    'World':  {'id': '1', 'first_row': 152, 'insert_at': 197},
    'Europe': {'id': '2', 'first_row': 59,  'insert_at': 104},
}
# World:  rows 152-154=hh, ..., 194-196=nms  → insert 6 rows at 197
# Europe: rows 59-61=hh,  ..., 101-103=nms  → insert 6 rows at 104

print(f'Reading from backup: {SRC}')

# ── Step 1: insert rows and fill zero data ─────────────────────────────────
wb = openpyxl.load_workbook(SRC)

for sheet_name, cfg in SHEETS.items():
    ws = wb[sheet_name]
    first_row  = cfg['first_row']
    insert_at  = cfg['insert_at']
    ws.insert_rows(insert_at, 6)

    for idx in range(N_EXISTING, len(sectors)):
        sector    = sectors[idx]
        start_row = first_row + idx * 3
        ws.cell(row=start_row,     column=1).value = sector
        ws.cell(row=start_row,     column=2).value = 'post-combustion'
        ws.cell(row=start_row + 1, column=2).value = 'pre-combustion'
        ws.cell(row=start_row + 2, column=2).value = 'oxyfuel combustion'
        for col in range(3, 11):   # columns C-J (8 time points)
            ws.cell(row=start_row,     column=col).value = 0
            ws.cell(row=start_row + 1, column=col).value = 0
            ws.cell(row=start_row + 2, column=col).value = 0
        print(f'  {sheet_name}: added zero rows {start_row}-{start_row+2} for {sector}')

wb.save(OUTPATH)
print(f'Saved with new rows to {OUTPATH}')

# ── Step 2: patch named ranges in workbook XML ─────────────────────────────
def shift_refs(text, sheet_name, threshold, delta=6):
    """Shift ALL $COL$ROW row numbers >= threshold in any range that belongs to sheet_name."""
    if f'{sheet_name}!' not in text:
        return text
    def fix(m):
        row = int(m.group(2))
        return f'{m.group(1)}{row + delta}' if row >= threshold else m.group(0)
    return re.sub(r'(\$[A-Z]+\$)(\d+)', fix, text)

# Build new named ranges for each sheet
new_ranges = []   # list of (name, localSheetId, ref)
for sheet_name, cfg in SHEETS.items():
    sheet_id  = cfg['id']
    first_row = cfg['first_row']
    for idx, sector in enumerate(sectors):
        r1 = first_row + idx * 3
        r2 = r1 + 2
        new_ranges.append((f'ccs_tech_share_{sector}', sheet_id, f'{sheet_name}!$C${r1}:$N${r2}'))

print(f'\nNew named ranges to add ({len(new_ranges)}):')
for name, sid, ref in new_ranges:
    print(f'  [sheet {sid}] {name} -> {ref}')

tmp = PATH + '.tmp'
with zipfile.ZipFile(OUTPATH, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/workbook.xml':
            tree = etree.fromstring(data)
            ns   = {'ns': NS}
            names_el = tree.find('ns:definedNames', ns)

            # Shift existing ranges in World and Europe sheets
            shifted = 0
            for dn in names_el.findall('ns:definedName', ns):
                old = dn.text or ''
                new = old
                for sheet_name, cfg in SHEETS.items():
                    new = shift_refs(new, sheet_name, cfg['insert_at'])
                if new != old:
                    dn.text = new
                    shifted += 1
            print(f'\nShifted {shifted} existing named ranges.')

            # Add new 16-sector named ranges (skip if already exists for same sheet)
            existing = {(dn.get('name'), dn.get('localSheetId'))
                        for dn in names_el.findall('ns:definedName', ns)}
            added = 0
            for name, sid, ref in new_ranges:
                if (name, sid) not in existing:
                    el = etree.SubElement(names_el, f'{{{NS}}}definedName')
                    el.set('name', name)
                    el.set('localSheetId', sid)
                    el.text = ref
                    added += 1
            print(f'Added {added} new named ranges.')

            data = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)

os.replace(tmp, OUTPATH)
# Attempt to replace final target; if locked, leave as OUTPATH for manual rename
try:
    os.replace(OUTPATH, PATH)
    print(f'Done. Written to {PATH}')
except PermissionError:
    print(f'NOTE: {PATH} is open — close it in Excel, then rename {OUTPATH} to climate.xlsx')
    print('Done (written to climate_16sec_new.xlsx).')
