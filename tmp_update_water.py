"""
Patch models/16sectors_qc/water.xlsx to support 16 SECTORS_and_HOUSEHOLDS.

Each sheet (World, Europe) has 3 water-type blocks of 15 rows each.
Inserts 2 new zero rows at the end of each block (bottom-to-top to avoid shift errors)
and updates named ranges accordingly.

Block layout before (per sheet):
  Row  1: year header (B1:P1)
  Row  2: "Blue Water" label
  Rows 3–17: blue data (15 sectors)
  Rows 18–19: empty
  Row 20: "Green Water" label
  Rows 21–35: green data (15 sectors)
  Rows 36–37: empty
  Row 38: "Gray Water" label
  Rows 39–53: gray data (15 sectors)

After inserting 2 rows per block (6 total, bottom-to-top):
  Rows  3–19: blue  (17 sectors)   named range B3:P19
  Rows 23–39: green (17 sectors)   named range B23:P39
  Rows 43–59: gray  (17 sectors)   named range B43:P59
"""
import json, openpyxl, zipfile, shutil, os
from lxml import etree

PATH   = 'models/16sectors_qc/water.xlsx'
BACKUP = 'models/16sectors_qc/water_14sec_backup.xlsx'
SUBS   = 'models/16sectors_qc/_subscripts_pymedeas_w.json'
NS     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

with open(SUBS) as f:
    sectors = json.load(f)['SECTORS_and_HOUSEHOLDS']  # 17 entries
new_sectors = sectors[15:]   # last 2: Other_c.s.p... and Private_Households
print(f'New sectors to add: {new_sectors}')

shutil.copy(PATH, BACKUP)
print(f'Backed up to {BACKUP}')

# ── Step 1: insert rows in openpyxl ──────────────────────────────────────────
wb = openpyxl.load_workbook(PATH)
UPDATE_SHEETS = ['World', 'Europe']   # skip Catalonia

for sname in UPDATE_SHEETS:
    ws = wb[sname]
    # Insert bottom-to-top so earlier insertions don't affect row numbers
    # Gray block: insert 2 rows before row 54 (after data ends at 53)
    ws.insert_rows(54, 2)
    # Green block: insert 2 rows before row 36 (after data ends at 35)
    ws.insert_rows(36, 2)
    # Blue block: insert 2 rows before row 18 (after data ends at 17)
    ws.insert_rows(18, 2)

    # Fill new rows: [sector_label, 0, 0, ...] for cols A(1) and B–P(2–16)
    for block_start, label_prefix in [(18, 'blue'), (38, 'green'), (58, 'gray')]:
        for idx, sector in enumerate(new_sectors):
            r = block_start + idx
            ws.cell(row=r, column=1).value = sector
            for col in range(2, 17):   # cols B-P (15 year columns)
                ws.cell(row=r, column=col).value = 0

    print(f'  {sname}: inserted rows, filled {len(new_sectors)*3} new sector cells')

wb.save(PATH)
print('Saved with new rows.')

# ── Step 2: update named ranges in workbook XML ───────────────────────────────
# New ranges after insertions (same for both World and Europe sheets):
#   historic_water_use_blue_water  : $B$3:$P$19   (was $B$3:$P$17)
#   historic_water_use_green_water : $B$23:$P$39  (was $B$21:$P$35)
#   historic_water_use_gray_water  : $B$43:$P$59  (was $B$39:$P$53)
#   year                           : $B$1:$P$1    (unchanged)
RANGE_UPDATES = {
    'historic_water_use_blue_water':  '$B$3:$P$19',
    'historic_water_use_green_water': '$B$23:$P$39',
    'historic_water_use_gray_water':  '$B$43:$P$59',
}

# Determine sheet IDs to update
with zipfile.ZipFile(PATH) as z:
    tree = etree.fromstring(z.read('xl/workbook.xml'))
    sheet_names = [s.get('name') for s in tree.findall(f'.//{{{NS}}}sheet')]
update_ids = {str(sheet_names.index(s)) for s in UPDATE_SHEETS}

tmp = PATH + '.tmp'
with zipfile.ZipFile(PATH, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/workbook.xml':
            tree = etree.fromstring(data)
            updated = 0
            for dn in tree.findall(f'.//{{{NS}}}definedName'):
                name = dn.get('name', '')
                sid  = dn.get('localSheetId')
                if name in RANGE_UPDATES and sid in update_ids:
                    sheet = sheet_names[int(sid)]
                    new_ref = f'{sheet}!{RANGE_UPDATES[name]}'
                    print(f'  [{sheet}] {name}: {dn.text} -> {new_ref}')
                    dn.text = new_ref
                    updated += 1
            print(f'Updated {updated} named ranges.')
            data = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)

os.replace(tmp, PATH)
print('Done.')
