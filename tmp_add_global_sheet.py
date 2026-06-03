"""
Patch models/16sectors_qc/economy.xlsx:
1. Remove accidental 'Global1' sheet
2. Add named ranges to existing Global sheet (transport_fraction, inland_transport_fraction)
3. Add historic_goverment_expenditures typo alias to World and Quebec via XML
"""
import shutil, zipfile, os, re
from lxml import etree
import openpyxl

PATH   = 'models/16sectors_qc/economy.xlsx'
BACKUP = 'models/16sectors_qc/economy_pre_global_backup.xlsx'
NS     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

shutil.copy(PATH, BACKUP)
print(f'Backed up to {BACKUP}')

# Step 1: remove Global1 sheet using openpyxl
wb = openpyxl.load_workbook(PATH)
if 'Global1' in wb.sheetnames:
    del wb['Global1']
    print('Removed accidental Global1 sheet')
wb.save(PATH)

# Step 2: add named ranges via XML (reliable approach)
# Determine sheet indices from the saved file
with zipfile.ZipFile(PATH) as z:
    tree = etree.fromstring(z.read('xl/workbook.xml'))
    sheets = [s.get('name') for s in tree.findall(f'.//{{{NS}}}sheet')]
    print('Sheets:', sheets)

GLOBAL_ID = str(sheets.index('Global'))
WORLD_ID  = str(sheets.index('World'))
QUEBEC_ID = str(sheets.index('Quebec'))

NEW_RANGES = [
    # Global sheet: transport fractions (data in row 2 and 3, cols B-Q)
    ('inland_transport_fraction', GLOBAL_ID, "Global!$B$2:$Q$2"),
    ('transport_fraction',        GLOBAL_ID, "Global!$B$3:$Q$3"),
    # Typo alias for model code that references 'historic_goverment_expenditures'
    ('historic_goverment_expenditures', WORLD_ID,  "World!$C$74:$Z$89"),
    ('historic_goverment_expenditures', QUEBEC_ID, "Quebec!$C$74:$Z$89"),
]

tmp = PATH + '.tmp'
with zipfile.ZipFile(PATH, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/workbook.xml':
            tree = etree.fromstring(data)
            names_el = tree.find(f'{{{NS}}}definedNames')
            if names_el is None:
                names_el = etree.SubElement(tree, f'{{{NS}}}definedNames')

            existing = {(dn.get('name'), dn.get('localSheetId'))
                        for dn in names_el.findall(f'{{{NS}}}definedName')}
            for name, sid, ref in NEW_RANGES:
                if (name, sid) not in existing:
                    el = etree.SubElement(names_el, f'{{{NS}}}definedName')
                    el.set('name', name)
                    el.set('localSheetId', sid)
                    el.text = ref
                    print(f'  Added [{sheets[int(sid)]}] {name} -> {ref}')
                else:
                    print(f'  Already exists: [{sheets[int(sid)]}] {name}')

            data = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)

os.replace(tmp, PATH)
print('Done.')
