"""
Backfill ALL 'na'/NaN cells for years 1995-2003 in economy.xlsx World and Quebec sheets
with the value from the same row at year 2004 (col L).

Layout: cols A,B = index/label; col C (col 3) = 1995; col L (col 12) = 2004.
Any cell at cols 3-11 that is 'na'/None gets replaced with the col 12 value of that row.

Also adds the same logic to QC_IOT.py comment.
"""
import shutil, openpyxl

PATH   = 'models/16sectors_qc/economy.xlsx'
BACKUP = 'models/16sectors_qc/economy_pre_fullbackfill_backup.xlsx'

YEAR_1995_COL = 3   # col C
YEAR_2003_COL = 11  # col K
YEAR_2004_COL = 12  # col L

shutil.copy(PATH, BACKUP)
print(f'Backed up to {BACKUP}')

wb = openpyxl.load_workbook(PATH)

for sheet_name in ['World', 'Quebec']:
    if sheet_name not in wb.sheetnames:
        continue
    ws = wb[sheet_name]
    patched_cells = 0
    skipped_rows  = 0

    for row in ws.iter_rows():
        # Get the 2004 reference value for this row (col L = col 12)
        val_2004 = ws.cell(row=row[0].row, column=YEAR_2004_COL).value

        # Skip rows where 2004 is also 'na'/None — nothing to backfill from
        if val_2004 is None or val_2004 == 'na':
            skipped_rows += 1
            continue

        # For each year column 1995-2003 (cols 3-11), replace 'na'/None with 2004 value
        for col in range(YEAR_1995_COL, YEAR_2003_COL + 1):
            cell = ws.cell(row=row[0].row, column=col)
            if cell.value is None or cell.value == 'na':
                cell.value = val_2004
                patched_cells += 1

    print(f'  {sheet_name}: backfilled {patched_cells} cells '
          f'({skipped_rows} rows skipped — no 2004 data)')

wb.save(PATH)
print('Done.')
