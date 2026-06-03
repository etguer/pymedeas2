"""
Patch economy.xlsx defined names so each named range stops at the correct column.

Only ranges whose current right column is Z are touched; all others are left alone.

  time_index2009 / time_index_2009  →  right col Q  (col 17, year 2009)
  time_index2014                    →  right col V  (col 22, year 2014)
  time_index_projection             →  right col AL (col 38, year 2030)
  all other data ranges ending in Z →  right col V  (col 22, year 2014)
"""
import re
import openpyxl

path = "models/16sectors_qc/economy.xlsx"
wb = openpyxl.load_workbook(path)

# Only match named ranges whose right-most cell column is exactly Z.
# Pattern: ...:$Z$<row> at end of attr_text
RIGHT_Z_RE = re.compile(r':\$Z(\$\d+)$')

def change_right_col(attr_text, new_col):
    """Replace trailing :$Z$N with :$<new_col>$N. No-op if right col isn't Z."""
    return RIGHT_Z_RE.sub(f':${new_col}\\1', attr_text)

TIME_INDEX_2009 = {'time_index2009', 'time_index_2009'}
TIME_INDEX_2014 = {'time_index2014'}
TIME_INDEX_PROJ = {'time_index_projection'}

print("Changes made:")
for ws in wb.worksheets:
    for dn in list(ws.defined_names.values()):
        name = dn.name
        old = dn.attr_text

        if not RIGHT_Z_RE.search(old):
            continue  # right column is not Z — skip

        if name in TIME_INDEX_2009:
            new = change_right_col(old, 'Q')
        elif name in TIME_INDEX_2014:
            new = change_right_col(old, 'V')
        elif name in TIME_INDEX_PROJ:
            new = change_right_col(old, 'AL')
        else:
            new = change_right_col(old, 'V')

        if new != old:
            dn.attr_text = new
            print(f"  [{ws.title}] {name}:  {old}  →  {new}")

wb.save(path)
print("\nDone — economy.xlsx saved.")
