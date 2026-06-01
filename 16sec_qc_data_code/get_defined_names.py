
import openpyxl

wb = openpyxl.load_workbook("./16sec_qc_data_code/original_and_mock_xlsx/14sec_original_data/economy.xlsx")

ws = wb['World']

for key in ws.defined_names:
    print(key)

