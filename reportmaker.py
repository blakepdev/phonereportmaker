from openpyxl import Workbook
from openpyxl import load_workbook
import datetime

workbookName = input("Name of workbook?: ")
exportName = input("Date for export: ")
workbook = load_workbook(filename=workbookName)
sheet = workbook.active

sheet.delete_cols(5,6)
sheet.delete_cols(6)
sheet.insert_cols(1)
sheet.insert_cols(4)

for cell in sheet['G']:
    if(cell.value == None):
        cell.value = 0
    elif('secs' in cell.value):
        cell.value = cell.value.replace(' secs', '')
        cell.value = int(cell.value)
    elif('mins' in cell.value):
        cell.value = cell.value.replace(' mins', '')
        time = int(cell.value)
        seconds = time * 60
        cell.value = seconds
    elif('min' in cell.value):
        cell.value = cell.value.replace(' min', '')
        time = int(cell.value)
        seconds = time * 60
        cell.value = seconds
    print(cell.value)

for cell in sheet['E']:
    if(cell.value != "Call Type"):
        cell.value = "Received"

    
workbook.save(filename=exportName + "_PhoneExport.xlsx")
