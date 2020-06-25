import openpyxl
from openpyxl.formula.translate import Translator

# Prepare the spreadsheets to copy from and paste too.

# File to be copied
workbook = 'A510_TB_Multiple BUs_LC_All_BUS'
workbook_url = 'report_samples/{}.xlsx'.format(workbook)
workbook_template = 'templates/{}.xltx' .format(workbook)
result_workbook = 'results/{}_combined.xlsx'.format(workbook)

wb = openpyxl.load_workbook(workbook_url)  # Add file name

# Paste range
# Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving, copiedData):
    countRow = 0
    for i in range(startRow, endRow + 1, 1):
        countCol = 0
        for j in range(startCol, endCol + 1, 1):
            cell = copiedData.cell(row=countRow+1, column=countCol+1)
            value = cell.value
            if cell.data_type == 'f':
                sheetReceiving.cell(row=i, column=j).value = Translator(value, cell.coordinate).translate_formula(sheetReceiving.cell(row=i, column=j).coordinate)
            else :
                sheetReceiving.cell(row=i, column=j).value = value
            countCol += 1
        countRow += 1


def mergeColumns():

    print("Processing...")
    startCol = 1
    for sn in wb.sheetnames:
        if sn == 'TOC':
            continue
        sheet1 = wb[sn]  # Add Sheet name
        startRow = 1
        endCol = startCol + sheet1.max_column -1
        endRow = sheet1.max_row
        print('sc:{} sr: {} ec: {} er: {}'.format(startCol, startRow, endCol, endRow))
        pasteRange(startCol, startRow, endCol, endRow,
                                  temp_sheet, sheet1)  # Change the 4 number values
        startCol = startCol + sheet1.max_column
    # You can save the template as another file to create a new file here too.s
    target.save(result_workbook)
    print("Range copied and pasted!")


target = openpyxl.load_workbook(workbook_template)
target.template = False
#target = copy(wb_template)
temp_sheet = target['Sheet']
mergeColumns()
