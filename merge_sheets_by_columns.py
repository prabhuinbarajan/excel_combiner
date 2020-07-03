import openpyxl
from openpyxl.formula.translate import Translator



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


def mergeColumns(source_workbook, target_worksheet):

    print("Processing...")
    startCol = 1
    for sn in source_workbook.sheetnames:
        if sn == 'TOC':
            continue
        source_worksheet = source_workbook[sn]  # Add Sheet name
        startRow = 1
        endCol = startCol + source_worksheet.max_column -1
        endRow = source_worksheet.max_row
        print('sc:{} sr: {} ec: {} er: {}'.format(startCol, startRow, endCol, endRow))
        pasteRange(startCol, startRow, endCol, endRow,
                                  target_worksheet, source_worksheet)  # Change the 4 number values
        startCol = startCol + source_worksheet.max_column
    # You can save the template as another file to create a new file here too.s
    print("Range copied and pasted!")

