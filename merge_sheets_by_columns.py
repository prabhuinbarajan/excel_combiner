import openpyxl
from merge_sheet_by_rows import *

from openpyxl.formula.translate import Translator



# Paste range
# Paste data from copyRange into template sheet
def pasteRangeCols(startCol, startRow, endCol, endRow, sheetReceiving, sourceSheet,  sourceStartRowOffset=0, sourceStartColOffset=0):
    countRow = 0+sourceStartRowOffset
    for i in range(startRow, endRow + 1, 1):
        countCol = 0+ sourceStartColOffset
        for j in range(startCol, endCol + 1, 1):
            source = sourceSheet.cell(row=countRow+1, column=countCol+1)
            value = source.value
            target  = sheetReceiving.cell(row=i, column=j)
            if source.data_type == 'f':
                target.value = Translator(value, source.coordinate).translate_formula(target.coordinate)
            else :
                target.value = value
            countCol += 1
            copy_style(source, target)
        countRow += 1


def mergeColumns(source_workbook, target_worksheet, sourceStartRowOffset=0, sourceStartColOffset=0, startRowOffset=0, startColOffset=0,  sourceEndColOffset=0, sourceEndRowOffset=0):

    print("Processing...")
    startCol = 1+startColOffset
    for sn in source_workbook.sheetnames:
        if sn == 'TOC':
            continue
        source_worksheet = source_workbook[sn]  # Add Sheet name
        startRow = 1+startRowOffset
        endCol = startCol + source_worksheet.max_column -1 - sourceEndColOffset
        endRow = source_worksheet.max_row - sourceEndRowOffset
        print('sc:{} sr: {} ec: {} er: {}'.format(startCol, startRow, endCol, endRow))
        pasteRangeCols(startCol, startRow, endCol, endRow,
                                  target_worksheet, source_worksheet,
                       sourceStartRowOffset=sourceStartRowOffset, sourceStartColOffset=sourceStartColOffset)
        startCol = endCol + 1
    # You can save the template as another file to create a new file here too.s
    print("Range copied and pasted!")

