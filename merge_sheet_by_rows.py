import openpyxl
from openpyxl.formula.translate import Translator
import re
from copy import copy


# Paste range
# Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving, copiedData, rowOffset = 0 , colOffset = 0):
    countRow = 1
    for i in range(startRow, endRow , 1):
        countCol = 1
        for j in range(startCol, endCol , 1):
            source = copiedData.cell(row=countRow+ rowOffset, column=countCol)
            target = sheetReceiving.cell(row=i, column=j)
            #if type(cell).__name__ == 'MergedCell':

            if source.data_type == 'f':
                target.value = Translator(source.value, source.coordinate).translate_formula(target.coordinate)
            else :
                target.value = source.value
            #if source.has_style:
            #    target._style = copy(source._style)
            if source.has_style:
                target.font = copy(source.font)
                target.border = copy(source.border)
                target.fill = copy(source.fill)
                target.number_format = copy(source.number_format)
                target.protection = copy(source.protection)
                target.alignment = copy(source.alignment)

            countCol += 1
        countRow += 1


def createMergedSheet(worksheet, regex_filter, workbook, startCol, startRow, initialRowOffset, postRowShrinkage):

    print("Processing...")
    itemList = list(filter(lambda i: regex_filter.match(i), workbook.sheetnames))
    for sn in itemList:
        sheet1 = workbook[sn]  # Add Sheet name
        endCol = sheet1.max_column
        endRow = startRow+ sheet1.max_row-initialRowOffset-postRowShrinkage
        print('sc:{} sr: {} ec: {} er: {} sn: {}'.format(startCol, startRow, endCol, endRow,  sn))
        pasteRange(startCol, startRow, endCol, endRow,
                                  worksheet, sheet1, initialRowOffset)  # Change the 4 number values
        startRow=endRow

    # You can save the template as another file to create a new file here too.s
    print("Range copied and pasted!")
