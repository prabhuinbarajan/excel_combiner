import openpyxl
from merge_sheet_by_rows import *

from openpyxl.formula.translate import Translator



# Paste range
# Paste data from copyRange into template sheet
def pasteRangeCols(startCol, startRow, endCol, endRow, sheetReceiving, sourceSheet,  sourceStartRowOffset=0, sourceStartColOffset=0, enableGrouping=False):
    countRow = 0+sourceStartRowOffset
    groupRange = None
    groupList = []
    for i in range(startRow, endRow + 1, 1):
        countCol = 0+ sourceStartColOffset
        if not sourceSheet.row_dimensions[i].hidden:
            if groupRange is not None:
                groupList.append(groupRange)
                groupRange = None
            groupRange = {'start' : i, 'end': i}
        else:
            groupRange['end'] = i

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
    if enableGrouping:
        for gr in groupList:
            if gr['start'] != gr['end']:
                sheetReceiving.row_dimensions.group(start=gr['start']+1, end=gr['end'], hidden=True)


def mergeColumns(source_workbook, target_worksheet, regex=r".*", sourceStartRowOffset=0, sourceStartColOffset=0, startRowOffset=0, startColOffset=0,  sourceEndColOffset=0, sourceEndRowOffset=0, columnGap = 0, enableGrouping=False):

    print("Processing...")
    startCol = 1+startColOffset
    first = True
    selector_regex = re.compile(regex)
    itemList = list(filter(lambda i: selector_regex.match(i), source_workbook.sheetnames))

    for sn in itemList:
        if sn == 'TOC':
            continue
        source_worksheet = source_workbook[sn]  # Add Sheet name
        startRow = 1+startRowOffset

        ssColOffset = 0 if first else sourceStartColOffset
        first = False

        endCol = startCol + (source_worksheet.max_column -1 - sourceEndColOffset-ssColOffset)
        endRow = source_worksheet.max_row - sourceEndRowOffset
        print('sc:{} sr: {} ec: {} er: {} sn: {} '.format(startCol, startRow, endCol, endRow, sn))

        pasteRangeCols(startCol, startRow, endCol, endRow,
                                  target_worksheet, source_worksheet,
                       sourceStartRowOffset=sourceStartRowOffset, sourceStartColOffset=ssColOffset, enableGrouping=enableGrouping)
        enableGrouping = False
        startCol = endCol + 1 + columnGap
    # You can save the template as another file to create a new file here too.s
    print("Range copied and pasted!")

