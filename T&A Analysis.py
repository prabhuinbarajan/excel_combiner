from merge_sheet_by_rows import *
import pandas as pd
from config_reader import *
from win32com.client import Dispatch


(TB_input_path,PL_input_path,template_path,TB_output_path,PL_output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)

# File to be copied
workbook_base = 'T&A Analysis'
workbook = fnmatch.filter(os.listdir(PL_input_path), '*{}*'.format(workbook_base))
workbook1 = workbook[0]
print("File Names are " + workbook[0])
workbook_url = r'{}{}'.format(PL_input_path,workbook[0])
print("Workbook URLs are " + workbook_url)
workbook_template = r'{}{}.xltx' .format(template_path,workbook_base)
print("Template URL is " + workbook_template)
result_workbook = r'{}{}_combined.xlsx'.format(PL_output_path,workbook1.rsplit('.',1)[0])
print("Result workbook URLs are " + result_workbook)



wb = openpyxl.load_workbook(workbook_url)  # Add file name
timeframes = ['XCEL']

target = openpyxl.load_workbook(workbook_template)
target.template = False

metadata = target['Metadata'].values
columns = next(metadata)[0:]
metadata_df = pd.DataFrame(metadata, columns=columns)
logical_groups = metadata_df['LogicalGroup'].unique().tolist()

for timeframe in timeframes:
    grandtotalRows = []
    balancecheckRows = []
    worksheet = target[timeframe]
    first = True
    max_row = 8
    for logical_group in logical_groups:
        logical_group_df = metadata_df.loc[metadata_df.LogicalGroup==logical_group]
        subtotal_groups = logical_group_df['SubtotalGroup'].unique().tolist()
        for subtotal_group in subtotal_groups:
            row_filter = logical_group_df.loc[logical_group_df.SubtotalGroup.isnull()] if subtotal_group is None else \
                logical_group_df.loc[logical_group_df.SubtotalGroup == subtotal_group]
            regex_str = "|".join(timeframe + " - " + str(x) for x in row_filter['Sheet'].tolist())
            sub_total_group_flag = subtotal_group is not None
            grand_total_group_flag = subtotal_group and len(row_filter) > 1
            regex_pattern = r'^({})'.format(regex_str)
            selector_regex = re.compile(regex_pattern)
            print(regex_pattern)
            startRow = max_row + 1
            if row_filter['Sheet'].tolist()[0] != 'XCEL Check':
                basesheet, max_row = createMergedSheet(worksheet, selector_regex, wb, startCol=1, startRow=startRow, initialRowOffset=9,
                                                   postRowShrinkage=8, subtotalRows=sub_total_group_flag,
                                                   totalColOffset=5, groupRows=True, totalColOffsetUpperBound=100,
                                                   grandTotal=grand_total_group_flag, grandTotalTitle=subtotal_group)
                grandtotalRows.append(worksheet[startRow])
            else :
                apply_grand_total(grandtotalRows, worksheet, max_row, 5, 100, worksheet[10],grandTotalTitle="Total BHN")
                balancecheckRows.append(worksheet[max_row])
                add_separator(worksheet, startCol=1, endCol=100, row=max_row + 1)
                add_separator(worksheet, startCol=1, endCol=100, row=max_row + 2)
                add_separator(worksheet, startCol=1, endCol=100, row=max_row + 3)
                add_separator(worksheet, startCol=1, endCol=100, row=max_row + 4)
                add_separator(worksheet, startCol=1, endCol=100, row=max_row + 5)
                max_row += 5
                startRow = max_row + 1
                basesheet, max_row = createMergedSheet(worksheet, selector_regex, wb, startCol=1, startRow=startRow, initialRowOffset=9,
                                                   postRowShrinkage=8, subtotalRows=sub_total_group_flag,
                                                   totalColOffset=5, groupRows=True, totalColOffsetUpperBound=100,
                                                   grandTotal=grand_total_group_flag, grandTotalTitle=subtotal_group)
                balancecheckRows.append(worksheet[max_row-2])
                balancecheck_start_row = max_row-2
            if first:
                worksheet['A5'].value = basesheet['F5'].value
                worksheet['A3'].value = basesheet['A3'].value
                first = False
        add_separator(worksheet,startCol=1, endCol=100,row=max_row+1)
        max_row+=1
    apply_grand_total(balancecheckRows, worksheet, max_row, 5 ,100, worksheet[10], grandTotalTitle="BalanceCheck" )
    add_separator(worksheet, startCol=1, endCol=26, row=max_row)
    worksheet.row_dimensions.group(start=balancecheck_start_row, end=max_row, hidden=True)
target.save(result_workbook)

xl = Dispatch("Excel.Application")
xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

wb1 = xl.Workbooks.Open(Filename=workbook_url)
wb2 = xl.Workbooks.Open(Filename=result_workbook)

ws1 = wb1.Worksheets('XCEL Project')
ws1.Copy(Before=wb2.Worksheets(1))

ws2 = wb1.Worksheets('BHN Acquisition')
ws2.Copy(Before=wb2.Worksheets(1))

ws3 = wb1.Worksheets('T&A Analysis')
ws3.Copy(Before=wb2.Worksheets(1))

wb1.Close(SaveChanges=False)
wb2.Close(SaveChanges=True)
xl.Quit()