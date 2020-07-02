from merge_sheet_by_rows import *
import pandas as pd



# File to be copied
workbook_base = 'Comp & Benefit Detail'
period = 'P5 2020'
workbook = '{} {}'.format(workbook_base, period)
workbook_url = 'report_samples/{}.xlsx'.format(workbook)
workbook_template = 'templates/{}.xltx' .format(workbook_base)
result_workbook = 'results/{}-{}-combined.xlsx'.format(workbook_base, period)



wb = openpyxl.load_workbook(workbook_url)  # Add file name
timeframes = ['MTD','QTD', 'YTD']

target = openpyxl.load_workbook(workbook_template)
target.template = False

metadata = target['Metadata'].values
columns = next(metadata)[0:]
metadata_df = pd.DataFrame(metadata, columns=columns)
logical_groups = metadata_df['LogicalGroup'].unique().tolist()

grandtotalRows = []
for timeframe in timeframes:
    worksheet = target[timeframe]
    first = True
    max_row = 8
    for logical_group in logical_groups:
        logical_group_df = metadata_df.loc[metadata_df.LogicalGroup==logical_group]
        subtotal_groups = logical_group_df['SubtotalGroup'].unique().tolist()
        for subtotal_group in subtotal_groups:
            row_filter = logical_group_df.loc[logical_group_df.SubtotalGroup.isnull()] if subtotal_group is None else \
                logical_group_df.loc[logical_group_df.SubtotalGroup == subtotal_group]
            regex_str = "|".join(str(x) + "-" + timeframe for x in row_filter['Sheet'].tolist())
            sub_total_group_flag = subtotal_group is not None
            grand_total_group_flag = subtotal_group and len(row_filter) > 1
            regex_pattern = r'^({})'.format(regex_str)
            selector_regex = re.compile(regex_pattern)
            print(regex_pattern)
            startRow = max_row
            basesheet, max_row = createMergedSheet(worksheet, selector_regex, wb, startCol=1, startRow=startRow+1, initialRowOffset=9,
                                                   postRowShrinkage=8, subtotalRows=sub_total_group_flag,
                                                   totalColOffset=6, groupRows=True, totalColOffsetUpperBound=26,
                                                   grandTotal=grand_total_group_flag, grandTotalTitle=subtotal_group)
            if(grand_total_group_flag) :
                grandtotalRows.append(worksheet[max_row])
            if len(subtotal_groups) == 1 :
                grandtotalRows.append(worksheet[startRow])
            if first:
                worksheet['A4'].value = basesheet['A4'].value
                worksheet['A6'].value = basesheet['K7'].value
                first = False
        add_separator(worksheet,startCol=1, endCol=26,row=max_row+1)
        max_row+=1


apply_grand_total(grandtotalRows, worksheet, max_row, 5 ,26, worksheet[10], grandTotalTitle="Total BHN" )
add_separator(worksheet,startCol=1, endCol=26,row=max_row)
target.save(result_workbook)