from merge_sheets_by_columns import  *
# Prepare the spreadsheets to copy from and paste too.

# File to be copied
period = 'P5 2020'
workbook_name = 'A510-BS-Seg-Total {}'.format(period)
template_name =  'A510-BS-Seg'
workbook_url = 'report_samples/{}.xlsx'.format(workbook_name,  period)
workbook_template = 'templates/{}.xltx' .format(template_name)
result_workbook = 'results/{}_combined.xlsx'.format(workbook_name)

wb = openpyxl.load_workbook(workbook_url)  # Add file name

target = openpyxl.load_workbook(workbook_template)
target.template = False

segment_groups = ["BS BU Seg CY", "BS BU Seg PY", "BS BU Seg Last Period PY"]
#segment_groups = ["BS BU Seg CY"]
for segment_group in segment_groups:
    seg_cy_regex =  "^({} \([0-9]*\))$".format(segment_group)
    target_sheet = target[segment_group]
    print (seg_cy_regex)
    mergeColumns(wb, target_sheet, regex=seg_cy_regex, startRowOffset=8, sourceStartRowOffset=4, sourceStartColOffset=5,
                 sourceEndColOffset=7, enableRowGrouping=True, enableColGrouping=True,  sourceEndRowOffset=10, applySourceOffsetFromFirst=False)
    target_sheet.column_dimensions.group(start='B', end='E', hidden=True, outline_level=3)

#target.save(result_workbook)


target_sheet = target['BS Seg Total']
target_sheet['B3'].value = wb['BS Seg Total (1)']['F5'].value
target_sheet['B4'].value = wb['BS Seg Total (2)']['F5'].value
target_sheet['B5'].value = wb['BS Seg Total (3)']['F5'].value


pl_regex = "^(BS Seg Total \([1-3]\))$"
mergeColumns(wb, target_sheet , regex=pl_regex, startRowOffset=9, sourceStartRowOffset=9, sourceStartColOffset=5,
             sourceEndColOffset=8, columnGap=2, enableRowGrouping=True, sourceEndRowOffset=7,applySourceOffsetFromFirst=False)
copy_data_in_range(worksheet=target_sheet, reference_row=target_sheet[10], col_rng=range(57,88),
                   row_range=range(11, target_sheet.max_row-1),  copy_format_column=6)
acquisition_regex =  "^(BS Seg Total \([4-5]\))$"
mergeColumns(wb, target_sheet , regex=acquisition_regex, startRowOffset=9, startColOffset=90,  sourceStartRowOffset=9, sourceStartColOffset=5,
             sourceEndColOffset=8, columnGap=2, enableRowGrouping=True, sourceEndRowOffset=7, applySourceOffsetFromFirst=True)


copy_data_in_range(worksheet=target_sheet, reference_row=target_sheet[10], col_rng=range(127,136),
                   row_range=range(11, target_sheet.max_row-1),  copy_format_column=6)

target.save(result_workbook)


#56 - 87
#126 -  135
