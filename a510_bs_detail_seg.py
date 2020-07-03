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
target_sheet = target['BS Seg Total']
mergeColumns(wb, target_sheet)
target.save(result_workbook)
