from merge_sheets_by_columns import  *
# Prepare the spreadsheets to copy from and paste too.

# File to be copied
workbook = 'A510_TB_Multiple BUs_LC_All_BUS'
workbook_url = 'report_samples/{}.xlsx'.format(workbook)
workbook_template = 'templates/{}.xltx' .format(workbook)
result_workbook = 'results/{}_combined.xlsx'.format(workbook)

wb = openpyxl.load_workbook(workbook_url)  # Add file name

target = openpyxl.load_workbook(workbook_template)
target.template = False
target_sheet = target['Sheet']
mergeColumns(wb, target_sheet)
target.save(result_workbook)
