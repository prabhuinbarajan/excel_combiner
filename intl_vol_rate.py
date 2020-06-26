from merge_sheet_by_rows import *

# Prepare the spreadsheets to copy from and paste too.

# File to be copied
workbook_base = 'Intl Vol Rate MTD QTD'

mtd_workbook = '{} P4 2020'.format(workbook_base)
ytd_workbook = 'Intl Vol Rate YTD P4 2020'
mtd_workbook_url = 'report_samples/{}.xlsx'.format(mtd_workbook)
ytd_workbook_url = 'report_samples/{}.xlsx'.format(ytd_workbook)

workbook_template = 'templates/{}.xltx' .format(workbook_base)
result_workbook = 'results/{}_combined.xlsx'.format(mtd_workbook)

mtd_qtd_wb = openpyxl.load_workbook(mtd_workbook_url)  # Add file name
ytd_wb = openpyxl.load_workbook(ytd_workbook_url)  # Add file name

mtd_regex = re.compile(r'^.*\-MTD')
qtd_regex = re.compile(r'^.*\-QTD')
ytd_regex = re.compile(r'^.*\-YTD')



target = openpyxl.load_workbook(workbook_template)
target.template = False
#target = copy(wb_template)
createMergedSheet(target['YTD'], ytd_regex, ytd_wb, 1, 10, 9, 8)

createMergedSheet(target['MTD'], mtd_regex, mtd_qtd_wb, 1, 10, 9, 8)
createMergedSheet(target['QTD'], qtd_regex, mtd_qtd_wb, 1, 10, 9, 8)


target.save(result_workbook)
