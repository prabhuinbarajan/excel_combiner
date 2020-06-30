from merge_sheet_by_rows import *

# Prepare the spreadsheets to copy from and paste too.

def apply_match_to_sheet(source, destination) :
    match = re.search('(.*) Actual vs (.*) - (.*)', source['A4'].value)
    destination['A7'].value = match.group(1) if match else None
    destination['A6'].value = match.group(2) if match else None
    #destination['A8'].value = match.group(3) if match else None


# File to be copied
workbook_base = 'Comm and Fee Rate Volume'
period = 'P5 2020'

mtd_workbook = '{}-MTD-{}'.format(workbook_base, period)
qtd_workbook = '{}-QTD-{}'.format(workbook_base, period)
ytd_workbook = '{}-YTD-{}'.format(workbook_base, period)

mtd_workbook_url = 'report_samples/{}.xlsx'.format(mtd_workbook)
qtd_workbook_url = 'report_samples/{}.xlsx'.format(qtd_workbook)
ytd_workbook_url = 'report_samples/{}.xlsx'.format(ytd_workbook)

workbook_template = 'templates/{}.xltx' .format(workbook_base)
result_workbook = 'results/{}-{}-combined.xlsx'.format(workbook_base, period)

mtd_wb = openpyxl.load_workbook(mtd_workbook_url)  # Add file name
qtd_wb = openpyxl.load_workbook(qtd_workbook_url)  # Add file name
ytd_wb = openpyxl.load_workbook(ytd_workbook_url)  # Add file name

mtd_regex = re.compile(r'^.*\-MTD')
qtd_regex = re.compile(r'^.*\-QTD')
ytd_regex = re.compile(r'^.*\-YTD')


target = openpyxl.load_workbook(workbook_template)
target.template = False
#target = copy(wb_template)
basesheet,er = createMergedSheet(target['YTD'], ytd_regex, ytd_wb, 1, 13, 9, 9 )
apply_match_to_sheet(basesheet, target['YTD'])
basesheet,er=createMergedSheet(target['MTD'], mtd_regex, mtd_wb, 1, 13, 9, 9)
apply_match_to_sheet(basesheet, target['MTD'])
basesheet,er=createMergedSheet(target['QTD'], qtd_regex, qtd_wb, 1, 13, 9, 9)
apply_match_to_sheet(basesheet, target['QTD'])


target.save(result_workbook)
