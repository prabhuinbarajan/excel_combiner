from merge_sheet_by_rows import *

# Prepare the spreadsheets to copy from and paste too.

# File to be copied
workbook_base = 'Dept Cost By Category'
period = 'P5 2020'

workbook = '{} {}'.format(workbook_base, period)


workbook_url = 'report_samples/{}.xlsx'.format(workbook)


workbook_template = 'templates/{}.xltx' .format(workbook_base)
result_workbook = 'results/{}-{}-combined.xlsx'.format(workbook_base, period)

wb = openpyxl.load_workbook(workbook_url)  # Add file name


selector_regex = re.compile(r'^(?!(TOC|Stock expense, 401K match and E|Bad Debt))')
stock_selector_regex = re.compile(r'^Stock expense, 401K match and E.*')
baddebt_selector_regex = re.compile(r'^Bad Debt.*')


target = openpyxl.load_workbook(workbook_template)
target.template = False
target_worksheet = target['Dept Cost by Category']
basesheet, max_row = createMergedSheet(target_worksheet, selector_regex, wb, 1, 13, 10, 8, subtotalRows=True,
                              totalColOffset=6, groupRows=True, grandTotal=True)
target_worksheet['A4'].value = basesheet['A3'].value
target_worksheet['A5'].value = basesheet['J10'].value
totalRow = max_row

stocks_sheet,stock_max_row = createMergedSheet(target_worksheet, stock_selector_regex, wb, 1, totalRow+1, 10, 8, subtotalRows=True,
                              totalColOffset=6, groupRows=True)

debt_sheet,dbt_max_row = createMergedSheet(target_worksheet, baddebt_selector_regex, wb, 1, stock_max_row, 10, 8, subtotalRows=True,
                              totalColOffset=6, groupRows=True)

target_worksheet.cell(row=dbt_max_row, column=1).value = "Grand Total"

for j in range(6, target_worksheet.max_column, 1):
    target_cell = target_worksheet.cell(row=dbt_max_row, column=j)
    source = target_worksheet.cell(row=dbt_max_row - 2, column=j)
    if source.data_type == 'f':
        target_cell.value = Translator(source.value, source.coordinate).translate_formula(target_cell.coordinate)
    else:
        target_cell.value = "=SUM({}+{}+{})".format(target_worksheet.cell(row=totalRow,column=j).coordinate,
                                               target_worksheet.cell(row=totalRow+1,column=j).coordinate,
                                               target_worksheet.cell(row=stock_max_row , column=j).coordinate)
    copy_style(source, target_cell)
target.save(result_workbook)
