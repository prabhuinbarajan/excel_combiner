from merge_sheet_by_rows import *
from config_reader import *
import xlwings as xw
# Prepare the spreadsheets to copy from and paste too.
(TB_input_path,PL_input_path,template_path,TB_output_path,PL_output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)

# File to be copied
workbook_base = fnmatch.filter(os.listdir(PL_input_path), '*PL37 Exe_Summary Table*')
workbook1 = workbook_base[0]

workbook_t19 = fnmatch.filter(os.listdir(PL_input_path), '*PL59_Dept Cost By Category*')
workbook_t18 = fnmatch.filter(os.listdir(PL_input_path), '*PL62_Dept Cost by Mgt Group*')
workbook_t17 = fnmatch.filter(os.listdir(PL_input_path), '*PL67_GAAP GA Line_BES*')

print("File Names are " + workbook_t19[0] + ", " + workbook_t18[0] + ", " + workbook_t17[0])

workbook_url = r'{}{}'.format(PL_input_path,workbook_base[0])
workbook_t19_url = r'{}{}'.format(PL_input_path,workbook_t19[0])
workbook_t18_url = r'{}{}'.format(PL_input_path,workbook_t18[0])
workbook_t17_url = r'{}{}'.format(PL_input_path,workbook_t17[0])
print("Workbook URLs are " + workbook_t19_url + ", " + workbook_t18_url + ", " + workbook_t17_url)

workbook_template = r'{}{}.xltx' .format(template_path,'Exe_Summary Table')
print("Template URL is " + workbook_template)

result_workbook = r'{}{}_combined.xlsx'.format(PL_output_path,workbook1.rsplit('.',1)[0])
print("Result workbook URLs are " + result_workbook)

t19_wb = openpyxl.load_workbook(workbook_t19_url)  # Add file name
t18_wb = openpyxl.load_workbook(workbook_t18_url)  # Add file name
t17_wb = openpyxl.load_workbook(workbook_t17_url)  # Add file name

timeframes = ['Dept Cost by Category','Dept Cost by Mgt Group']
selector_regex = re.compile(r'^(?!(TOC|Stock expense, 401K match and E|Bad Debt))')
stock_selector_regex = re.compile(r'^Stock expense, 401K match and E.*')
baddebt_selector_regex = re.compile(r'^Bad Debt.*')


target = openpyxl.load_workbook(workbook_template)
target.template = False
for timeframe in timeframes:
    target_worksheet = target[timeframe]
    first = True
    if timeframe == 'Dept Cost by Category':
        source_wb = t19_wb
        startrow = 13
        initialRowOffset = 10
    if timeframe == 'Dept Cost by Mgt Group':
        source_wb = t18_wb
        startrow = 14
        initialRowOffset = 11
    basesheet, max_row = createMergedSheet(target_worksheet, selector_regex, source_wb, 1, startrow, initialRowOffset, 8, subtotalRows=True,
                              totalColOffset=6, groupRows=True, grandTotal=True)
    target_worksheet['A4'].value = basesheet['A3'].value
    if timeframe == 'Dept Cost by Mgt Group':
        target_worksheet['A5'].value = basesheet['J11'].value
    else:
        target_worksheet['A5'].value = basesheet['J10'].value
    totalRow = max_row

    stocks_sheet,stock_max_row = createMergedSheet(target_worksheet, stock_selector_regex, source_wb, 1, totalRow+1, initialRowOffset, 8, subtotalRows=True,
                              totalColOffset=6, groupRows=True)

    debt_sheet,dbt_max_row = createMergedSheet(target_worksheet, baddebt_selector_regex, source_wb, 1, stock_max_row, initialRowOffset, 8, subtotalRows=True,
                              totalColOffset=6, groupRows=True)

    target_worksheet.cell(row=dbt_max_row, column=1).value = "TOTAL BLACKHAWK"

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

wb1 = xw.Book(workbook_url)
wb2 = xw.Book(result_workbook)

# copying 11 sheets into combined workbook from the first collection of exec summary table report
ws1 = wb1.sheets('Sales&Marketing')
ws1.api.Copy(Before=wb2.sheets(1).api)
ws2 = wb1.sheets('Proc&Srv')
ws2.api.Copy(Before=wb2.sheets(1).api)
ws3 = wb1.sheets('PDC')
ws3.api.Copy(Before=wb2.sheets(1).api)
ws4 = wb1.sheets('ProdSales')
ws4.api.Copy(Before=wb2.sheets(1).api)
ws5 = wb1.sheets('ProgFee')
ws5.api.Copy(Before=wb2.sheets(1).api)
ws6 = wb1.sheets('ISS & Others')
ws6.api.Copy(Before=wb2.sheets(1).api)
ws7 = wb1.sheets('PMF')
ws7.api.Copy(Before=wb2.sheets(1).api)
ws8 = wb1.sheets('NDC')
ws8.api.Copy(Before=wb2.sheets(1).api)
ws9 = wb1.sheets('GDC')
ws9.api.Copy(Before=wb2.sheets(1).api)
ws10 = wb1.sheets('Comm&Fee')
ws10.api.Copy(Before=wb2.sheets(1).api)
ws11 = wb1.sheets('TDV')
ws11.api.Copy(Before=wb2.sheets(1).api)
wb2.save()
wb2.app.quit()