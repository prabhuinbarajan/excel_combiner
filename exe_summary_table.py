from merge_sheet_by_rows import *
from config_reader import *
import xlwings as xw
# import pandas as pd
# Prepare the spreadsheets to copy from and paste too.
(TB_input_path,PL_input_path,template_path,TB_output_path,PL_output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)

# File to be copied
workbook_base = fnmatch.filter(os.listdir(PL_input_path), '*PL37 Exe_Summary Table*')
workbook1 = workbook_base[0]

workbook_t19 = fnmatch.filter(os.listdir(PL_input_path), '*PL59_Dept Cost By Category*')
workbook_t18 = fnmatch.filter(os.listdir(PL_input_path), '*PL62_Dept Cost by Mgt Group*')
workbook_t17 = fnmatch.filter(os.listdir(PL_input_path), '*PL67_GAAP GA Line_BES*')
workbook_t20 = fnmatch.filter(os.listdir(PL_input_path), '*PL60_Dept Cost by Category by Seg*')
workbook_t21 = fnmatch.filter(os.listdir(PL_input_path), '*PL61_Dept Cost By GAAP Line*')
workbook_t22 = fnmatch.filter(os.listdir(PL_input_path), '*PL65_FTE ACT VS FCST*')
workbook_t23 = fnmatch.filter(os.listdir(PL_input_path), '*PL66_FTE CY vs PY*')
workbook_t24 = fnmatch.filter(os.listdir(PL_input_path), '*PL63_EMP HC ACT VS FCST*')
workbook_t25 = fnmatch.filter(os.listdir(PL_input_path), '*PL64_EMP HC CY vs PY*')
workbook_t26 = fnmatch.filter(os.listdir(PL_input_path), '*PL57_CNTR HC Act vs FCST*')
workbook_t27 = fnmatch.filter(os.listdir(PL_input_path), '*PL58_CNTR HC CY vs PY*')

print("File Names are " + workbook_t19[0] + ", " + workbook_t18[0] + ", " + workbook_t17[0] + ", " + workbook_t20[0]
      + ", " + workbook_t21[0] + ", " + workbook_t22[0] + ", " + workbook_t23[0] + ", " + workbook_t24[0]
      + ", " + workbook_t25[0] + ", " + workbook_t26[0] + ", " + workbook_t27[0])

workbook_url = r'{}{}'.format(PL_input_path,workbook_base[0])
workbook_t19_url = r'{}{}'.format(PL_input_path,workbook_t19[0])
workbook_t18_url = r'{}{}'.format(PL_input_path,workbook_t18[0])
workbook_t17_url = r'{}{}'.format(PL_input_path,workbook_t17[0])
workbook_t20_url = r'{}{}'.format(PL_input_path,workbook_t20[0])
workbook_t21_url = r'{}{}'.format(PL_input_path,workbook_t21[0])
workbook_t22_url = r'{}{}'.format(PL_input_path,workbook_t22[0])
workbook_t23_url = r'{}{}'.format(PL_input_path,workbook_t23[0])
workbook_t24_url = r'{}{}'.format(PL_input_path,workbook_t24[0])
workbook_t25_url = r'{}{}'.format(PL_input_path,workbook_t25[0])
workbook_t26_url = r'{}{}'.format(PL_input_path,workbook_t26[0])
workbook_t27_url = r'{}{}'.format(PL_input_path,workbook_t27[0])
print("Workbook URLs are " + workbook_t19_url + ", " + workbook_t18_url + ", " + workbook_t17_url
      + ", " + workbook_t20_url + ", " + workbook_t21_url + ", " + workbook_t22_url + ", " + workbook_t23_url
      + ", " + workbook_t24_url + ", " + workbook_t25_url + ", " + workbook_t26_url + ", " + workbook_t27_url)

workbook_template = r'{}{}.xltx' .format(template_path,'Exe_Summary Table')
print("Template URL is " + workbook_template)

result_workbook = r'{}{}_combined.xlsx'.format(PL_output_path,workbook1.rsplit('.',1)[0])
print("Result workbook URLs are " + result_workbook)

t19_wb = openpyxl.load_workbook(workbook_t19_url)  # Add file name
t18_wb = openpyxl.load_workbook(workbook_t18_url)  # Add file name
t17_wb = openpyxl.load_workbook(workbook_t17_url)  # Add file name
t20_wb = openpyxl.load_workbook(workbook_t20_url)  # Add file name
t21_wb = openpyxl.load_workbook(workbook_t21_url)  # Add file name
t22_wb = openpyxl.load_workbook(workbook_t22_url)  # Add file name
t23_wb = openpyxl.load_workbook(workbook_t23_url)  # Add file name
t24_wb = openpyxl.load_workbook(workbook_t24_url)  # Add file name
t25_wb = openpyxl.load_workbook(workbook_t25_url)  # Add file name
t26_wb = openpyxl.load_workbook(workbook_t26_url)  # Add file name
t27_wb = openpyxl.load_workbook(workbook_t27_url)  # Add file name

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
# Merge GAAP GA Line_S BES sheets into one
target_worksheet = target['Table17 GAAP GA Line_S BES']

ustotal_regex = re.compile(r'^US Total - GAAP Line_S BS')
ustotal_sheet, ustotal_max_row = createMergedSheet(target_worksheet, ustotal_regex, t17_wb, 1, 13, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
target_worksheet['A4'].value = ustotal_sheet['A3'].value
target_worksheet['A5'].value = ustotal_sheet['J7'].value

selector_regex = re.compile(r'^Us Core.*')
uscore_group = 'US Core'
basesheet, uscore_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, ustotal_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=uscore_group)
selector_regex = re.compile(r'^US Other.*')
usother_group = 'US Other'
basesheet, usother_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, uscore_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=usother_group)
selector_regex = re.compile(r'^America.*')
america_group = 'Americas Other'
basesheet, america_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, usother_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=america_group)
selector_regex = re.compile(r'^EMEA.*')
emea_group = 'EMEA'
basesheet, emea_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, america_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=emea_group)
selector_regex = re.compile(r'^ASIA.*')
asia_group = 'ASIA_PAC'
basesheet, asia_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, emea_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=asia_group)

target_worksheet.row_dimensions.group(start=ustotal_max_row+1, end=asia_max_row, hidden=True)

intotal_regex = re.compile(r'^International Total - GAAP Line')
intotal_sheet, intotal_max_row = createMergedSheet(target_worksheet, intotal_regex, t17_wb, 1, asia_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
bestotal_regex = re.compile(r'^BES Total - GAAP Line_S BS')
bestotal_sheet, bestotal_max_row = createMergedSheet(target_worksheet, bestotal_regex, t17_wb, 1, intotal_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
selector_regex = re.compile(r'^BES ISP.*')
besisp_group = 'BES ISP'
basesheet, besisp_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, bestotal_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besisp_group)
selector_regex = re.compile(r'^Incentec.*')
besinc_group = 'BES Incentec'
basesheet, besinc_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, besisp_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besinc_group)
selector_regex = re.compile(r'^Parago.*')
besparago_group = 'BES Parago'
basesheet, besparago_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, besinc_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besparago_group)
selector_regex = re.compile(r'^BES Elim.*')
beselim_group = 'BES Elim'
basesheet, beselim_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, besparago_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=beselim_group)
selector_regex = re.compile(r'^SVM.*')
bessvm_group = 'BES SVM'
basesheet, bessvm_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, beselim_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=bessvm_group)
selector_regex = re.compile(r'^GC.*')
besgc_group = 'BES GC.COM'
basesheet, besgc_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, bessvm_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besgc_group)
selector_regex = re.compile(r'^Touchpoint.*')
bestouch_group = 'BES Touchpoint'
basesheet, bestouch_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, besgc_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=bestouch_group)
selector_regex = re.compile(r'^Extra Measures.*')
besextra_group = 'BES Extra Measures'
basesheet, besextra_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, bestouch_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besextra_group)
selector_regex = re.compile(r'^Loyalty.*')
besloyalty_group = 'BES Loyalty'
basesheet, besloyalty_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, besextra_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besloyalty_group)
selector_regex = re.compile(r'^Achievers (?!Total).*')
besach_group = 'ACHIEVERS'
basesheet, besach_max_row = createMergedSheet(target_worksheet, selector_regex, t17_wb, 1, besloyalty_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besach_group)


target_worksheet.row_dimensions.group(start=bestotal_max_row+1, end=besach_max_row, hidden=True)

achtotal_regex = re.compile(r'^Achievers Total - GAAP Line_S B')
achtotal_sheet, achtotal_max_row = createMergedSheet(target_worksheet, achtotal_regex, t17_wb, 1, besach_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
bhntotal_regex = re.compile(r'^Blackhawk Total - GAAP Line_S B')
bhntotal_sheet, bhntotal_max_row = createMergedSheet(target_worksheet, bhntotal_regex, t17_wb, 1, achtotal_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)

######   Start of Processing PL60_Dept Cost by Category by Seg
target_worksheet = target['Dept Cost by Category by Seg']

ustotal_regex = re.compile(r'^Us Total - Dept Cost By Categor')
ustotal_sheet, ustotal_max_row = createMergedSheet(target_worksheet, ustotal_regex, t20_wb, 1, 13, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
target_worksheet['A4'].value = ustotal_sheet['A3'].value
target_worksheet['A5'].value = ustotal_sheet['J7'].value

selector_regex = re.compile(r'^Us Core.*')
uscore_group = 'US Core'
basesheet, uscore_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, ustotal_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=uscore_group)
selector_regex = re.compile(r'^US Other.*')
usother_group = 'US Other'
basesheet, usother_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, uscore_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=usother_group)
selector_regex = re.compile(r'^America.*')
america_group = 'Americas Other'
basesheet, america_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, usother_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=america_group)
selector_regex = re.compile(r'^EMEA.*')
emea_group = 'EMEA'
basesheet, emea_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, america_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=emea_group)
selector_regex = re.compile(r'^ASIA.*')
asia_group = 'ASIA_PAC'
basesheet, asia_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, emea_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=asia_group)

target_worksheet.row_dimensions.group(start=ustotal_max_row+1, end=asia_max_row, hidden=True)

intotal_regex = re.compile(r'^International Total - Dept Cost')
intotal_sheet, intotal_max_row = createMergedSheet(target_worksheet, intotal_regex, t20_wb, 1, asia_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
bestotal_regex = re.compile(r'^BES Total - Dept Cost By Catego')
bestotal_sheet, bestotal_max_row = createMergedSheet(target_worksheet, bestotal_regex, t20_wb, 1, intotal_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
selector_regex = re.compile(r'^BES ISP.*')
besisp_group = 'BES ISP'
basesheet, besisp_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, bestotal_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besisp_group)
selector_regex = re.compile(r'^Incentec.*')
besinc_group = 'BES Incentec'
basesheet, besinc_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, besisp_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besinc_group)
selector_regex = re.compile(r'^Parago.*')
besparago_group = 'BES Parago'
basesheet, besparago_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, besinc_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besparago_group)
selector_regex = re.compile(r'^BES Elim.*')
beselim_group = 'BES Elim'
basesheet, beselim_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, besparago_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=beselim_group)
selector_regex = re.compile(r'^SVM.*')
bessvm_group = 'BES SVM'
basesheet, bessvm_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, beselim_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=bessvm_group)
selector_regex = re.compile(r'^GC.*')
besgc_group = 'BES GC.COM'
basesheet, besgc_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, bessvm_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besgc_group)
selector_regex = re.compile(r'^Touchpoint.*')
bestouch_group = 'BES Touchpoint'
basesheet, bestouch_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, besgc_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=bestouch_group)
selector_regex = re.compile(r'^Extra Measures.*')
besextra_group = 'BES Extra Measures'
basesheet, besextra_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, bestouch_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besextra_group)
selector_regex = re.compile(r'^Loyalty.*')
besloyalty_group = 'BES Loyalty'
basesheet, besloyalty_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, besextra_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besloyalty_group)
selector_regex = re.compile(r'^Achievers (?!Total).*')
besach_group = 'ACHIEVERS'
basesheet, besach_max_row = createMergedSheet(target_worksheet, selector_regex, t20_wb, 1, besloyalty_max_row+1, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=besach_group)


target_worksheet.row_dimensions.group(start=bestotal_max_row+1, end=besach_max_row, hidden=True)

achtotal_regex = re.compile(r'^Achievers Total - GAAP Line_S B')
achtotal_sheet, achtotal_max_row = createMergedSheet(target_worksheet, achtotal_regex, t20_wb, 1, besach_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
bhntotal_regex = re.compile(r'^Blackhawk Total - Dept Cost By ')
bhntotal_sheet, bhntotal_max_row = createMergedSheet(target_worksheet, bhntotal_regex, t20_wb, 1, achtotal_max_row+1, 9, 8,
                                       subtotalRows=False,
                                       totalColOffset=5, groupRows=False, grandTotal=False)
######   End of processing PL60_Dept Cost by Category by Seg
######   Start of Processing PL61_Dept Cost By GAAP Line
target_worksheet = target['Dept Cost by GAAP Line']

selector_regex = re.compile(r'^(Technology|Ops, CC, Merch, Risk)')
tps_group = 'Total Processing and Service - Dept. Costs'
tpssheet, tps_max_row = createMergedSheet(target_worksheet, selector_regex, t21_wb, 1, 13, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=tps_group)
selector_regex = re.compile(r'^(Cardpool & Product)')
tcp_group = 'Total Costs of Products Sold - Dept. Costs'
tcpsheet, tcp_max_row = createMergedSheet(target_worksheet, selector_regex, t21_wb, 1, tps_max_row+2, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=tcp_group)
selector_regex = re.compile(r'^(Prod Mgt & Bus Dev|Marketing|US Sales|BES Sales|International Sales|International Sales - US)')
tsm_group = 'Total Sales and Marketing - Dept. Costs'
tsmsheet, tsm_max_row = createMergedSheet(target_worksheet, selector_regex, t21_wb, 1, tcp_max_row+2, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=tsm_group)
selector_regex = re.compile(r'^(Legal|Acctg & Finance|HR|It Admin|Executive|Other|Bad Debt)')
tga_group = 'Total General and Admin - Dept. Costs'
tgasheet, tga_max_row = createMergedSheet(target_worksheet, selector_regex, t21_wb, 1, tsm_max_row+2, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=tga_group)

target_worksheet.cell(row=tga_max_row+2, column=1).value = "TOTAL BLACKHAWK"
for j in range(6, target_worksheet.max_column, 1):
    target_cell = target_worksheet.cell(row=tga_max_row+2, column=j)
    source = target_worksheet.cell(row=tga_max_row - 1, column=j)
    if source.data_type == 'f':
        target_cell.value = Translator(source.value, source.coordinate).translate_formula(target_cell.coordinate)
    else:
        target_cell.value = "=SUM({}+{}+{}+{})".format(target_worksheet.cell(row=tps_max_row, column=j).coordinate,
                                                    target_worksheet.cell(row=tcp_max_row, column=j).coordinate,
                                                    target_worksheet.cell(row=tsm_max_row, column=j).coordinate,
                                                    target_worksheet.cell(row=tga_max_row, column=j).coordinate)
    copy_style(source, target_cell)

target_worksheet['A4'].value = tpssheet['A3'].value
target_worksheet['A5'].value = tpssheet['J7'].value
######   End of Processing PL61_Dept Cost By GAAP Line
######   Start of Processing FTE, EMP & CNTR HC
ftetimeframes = ['FTE ACT vs FCST','FTE CY vs PY','EMP ACT vs FCST','EMP CY vs PY','CNTR ACT vs FCST','CNTR CY vs PY']
fteselector_regex = re.compile(r'^(?!(TOC))')

for ftetimeframe in ftetimeframes:
    target_worksheet = target[ftetimeframe]
    first = True
    if ftetimeframe == 'FTE ACT vs FCST':
        source_wb = t22_wb
        grouptitle = 'Total FTE HC'
    if ftetimeframe == 'FTE CY vs PY':
        source_wb = t23_wb
        grouptitle = 'Total FTE HC'
    if ftetimeframe == 'EMP ACT vs FCST':
        source_wb = t24_wb
        grouptitle = 'Total Employee HC'
    if ftetimeframe == 'EMP CY vs PY':
        source_wb = t25_wb
        grouptitle = 'Total Employee HC'
    if ftetimeframe == 'CNTR ACT vs FCST':
        source_wb = t26_wb
        grouptitle = 'Total Contractor HC'
    if ftetimeframe == 'CNTR CY vs PY':
        source_wb = t27_wb
        grouptitle = 'Total Contractor HC'
    basesheet, max_row = createMergedSheet(target_worksheet, fteselector_regex, source_wb, 1, 10, 9, 8, subtotalRows=True,
                              totalColOffset=5, groupRows=True, grandTotal=True, grandTotalTitle=grouptitle)
    target_worksheet['A3'].value = basesheet['A3'].value
    target_worksheet['A4'].value = basesheet['A4'].value
    target_worksheet['A5'].value = basesheet['T7'].value
######   End of Processing FTE, EMP & CNTR HC
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