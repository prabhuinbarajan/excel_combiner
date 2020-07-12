from merge_sheets_by_columns import  *
import pandas as pd
from config_reader import *

(TB_input_path,PL_input_path,template_path,TB_output_path,PL_output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)

# File to be copied
workbook_base = 'A510-BSTNDLC Detail - by Account Category'
workbook = fnmatch.filter(os.listdir(TB_input_path), '*{}*'.format(workbook_base))
workbook1 = workbook[0]
print("File Names are " + workbook[0])
workbook_url = r'{}{}'.format(TB_input_path,workbook[0])
print("Workbook URLs are " + workbook_url)
workbook_template = r'{}{}.xltx' .format(template_path,workbook_base)
print("Template URL is " + workbook_template)
result_workbook = r'{}{}_combined.xlsx'.format(TB_output_path,workbook1.rsplit('.',1)[0])
print("Result workbook URLs are " + result_workbook)



source_workbook = openpyxl.load_workbook(workbook_url)  # Add file name


target = openpyxl.load_workbook(workbook_template)
target.template = False

metadata = target['Metadata'].values
columns = next(metadata)[0:]
metadata_df = pd.DataFrame(metadata, columns=columns)
logical_groups = ['Assets','Liabilities','Equities']
grandtotalRows = []
target_sheet = target['Sheet']
source_header_worksheet = source_workbook.worksheets[1]


copy_data_in_range(target_worksheet=target_sheet, source_worksheet=source_header_worksheet, reference_row=source_header_worksheet[6], col_rng=range(7, source_header_worksheet.max_column - 11),
                   row_range=range(9, 9))

first = True
max_row = 9

for logical_group in logical_groups:

    row_filter = metadata_df.loc[metadata_df.LogicalGroup.isnull()] if logical_group is None else \
        metadata_df.loc[metadata_df.LogicalGroup == logical_group]
    regex_str = "|".join(str(x[:31]).replace("(", "\(").replace(")", "\)")  for x in row_filter['Sheet Name'].tolist())
    regex_pattern = r'^({})$'.format(regex_str)
    natural_state = row_filter.iloc[0].values[2]
    selector_regex = re.compile(regex_pattern)
    print(regex_pattern)

    startRow = max_row

    basesheet, max_row = createMergedSheet(target_sheet, selector_regex, source_workbook, startCol=1, startRow=startRow + 1, initialRowOffset=9,
                                           postRowShrinkage=8, subtotalRows=False,
                                           totalColOffset=8, groupRows=True,
                                           grandTotal=True, grandTotalTitle=logical_group,sourceColEndOffset=10,natural_state=natural_state)
    grandtotalRows.append(max_row)
    if first:
        target_sheet['A5'].value = basesheet['A4'].value
        first = False
    #add_separator(target_sheet, startCol=1, endCol=26, row=max_row + 1)
    max_row+=1
    #apply_grand_total(grandtotalRows, target_sheet, max_row, 5, 26, target_sheet[10], grandTotalTitle="Total BHN")
    #add_separator(target_sheet, startCol=1, endCol=26, row=max_row)
checkfigure_list = []
target_sheet.cell(max_row+1,1).value = 'Check Figure'
for col in range(8, target_sheet.max_column):

    col_formulae = ''
    for k in grandtotalRows:
        grand_total_cell = target_sheet.cell(k,col).coordinate
        col_formulae = grand_total_cell if col_formulae == '' else '{},{}'.format(col_formulae,grand_total_cell)
    target_sheet.cell(max_row+1,col).value = '=sum({})'.format(col_formulae)
target.save(result_workbook)