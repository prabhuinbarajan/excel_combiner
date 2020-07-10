from merge_sheets_by_columns import  *
from config_reader import *
# Prepare the spreadsheets to copy from and paste too.
(TB_input_path,PL_input_path,template_path,TB_output_path,PL_output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)

# File to be copied
workbook_base = 'HA01_A510 BS Detail By Segment'
workbook = fnmatch.filter(os.listdir(TB_input_path), '*{}*'.format(workbook_base))
workbook1 = workbook[0]
print("File Names are " + workbook[0])
workbook_url = r'{}{}'.format(TB_input_path,workbook[0])
print("Workbook URLs are " + workbook_url)

workbook_template = r'{}{}.xltx' .format(template_path,workbook_base)
print("Template URL is " + workbook_template)

result_workbook = r'{}{}_combined.xlsx'.format(TB_output_path,workbook1.rsplit('.',1)[0])
print("Result workbook URLs are " + result_workbook)

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
             sourceEndColOffset=8, columnGap=2, enableRowGrouping=True, sourceEndRowOffset=8,applySourceOffsetFromFirst=False)
copy_data_in_range(target_worksheet=target_sheet, source_worksheet=target_sheet, reference_row=target_sheet[10], col_rng=range(57, 88),
                   row_range=range(11, target_sheet.max_row-1), copy_format_column=6)
acquisition_regex =  "^(BS Seg Total \([4-5]\))$"
mergeColumns(wb, target_sheet , regex=acquisition_regex, startRowOffset=9, startColOffset=90,  sourceStartRowOffset=9, sourceStartColOffset=5,
             sourceEndColOffset=8, columnGap=2, enableRowGrouping=True, sourceEndRowOffset=8, applySourceOffsetFromFirst=True)


copy_data_in_range(target_worksheet=target_sheet, source_worksheet=target_sheet, reference_row=target_sheet[10], col_rng=range(127, 136),
                   row_range=range(11, target_sheet.max_row-1), copy_format_column=6)

target.save(result_workbook)


#56 - 87
#126 -  135
