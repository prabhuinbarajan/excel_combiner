from merge_sheets_by_columns import  *
from config_reader import *
# Prepare the spreadsheets to copy from and paste too.
(input_path,template_path,output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)
# File to be copied
workbook_base = 'A510_TB_Multiple BUs_LC_All Bus'
workbook = fnmatch.filter(os.listdir(input_path), '*A510_TB_Multiple BUs_LC_All Bus*')
workbook1 = workbook[0]
print("File Names are " + workbook[0])
workbook_url = r'{}{}'.format(input_path,workbook[0])
print("Workbook URLs are " + workbook_url)
workbook_template = r'{}{}.xltx' .format(template_path,workbook_base)
print("Template URL is " + workbook_template)
result_workbook = r'{}{}_combined.xlsx'.format(output_path,workbook1.rsplit('.',1)[0])
print("Result workbook URLs are " + result_workbook)

wb = openpyxl.load_workbook(workbook_url)  # Add file name

target = openpyxl.load_workbook(workbook_template)
target.template = False
target_sheet = target['Sheet']
mergeColumns(wb, target_sheet)
target.save(result_workbook)
