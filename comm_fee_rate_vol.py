from merge_sheet_by_rows import *
from config_reader import *

# Prepare the spreadsheets to copy from and paste too.

def apply_match_to_sheet(source, destination) :
    match = re.search('(.*)Actual vs(.*)-(.*)', source['A4'].value)
    destination['A7'].value = match.group(1) if match else None
    destination['A6'].value = match.group(2) if match else None
    #destination['A8'].value = match.group(3) if match else None

(input_path,template_path,output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)
# File to be copied

mtd_workbook = fnmatch.filter(os.listdir(input_path), '*Comm and Fee Rate Volume-MTD*')
qtd_workbook = fnmatch.filter(os.listdir(input_path), '*Comm and Fee Rate Volume-QTD*')
ytd_workbook = fnmatch.filter(os.listdir(input_path), '*Comm and Fee Rate Volume-YTD*')
print("File Names are " + mtd_workbook[0] + ", " + qtd_workbook[0] + ", " + ytd_workbook[0])

workbook_base = 'Comm and Fee Rate Volume'

mtd_workbook_url = r'{}{}'.format(input_path,mtd_workbook[0])
ytd_workbook_url = r'{}{}'.format(input_path,ytd_workbook[0])
qtd_workbook_url = r'{}{}'.format(input_path,qtd_workbook[0])
print("Workbook URLs are " + mtd_workbook_url + ", " + ytd_workbook_url + ", " + qtd_workbook_url)

workbook_template = r'{}{}.xltx' .format(template_path,workbook_base)
print("Template URL is " + workbook_template)

result_workbook = r'{}P{} {} HA06_{}_combined.xlsx'.format(output_path,myper,myyear,workbook_base)
print("Result workbook URLs are " + result_workbook)

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
