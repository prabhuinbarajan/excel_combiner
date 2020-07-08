from merge_sheet_by_rows import *
from config_reader import *

(input_path,template_path,output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)
# File to be copied
workbook1 = fnmatch.filter(os.listdir(input_path), '*Intl Vol Rate - Asia*')
workbook2 = fnmatch.filter(os.listdir(input_path), '*Intl Vol Rate - YTD*')
workbook3 = fnmatch.filter(os.listdir(input_path), '*Intl Vol Rate - MTD*')
workbook4 = fnmatch.filter(os.listdir(input_path), '*Intl Vol Rate - QTD*')
print("File Names are " + workbook1[0] + ", " + workbook2[0] + ", " + workbook3[0] + ", " + workbook4[0])

workbook_base = 'Intl Vol Rate'
mtd_workbook_url = r'{}{}'.format(input_path,workbook3[0])
ytd_workbook_url = r'{}{}'.format(input_path,workbook2[0])
qtd_workbook_url = r'{}{}'.format(input_path,workbook4[0])
asia_workbook_url = r'{}{}'.format(input_path,workbook1[0])
print("Workbook URLs are " + mtd_workbook_url + ", " + ytd_workbook_url + ", " + qtd_workbook_url + ", " + asia_workbook_url)

workbook_template = r'{}{}.xltx' .format(template_path,workbook_base)
print("Template URL is " + workbook_template)

result_workbook = r'{}{}_combined.xlsx'.format(output_path,workbook_base)
result_workbook_asia = r'{}{}_Asia_combined.xlsx'.format(output_path, workbook_base)
print("Result workbook URLs are " + result_workbook + ", " + result_workbook_asia)

mtd_wb = openpyxl.load_workbook(mtd_workbook_url)  # Add file name
ytd_wb = openpyxl.load_workbook(ytd_workbook_url)  # Add file name
qtd_wb = openpyxl.load_workbook(qtd_workbook_url)  # Add file name
asia_wb = openpyxl.load_workbook(asia_workbook_url)  # Add file name

mtd_regex = re.compile(r'^.*\-MTD')
qtd_regex = re.compile(r'^.*\-QTD')
ytd_regex = re.compile(r'^.*\-YTD')

target = openpyxl.load_workbook(workbook_template)
target.template = False

#target = copy(wb_template)
createMergedSheet(target['YTD'], ytd_regex, ytd_wb, 1, 10, 9, 8)
createMergedSheet(target['MTD'], mtd_regex, mtd_wb, 1, 10, 9, 8)
createMergedSheet(target['QTD'], qtd_regex, qtd_wb, 1, 10, 9, 8)
target.save(result_workbook)

target1 = openpyxl.load_workbook(workbook_template)
target1.template = False

createMergedSheet(target1['YTD'], ytd_regex, asia_wb, 1, 10, 9, 8)
createMergedSheet(target1['MTD'], mtd_regex, asia_wb, 1, 10, 9, 8)
createMergedSheet(target1['QTD'], qtd_regex, asia_wb, 1, 10, 9, 8)
target1.save(result_workbook_asia)
