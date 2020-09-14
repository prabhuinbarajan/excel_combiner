from openpyxl import Workbook
from config_reader import *
from win32com.client import Dispatch

# Prepare the spreadsheets to copy from and paste too.
(TB_input_path, PL_input_path, template_path, TB_output_path, PL_output_path, myyear, myper) = \
    get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)

# File to be copied

workbook_t19 = fnmatch.filter(os.listdir(PL_input_path), '*PL14_Executive Summary Segment Report Collection LV2 Act vs Plan MTD*')
workbook_t18 = fnmatch.filter(os.listdir(PL_input_path), '*PL14_Executive Summary Segment Report Collection LV2 Act vs Plan QTD*')
workbook_t17 = fnmatch.filter(os.listdir(PL_input_path), '*PL14_Executive Summary Segment Report Collection LV2 Act vs Plan YTD*')
workbook_t20 = fnmatch.filter(os.listdir(PL_input_path), '*PL14_Executive Summary Segment Report Collection LV2 CY vs PY MTD*')
workbook_t21 = fnmatch.filter(os.listdir(PL_input_path), '*PL14_Executive Summary Segment Report Collection LV2 CY vs PY QTD*')
workbook_t22 = fnmatch.filter(os.listdir(PL_input_path), '*PL14_Executive Summary Segment Report Collection LV2 CY vs PY YTD*')

print("File Names are " + workbook_t19[0] + ", " + workbook_t18[0] + ", " + workbook_t17[0] + ", " + workbook_t20[0]
      + ", " + workbook_t21[0] + ", " + workbook_t22[0])

workbook_t19_url = r'{}{}'.format(PL_input_path,workbook_t19[0])
workbook_t18_url = r'{}{}'.format(PL_input_path,workbook_t18[0])
workbook_t17_url = r'{}{}'.format(PL_input_path,workbook_t17[0])
workbook_t20_url = r'{}{}'.format(PL_input_path,workbook_t20[0])
workbook_t21_url = r'{}{}'.format(PL_input_path,workbook_t21[0])
workbook_t22_url = r'{}{}'.format(PL_input_path,workbook_t22[0])

print("Workbook URLs are " + workbook_t19_url + ", " + workbook_t18_url + ", " + workbook_t17_url
      + ", " + workbook_t20_url + ", " + workbook_t21_url + ", " + workbook_t22_url)

result = workbook_t19[0]
result_workbook = r'{}{}_combined.xlsx'.format(PL_output_path,result[:64])
print("Result workbook URLs are " + result_workbook)

wb = Workbook()
wb.save(result_workbook)

xl = Dispatch("Excel.Application")
xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

wb1 = xl.Workbooks.Open(Filename=workbook_t19_url)
wb2 = xl.Workbooks.Open(Filename=workbook_t18_url)
wb3 = xl.Workbooks.Open(Filename=workbook_t17_url)
wb4 = xl.Workbooks.Open(Filename=workbook_t20_url)
wb5 = xl.Workbooks.Open(Filename=workbook_t21_url)
wb6 = xl.Workbooks.Open(Filename=workbook_t22_url)
wb7 = xl.Workbooks.Open(Filename=result_workbook)

ws1 = wb1.Worksheets('LV2 Act vs Plan-MTD')
ws1.Copy(Before=wb7.Worksheets(1))

ws2 = wb2.Worksheets('LV2 Act vs Plan-QTD')
ws2.Copy(Before=wb7.Worksheets(1))

ws3 = wb3.Worksheets('LV2 Act vs Plan-YTD')
ws3.Copy(Before=wb7.Worksheets(1))

ws4 = wb4.Worksheets('LV2 CY vs PY-MTD')
ws4.Copy(Before=wb7.Worksheets(1))

ws5 = wb5.Worksheets('LV2 CY vs PY-QTD')
ws5.Copy(Before=wb7.Worksheets(1))

ws6 = wb6.Worksheets('LV2 CY vs PY-YTD')
ws6.Copy(Before=wb7.Worksheets(1))

wb1.Close(SaveChanges=False)
wb2.Close(SaveChanges=False)
wb3.Close(SaveChanges=False)
wb4.Close(SaveChanges=False)
wb5.Close(SaveChanges=False)
wb6.Close(SaveChanges=False)
wb7.Close(SaveChanges=True)
xl.Quit()