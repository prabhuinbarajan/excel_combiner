import openpyxl
from openpyxl.formula.translate import Translator

# Prepare the spreadsheets to copy from and paste too.

# File to be copied
import pandas as pd
import re

from openpyxl.utils.dataframe import dataframe_to_rows


def pasteRange(ws, rows, startRow,totalsFrom, title=''):
    range_text = ''
    lastRow = 1
    lastCol = 1
    first=True
    lr = None
    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row, 1):
            lr = row
            ws.cell(row=r_idx+startRow, column=c_idx, value=value)
            lastRow = r_idx+startRow
            lastCol = c_idx
            if first:
                range_text=ws.cell(row=lastRow, column=lastCol).coordinate
                first=False

    if (totalsFrom >= 0):

        for c_idx, value in enumerate(lr, 1):
            if c_idx == 1 :
                ws.cell(row=lastRow + 1, column=c_idx, value=title)
            if c_idx > totalsFrom:
                column_letter =  ws.cell(row=startRow, column=c_idx).column_letter
                range_text = column_letter +str(startRow+1) + ":" + column_letter+str(lastRow)
                ws.cell(row=lastRow+1, column=c_idx,  value='=SUM('+range_text + ')')
    return lastRow+1, lr


p2020_data_file = 'report_samples/Actual-P1-P5-2020.csv'
p2019_data_file = 'report_samples/Actual-P1-P13-2019.csv'
interested_period = 'P4 2020'
num_periods = 13
account_meta = 'config/account_meta.csv'
workbook = 'A510-MOEND_606-2'
workbook_url = 'report_samples/{}.xlsx'.format(workbook)
workbook_template = 'templates/{}.xltx' .format(workbook)
result_workbook = 'results/{}_combined.xlsx'.format(workbook)
period_columns_regex = re.compile(r'^P[0-9]*')

account_meta = pd.read_csv(account_meta, dtype=str)

df_2019 = pd.read_csv(p2019_data_file)
df_2019= df_2019.loc[:, ~df_2019.columns.str.contains('^Unnamed')]
df_2019.drop(['Department', 'Company', 'Budget Entity'], axis=1, inplace=True)
df_2020 = pd.read_csv(p2020_data_file)
df_2020.drop(['Department', 'Company', 'Budget Entity'], axis=1, inplace=True)
df_2020= df_2020.loc[:, ~df_2020.columns.str.contains('^Unnamed')]

df_2019['AffiliateRollup'] = df_2019['Affiliate'].str.slice(0, 4, 1)
df_2020['AffiliateRollup'] = df_2020['Affiliate'].str.slice(0, 4, 1)

#result = pd.merge(df_2019, df_2020, how='outer', on=['Account', 'AffiliateRollup', 'Affiliate'])
result = pd.concat([df_2019, df_2020])
result.drop(['Affiliate'], axis=1, inplace=True)
result.rename(columns = {'AffiliateRollup':'Affiliate'}, inplace = True)

columns = list(filter(lambda i: period_columns_regex.match(i), result.columns))

index_of_period= columns.index(interested_period)
columns = columns[index_of_period-num_periods:index_of_period+1]
aggregate_columns = {columns[i]: 'sum' for i in range(0, len(columns))}


result = pd.merge(result, account_meta, how='left', on=['Account'])

agg_result = result.groupby(['AC', 'Account', 'Affiliate']).agg(aggregate_columns).reset_index()
sub_totals = result.groupby(['AC']).agg(aggregate_columns).reset_index()



check_figure = sub_totals.agg(['sum'])

combined_result = pd.merge(agg_result, account_meta, how='outer', on=['AC', 'Account'])
enriched_trend_cols = ['AC', 'Account Category', 'Account', 'Acct Cat', 'Affiliate', 'Account Description']
enriched_trend_cols.extend(columns)

for col in ['Account Category', 'Account', 'Acct Cat', 'Affiliate', 'Account Description']:
    check_figure[col] = ''
    sub_totals[col] = ''

final_result = combined_result.reindex(enriched_trend_cols,axis=1)
sub_totals = sub_totals.reindex(enriched_trend_cols,axis=1)
check_figure = check_figure.reindex(enriched_trend_cols,axis=1)
check_figure = check_figure.drop(['AC'], axis=1)


target = openpyxl.load_workbook(workbook_template)
target.template = False
sheet = target['Sheet']

column_header = pd.DataFrame(columns=final_result.columns)
column_header.drop(['AC'], axis = 1, inplace=True)
header_rows = dataframe_to_rows(column_header, index=False, header=True)
pasteRange(sheet, header_rows, 8,-1)

startRow = 9
rows = None
totalRows = {}
lastRowRecord = None
for asset_category in ['Assets', 'Liabilities','Equities']:
    ac_results = final_result[final_result['AC'] == asset_category]
    ac_results = ac_results.drop(['AC'], axis=1)
    rows = dataframe_to_rows(ac_results, index=False, header=False)
    lr , lastRowRecord = pasteRange(sheet, rows, startRow, 5, 'Total '+ asset_category)
    totalRows[asset_category] = lr
    startRow += ac_results.shape[0]+1


for c_idx, value in enumerate(lastRowRecord,1):
    if c_idx == 1 :
        sheet.cell(row=startRow + 2, column=c_idx, value='check figure')

    if c_idx > 5:
        column_letter = sheet.cell(row=startRow+2, column=c_idx).column_letter
        range_text = column_letter + str(totalRows['Assets']) + \
                     '+' + column_letter + str(totalRows['Liabilities']) + \
                     '+' + column_letter + str(totalRows['Equities'])
        sheet.cell(row=startRow+2, column=c_idx,  value='=' + range_text )

target.save(result_workbook)
print(agg_result)


