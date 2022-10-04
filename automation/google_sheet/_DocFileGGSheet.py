import gspread

sa = gspread.service_account(filename='./.config/gspread/service_account.json')
sh = sa.open("dataTest")

wks = sh.worksheet("template_1")

print('Rows: ', wks.row_count)
print('Cols: ', wks.col_count)

print(wks.acell('G15').value)
print(wks.get('A10:I17'))