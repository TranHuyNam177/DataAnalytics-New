from blackscholes import *


def run():
    folder = join(dirname(realpath(__file__)),'backtest_result')
    file_names = [name for name in listdir(folder) if
                  name.endswith('.xlsx') and name.startswith('C') and '_edited' not in name]

    for file in file_names:

        df = pd.read_excel(join(folder,file),index_col=0)
        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
        writer = pd.ExcelWriter(join(folder,file),engine='xlsxwriter')
        df.to_excel(writer,sheet_name='Sheet1',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Themes
        title = workbook.add_format({
            'bold':True,
            'align':'center',
            'fg_color':'#2F4F4F',
            'font_color':'white',
            'border':1
        })
        text_left = workbook.add_format({
            'align':'left',
            'valign':'top',
            'border':1,
        })
        number = workbook.add_format({
            'align':'right',
            'valign':'top',
            'border':1,
            'num_format':'0'
        })
        number_decimals = workbook.add_format({
            'align':'right',
            'valign':'top',
            'border':1,
            'num_format':'0.000000'
        })
        text_right = workbook.add_format({
            'align':'right',
            'valign':'top',
            'border':1,
        })

        # Format
        for col in range(0,df.shape[1]):
            worksheet.write(0,col,df.columns[col],title)

        col_format = [text_right,number,text_left,text_right]+[number]*8+[number_decimals]*3+[number]
        for row in range(0,df.shape[0]):
            for col in range(0,df.shape[1]):
                data = df.iloc[row,col]
                try:
                    worksheet.write(row+1,col,data,col_format[col])
                except (Exception,):
                    worksheet.write(row+1,col,'NaN',col_format[col])

        # Column Fit
        worksheet.set_column(0,0,15)
        worksheet.set_column(1,1,21)
        worksheet.set_column(2,2,21)
        worksheet.set_column(3,3,32)
        worksheet.set_column(4,4,17)
        worksheet.set_column(5,5,22)
        worksheet.set_column(6,6,27)
        worksheet.set_column(7,7,32)
        worksheet.set_column(8,8,19)
        worksheet.set_column(9,9,15)
        worksheet.set_column(10,10,20)
        worksheet.set_column(11,11,17)
        worksheet.set_column(12,12,18)
        worksheet.set_column(13,13,19)
        worksheet.set_column(14,14,21)
        worksheet.set_column(15,15,23)

        writer.close()
