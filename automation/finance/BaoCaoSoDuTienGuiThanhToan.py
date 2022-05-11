from automation.finance import *


def run(  # chạy hàng ngày
    run_time=None,
):

    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    t0_date = info['end_date'].replace('.','-')
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        SELECT 
            [t].[Date],
            [t].[Bank],
            [t].[AccountNumber],
            [t].[Balance],
            [t].[Currency]
        FROM [BankCurrentBalance] [t]
        WHERE [t].[Date] = '{t0_date}' AND [t].[Currency] = 'VND'
        ORDER BY [t].[Bank], [t].[AccountNumber]
        """,
        connect_DWH_CoSo
    )
    table['Balance'] = table['Balance'].fillna(0)

    ###################################################
    ###################################################
    ###################################################

    file_date = dt.datetime.strptime(t0_date,'%Y-%m-%d').strftime('%d.%m.%Y')
    file_name = f'Báo cáo số dư tiền gửi thanh toán {file_date}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    company_name_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    company_info_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    empty_row_format = workbook.add_format(
        {
            'bottom':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    sheet_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':14,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    sub_title_date_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    headers_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'#,##0_);(#,##0)'
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'num_format':'dd/mm/yyyy'
        }
    )

    headers = [
        'No.',
        'Date',
        'Bank',
        'Account Number',
        'Current Balance',
        'Currency',
    ]

    sheet_title_name = 'CURRENT ACCOUNT BALANCE'
    sub_title_name = f'Date {file_date}'

    worksheet = workbook.add_worksheet(f'CurrentAccountBalance')
    worksheet.hide_gridlines(option=2)
    worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.66,'y_scale':0.71})

    worksheet.set_column('A:A',6)
    worksheet.set_column('B:C',13)
    worksheet.set_column('D:E',20)
    worksheet.set_column('F:F',13)

    worksheet.merge_range('C1:F1',CompanyName,company_name_format)
    worksheet.merge_range('C2:F2',CompanyAddress.replace('Thành phố Hồ Chí Minh','TP.HCM'),company_info_format)
    worksheet.merge_range('C3:F3',CompanyPhoneNumber,company_info_format)
    worksheet.merge_range('A7:F7',sheet_title_name,sheet_title_format)
    worksheet.merge_range('A8:F8',sub_title_name,sub_title_date_format)
    worksheet.write_row('A4',['']*len(headers),empty_row_format)

    worksheet.write_row('A10',headers,headers_format)
    worksheet.write_column('A11',np.arange(table.shape[0])+1,text_center_format)
    worksheet.write_column('B11',table['Date'],date_format)
    worksheet.write_column('C11',table['Bank'],text_center_format)
    worksheet.write_column('D11',table['AccountNumber'],text_center_format)
    worksheet.write_column('E11',table['Balance'],money_format)
    worksheet.write_column('F11',table['Currency'],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
