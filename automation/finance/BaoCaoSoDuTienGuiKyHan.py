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
            [t].[TermDays],
            [t].[TermMonths],
            [t].[InterestRate],
            [t].[IssueDate],
            [t].[ExpireDate],
            [t].[Balance],
            [t].[InterestAmount],
            [t].[Currency]
        FROM [BankDepositBalance] [t]
        WHERE [t].[Date] = '{t0_date}'
        ORDER BY [t].[Bank], [t].[AccountNumber]
        """,
        connect_DWH_CoSo,
    )

    ###################################################
    ###################################################
    ###################################################

    file_date = dt.datetime.strptime(t0_date,'%Y-%m-%d').strftime('%d.%m.%Y')
    file_name = f'Báo cáo số dư tiền gửi có kỳ hạn {file_date}.xlsx'
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
    integer_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0'
        }
    )
    pct_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0.000%'
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
        'No.', # A
        'Date', # B
        'Bank', # C
        'Account Number', # D
        'Term Days', # E
        'Term Months', # F
        'Interest Rate', # G
        'Issue Date', # H
        'Expire Date', # I
        'Balance', # J
        'Interest Amount', # K
        'Currency', # L
    ]

    sheet_title_name = 'DEPOSIT ACCOUNT BALANCE'
    sub_title_name = f'Date {file_date}'

    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.66,'y_scale':0.71})

    worksheet.set_column('A:A',6)
    worksheet.set_column('B:C',13)
    worksheet.set_column('D:D',20)
    worksheet.set_column('E:G',10)
    worksheet.set_column('H:I',13)
    worksheet.set_column('J:K',20)
    worksheet.set_column('L:L',10)

    worksheet.merge_range('C1:L1',CompanyName,company_name_format)
    worksheet.merge_range('C2:L2',CompanyAddress,company_info_format)
    worksheet.merge_range('C3:L3',CompanyPhoneNumber,company_info_format)
    worksheet.merge_range('A7:L7',sheet_title_name,sheet_title_format)
    worksheet.merge_range('A8:L8',sub_title_name,sub_title_date_format)
    worksheet.write_row('A4',['']*len(headers),empty_row_format)

    worksheet.write_row('A10',headers,headers_format)
    worksheet.write_column('A11',np.arange(table.shape[0])+1,text_center_format)
    worksheet.write_column('B11',table['Date'],date_format)
    worksheet.write_column('C11',table['Bank'],text_center_format)
    worksheet.write_column('D11',table['AccountNumber'],text_center_format)
    worksheet.write_column('E11',table['TermDays'],integer_format)
    worksheet.write_column('F11',table['TermMonths'],integer_format)
    worksheet.write_column('G11',table['InterestRate'],pct_format)
    worksheet.write_column('H11',table['IssueDate'],date_format)
    worksheet.write_column('I11',table['ExpireDate'],date_format)
    worksheet.write_column('J11',table['Balance'],money_format)
    worksheet.write_column('K11',table['InterestAmount'],money_format)
    worksheet.write_column('L11',table['Currency'],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
