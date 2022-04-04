from automation.trading_service.thanhtoanbutru import *


def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    t0_date =info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        SELECT * FROM (
        SELECT
            [b].[branch_name],
            [t].[account_code],
            [a].[customer_name],
            ISNULL([c].[cash_balance_at_phs],0) [cash_at_phs],
            ISNULL([c].[cash_balance_at_vsd],0) [cash_at_vsd],
            ISNULL([t].[total_fee_tax],0) [total_fee_tax],
            ISNULL([d].[deferred_payment_amount_closing],0) + ISNULL([d].[deferred_payment_fee_closing],0) [deferred_amount],
            ISNULL([t].[nav],0) [nav]
        FROM
            [320200_tradingaccount] [t]
        LEFT JOIN
            [relationship] [r]
            ON [r].[account_code] = [t].[account_code] AND [r].[date] = [t].[date]
        LEFT JOIN
            [branch] [b]
            ON [b].[branch_id] = [r].[branch_id]
        LEFT JOIN
            [account] [a]
            ON [a].[account_code] = [r].[account_code]
        LEFT JOIN
            [rdt0121] [c]
            ON [c].[account_code] = [t].[account_code] AND [c].[date] = [t].[date]
        LEFT JOIN
            [rdt0141] [d]
            ON [d].[sub_account] = [r].[sub_account] AND [d].[date] = [t].[date]
        WHERE [t].[date] = '{t0_date}'
        ) [table]
        WHERE [cash_at_phs] <> 0
            OR [cash_at_vsd] <> 0
            OR [total_fee_tax] <> 0
            OR [deferred_amount] <> 0
            OR [nav] <> 0
        """,
        connect_DWH_PhaiSinh
    )

    ###################################################
    ###################################################
    ###################################################

    eod = dt.datetime.strptime(t0_date,"%Y/%m/%d").strftime("%d.%m.%Y")
    file_name = f'Báo cáo phái sinh theo dõi TKKH cần chuyển tiền từ VSD về TKGD {eod}.xlsx'
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
    company_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    sheet_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':16,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    from_to_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
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
            'font_size':12,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    num_bold_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    num_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'top',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'top',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    footer_dmy_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    footer_text_format = workbook.add_format(
        {
            'bold':True,
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)

    # Insert Phu Hung Logo
    worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.58,'y_scale':0.66})

    # Set Column Width and Row Height
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:B',20)
    worksheet.set_column('C:C',17)
    worksheet.set_column('D:D',26)
    worksheet.set_column('E:I',21)

    space = ' '*26
    worksheet.merge_range('B1:I1',f'{space}{CompanyName}',company_name_format)
    worksheet.merge_range('B2:I2',f'{space}{CompanyAddress}',company_format)
    worksheet.merge_range('B3:I3',f'{space}{CompanyPhoneNumber}',company_format)
    worksheet.merge_range('A6:I6','BÁO CÁO THEO DÕI CÁC TK CẦN XỬ LÝ TRÁNH ÂM TIỀN TRÊN FDS',sheet_title_format)
    title_date = f'{t0_date[-2:]}/{t0_date[5:7]}/{t0_date[:4]}'
    worksheet.merge_range('A7:I7',f'Ngày {title_date}',from_to_format)

    headers = [
        'STT',
        'Tên chi nhánh',
        'Tài khoản ký quỹ',
        'Tên khách hàng',
        'Số tiền tại công ty',
        'Số tiền ký quỹ tại VSD',
        'Nợ chậm trả',
        'Tổng giá trị phí thuế',
        'Giá trị tài sản ròng',
    ]
    worksheet.write_row('A9',headers,headers_format)
    worksheet.write_column('A10',np.arange(table.shape[0])+1,num_format)
    worksheet.write_column('B10',table['branch_name'],text_left_format)
    worksheet.write_column('C10',table['account_code'],text_center_format)
    worksheet.write_column('D10',table['customer_name'],text_left_format)
    worksheet.write_column('E10',table['cash_at_phs'],num_format)
    worksheet.write_column('F10',table['cash_at_vsd'],num_format)
    worksheet.write_column('G10',table['deferred_amount'],num_format)
    worksheet.write_column('H10',table['total_fee_tax'],num_format)
    worksheet.write_column('I10',table['nav'],num_format)

    sum_row = table.shape[0]+10
    worksheet.merge_range(f'A{sum_row}:D{sum_row}','Tổng',headers_format)

    for col in 'EFGHI':
        worksheet.write(f'{col}{sum_row}',f'=SUBTOTAL(9,{col}10:{col}{sum_row-1})',num_bold_format)

    worksheet.merge_range(f'A{sum_row+4}:D{sum_row+4}','Người lập',footer_text_format)
    worksheet.merge_range(
        f'H{sum_row+3}:I{sum_row+3}',
        f'Ngày {t0_date[-2:]} tháng {t0_date[5:7]} năm {t0_date[:4]}',
        footer_dmy_format
    )
    worksheet.merge_range(f'H{sum_row+4}:I{sum_row+4}','Người duyệt',footer_text_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
