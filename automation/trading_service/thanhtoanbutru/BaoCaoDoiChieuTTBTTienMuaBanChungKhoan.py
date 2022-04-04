"""
I. RCF1002 -> cash_balance
   ROD0040 -> trading_record
II. Rule:
    1. Phần Kết quả khớp lệnh (ROD0040):
        - Nếu loại lệnh = MUA -> điều kiện ngày là T0
        - Nếu loại lệnh = BÁN -> điều kiện ngày là T-2
    2. Phần Theo giao dịch tiền (RCF1002)
        - Điều kiện ngày là T0
    3. Cột Ngày giao dịch
        - Nếu Loại lệnh = MUA -> ngày giao dịch = T0
        - Nếu loại lệnh = BÁN -> ngày giao dịch T-2
    4. Cột Ngày thanh toán = T0
    5. Cột thuế
        - Thuế trong 'Kết quả khớp lệnh' (ROD0040) = tax_of_selling + tax_of_share_dividend
        - Thuế trong 'Theo giao dịch tiền' (RCF1002) lấy giá trị của cột decrease với transaction_id = 0066 (lấy toàn bộ)
"""
from automation.trading_service.thanhtoanbutru import *


# DONE
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    t0_date = info['end_date'].replace('/','-')
    t1_date = bdate(t0_date,-1)
    t2_date = bdate(t0_date,-2)
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    # query ROD0040
    order_table = pd.read_sql(
        f"""
        (
        SELECT
            '{t0_date}' [date],
            [trading_record].[sub_account],
            CASE
                WHEN [trading_record].[type_of_order] = 'S' THEN N'Bán'
            END [type_of_order],
            SUM([trading_record].[value]) [value],
            SUM([trading_record].[fee]) [fee],
            SUM([trading_record].[tax_of_selling] + [trading_record].[tax_of_share_dividend]) [tax]
        FROM [trading_record]
        WHERE (([trading_record].[date] = '{t2_date}' AND [trading_record].[settlement_period] = 2)
                OR ([trading_record].[date] = '{t1_date}' AND [trading_record].[settlement_period] = 1))
            AND [trading_record].[type_of_order] = 'S'
        GROUP BY
            [trading_record].[date],
            [trading_record].[sub_account],
            [trading_record].[type_of_order]
        )
        UNION ALL
        (
        SELECT
            [trading_record].[date],
            [trading_record].[sub_account],
            CASE
                WHEN [trading_record].[type_of_order] = 'B' THEN N'Mua'
            END [type_of_order],
            SUM([trading_record].[value]) [value],
            SUM([trading_record].[fee]) [fee],
            SUM([trading_record].[tax_of_selling] + [trading_record].[tax_of_share_dividend]) [tax]
        FROM 
            [trading_record]
        WHERE 
            [trading_record].[date] = '{t0_date}' AND [trading_record].[type_of_order] = 'B'
        GROUP BY
            [trading_record].[date],
            [trading_record].[sub_account],
            [trading_record].[type_of_order]
        )
        ORDER BY
            [date],
            [sub_account],
            [type_of_order]
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # query RCF1002
    cash_table = pd.read_sql(
        f"""
        WITH 
        [cf_value] AS (
            SELECT 
                [cash_balance].[date],
                [cash_balance].[sub_account],
                N'Mua' AS [type_of_order],
                [cash_balance].[decrease] AS [value]
            FROM 
                [cash_balance]
            WHERE
                [cash_balance].[transaction_id] = '8865' AND [cash_balance].[date] = '{t0_date}'
            UNION ALL
            SELECT
                [cash_balance].[date],
                [cash_balance].[sub_account],
                N'Bán' AS [type_of_order],
                [cash_balance].[increase] AS [value]
            FROM 
                [cash_balance]
            WHERE
                [cash_balance].[transaction_id] = '8866' AND [cash_balance].[date] = '{t0_date}'
            ),
        [cf_fee] AS (
            SELECT 
                [cash_balance].[date],
                [cash_balance].[sub_account],
                N'Mua' AS [type_of_order],
                [cash_balance].[decrease] AS [fee]
            FROM 
                [cash_balance]
            WHERE
                [cash_balance].[transaction_id] = '8855' AND [cash_balance].[date] = '{t0_date}'
            UNION ALL
            SELECT 
                [cash_balance].[date],
                [cash_balance].[sub_account],
                N'Bán' AS [type_of_order],
                [cash_balance].[decrease] AS [fee]
            FROM 
                [cash_balance]
            WHERE
                [cash_balance].[transaction_id] = '8856' AND [cash_balance].[date] = '{t0_date}'
            ),
        [cf_tax] AS ( 
            SELECT 
                [cash_balance].[date],
                [cash_balance].[sub_account],
                N'Bán' AS [type_of_order],
                [cash_balance].[decrease] AS [tax]
            FROM 
                [cash_balance]
            WHERE
                [cash_balance].[transaction_id] = '0066' AND [cash_balance].[date] = '{t0_date}'
            )
        SELECT
            [a].[date],
            [a].[sub_account],
            [a].[type_of_order],
            [v].[value],
            [f].[fee],
            [t].[tax]
        FROM
            (
            SELECT DISTINCT [date],[sub_account],[type_of_order] FROM [cf_value]
            UNION
            SELECT DISTINCT [date],[sub_account],[type_of_order] FROM [cf_fee]
            UNION
            SELECT DISTINCT [date],[sub_account],[type_of_order] FROM [cf_tax]
            ) [a]
        LEFT JOIN (SELECT [cf_value].[date],[cf_value].[sub_account],[cf_value].[type_of_order],SUM([cf_value].[value]) [value] FROM [cf_value] GROUP BY [cf_value].[date],[cf_value].[sub_account],[cf_value].[type_of_order]) [v] 
            ON [a].[date] = [v].[date] AND [a].[sub_account] = [v].[sub_account] AND [a].[type_of_order] = [v].[type_of_order]
        LEFT JOIN (SELECT [cf_fee].[date],[cf_fee].[sub_account],[cf_fee].[type_of_order],SUM([cf_fee].[fee]) [fee] FROM [cf_fee] GROUP BY [cf_fee].[date],[cf_fee].[sub_account],[cf_fee].[type_of_order]) [f] 
            ON [a].[date] = [f].[date] AND [a].[sub_account] = [f].[sub_account] AND [a].[type_of_order] = [f].[type_of_order]
        LEFT JOIN (SELECT [cf_tax].[date],[cf_tax].[sub_account],[cf_tax].[type_of_order],SUM([cf_tax].[tax]) [tax] FROM [cf_tax] GROUP BY [cf_tax].[date],[cf_tax].[sub_account],[cf_tax].[type_of_order]) [t] 
            ON [a].[date] = [t].[date] AND [a].[sub_account] = [t].[sub_account] AND [a].[type_of_order] = [t].[type_of_order]
        
        ORDER BY [date], [sub_account], [type_of_order]
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    ).fillna(0)
    customer_info = pd.read_sql(
        f"""
        SELECT 
            [sub_account].[sub_account],
            [sub_account].[account_code],
            [account].[customer_name]
        FROM [sub_account]
        LEFT JOIN 
            [account]
        ON 
            [account].[account_code] = sub_account.account_code
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # các tài khoản mở tại công ty nhưng lưu ký ở nơi khác PHS sẽ ko theo dõi dòng tiền -> loại ra
    order_table = order_table.loc[order_table.index.isin(cash_table.index)]
    # join
    table = customer_info.join(
        order_table,on='sub_account',
        how='right',
    ).join(
        cash_table.reset_index().set_index(['sub_account','date','type_of_order']),
        on=['sub_account','date','type_of_order'],
        how='outer',
        rsuffix='_cash',
    )
    table.rename({'date':'pay_date','value':'value_order','fee':'fee_order','tax':'tax_order'},axis=1,inplace=True)
    for diff_col,order_col,cash_col in zip(
            ['diff_value','diff_fee','diff_tax'],
            ['value_order','fee_order','tax_order'],
            ['value_cash','fee_cash','tax_cash'],
    ):
        table[diff_col] = table[order_col]-table[cash_col]

    table.insert(3,'trade_date',table['pay_date'])
    table.loc[table['type_of_order']=='Bán','trade_date'] = t2_date

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU THANH TOÁN BÙ TRỪ TIỀN MUA BÁN CHỨNG KHOÁN
    report_date = dt.datetime.strptime(t0_date,'%Y-%m-%d').strftime('%d.%m.%Y')
    file_name = f'Báo cáo Đối chiếu TTBT tiền mua bán chứng khoán {report_date}.xlsx'
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
    sheet_subtitle_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    from_to_format = workbook.add_format(
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
    text_left_format = workbook.add_format(  # for customer name only
        {
            'border':1,
            'align':'left',
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
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    sum_money_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'dd/mm/yyyy'
        }
    )
    footer_dmy_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
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
    headers = [
        'STT',
        'Ngày giao dịch',
        'Ngày thanh toán tiền',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Loại lệnh',
        'Theo kết quả khớp lệnh',
        'Theo giao dịch tiền',
        'Lệch',
    ]
    sub_headers = [
        'Giá trị khớp',
        'Phí',
        'Thuế',
    ]
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU THANH TOÁN BÙ TRỪ TIỀN MUA BÁN CHỨNG KHOÁN TẠI PHS'
    sheet_subtitle_name = '(KHÔNG BAO GỒM CÁC TK LƯU KÝ NƠI KHÁC)'
    sub_title_date = dt.datetime.strptime(t0_date,"%Y-%m-%d").strftime("%d/%m/%Y")
    sub_title_name = f'Từ ngày {sub_title_date} đến {sub_title_date}'

    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)

    worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.66,'y_scale':0.71})
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:E',13)
    worksheet.set_column('F:F',28)
    worksheet.set_column('G:G',9)
    worksheet.set_column('H:P',13)

    worksheet.merge_range('C1:I1',CompanyName.upper(),company_name_format)
    worksheet.merge_range('C2:I2',CompanyAddress,company_info_format)
    worksheet.merge_range('C3:I3',CompanyPhoneNumber,company_info_format)
    worksheet.merge_range('A6:P6',sheet_title_name,sheet_title_format)
    worksheet.merge_range('A7:P7',sheet_subtitle_name,sheet_subtitle_format)
    worksheet.merge_range('A8:P8',sub_title_name,from_to_format)
    for col,header in zip(range(len(headers)-3),headers[:-3]):
        worksheet.merge_range(9,col,10,col,header,headers_format)
    worksheet.merge_range('H10:J10',headers[-3],headers_format)
    worksheet.merge_range('K10:M10',headers[-2],headers_format)
    worksheet.merge_range('N10:P10',headers[-1],headers_format)
    sum_start_row = table.shape[0]+12
    worksheet.merge_range(
        f'A{sum_start_row}:G{sum_start_row}',
        'Tổng',
        headers_format
    )
    footer_start_row = sum_start_row+2
    worksheet.merge_range(
        f'N{footer_start_row}:P{footer_start_row}',
        f'Ngày     tháng     năm        ',
        footer_dmy_format
    )
    worksheet.merge_range(
        f'N{footer_start_row+1}:P{footer_start_row+1}',
        'Người duyệt',
        footer_text_format
    )
    worksheet.merge_range(
        f'C{footer_start_row+1}:E{footer_start_row+1}',
        'Người lập',
        footer_text_format
    )
    # write row & column
    worksheet.write_row('H11',sub_headers*3,headers_format)
    worksheet.write_row('A4',['']*(len(headers)+len(sub_headers)+3),empty_row_format)
    worksheet.write_column('A12',np.arange(table.shape[0])+1,text_center_format)
    worksheet.write_column('B12',table['trade_date'],date_format)
    worksheet.write_column('C12',table['pay_date'],date_format)
    worksheet.write_column('D12',table['account_code'],text_center_format)
    worksheet.write_column('E12',table.index,text_center_format)
    worksheet.write_column('F12',table['customer_name'].str.title(),text_left_format)
    worksheet.write_column('G12',table['type_of_order'],text_center_format)
    worksheet.write_column('H12',table['value_order'],money_format)
    worksheet.write_column('I12',table['fee_order'],money_format)
    worksheet.write_column('J12',table['tax_order'],money_format)
    worksheet.write_column('K12',table['value_cash'],money_format)
    worksheet.write_column('L12',table['fee_cash'],money_format)
    worksheet.write_column('M12',table['tax_cash'],money_format)
    worksheet.write_column('N12',table['diff_value'],money_format)
    worksheet.write_column('O12',table['diff_fee'],money_format)
    worksheet.write_column('P12',table['diff_tax'],money_format)
    worksheet.write(f'D{footer_start_row+1}','Người lập',footer_text_format)
    for col in 'HIJKLMNOP':
        worksheet.write(f'{col}{sum_start_row}',f'=SUBTOTAL(9,{col}12:{col}{sum_start_row-1})',sum_money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
