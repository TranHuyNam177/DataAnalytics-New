"""
1. quarterly, ngày cuối của kỳ trước = 2021-09-30
2. bản chất của báo cáo này là xuất gốc từ RMR1062, bớt 1 số dòng va thêm 1 số dòng (Mrs. Tuyết)
3. Báo cáo RLN0006: lấy loại vay của trả trậm, margin, và bảo lãnh (bảo lãnh trừ tk tự doanh 022P... ) ra
4. Dựa theo file cách làm
    - cột H và cột O: bắt buộc phải khớp với báo cáo RLN0006
    - cột J: bắt buộc phải  khớp với báo cáo RCI0001
    - cột K: là số tiền còn có thể UTTB: cách lấy: Lấy ROD0040 (2 ngày làm việc cuối cùng
của tháng) cột giá trị bán trừ ( phí bán và thuế bán) trừ tiếp số tiền đã ứng (xuất
RCI0015, 2 ngày làm việc cuối cùng, chọn ngày bán là 2 ngày làm việc cuối cùng) là ra.
    - có thể LOẠI các tài khoản mà giá trị của cả 5 cột (H, J, K, L, O) đều bằng 0
    (chú ý: miễn sao các cột này có số và khớp với mấy báo cáo kia là dc,
    còn nếu họ có số ở các cột này thì không dc phép loại)
    - kết quả xuất ra có thể ít hơn hoặc bằng so với file kết quả của báo cáo
5. Không chạy lùi trước ngày 23/12/2021 được (vì chưa bắt đầu lưu VCF0051)
"""

from automation.trading_service.thanhtoanbutru import *


# DONE
def run(
    run_time=None
):  # BC quý, chạy vào ngày đầu quý sau

    start = time.time()
    info = get_info('quarterly',run_time)
    t0_date = bdate(bdate(info['end_date'],1),-1)
    t1_date = bdate(t0_date,-1)
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        WITH 
        [i] AS (
            SELECT 
                [branch].[branch_name],
                [relationship].[sub_account],
                [relationship].[account_code],
                [account].[customer_name]
            FROM [relationship]
            RIGHT JOIN
                [vcf0051]
            ON [relationship].[sub_account] = [vcf0051].[sub_account]
                AND [relationship].[date] = [vcf0051].[date]
            LEFT JOIN 
                [branch] 
            ON [relationship].[branch_id] = [branch].[branch_id]
            LEFT JOIN 
                [account] 
            ON [relationship].[account_code] = [account].[account_code]
            WHERE [relationship].[date] = '{t0_date}'
                AND [vcf0051].[status] IN ('A','B')
        ),
        [b] AS (
            SELECT
                [rmr1062].[sub_account],
                ISNULL([rmr1062].[credit_line],0) [credit_line],
                ISNULL([rmr1062].[total_outstanding],0) [total_outstanding],
                ISNULL([rmr1062].[total_cash],0) [total_cash],
                ISNULL([rmr1062].[total_margin_value],0) [total_margin],
                ISNULL([rmr1062].[total_asset_value],0) [total_asset],
                ISNULL([rmr1062].[total_outstanding_plus_interest],0) [total_outs_plus_int]
            FROM
                [rmr1062]
            WHERE
                [rmr1062].[margin_account] = 1 AND [rmr1062].[date] = '{t0_date}'
        ),
        [p] AS (
            SELECT
                [sub_account_report].[sub_account],
                [sub_account_report].[rtt],
                [sub_account_report].[buying_power]
            FROM
                [sub_account_report]
            WHERE
                [sub_account_report].[date] = '{t0_date}'
        ),
        [c] AS (
            SELECT
                [sub_account_deposit].[sub_account],
                [sub_account_deposit].[closing_balance] [cash]
            FROM
                [sub_account_deposit]
            WHERE
                [sub_account_deposit].[date] = '{t0_date}'
        ),
        [s] AS (
            SELECT
                [trading_record].[sub_account],
                (
                    SUM(ISNULL([trading_record].[value],0))
                    - SUM(ISNULL([trading_record].[fee],0))
                    - SUM(ISNULL([trading_record].[tax_of_selling],0)) 
                    - SUM(ISNULL([trading_record].[tax_of_share_dividend],0))
                ) [value]
            FROM
                [trading_record]
            WHERE
                [trading_record].[date] BETWEEN '{t1_date}' AND '{t0_date}'
                AND [trading_record].[type_of_order] = 'S'
            GROUP BY
                [trading_record].[sub_account]
        ),
        [r] AS (
            SELECT
                [payment_in_advance].[sub_account],
                SUM(ISNULL([payment_in_advance].[receivable],0)) [receivable]
            FROM
                [payment_in_advance]
            WHERE
                [payment_in_advance].[date] BETWEEN '{t1_date}' AND '{t0_date}'
                AND [payment_in_advance].[trading_date] BETWEEN '{t1_date}' AND '{t0_date}'
            GROUP BY [payment_in_advance].[sub_account]
        )
        SELECT [all].* FROM (
            SELECT 
                [i].*,
                [b].[credit_line],
                ISNULL([p].[rtt],0) [rtt],
                ISNULL([c].[cash],0) [cash_rci0001],
                ISNULL([s].[value],0) - ISNULL([r].[receivable],0) [remain_pia],
                ISNULL([p].[buying_power],0) [buying_power],
                [b].[total_outstanding],
                [b].[total_cash],
                [b].[total_margin],
                [b].[total_asset],
                CASE WHEN [b].[total_outs_plus_int] >= 0 THEN [b].[total_outs_plus_int]
                 ELSE 0
                 END [total_outs_plus_int], -- data của FLEX đang sai ở 022C062345 (-9đ), phải lên 0đ
                50 [rai]
            FROM [i]
            INNER JOIN [b] ON [b].[sub_account] = [i].[sub_account]
            LEFT JOIN [p] ON [p].[sub_account] = [i].[sub_account]
            LEFT JOIN [c] ON [c].[sub_account] = [i].[sub_account]
            LEFT JOIN [r] ON [r].[sub_account] = [i].[sub_account]
            LEFT JOIN [s] ON [s].[sub_account] = [i].[sub_account]
        ) [all]
        WHERE ([all].[rtt] <> 0 
            OR [all].[remain_pia] <> 0
            OR [all].[buying_power] <> 0
            OR [all].[total_outstanding] <> 0
            OR [all].[total_asset] <> 0)
            
    """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    file_date = dt.datetime.strptime(t0_date,"%Y-%m-%d").strftime("%d.%m.%Y")
    file_name = f'Dữ liệu gửi kiểm toán {file_date}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    info_company_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vcenter',
            'bold':True,
            'font_size':8,
            'font_name':'Times New Roman',
        }
    )
    title_format = workbook.add_format(
        {
            'top':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':18,
            'font_name':'Times New Roman',
        }
    )
    period_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'top',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    headers_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True,
            'bg_color':'#00b050',
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
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
    num_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0)'
        }
    )
    sum_name_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    sum_num_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0)'
        }
    )
    headers = [
        'No.',
        'Branch',
        'Account',
        'Tiểu khoản MR',
        'Name',
        'Creditline',
        'Margin Ratio',
        'Total Cash\nRCI0001',
        'UTTB còn lại',
        'Buying power',
        'Total Outstanding',
        'Total Cash',
        'Total Margin Value',
        'Total Asset Value',
        'Total Outstanding plus Total Interest',
        'Rai - Mortgage Ratio of Account',
    ]

    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:B',16)
    worksheet.set_column('C:D',13)
    worksheet.set_column('E:E',24)
    worksheet.set_column('F:P',18)
    worksheet.set_row(4,22)

    worksheet.merge_range('A1:F1',CompanyName,info_company_format)
    worksheet.merge_range('A2:F2',CompanyAddress,info_company_format)
    worksheet.merge_range('A3:F3',CompanyPhoneNumber,info_company_format)
    worksheet.merge_range('A4:P4','ACCOUNT DETAILS REPORT',title_format)
    worksheet.merge_range('A5:P5',f'Kỳ báo cáo: {period}',period_format)
    worksheet.write_row('A6',headers,headers_format)
    worksheet.write_column('A7',np.arange(table.shape[0])+1,text_center_format)
    worksheet.write_column('B7',table['branch_name'],text_left_format)
    worksheet.write_column('C7',table['account_code'],text_center_format)
    worksheet.write_column('D7',table['sub_account'],text_center_format)
    worksheet.write_column('E7',table['customer_name'],text_left_format)
    worksheet.write_column('F7',table['credit_line'],num_format)
    worksheet.write_column('G7',table['rtt'],num_format)
    worksheet.write_column('H7',table['cash_rci0001'],num_format)
    worksheet.write_column('I7',table['remain_pia'],num_format)
    worksheet.write_column('J7',table['buying_power'],num_format)
    worksheet.write_column('K7',table['total_outstanding'],num_format)
    worksheet.write_column('L7',table['total_cash'],num_format)
    worksheet.write_column('M7',table['total_margin'],num_format)
    worksheet.write_column('N7',table['total_asset'],num_format)
    worksheet.write_column('O7',table['total_outs_plus_int'],num_format)
    worksheet.write_column('P7',table['rai'],num_format)

    sum_row = table.shape[0]+7
    worksheet.merge_range(f'A{sum_row}:E{sum_row}','TỔNG',sum_name_format)
    for col in 'FGHIJKLMNOP':
        worksheet.write(f'{col}{sum_row}',f'=SUBTOTAL(9,{col}7:{col}{sum_row-1})',sum_num_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
