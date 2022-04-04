"""
Duy trì 2 file:
    1. Xuất tại thời điểm ngày đang chạy (p1)
    2. Xuất tổng tiền từ ngày làm việc cuối cùng của tháng trước tới ngày đang chạy (p2)
"""

from automation.trading_service.thanhtoanbutru import *


# DONE
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    t0_date = info['end_date'].replace('/','-')
    year = t0_date[:4]
    month = t0_date[5:7]
    start_date = bdate(f"{year}-{month}-{'01'}",-1)
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        WITH [info] AS (
            SELECT 
                [relationship].[account_code],
                [relationship].[sub_account],
                [account].[customer_name],
                [account].[account_type],
                [branch].[branch_name],
                [broker].[broker_name]
            FROM
                [relationship]
            LEFT JOIN
                [account]
            ON
                [relationship].[account_code] = [account].[account_code]
            LEFT JOIN
                [broker]
            ON 
                [relationship].[broker_id] = [broker].[broker_id]
            LEFT JOIN
                [branch]
            ON
                [relationship].[branch_id] = [branch].[branch_id]
            WHERE
                [relationship].[date] = '{t0_date}'
        ),
        [interest_actu] AS (
            SELECT
                [rcf0011].[sub_account],
                SUM(ISNULL([rcf0011].[interest],0)) [interest_actu]
            FROM
                [rcf0011]
            WHERE
                [rcf0011].[date] BETWEEN '{start_date}' AND '{t0_date}'
            GROUP BY
                [rcf0011].[sub_account]
        ),
        [interest_paid] AS (
            SELECT
                [cash_balance].[sub_account],
                SUM(ISNULL([cash_balance].[increase],0)) [interest_paid]
            FROM
                [cash_balance]
            WHERE
                [cash_balance].[remark] LIKE N'%lãi%tiền gửi%'
            AND
                [cash_balance].[date] BETWEEN '{start_date}' AND '{t0_date}'
            GROUP BY
                [cash_balance].[sub_account]
        ),
        [balance] AS (
            SELECT
                [sub_account_deposit].[sub_account],
                SUM(ISNULL([sub_account_deposit].[closing_balance],0)) [closing_balance]
            FROM
                [sub_account_deposit]
            WHERE
                [sub_account_deposit].[date] BETWEEN '{start_date}' AND '{t0_date}'
            GROUP BY
                [sub_account_deposit].[sub_account]
        )
        
        SELECT 
            [final].*,
            ROUND([final].[interest_actu],2) - ROUND([final].[interest_calc],2) [diff]
            
        FROM (
            SELECT 
                [info].[account_code],
                [info].[account_type],
                [info].[sub_account],
                [info].[customer_name],
                [info].[branch_name],
                [info].[broker_name],
                ISNULL([interest_actu].[interest_actu],0) [interest_actu],
                ISNULL([interest_paid].[interest_paid],0) [interest_paid],
                CASE
                    WHEN ([info].[account_type] LIKE N'%tự doanh%') OR ([info].[account_type] LIKE N'%nước ngoài%')
                        THEN 0
                    WHEN [info].[account_type] LIKE N'%trong nước%'
                        THEN ISNULL([balance].[closing_balance],0)*0.1/100/360
                END [interest_calc]
            FROM [info]
            LEFT JOIN [interest_actu]
                ON [info].[sub_account] = [interest_actu].[sub_account]
            LEFT JOIN [interest_paid]
                ON [info].[sub_account] = [interest_paid].[sub_account]
            LEFT JOIN [balance]
                ON [info].[sub_account] = [balance].[sub_account]
            ) [final]
            
        WHERE
            [final].[interest_actu] <> 0 OR [final].[interest_calc] <> 0
            
        """,
            connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    fstart_date = dt.datetime.strptime(start_date,'%Y-%m-%d').strftime('%d.%m.%Y')
    ft0_date = dt.datetime.strptime(t0_date,'%Y-%m-%d').strftime('%d.%m.%Y')
    file_name = f'Báo cáo Đối chiếu lãi tiền gửi phát sinh trên TKKH {fstart_date} đến {ft0_date}.xlsx'
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
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Lãi tiền gửi cộng dồn trên TKKH',
        'Số tiền lãi đã trả',
        'Lãi tiền gửi cộng dồn tính lại',
        'Bất thường',
        'Chi nhánh quản lý',
        'Môi giới quản lý',
    ]

    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU LÃI TIỀN GỬI PHÁT SINH TRÊN TÀI KHOẢN KHÁCH HÀNG'
    sub_title_name = f'Từ ngày {fstart_date} đến ngày {ft0_date}'
    sum_name = 'Cộng phát sinh'

    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.66,'y_scale':0.71})

    worksheet.set_column('A:A',6)
    worksheet.set_column('B:C',13)
    worksheet.set_column('D:D',28)
    worksheet.set_column('E:H',20)
    worksheet.set_column('I:I',20)
    worksheet.set_column('J:J',28)
    worksheet.set_row(6,18)
    worksheet.set_row(9,30)

    sum_start_row = table.shape[0]+11
    worksheet.merge_range('C1:J1',CompanyName,company_name_format)
    worksheet.merge_range('C2:J2',CompanyAddress,company_info_format)
    worksheet.merge_range('C3:J3',CompanyPhoneNumber,company_info_format)
    worksheet.merge_range('A7:J7',sheet_title_name,sheet_title_format)
    worksheet.merge_range('A8:J8',sub_title_name,sub_title_date_format)
    text = f'Ngày    tháng    năm    '
    worksheet.merge_range(f'H{sum_start_row+2}:J{sum_start_row+2}',text,footer_dmy_format)
    worksheet.merge_range(f'A{sum_start_row+3}:C{sum_start_row+3}','Người lập',footer_text_format)
    worksheet.merge_range(f'H{sum_start_row+3}:J{sum_start_row+3}','Người duyệt',footer_text_format)
    worksheet.write_row('A4',['']*len(headers),empty_row_format)
    worksheet.write_row('A10',headers,headers_format)
    worksheet.write_column('A11',np.arange(table.shape[0])+1,text_center_format)
    worksheet.write_column('B11',table['account_code'],text_center_format)
    worksheet.write_column('C11',table['sub_account'],text_center_format)
    worksheet.write_column('D11',table['customer_name'].str.title(),text_left_format)
    worksheet.write_column('E11',table['interest_actu'],money_format)
    worksheet.write_column('F11',table['interest_paid'],money_format)
    worksheet.write_column('G11',table['interest_calc'],money_format)
    worksheet.write_column('H11',table['diff'],money_format)
    worksheet.write_column('I11',table['branch_name'],text_left_format)
    worksheet.write_column('J11',table['broker_name'].str.title(),text_left_format)
    worksheet.merge_range(f'A{sum_start_row}:D{sum_start_row}',sum_name,headers_format)
    for col in 'EFGH':
        worksheet.write(f'{col}{sum_start_row}',f'=SUBTOTAL(9,{col}11:{col}{sum_start_row-1})',sum_money_format)
    worksheet.merge_range(f'I{sum_start_row}:J{sum_start_row}','',text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
