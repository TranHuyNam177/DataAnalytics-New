from automation.trading_service.giaodichluuky import *


def run(
    run_time=None
):
    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    # Phi chuyen khoan thanh toan bu tru xuat theo ngay duoc dieu chinh
    def adjust_time(x):
        x_dt = dt.datetime.strptime(x,'%Y/%m/%d')
        if x_dt.weekday() in holidays.WEEKEND or x_dt in holidays.VN():
            result = bdate(x,-3)
        else:
            result = bdate(x,-2)
        return result

    """
    Chỗ này rule lấy ngày ở GDLK không thực sự rõ ràng, có khả năng sai.
    Tháng nào lấy ngày sai thì hardcode ngày đúng tại dòng ngay dưới
    """

    adj_start_date, adj_end_date = adjust_time(start_date), adjust_time(end_date)  # thay đổi tại đây

    transfer_fee = pd.read_sql(
        f"""
        WITH [r] AS (
            SELECT DISTINCT 
                [relationship].[date],
                [relationship].[account_code],
                [relationship].[branch_id]
            FROM [relationship]
            WHERE [relationship].[date] BETWEEN '{adj_start_date}' AND '{end_date}'
        ),
        [cal_sum_volume] AS (
            SELECT
                [date],
                [sub_account],
                [volume],
                SUM([volume]) OVER(PARTITION BY [date], [ticker]) [sum_volume]
            FROM [trading_record]
            WHERE [date] BETWEEN '{adj_start_date}' AND '{adj_end_date}'
            AND [type_of_order] = 'S' 
            AND [depository_place] = 'Tai PHS'
        ),
        [cal_vol_percent] AS (
            SELECT
                [cal_sum_volume].*,
                ([cal_sum_volume].[volume] / [cal_sum_volume].[sum_volume]) [vol_percent]
            FROM [cal_sum_volume]
        ),
        [rod0040] AS (
            SELECT
                [cal_vol_percent].*,
                CASE
                    WHEN [cal_vol_percent].[sum_volume] > 1000000 THEN [cal_vol_percent].[vol_percent] * 1000000
                    ELSE [cal_vol_percent].[volume]
                END [charged_volume]
            FROM [cal_vol_percent]
        ),
        [rse2009] AS (
            SELECT
                [r].[branch_id],
                [deposit_withdraw_stock].[date], 
                [deposit_withdraw_stock].[account_code],
                [creator],
                CASE
                    WHEN [deposit_withdraw_stock].[volume] >= 1000000 THEN 1000000
                    ELSE [deposit_withdraw_stock].[volume]
                END [volume]
                FROM [deposit_withdraw_stock]
                LEFT JOIN [r] 
                ON [r].[account_code] = [deposit_withdraw_stock].[account_code] AND [r].[date] = [deposit_withdraw_stock].[date]
                WHERE [deposit_withdraw_stock].[date] BETWEEN '{start_date}' AND '{end_date}'
                AND [creator] NOT IN ('User Online trading')
                AND [transaction] = 'Chuyen CK'
        ),
        [fee_CTCK] AS (
            SELECT
                MAX([rse2009].[branch_id]) [branch_id],
                SUM([rse2009].[volume]) [volume],
                SUM(CASE
                    WHEN [rse2009].[date] > '2020-3-19' THEN [rse2009].[volume] * 0.3
                    ELSE [rse2009].[volume] * 0.5
                END) [fee]
            FROM [rse2009]
            GROUP BY [rse2009].[branch_id]
        ),
        [fee_TTBT] AS (
            SELECT
                [r].[branch_id],
                SUM([rod0040].[volume]) [volume],
                SUM(
                    CASE
                        WHEN [rod0040].[date] > '2020-3-16' THEN [rod0040].[charged_volume] * 0.3
                        ELSE [rod0040].[charged_volume] * 0.5
                    END
                ) [fee]
            FROM [rod0040]
            LEFT JOIN [sub_account] on [sub_account].[sub_account] = [rod0040].[sub_account]
            LEFT JOIN [r]
            ON [r].[account_code] = [sub_account].[account_code] AND [r].[date] = [rod0040].[date]
            GROUP BY [r].[branch_id]
        ),
        [res] AS (
            SELECT
                [branch].[branch_id] [Mã Chi Nhánh],
                CASE 
                    WHEN [branch].[branch_id] = '0001' THEN N'HQ'
                    WHEN [branch].[branch_id] = '0101' THEN N'Quận 3'
                    WHEN [branch].[branch_id] = '0102' THEN N'PMH'
                    WHEN [branch].[branch_id] = '0104' THEN N'Q7'
                    WHEN [branch].[branch_id] = '0105' THEN N'TB'
                    WHEN [branch].[branch_id] = '0116' THEN N'P.QLTK1'
                    WHEN [branch].[branch_id] = '0111' THEN N'InB1'
                    WHEN [branch].[branch_id] = '0113' THEN N'IB'
                    WHEN [branch].[branch_id] = '0201' THEN N'Hà Nội'
                    WHEN [branch].[branch_id] = '0202' THEN N'TX'
                    WHEN [branch].[branch_id] = '0301' THEN N'Hải Phòng'
                    WHEN [branch].[branch_id] = '0117' THEN N'Quận 1'
                    WHEN [branch].[branch_id] = '0118' THEN N'P.QLTK3'
                    WHEN [branch].[branch_id] = '0119' THEN N'InB2'
                END [Tên Chi Nhánh],
                ISNULL([fee_TTBT].[fee], 0) [fee_TTBT],
                ISNULL([fee_TTBT].[volume], 0) [fee_TTBT_volume],
                ISNULL([fee_CTCK].[fee], 0) [fee_CTCK],
                ISNULL([fee_CTCK].[volume], 0) [fee_CTCK_volume],
                ISNULL([fee_TTBT].[fee], 0) + ISNULL([fee_CTCK].[fee], 0) [Phí Chuyển Khoản],
                (ISNULL([fee_TTBT].[volume], 0) + ISNULL([fee_CTCK].[volume], 0)) [sum Số Lượng],
                (ROUND(ISNULL([fee_TTBT].[fee], 0), 0) + ROUND(ISNULL([fee_CTCK].[fee], 0), 0)) [sum Phí]
            FROM [fee_TTBT]
            FULL OUTER JOIN [fee_CTCK] ON [fee_CTCK].[branch_id] = [fee_TTBT].[branch_id]
            FULL OUTER JOIN [branch] ON [fee_TTBT].[branch_id] = [branch].[branch_id]
        )
        SELECT
            ROW_NUMBER() OVER (ORDER BY [Mã Chi Nhánh]) [STT],
            [res].*
        FROM [res]
        WHERE [res].[Tên Chi Nhánh] IS NOT NULL
        ORDER BY [Mã Chi Nhánh]
        """,
        connect_DWH_CoSo
    )

    ########################################################################################
    ########################################################################################
    ########################################################################################

    table_title = f'PHÍ CHUYỂN KHOẢN {period}'
    # Write to Báo cáo phí chuyển khoản
    file_name = f'Báo cáo phí chuyển khoản {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',20)
    worksheet.set_column('C:C',15)
    worksheet.set_column('D:D',18)

    title_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'bold':True,
            'align':'center',
        }
    )
    header_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    stt_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'align':'center',
            'border':1,
        }
    )
    tenchinhanh_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'border':1
        }
    )
    machinhanh_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'align':'center',
            'border':1
        }
    )
    philuuky_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    tong_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    tongphiluuky_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'bold':True,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1,
        }
    )
    datenote_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'italic':True,
        }
    )
    sum_fee = transfer_fee['Phí Chuyển Khoản'].sum()
    worksheet.merge_range('A1:D1',table_title,title_format)
    worksheet.write('E1',f'Dữ liệu từ {adj_start_date} đến {adj_end_date}',datenote_format)
    worksheet.write_row(2,0,['STT','Tên Chi Nhánh','Mã Chi Nhánh','Phí Chuyển Khoản'],header_format)
    worksheet.write_column(3,0,transfer_fee['STT'],stt_format)
    worksheet.write_column(3,1,transfer_fee['Tên Chi Nhánh'],tenchinhanh_format)
    worksheet.write_column(3,2,transfer_fee['Mã Chi Nhánh'],machinhanh_format)
    worksheet.write_column(3,3,transfer_fee['Phí Chuyển Khoản'],philuuky_format)
    tong_row = transfer_fee.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,2,'Tổng',tong_format)
    worksheet.write(tong_row,3,sum_fee,tongphiluuky_format)
    writer.close()

    ###################################################################################
    ###################################################################################
    ###################################################################################

    # Write to Báo cáo phí chuyển khoản tổng hợp
    file_name = f'Báo cáo phí chuyển khoản tổng hợp {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',20)
    worksheet.set_column('C:J',15)

    header_format = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'border':1,
        }
    )
    tongcong_format = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'align':'center',
            'border':1,
        }
    )
    column_sum_format = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'align':'center',
            'border':1,
        }
    )
    branch_id_format = workbook.add_format(
        {
            'num_format':'@',
            'font_name':'Times New Roman',
            'align':'center',
            'border':1,
        }
    )
    branch_name_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'valign':'vcenter',
            'border':1,
        }
    )
    number_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    last_col_format = workbook.add_format(
        {
            'bg_color':'FFF300',
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    datenote_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'italic':True,
        }
    )
    worksheet.write('I2','Số lượng',header_format)
    worksheet.write('J2','Phí',header_format)
    worksheet.merge_range('A1:A2','Mã Chi Nhánh',header_format)
    worksheet.merge_range('B1:B2','Tên Chi Nhánh',header_format)
    worksheet.merge_range('C1:E1','Chuyển khoản thanh toán bù trừ',header_format)
    worksheet.merge_range('F1:H1','Chuyển khoản chứng khoán sang CTCK khác',header_format)
    worksheet.merge_range('I1:J1','Tổng cộng',header_format)
    worksheet.write_row('C2',['Số lượng','Phí','Phí làm tròn']*2,header_format)
    worksheet.write_column('A3',transfer_fee['Mã Chi Nhánh'],branch_id_format)
    worksheet.write_column('B3',transfer_fee['Tên Chi Nhánh'],branch_name_format)
    worksheet.write_column('C3',transfer_fee['fee_TTBT_volume'],number_format)
    worksheet.write_column('D3',transfer_fee['fee_TTBT'],number_format)
    worksheet.write_column('E3',np.round(transfer_fee['fee_TTBT'],0),number_format)
    worksheet.write_column('F3',transfer_fee['fee_CTCK_volume'],number_format)
    worksheet.write_column('G3',transfer_fee['fee_CTCK'],number_format)
    worksheet.write_column('H3',np.round(transfer_fee['fee_CTCK'],0),number_format)
    worksheet.write_column('I3',transfer_fee['sum Số Lượng'],number_format)
    worksheet.write_column('J3',transfer_fee['sum Phí'],last_col_format)

    sum_row = transfer_fee.shape[0]+3
    worksheet.merge_range(f'A{sum_row}:B{sum_row}','Tổng cộng',tongcong_format)
    worksheet.write(f'C{sum_row}',transfer_fee['fee_TTBT_volume'].sum(),column_sum_format)
    worksheet.write(f'D{sum_row}',transfer_fee['fee_TTBT'].sum(),column_sum_format)
    worksheet.write(f'E{sum_row}',np.round(transfer_fee['fee_TTBT'],0).sum(),column_sum_format)
    worksheet.write(f'F{sum_row}',transfer_fee['fee_CTCK_volume'].sum(),column_sum_format)
    worksheet.write(f'G{sum_row}',transfer_fee['fee_CTCK'].sum(),column_sum_format)
    worksheet.write(f'H{sum_row}',np.round(transfer_fee['fee_CTCK'],0).sum(),column_sum_format)
    worksheet.write(f'I{sum_row}',transfer_fee['sum Số Lượng'].sum(),column_sum_format)
    worksheet.write(f'J{sum_row}',transfer_fee['sum Phí'].sum(),column_sum_format)

    worksheet.write(f'A{sum_row+3}',f'Dữ liệu từ {adj_start_date} đến {adj_end_date}',datenote_format)
    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
