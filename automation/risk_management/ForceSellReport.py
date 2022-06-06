from automation.risk_management import *
from datawarehouse import BDATE


def run(  # chạy hàng ngày
    run_time=dt.datetime.now()
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    dataDate = info['end_date']
    reportDate = BDATE(dataDate,1)
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    badDebtAccounts = [
        '022P002222',
        '022C006827',
        '022C012621',
        '022C012620',
        '022C012622',
        '022C089535',
        '022C050302',
        '022C089950',
        '022C089957',
    ]

    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        WITH 
        [KhachHang] AS (
            SELECT
                [relationship].[sub_account] [TieuKhoan],
                [relationship].[account_code] [TaiKhoan],
                [relationship].[branch_id] [MaChiNhanh],
                [relationship].[broker_id] [MaMoiGioi],
                [branch].[branch_name] [TenChiNhanh],
                [broker].[broker_name] [TenMoiGioi],
                [account].[customer_name] [TenKhachHang]
            FROM [relationship]
            LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
            LEFT JOIN [broker] ON [broker].[broker_id] = [relationship].[broker_id]
            LEFT JOIN [account] ON [account].[account_code] = [relationship].[account_code]
            WHERE [relationship].[date] = '{dataDate}'
        ),
        [DuNo] AS (
            SELECT
                [account_code] [TaiKhoan],
                [principal_outstanding] + [interest_outstanding] + [fee_outstanding] [TongDuNo]
            FROM [margin_outstanding]
            WHERE [type] = N'Trả chậm' AND [date] = '{dataDate}' 
        ),
        [MoiGioi] AS (
            SELECT [TieuKhoan], [DTMG], [EmailMG]
            FROM [VMR0002]
            WHERE [Ngay] = '{dataDate}'
        ),
        [MonVay] AS (
            SELECT 
                [sub_account].[account_code] [TaiKhoan],
                [VLN0001].[MonVay],
                SUM([VLN0001].[GocQuaHan]) + SUM([VLN0001].[LaiQuaHan]) [NoQuaHan],
                MAX(DATEDIFF(DAY,[VLN0001].[NgayGiaiNgan],[VLN0001].[NgayDHKyHan2])) [SoNgayQuaHan]
            FROM [VLN0001]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [VLN0001].[TieuKhoan]
            WHERE [Ngay] = '{dataDate}' AND [GocQuaHan] <> 0 AND [MonVay] IN ('Credit line','Delayed payment')
            GROUP BY [sub_account].[account_code], [VLN0001].[MonVay]
        ),
        [TongHop] AS (
            SELECT
                [KhachHang].[TenChiNhanh],
                [KhachHang].[TaiKhoan],
                CASE
                    WHEN 
                        [ForceSell].[guarantee_debt] <> 0 
                        AND [ForceSell].[actual_mr_ratio] >= 85 
                    THEN 'No HM'
                    WHEN [ForceSell].[guarantee_debt] <> 0 
                        AND [ForceSell].[actual_mr_ratio] < 85 
                    THEN 'No HM/MR'
                    WHEN [KhachHang].[TaiKhoan] IN {iterable_to_sqlstring(badDebtAccounts)} 
                        AND ISNULL([DuNo].[TongDuNo],0) > 0 AND [ForceSell].[actual_dp_ratio] >= 100
                    THEN 'DP Qua Han'
                    WHEN [KhachHang].[TaiKhoan] IN {iterable_to_sqlstring(badDebtAccounts)} 
                        AND ISNULL([DuNo].[TongDuNo],0) > 0 AND [ForceSell].[actual_dp_ratio] < 100
                    THEN 'DP Qua Han/DP'
                    WHEN [KhachHang].[TaiKhoan] IN {iterable_to_sqlstring(badDebtAccounts)} 
                        AND ISNULL([DuNo].[TongDuNo],0) = 0
                    THEN 'MR Qua Han 180n'
                    WHEN [ForceSell].[guarantee_debt] = 0 
                        AND [MonVay].[MonVay] = 'Credit line'
                        AND [MonVay].[SoNgayQuaHan] < 180
                        AND [ForceSell].[actual_mr_ratio] < 85
                    THEN 'MR Qua Han 90n/MR'
                    WHEN [ForceSell].[guarantee_debt] = 0 
                        AND [MonVay].[MonVay] = 'Credit line'
                        AND [MonVay].[SoNgayQuaHan] < 180
                    THEN 'MR Qua Han 90n'
                    WHEN [ForceSell].[guarantee_debt] = 0 
                        AND [MonVay].[MonVay] = 'Credit line'
                        AND [MonVay].[SoNgayQuaHan] >= 180
                        AND [ForceSell].[actual_mr_ratio] < 85
                    THEN 'MR Qua Han 180n/MR'
                    WHEN [ForceSell].[guarantee_debt] = 0 
                        AND [MonVay].[MonVay] = 'Credit line'
                        AND [MonVay].[SoNgayQuaHan] >= 180
                    THEN 'MR Qua Han 180n'
                    WHEN [MonVay].[MonVay] = 'Delayed Payment' AND [ForceSell].[actual_dp_ratio] >= 100
                    THEN 'DP Qua Han'
                    WHEN [MonVay].[MonVay] = 'Delayed Payment' AND [ForceSell].[actual_dp_ratio] < 100
                    THEN 'DP Qua Han/DP'
                    WHEN ISNULL([DuNo].[TongDuNo],0) > 0
                        AND [ForceSell].[actual_dp_ratio] < 100
                    THEN 'DP'
                    WHEN ISNULL([DuNo].[TongDuNo],0) > 0 
                        AND [ForceSell].[actual_dp_ratio] >= 100
                        AND [ForceSell].[actual_mr_ratio] < 85
                    THEN 'MR/Co Suc Mua'
                    WHEN ISNULL([DuNo].[TongDuNo],0) = 0 
                        AND [ForceSell].[actual_mr_ratio] < 85
                    THEN 'MR/Ko Suc Mua' 
                ELSE ''                   
                END [TinhTrangForceSell],
                '' [XuLyForceSell],
                [KhachHang].[TenKhachHang],
                [ForceSell].[actual_mr_ratio] [TLThucTeMR],
                [ForceSell].[actual_dp_ratio] [TLThucTeTC],
                [ForceSell].[date_of_first_call] [NgayCallDauTien],
                [ForceSell].[selling_value] [TienMatVe100],
                [ForceSell].[executing_amount] [TienMatVe85],
                CASE
                    WHEN [ForceSell].[actual_asset_to_guarantee] - [ForceSell].[converted_asset] > 0
                    THEN [ForceSell].[actual_asset_to_guarantee] - [ForceSell].[converted_asset]
                    ELSE 0
                END [TienMatVeRTTDP],
                [ForceSell].[guarantee_debt] [NoHanMuc],
                [ForceSell].[mr_dp_overdue_amount] [NoMRTCQuaHan],
                [KhachHang].[TenMoiGioi],
                [ForceSell].[mr_dp_due_amount] [NoMRTCDenHan],
                [ForceSell].[contract_type] [MaLoaiHinh],
                [ForceSell].[total_cash] [TongTien],
                ISNULL([DuNo].[TongDuNo],0) [DP],
                [ForceSell].[actual_ratio] [TLThucTe],
                [ForceSell].[actual_t0_ratio] [TLThucTeT0],
                [ForceSell].[depository_fee_debt] [NoPhiLK],
                0 AS [ChoBan],
                [ForceSell].[date_of_last_sms] [NgaySMSCuoiCung],
                [ForceSell].[time_of_last_sms] [ThoiGianSMSCuoiCung],
                [ForceSell].[days_maintain_call] [SoNgayCallConLai],
                [ForceSell].[days_warning] [SoNgayCanhBao],
                [ForceSell].[date_of_last_call] [NgayCallCuoiCung],
                [ForceSell].[date_of_trigger] [NgayTrigger],
                [ForceSell].[additional_deposit_amount] [SoTienNopThem],
                [ForceSell].[selling_amount] [SoTienCanPhaiBan],
                [ForceSell].[executing_overdue_amount] [SoTienQuaHanCanXuLy],
                [ForceSell].[remain_cash_after_sell] [SoTienDuSauBan],
                [ForceSell].[days_execution] [SoNgayRoiVaoXuLy],
                [ForceSell].[sub_account] [TieuKhoan],
                [ForceSell].[sub_account_type] [TenLoaiHinh],
                '' [DTLienLac],
                '' [Email],
                [ForceSell].[careby] [DiemHoTro],
                [MoiGioi].[EmailMG] [EmailMG],
                [MoiGioi].[DTMG] AS [DTMG],
                [KhachHang].[MaChiNhanh],
                [ForceSell].[force_sell_ordered_value] [TongGTLenhGiaiChap],
                [ForceSell].[force_sell_matched_value] [TongGTKhopBanGiaiChap],
                [ForceSell].[rate_add] [TLBoSung],
                [ForceSell].[days_add] [SoNgayCongBoSung],
                [ForceSell].[right_event] [SuKienQuyen],
                [ForceSell].[days_base_call] [SoNgayCallCoSo],
                [ForceSell].[safe_rate] [TLAnToan],
                [ForceSell].[total_loan_amount] [TongDuNoVay],
                [ForceSell].[converted_asset] [TaiSanVayQuyDoi],
                [ForceSell].[actual_asset_to_guarantee] [TSThucBaoDamTyLeKyQuyDuyTri]
            FROM [vmr0003] [ForceSell]
            LEFT JOIN [KhachHang] ON [KhachHang].[TieuKhoan] = [ForceSell].[sub_account]
            LEFT JOIN [MoiGioi] ON [MoiGioi].[TieuKhoan] = [ForceSell].[sub_account]
            LEFT JOIN [DuNo] ON [DuNo].[TaiKhoan] = [KhachHang].[TaiKhoan]
            LEFT JOIN [MonVay] ON [MonVay].[TaiKhoan] = [KhachHang].[TaiKhoan]
            WHERE [ForceSell].[date] = '{dataDate}'
        )
        SELECT 
            [TongHop].*,
            CASE
                WHEN [TongHop].[TaiKhoan] IN {iterable_to_sqlstring(badDebtAccounts)} THEN 'BadDebt'
                ELSE ''
            END [BadDebt]
        FROM [TongHop]
        ORDER BY [BadDebt], [TaiKhoan]
        """,
        connect_DWH_CoSo,
    )
    ###################################################
    ###################################################
    ###################################################

    reportDay = reportDate[-2:]
    reportMonth = reportDate[5:7]
    reportYear = reportDate[0:4]
    file_name = f'Force sell {reportMonth}.{reportDay}.{reportYear}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    header_normal_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Calibri',
            'text_wrap':True
        }
    )
    header_red_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Calibri',
            'text_wrap':True,
            'font_color':'#FF0000',
        }
    )
    account_normal_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Calibri',
        }
    )
    account_baddebt_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Calibri',
            'bg_color':'#ED7D31',
        }
    )
    text_left_format = workbook.add_format(
        {
            'align':'left',
            'valign':'top',
            'font_size':10,
            'font_name':'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Calibri',
            'num_format': '#,##0'
        }
    )
    money_red_format = workbook.add_format(
        {
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Calibri',
            'num_format': '#,##0',
            'font_color':'#FF0000',
        }
    )
    number_format = workbook.add_format(
        {
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Calibri',
        }
    )
    number_red_format = workbook.add_format(
        {
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Calibri',
            'font_color':'#FF0000',
        }
    )
    date_format = workbook.add_format(
        {
            'align':'left',
            'valign':'top',
            'font_size':10,
            'font_name':'Calibri',
            'num_format':'dd/mm/yyyy'
        }
    )
    ###################################################
    ###################################################
    ###################################################

    headers = [
        'Tên chi nhánh',
        'Số TK lưu ký',
        'Tình Trạng Force sell',
        'Xử lý force sell',
        'Tên khách hàng',
        'TL thực tế MR',
        'TL thực tế TC',
        'Ngày bắt đầu Call',
        'Tiền mặt nộp về 100%',
        'Tiền mặt nộp về 85%',
        'Tiền mặt nộp về Rtt_DP',
        'Nợ hạn mức',
        'Nợ MR + TC quá hạn',
        'Tên MG',
        'Nợ MR + TC đến hạn',
        'Mã loại hình',
        'Tổng tiền',
        'DP',
        'TL thực tế',
        'TL thực tế T0',
        'Nợ phí LK',
        'Chờ bán',
        'Ngày SMS cuối',
        'Thời gian SMS cuối',
        'Số ngày duy trì call',
        'Số ngày rơi vào cảnh báo',
        'Ngày bắt đầu call',
        'Ngày bắt đầu trigger',
        'Số tiền nộp thêm',
        'Số tiền cần phải bán',
        'Số tiền quá hạn cần phải xử lý',
        'Số tiền dư sau bán',
        'Số ngày rơi vào xử lý',
        'Số tiểu khoản',
        'Tên loại hình',
        'ĐT liên lạc',
        'E-mail',
        'Điểm hỗ trợ',
        'E-mail MG',
        'ĐT MG',
        'Code chi nhánh',
        'Tổng GT lệnh giải chấp',
        'Tổng GT khớp bán giải chấp',
        'Tỷ lệ bổ sung',
        'Số ngày cộng bổ sung',
        'Chạm sự kiện quyền',
        'Số ngày call cơ sở',
        'TL an toàn',
        'Tổng dư nợ vay',
        'Tài sản vay qui đổi',
        'TS thực có tối thiểu để bảo đảm TLKQ duy trì',
    ]
    worksheet = workbook.add_worksheet('Sheet1')
    worksheet.set_column('A:A',10)
    worksheet.set_column('B:B',10)
    worksheet.set_column('C:D',8)
    worksheet.set_column('E:E',25)
    worksheet.set_column('F:G',9)
    worksheet.set_column('H:H',10)
    worksheet.set_column('I:M',13)
    worksheet.set_column('N:N',25)
    worksheet.set_column('O:O',7)
    worksheet.set_column('P:P',5)
    worksheet.set_column('Q:R',11)
    worksheet.set_column('S:V',9)
    worksheet.set_column('W:W',10)
    worksheet.set_column('X:Z',7)
    worksheet.set_column('AA:AB',10)
    worksheet.set_column('AC:AE',13)
    worksheet.set_column('AF:AG',6)
    worksheet.set_column('AH:AH',10)
    worksheet.set_column('AI:AV',8)
    worksheet.set_column('AW:AY',14)

    for num,header in enumerate(headers):
        if header in ('Tình Trạng Force sell','Xử lý force sell','Tiền mặt nộp về Rtt_DP','DP'):
            fmt = header_red_format
        else:
            fmt = header_normal_format
        worksheet.write(0,num,header,fmt)

    # Format lần 1 cho TH tổng quát
    for colNum,colName in enumerate(table.columns[:-1]):
        if colName == 'TaiKhoan':
            fmt = account_normal_format
        elif colName.lower().startswith('tl'):
            fmt = number_format
        elif pd.api.types.is_numeric_dtype(table[colName]):
            fmt = money_format
        elif pd.api.types.is_datetime64_dtype(table[colName]):
            table[colName] = table[colName].replace(pd.NaT,dt.datetime(9999,12,31))
            fmt = date_format
        else:
            fmt = text_left_format
        worksheet.write_column(1,colNum,table[colName],fmt)

    # Format lần 2 các TH cần lưu ý
    for row in range(table.shape[0]):
        Account = table.loc[table.index[row],'TaiKhoan']
        NoHanMuc = table.loc[table.index[row],'NoHanMuc']
        TLThucTeMR = table.loc[table.index[row],'TLThucTeMR']
        NoMRTCQuaHan = table.loc[table.index[row],'NoMRTCQuaHan']
        NoDP = table.loc[table.index[row],'DP']
        TLThucTeDP = table.loc[table.index[row],'TLThucTeTC']

        writtenRow = row + 1
        for col,colName in enumerate(table.columns):
            if colName == 'TaiKhoan' and Account in badDebtAccounts:
                worksheet.write(writtenRow,col,Account,account_baddebt_format)
        for col,colName in enumerate(table.columns):
            if NoDP > 0 and TLThucTeDP < 100:
                if colName in ('TienMatVeRTTDP','TLThucTeTC'):
                    value = table.loc[table.index[row],colName]
                    worksheet.write(writtenRow,col,value,number_red_format)
            else:
                if NoHanMuc != 0:
                    condition1 = TLThucTeMR >= 85 and colName in ('NoHanMuc',)
                    condition2 = TLThucTeMR < 85 and colName in ('TLThucTeMR','TienMatVe100','TienMatVe85')
                    if condition1 or condition2:
                        value = table.loc[table.index[row],colName]
                        worksheet.write(writtenRow,col,value,money_red_format)
                else:
                    condition1 = NoMRTCQuaHan != 0 and colName in ('NoMRTCQuaHan',)
                    condition2 = NoMRTCQuaHan == 0 and colName in ('TienMatVe100','TienMatVe85')
                    condition3 = NoMRTCQuaHan == 0 and colName in ('TLThucTeMR',)
                    if condition1 or condition2:
                        value = table.loc[table.index[row],colName]
                        worksheet.write(writtenRow,col,value,money_red_format)
                    elif condition3:
                        value = table.loc[table.index[row],colName]
                        worksheet.write(writtenRow,col,value,number_red_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

