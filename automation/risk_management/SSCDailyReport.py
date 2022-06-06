from automation.risk_management import *
from datawarehouse import BDATE

def run(  # chạy hàng ngày sau batch giữa ngày
    run_time=dt.datetime.now()
):
    
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    dataDate = info['end_date']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    # "Vốn chủ sở hữu"
    equityPHS = 1565140

    detailTable = pd.read_sql(
        f"""
        WITH 
        [DanhMuc] AS (
            SELECT [MaCK]
            FROM [DanhMucChoVayMargin]
            WHERE [Ngay] = (SELECT MAX([Ngay]) FROM [DanhMucChoVayMargin])
                AND [MaCK] NOT IN ('CII','CTI','POM','SKG','VIC')
        ),
        [CKNY] AS (
            SELECT 
                [Ticker] [MaCK], 
                TRY_CAST([t].[value] AS BIGINT) / 1000 [KLNY] 
            FROM [DWH-ThiTruong].[dbo].[SecuritiesInfoVSD] 
            CROSS APPLY STRING_SPLIT([DWH-ThiTruong].[dbo].[SecuritiesInfoVSD].[Value],' ') [t] 
            WHERE [Attribute] LIKE '%Quantity%registered%' AND [t].[value] LIKE '[0-9]%'
        ),
        [Tien] AS (
            SELECT
                [account_code] [TaiKhoan],
                SUM([cash]) [TienMat]
            FROM [rmr0062]
            WHERE [date] = '{dataDate}' AND [rmr0062].[loan_type] = 1
            GROUP BY [account_code]
        ),
        [DuNo] AS (
            SELECT
                [account_code] [TaiKhoan],
                SUM([principal_outstanding]) [DuNoGoc]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{dataDate}'
                AND [margin_outstanding].[type] IN (N'Margin',N'Trả chậm',N'Bảo lãnh')
            GROUP BY [account_code]
        ),
        [Room] AS (
            SELECT
                [ticker] [MaCK],
                [sub_account] [TieuKhoan],
                SUM([used_system_room]) + SUM([vmr0104].[special_room]) [TongRoomDaSuDung]
            FROM [vmr0104]
            WHERE [date] = '{dataDate}'
            GROUP BY [ticker], [sub_account]
        ),
        [QuanHe] AS (
            SELECT
                [account_code] [TaiKhoan],
                [sub_account] [TieuKhoan]
            FROM [relationship]
            WHERE [date] = '{dataDate}'
        ),
        [ThiTruong] AS (
            SELECT
                [Ticker] [MaCK],
                [Close] * 1000 [GiaDongCua]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
            WHERE [Date] = '{dataDate}'
        ),
        [SanGiaoDich] AS (
            SELECT
                [Ticker] [MaCK],
                [Exchange] [San]
            FROM [DWH-ThiTruong].[dbo].[DanhSachMa]
            WHERE [Date] = (SELECT MAX([Date]) FROM [DWH-ThiTruong].[dbo].[DanhSachMa])
        ), 
        [RawTable] AS (
            SELECT
                {equityPHS} AS [VCSH],
                [DanhMuc].[MaCK],
                [Room].[TongRoomDaSuDung],
                [Room].[TongRoomDaSuDung] * [ThiTruong].[GiaDongCua] [TaiSanTheoCoPhieu],
                [Room].[TongRoomDaSuDung] * [ThiTruong].[GiaDongCua] / SUM([Room].[TongRoomDaSuDung] * [ThiTruong].[GiaDongCua]) OVER (PARTITION BY [QuanHe].[TaiKhoan]) [TyTrong],
                CASE
                    WHEN [DuNo].[DuNoGoc] - [Tien].[TienMat] > 0
                        THEN [DuNo].[DuNoGoc] - [Tien].[TienMat]
                    ELSE 0
                END [DuNoRong]
            FROM [DanhMuc]
            LEFT JOIN [Room] ON [Room].[MaCK] = [DanhMuc].[MaCK]
            LEFT JOIN [QuanHe] ON [QuanHe].[TieuKhoan] = [Room].[TieuKhoan]
            LEFT JOIN [ThiTruong] ON [ThiTruong].[MaCK] = [Room].[MaCK]
            LEFT JOIN [DuNo] ON [DuNo].[TaiKhoan] = [QuanHe].[TaiKhoan]
            LEFT JOIN [Tien] ON [Tien].[TaiKhoan] = [QuanHe].[TaiKhoan]
        )
        SELECT
            [RawTable].[MaCK],
            CASE
                WHEN SUM([RawTable].[DuNoRong]*[RawTable].[TyTrong])/1000000 > MAX([RawTable].[VCSH])*0.1
                    THEN MAX([RawTable].[VCSH])*0.1
                ELSE SUM([RawTable].[DuNoRong]*[RawTable].[TyTrong])/1000000
            END [DuNoChoVayGDKQ],
            CASE
                WHEN SUM([RawTable].[TongRoomDaSuDung]) / 1000 / MAX([CKNY].[KLNY]) > 0.05
                    THEN MAX([CKNY].[KLNY])*0.05
                ELSE SUM([RawTable].[TongRoomDaSuDung]) / 1000
            END [SoLuongCKChoVayCTCK],
            MAX([CKNY].[KLNY]) [KLNY],
            CASE
                WHEN SUM([RawTable].[DuNoRong]*[RawTable].[TyTrong])/1000000 > MAX([RawTable].[VCSH])*0.1
                    THEN 0.1
                ELSE SUM([RawTable].[DuNoRong]*[RawTable].[TyTrong])/1000000/MAX([RawTable].[VCSH])
            END [DuNo/VCSH],
            MAX([RawTable].[VCSH]) [VonChuSoHuuCTCK],
            CASE
                WHEN SUM([RawTable].[TongRoomDaSuDung]) / 1000 / MAX([CKNY].[KLNY]) > 0.05
                    THEN 0.05
                ELSE SUM([RawTable].[TongRoomDaSuDung]) / 1000 / MAX([CKNY].[KLNY]) 
            END [CKChoVay/CKNY],
            MAX([SanGiaoDich].[San]) [SanGiaoDich]
        FROM [RawTable]
        LEFT JOIN [CKNY] ON [CKNY].[MaCK] = [RawTable].[MaCK]
        LEFT JOIN [SanGiaoDich] ON [RawTable].[MaCK] = [SanGiaoDich].[MaCK]
        GROUP BY [RawTable].[MaCK]
        ORDER BY [RawTable].[MaCK]
        """,
        connect_DWH_CoSo
    )
    pivotTable = detailTable.groupby('SanGiaoDich')[['SoLuongCKChoVayCTCK','DuNoChoVayGDKQ']].sum()

    ###################################################
    ###################################################
    ###################################################

    t0_day = dataDate[-2:]
    t0_month = dataDate[5:7]
    t0_year = dataDate[0:4]
    file_name = f'180426__RMD_SCMS_Bao cao ngay truoc 8AM {t0_day}{t0_month}{t0_year} New.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    # Format
    title_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    sub_title_format = workbook.add_format(
        {
            'border':1,
            'italic':True,
            'align':'right',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    headers_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    text_bold_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    number_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'#,##0_);(#,##0)',
        }
    )
    number_bold_format = workbook.add_format(
        {
            'bold':True,
            'align':'right',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'#,##0_);(#,##0)',
        }
    )
    number_incell_bold_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'right',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'#,##0_);(#,##0)',
        }
    )
    percent_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'0.000%',
        }
    )
    ###################################################
    ###################################################
    ###################################################

    title_sheet = 'TÌNH HÌNH GIAO DỊCH KÝ QUỸ'
    sub_title = 'Đơn vị : nghìn cổ phiếu, triệu đồng'
    headers = [
        'STT',
        'Mã CK',
        'Dư nợ cho vay GDKQ (1)',
        'Số lượng chứng khoán cho vay của CTCK (2)',
        'Số lượng chứng khoán niêm yết của TCNY (3)',
        'Vốn chủ sở hữu của CTCK (4)',
        'Tỷ lệ dư nợ/VCSH (1)/(4)',
        'Tỷ lệ CK cho vay/CKNY (2)/(3)',
        'Sàn',
    ]
    sub_headers = [
        'Sàn',
        'Khối lượng TSĐB',
        'Dư nợ mã',
    ]
    worksheet = workbook.add_worksheet('Báo Cáo')
    worksheet.set_column('A:B',8)
    worksheet.set_column('C:C',10)
    worksheet.set_column('D:D',15)
    worksheet.set_column('E:E',16)
    worksheet.set_column('F:G',12)
    worksheet.set_column('H:H',14)
    worksheet.set_column('I:I',9)
    worksheet.set_column('K:K',15)
    worksheet.set_column('L:L',20)
    worksheet.set_column('M:M',11)
    worksheet.set_row(2,66)

    worksheet.merge_range('A1:I1',title_sheet,title_format)
    worksheet.merge_range('A2:I2',sub_title,sub_title_format)
    worksheet.write_row('A3',headers,headers_format)
    worksheet.write_column('A4',np.arange(detailTable.shape[0])+1,text_center_format)
    worksheet.write_column('B4',detailTable['MaCK'],text_center_format)
    worksheet.write_column('C4',detailTable['DuNoChoVayGDKQ'],number_format)
    worksheet.write_column('D4',detailTable['SoLuongCKChoVayCTCK'],number_format)
    worksheet.write_column('E4',detailTable['KLNY'],number_format)
    worksheet.write_column('F4',detailTable['VonChuSoHuuCTCK'],number_format)
    worksheet.write_column('G4',detailTable['DuNo/VCSH'],percent_format)
    worksheet.write_column('H4',detailTable['CKChoVay/CKNY'],percent_format)
    worksheet.write_column('I4',detailTable['SanGiaoDich'],text_center_format)
    sum_row = detailTable.shape[0] + 4
    worksheet.write(f'C{sum_row}',detailTable['DuNoChoVayGDKQ'].sum(),number_bold_format)
    worksheet.write(f'D{sum_row}',detailTable['SoLuongCKChoVayCTCK'].sum(),number_bold_format)
    worksheet.write_row('K3',sub_headers,headers_format)
    worksheet.write_column('K4',pivotTable.index,text_center_format)
    worksheet.write_column('L4',pivotTable['SoLuongCKChoVayCTCK'],number_format)
    worksheet.write_column('M4',pivotTable['DuNoChoVayGDKQ'],number_format)
    worksheet.write('K6','Total',headers_format)
    worksheet.write_row('L6',pivotTable.sum(),number_incell_bold_format)
    worksheet.merge_range('K7:L7','200% VCSH',text_bold_format)
    worksheet.write('M7',equityPHS*2,number_bold_format)
    worksheet.write('M8','=M6-M7',number_bold_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')
