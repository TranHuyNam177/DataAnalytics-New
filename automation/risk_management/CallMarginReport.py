from automation.risk_management import *
from datawarehouse import BDATE


"""
TK bị lệch ngày 29/04/2022
1. 022C017132	0117000598	ĐINH TIÊN HOÀNG
- Trạng thái call trên DB là No - Trạng thái call từ báo cáo bên QLRR là Yes
2. 022C018098	0203001688	NGUYỄN KIM NHUNG
- DB ngày 29/4/2022 không trả ra kết quả
3. 022C024544	0202001405	NGUYỄN THỊ CHĂM
- DB ngày 29/4/2022 không trả ra kết quả
4. 022C027172	0102002141	DƯƠNG THỊ CÁT ĐẰNG
- DB ngày 29/4/2022 không trả ra kết quả
5. 022C042269	0202003304	ĐẶNG NHÂN THỦY
- DB ngày 29/4/2022 không trả ra kết quả
6. 022C042343	0101003903	PHAN THỊ NGỌC NỮ
- DB ngày 29/4/2022 không trả ra kết quả
7. 022C087587	0101087587	VŨ KIM LIÊN
- DB ngày 29/4/2022 không trả ra kết quả
8. 022C240592	0117002925	NGUYỄN VĂN HỒ
- Trạng thái call trên DB là No - Trạng thái call từ báo cáo bên QLRR là Yes
9. 022C357999	0202002369	LÊ CHÍ CƯỜNG
- DB ngày 29/4/2022 không trả ra kết quả
10. 022C567803	0117002907	NGUYỄN ĐÌNH TƯ
- Trạng thái call trên DB là No - Trạng thái call từ báo cáo bên QLRR là Yes
11. 022C777999	0201001455	NGUYỄN XUÂN CỬ
- DB ngày 29/4/2022 không trả ra kết quả
"""
"""
TK bị lệch ngày 09/05/2022
1. TK bị dư so với file mẫu:
022C012610 - có trên SQL nhưng ko có trên Flex (dữ liệu chốt sau batch cuối ngày 06/05/2022)
2. TK bị lệch số liệu:
022C076999 - Lệch cột Số tiền phải bán,cột Supplementary Amount,cột Overdue + Due to date amount,cột Số tiền nộp thêm gốc 
và cột THIẾU HỤT
(so SQL với Flex)
    - Số tiền phải bán:
        SQL:  1,196,063,222 + FLEX: 622,617,641
    - Supplementary Amount:
        SQL:  1,196,063,222 + FLEX: 622,617,641
    - Overdue + Due to date amount
        SQL:  1,195,690,750 + FLEX: 0
    - Số tiền nộp thêm gốc:
        SQL:  1,196,063,222 + FLEX: 622,617,641
    - THIẾU HỤT:
        SQL:  1,196,063,222 + FLEX: 622,617,641
"""


def run(  # chạy hàng ngày
    run_time=dt.datetime.now()
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    t0_date = info['end_date']
    t1_date = BDATE(t0_date,1)
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        WITH [i] AS (
            SELECT 
                [relationship].[date],
                [relationship].[account_code],
                [relationship].[broker_id],
                [relationship].[sub_account],
                [account].[customer_name],
                [broker].[broker_name]
            FROM [relationship]
            LEFT JOIN [account] ON [relationship].[account_code] = [account].[account_code]
            LEFT JOIN [broker] ON [relationship].[broker_id] = [broker].[broker_id]
            WHERE [relationship].[date] = '{t0_date}'
        )
        SELECT 
            [i].[account_code] [AccountCode],
            [t].[TieuKhoan],
            [i].[customer_name] [CustomerName],
            [t].[MaLoaiHinh],
            [t].[TenLoaiHinh],
            [i].[broker_name] [BrokerName],
            [t].[SoNgayDuyTriCall],
            [t].[NgayBatDauCall],
            [t].[NgayHanCuoiCall],
            [t].[TrangThaiCall],
            [t].[LoaiCall],
            [t].[SoNgayCallVuot],
            [t].[SoTienPhaiNop],
            [t].[SoTienPhaiBan],
            [t].[SoTienDenHanVaQuaHan],
            [t].[SoTienNopThemGoc],
            [t].[ChamSuKienQuyen],
            [t].[TLThucTe],
            [t].[TLThucTeMR],
            [t].[TLThucTeTC],
            [t].[TyLeAnToan],
            [t].[ToDuNoVay],
            [t].[NoMRTCBL],
            [t].[TaiSanVayQuiDoi],
            [t].[TSThucCoToiThieuDeBaoDamTLKQDuyTri],
            [t].[ThieuHut],
            [t].[EmailMG],
            [t].[DTMG]
        FROM [VMR0002] [t]
        LEFT JOIN [i] ON [i].[sub_account] = [t].[TieuKhoan] 
            AND [i].[date] = [t].[Ngay]
        WHERE 
            [t].[Ngay] = '{t0_date}'
            AND [t].[TenLoaiHinh] = N'Margin'
            AND [t].[TrangThaiCall] = N'Yes'
            AND [t].[LoaiCall] <> ''
            AND [t].[NgayBatDauCall] IS NOT NULL 
            AND [t].[NgayHanCuoiCall] IS NOT NULL
        ORDER BY [account_code]
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    t1_day = t1_date[-2:]
    t1_month = calendar.month_name[int(t1_date[5:7])]
    t1_year = t1_date[0:4]
    file_name = f'Call Margin Report on {t1_day} {t1_month} {t1_year}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Format
    headers_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'top',
            'font_size':12,
            'font_name':'Calibri',
            'text_wrap':True
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    stt_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    decimal_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        }
    )
    integer_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format':'dd/mm/yyyy'
        }
    )

    ###################################################
    ###################################################
    ###################################################

    headers = [
        'No',
        'Account No',
        'Số tiểu khoản',
        'Name',
        'Mã loại hình',
        'Tên loại hình',
        'Broker Name',
        'Số ngày duy trì call',
        'Call date',
        'Call deadline',
        'Trạng thái call',
        'Call Type',
        'Số ngày call vượt',
        'Supplementary Amount',
        'Số tiền phải bán',
        'Overdue + Due to date amount',
        'Số tiền nộp thêm gốc',
        'Ex-right',
        'TL thực tế',
        'Rtt-MR',
        'Rtt-DP',
        'Rat',
        'Total Outstanding',
        'Nợ MR + TC + BL',
        'Tài sản vay qui đổi',
        'TS thực có tối thiểu để bảo đảm TLKQ duy trì',
        'Thiếu hụt',
        'E-mail MG',
        'ĐT MG',
    ]

    worksheet = workbook.add_worksheet('Sheet1')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:ZZ',18)
    worksheet.set_column('A:A',4)
    worksheet.set_column('B:B',16)
    worksheet.set_column('D:D',27)
    worksheet.set_column('G:G',28)
    worksheet.set_column('I:J',13)
    worksheet.set_column('L:L',19)
    worksheet.set_column('N:N',22)
    worksheet.set_column('P:P',18)
    worksheet.set_column('R:R',12)
    worksheet.set_column('T:U',11)
    worksheet.set_column('V:V',8)
    worksheet.set_column('W:W',18)
    worksheet.set_column('AA:AC',0)
    worksheet.set_row(0,37)

    for col in 'CEFHKMOQSXYZ':
        worksheet.set_column(f'{col}:{col}',0)

    worksheet.write_row('A1',headers,headers_format)
    worksheet.write_column('A2',np.arange(table.shape[0])+1,stt_format)
    for colNum, colName in enumerate(table.columns,1):
        if colName.lower().startswith('tl'):
            fmt = decimal_format
        elif colName.lower().startswith('songay'):
            fmt = integer_format
        elif pd.api.types.is_numeric_dtype(table[colName]):
            fmt = money_format
        elif pd.api.types.is_datetime64_dtype(table[colName]):
            fmt = date_format
        else:
            fmt = text_left_format
        worksheet.write_column(1,colNum,table[colName],fmt)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

