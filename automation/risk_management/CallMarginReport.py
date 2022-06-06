from automation.risk_management import *
from datawarehouse import BDATE


def run(  # chạy hàng ngày
    run_time=dt.datetime.now()
):
    start = time.time()
    info = get_info('daily',run_time)
    dataDate = info['end_date']
    period = dataDate.replace('-','.')
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
            WHERE [relationship].[date] = '{dataDate}'
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
            [t].[Ngay] = '{dataDate}'
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

    reportDate = BDATE(dataDate,1)
    t1_day = reportDate[-2:]
    t1_month = calendar.month_name[int(reportDate[5:7])]
    t1_year = reportDate[0:4]
    file_name = f'Call Margin Report on {t1_month} {t1_day} {t1_year}.xlsx'
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
            'text_wrap':True,
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
    worksheet.set_column('AA:AC',18,options={'hidden':1})
    worksheet.set_row(0,37)

    for col in 'CEFHKMOQSXYZ':
        worksheet.set_column(f'{col}:{col}',18,options={'hidden':1})

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

