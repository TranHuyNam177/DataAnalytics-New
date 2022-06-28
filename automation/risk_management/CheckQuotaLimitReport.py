from automation.risk_management import *
from datawarehouse import *


def generateTempData():

    """
    Chạy lúc 16:10 chiều mỗi ngày, nếu ngày thường thì save file
    """

    info = get_info('daily',dt.datetime.now())
    dataDate = info['end_date']
    SYNC(connect_DWH_CoSo,'spVMR0001',FrDate=dataDate,ToDate=dataDate)
    SYNC(connect_DWH_CoSo,'spVMR9003',FrDate=dataDate,ToDate=dataDate)
    SYNC(connect_DWH_CoSo,'sprelationship',FrDate=dataDate,ToDate=dataDate)
    table = pd.read_sql(
        f"""
        WITH 
        [i] AS (
            SELECT
                [relationship].[sub_account],
                [relationship].[account_code],
                [relationship].[broker_id],
                [account].[customer_name],
                [broker].[broker_name]
            FROM [relationship]
            LEFT JOIN [account] ON [relationship].[account_code] = [account].[account_code]
            LEFT JOIN [broker] ON [relationship].[broker_id] = [broker].[broker_id]
            WHERE [relationship].[date] = '{dataDate}'
        )
        SELECT
            [t].[MaLoaiHinh],
            [t].[TenLoaiHinh],
            [i].[account_code] [SoTKLuuKy],
            [t].[TieuKhoan] [SoTieuKhoan],
            [i].[customer_name] [TenKhachHang],
            [i].[broker_name] [TenMoiGioi],
            [t].[Tien],
            [t].[TLMRThucTe],
            [t].[TLTCThucTe],
            [c].[DuTinhGiaiNganT0]
        FROM [VMR0001] [t]
        LEFT JOIN [i] ON [i].[sub_account] = [t].[TieuKhoan]
        LEFT JOIN [VMR9003] [c] ON [i].[sub_account] = [c].[TieuKhoan] AND [c].[Ngay] = [t].[Ngay]
        WHERE [t].[Ngay] = '{dataDate}'
        """,
        connect_DWH_CoSo,
    )
    table.to_pickle(join(dirname(__file__),'temp',f'TempDataQuotaLimit_{dataDate.replace(".","")}.pickle'))

def run(
    run_time=dt.datetime.now()
):

    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    dataDate = info['end_date'].replace('.','-')
    t1Date = BDATE(dataDate,-1)
    reportDate = BDATE(dataDate,1)
    folder_name = info['folder_name']

    # create_folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ################################################################
    ################################################################
    ################################################################

    dataPart1 = pd.read_pickle(
        join(dirname(__file__),'temp',f'TempDataQuotaLimit_{dataDate.replace("-","")}.pickle')
    )
    dataPart2 = pd.read_sql(
        f"""
        WITH
        [VMR0001T1] AS (
            SELECT
                [sub_account].[account_code] [TaiKhoan],
                [VMR0001].[TLMRThucTe] [TLMRDN],
                [VMR0001].[TLTCThucTe] [TLTCDN]
            FROM [VMR0001]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [VMR0001].[TieuKhoan]
            WHERE [VMR0001].[Ngay] = '{t1Date}'
				AND [VMR0001].[TenLoaiHinh] LIKE 'MR%'
        ),
        [RLN0005T0] AS (
            SELECT
                [sub_account].[account_code] [TaiKhoan],
                SUM([RLN0005].[SoTienCapBaoLanh]) [SoTienCapBaoLanh]
            FROM [RLN0005]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [RLN0005].[TieuKhoan]
			WHERE [RLN0005].[Ngay] = '{dataDate}'
			GROUP BY [sub_account].[account_code]
        ),
        [RLN0006TCT0] AS (
            SELECT
                [account_code] [TaiKhoan],
                SUM([principal_outstanding])+SUM([interest_outstanding])+SUM([fee_outstanding]) [TongDuNo]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{dataDate}' AND [margin_outstanding].[type] = N'Trả chậm'
			GROUP BY [account_code]
        ),
        [RLN0006BLT0] AS (
            SELECT
                [account_code] [TaiKhoan],
                SUM([principal_outstanding]) [DuNoBaoLanh]
            FROM [margin_outstanding]
            WHERE [type] = N'Bảo lãnh' AND [date] = '{dataDate}'
			GROUP BY [account_code]
        ),
        [RSA0004T0] AS (
            SELECT
                [sub_account].[account_code] [TaiKhoan],
                SUM([transactional_record].[amount]) [Tien],
                MAX([transactional_record].[GioDuyet]) [GioDuyet]
            FROM [transactional_record]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [transactional_record].[sub_account]
            WHERE LEN([transactional_record].[sub_account]) = 10
                AND [transactional_record].[date] = '{dataDate}'
                AND ([transactional_record].[transaction_id] = '1141' 
                    OR [transactional_record].[transaction_id] = '1120' AND [transactional_record].[remark] LIKE N'Nhận chuyển khoản nội bộ %Thường%'
                )
                AND [transactional_record].[GioDuyet] BETWEEN '16:10:00' AND '19:00:00'
            GROUP BY [sub_account].[account_code]
        )
        SELECT
            COALESCE([VMR0001T1].[TaiKhoan],[RLN0005T0].[TaiKhoan],[RLN0006TCT0].[TaiKhoan],[RLN0006BLT0].[TaiKhoan],[RSA0004T0].[TaiKhoan]) [SoTKLuuKy],
            [VMR0001T1].[TLMRDN],
            [VMR0001T1].[TLTCDN],
            [RLN0005T0].[SoTienCapBaoLanh] [RLN0005],
            [RLN0006TCT0].[TongDuNo] [RLN0006],
            [RSA0004T0].[Tien] [TienRSA0004],
            [RSA0004T0].[GioDuyet] [TimeRSA0004],
            [RLN0006BLT0].[DuNoBaoLanh] [DuNoBaoLanh]
        FROM [VMR0001T1]
        FULL JOIN [RLN0005T0] ON [RLN0005T0].[TaiKhoan] = [VMR0001T1].[TaiKhoan]
        FULL JOIN [RLN0006TCT0] ON [RLN0006TCT0].[TaiKhoan] = [VMR0001T1].[TaiKhoan]
        FULL JOIN [RLN0006BLT0] ON [RLN0006BLT0].[TaiKhoan] = [VMR0001T1].[TaiKhoan]
        FULL JOIN [RSA0004T0] ON [RSA0004T0].[TaiKhoan] = [VMR0001T1].[TaiKhoan]
        WHERE [RLN0005T0].[SoTienCapBaoLanh] IS NOT NULL
        """,
        connect_DWH_CoSo,
    )
    table = pd.merge(dataPart1,dataPart2,how='left',on='SoTKLuuKy')
    table.dropna(subset=['RLN0005'],inplace=True)
    table.sort_values('RLN0005',ascending=False,inplace=True)

    ################################################################
    ################################################################
    ################################################################

    reportDay = reportDate[-2:]
    reportMonth = reportDate[5:7]
    reportYear = reportDate[:4]
    file_name = f'Checking Quota {reportDay}{reportMonth}{reportYear}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    # Set Format
    header_1_format = workbook.add_format({
        'align':'left',
        'valign':'vcenter',
        'text_wrap':True,
        'font_size':11,
        'font_name':'Calibri'
    })
    header_2_format = workbook.add_format({
        'bold':True,
        'align':'center',
        'valign':'vcenter',
        'text_wrap':True,
        'font_size':11,
        'font_name':'Calibri',
        'bg_color':'#FFFF00'
    })
    text_left_format = workbook.add_format({
        'align':'left',
        'valign':'vcenter',
        'font_size':11,
        'font_name':'Calibri'
    })
    normal_account_format = workbook.add_format({
        'align':'left',
        'valign':'vcenter',
        'font_size':11,
        'font_name':'Calibri'
    })
    suspected_account_format = workbook.add_format({
        'align':'left',
        'valign':'vcenter',
        'font_size':11,
        'font_name':'Calibri',
        'bg_color':'#00B050',
    })
    violated_account_format = workbook.add_format({
        'align':'left',
        'valign':'vcenter',
        'font_size':11,
        'font_name':'Calibri',
        'font_color':'#FF0000',
        'bg_color': '#FFFF00'
    })
    number_format = workbook.add_format({
        'align':'right',
        'valign':'vcenter',
        'font_size':11,
        'font_name':'Calibri'
    })
    money_format = workbook.add_format({
        'align':'right',
        'valign':'vcenter',
        'font_size':11,
        'font_name':'Calibri',
        'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
    })
    header_1 = [
        'Mã loại hình',
        'Tên loại hình',
        'Số TK lưu ký',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Tên MG',
        'Tiền',
        'TL MR thực tế',
        'TL TC thực tế'
    ]
    header_2 = [
        'TL MR DN',
        'TL TC DN',
        'RLN0005',
        'RLN0006',
        'VMR9003',
        'Note',
        'DuNoBaoLanh RLN0006',
        'Tien RSA0004',
        'Thoi Gian RSA0004',
    ]
    worksheet = workbook.add_worksheet('Sheet1')
    # Set Columns & Rows
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',10)
    worksheet.set_column('C:D',13)
    worksheet.set_column('E:F',20)
    worksheet.set_column('G:R',14)

    def NaN2NA(worksheet,row,col,number,format_=None):
        if number != number: # np.NaN
            return worksheet.write(row,col,'#N/A',format_)
    worksheet.add_write_handler(float,NaN2NA)

    worksheet.write_row('A1',header_1,header_1_format)
    worksheet.write_row('J1',header_2,header_2_format)
    worksheet.write_column('A2',table['MaLoaiHinh'],text_left_format)
    worksheet.write_column('B2',table['TenLoaiHinh'],text_left_format)
    for rowNum,account in enumerate(table['SoTKLuuKy']):
        if table.loc[table.index[rowNum],'DuTinhGiaiNganT0'] != 0:
            if table.loc[table.index[rowNum],'TLMRThucTe'] < table.loc[table.index[rowNum],'TLMRDN']:
                if table.loc[table.index[rowNum],'TLMRThucTe'] >= 100:
                    fmt = suspected_account_format
                else:
                    fmt = violated_account_format
            else:
                if pd.isnull(table.loc[table.index[rowNum],'DuNoBaoLanh']): # không có giá trị
                    fmt = normal_account_format
                else: # có giá trị
                    fmt = violated_account_format
        else:
            if table.loc[table.index[rowNum],'TLMRThucTe'] < table.loc[table.index[rowNum],'TLMRDN']:
                if table.loc[table.index[rowNum],'TLMRThucTe'] >= 100:
                    fmt = suspected_account_format
                else:
                    if pd.isnull(table.loc[table.index[rowNum],'RLN0006']): # không có giá trị
                        fmt = violated_account_format
                    else:
                        if table.loc[table.index[rowNum],'TLTCThucTe'] < 100:
                            fmt = violated_account_format
            else:
                fmt = normal_account_format
        worksheet.write(f'C{rowNum+2}',account,fmt)
    worksheet.write_column('D2',table['SoTieuKhoan'],text_left_format)
    worksheet.write_column('E2',table['TenKhachHang'],text_left_format)
    worksheet.write_column('F2',table['TenMoiGioi'],text_left_format)
    worksheet.write_column('G2',table['Tien'],money_format)
    worksheet.write_column('H2',table['TLMRThucTe'],number_format)
    worksheet.write_column('I2',table['TLTCThucTe'],number_format)
    worksheet.write_column('J2',table['TLMRDN'],number_format)
    worksheet.write_column('K2',table['TLTCDN'],number_format)
    worksheet.write_column('L2',table['RLN0005'],money_format)
    worksheet.write_column('M2',table['RLN0006'],money_format)
    worksheet.write_column('N2',table['DuTinhGiaiNganT0'],money_format) # VMR9003
    worksheet.write_column('O2',['']*table.shape[0],text_left_format)
    worksheet.write_column('P2',table['DuNoBaoLanh'],money_format)
    worksheet.write_column('Q2',table['TienRSA0004'],money_format)
    worksheet.write_column('R2',table['TimeRSA0004'],money_format)

    ################################################################
    ################################################################
    ################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')



