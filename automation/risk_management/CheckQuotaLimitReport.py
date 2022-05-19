from automation.risk_management import *
from datawarehouse import *


def generateTempData():
    """
    Chạy lúc 16:10 chiều mỗi ngày, nếu ngày thường thì save file, nếu là ngày nghỉ thì không làm gì cả
    """

    info = get_info('daily',dt.datetime.now())
    dataDate = info['end_date']
    timeCheck = pd.read_sql(
        f"""
        SELECT [Work] FROM [Date] WHERE [Date].[Date] = '{dataDate}'
        """,
        connect_DWH_CoSo,
    ).squeeze()

    if not timeCheck: # Ngày nghỉ
        return # dừng hàm, không làm gì cả

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
        LEFT JOIN [VMR9003] [c] ON [i].[sub_account] = [c].[TieuKhoan]
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
    t0_date = info['end_date'].replace('.', '-')
    t1_date = BDATE(t0_date,-1)
    t2_date = BDATE(t0_date,-2)
    folder_name = info['folder_name']

    # create_folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ################################################################
    ################################################################
    ################################################################

    dataPart1 = pd.read_pickle(
        join(dirname(__file__),'temp',f'TempDataQuotaLimit_{t2_date.replace(".","")}.pickle')
    )
    dataPart2 = pd.read_sql(
        f"""
        WITH
        [VMR0001T2] AS (
            SELECT
                [sub_account].[account_code] [TaiKhoan],
                [VMR0001].[TLMRThucTe] [TLMRDN],
                [VMR0001].[TLTCThucTe] [TLTCDN]
            FROM [VMR0001]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [VMR0001].[TieuKhoan]
            WHERE [VMR0001].[Ngay] = '{t2_date}'
            AND [VMR0001].[TenLoaiHinh] = 'Margin'
        ),
        [RLN0005T1] AS (
            SELECT
                [sub_account].[account_code] [TaiKhoan],
                [RLN0005].[SoTienCapBaoLanh]
            FROM [RLN0005]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [RLN0005].[TieuKhoan]
            WHERE [RLN0005].[Ngay] = '{t1_date}'
        ),
        [RLN0006T1] AS (
            SELECT
                [margin_outstanding].[account_code] [TaiKhoan],
                [principal_outstanding]+[interest_outstanding]+[fee_outstanding] [SumOutstanding]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{t1_date}' AND [margin_outstanding].[type] = N'Trả chậm'
        ),
        [RSA0004T1] AS (
            SELECT
                [sub_account].[account_code] [TaiKhoan],
                [transactional_record].[amount] [Tien],
                [transactional_record].[time] [Time]
            FROM [transactional_record]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [transactional_record].[TieuKhoan]
            WHERE [transactional_record].[date] = '{t1_date}'
        )
        SELECT
            COALESCE([VMR0001T2].[TaiKhoan],[RLN0005T1].[TaiKhoan],[RLN0006T1].[TaiKhoan]) [TaiKhoan],
            [VMR0001T2].[TLMRDN],
            [VMR0001T2].[TLTCDN],
            [RLN0005T1].[SoTienCapBaoLanh] [RLN0005],
            [RLN0006T1].[SumOutstanding] [RLN0006],
            [RSA0004T1].[Tien] [Tien RSA0004]
            [RSA0004T!].[GioDuyet] [Time RSA0004]
        FROM [VMR0001T2]
        FULL JOIN [RLN0005T1] ON [RLN0005T1].[TaiKhoan] = [VMR0001T2].[TaiKhoan]
        FULL JOIN [RLN0006T1] ON [RLN0006T1].[TaiKhoan] = [VMR0001T2].[TaiKhoan]
        FULL JOIN [RSA0004T1] ON [RSA0004T1].[TaiKhoan] = [VMR0001T2].[TaiKhoan]
        """,
        connect_DWH_CoSo,
    ) # Thiếu cột RLN0006 BL
    table = dataPart1.join(dataPart2,how='left')
    table.sort_values('RLN0005',ascending=False,inplace=True)

    ################################################################
    ################################################################
    ################################################################

    file_name=f'Checking Quota {t0_day}{t0_month}{t0_year}.xlsx'
    writer=pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook=writer.book

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
        'font_color':'#FF0000',
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
        'Note'
    ]
    worksheet = workbook.add_worksheet('Sheet1')
    # Set Columns & Rows
    worksheet.set_column('A:B',9)
    worksheet.set_column('C:C',16)
    worksheet.set_column('D:F',8)
    worksheet.set_column('G:G',15)
    worksheet.set_column('H:K',9)
    worksheet.set_column('L:L',15)
    worksheet.set_column('M:M',9)
    worksheet.set_column('N:N',15)
    worksheet.set_column('O:O',9)

    worksheet.write_row('A1',header_1,header_1_format)
    worksheet.write_row('J1',header_2,header_2_format)
    worksheet.write_column('A2',table['MaLoaiHinh'],text_left_format)
    worksheet.write_column('B2',table['TenLoaiHinh'],text_left_format)

    for rowNum,account in enumerate(table['account_code']):
        if table.loc[table.index[rowNum],'VMR9003'] != 0:
            if table.loc[table.index[rowNum],'TLMRThucTe'] < table.loc[table.index[rowNum],'TLMRDN']:
                if table.loc[table.index[rowNum],'VMR9003'] >= 1:
                    fmt = suspected_account_format
                else:
                    fmt = violated_account_format
            else:
                if table.loc[table.index[rowNum],'RLN0006'] != table.loc[table.index[rowNum],'RLN0006']: # không có giá trị
                    fmt = normal_account_format
                else: # có giá trị
                    fmt = violated_account_format
        else:
            if table.loc[table.index[rowNum],'TLMRThucTe'] < table.loc[table.index[rowNum],'TLMRDN']:
                if table.loc[table.index[rowNum],'TLMRThucTe'] >= 1:
                    fmt = suspected_account_format
                else:
                    fmt = violated_account_format
            else:
                fmt = normal_account_format

        worksheet.write(f'C{rowNum+2}',account,fmt)
    
    worksheet.write_column('D2',table['sub_account'],text_left_format)
    worksheet.write_column('E2',table['customer_name'],text_left_format)
    worksheet.write_column('F2',table['broker_name'],text_left_format)
    worksheet.write_column('G2',table['Tien'],money_format)
    worksheet.write_column('H2',table['TLMRThucTe'],number_format)
    worksheet.write_column('I2',table['TLTCThucTe'],number_format)
    worksheet.write_column('J2',table['TLMRDN'],number_format)
    worksheet.write_column('K2',table['TLTCDN'],number_format)
    worksheet.write_column('L2',table['RLN0005'],money_format)
    worksheet.write_column('M2',table['RLN0006'],money_format)
    worksheet.write_column('N2',table['VMR9003'],money_format)
    worksheet.write_column('O2',['']*table.shape[0],text_left_format)

    worksheet.write_column('O2',['']*table.shape[0],text_left_format)
    worksheet.write_column('O2',['']*table.shape[0],text_left_format)

    ################################################################
    ################################################################
    ################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')



