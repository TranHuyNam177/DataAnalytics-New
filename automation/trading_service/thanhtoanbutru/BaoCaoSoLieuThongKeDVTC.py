"""
BC tháng, chạy vào làm việc đầu tiên tháng sau
Ko chạy lùi trước 23/12/2021 được (vì chưa bắt đầu lưu VCF0051)
"""

from automation.trading_service.thanhtoanbutru import *


# DONE
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('monthly',run_time)
    period = info['period']
    folder_name = info['folder_name']
    end_date = info['end_date']
    start_date = bdate(info['start_date'],-1)

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    outstandings = pd.read_sql(
        f"""
        WITH 
        [opening] AS (
            SELECT
                [margin_outstanding].[type],
                SUM([margin_outstanding].[principal_outstanding]) [o_outs]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{start_date}'
                AND [margin_outstanding].[account_code] NOT IN (
                    SELECT [account_code] FROM [account] WHERE [account_type] = N'Tự doanh')
            GROUP BY [margin_outstanding].[type]
        ),
        [closing] AS (
            SELECT 
                [margin_outstanding].[type],
                SUM([margin_outstanding].[principal_outstanding]) [c_outs]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{end_date}'
                AND [margin_outstanding].[account_code] NOT IN (
                    SELECT [account_code] FROM [account] WHERE [account_type] = N'Tự doanh')
            GROUP BY [margin_outstanding].[type]
        )
        SELECT
            [t].[type],
            [t].[o_outs],
            [t].[c_outs] - [t].[o_outs] [d_oust],
            [t].[c_outs]
        FROM (
            SELECT
                CASE COALESCE([closing].[type], [opening].[type])
                    WHEN N'Ứng trước cổ tức' THEN 'UTCT'
                    WHEN N'Trả chậm' THEN 'DP'
                    WHEN N'Margin' THEN 'MR'
                    WHEN N'Bảo lãnh' THEN 'BL'
                    ELSE [closing].[type]
                END [type],
                ISNULL([opening].[o_outs],0) [o_outs],
                ISNULL([closing].[c_outs],0) [c_outs]
            FROM [opening]
            FULL JOIN [closing]
                ON [opening].[type] = [closing].[type]) [t]
        ORDER BY 
            CASE [t].[type]
                WHEN N'UTCT' THEN 1
                WHEN N'DP' THEN 2
                WHEN N'MR' THEN 3
                WHEN N'BL' THEN 4
                WHEN N'Cầm cố' THEN 5
            END ASC;
        """,
        connect_DWH_CoSo,
    )

    c_margin_table = pd.read_sql(
        f"""
        SELECT DISTINCT
            [vcf0051].[sub_account], 
            -- [vcf0051].[contract_type],
            CASE
                WHEN
                    CHARINDEX('MR1',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('VIPCN',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'vipcnt1'
                WHEN
                    CHARINDEX('MR1',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('SILV',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'silvt1'
                WHEN
                    CHARINDEX('MR0',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'goldt0'
                WHEN
                    CHARINDEX('MR2',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'goldt2'
                ELSE 'NOR'
            END [dim_1],
            CASE
                WHEN
                    CHARINDEX('VIPCN',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP3',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'vipcndp3'
                WHEN
                    CHARINDEX('SILV',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP4',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'silvdp4'
                WHEN
                    CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP5',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'golddp5'
                WHEN
                    CHARINDEX('NOR',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP2',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'normaldp2'
                ELSE 'NODP'
            END [dim_2],
            CASE
                WHEN
                    CHARINDEX('SILV',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'silver'
                WHEN
                    CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'golden'
                WHEN
                    CHARINDEX('VIPCN',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'vipcn'
                ELSE 'NOR'
            END [dim_3]
            
        FROM 
            [vcf0051]
        
        WHERE
            [vcf0051].[date] = '{end_date}'
            AND[vcf0051].[contract_type] NOT LIKE N'%Thường%'
            AND [vcf0051].[contract_type] NOT LIKE N'%Tự doanh%'
            AND [vcf0051].[status] IN ('A','B')
            
        ORDER BY 
            [vcf0051].[sub_account]
            
        """,
        connect_DWH_CoSo,
    )

    # get last month's number of accounts
    o_year = start_date[:4]
    o_month = start_date[5:7]

    saved_folder = join(dirname(__file__),'soluongtaikhoankhachvip')
    saved_file = join(saved_folder,f'{o_year}.{o_month}.pickle')
    o_accounts = pd.read_pickle(saved_file)

    c_margin_accounts = c_margin_table.shape[0]
    o_margin_accounts = o_accounts['margin']

    # Dim 1
    count_table = c_margin_table['dim_1'].value_counts()
    c_dim_1 = count_table.reindex(['vipcnt1','silvt1','goldt0','goldt2']).fillna(0)
    o_dim_1 = o_accounts[['vipcnt1','silvt1','goldt0','goldt2']]
    d_dim_1 = c_dim_1-o_dim_1

    # Dim 2
    count_table = c_margin_table['dim_2'].value_counts()
    c_dim_2 = count_table.reindex(['normaldp2','vipcndp3','silvdp4','golddp5']).fillna(0)
    o_dim_2 = o_accounts[['normaldp2','vipcndp3','silvdp4','golddp5']]
    d_dim_2 = c_dim_2-o_dim_2

    # Dim 3
    count_table = c_margin_table['dim_3'].value_counts()
    c_dim_3 = count_table.reindex(['vipcn','silver','golden']).fillna(0)
    o_dim_3 = o_accounts[['vipcn','silver','golden']]
    d_dim_3 = c_dim_3-o_dim_3

    # export to pickle (for next month's reference)
    m_series = pd.Series(c_margin_accounts,index=['margin'],name='count')
    exported_series = pd.concat([m_series,c_dim_1,c_dim_2,c_dim_3])
    c_year = end_date[:4]
    c_month = end_date[5:7]
    exported_file = join(saved_folder,f'{c_year}.{c_month}.pickle')
    exported_series.to_pickle(exported_file)

    outstandings_inb01 = pd.read_sql(
        f"""
        SELECT
            CASE [margin_outstanding].[type]
                WHEN N'Ứng trước cổ tức' THEN 'UTCT'
                WHEN N'Trả chậm' THEN 'DP'
                WHEN N'Margin' THEN 'MR'
                WHEN N'Bảo lãnh' THEN 'BL'
                ELSE [margin_outstanding].[type]
            END [type],
            SUM([margin_outstanding].[principal_outstanding]) [c_outs]
        FROM [margin_outstanding]
        LEFT JOIN (
            SELECT DISTINCT
                [relationship].[account_code], 
                [relationship].[branch_id]
            FROM [relationship]
            WHERE [relationship].[date] = '{end_date}'
            ) [r]
            ON [r].[account_code] = [margin_outstanding].[account_code] 
        WHERE 
            [margin_outstanding].[date] = '{end_date}' 
            AND [r].[branch_id] = '0111'
            AND [margin_outstanding].[account_code] NOT IN (
                SELECT [account_code] FROM [account] WHERE [account_type] = N'Tự doanh')
        GROUP BY [type]
        ORDER BY 
            CASE [type]
                WHEN N'UTCT' THEN 1
                WHEN N'DP' THEN 2
                WHEN N'MR' THEN 3
                WHEN N'BL' THEN 4
                WHEN N'Cầm cố' THEN 5
            END ASC;
        """,
        connect_DWH_CoSo,
        index_col='type',
    ).squeeze(axis=1).reindex(['UTTB','DP','MR','BL','Cầm cố']).fillna(0)

    margin_accounts_inb01 = pd.read_sql(
        f"""
        SELECT 
            COUNT(DISTINCT [vcf0051].[sub_account]) [count]
        FROM 
            [vcf0051]
        WHERE
            [vcf0051].[date] = '{end_date}'
            AND [vcf0051].[contract_type] NOT LIKE N'%Thường%'
            AND [vcf0051].[contract_type] NOT LIKE N'%Tự doanh%'
            AND [vcf0051].[status] IN ('A','B')
            AND [vcf0051].[sub_account] IN (
                SELECT 
                    [relationship].[sub_account]
                FROM [relationship]
                WHERE [relationship].[date] = '{end_date}'
                    AND [relationship].[branch_id] = '0111'
        )
        """,
        connect_DWH_CoSo,
    ).squeeze()

    """
    chỗ này lấy file danh sách nợ xấu gần nhất của RMD -> cần được tối ưu khi
    làm ở RMD
    """

    folder = r"\\192.168.10.101\phs-storge-2018\RiskManagementDept" \
             r"\RMD_Data\Luu tru van ban\Monthly Report" \
             r"\2. Monthly Strategy Meeting Report"

    path = join(folder,end_date[:4])
    if not isdir(path):
        path = join(folder,f'{int(end_date[:4])-1}')
    else:
        files = [file for file in listdir(path) if 'Monthly' in file]
        if len(files)==0:
            path = join(folder,f'{int(end_date[:4])-1}')

    files = [file for file in listdir(path) if 'Monthly' in file]

    def get_month(x):
        pstring = x.split('.')[0].split('_')[-1][:3]
        month = month_mapper(pstring)
        return month

    months = []
    for file in files:
        try:
            months.append(int(get_month(file)))
        except (Exception,):  # một số thằng lưu tên không theo chuẩn
            months.append(0)
    file = files[months.index(max(months))]

    bad_loan_path = join(path,file)
    bad_loan = pd.read_excel(bad_loan_path,sheet_name='Bad loan',usecols=[0,1,4])
    bad_loan.columns = ['account_code','principal','group']
    last_row = (~bad_loan['account_code'].isna()).idxmin()
    bad_loan = bad_loan.iloc[:last_row]
    bad_loan.dropna(subset=['principal'],inplace=True)
    bad_loan = bad_loan.loc[bad_loan['principal']!=0]
    bad_loan['group'] = bad_loan['group'].str.replace(' ','').str.replace("case",'')
    bad_loan['group'].fillna(method='ffill',inplace=True)

    debt_type = pd.read_sql(
        f"""
        SELECT DISTINCT
        	[margin_outstanding].[account_code],
            CASE [margin_outstanding].[type]
                WHEN N'Ứng trước cổ tức' THEN 'UTCT'
                WHEN N'Trả chậm' THEN 'DP'
                WHEN N'Margin' THEN 'MR'
                WHEN N'Bảo lãnh' THEN 'BL'
            ELSE [margin_outstanding].[type]
            END [type]
        FROM [margin_outstanding]
        WHERE [margin_outstanding].[account_code] IN {iterable_to_sqlstring(bad_loan['account_code'])}
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    ).squeeze()

    bad_loan['type'] = bad_loan['account_code'].map(debt_type)
    name = pd.read_sql(
        f"""
        SELECT [account].[account_code], [account].[customer_name] 
        FROM [account]
        WHERE [account].[account_code] IN {iterable_to_sqlstring(bad_loan['account_code'])}
        """,
        connect_DWH_CoSo,
        index_col='account_code',
    ).squeeze()
    bad_loan['cutomer_name'] = bad_loan['account_code'].map(name)
    bad_loan.sort_values('type',inplace=True)

    bad_loan_g = bad_loan.groupby('type')['principal'].sum()
    outstandings = outstandings.join(bad_loan_g,'type','outer').fillna(0)
    outstandings.rename({'principal':'bad_loan'},axis=1,inplace=True)

    ###################################################
    ###################################################
    ###################################################

    file_name = f'Số liệu thống kê DVTC tháng {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':14,
            'font_name':'Times New Roman',
            'bg_color':'#4BABC6',
        }
    )
    percent_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'0.00%'
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    bad_loan_header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'bg_color':'#FFBF00'
        }
    )
    num_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
        }
    )
    num_bold_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
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
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    text_noborder_format = workbook.add_format(
        {
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    worksheet = workbook.add_worksheet('DVTC')
    worksheet.hide_gridlines(option=2)
    headers = [
        '',
        'Loại',
        'Đầu kỳ',
        'Trong kỳ',
        'Cuối kỳ',
        'Nợ xấu',
        'Ghi chú',
        'INB-01',
        'Tỷ lệ',
    ]
    worksheet.set_column('A:A',26)
    worksheet.set_column('B:B',16)
    worksheet.set_column('C:F',20)
    worksheet.set_column('G:G',27)
    worksheet.set_column('H:H',16)
    worksheet.set_column('I:I',10)

    worksheet.merge_range('A1:I1','TOÀN CÔNG TY',title_format)
    worksheet.write_row('A2',headers,header_format)
    worksheet.write_column('B3',outstandings['type'],text_center_format)
    worksheet.write_column('C3',outstandings['o_outs'],num_format)
    worksheet.write_column('D3',outstandings['d_oust'],num_format)
    worksheet.write_column('E3',outstandings['c_outs'],num_format)
    worksheet.write_column('F3',outstandings['bad_loan'],num_format)
    worksheet.write('F9',0,num_format)

    worksheet.merge_range('A3:A7','Tổng Dư nợ gốc',text_left_format)
    worksheet.write('A8','',text_left_format)
    worksheet.write('B8','TỔNG',header_format)
    worksheet.write_row('C8',outstandings.iloc[:,1:].sum(),num_bold_format)
    worksheet.write('A9','Dư nợ báo cáo SSC',text_left_format)
    worksheet.write_row('C9',[0,'=E9-C9',0],num_format)

    worksheet.write('A10','Tổng số lượng TK Margin\n(Active & Block)',text_left_format)
    worksheet.write('B10','',text_left_format)
    worksheet.write_row(
        'C10',
        [o_margin_accounts,c_margin_accounts-o_margin_accounts,c_margin_accounts],
        num_format,
    )
    worksheet.merge_range(
        'A11:A14',
        'Tổng số lượng TK VIP miễn lãi vay MR\n(Active &Block)',
        text_left_format
    )
    mapper = {
        'vipcnt1':'VIPCN T1',
        'silvt1':'SILV T1',
        'goldt0':'GOLD T0',
        'goldt2':'GOLD T2',
    }
    worksheet.write_column('B11',o_dim_1.index.map(mapper),text_center_format)
    worksheet.write_column('C11',o_dim_1,num_format)
    worksheet.write_column('D11',d_dim_1,num_format)
    worksheet.write_column('E11',c_dim_1,num_format)

    worksheet.merge_range(
        'A15:A18',
        'Tổng số lượng TK sử dụng hạn mức DP\n(Active &Block)',
        text_left_format
    )
    mapper = {
        'normaldp2':'NORMAL DP+2',
        'vipcndp3':'VIPCN DP+3',
        'silvdp4':'SILV DP+4',
        'golddp5':'GOLD DP+5',
    }
    worksheet.write_column('B15',o_dim_2.index.map(mapper),text_center_format)
    worksheet.write_column('C15',o_dim_2,num_format)
    worksheet.write_column('D15',d_dim_2,num_format)
    worksheet.write_column('E15',c_dim_2,num_format)

    worksheet.write_column(
        'A19',
        ['Tổng số lượng TK VIPCN','Tổng số lượng TK SILVER','Tổng số lượng TK GOLDEN'],
        text_left_format,
    )
    worksheet.write_column('B19',['']*3,text_center_format)
    worksheet.write_column('C19',o_dim_3,num_format)
    worksheet.write_column('D19',d_dim_3,num_format)
    worksheet.write_column('E19',c_dim_3,num_format)

    worksheet.write_column('F10',['']*12,num_format)
    for col in ['G','H','I']:
        worksheet.write_column(f'{col}3',['']*19,num_format)

    worksheet.write_column('H3',outstandings_inb01,num_format)
    worksheet.write('H8',outstandings_inb01.sum(),num_bold_format)
    worksheet.write('H10',margin_accounts_inb01,num_format)
    worksheet.write(f'I5',f'=H5/E5',percent_format)
    worksheet.write('I10','=H10/E10',percent_format)

    worksheet.merge_range('A23:E23','DANH SÁCH KH NỢ XẤU',bad_loan_header_format)
    headers = [
        'STT',
        'Số TKCK',
        'Dư nợ gốc',
        'Note',
        'Loại nợ',
    ]
    worksheet.write_row('A24',headers,header_format)
    worksheet.write_column('A25',np.arange(bad_loan.shape[0])+1,text_center_format)
    worksheet.write_column('B25',bad_loan['account_code'],text_center_format)
    worksheet.write_column('C25',bad_loan['principal'],num_format)
    worksheet.write_column('D25',bad_loan['group'],text_center_format)
    worksheet.write_column('E25',bad_loan['type'],text_center_format)
    worksheet.write_column('F25',bad_loan['cutomer_name'],text_noborder_format)

    last_row = 25+bad_loan.shape[0]
    worksheet.merge_range(f'A{last_row}:B{last_row}','Tổng',header_format)
    worksheet.write(f'C{last_row}',bad_loan['principal'].sum(),num_bold_format)
    worksheet.write_row(f'D{last_row}',['']*2,header_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
