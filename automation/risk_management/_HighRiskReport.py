from automation.risk_management import *
from datawarehouse import *
from news_collector import scrape_ticker_by_exchange_VN_HN30

"""
I. Bảng gốc là bảng RMR0035 (high_risk_account)

II. Cột % giá hòa vốn & Max loan price = (Giá hòa vốn - (Max price * tỷ lệ/100))/(Max price * tỷ lệ/100) 

III. Công thức tính cột total loan: 
    total_loan = (total_outstanding - cash)/1000000

IV. CÁC TÀI KHOẢN BỊ LỆCH
    A. Sheet High Risk accounts
    1. Tài khoản 022C005417 - DXG, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 857,702,803 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 0
    --> cột Total loan bị tính sai (số đúng: 626, số sai: -231.77)
    2. Tài khoản 022C010788 - HCM, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 0 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 125,685,000
    --> cột Total loan bị tính sai (số đúng: 120 , số sai: 246.14)
    3. Tài khoản 022C018611 - HDC, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 92,523,315 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 0
    --> cột Total loan bị tính sai (số đúng: 229 , số sai: 136.02)
    4. Tài khoản 022C018636 - NVL, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 27,066,504,482 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 14,005,026,009 
    --> cột Total loan bị tính sai (số đúng: 13,089, số sai: 28)
    5. Tài khoản  022C026454 - POW, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 0 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 343,234,763 
    --> cột Total loan bị tính sai (số đúng: 2, số sai: 345.31)
    6. Tài khoản  022C029247 - HTP, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 3,696,328,907 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 1  
    --> cột Total loan bị tính sai (số đúng: 7,114 , số sai: 3417.44)
    7. Tài khoản  022C029985 - HSG, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 0 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 105,984,375  
    --> cột Total loan bị tính sai (số đúng: 159, số sai: 264.66)
    8. Tài khoản  022C993388 - C47, query trên SQL bảng high_risk_account (ngày 20/05/2022) cột cash trả ra
    giá trị 1,098,614,029 còn cột cash trong báo cáo bên QLRR (bảng RMR0035) có giá trị là 512,236,200
    --> cột Total loan bị tính sai (số đúng: 5,989, số sai: 5402.43)
    9. 
        - Các tài khoản 022C957139, 022C957140, 022C957141 - AGG cũng bị lệch cột cash trong bảng high_risk_account
    so với báo cáo RMR0035 bên QLRR như các trường hợp trên
        - Các tài khoản 022C102573, 022C280896, 022C887979 - DXG cũng bị lệch cột cash trong bảng high_risk_account
    so với báo cáo RMR0035 bên QLRR như các trường hợp trên
    10. Tài khoản 022C299666, 022C076967 - SHB, các giá trị đều khớp với báo cáo bên QLRR, nhưng bên QLRR, SHB thuộc 
    nhóm HN30 còn trên web SHB không nằm trong HN30

V. Sheet Liquidity Deal Report
    1. Rule
    - Rule gốc:
        + Approved Quantity: VPR0108 - cột D, điều kiện: lấy những mã có tổng số lượng phê duyệt > 0
        + Set up: không cần lấy cột này
    - Rule tạm áp dụng: Hiện tại, QLRR đang control nguồn 
    => Cột set up sẽ lấy theo báo cáo VPR0108 - cột D
    => Cột Approved Quantity: đọc file tuần trước
"""


def to_pickle(
    df: pd.DataFrame(),
    path: str
):
    return df.to_pickle(path)


def process_data(
    sqlString: str,
    t: str,
    df: pd.DataFrame
) -> pd.DataFrame():

    table = pd.read_sql(
        f"""
        WITH [d] AS (
            SELECT
                ROW_NUMBER() OVER(ORDER BY [sub].[Date] DESC) AS [num_row],
                [sub].[Date]
            FROM (
                SELECT DISTINCT
                    [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date]
                FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
                WHERE DATEDIFF(DAY,[DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date],'{t}') between 0 and 150
            ) [sub]
        ),
        [avg_3m] AS (
            SELECT
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker] [stock],
                AVG([DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Volume]) [total_volume]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
            WHERE [DuLieuGiaoDichNgay].[Date] >= (SELECT [d].[Date] FROM [d] WHERE [d].[num_row] = 66)
            AND [DuLieuGiaoDichNgay].[Date] <= '{t}'
            GROUP BY [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker]
        ),
        [thitruong] AS (
            SELECT
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker] [ticker],
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Volume] [match_trade_volume],
                ([DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Close] * 1000) [closed_price],
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[High]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
            WHERE [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date] = '{t}'
        ),
        [r] AS (
            SELECT
                [relationship].[sub_account],
                [relationship].[account_code],
                [relationship].[branch_id],
                CASE
                WHEN [relationship].[branch_id] = '0001' THEN N'Headquarter'
                WHEN [relationship].[branch_id] = '0101' THEN N'Dist.03'
                WHEN [relationship].[branch_id] = '0102' THEN N'PMH T.F'
                WHEN [relationship].[branch_id] = '0104' THEN N'Dist.07'
                WHEN [relationship].[branch_id] = '0105' THEN N'Tan Binh'
                WHEN [relationship].[branch_id] = '0111' THEN N'Institutional Business'
                WHEN [relationship].[branch_id] = '0113' THEN N'IB'
                WHEN [relationship].[branch_id] = '0117' THEN N'Dist.01'
                WHEN [relationship].[branch_id] = '0118' THEN N'AMD-03'
                WHEN [relationship].[branch_id] = '0119' THEN N'Institutional Business 02'
                WHEN [relationship].[branch_id] = '0201' THEN N'Ha Noi'
                WHEN [relationship].[branch_id] = '0202' THEN N'Thanh Xuan'
                WHEN [relationship].[branch_id] = '0203' THEN N'Cau Giay'
                WHEN [relationship].[branch_id] = '0301' THEN N'Hai Phong'
                ELSE
                    [branch].[branch_name]
            END [location]
            FROM [relationship]
            LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
            LEFT JOIN [vcf0051] ON [vcf0051].[sub_account] = [relationship].[sub_account]
            AND [vcf0051].[date] = [relationship].[date]
            WHERE [relationship].[date] = '{t}'
            AND [vcf0051].[contract_type] LIKE N'MR%'
        ),
        [rmr15] AS (
            SELECT
                [rmr0015].[sub_account],
                (SUM([rmr0015].[market_value])/1000000) [total_asset_val]
            FROM [rmr0015]
            WHERE [rmr0015].[date] = '{t}'
            GROUP BY [rmr0015].[sub_account]
        ),
        [rmr35_pivot] AS (
            SELECT
                [account_code],
                [stock],
                SUM([SCR]) [SCR_val],
                SUM([DL]) [DL_val]
            FROM (
                SELECT
                    [high_risk_account].[date],
                    [high_risk_account].[account_code],
                    [high_risk_account].[stock],
                    [high_risk_account].[type],
                    [high_risk_account].[value]
                FROM [high_risk_account]
                WHERE [high_risk_account].[date] = '{t}'
            ) [t1]
            PIVOT (
                MAX([t1].[value]) FOR [type] IN (SCR, DL)
            ) [t2]
            GROUP BY [date], [account_code], [stock]
        ),
        [vpr09] AS (
            SELECT
                [vpr0109_CL01].[ticker_code],
                [vpr0109_CL01].[margin_ratio] [mr_loan_ratio],
                [vpr0109_TC01].[dp_loan_ratio],
                [vpr0109_CL01].[margin_max_price] [max_loan_price]
            FROM [vpr0109] [vpr0109_CL01]
            JOIN (
                SELECT
                    [vpr0109].[date],
                    [vpr0109].[ticker_code],
                    [vpr0109].[margin_ratio] [dp_loan_ratio]
                FROM [vpr0109]
                WHERE [vpr0109].[room_code] = 'TC01_PHS'
            ) [vpr0109_TC01]
            ON [vpr0109_TC01].[ticker_code] = [vpr0109_CL01].[ticker_code]
            AND [vpr0109_TC01].[date] = [vpr0109_CL01].[date]
            WHERE [vpr0109_CL01].[date] = '{t}'
            AND [vpr0109_CL01].[room_code] = 'CL01_PHS'
        ),
        [237] AS (
            SELECT
                [230007].[ticker],
                [230007].[system_total_room] [general_room]
            FROM [230007]
            WHERE [230007].[date] = '{t}'
        ),
        [rmr35] AS (
            SELECT
                CONCAT([high_risk_account].[account_code],[high_risk_account].[stock]) [TK_stock],
                [high_risk_account].[account_code],
                [high_risk_account].[account_code] [depository_account],
                [high_risk_account].[stock],
                [high_risk_account].[quantity],
                [high_risk_account].[price],
                (([high_risk_account].[total_outstanding]-[high_risk_account].[cash])/1000000) [total_loan],
                ([high_risk_account].[market_value]/1000000) [margin_val]
            FROM [high_risk_account]
            WHERE [high_risk_account].[date] = '{t}'
        )
        SELECT DISTINCT
            CONCAT([r].[location],[rmr35].[depository_account],[rmr35].[stock]) [0],
            [rmr35].[TK_stock],
            [r].[branch_id],
            [r].[location],
            [rmr35].[depository_account],
            [rmr35].[stock],
            [rmr35].[quantity],
            [rmr35].[price],
            [rmr15].[total_asset_val],
            [rmr35].[total_loan],
            [rmr35].[margin_val],
            [rmr35_pivot].[SCR_val],
            [rmr35_pivot].[DL_val],
            [vpr09].[mr_loan_ratio],
            [vpr09].[dp_loan_ratio],
            [vpr09].[max_loan_price],
            [237].[general_room],
            CASE
                WHEN ([rmr15].[total_asset_val] - [rmr35].[margin_val]) > 0
                THEN ([rmr35].[total_loan]-([rmr15].[total_asset_val]-[rmr35].[margin_val]))*1000000/[rmr35].[quantity]
                ELSE ([rmr35].[total_loan])*1000000/[rmr35].[quantity]
            END [breakeven_price],
            [avg_3m].[total_volume] [avg_vol_3m],
            [thitruong].[match_trade_volume],
            [thitruong].[closed_price],
            [thitruong].[High] [price_changes]
        FROM [rmr35]
        LEFT JOIN [r] ON [r].[account_code] = [rmr35].[depository_account] 
        LEFT JOIN [rmr35_pivot] ON [rmr35_pivot].[account_code] = [rmr35].[depository_account]
        AND [rmr35_pivot].[stock] = [rmr35].[stock]
        LEFT JOIN [vpr09] ON [vpr09].[ticker_code] = [rmr35].[stock]
        LEFT JOIN [237] ON [237].[ticker] = [rmr35].[stock]
        LEFT JOIN [thitruong] ON [thitruong].[ticker] = [rmr35].[stock]
        LEFT JOIN [rmr15] ON [rmr15].[sub_account] = [r].[sub_account]
        LEFT JOIN [avg_3m] ON [avg_3m].[stock] = [rmr35].[stock]
        {sqlString}
        """,
        connect_DWH_CoSo,
        index_col='TK_stock'
    )
    table = table.merge(df['special_room'],how='left',left_index=True,right_index=True).fillna(0)

    table['minP_ML'] = table[['price','max_loan_price']].min(axis=1)

    table['total_potential_outstanding'] = \
        (((table['general_room'] + table['special_room']) * table['minP_ML'] * table[
            'mr_loan_ratio'] / 100) / 1000000000).round(0)

    # Tìm max(break-even price, 0)
    gia_hv = table['breakeven_price'].apply(lambda x:0 if x < 0 else x)
    # Tìm max(mr loan ratio, dp loan ratio)
    ty_le = table[['mr_loan_ratio','dp_loan_ratio']].max(axis=1)
    # ML -> Max loan price
    table['%giaHoaVon_ML'] = (gia_hv - table['minP_ML'] * ty_le / 100) / (table['max_loan_price'] * ty_le / 100)
    # MP -> market price
    table['%giaHoaVon_MP'] = (table['closed_price'] - table['breakeven_price']) / table['closed_price']
    table[''] = (table['match_trade_volume'] - table['avg_vol_3m']) / table['avg_vol_3m']
    table = table.replace([np.inf,-np.inf,np.nan],0)

    return table


def run(  # chạy hàng ngày
    run_time=dt.datetime.now()
):
    start = time.time()
    info = get_info('weekly',run_time)
    period = info['period']
    # t0_date = info['end_date']
    # date_last_week = BDATE(t0_date, -5)
    t0_date = '2022-05-20'
    date_last_week = '2022-05-13'
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    # process date string
    t0_day = t0_date[-2:]
    t0_month = t0_date[5:7]
    t0_year = t0_date[0:4]

    day_last_w = date_last_week[-2:]
    month_last_w = date_last_week[5:7]
    year_last_w = date_last_week[0:4]

    ###################################################
    ###################################################
    ###################################################

    # get ticker in VN30 and HN30
    ticker = scrape_ticker_by_exchange_VN_HN30.run()
    # read file excel special room
    file_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\RMD staff\Huy' \
                fr'\SR HIGH RISK\{t0_year}\{t0_month}\SR High Risk {t0_day}.{t0_month}.{t0_year}.xlsx'
    specialRoom = pd.read_excel(file_path)
    specialRoom.rename(columns={'Tổng số lượng':'special_room'},inplace=True)
    specialRoom = specialRoom.set_index('TK&Stock(1)')

    sqlStr = 'ORDER BY [depository_account],[stock]'
    table = process_data(sqlStr,f'{t0_date}',specialRoom)
    table = table.merge(ticker,how='left',left_on='stock',right_index=True)

    ###################################################
    ###################################################
    ###################################################

    file_name = f'Summary High Risk_{t0_day}{t0_month}{t0_year}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Sheet High Risk
    # Format
    headers0_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    headers1_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#B4C6E7'
        }
    )
    headers2_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'color':'#FF0000'
        }
    )
    headers3_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    headers4_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'color':'#FF0000',
            'bg_color':'#B4C6E7'
        }
    )
    headers5_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#C6E0B4'
        }
    )
    headers6_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#ED7D31'
        }
    )
    empty_column = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FF0000'
        }
    )
    red_column = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FF0000'
        }
    )
    text_left_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vbottom',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    text_left_bg_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#B4C6E7'
        }
    )
    text_left_arial_format = workbook.add_format(
        {
            'align':'left',
            'valign':'top',
            'font_size':10,
            'font_name':'Arial'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#B4C6E7'
        }
    )
    num_right_format = workbook.add_format(
        {
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Arial'
        }
    )
    price_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'#,##0_);(#,##0)',
            'bg_color':'#B4C6E7'
        }
    )
    SCR_DL_format = workbook.add_format(
        {
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Arial',
            'num_format':'#,##0.0000'
        }
    )
    ratio_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'bg_color':'#B4C6E7'
        }
    )
    market_price_format = workbook.add_format(
        {
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'bg_color':'#B4C6E7'
        }
    )
    giahoavonML_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
            'bg_color':'#B4C6E7'
        }
    )
    giahoavonMP_format = workbook.add_format(
        {
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
            'bg_color':'#FFF2CC'
        }
    )
    percent_format = workbook.add_format(
        {
            'align':'right',
            'valign':'vbottom',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0%',
        }
    )
    # WRITE EXCEL
    high_risk_sheet = workbook.add_worksheet('High risk')
    high_risk_sheet.set_column('B:B',13)
    high_risk_sheet.set_column('C:C',20)
    high_risk_sheet.set_column('D:D',13)
    high_risk_sheet.set_column('F:F',15)
    high_risk_sheet.set_column('H:H',11.5)
    high_risk_sheet.set_column('I:K',10)
    high_risk_sheet.set_column('L:L',15)
    high_risk_sheet.set_column('M:M',12.5)
    high_risk_sheet.set_column('N:N',10)
    high_risk_sheet.set_column('O:Q',9)
    high_risk_sheet.set_column('R:R',11)
    high_risk_sheet.set_column('S:S',13)
    high_risk_sheet.set_column('T:T',19)
    high_risk_sheet.set_column('U:U',12)
    high_risk_sheet.set_column('V:V',14)
    high_risk_sheet.set_column('W:W',13)
    high_risk_sheet.set_column('X:X',7.5)
    high_risk_sheet.set_column('Y:Y',13.5)
    high_risk_sheet.set_column('Z:Z',12)
    high_risk_sheet.set_column('AA:AA',2)
    high_risk_sheet.set_column('AB:AB',7.5)
    high_risk_sheet.set_column('AC:AC',9)
    high_risk_sheet.set_column('AD:AE',7)
    high_risk_sheet.set_row(1,105.5)

    high_risk_sheet.write('A2',0,headers0_format)
    high_risk_sheet.write_row('B2',['VN&HNX30','TK&Stock(1)'],headers1_format)
    high_risk_sheet.write('D2','RMR0035',headers2_format)
    high_risk_sheet.write_row('E2',['0','Depository Account','Stock (1)','Quantity','Price'],headers3_format)
    high_risk_sheet.write_row(
        'J2',
        [
            'Total Asset Value (Million dong)',
            'Total Loan (Total Outs - Total Cash) (Million dong)',
            'Margin value (Million dong)'
        ],
        headers1_format
    )
    high_risk_sheet.write_row('M2',['SCR Value','DL Value'],headers3_format)
    high_risk_sheet.write_row('O2',['MR Loan ratio (%)','DP Loan ratio (%)','Max loan price (dong)'],
                              headers4_format)
    high_risk_sheet.write_row(
        'R2',
        [
            'General room (approved)',
            'Special room (approved)',
            'Break-even price (dong)',
            'Total potential outstanding (Billion dong)'
        ],
        headers1_format
    )
    high_risk_sheet.write_row('V2',['Average Volume 3M','Total matched trading volume today'],headers3_format)
    high_risk_sheet.write('X2','% giá hòa vốn & Max loan price',headers5_format)
    high_risk_sheet.write_row('Y2',['Maket price','% giá hòa vốn & Maket price'],headers6_format)
    high_risk_sheet.write('AA2','',empty_column)
    high_risk_sheet.write('AB2','Price changes (%)',headers5_format)
    high_risk_sheet.write_row(
        'AC2',
        ['Total put through volume today','Market capitalization (billion dong)'],
        headers3_format
    )
    high_risk_sheet.write('AE2','',headers0_format)
    high_risk_sheet.write_column('A3',table['0'],text_left_format)
    high_risk_sheet.write_column('B3',[''] * table.shape[0],text_left_bg_format)
    high_risk_sheet.write_column('C3',table.index,text_center_format)
    high_risk_sheet.write_column('D3',table['branch_id'],text_left_arial_format)
    high_risk_sheet.write_column('E3',table['location'],text_center_format)
    high_risk_sheet.write_column('F3',table['depository_account'],text_left_arial_format)
    high_risk_sheet.write_column('G3',table['stock'],text_left_arial_format)
    high_risk_sheet.write_column('H3',table['quantity'],num_right_format)
    high_risk_sheet.write_column('I3',table['price'],price_format)
    high_risk_sheet.write_column('J3',table['total_asset_val'],price_format)
    high_risk_sheet.write_column('K3',table['total_loan'],price_format)
    high_risk_sheet.write_column('L3',table['margin_val'],price_format)
    high_risk_sheet.write_column('M3',table['SCR_val'],SCR_DL_format)
    high_risk_sheet.write_column('N3',table['DL_val'],SCR_DL_format)
    high_risk_sheet.write_column('O3',table['mr_loan_ratio'],ratio_format)
    high_risk_sheet.write_column('P3',table['dp_loan_ratio'],ratio_format)
    high_risk_sheet.write_column('Q3',table['max_loan_price'],price_format)
    high_risk_sheet.write_column('R3',table['general_room'],price_format)
    high_risk_sheet.write_column('S3',table['special_room'],price_format)
    high_risk_sheet.write_column('T3',table['breakeven_price'],price_format)
    high_risk_sheet.write_column('U3',table['total_potential_outstanding'],price_format)
    high_risk_sheet.write_column('V3',table['avg_vol_3m'],price_format)
    high_risk_sheet.write_column('W3',table['match_trade_volume'],ratio_format)
    high_risk_sheet.write_column('X3',table['%giaHoaVon_ML'],giahoavonML_format)
    high_risk_sheet.write_column('Y3',table['closed_price'],market_price_format)
    high_risk_sheet.write_column('Z3',table['%giaHoaVon_MP'],giahoavonMP_format)
    high_risk_sheet.write_column('AA3',[''] * table.shape[0],red_column)
    high_risk_sheet.write_column('AB3',table['price_changes'],giahoavonML_format)
    high_risk_sheet.write_column('AC3',[''] * table.shape[0],price_format)
    high_risk_sheet.write_column('AD3',[''] * table.shape[0],ratio_format)
    high_risk_sheet.write_column('AE3',table[''],percent_format)

    ###################################################
    ###################################################
    ###################################################

    # sheet Special Room
    specialRoom = specialRoom.fillna('')
    # Format
    header_color_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#B4C6E7'
        }
    )
    header_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Arial',
        }
    )
    header2_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri',
            'color':'#FF0000'
        }
    )
    header3_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Arial',
            'bg_color':'#FFC000'
        }
    )
    text_left_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vbottom',
            'font_size':10,
            'font_name':'Arial'
        }
    )
    money_format = workbook.add_format(
        {
            'font_size':10,
            'font_name':'Arial',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
        }
    )

    # WRITE EXCEL
    special_room_sheet = workbook.add_worksheet('Special Room')
    special_room_sheet.set_tab_color('yellow')
    special_room_sheet.set_column('A:A',26.5)
    special_room_sheet.set_column('B:B',11)
    special_room_sheet.set_column('C:C',8.5)
    special_room_sheet.set_column('D:D',12)
    special_room_sheet.set_column('E:E',14)
    special_room_sheet.set_column('F:F',17)
    special_room_sheet.set_column('G:G',18.5)
    special_room_sheet.set_column('H:H',17)
    special_room_sheet.set_column('I:I',9.5)
    special_room_sheet.set_column('J:J',10.5)
    special_room_sheet.set_column('K:K',8.5)
    special_room_sheet.set_column('L:L',14)

    headers = [
        'Code',
        'Mã Room',
        'Account',
        'Mã CK',
        'Tổng số lượng'
    ]
    special_room_sheet.write('A1','TK&Stock(1)',header_color_format)
    special_room_sheet.write_row('B1',headers,header_format)
    special_room_sheet.write('G1','Guarantee stock in normal a/c',header2_format)
    special_room_sheet.write_row('H1',['Note','g'],header_format)
    special_room_sheet.write('J1','Group/deal',header3_format)
    special_room_sheet.write_row('K1',['','Note','',''],header_format)

    special_room_sheet.write_column('A2',specialRoom.index,text_left_format)
    special_room_sheet.write_column('B2',specialRoom['Code'],header_color_format)
    special_room_sheet.write_column('C2',specialRoom['Mã Room'],text_left_format)
    special_room_sheet.write_column('D2',specialRoom['Account'],text_left_format)
    special_room_sheet.write_column('E2',specialRoom['Mã CK'],text_left_format)
    special_room_sheet.write_column('F2',specialRoom['special_room'],money_format)
    special_room_sheet.write_column('G2',specialRoom['Guarantee stock in normal a/c'],money_format)
    special_room_sheet.write_column('H2',specialRoom['Note'],money_format)
    special_room_sheet.write_column('I2',specialRoom['g'],money_format)
    special_room_sheet.write_column('J2',specialRoom['Group/deal'],text_left_format)
    special_room_sheet.write_column('K2',specialRoom['Unnamed: 10'],text_left_format)
    special_room_sheet.write_column('L2',specialRoom['Note.1'],text_left_format)
    special_room_sheet.write_column('M2',specialRoom['Unnamed: 12'],text_left_format)
    special_room_sheet.write_column('N2',specialRoom['Unnamed: 13'],text_left_format)

    ###################################################
    ###################################################
    ###################################################

    # sheet High Risk Accounts

    # loc trực tiếp trong SQL, nhưng do không có cột VN30&HN30 để check nên dùng pandas để loc
    # condition_loc = f"""
    #     WHERE
    #         ([rmr35].[total_loan] > 200 AND [rmr35_pivot].[SCR_val] > 90)
    #         OR ([rmr35].[total_loan] > 200 AND [rmr35_pivot].[SCR_val] > 70 AND [rmr35_pivot].[DL_val] > 0.5)
    #         OR ([rmr35].[total_loan] > 200 AND [rmr35_pivot].[SCR_val] > 20 AND [rmr35_pivot].[DL_val] > 1.5)
    # """
    # final_table = process_data(condition_loc, t0_date, specialRoom)
    # final_table = final_table.loc[~(final_table['breakeven_price'] < 0)]

    table_loc_1 = table.loc[(table['total_loan'] > 200) & (table['SCR_val'] > 90)]
    table_loc_2 = table.loc[(table['total_loan'] > 200) & (table['SCR_val'] > 70) & (table['DL_val'] > 0.5)]
    table_loc_3 = table.loc[(table['total_loan'] > 200) & (table['exchange'].isnull()) & (table['SCR_val'] > 70)]
    table_loc_4 = table.loc[
        (table['total_loan'] > 200) &
        (table['exchange'].isnull()) &
        (table['SCR_val'] > 20) &
        (table['DL_val'] > 1.5)
        ]
    high_risk_account = pd.concat(
        [table_loc_1,table_loc_2,table_loc_3,table_loc_4]
    ).drop_duplicates().reset_index(drop=True)
    high_risk_account = high_risk_account.loc[~(high_risk_account['breakeven_price'] < 0)]

    # save today result to a pickle file
    path = join(dirname(__file__),'pickle_file','HighRisk_last_week',f'highrisk_{t0_day}{t0_month}{t0_year}.pickle')
    to_pickle(high_risk_account,path)

    # read file pickle high risk last week
    hr_last_w = pd.read_pickle(
        join(dirname(__file__),'pickle_file','HighRisk_last_week',
             f'highrisk_{day_last_w}{month_last_w}{year_last_w}.pickle')
    )
    high_risk_account.loc[high_risk_account['depository_account'].isin(hr_last_w['depository_account']),'stt'] = 'old'
    high_risk_account['stt'] = high_risk_account['stt'].fillna('new')
    high_risk_account = high_risk_account.sort_values(['stt','special_room','stock'],ascending=[True,False,True],
                                                      ignore_index=True)

    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    new_case_format = workbook.add_format(
        {
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#00B050'
        }
    )
    old_case_format = workbook.add_format(
        {
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FFC000'
        }
    )
    new_account_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#00B050'
        }
    )
    old_account_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FFC000'
        }
    )
    headers1_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    headers2_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'color':'#FF0000'
        }
    )
    headers3_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#ED7D31'
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
    num_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'#,##0'
        }
    )
    dl_val_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'#,##0.00'
        }
    )
    breakeven_price_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    percent_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0%'
        }
    )
    percent_breakeven_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0.00%'
        }
    )

    header_1_hr = [
        'No.',
        '',
        'Location',
        'Depository Account',
        'Stock (1)',
        'Quantity',
        'Price',
        'Total Asset Value (Million dong)'
    ]
    header_2_hr = [
        'Margin value (Million dong)',
        'SCR Value',
        'DL Value',
        'MR Loan ratio (%)',
        'DP Loan ratio (%)',
        'Max loan price (dong)',
        'General room (approved)',
        'Special room (approved)'
    ]
    header_3_hr = [
        'Total potential outstanding (Billion dong)',
        'Average Volume 3M',
        'Total matched trading volume today',
        '% Break-even price & Max loan price'
    ]

    # WRITE EXCEL
    high_risk_account_sheet = workbook.add_worksheet('High Risk accounts')
    high_risk_account_sheet.set_column('A:A',5)
    high_risk_account_sheet.set_column('B:B',0)
    high_risk_account_sheet.set_column('C:C',15.5)
    high_risk_account_sheet.set_column('D:D',14)
    high_risk_account_sheet.set_column('E:E',9)
    high_risk_account_sheet.set_column('F:F',11.5)
    high_risk_account_sheet.set_column('G:G',10.5)
    high_risk_account_sheet.set_column('H:H',11)
    high_risk_account_sheet.set_column('I:I',14)
    high_risk_account_sheet.set_column('J:J',11)
    high_risk_account_sheet.set_column('K:K',10)
    high_risk_account_sheet.set_column('L:L',9)
    high_risk_account_sheet.set_column('M:N',10)
    high_risk_account_sheet.set_column('O:O',12)
    high_risk_account_sheet.set_column('P:Q',13.5)
    high_risk_account_sheet.set_column('R:R',12)
    high_risk_account_sheet.set_column('S:S',13)
    high_risk_account_sheet.set_column('T:T',12)
    high_risk_account_sheet.set_column('U:U',14)
    high_risk_account_sheet.set_column('V:V',0)
    high_risk_account_sheet.set_column('W:X',12)
    high_risk_account_sheet.set_column('Y:Y',22.5)
    high_risk_account_sheet.set_column('Z:Z',24)
    high_risk_account_sheet.set_row(4,73.5)

    high_risk_account_sheet.write('A1','HIGH RISK ACCOUNTS',title_format)
    t0_month_title = calendar.month_name[int(t0_month)]
    t0_year_title = t0_year[2:4]
    high_risk_account_sheet.write('C2',f'{t0_day}-{t0_month_title}-{t0_year_title}',text_center_format)
    high_risk_account_sheet.write('C3','',old_case_format)
    high_risk_account_sheet.write('D3','Old cases',text_center_format)
    high_risk_account_sheet.write('C4','',new_case_format)
    high_risk_account_sheet.write('D4','New cases',text_center_format)
    high_risk_account_sheet.write_row('A5',header_1_hr,headers1_format)
    high_risk_account_sheet.write_rich_string(
        'I5',headers2_format,'Total Loan',headers1_format,' (Total Outs - Total Cash) (Million dong)',
        headers1_format
    )
    high_risk_account_sheet.write_row('J5',header_2_hr,headers1_format)
    high_risk_account_sheet.write('R5','Break-even price (dong)',headers2_format)
    high_risk_account_sheet.write_row('S5',header_3_hr,headers1_format)
    high_risk_account_sheet.write_row('W5',['Reference price','% Reference price & Break-even price'],
                                      headers3_format)
    high_risk_account_sheet.write_row('Y5',['Industry','Note'],headers1_format)
    high_risk_account_sheet.write_column('A6',np.arange(high_risk_account.shape[0]) + 1,text_center_format)
    high_risk_account_sheet.write_column('B6',[''] * high_risk_account.shape[0],text_center_format)
    high_risk_account_sheet.write_column('C6',high_risk_account['location'],text_center_format)

    for idx,val in high_risk_account.iterrows():
        if val['stt'] == 'old':
            fmt = old_account_format
        else:
            fmt = new_account_format
        high_risk_account_sheet.write(f'D{idx + 6}',val['depository_account'],fmt)

    high_risk_account_sheet.write_column('E6',high_risk_account['stock'],text_center_format)
    high_risk_account_sheet.write_column('F6',high_risk_account['quantity'],num_center_format)
    high_risk_account_sheet.write_column('G6',high_risk_account['price'],num_center_format)
    high_risk_account_sheet.write_column('H6',high_risk_account['total_asset_val'],num_center_format)
    high_risk_account_sheet.write_column('I6',high_risk_account['total_loan'],num_center_format)
    high_risk_account_sheet.write_column('J6',high_risk_account['margin_val'],num_center_format)
    high_risk_account_sheet.write_column('K6',high_risk_account['SCR_val'],num_center_format)
    high_risk_account_sheet.write_column('L6',high_risk_account['DL_val'],dl_val_format)
    high_risk_account_sheet.write_column('M6',high_risk_account['mr_loan_ratio'],num_center_format)
    high_risk_account_sheet.write_column('N6',high_risk_account['dp_loan_ratio'],num_center_format)
    high_risk_account_sheet.write_column('O6',high_risk_account['max_loan_price'],num_center_format)
    high_risk_account_sheet.write_column('P6',high_risk_account['general_room'],num_center_format)
    high_risk_account_sheet.write_column('Q6',high_risk_account['special_room'],num_center_format)
    high_risk_account_sheet.write_column('R6',high_risk_account['breakeven_price'],breakeven_price_format)
    high_risk_account_sheet.write_column('S6',high_risk_account['total_potential_outstanding'],num_center_format)
    high_risk_account_sheet.write_column('T6',high_risk_account['avg_vol_3m'],num_center_format)
    high_risk_account_sheet.write_column('U6',high_risk_account['match_trade_volume'],num_center_format)
    high_risk_account_sheet.write_column('V6',high_risk_account['%giaHoaVon_ML'],percent_breakeven_format)
    high_risk_account_sheet.write_column('W6',high_risk_account['closed_price'],breakeven_price_format)
    high_risk_account_sheet.write_column('X6',high_risk_account['%giaHoaVon_MP'],percent_format)
    high_risk_account_sheet.write_column('Y6',[''] * high_risk_account.shape[0],text_center_format)
    high_risk_account_sheet.write_column('Z6',[''] * high_risk_account.shape[0],text_center_format)

    ###################################################
    ###################################################
    ###################################################

    # sheet Liquidity Deal Report
    # read file excel Summary High Risk last week - Sheet Liquidity Deal Report
    file_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\Daily Report' \
                fr'\04. High Risk & Liquidity\High Risk\{year_last_w}\{month_last_w}.{year_last_w}' \
                fr'\Summary High Risk_{day_last_w}{month_last_w}{year_last_w}.xlsx'
    liquidity_deal_last_w = pd.read_excel(
        file_path,
        sheet_name='Liquidity Deal Report',
        skiprows=1,
        dtype={'Mã Room':object}
    ).iloc[:,1:-1]
    liquidity_deal_last_w = liquidity_deal_last_w.dropna(how='all',
                                                         subset=['Code','Setup date','Approved Date','Location'])
    liquidity_deal_last_w = liquidity_deal_last_w.reset_index(drop=True)

    liquidity_deal_last_w = liquidity_deal_last_w.fillna(method='ffill').set_index('Code')

    query_sql = pd.read_sql(
        f"""
        WITH 
        [d] AS (
            SELECT
                ROW_NUMBER() OVER(ORDER BY [sub].[Date] DESC) AS [num_row],
                [sub].[Date]
            FROM (
                SELECT DISTINCT
                    [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date]
                FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
                WHERE DATEDIFF(DAY,[DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date],'{t0_date}') between 0 and 150
            ) [sub]
        ),
        [avg_3m] AS (
            SELECT
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker] [stock],
                AVG([DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Volume]) [total_volume]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
            WHERE [DuLieuGiaoDichNgay].[Date] >= (SELECT [d].[Date] FROM [d] WHERE [d].[num_row] = 66)
            AND [DuLieuGiaoDichNgay].[Date] <= '{t0_date}'
            GROUP BY [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker]
        ),
        [vpr08] AS (
            SELECT
                [vpr0108].[room_code],
                CONCAT([vpr0108].[room_code],[vpr0108].[ticker]) [code],
                [vpr0108].[ticker],
                [vpr0108].[total_volume] [set_up],
                [vpr0108].[used_volume] [used_quantity]
            FROM [vpr0108]
            WHERE [vpr0108].[date] = '{t0_date}'
        ),
        [vpr09] AS (
            SELECT
                [vpr0109_CL01].[ticker_code],
                [vpr0109_CL01].[margin_ratio] [mr_approved_ratio],
                [vpr0109_TC01].[dp_approved_ratio],
                [vpr0109_CL01].[margin_max_price] [maximum_loan_price]
            FROM [vpr0109] [vpr0109_CL01]
            JOIN (
                SELECT
                    [vpr0109].[date],
                    [vpr0109].[ticker_code],
                    [vpr0109].[margin_ratio] [dp_approved_ratio]
                FROM [vpr0109]
                WHERE [vpr0109].[room_code] = 'TC01_PHS'
            ) [vpr0109_TC01]
            ON [vpr0109_TC01].[ticker_code] = [vpr0109_CL01].[ticker_code]
            AND [vpr0109_TC01].[date] = [vpr0109_CL01].[date]
            WHERE [vpr0109_CL01].[date] = '{t0_date}'
            AND [vpr0109_CL01].[room_code] = 'CL01_PHS'
        ),
        [thitruong] AS (
            SELECT
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker] [ticker],
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Volume] [match_trade_volume]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
            WHERE [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date] = '{t0_date}'
        )
        SELECT
            [vpr08].[room_code],
            [vpr08].[code],
            [vpr08].[ticker],
            [vpr08].[set_up],
            [vpr09].[mr_approved_ratio],
            [vpr09].[dp_approved_ratio],
            [vpr09].[maximum_loan_price],
            [vpr08].[used_quantity],
            [avg_3m].[total_volume] [avg_liquidity_vol_3m],
            [thitruong].[match_trade_volume] [liquidity_today]
        FROM [vpr08]
        FULL JOIN [vpr09] ON [vpr09].[ticker_code] = [vpr08].[ticker]
        FULL JOIN [thitruong] ON [thitruong].[ticker] = [vpr08].[ticker]
        FULL JOIN [avg_3m] ON [avg_3m].[stock] = [vpr08].[ticker]
        WHERE [vpr08].[ticker] IS NOT NULL
        AND (
            [vpr09].[mr_approved_ratio] IS NOT NULL
            OR [vpr09].[dp_approved_ratio] IS NOT NULL
        )
        """,
        connect_DWH_CoSo,
        index_col='code'
    )
    liquidity_table = liquidity_deal_last_w.merge(query_sql,how='outer',left_index=True,right_index=True)
    liquidity_table = liquidity_table.loc[
        ~(liquidity_table['Approved Quantity'].isnull()) | (liquidity_table['set_up'] != 0)]
    liquidity_table = liquidity_table.rename_axis('code').reset_index()
    liquidity_table = liquidity_table.sort_values(['room_code','Location'],ascending=[True,False],ignore_index=True)
    for colNum,colName in enumerate(liquidity_table.columns):
        if pd.api.types.is_numeric_dtype(liquidity_table[colName]):
            liquidity_table[colName] = liquidity_table[colName].fillna(0)
        else:
            liquidity_table[colName] = liquidity_table[colName].fillna('')

    ratio = liquidity_table[['mr_approved_ratio','dp_approved_ratio']].max(axis=1)
    liquidity_table['pOuts'] = liquidity_table['Approved Quantity'] * liquidity_table[
        'maximum_loan_price'] * ratio / 100
    liquidity_table['approve/avg_liquidity_vol_3m'] = liquidity_table['Approved Quantity'] / liquidity_table[
        'avg_liquidity_vol_3m']
    liquidity_table = liquidity_table.replace([np.inf,-np.inf],0)
    # loc 3 dòng cuối cùng
    loc_last_3rows = liquidity_table.iloc[-3:].fillna('')
    loc_last_3rows.drop(loc_last_3rows.loc[:,'room_code':],inplace=True,axis=1)
    # liquidity table không có 3 dòng cuối
    liquidity_table = liquidity_table.iloc[:-3]
    liquidity_table = liquidity_table.drop(
        liquidity_table.loc[:,'Set up':'Approval room/3months average liquidity volume'],axis=1
    )

    # tìm min, max index để merge range giá trị của cột Account trong excel
    min_idx = liquidity_table.groupby(['Account','room_code']).apply(lambda x:min(tuple(x.index))).rename('min_idx')
    max_idx = liquidity_table.groupby(['Account','room_code']).apply(lambda x:max(tuple(x.index))).rename('max_idx')
    min_max_idx = pd.merge(min_idx,max_idx,left_index=True,right_index=True)
    liquidity_table = liquidity_table.merge(min_max_idx,how='left',left_on=['Account','room_code'],right_index=True)

    # WRITE EXCEL
    # Format
    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    headers1_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'bg_color':'#70AD47'
        }
    )
    headers2_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'bg_color':'#70AD47'
        }
    )
    headers3_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'bg_color':'#FFC000'
        }
    )
    headers4_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'color':'#FF0000'
        }
    )
    text_center_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    empty_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'bg_color':'#FFFF00'
        }
    )
    text_left_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    date_center_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'dd/mm/yyyy'
        }
    )
    date_empty_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'dd/mm/yyyy',
            'bg_color':'#FFFF00'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    sum_money_format = workbook.add_format(
        {
            'bold':True,
            'align':'right',
            'valign':'vbottom',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    percent_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        }
    )
    header1_liquidity = [
        'No.',
        'Mã Room',
        'Code',
        'Setup date'
    ]
    header2_liquidity = [
        'Approved Date',
        'Stock',
        'Location',
        'Account',
        'Approved Quantity',
        'Set up',
        'MR Approved Ratio (%)',
        'DP Approved Ratio (%)',
        'Maximum loan price'
    ]
    header3_liquidity = [
        'Used quantity today',
        'Average Liquidity Volume of 03 months',
        'Liquidity today'
    ]
    header4_liquidity = [
        'Approval room/3months average liquidity volume',
        'Commitment Fix Type other special'
    ]
    liquidity_deal_sheet = workbook.add_worksheet('Liquidity Deal Report')
    liquidity_deal_sheet.set_column('A:A',5)
    liquidity_deal_sheet.set_column('B:D',0)
    liquidity_deal_sheet.set_column('E:E',13)
    liquidity_deal_sheet.set_column('E:G',13)
    liquidity_deal_sheet.set_column('H:H',35)
    liquidity_deal_sheet.set_column('I:J',18.5)
    liquidity_deal_sheet.set_column('K:K',16.5)
    liquidity_deal_sheet.set_column('L:M',14.5)
    liquidity_deal_sheet.set_column('N:N',27)
    liquidity_deal_sheet.set_column('O:P',17)
    liquidity_deal_sheet.set_column('Q:Q',20)
    liquidity_deal_sheet.set_column('R:R',15)
    liquidity_deal_sheet.set_column('S:S',20)
    liquidity_deal_sheet.set_row(1,71)

    liquidity_deal_sheet.write('E1',"DAILY  DEAL's LIQUIDITY REPORT",title_format)
    liquidity_deal_sheet.write_row('A2',header1_liquidity,headers1_format)
    liquidity_deal_sheet.write_row('E2',header2_liquidity,headers2_format)
    liquidity_deal_sheet.write('N2','P.Outs',headers3_format)
    liquidity_deal_sheet.write_row('O2',header3_liquidity,headers4_format)
    liquidity_deal_sheet.write_row('R2',header4_liquidity,headers1_format)
    liquidity_deal_sheet.write_column('A3',liquidity_table.index + 1,text_center_format)
    liquidity_deal_sheet.write_column('B3',liquidity_table['room_code'],text_center_format)
    liquidity_deal_sheet.write_column('C3',liquidity_table['code'],text_center_format)
    liquidity_deal_sheet.write_column('F3',liquidity_table['ticker'],text_center_format)
    for idx,val in liquidity_table.iterrows():
        if val['Setup date'] == '' or val['Approved Date'] == '':
            fmt = date_empty_format
        elif val['Location'] == '':
            fmt = empty_format
        else:
            fmt = date_center_format
        liquidity_deal_sheet.write(f'D{idx + 3}',val['Setup date'],fmt)
        liquidity_deal_sheet.write(f'E{idx + 3}',val['Approved Date'],fmt)
        liquidity_deal_sheet.write(f'G{idx + 3}',val['Location'],fmt)
    for idx,val in liquidity_table.iterrows():
        if val['min_idx'] == val['max_idx'] and val['Account'] == '':
            liquidity_deal_sheet.write(f"H{val['min_idx'] + 3}",val['Account'],empty_format)
        elif val['min_idx'] == val['max_idx']:
            liquidity_deal_sheet.write(f"H{val['min_idx'] + 3}",val['Account'],text_center_format)
        elif val['min_idx'] != val['max_idx'] and val['Account'] == '':
            liquidity_deal_sheet.merge_range(f"H{val['min_idx'] + 3}:H{val['max_idx'] + 3}",val['Account'],empty_format)
        else:
            liquidity_deal_sheet.merge_range(
                f"H{val['min_idx'] + 3}:H{val['max_idx'] + 3}",
                val['Account'],
                text_center_format
            )

    liquidity_deal_sheet.write_column('I3',liquidity_table['Approved Quantity'],money_format)
    liquidity_deal_sheet.write_column('J3',liquidity_table['set_up'],money_format)
    liquidity_deal_sheet.write_column('K3',liquidity_table['mr_approved_ratio'],text_center_format)
    liquidity_deal_sheet.write_column('L3',liquidity_table['dp_approved_ratio'],text_center_format)
    liquidity_deal_sheet.write_column('M3',liquidity_table['maximum_loan_price'],money_format)
    liquidity_deal_sheet.write_column('N3',liquidity_table['pOuts'],money_format)
    liquidity_deal_sheet.write_column('O3',liquidity_table['used_quantity'],money_format)
    liquidity_deal_sheet.write_column('P3',liquidity_table['avg_liquidity_vol_3m'],money_format)
    liquidity_deal_sheet.write_column('Q3',liquidity_table['liquidity_today'],money_format)
    liquidity_deal_sheet.write_column('R3',liquidity_table['approve/avg_liquidity_vol_3m'],percent_format)
    liquidity_deal_sheet.write_column('S3',[''] * liquidity_table.shape[0],text_left_format)

    sum_start_row = liquidity_table.shape[0] + 3
    for col in 'IN':
        liquidity_deal_sheet.write(
            f'{col}{sum_start_row}',
            f'=SUBTOTAL(9,{col}3:{col}{sum_start_row - 1})',
            sum_money_format
        )

    last_three_row = sum_start_row + 2
    liquidity_deal_sheet.write_column(f'A{last_three_row}',loc_last_3rows.index + 1,text_center_format)
    liquidity_deal_sheet.write_column(f'B{last_three_row}',loc_last_3rows['Mã Room'],text_center_format)
    liquidity_deal_sheet.write_column(f'C{last_three_row}',loc_last_3rows['code'],text_center_format)
    liquidity_deal_sheet.write_column(f'D{last_three_row}',loc_last_3rows['Setup date'],date_center_format)
    liquidity_deal_sheet.write_column(f'E{last_three_row}',loc_last_3rows['Approved Date'],date_center_format)
    liquidity_deal_sheet.write_column(f'F{last_three_row}',loc_last_3rows['Stock '],text_center_format)
    liquidity_deal_sheet.write_column(f'G{last_three_row}',loc_last_3rows['Location'],date_center_format)
    liquidity_deal_sheet.write_column(f'H{last_three_row}',loc_last_3rows['Account'],text_center_format)
    liquidity_deal_sheet.write_column(f'I{last_three_row}',loc_last_3rows['Approved Quantity'],money_format)
    liquidity_deal_sheet.write_column(f'J{last_three_row}',loc_last_3rows['Set up'],money_format)
    liquidity_deal_sheet.write_column(
        f'K{last_three_row}',
        loc_last_3rows[' MR Approved Ratio (%)'],
        text_center_format
    )
    liquidity_deal_sheet.write_column(
        f'L{last_three_row}',
        loc_last_3rows[' DP Approved Ratio (%)'],
        text_center_format
    )
    liquidity_deal_sheet.write_column(f'M{last_three_row}',loc_last_3rows['Maximum loan price'],money_format)
    liquidity_deal_sheet.write_column(f'N{last_three_row}',loc_last_3rows['P.Outs'],money_format)
    liquidity_deal_sheet.write_column(f'O{last_three_row}',loc_last_3rows['Used quantity today'],money_format)
    liquidity_deal_sheet.write_column(
        f'P{last_three_row}',
        loc_last_3rows['Average Liquidity Volume of 03 months\n'],
        money_format
    )
    liquidity_deal_sheet.write_column(f'Q{last_three_row}',loc_last_3rows['Liquidity today\n'],money_format)
    liquidity_deal_sheet.write_column(
        f'R{last_three_row}',
        loc_last_3rows['Approval room/3months average liquidity volume'],
        percent_format
    )
    liquidity_deal_sheet.write_column(f'S{last_three_row}',['']*loc_last_3rows.shape[0],text_left_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start,1)}s')
