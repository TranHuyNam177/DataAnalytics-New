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

    ###################################################
    ###################################################
    ###################################################

    generalRoomTable = pd.read_sql(
        f"""
        WITH
        [DanhMuc] AS (
            SELECT [MaCK]
            FROM [DanhMucChoVayMargin]
            WHERE [Ngay] = (SELECT MAX([Ngay]) FROM [DanhMucChoVayMargin])
        ),
        [ThongTinCK] AS (
            SELECT 
                [t].[Ticker] [MaCK],
                [t].[Value] [TenCK]
            FROM [DWH-ThiTruong].[dbo].[SecuritiesInfoVSD] [t]
            WHERE [t].[Attribute] = 'Issuers name'
        ),
        [ThiTruong] AS (
            SELECT
                [t].[Ticker] [MaCK],
                [t].[Close] * 1000 [ClosePrice]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay] [t]
            WHERE [t].[Date] = '{dataDate}'
        ),
        [RoomInfo] AS (
            SELECT
                [t].[ticker] [MaCK],
                SUM([t].[system_used_room]) [SystemUsedRoom],
                SUM([t].[used_special_room]) [UsedSpecialRoom],
                SUM([t].[system_used_room]) + SUM([t].[used_special_room]) [TotalUsedRoom]
            FROM [230007] [t]
            WHERE [t].[date] = '{dataDate}'
            GROUP BY [t].[date], [t].[ticker]
        ),
        [MarginPrice] AS (
            SELECT
                [t].[date] [Ngay],
                [t].[ticker_code] [MaCK],
                [t].[margin_max_price] [MaxPrice],
                [t].[margin_ratio] [MarginRatio]
            FROM [vpr0109] [t]
            WHERE [t].[room_code] LIKE N'CL101%' OR [t].[room_code] LIKE N'TC01%'
                AND [t].[date] = '{dataDate}'
        ),
        [TongHop] AS (
        SELECT
            [DanhMuc].[MaCK],
            [ThongTinCK].[TenCK],
            [RoomInfo].[SystemUsedRoom],
            [RoomInfo].[UsedSpecialRoom],
            [RoomInfo].[SystemUsedRoom] + [RoomInfo].[UsedSpecialRoom] [TotalUsedRoom],
            [ThiTruong].[ClosePrice],
            CASE
                WHEN [ThiTruong].[ClosePrice] < [MarginPrice].[MaxPrice]
                    THEN [ThiTruong].[ClosePrice]
                ELSE [MarginPrice].[MaxPrice]
            END [MinPrice],
            [MarginPrice].[MarginRatio],
            [MarginPrice].[MaxPrice]
        FROM [DanhMuc]
        LEFT JOIN [RoomInfo] ON [DanhMuc].[MaCK] = [RoomInfo].[MaCK]
        LEFT JOIN [MarginPrice] ON [DanhMuc].[MaCK] = [MarginPrice].[MaCK]
        LEFT JOIN [ThongTinCK] ON [DanhMuc].[MaCK] = [ThongTinCK].[MaCK]
        LEFT JOIN [ThiTruong] ON [DanhMuc].[MaCK] = [ThiTruong].[MaCK]
        )
        SELECT
            [TongHop].[MaCK],
            [TongHop].[TenCK],
            [TongHop].[SystemUsedRoom],
            [TongHop].[UsedSpecialRoom],
            [TongHop].[TotalUsedRoom],
            [TongHop].[ClosePrice],
            [Tonghop].[MinPrice],
            [TongHop].[MinPrice] * [TongHop].[MarginRatio] * [TongHop].[SystemUsedRoom] / 100 [GeneralRoomOutstanding],
            [TongHop].[MarginRatio],
            [TongHop].[MaxPrice]
        FROM [TongHop]
        """,
        connect_DWH_CoSo,
    )
    closePrice = generalRoomTable[['MaCK','ClosePrice']].set_index('MaCK').squeeze()
    def findBranch(roomName):
        regex = re.compile('\(.{2}\)|\(.{3}\)|\(.{4}\)')
        result = regex.search(roomName)
        if result is not None:
            return result.group().replace('(','').replace(')','')
    # VPR0108
    vpr0108 = pd.read_sql(
        f"""
        SELECT
            [t].[room_name] [RoomName],
            [t].[ticker] [Ticker],
            [t].[used_volume] [UsedRoom]
        FROM [vpr0108] [t]
        WHERE [t].[date] = '{dataDate}'
            AND [t].[used_volume] > 0
        """,
        connect_DWH_CoSo
    )
    vpr0108['Key'] = vpr0108['Ticker'] + '_' + vpr0108['RoomName'].map(findBranch)
    vpr0108 = vpr0108.dropna(subset=['Key'])
    vpr0108 = vpr0108.sort_values('Ticker').reset_index()
    vpr0108 = vpr0108[['Key','UsedRoom']]
    # VPR0109
    vpr0109 = pd.read_sql(
        f"""
        SELECT
            [t].[room_name] [RoomName],
            [t].[margin_ratio] [MarginRatio],
            [t].[ticker_code] [Ticker],
            [t].[margin_max_price] [MaxPrice]
        FROM [vpr0109] [t]
        WHERE [t].[date] = '{dataDate}'
            AND [t].[room_code] LIKE 'CL%'
            AND [t].[margin_ratio] > 0
        """,
        connect_DWH_CoSo
    )
    vpr0109['Key'] = None
    for row in range(vpr0109.shape[0]):
        ticker = vpr0109.loc[vpr0109.index[row],'Ticker']
        room_name = vpr0109.loc[vpr0109.index[row],'RoomName']
        if ticker in room_name:
            vpr0109.loc[vpr0109.index[row],'Key'] = ticker + '_' + findBranch(room_name)
    vpr0109 = vpr0109.dropna(subset=['Key'])
    vpr0109 = vpr0109.sort_values('Ticker').reset_index()
    vpr0109 = vpr0109[['Key','MarginRatio','MaxPrice']]
    vpr0109 = vpr0109.drop_duplicates()
    # Aggregate
    specialRoomTable = pd.merge(vpr0109,vpr0108,how='inner',on='Key')
    specialRoomTable = specialRoomTable.groupby('Key',as_index=False).agg({
        'MarginRatio': np.mean,
        'MaxPrice': np.mean,
        'UsedRoom': np.sum
    })
    specialRoomTable.insert(1,'MaCK',specialRoomTable['Key'].str.split('_').str.get(0))
    specialRoomTable.insert(4,'ClosePrice',specialRoomTable['MaCK'].map(closePrice))
    specialRoomTable['SpecialRoomOutstanding'] = specialRoomTable.apply(
        lambda x: min(x['MaxPrice'],x['ClosePrice'])*x['MarginRatio']*x['UsedRoom']/100,
        axis=1,
    )
    specialRoomTable = specialRoomTable.groupby('MaCK')[['UsedRoom','SpecialRoomOutstanding']].sum()

    finalTable = pd.merge(generalRoomTable,specialRoomTable,how='left',on='MaCK')
    finalTable = finalTable.fillna(0)
    finalTable['RemainingAsGeneralRoom'] = finalTable['UsedSpecialRoom'] - finalTable['UsedRoom']
    finalTable['RemainingAsGeneralOutstanding'] = finalTable['RemainingAsGeneralRoom'] * finalTable['MinPrice'] * finalTable['MarginRatio'] / 100
    finalTable['TotalOutstanding'] = finalTable['GeneralRoomOutstanding'] + finalTable['SpecialRoomOutstanding'] + finalTable['RemainingAsGeneralOutstanding']
    finalTable = finalTable[['MaCK','TenCK','SystemUsedRoom','UsedSpecialRoom','TotalUsedRoom','ClosePrice','MaxPrice','MarginRatio','TotalOutstanding']]
    finalTable = finalTable.sort_values('TotalOutstanding',ascending=False)
    finalTable = finalTable.reset_index(drop=True)

    ###################################################
    ###################################################
    ###################################################

    reportDay = reportDate[-2:]
    reportMonth = calendar.month_name[int(reportDate[5:7])]
    reportYear = reportDate[0:4]
    file_name = f'Top 30 biggest outstanding {reportMonth} {reportDay} {reportYear}.xlsx'
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
    title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 14,
            'font_name': 'Times New Roman',
        }
    )
    headers_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'text_wrap': True
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial'
        }
    )
    number_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial'
        }
    )
    rat_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )

    ###################################################
    ###################################################
    ###################################################

    title_sheet = 'TOP 30 BIGGEST OUTSTANDING'
    sub_date = dt.datetime.strptime(reportDate,'%Y-%m-%d').strftime('%d/%m/%Y')
    headers = [
        'No',
        'Code stock',
        'Name',
        'Used general room',
        'Used special room',
        'Total used room (previous session)',
        'Closed price (previous session)',
        'Max Price',
        'Ratio',
        'Total outstanding'
    ]
    worksheet = workbook.add_worksheet('Top 30')
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',12)
    worksheet.set_column('C:C',60)
    worksheet.set_column('D:I',16,options={'hidden':1})
    worksheet.set_column('J:J',23)
    for row in range(34,5000):
        worksheet.set_row(row,options={'hidden':True})

    worksheet.write('C1',title_sheet,title_format)
    worksheet.write('C2',sub_date,title_format)
    worksheet.write_row('A4',headers,headers_format)
    worksheet.write_column('A5',finalTable.index+1,number_format)
    worksheet.write_column('B5',finalTable['MaCK'],text_left_format)
    worksheet.write_column('C5',finalTable['TenCK'],text_left_format)
    worksheet.write_column('D5',finalTable['SystemUsedRoom'],money_format)
    worksheet.write_column('E5',finalTable['UsedSpecialRoom'],money_format)
    worksheet.write_column('F5',finalTable['TotalUsedRoom'],money_format)
    worksheet.write_column('G5',finalTable['ClosePrice'],money_format)
    worksheet.write_column('H5',finalTable['MaxPrice'],money_format)
    worksheet.write_column('I5',finalTable['MarginRatio'],rat_format)
    worksheet.write_column('J5',finalTable['TotalOutstanding'],money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')} ::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

