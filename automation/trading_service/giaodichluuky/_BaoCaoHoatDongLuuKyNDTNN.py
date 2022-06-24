from automation.trading_service.giaodichluuky import *
from request.stock import ta
from news_collector import scrape_ticker_by_exchange
from news_collector import scrape_derivatives_price
from datawarehouse.DWH_CoSo.de083 import monthly_update

"""
BC này không chạy lùi ngày được (do scrape bảng điện thời điểm hiện tại), 
tuy nhiên ko ảnh hưởng quá nhiều
"""

def run(
    run_time=None,
):

    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']
    now = dt.datetime.now()

    start_date = bdate(start_date,-1).replace('/','-')
    end_date = end_date.replace('/','-')

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    # update dữ liệu từ file VSD lên data warehouse mới chạy được BC
    year = int(period.split('.')[1])
    month = int(period.split('.')[0])
    monthly_update(year,month)

    account_type = pd.read_sql(
        """SELECT account_code, account_type FROM account""",
        connect_DWH_CoSo,
        index_col='account_code',
    ).squeeze()

    foreign_investor_holding = pd.read_sql(
        f"""
        SELECT 
            [de083].[date],
            [de083].[account_code],
            [account].[account_type],
            [account].[nationality],
            [account].[foreign_trading_code],
            [de083].[ticker],
            [securities_list].[asset_type],
            [de083].[volume]
        FROM [de083]
        LEFT JOIN [account] ON [de083].[account_code] = [account].[account_code]
        LEFT JOIN [securities_list] ON [securities_list].[ticker] = [de083].[ticker]
        WHERE ([account].[account_type] LIKE N'%ngoài')
        AND [de083].[date] BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col='date',
    )
    foreign_investor_holding['asset_type'] = foreign_investor_holding['asset_type'].map({
        'Cổ phiếu thường': 'Co phieu niem yet',
        'Quyền chọn': 'Quyen chon',
        'Chứng chỉ quỹ': 'Chung chi quy',
        'Trái phiếu': 'Trai phieu',
        'Chứng quyền': 'Chung quyen',
        'Trái phiếu doanh nghiệp': 'Trai phieu doanh nghiep',
        'Trái phiếu chuyển đổi': 'Trai phieu chuyen doi',
        'Chứng khoán phái sinh': 'Chung khoan phai sinh',
    })
    foreign_investor_holding.loc[foreign_investor_holding.index.min(),'period'] = 'opening'
    foreign_investor_holding.loc[foreign_investor_holding.index.max(),'period'] = 'closing'
    foreign_investor_holding = foreign_investor_holding.dropna(subset=['period'])
    foreign_investor_holding.reset_index(drop=True,inplace=True)

    cash_balance = pd.read_sql(
        f"""
        SELECT 
            [sub_account].[account_code],
            [sub_account_deposit].[date],
            [sub_account_deposit].[closing_balance] AS [balance]
        FROM [sub_account_deposit]
        LEFT JOIN [sub_account]
        ON [sub_account_deposit].[sub_account] = [sub_account].[sub_account]
        WHERE [sub_account_deposit].[date] IN ('{start_date}','{end_date}')
        """,
        connect_DWH_CoSo,
        index_col='date'
    )
    cash_balance.loc[cash_balance.index.min(),'period'] = 'opening'
    cash_balance.loc[cash_balance.index.max(),'period'] = 'closing'
    cash_balance.reset_index(drop=True,inplace=True)
    foreign_account = pd.read_sql(
        f"SELECT account_code FROM account WHERE account.account_type LIKE '%Ngoài'",
        connect_DWH_CoSo,
    ).squeeze()
    cash_balance = cash_balance.loc[cash_balance['account_code'].isin(foreign_account)]
    # lay san giao dich, xep loai lai co phieu
    exchange = scrape_ticker_by_exchange.run(False)
    foreign_investor_holding['exchange'] = foreign_investor_holding['ticker'].map(exchange.squeeze())
    foreign_investor_holding['exchange'].fillna('OTC',inplace=True)
    cophieuupcom_mask = (foreign_investor_holding['exchange']=='UPCOM')&(foreign_investor_holding['asset_type']=='Co phieu niem yet')
    foreign_investor_holding.loc[cophieuupcom_mask,'asset_type'] = 'Co phieu upcom'
    loainkhac_mask = (foreign_investor_holding['exchange']=='OTC')&(foreign_investor_holding['asset_type']=='Co phieu niem yet')
    foreign_investor_holding.loc[loainkhac_mask,'asset_type'] = 'Loai khac'

    # Phái sinh
    derivatives = pd.read_sql(
        f"""
        SELECT 
			[t].[account_code],
            [a].[account_type],
            [a].[nationality],
            [a].[foreign_trading_code],
            [t].[ticker],
            'Chung khoan phai sinh' [asset_type],
            [t].[volume],
            [t].[period],
            '' [exchange]
		FROM (
            SELECT 
                [rdo0009].[account_code],
                [rdo0009].[ticker],
                SUM([rdo0009].[o_long_volume] + [rdo0009].[o_short_volume]) [volume],
                'opening' [period]
            FROM [rdo0009]
            WHERE 
                [rdo0009].[date] = (
                    SELECT MAX([rdo0009].[date]) FROM [rdo0009] WHERE [rdo0009].[date] <= '{start_date}'
                )
                AND [rdo0009].[account_code] LIKE '022F%'
            GROUP BY [rdo0009].[account_code], [rdo0009].[ticker]

		    UNION ALL
            
            SELECT 
                [rdo0009].[account_code],
                [rdo0009].[ticker],
                SUM([rdo0009].[o_long_volume] + [rdo0009].[o_short_volume]) [volume],
                'closing' [period]
            FROM [rdo0009]
            WHERE 
                [rdo0009].[date] = (
                    SELECT MAX([rdo0009].[date]) FROM [rdo0009] WHERE [rdo0009].[date] <= '{end_date}'
                ) 
                AND [rdo0009].[account_code] LIKE '022F%'
            GROUP BY [rdo0009].[account_code], [rdo0009].[ticker]
        ) [t]
        LEFT JOIN [account] [a]
            ON [a].[account_code] = [t].[account_code]
        """,
        connect_DWH_PhaiSinh,
    )
    foreign_investor_holding = pd.concat([foreign_investor_holding,derivatives])

    # =========================================================================
    # Du lieu sheet I
    table_sheetI = foreign_investor_holding.groupby(['asset_type','ticker','period'])['volume'].sum()
    table_sheetI = table_sheetI.unstack(level=2,fill_value=0)
    table_sheetI = table_sheetI.reindex(['closing','opening','volume'],axis=1,fill_value=0)
    table_sheetI['change'] = table_sheetI['closing'] - table_sheetI['opening']
    table_sheetI.reset_index(level=1,inplace=True)

    # lay gia
    price = pd.DataFrame(
        index=table_sheetI.loc[table_sheetI.index!='Loai khac','ticker'].unique()
    )
    cw_list = price.loc[(price.index.str.startswith('C'))&(price.index.str.len()==8)].index.to_list()
    cw_price = pd.read_sql(
        f"EXEC [dbo].[spCoveredWarrantIntraday] " \
        f"@FROM_DATE = N'{end_date}', " \
        f"@TO_DATE = N'{end_date}'",
        connect_RMD,
        index_col='SYMBOL',
    )
    cw_price = cw_price.loc[cw_list,'CLOSE']
    derivatives_list = price.index[price.index.str.startswith('VN30F')].to_list()
    priceDate = bdate(end_date,1)
    priceDateDT = dt.datetime(int(priceDate[:4]),int(priceDate[5:7]),int(priceDate[-2:]))
    derivatives_price = scrape_derivatives_price.run(derivatives_list,priceDateDT)

    def price_mapper(x):
        if x in cw_list:
            price = cw_price.loc[x]*1000
        elif x in derivatives_list:
            price = derivatives_price.loc[x]*100000 # giá * hệ số nhân HĐ
        else:  # co phieu thuong
            price = ta.hist(x,start_date,end_date).tail(1)['close'].squeeze()*1000
        return price

    price['last_price'] = price.index.map(price_mapper)

    # Reformat bảng
    full_type = pd.Index([
        'Tong tin phieu',
        'Tin phieu',
        'Tong trai phieu',
        'Tong trai phieu chinh phu',
        'Trai phieu chinh phu',
        'Tong trai phieu chinh quyen dia phuong',
        'Trai phieu chinh quyen dia phuong',
        'Tong trai phieu doanh nghiep',
        'Trai phieu doanh nghiep',
        'Tong co phieu',
        'Tong co phieu niem yet',
        'Co phieu niem yet',
        'Tong co phieu upcom',
        'Co phieu upcom',
        'Tong gia tri von gop mua co phan va co phieu khac',
        'Gia tri von gop mua co phan va co phieu khac',
        'Tong chung chi quy',
        'Chung chi quy',
        'Tong chung khoan phai sinh',
        'Chung khoan phai sinh',
        'Tong loai khac',
        'Loai khac',
    ])
    appeared_type = table_sheetI.index.unique()
    missing_type = full_type.difference(appeared_type)
    for t in missing_type:  # enlargement only work on single index
        table_sheetI.loc[t,'ticker'] = ''
        if t in [
            'Trai phieu chinh phu',
            'Trai phieu chinh quyen dia phuong',
            'Trai phieu doanh nghiep',
            'Co phieu niem yet',
            'Co phieu upcom',
            'Gia tri von gop mua co phan va co phieu khac',
            'Chung khoan phai sinh',
        ]:
            table_sheetI.loc[t,['closing','opening','change']] = 0  # missing value means 0
    for t in missing_type:
        if t in [
            'Tong tin phieu',
            'Tong trai phieu chinh phu',
            'Tong trai phieu doanh nghiep',
            'Tong trai phieu chinh quyen dia phuong',
            'Tong co phieu niem yet',
            'Tong co phieu upcom',
            'Tong gia tri von gop mua co phan va co phieu khac',
            'Tong chung chi quy',
            'Tong chung khoan phai sinh',
            'Tong loai khac',
        ]:
            def get_elems(x):
                result = x.replace('Tong ','')
                first_char = result[0].upper()
                result = first_char+result[1:]
                return result

            sum_series = table_sheetI.loc[get_elems(t),['closing','opening','change']].sum()
            table_sheetI.loc[t,['closing','opening','change']] = sum_series
    # Tong trai phieu
    sub_level = ['Tong trai phieu chinh phu','Tong trai phieu chinh quyen dia phuong','Tong trai phieu doanh nghiep']
    sum_series = table_sheetI.loc[sub_level,['closing','opening','change']].sum()
    table_sheetI.loc['Tong trai phieu',['closing','opening','change']] = sum_series
    # Tong co phieu
    sub_level = ['Tong co phieu niem yet','Tong co phieu upcom','Tong gia tri von gop mua co phan va co phieu khac']
    sum_series = table_sheetI.loc[sub_level,['closing','opening','change']].sum()
    table_sheetI.loc['Tong co phieu',['closing','opening','change']] = sum_series
    # Tong chung khoan phai sinh
    sum_series = table_sheetI.loc['Chung khoan phai sinh',['closing','opening','change']].sum()
    table_sheetI.loc['Tong chung khoan phai sinh',['closing','opening','change']] = sum_series
    # Tính Cash
    table_sheetI = table_sheetI.loc[full_type]
    table_sheetI.loc['Tong tien','ticker'] = ''
    table_sheetI.loc['Tong tien',['closing','opening']] = cash_balance.groupby('period')['balance'].sum()
    table_sheetI.loc['Tong tien','change'] = table_sheetI.loc['Tong tien','closing']-table_sheetI.loc['Tong tien','opening']
    # Tính Value
    table_sheetI.loc['Tong cong','ticker'] = ''
    volume = table_sheetI.loc[table_sheetI['ticker']!='',['ticker','closing']].set_index('ticker')
    value = volume['closing']*price['last_price']
    # Tính Tổng Cộng
    table_sheetI.loc['Tong cong','closing'] = table_sheetI.loc['Tong tien','closing']+value.sum()
    table_sheetI.fillna(0,inplace=True)
    table_sheetI.rename(
        {
            'Tong tin phieu':'A. Tín phiếu',
            'Tin phieu':'',
            'Tong trai phieu':'B. Trái phiếu',
            'Tong trai phieu chinh phu':'Trái phiếu chính phủ',
            'Trai phieu chinh phu':'',
            'Tong trai phieu chinh quyen dia phuong':'Tổng trái phiếu chính quyền địa phương',
            'Trai phieu chinh quyen dia phuong':'',
            'Tong trai phieu doanh nghiep':'Trái phiếu doanh nghiệp',
            'Trai phieu doanh nghiep':'',
            'Tong co phieu':'C. Cổ phiếu',
            'Tong co phieu niem yet':'Cổ phiếu niêm yết',
            'Co phieu niem yet':'',
            'Tong co phieu upcom':'Cổ phiếu công ty đại chúng đăng ký giao dịch (upcom)',
            'Co phieu upcom':'',
            'Tong gia tri von gop mua co phan va co phieu khac':'Giá trị vốn góp mua cổ phần và cổ phiếu khác',
            'Gia tri von gop mua co phan va co phieu khac':'',
            'Tong chung chi quy':'D. Chứng chỉ quỹ, đơn vị quỹ thành viên',
            'Chung chi quy':'',
            'Tong chung khoan phai sinh':'Đ. Chứng khoán phái sinh',
            'Chung khoan phai sinh':'',
            'Tong loai khac':'E. Các loại chứng khoán khác',
            'Loai khac':'',
            'Tong tien':'F. Tiền (VND) và các khoản tương đương tiền (Chứng chỉ tiền gửi và các công cụ thị trường tiền tệ…)',
            'Tong cong':'Tổng cộng',
        },
        inplace=True
    )

    # =========================================================================
    # Du lieu sheet II
    foreign_business_type = pd.read_excel(
        join(dirname(dirname(__file__)),'foreign_business_type.xlsx'),
        index_col='account_code',
        squeeze=True,
    )
    cash_balance['account_type'] = cash_balance['account_code'].map(account_type)
    cash_balance = cash_balance.loc[cash_balance['period']=='closing']
    cash_balance_tochuc = cash_balance.loc[cash_balance['account_type']=='Tổ chức nước ngoài',['account_code','balance']]
    cash_balance_tochuc = cash_balance_tochuc.set_index('account_code').squeeze()
    cash_balance_canhan = cash_balance.loc[cash_balance['account_type']=='Cá nhân nước ngoài','balance'].sum()

    tochuc_sheetII = foreign_investor_holding.loc[foreign_investor_holding['period']=='closing'].copy()
    tochuc_sheetII.insert(6,'value',tochuc_sheetII['ticker'].map(price.squeeze())*tochuc_sheetII['volume'])
    tochuc_sheetII = tochuc_sheetII.groupby(
        ['account_type','account_code','nationality','foreign_trading_code','asset_type']
    )['value'].sum()
    tochuc_sheetII.sort_index(level=0,ascending=False,inplace=True)
    tochuc_sheetII = tochuc_sheetII.reset_index([0,2,3,4])

    canhan_sheetII = tochuc_sheetII.loc[tochuc_sheetII['account_type']=='Cá nhân nước ngoài']
    canhan_sheetII = canhan_sheetII.groupby(['account_type','asset_type'])['value'].sum().droplevel(level=0)
    full_type = {
        'Tin phieu':'Gia tri tin phieu',
        'Trai phieu ngan':'Gia tri trai phieu ngan han',
        'Trai phieu trung han':'Gia tri trai phieu trung han',
        'Trai phieu dai han':'Gia tri trai phieu dai han',
        'Co phieu niem yet':'Gia tri co phieu niem yet',
        'Chung chi quy':'Gia tri chung chi quy',
        'Co phieu upcom':'Gia tri co phieu upcom',
        'Gia tri von gop mua co phan va co phieu khac':'Gia tri von gop mua co phan va co phieu khac',
        'Loai khac':'Gia tri loai khac',
        'Chung khoan phai sinh': 'Chung khoan phai sinh',
        'Tien':'Gia tri tien',
        'Tong gia tri danh muc':'Tong gia tri danh muc',
    }
    canhan_sheetII.rename(full_type,axis=0,inplace=True)
    canhan_sheetII = canhan_sheetII.reindex(full_type.values())
    canhan_sheetII.loc['Gia tri co phieu niem yet va chung chi quy'] = canhan_sheetII.loc[['Gia tri co phieu niem yet','Gia tri chung chi quy']].sum()
    canhan_sheetII.drop(['Gia tri co phieu niem yet','Gia tri chung chi quy'],inplace=True)
    canhan_sheetII.loc['Gia tri von gop mua co phan va chung khoan khac'] = canhan_sheetII.loc[['Gia tri von gop mua co phan va co phieu khac','Gia tri loai khac','Chung khoan phai sinh']].sum()
    canhan_sheetII.drop(['Gia tri von gop mua co phan va co phieu khac','Gia tri loai khac'],inplace=True)
    canhan_sheetII.loc['Gia tri tien'] = cash_balance_canhan
    canhan_sheetII.loc['Tong gia tri danh muc'] = canhan_sheetII.sum()
    canhan_sheetII.name = 'Tổng (2)'
    canhan_sheetII = canhan_sheetII.to_frame().transpose()
    canhan_sheetII.loc['B-Cá nhân'] = np.nan
    canhan_sheetII = canhan_sheetII.loc[['B-Cá nhân','Tổng (2)']]

    tochuc_sheetII = tochuc_sheetII.loc[tochuc_sheetII['account_type']=='Tổ chức nước ngoài']
    tochuc_sheetII.insert(2,'business_type',foreign_business_type)
    # to avoid double count
    original_tochuc_sheetII = tochuc_sheetII.copy()
    tochuc_sheetII.drop(['asset_type','value'],axis=1,inplace=True)
    tochuc_sheetII.drop_duplicates(inplace=True)
    # Tin phieu (ok)
    tochuc_sheetII['Gia tri tin phieu'] = original_tochuc_sheetII.loc[original_tochuc_sheetII['asset_type']=='Tin phieu','value']
    tinphieu_sum = tochuc_sheetII['Gia tri tin phieu'].sum()
    if tinphieu_sum == 0:
        tochuc_sheetII['Ty le tin phieu'] = 0
    else:
        tochuc_sheetII['Ty le tin phieu'] = tochuc_sheetII['Gia tri tin phieu']/tinphieu_sum

    # Trai phieu (khong phan duoc theo ky han nen mac dinh khong co gia tri -> mat tinh tong quat)
    tochuc_sheetII[[
        'Gia tri trai phieu ngan han',
        'Ty le trai phieu ngan han',
        'Gia tri trai phieu trung han',
        'Ty le trai phieu trung han',
        'Gia tri trai phieu dai han',
        'Ty le trai phieu dai han',
    ]] = np.nan

    # Co phieu niem yet va chung chi quy
    asset_type_sum = original_tochuc_sheetII.loc[original_tochuc_sheetII['asset_type'].isin(['Co phieu niem yet','Chung chi quy']),'value']
    asset_type_sum = asset_type_sum.groupby('account_code').sum()
    tochuc_sheetII['Gia tri co phieu niem yet va chung chi quy'] = asset_type_sum
    cpccq_sum = tochuc_sheetII['Gia tri co phieu niem yet va chung chi quy'].sum()
    if cpccq_sum == 0:
        tochuc_sheetII['Ty le co phieu niem yet va chung chi quy'] = 0
    else:
        tochuc_sheetII['Ty le co phieu niem yet va chung chi quy'] = tochuc_sheetII['Gia tri co phieu niem yet va chung chi quy']/cpccq_sum

    # Co phieu upcom
    tochuc_sheetII['Gia tri co phieu upcom'] = original_tochuc_sheetII.loc[original_tochuc_sheetII['asset_type']=='Co phieu upcom','value']
    cophieuupcom_sum = tochuc_sheetII['Gia tri co phieu upcom'].sum()
    if cophieuupcom_sum == 0:
        tochuc_sheetII['Ty le co phieu upcom'] = 0
    else:
        tochuc_sheetII['Ty le co phieu upcom'] = tochuc_sheetII['Gia tri co phieu upcom']/cophieuupcom_sum

    # Gia tri von gop mua co phan va chung khoan khac
    asset_type_sum = original_tochuc_sheetII.loc[original_tochuc_sheetII['asset_type'].isin(['Gia tri von gop mua co phan va co phieu khac','Chung khoan phai sinh','Loai khac']),'value']
    asset_type_sum = asset_type_sum.groupby('account_code').sum()
    tochuc_sheetII['Gia tri von gop mua co phan va chung khoan khac'] = asset_type_sum
    value_sum = tochuc_sheetII['Gia tri von gop mua co phan va chung khoan khac'].sum()
    if value_sum == 0:
        tochuc_sheetII['Ty le von gop mua co phan va chung khoan khac'] = 0
    else:
        tochuc_sheetII['Ty le von gop mua co phan va chung khoan khac'] = tochuc_sheetII['Gia tri von gop mua co phan va chung khoan khac']/value_sum

    # Tien
    tochuc_sheetII['Gia tri tien'] = cash_balance_tochuc
    tien_sum = tochuc_sheetII['Gia tri tien'].sum()
    if tien_sum == 0:
        tochuc_sheetII['Ty le tien'] = 0
    else:
        tochuc_sheetII['Ty le tien'] = tochuc_sheetII['Gia tri tien']/tochuc_sheetII['Gia tri tien'].sum()

    # Tong gia tri danh muc
    tonggiatridanhmuc = tochuc_sheetII.loc[:,[col.startswith('Gia tri') for col in tochuc_sheetII.columns]].sum(axis=1)
    tochuc_sheetII['Tong gia tri danh muc'] = tonggiatridanhmuc

    # Dong Tong (1)
    tochuc_sheetII.reset_index(inplace=True)
    tochuc_sheetII.loc['Tổng (1)'] = np.nan
    for col in tochuc_sheetII.columns:
        if col.startswith('Gia tri') or col=='Tong gia tri danh muc':
            tochuc_sheetII.loc['Tổng (1)',col] = tochuc_sheetII.loc[tochuc_sheetII.index[:-1],col].fillna(0).sum(skipna=False)
        elif col.startswith('Ty le'):
            tochuc_sheetII.loc['Tổng (1)',col] = tochuc_sheetII.loc[tochuc_sheetII.index[:-1],col].fillna(0).mean(skipna=False)

    first_row_sheetII = pd.DataFrame(columns=tochuc_sheetII.columns,index=['A-Tổ chức'])
    tong_row_sheetII = pd.DataFrame(columns=tochuc_sheetII.columns,index=['Tổng = (1) + (2)'])
    table_sheetII = pd.concat([
        first_row_sheetII,
        tochuc_sheetII,
        canhan_sheetII,
        tong_row_sheetII,
    ])
    fill_with_zero_col = [col.startswith('Gia tri') or col.startswith('Ty le') or col.startswith('Tong gia tri') for col in table_sheetII.columns]
    fill_with_zero_empty = [col for col in tochuc_sheetII.columns if col not in fill_with_zero_col]
    table_sheetII.loc[:,fill_with_zero_col] = table_sheetII.loc[:,fill_with_zero_col].fillna(0)
    table_sheetII.loc[:,fill_with_zero_empty] = table_sheetII.loc[:,fill_with_zero_empty].fillna('')
    table_sheetII = table_sheetII[[  # ensure column order
        'account_code',
        'account_type',
        'nationality',
        'business_type',
        'foreign_trading_code',
        'Gia tri tin phieu',
        'Ty le tin phieu',
        'Gia tri trai phieu ngan han',
        'Ty le trai phieu ngan han',
        'Gia tri trai phieu trung han',
        'Ty le trai phieu trung han',
        'Gia tri trai phieu dai han',
        'Ty le trai phieu dai han',
        'Gia tri co phieu niem yet va chung chi quy',
        'Ty le co phieu niem yet va chung chi quy',
        'Gia tri co phieu upcom',
        'Ty le co phieu upcom',
        'Gia tri von gop mua co phan va chung khoan khac',
        'Ty le von gop mua co phan va chung khoan khac',
        'Gia tri tien',
        'Ty le tien',
        'Tong gia tri danh muc',
    ]]
    table_sheetII.index = list(map(str,table_sheetII.index))
    for col in table_sheetII.columns:
        if col.startswith('Gia tri') or col=='Tong gia tri danh muc':
            table_sheetII.loc['Tổng = (1) + (2)',col] = table_sheetII.loc['Tổng (1)',col]+table_sheetII.loc[
                'Tổng (2)',col]

    # =========================================================================
    # Write to Báo cáo hoạt động lưu ký nhà đầu tư nước ngoài
    file_name = f'Báo cáo hoạt động lưu ký nhà đầu tư nước ngoài {period} bk.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    # =========================================================================
    # Write to sheet Gioi Thieu Sheet
    gioithieu_sheet = workbook.add_worksheet('GioiThieu')
    gioithieu_sheet.hide_gridlines(option=2)
    # set column width
    gioithieu_sheet.set_column('A:N',9)
    gioithieu_sheet.set_row(2,21)
    gioithieu_sheet.set_row(6,27)
    gioithieu_sheet.set_row(25,95)
    text_center_bold = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':10.5,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter'
        }
    )
    text_left_bold = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':10.5,
            'text_wrap':True,
            'valign':'vcenter'
        }
    )
    text_right_normal = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':10.5,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
        }
    )
    text_center_normal = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':10.5,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter'
        }
    )
    text_center_normal_small = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter'
        }
    )
    gioithieu_sheet.merge_range('A1:D1',CompanyName,text_center_bold)
    gioithieu_sheet.merge_range('I1:N1','CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',text_center_bold)
    gioithieu_sheet.merge_range('I2:N2','Độc lập - Tự do - Hạnh phúc',text_center_bold)
    gioithieu_sheet.write('A3','Số:',text_right_normal)
    gioithieu_sheet.merge_range('B3:D3',period.replace('.','/')+'/BC-DVKH',text_left_bold)
    gioithieu_sheet.merge_range(
        'A4:F5',
        'V/v báo cáo thống kê danh mục lưu ký của nhà đầu tư nước ngoài, '
        'tổ chức phát hành chứng chỉ lưu ký tại nước ngoài (PLIII-TT51/2011/TT-BTC)',
        text_center_normal_small,
    )
    gioithieu_sheet.write('H4','Hà Nội',text_center_normal)
    gioithieu_sheet.write('I4','Ngày',text_center_normal)
    gioithieu_sheet.write('J4',convert_int(now.day),text_center_bold)
    gioithieu_sheet.write('K4','Tháng',text_center_normal)
    gioithieu_sheet.write('L4',convert_int(now.month),text_center_bold)
    gioithieu_sheet.write('M4','Năm',text_center_normal)
    gioithieu_sheet.write('N4',now.year,text_center_bold)
    gioithieu_sheet.merge_range(
        'A7:N7',
        'BÁO CÁO THỐNG KÊ DANH MỤC LƯU KÝ CỦA NHÀ ĐẦU TƯ NƯỚC NGOÀI, TỔ CHỨC PHÁT HÀNH CHỨNG CHỈ LƯU KÝ TẠI NƯỚC NGOÀI'
        '\n'
        '(PLIII-TT51/2011/TT-BTC)',
        text_center_bold,
    )
    gioithieu_sheet.merge_range('C9:E9','Kỳ báo cáo:',text_right_normal)
    gioithieu_sheet.write('G9','Tháng',text_right_normal)
    gioithieu_sheet.write('H9',period.split('.')[0],text_center_bold)
    gioithieu_sheet.write('I9','Năm',text_right_normal)
    gioithieu_sheet.write('J9',period.split('.')[1],text_center_bold)
    gioithieu_sheet.merge_range('C11:D11','Thời điểm báo gửi cáo:',text_right_normal)
    gioithieu_sheet.write('E11','Ngày',text_center_normal)
    gioithieu_sheet.write('F11',convert_int(now.day),text_center_bold)
    gioithieu_sheet.write('G11','Tháng',text_right_normal)
    gioithieu_sheet.write('H11',period.split('.')[0],text_center_bold)
    gioithieu_sheet.write('I11','Năm',text_right_normal)
    gioithieu_sheet.write('J11',period.split('.')[1],text_center_bold)
    gioithieu_sheet.merge_range(
        'A22:N22',
        'Chúng tôi xin cam kết chịu trách nhiệm hoàn toàn về sự trung thực, đầy đủ chính xác '
        'của nội dung Giấy thông báo này và tài liệu kèm theo.',
        text_center_bold,
    )
    gioithieu_sheet.merge_range('B24:D24','Lập biểu',text_center_bold)
    gioithieu_sheet.merge_range('F24:H24','Kiểm soát',text_center_bold)
    gioithieu_sheet.merge_range('J24:N24','Đại diện có thẩm quyền của thành viên',text_center_bold)
    gioithieu_sheet.merge_range('J25:N25','(Ký số, ghi rõ họ tên, chức vụ)',text_center_normal)
    gioithieu_sheet.write('B27','Họ và tên',text_right_normal)
    gioithieu_sheet.merge_range('C27:D27','ĐIỀN TÊN',text_center_bold)
    gioithieu_sheet.write('F27','Họ và tên',text_right_normal)
    gioithieu_sheet.merge_range('G27:H27','Trần Thu Trang',text_center_bold)
    gioithieu_sheet.write('K27','Họ và tên',text_right_normal)
    gioithieu_sheet.merge_range('L27:M27','Chen Chia Ken',text_center_bold)
    gioithieu_sheet.write('B29','Chức vụ',text_right_normal)
    gioithieu_sheet.merge_range('C29:D29','Nhân viên',text_center_normal)
    gioithieu_sheet.write('F29','Chức vụ',text_right_normal)
    gioithieu_sheet.merge_range('G29:H29','Giám đốc K.DVKH',text_center_normal)
    gioithieu_sheet.write('K29','Chức vụ',text_right_normal)
    gioithieu_sheet.merge_range('L29:M29','Tổng giám đốc',text_center_normal)

    # =========================================================================
    # Write to Sheet I
    sheet_I = workbook.add_worksheet('I')
    sheet_I.hide_gridlines(option=2)
    # set column width
    sheet_I.set_column('A:A',35)
    sheet_I.set_column('B:B',23)
    sheet_I.set_column('C:E',20)
    sheet_I.set_row(0,38)
    sheet_I.set_row(2,24)
    sheet_I.set_row(4,30)
    headline_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    date_text_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    header_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'bg_color':'#A4CAFA',
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    normal_text_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    normal_number_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* -#,##0__;_(* "-"??_);_(@_)',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bold_text_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'bold':True,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bold_number_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* -#,##0__;_(* "-"??_);_(@_)',
            'font_size':11,
            'bold':True,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bold_footer_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    normal_footer_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'top',
            'text_wrap':True,
        }
    )
    sheet_I.merge_range(
        'A1:E1',
        'BÁO CÁO THỐNG KÊ DANH MỤC LƯU KÝ CỦA NHÀ ĐẦU TƯ NƯỚC NGOÀI, '
        'TỔ CHỨC PHÁT HÀNH CHỨNG CHỈ LƯU KÝ TẠI NƯỚC NGOÀI (PLIII-TT51/2011/TT-BTC)',
        headline_fmt
    )
    sheet_I.merge_range(
        'A2:E2',
        f'(Tháng {period.split(".")[0]} năm {period.split(".")[1]})',
        date_text_fmt
    )
    sheet_I.merge_range(
        'A3:E3',
        'I. Báo cáo chi tiết theo danh mục',
        headline_fmt
    )
    sheet_I.merge_range('A4:A5','Danh mục tài sản',header_fmt)
    sheet_I.merge_range('B4:B5','Danh mục (theo mã chứng khoán)',header_fmt)
    sheet_I.merge_range('C4:E4','Số lượng chứng khoán lưu ký',header_fmt)
    sheet_I.write_row('C5',['Kỳ báo cáo','Kỳ báo cáo trước','Thay đổi so với kỳ báo cáo trước(+/-)'],header_fmt)

    sheet_I.write_column('A6',table_sheetI.index,normal_text_cell_fmt)
    sheet_I.write_column('B6',table_sheetI['ticker'],normal_text_cell_fmt)
    for start_cell,col_name in zip(['C6','D6','E6'],['closing','opening','change']):
        sheet_I.write_column(start_cell,table_sheetI[col_name],normal_number_cell_fmt)
    end_of_table_row = 4+table_sheetI.shape[0]

    # overwrite normal row by bold version
    sheet_I.write(end_of_table_row,0,'Tổng cộng',bold_text_cell_fmt)
    sheet_I.write(end_of_table_row,1,'',bold_text_cell_fmt)
    sheet_I.write(end_of_table_row,2,table_sheetI.loc['Tổng cộng','closing'],bold_number_cell_fmt)
    sheet_I.write(end_of_table_row,3,'XXXXXX',bold_number_cell_fmt)
    sheet_I.write(end_of_table_row,4,'XXXXXX',bold_number_cell_fmt)

    sheet_I.merge_range(
        end_of_table_row+2,0,end_of_table_row+2,4,
        'Chúng tôi xin cam kết chịu trách nhiệm hoàn toàn về sự trung thực, '
        'đầy đủ chính xác của nội dung Giấy thông báo này và tài liệu kèm theo.',
        bold_footer_fmt
    )
    sheet_I.write_row(end_of_table_row+4,0,['Lập biểu','Kiểm soát'],bold_footer_fmt)
    sheet_I.merge_range(end_of_table_row+4,2,end_of_table_row+4,4,'Đại diện có thẩm quyền của thành viên',
                        bold_footer_fmt)
    sheet_I.merge_range(end_of_table_row+5,2,end_of_table_row+5,4,'(Ký số, ghi rõ họ tên, chức vụ)',bold_footer_fmt)
    sheet_I.write_row(end_of_table_row+11,0,['Điền tên vào đây','Trần Thu Trang'],bold_footer_fmt)
    sheet_I.merge_range(end_of_table_row+11,2,end_of_table_row+11,4,'Chen Chia Ken',bold_footer_fmt)
    sheet_I.write_row(end_of_table_row+12,0,['Chức vụ: Nhân viên','Chức vụ: Giám đốc K.DVKH'],normal_footer_fmt)
    sheet_I.merge_range(end_of_table_row+12,2,end_of_table_row+12,4,'Chức vụ: Tổng Giám đốc',normal_footer_fmt)

    # =========================================================================
    # Write to Sheet II
    sheet_II = workbook.add_worksheet('II')
    sheet_II.hide_gridlines(option=2)
    sheet_II.set_column('A:A',17)
    sheet_II.set_column('B:B',14)
    sheet_II.set_column('C:C',25)
    sheet_II.set_column('D:D',40)
    sheet_II.set_column('E:E',9)
    wide_coulumns = ['F','H','J','L','N','P','R','T','V']
    for col in wide_coulumns:
        sheet_II.set_column(f'{col}:{col}',20)
    narrow_columns = ['G','I','K','M','O','Q','S','U']
    for col in narrow_columns:
        sheet_II.set_column(f'{col}:{col}',10)
    sheet_II.set_row(0,24)

    headline_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    sub_headline_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    footer_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'valign':'top',
            'text_wrap':True,
        }
    )
    header_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'bg_color':'#A4CAFA',
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    normal_text_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bold_text_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'bold':True,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    normal_number_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* -#,##0__;_(* "-"??_);_(@_)',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bold_number_cell_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* -#,##0__;_(* "-"??_);_(@_)',
            'font_size':11,
            'bold':True,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bold_pct_cell_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'num_format':'0.0%',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    sheet_II.merge_range(
        'A1:V1',
        'BÁO CÁO THỐNG KÊ DANH MỤC LƯU KÝ CỦA NHÀ ĐẦU TƯ NƯỚC NGOÀI, '
        'TỔ CHỨC PHÁT HÀNH CHỨNG CHỈ LƯU KÝ TẠI NƯỚC NGOÀI (PLIII-TT51/2011/TT-BTC)',
        headline_fmt
    )
    sheet_II.merge_range(
        'A2:V2',
        f'(Tháng {period.split(".")[0]} năm {period.split(".")[1]})',
        sub_headline_fmt,
    )
    sheet_II.merge_range(
        'A3:V3',
        'II. Báo cáo cơ cấu danh mục theo tỷ trọng đầu tư của tổ chức và cá nhân',
        headline_fmt,
    )
    sheet_II.merge_range('A5:B9','Tên khách hàng',header_fmt)
    sheet_II.merge_range('C5:C9','Quốc tịch',header_fmt)
    sheet_II.merge_range('D5:D9','Loại hình đối với tổ chức',header_fmt)
    sheet_II.merge_range('E5:E9','Mã số giao dịch chứng khoán',header_fmt)
    sheet_II.merge_range('F5:G8','Tín phiếu',header_fmt)
    sheet_II.merge_range('H5:M5','Trái phiếu',header_fmt)
    sheet_II.merge_range('N5:S5','Cổ phiếu/Chứng chỉ quỹ',header_fmt)
    sheet_II.merge_range('T5:U8','Tiền và các khoản tương đương tiền (chứng chỉ tiền gửi...)',header_fmt)
    sheet_II.merge_range('V5:V8','Tổng giá trị danh mục',header_fmt)
    sheet_II.merge_range('H6:M6','Thời gian còn lại tới khi đáo hạn',header_fmt)
    sheet_II.merge_range('H7:I7','Ngắn hạn',header_fmt)
    sheet_II.merge_range('J7:K7','Trung hạn',header_fmt)
    sheet_II.merge_range('L7:M7','Dài hạn',header_fmt)
    sheet_II.merge_range('H8:I8','Dưới 12 tháng',header_fmt)
    sheet_II.merge_range('J8:K8','Từ 12 tháng đến 24 tháng',header_fmt)
    sheet_II.merge_range('L8:M8','Trên 24 tháng',header_fmt)
    sheet_II.merge_range('N6:O8','Cổ phiếu niêm yết, chứng chỉ quỹ niêm yết',header_fmt)
    sheet_II.merge_range('P6:Q8','Cổ phiếu công ty đại chúng đăng ký giao dịch (upcom)',header_fmt)
    sheet_II.merge_range('R6:S8','Giá trị vốn góp mua cổ phần, quỹ thành viên và chứng khoán khác',header_fmt)
    sheet_II.write_row('F9',['Giá trị','Tỷ lệ (%)']*8+['Giá trị'],header_fmt)
    start_row = xlsxwriter.utility.xl_cell_to_rowcol('A10')[0]
    for row,elem in enumerate(table_sheetII.index):
        if elem in ['A-Tổ chức','B-Cá nhân']:
            fmt = bold_text_cell_fmt
        else:
            fmt = normal_text_cell_fmt
        if len(elem)<=3:  # able to accomodate 3-digit number
            elem = ''
        sheet_II.write(start_row+row,0,elem,fmt)
    start_cells = ['B10','C10','D10','E10']
    df_col_names = ['account_code','nationality','business_type','foreign_trading_code']
    for cell,col_name in zip(start_cells,df_col_names):
        sheet_II.write_column(cell,table_sheetII[col_name],normal_text_cell_fmt)
    start_row,start_col = xlsxwriter.utility.xl_cell_to_rowcol('F10')
    indices = table_sheetII.index
    columns = [col for col in table_sheetII.columns if 'Gia tri' in col or 'Ty le' in col or 'Tong gia tri' in col]
    for row,row_name in enumerate(indices):
        for col,col_name in enumerate(columns):
            if row_name.startswith('A-') or row_name.startswith('B-'):
                fmt = normal_text_cell_fmt
                value = ''
            elif 'Ty le' in col_name:
                if row_name=='Tổng (2)' or row_name=='Tổng = (1) + (2)':
                    fmt = normal_text_cell_fmt
                    value = ''
                else:
                    fmt = bold_pct_cell_fmt
                    value = table_sheetII.loc[row_name,col_name]
            elif 'Tong gia tri' in col_name or 'Tổng' in row_name:
                fmt = bold_number_cell_fmt
                value = table_sheetII.loc[row_name,col_name]
            else:
                fmt = normal_number_cell_fmt
                value = table_sheetII.loc[row_name,col_name]
            sheet_II.write(start_row+row,start_col+col,value,fmt)
    end_row = start_row+table_sheetII.shape[0]
    sheet_II.set_row(end_row+1,85)
    sheet_II.merge_range(
        end_row+1,0,end_row+1,3,
        'Ghi chú:\n'
        '1) Giá trị tính theo giá thị trường vào thời điểm báo cáo; tín phiếu, trái phiếu, '
        'chứng chỉ tiền gửi và cổ phiếu không có giao dịch thì tính theo mệnh giá hoặc giá trị mua vào;\n'
        '2) Giá trị tài sản tính theo đơn vị VND;\n'
        '3) Giá trị danh mục của tổ chức, cá nhân được sắp xếp theo thứ tự từ lớn nhất đến nhỏ nhất.\n',
        footer_fmt
    )

    # =========================================================================
    # Write to Sheet III
    sheet_III = workbook.add_worksheet('III')
    sheet_III.hide_gridlines(option=2)

    sheet_III.set_column('A:A',34)
    sheet_III.set_column('B:B',3)
    sheet_III.set_column('C:C',13)
    sheet_III.set_column('D:K',11)

    sheet_III.set_row(1,39)
    sheet_III.set_row(5,32)
    sheet_III.set_row(6,35)
    sheet_III.set_row(7,23)

    title_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    date_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    sub_title_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    header_fmt = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':12,
            'bg_color':'#99CCFF',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    text_cell_fmt = workbook.add_format(
        {
            'border':1,
            'font_name':'Times New Roman',
            'font_size':12,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    number_cell_fmt = workbook.add_format(
        {
            'border':1,
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'#,##0.00',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    title = 'BÁO CÁO THỐNG KÊ DANH MỤC LƯU KÝ CỦA NHÀ ĐẦU TƯ NƯỚC NGOÀI, ' \
            'TỔ CHỨC PHÁT HÀNH CHỨNG CHỈ LƯU KÝ TẠI NƯỚC NGOÀI (PLIII-TT51/2011/TT-BTC)'
    sheet_III.merge_range('A2:K2',title,title_fmt)
    sheet_III.merge_range('A3:K3',f"Tháng {period.split('.')[0]}/năm {period.split('.')[1]}",date_fmt)
    sub_title = 'III. Hoạt động kinh doanh chứng khoán của thành viên lưu ký là chi nhánh các tổ chức tín dụng nước ngoài, ' \
                'tổ chức tín dụng 100% vốn nước ngoài thành lập tại Việt Nam'
    sheet_III.merge_range('A4:K4',sub_title,sub_title_fmt)
    sheet_III.merge_range('A6:B8','STT',header_fmt)
    sheet_III.merge_range('C6:C7','Loại tài sản/ Mã chứng khoán',header_fmt)
    sheet_III.merge_range('D6:E6','Mua trong kỳ',header_fmt)
    sheet_III.merge_range('F6:G6','Bán trong kỳ',header_fmt)
    sheet_III.merge_range('H6:I6','Mua thuần trong kỳ',header_fmt)
    sheet_III.merge_range('J6:K6','Số dư cuối kỳ',header_fmt)
    sheet_III.write_row('D7',['Khối lượng','Giá trị']*4,header_fmt)
    sheet_III.write_row('C8',[f'({i})' for i in np.arange(5)+2],header_fmt)
    sheet_III.write('H8','(7) = (3)-(5)',header_fmt)
    sheet_III.write('I8','(8) = (4)-(6)',header_fmt)
    sheet_III.write('J8','(9)',header_fmt)
    sheet_III.write('K8','(10)',header_fmt)
    row_labels = [
        'A. Tín phiếu',
        '',
        'Tổng',
        'B. Trái phiếu',
        'B1.Trái phiếu có thời gian tới khi đáo hạn còn lại dưới 12 tháng',
        '',
        'Tổng',
        'B2. Trái phiếu có thời gian tới khi đáo hạn còn lại từ 12 tháng đến 24 tháng',
        '',
        'Tổng',
        'B3. Trái phiếu có thời gian tới khi đáo hạn còn lại trên 24 tháng đến 60 tháng',
        '',
        'Tổng',
        'B4. Trái phiếu có thời gian đáo hạn trên 60 tháng',
        '',
        'Tổng',
        'C. Cổ phiếu niêm yết, chứng chỉ quỹ niêm yết',
        '',
        'Tổng',
        'D. Cổ phiếu công ty đại chúng đăng ký giao dịch (upcom)',
        '',
        'Tổng',
        'Đ. Giá trị vốn góp mua cổ phần, đơn vị quỹ thành viên',
        '',
        'Tổng',
        'E. Các loại chứng khoán khác',
        '',
        'Tổng',
        'G. Tiền và các khoản tương đương tiền (Chứng chỉ tiền gửi và các công cụ thị trường tiền tệ ..)',
        '1. Tiền',
        '',
        '2. Chứng chỉ tiền gửi và công cụ thị trường tiền tệ...',
        '',
        'Tổng',
        'Tổng',
        'Ghi chú:',
        '1) Giá trị chứng khoán tính theo giá thị trường vào thời điểm báo cáo;',
        'Đối với chứng khoán không có giao dịch, giá trị tính theo giá mua vào hoặc mệnh giá;',
        '2) Giá trị tài sản tính theo đơn vị VND.'
    ]
    sheet_III.write_column('A9',row_labels,text_cell_fmt)
    for col in np.arange(10)+1:
        for row in np.arange(39)+8:
            if col in [1,2] or row in [8,11,12,15,18,21,24,27,30,33,36,39,42,43,44,45,46]:
                value = ''
                fmt = text_cell_fmt
            else:
                value = 0
                fmt = number_cell_fmt
            sheet_III.write(row,col,value,fmt)

    # =========================================================================

    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
