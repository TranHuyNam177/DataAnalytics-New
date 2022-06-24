from request import *
from request.stock import ta
from datawarehouse import BATCHINSERT, DELETE

def run(
    fromDate:dt.datetime,
    toDate:dt.datetime,
):

    """
    This function get data from API: 'https://api.phs.vn/market/utilities.svc/GetShareIntraday'
    and write to [DWH-ThiTruong].[DuLieuGiaoDichNgay]
    """

    start = time.time()
    URL = 'https://api.phs.vn/market/utilities.svc/GetShareIntraday'

    if fromDate < dt.datetime(year=2018,month=1,day=1):
        raise Exception('Only data after 2018-01-01 is reliable')

    fromDateString = fromDate.strftime('%Y-%m-%d')
    toDateString = toDate.strftime('%Y-%m-%d')

    tickers = pd.read_sql(
        f"""
        SELECT DISTINCT
            [DanhSachMa].[Ticker] 
        FROM [DanhSachMa]
        """,
        connect_DWH_ThiTruong,
    ).squeeze()

    with requests.Session() as session:
        retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
        adapter = requests.adapters.HTTPAdapter(max_retries=retry)
        session.mount('https://',adapter)
        frames = []
        for ticker in tickers:
            r = session.post(
                URL,
                data=json.dumps({
                    'symbol':ticker,
                    'fromdate':fromDateString,
                    'todate':toDateString,
                }),
                headers={'content-type':'application/json'}
            )
            frame = pd.DataFrame(json.loads(r.json()['d']))
            if not frame.empty:
                frame.loc[frame['open_price']==0,'open_price'] = frame['prior_price']
                frame.loc[frame['close_price']==0,'close_price'] = frame['prior_price']
                frame.loc[frame['high']==0,'high'] = frame['prior_price']
                frame.loc[frame['low']==0,'low'] = frame['prior_price']
            print(f'Extracted {ticker}')
            frames.append(frame)

    table = pd.concat(frames)
    nameMapper = {
        'trading_date':'Date',
        'symbol':'Ticker',
        'prior_price':'Ref',
        'open_price':'Open',
        'close_price':'Close',
        'high':'High',
        'low':'Low',
        'total_volume':'Volume',
        'total_value':'Value',
    }
    table.rename(columns=nameMapper,inplace=True)
    table = table.reindex(nameMapper.values(),axis=1)
    table[['Volume','Value']] = table[['Volume','Value']].astype(np.float64)

    def converter(date_str:str):
        month,day,year = np.array(date_str.split('/'),dtype=int)
        return dt.datetime(year,month,day)
    table['Date'] = table['Date'].str.split().str.get(0).map(converter)

    print(f'Inserting to [DWH-ThiTruong]...')
    DELETE(connect_DWH_ThiTruong,'DuLieuGiaoDichNgay',f"""WHERE [Date] BETWEEN '{fromDateString}' AND '{toDateString}'""")
    BATCHINSERT(connect_DWH_ThiTruong,'DuLieuGiaoDichNgay',table)

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
