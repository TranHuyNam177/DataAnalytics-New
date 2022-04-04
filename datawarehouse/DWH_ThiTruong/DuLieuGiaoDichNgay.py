from request import *
from request.stock import ta
from datawarehouse import INSERT, DELETE

def update(
    fromDate: dt.datetime,
    toDate: dt.datetime
):

    """
    This function updates data to table [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]

    :param fromDate: Ngày bắt đầu cập nhật
    :param toDate: Ngày cuối cùng cập nhật
    """

    tickers = pd.read_sql(
        f"""
        SELECT [ticker]
        FROM [DanhSachMa]
        WHERE [date] = (SELECT MAX([date]) FROM [DanhSachMa])
        """,
        connect_DWH_ThiTruong
    ).squeeze()

    fromDateStr = fromDate.strftime('%Y-%m-%d')
    toDateStr = toDate.strftime('%Y-%m-%d')
    for ticker in tickers:
        if len(ticker) > 3:
            continue
        try:
            print(f'{ticker} inserted')
            table = ta.hist(ticker,fromDateStr,toDateStr)
            table['trading_date'] = pd.to_datetime(table['trading_date'],format='%Y-%m-%d')
            table = table.rename(
                {
                    'date':'Date',
                    'ticker':'Ticker',
                    'ref':'Ref',
                    'open':'Open',
                    'close':'Close',
                    'high':'High',
                    'low':'Low',
                    'trading_date':'Date',
                    'total_volume':'Volume',
                    'total_value':'Value'
                },
                axis=1
            )
            cols = ['Date','Ticker','Ref','Open','Close','High','Low','Volume','Value']
            table = table[cols]
            multcols = ['Ref','Open','Close','High','Low']
            table[multcols] = table[multcols] * 1000
            DELETE(
                connect_DWH_ThiTruong,
                "DuLieuGiaoDichNgay",
                f"WHERE ([Ticker] = '{ticker}') AND ([Date] BETWEEN '{fromDateStr}' AND '{toDateStr}')")
            INSERT(connect_DWH_ThiTruong,'DuLieuGiaoDichNgay',table)
        except KeyError:
            print(f'{ticker} không có giá ở ta.hist')

