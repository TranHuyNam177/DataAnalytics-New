import pandas as pd

from request.stock import *
from datawarehouse import *


def run(
    exchange:str='HOSE',
) -> pd.DataFrame:

    """
    This method returns list of tickers that up ceil (fc_type='ceil')
    or down floor (fc_type='floor') in a given exchange of a given segment
    in n consecutive trading days

    :param exchange: allow values in fa.exchanges. Do not allow 'all'
    :return: pd.DataFrame (columns: 'Ticker', 'Exchange', 'Consecutive Days'
    """

    start_time = time.time()
    now = dt.datetime.now().strftime('%Y-%m-%d')
    # xét 3 tháng gần nhất, nếu sửa ở đây thì phải sửa dòng ngay dưới
    since = bdate(now,-22*3)
    three_month_ago = since

    mrate_series = internal.margin['mrate']
    drate_series = internal.margin['drate']
    maxprice_series = internal.margin['max_price']
    general_room_seires = internal.margin['general_room']
    special_room_series = internal.margin['special_room']
    total_room_series = internal.margin['total_room']

    full_tickers = internal.mlist(exchanges=[exchange])
    full_tickers = tuple(full_tickers)

    if not full_tickers:
        return pd.DataFrame()

    records = []
    df = pd.read_sql(
        f"""
        SELECT
            [DuLieuGiaoDichNgay].[Date] [trading_date],
            [DuLieuGiaoDichNgay].[Ticker] [stock],
            CAST(ROUND([DuLieuGiaoDichNgay].[Ref] * 1000, 0) AS INT) [ref],
            CAST(ROUND([DuLieuGiaoDichNgay].[Close] * 1000, 0) AS INT) [close],
            CAST([DuLieuGiaoDichNgay].[Volume] AS INT) [total_volume]
        FROM [DuLieuGiaoDichNgay]
        WHERE [DuLieuGiaoDichNgay].[Date] BETWEEN '{since}' AND '{now}'
        AND [DuLieuGiaoDichNgay].[Ticker] IN {full_tickers}
        """,
        connect_DWH_ThiTruong,
        index_col='trading_date'
    )
    df['floor'] = df['ref'].apply(
        fc_price,
        price_type='floor',
        exchange=exchange
    )
    df.index = pd.to_datetime(df.index).strftime('%Y-%m-%d')

    for stock in df['stock'].unique():
        df_sub = df[df['stock']==stock].sort_index()
        # giam san lien tiep
        n_floor = 0
        for i in range(df_sub.shape[0]):
            condition1 = (df_sub.loc[df_sub.index[-i-1:],'floor']==df_sub.loc[df_sub.index[-i-1:],'close']).all()
            condition2 = (df_sub.loc[df_sub.index[-i-1:],'floor']!=df_sub.loc[df_sub.index[-i-1:],'ref']).all()
            # the second condition is to ignore trash tickers whose price
            # less than 1000 VND (a single price step equivalent to
            # more than 7%(HOSE), 10%(HNX), 15%(UPCOM))
            if condition1 and condition2:
                n_floor += 1
            else:
                break
        # mat thanh khoan trong 1 thang
        avg_vol_1m = df_sub.loc[df_sub.index[-22]:,'total_volume'].mean()
        n_illiquidity = (df_sub.loc[df_sub.index[-22]:,'total_volume']<avg_vol_1m).sum()
        # thanh khoan ngay gan nhat so voi thanh khoan trung binh 1 thang
        volume = df_sub.loc[df_sub.index[-1],'total_volume']
        volume_change_1m = volume/avg_vol_1m-1

        n_illiquidity_bmk = 1
        n_floor_bmk = 1

        condition1 = n_floor>=n_floor_bmk
        condition2 = n_illiquidity>=n_illiquidity_bmk

        if condition1 or condition2:
            print(stock,'::: Warning')
            mrate = mrate_series.loc[stock]
            drate = drate_series.loc[stock]
            avg_vol_1w = df_sub.loc[df_sub.index[-5]:,'total_volume'].mean()
            avg_vol_1m = df_sub.loc[df_sub.index[-22]:,'total_volume'].mean()
            avg_vol_3m = df_sub.loc[three_month_ago:,'total_volume'].mean()  # để tránh out-of-bound error
            max_price = maxprice_series.loc[stock]
            general_room = general_room_seires.loc[stock]
            special_room = special_room_series.loc[stock]
            total_room = total_room_series.loc[stock]
            room_on_avg_vol_3m = total_room/avg_vol_3m
            record = pd.DataFrame({
                'Stock':[stock],
                'Exchange':[exchange],
                'Tỷ lệ vay KQ (%)':[mrate],
                'Tỷ lệ vay TC (%)':[drate],
                'Giá vay / Giá TSĐB tối đa (VND)':[max_price],
                'General Room':[general_room],
                'Special Room':[special_room],
                'Total Room':[total_room],
                'Consecutive Floor Days':[n_floor],
                'Last day Volume':[volume],
                '% Last day volume / 1M Avg.':[volume_change_1m],
                '1W Avg. Volume':[avg_vol_1w],
                '1M Avg. Volume':[avg_vol_1m],
                '3M Avg. Volume':[avg_vol_3m],
                'Approved Room / Avg. Liquidity 3 months':[room_on_avg_vol_3m],
                '1M Illiquidity Days':[n_illiquidity],
            })
            records.append(record)
        else:
            print(stock,'::: Success')

    print('-------------------------')

    result_table = pd.concat(records,ignore_index=True)
    result_table.sort_values('Consecutive Floor Days',ascending=False,inplace=True)

    print('Finished!')
    print("Total execution time is: %s seconds"%np.round(time.time()-start_time,2))

    return result_table
