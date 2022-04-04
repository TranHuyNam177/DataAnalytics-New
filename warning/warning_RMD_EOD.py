from request.stock import *


def run(
    mlist:bool=True,
    exchange:str='HOSE',
    segment:str='all'
) \
        -> pd.DataFrame:
    """
    This method returns list of tickers that up ceil (fc_type='ceil')
    or down floor (fc_type='floor') in a given exchange of a given segment
    in n consecutive trading days

    :param mlist: report margin list only (True) or not (False)
    :param exchange: allow values in fa.exchanges. Do not allow 'all'
    :param segment: allow values in fa.segments or 'all'.
    For mlist=False only, if mlist=True left as default

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

    if mlist is True:
        full_tickers = internal.mlist(exchanges=[exchange])
    else:
        full_tickers = fa.tickers(segment,exchange)

    records = []
    for ticker in full_tickers:
        df = ta.hist(ticker,fromdate=since,todate=now)[['trading_date','ref','close','total_volume']]
        df.set_index('trading_date',drop=True,inplace=True)
        df[['ref','close']] = (df[['ref','close']]*1000).round(0).astype(np.int64)
        df['total_volume'] = df['total_volume'].astype(np.int64)
        df['floor'] = df['ref'].apply(
            fc_price,
            price_type='floor',
            exchange=exchange
        )
        # giam san lien tiep
        n_floor = 0
        for i in range(df.shape[0]):
            condition1 = (df.loc[df.index[-i-1:],'floor']==df.loc[df.index[-i-1:],'close']).all()
            condition2 = (df.loc[df.index[-i-1:],'floor']!=df.loc[df.index[-i-1:],'ref']).all()
            # the second condition is to ignore trash tickers whose price
            # less than 1000 VND (a single price step equivalent to
            # more than 7%(HOSE), 10%(HNX), 15%(UPCOM))
            if condition1 and condition2:
                n_floor += 1
            else:
                break
        # mat thanh khoan trong 1 thang
        avg_vol_1m = df.loc[df.index[-22]:,'total_volume'].mean()
        n_illiquidity = (df.loc[df.index[-22]:,'total_volume']<avg_vol_1m).sum()
        # thanh khoan ngay gan nhat so voi thanh khoan trung binh 1 thang
        volume = df.loc[df.index[-1],'total_volume']
        volume_change_1m = volume/avg_vol_1m-1

        n_illiquidity_bmk = 1
        n_floor_bmk = 1

        condition1 = n_floor>=n_floor_bmk
        condition2 = n_illiquidity>=n_illiquidity_bmk

        if condition1 or condition2:
            print(ticker,'::: Warning')
            mrate = mrate_series.loc[ticker]
            drate = drate_series.loc[ticker]
            avg_vol_1w = df.loc[df.index[-5]:,'total_volume'].mean()
            avg_vol_1m = df.loc[df.index[-22]:,'total_volume'].mean()
            avg_vol_3m = df.loc[three_month_ago:,'total_volume'].mean()  # để tránh out-of-bound error
            max_price = maxprice_series.loc[ticker]
            general_room = general_room_seires.loc[ticker]
            special_room = special_room_series.loc[ticker]
            total_room = total_room_series.loc[ticker]
            room_on_avg_vol_3m = total_room/avg_vol_3m
            record = pd.DataFrame({
                'Stock':[ticker],
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
            print(ticker,'::: Success')

        print('-------------------------')

    result_table = pd.concat(records,ignore_index=True)
    result_table.sort_values('Consecutive Floor Days',ascending=False,inplace=True)

    print('Finished!')
    print("Total execution time is: %s seconds"%np.round(time.time()-start_time,2))

    return result_table
