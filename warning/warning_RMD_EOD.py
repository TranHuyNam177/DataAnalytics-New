import datetime as dt
import pandas as pd
import time
from datawarehouse import connect_DWH_CoSo, BDATE
from function import fc_price


def run(
    run_time=dt.datetime.now()
) -> pd.DataFrame:

    start_time = time.time()
    _rundate = run_time.strftime('%Y-%m-%d')
    # xét 3 tháng gần nhất
    _3M_ago = BDATE(_rundate,-66)
    _12D_ago = BDATE(_rundate,-12)
    _1D_ago = BDATE(_rundate,-1)
    
    table = pd.read_sql(
        f"""
        WITH
        [DanhMuc] AS (
            SELECT
                [MaCK],
                [SanGiaoDich],
                [TyLeVayKyQuy],
                [TyLeVayTheChap],
                [GiaVayGiaTaiSanDamBaoToiDa],
                [RoomChung],
                [RoomRieng],
                [TongRoom]
            FROM [DWH-CoSo].[dbo].[DanhMucChoVayMargin] 
            WHERE [Ngay] = '{_1D_ago}'
        ),
        [RawTable] AS (
            SELECT
		        [ThiTruong].[Date],
                [DanhMuc].*,
                [ThiTruong].[Ref] * 1000 [RefPrice],
                [ThiTruong].[Close] * 1000 [ClosePrice],
                [ThiTruong].[Volume],
                AVG([ThiTruong].[Volume]) OVER (PARTITION BY [Ticker] ORDER BY [Date] ROWS BETWEEN 65 PRECEDING AND CURRENT ROW) [AvgVolume3M],
                AVG([ThiTruong].[Volume]) OVER (PARTITION BY [Ticker] ORDER BY [Date] ROWS BETWEEN 21 PRECEDING AND CURRENT ROW) [AvgVolume1M],	
                AVG([ThiTruong].[Volume]) OVER (PARTITION BY [Ticker] ORDER BY [Date] ROWS BETWEEN 4 PRECEDING AND CURRENT ROW) [AvgVolume1W],
                CASE
                    WHEN [ThiTruong].[Volume] < AVG([ThiTruong].[Volume]) OVER (PARTITION BY [Ticker] ORDER BY [Date] ROWS BETWEEN 21 PRECEDING AND CURRENT ROW)
                        THEN 1
                    ELSE 0
                END [FlagIlliquidity1M]
            FROM [DanhMuc]
            LEFT JOIN [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay] [ThiTruong]
                ON [DanhMuc].[MaCK] = [ThiTruong].[Ticker]
            WHERE [ThiTruong].[Date] BETWEEN '{_3M_ago}' AND '{_rundate}' 
                AND [ThiTruong].[Ref] IS NOT NULL -- bỏ ngày nghỉ
        )
        SELECT 
            [RawTable].*,
            CASE WHEN [AvgVolume3M] <> 0 THEN [RawTable].[TongRoom] / [RawTable].[AvgVolume3M] ELSE 0 END [ApprovedRoomOnAvgVolume3M],
            CASE WHEN [AvgVolume3M] <> 0 THEN [RawTable].[Volume] / [RawTable].[AvgVolume1M] - 1 ELSE 0 END [LastDayVolumeOnAvgVolume1M],
            SUM([RawTable].[FlagIlliquidity1M]) OVER (PARTITION BY [MaCK] ORDER BY [Date] ROWS BETWEEN 21 PRECEDING AND CURRENT ROW) [CountIlliquidity1M]
        FROM [RawTable]
        WHERE [RawTable].[AvgVolume1M] <> 0 -- 1 tháng vừa rồi có giao dịch (để đảm bảo không bị lỗi chia cho 0)
        ORDER BY [RawTable].[MaCK], [RawTable].[Date]
        """,
        connect_DWH_CoSo
    )
    table = table.drop('FlagIlliquidity1M',axis=1)
    table = table.loc[table['Date']>=dt.datetime.strptime(_12D_ago,'%Y-%m-%d')] # 12 phiên gần nhất
    table['FloorPrice'] = table.apply(
        lambda x: fc_price(x['RefPrice'],'floor',x['SanGiaoDich']),
        axis=1,
    )
    records = []
    for stock in table['MaCK'].unique():
        subTable = table[table['MaCK']==stock]
        # giam san lien tiep
        n_floor = 0
        for i in range(subTable.shape[0]):
            condition1 = (subTable.loc[subTable.index[-i-1:],'FloorPrice']==subTable.loc[subTable.index[-i-1:],'ClosePrice']).all()
            condition2 = (subTable.loc[subTable.index[-i-1:],'FloorPrice']!=subTable.loc[subTable.index[-i-1:],'RefPrice']).all()
            # condition2 is to ignore trash tickers in which a single price step leads to floor price
            if condition1 and condition2:
                n_floor += 1
            else:
                break
        subTable = subTable.iloc[[-1]]
        subTable.insert(subTable.shape[1],'ConsecutiveFloors',n_floor)
        records.append(subTable)
        print(stock,'::: Done')

    print('-------------------------')
    result_table = pd.concat(records,ignore_index=True)
    result_table.sort_values('ConsecutiveFloors',ascending=False,inplace=True)
    nameMapper = {
        'MaCK':'Stock',
        'SanGiaoDich':'Exchange',
        'TyLeVayKyQuy':'Tỷ lệ vay KQ (%)',
        'TyLeVayTheChap':'Tỷ lệ vay TC (%)',
        'GiaVayGiaTaiSanDamBaoToiDa':'Giá vay / Giá TSĐB tối đa (VND)',
        'RoomChung':'General Room',
        'RoomRieng':'Special Room',
        'TongRoom':'Total Room',
        'ConsecutiveFloors':'Consecutive Floor Days',
        'Volume':'Last day Volume',
        'LastDayVolumeOnAvgVolume1M':'% Last day volume / 1M Avg.',
        'AvgVolume1W':'1W Avg. Volume',
        'AvgVolume1M':'1M Avg. Volume',
        'AvgVolume3M':'3M Avg. Volume',
        'ApprovedRoomOnAvgVolume3M':'Approved Room / Avg. Liquidity 3 months',
        'CountIlliquidity1M':'1M Illiquidity Days',
    }
    result_table = result_table.reindex(nameMapper.keys(),axis=1)
    result_table = result_table.rename(nameMapper,axis=1)

    print('Finished!')
    print(f"Total execution time is: {round(time.time()-start_time,2)} seconds")

    return result_table
