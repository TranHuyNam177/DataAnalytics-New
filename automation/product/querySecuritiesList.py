import os
import argparse
import pandas as pd
import datetime as dt
import pyodbc

# DWH-CoSo Database Information
driver_DWH_CoSo = '{SQL Server}'
server_DWH_CoSo = 'SRV-RPT'
db_DWH_CoSo = 'DWH-CoSo'
connect_DWH_CoSo = pyodbc.connect(
    f'Driver={driver_DWH_CoSo};'
    f'Server={server_DWH_CoSo};'
    f'Database={db_DWH_CoSo};'
    f'uid=namtran;'
    f'pwd=nam!tran@2021'
)

def createJSON(inputDate, outputPath):
    date = dt.datetime.strptime(inputDate, "%Y%m%d").strftime("%Y-%m-%d")
    query020004 = pd.read_sql(
        f"""
        WITH
        [engName] AS (
            SELECT
                [infoVSD].[Ticker],
                [infoVSD].[Value] [tenTiengAnh]
            FROM [DWH-ThiTruong].[dbo].[SecuritiesInfoVSD] [infoVSD]
            WHERE Attribute = 'Issuers name'
        )
        , [tcph] AS (
            SELECT
                [infoVSD].[Ticker],
                [infoVSD].[Value] [toChucPhatHanh]
            FROM [DWH-ThiTruong].[dbo].[SecuritiesInfoVSD] [infoVSD]
            WHERE Attribute = N'Tên TCPH'
        )
        , [lastDateResult] AS (
            SELECT
                MAX([date]) [date],
                [securities_name]
            FROM [020004_securities]
            WHERE [date] <= '{date}'
            GROUP BY [securities_name]
        )
        SELECT
        DISTINCT -- distinct vì có trường hợp 1 ngày có 2 mã giống nhau nhưng khác type, exchange
            [020004_securities].[securities_name] [Ma CK],
            [tcph].[toChucPhatHanh] [To Chuc Phat Hanh],
            [engName].[tenTiengAnh] [Ten Tieng Anh]
        FROM [020004_securities]
        INNER JOIN [lastDateResult] ON [lastDateResult].[date] = [020004_securities].[date]
        AND [lastDateResult].[securities_name] = [020004_securities].[securities_name]
        LEFT JOIN [engName] ON [engName].[Ticker] = [020004_securities].[securities_name]
        LEFT JOIN [tcph] ON [tcph].[Ticker] = [020004_securities].[securities_name]
        WHERE 
            [020004_securities].[exchange] <> 'WFT'
            AND [020004_securities].[type] <> N'Quyền chọn'
            AND [020004_securities].[securities_name] <> 'PHS_SIC'
        ORDER BY [020004_securities].[securities_name]
        """,
        connect_DWH_CoSo
    )
    jsonFile = query020004.to_json(
        os.path.join(outputPath, "dataMaCK.json"),
        orient="records",
        lines=True,
        force_ascii=False
    )
    return jsonFile


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--inputDate", required=True, help="input date with format: YYYYmmdd")
    ap.add_argument("-o", "--outputPath", required=True, help="path to output json file")
    args = vars(ap.parse_args())
    createJSON(args["inputDate"], args["outputPath"])
