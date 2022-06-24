from request import *
from datawarehouse import BATCHINSERT, DELETE

def run(
    run_time=dt.datetime.now()
):
    colNames = [
        'MaCK',
        'TyLeVayKyQuy',
        'TyLeTaiSanDamBaoKyQuy',
        'TyLeVayTheChap',
        'TyLeTaiSanDamBaoTheChap',
        'GiaVayGiaTaiSanDamBaoToiDa',
        'RoomChung',
        'RoomRieng',
        'TongRoom',
        'SanGiaoDich',
    ]
    sqlDate = run_time.replace(hour=0,minute=0,second=0,microsecond=0)
    sqlDateString = sqlDate.strftime('%Y-%m-%d')
    if sqlDate < dt.datetime(2021,9,24):
        folderPATH = fr'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\Information Public\Intranet\Intranet\{sqlDate.year}'
        usecols = 'B:I'
        colNames.remove('RoomRieng')
        colNames.remove('TongRoom')
    else:
        folderPATH = fr'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Margin List'
        usecols = 'B:K'
    files = [f for f in listdir(folderPATH) if 'copy' not in f.lower()]
    fileDates = [dt.datetime.strptime(f.split('_')[1][:10],'%d.%m.%Y') for f in files]
    strippedDates = [d for d in fileDates if d <= sqlDate]
    if strippedDates:
        latestDate = max(strippedDates).strftime('%d.%m.%Y')
        latestFile = first(files, lambda file: latestDate in file)
        filePATH = join(folderPATH,latestFile)
        table = pd.read_excel(filePATH,usecols=usecols,names=colNames).dropna(how='any')
        if table.empty:
            raise IOError("Không thể đọc dữ liệu từ file của RMD")
        for col in colNames:
            if 'TyLe' in col:
                table = table.loc[table[col].isin(range(0,101))]
            if col in ('GiaVayGiaTaiSanDamBaoToiDa','RoomChung'):
                table = table.loc[table[col].isin(range(0,10000000))]
    else:
        table = pd.read_sql(
            f"""
            SELECT {','.join(colNames)} FROM [DanhMucChoVayMargin] 
            WHERE [Ngay] = (SELECT MAX([Ngay]) FROM [DanhMucChoVayMargin] WHERE [Ngay] <= '{sqlDateString}')
            """,
            connect_DWH_CoSo,
        )
    table.insert(0,'Ngay',sqlDate)
    DELETE(connect_DWH_CoSo,'DanhMucChoVayMargin',f"WHERE [Ngay] = '{sqlDateString}'")
    BATCHINSERT(connect_DWH_CoSo,'DanhMucChoVayMargin',table)

