from request import *
from datawarehouse import BATCHINSERT, DELETE

def run(
    run_time=dt.datetime.now()
):

    folderPATH = fr'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\Information Public\Intranet\Intranet\{run_time.year}'
    files = listdir(folderPATH)
    fileDates = [dt.datetime.strptime(f.split('_')[1][:10],'%d.%m.%Y') for f in files]
    strippedDates = [d for d in fileDates if d <= run_time]
    latestDate = max(strippedDates).strftime('%d.%m.%Y')
    latestFile = first(files, lambda file: latestDate in file)
    filePATH = join(folderPATH,latestFile)

    colNames = [
        'MaCK',
        'TyLeVayKyQuy',
        'TyLeTaiSanDamBaoKyQuy',
        'TyLeVayTheChap',
        'TyLeTaiSanDamBaoTheChap',
        'GiaVayGiaTaiSanDamBaoToiDa',
        'RoomChung',
        'SanGiaoDich',
    ]
    table = pd.read_excel(filePATH,usecols='B:I',names=colNames,).dropna(how='any')
    if table.empty:
        raise IOError("Không thể đọc dữ liệu từ file của RMD")
    for col in colNames:
        if 'TyLe' in col:
            table = table.loc[table[col].isin(range(0,101))]
        if col in ('GiaVayGiaTaiSanDamBaoToiDa','RoomChung'):
            table = table.loc[table[col].isin(range(0,10000000))]
    sqlDate = run_time.replace(hour=0,minute=0,second=0,microsecond=0)
    table.insert(0,'Ngay',sqlDate)
    sqlDateString = sqlDate.strftime('%Y-%m-%d')
    DELETE(connect_DWH_CoSo,'DanhMucChoVayMargin',f"WHERE [Ngay] = {sqlDateString}")
    BATCHINSERT(connect_DWH_CoSo,'DanhMucChoVayMargin',table)
