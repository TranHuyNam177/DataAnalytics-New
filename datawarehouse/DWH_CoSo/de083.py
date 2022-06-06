import calendar

from request import *
from datawarehouse import *
from automation import convert_int


def update(
    update_date:dt.datetime, # lấy ngày cuối tháng
):

    """
    This function get data from Securities Service Division and insert to [DWH-CoSo].[dbo].[de083]

    :param update_date: date to update
    """

    year, month, day = (convert_int(x) for x in (update_date.year,update_date.month,update_date.day))
    insertDate = f'{year}.{month}.{day}'
    DELETE(connect_DWH_CoSo,'de083',f"WHERE [de083].[date] = '{insertDate}'")
    root = r'C:\Users\hiepdang\Shared Folder\Trading Service\Report\GiaoDichLuuKy\FilesFromVSD\SoDuChungKhoan'
    folder = join(root,year,f'THÁNG {month}',f'{day}.{month}.{year}C')
    files = [f for f in listdir(folder) if f[:2] in ('BO','DC','HN','HO','UP') and f.split('.')[-1]=='csv']
    for file in files:
        df = pd.read_csv(
            join(folder,file),
            delimiter=';',
            usecols=[10,13,14],
            header=None,
            names=['account_code','amount','ticker']
        )
        df.insert(0,'date',dt.datetime(int(year),int(month),int(day)))
        df = df[[
            'date',
            'account_code',
            'ticker',
            'amount',
        ]]
        BATCHINSERT(connect_DWH_CoSo,'de083',df)



def init():

    """
    This function get ALL data from Securities Service Division and insert to [DWH-CoSo].[dbo].[de083]
    Used to initiate data into the empy [DWH-CoSo].[dbo].[de083]
    """

    root = r'C:\Users\hiepdang\Shared Folder\Trading Service\Report\GiaoDichLuuKy\FilesFromVSD\SoDuChungKhoan'
    for yearFolder in listdir(root):
        for monthFolder in listdir(join(root,yearFolder)):
            for dateFolder in listdir(join(root,yearFolder,monthFolder)):
                year = dateFolder.split('.')[2].replace('C','')
                month = dateFolder.split('.')[1]
                day = dateFolder.split('.')[0]
                update(f'{year}-{month}-{day}')


def monthly_update(
    year: int,
    month:int,
):

    """
    This function update data for last month (GDLK must copy/paste data to specied folder first)

    :param month: month to update
    :param year: year to update
    """

    start_day, end_day = calendar.monthrange(year,month)
    for d in range(start_day,end_day+1):
        update_date = dt.datetime(year,month,d)
        date_string = update_date.strftime('%Y-%m-%d')
        try:
            update(update_date)
            print(f'Insert thành công ngày {date_string}')
        except pyodbc.DataError: # Catch lỗi các ngày không có thực như: 29/02, 31/04, etc.
            print(f'Thực tế không có ngày {date_string}')
        except FileNotFoundError: # Không phải ngày nào cũng có số
            print(f'VSD không có file ngày {date_string}')
