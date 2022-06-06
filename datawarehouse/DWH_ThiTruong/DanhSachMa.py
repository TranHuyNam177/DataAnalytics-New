from request import *
from datawarehouse import BATCHINSERT, DELETE
from news_collector import scrape_ticker_by_exchange

def update():

    """
    This function updates data to table [DWH-ThiTruong].[dbo].[DanhSachMa]
    """

    table = scrape_ticker_by_exchange.run(True).reset_index()
    d = dt.datetime.now()
    table.insert(0,'Date',dt.datetime(d.year,d.month,d.day))
    table = table.rename({'ticker':'Ticker','exchange':'Exchange'},axis=1)
    DELETE(connect_DWH_ThiTruong,"DanhSachMa",f"""WHERE [Date] = '{d.strftime("%Y-%m-%d")}'""")
    BATCHINSERT(connect_DWH_ThiTruong,'DanhSachMa',table)



