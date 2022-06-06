from request import *
from datawarehouse import BATCHINSERT, DROPDUPLICATES
from news_analysis import get_news

def update(hours):

    """
    This function updates data to table [DWH-ThiTruong].[dbo].[TinChungKhoan]

    :param hours: number of hours to read news historically (đọc lùi bao nhiêu giờ)
    """

    Table1 = get_news.cafef(hours).run()
    Table2 = get_news.ndh(hours).run()
    Table3 = get_news.vietstock(hours).run()
    Table4 = get_news.tinnhanhchungkhoan(hours).run()

    newsTable = pd.concat([Table1,Table2,Table3,Table4])

    BATCHINSERT(connect_DWH_ThiTruong,'TinChungKhoan',newsTable)
    DROPDUPLICATES(connect_DWH_ThiTruong,'TinChungKhoan','URL') # xóa dòng có URL trùng nhau

