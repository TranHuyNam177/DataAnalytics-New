from implementation import TaskMonitor
from request.stock import *

@TaskMonitor
def RMD_BaoCaoGiaHoaVonDanhMuc():
    from automation.risk_management import BaoCaoGiaHoaVonDanhMuc
    BaoCaoGiaHoaVonDanhMuc.run()

@TaskMonitor
def RMD_TinChungKhoan():
    from news_analysis import classify
    t = dt.datetime.now().time()
    if t < dt.time(hour=12):
        classify.FilterNewsByKeywords(hours=16)
    else:
        classify.FilterNewsByKeywords(hours=8)

