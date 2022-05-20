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

@TaskMonitor
def RMD_SaveTempData1():
    from automation.risk_management import CheckQuotaLimitReport
    CheckQuotaLimitReport.generateTempData()

@TaskMonitor
def RMD_CallMarginReport():
    from automation.risk_management import CallMarginReport
    CallMarginReport.run(dt.datetime.now())

@TaskMonitor
def RMD_QuotaLimitViolationReport():
    from automation.risk_management import QuotaLimitViolation
    QuotaLimitViolation.run(dt.datetime.now())

@TaskMonitor
def RMD_MarketPressureReport():
    from automation.risk_management import MarketPressure
    MarketPressure.run(dt.datetime.now())