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
    CallMarginReport.run()

@TaskMonitor
def RMD_QuotaLimitViolationReport():
    from automation.risk_management import QuotaLimitViolationReport
    QuotaLimitViolationReport.run()

@TaskMonitor
def RMD_CheckQuotaLimitReport():
    from automation.risk_management import CheckQuotaLimitReport
    CheckQuotaLimitReport.run()

@TaskMonitor
def RMD_MarketPressureReport():
    from automation.risk_management import MarketPressureReport
    MarketPressureReport.run()

@TaskMonitor
def RMD_Top30BiggestOutstandingReport():
    from automation.risk_management import Top30BiggestOutstandingReport
    Top30BiggestOutstandingReport.run()

@TaskMonitor
def RMD_ForceSellReport():
    from automation.risk_management import ForceSellReport
    ForceSellReport.run()

@TaskMonitor
def RMD_SSCDailyReport():
    from automation.risk_management import SSCDailyReport
    SSCDailyReport.run()

