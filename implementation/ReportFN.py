from implementation import TaskMonitor


@TaskMonitor
def FN_BaoCaoTienGuiThanhToan():
    from automation.finance import BaoCaoSoDuTienGuiThanhToan
    import datetime as dt
    # chạy buổi sáng ngày hôm sau
    BaoCaoSoDuTienGuiThanhToan.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def FN_BaoCaoTienGuiKyHan():
    from automation.finance import BaoCaoSoDuTienGuiKyHan
    import datetime as dt
    # chạy buổi sáng ngày hôm sau
    BaoCaoSoDuTienGuiKyHan.run(dt.datetime.now()-dt.timedelta(days=1))