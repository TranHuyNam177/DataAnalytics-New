from implementation import TaskMonitor


@TaskMonitor
def DWH_CoSo_Update_Today():
    from datawarehouse.DWH_CoSo import UPDATE
    UPDATE()

@TaskMonitor
def DWH_CoSo_Update_BackDate():
    from datawarehouse.DWH_CoSo import UPDATEBACKDATE
    import datetime as dt
    hour = dt.datetime.now().hour
    if 22 <= hour <= 24 or 0 <= hour <= 5:
        days = 5
    else:
        days = 1
    for day in range(1,days+1): # 1,2,3,...,day
        UPDATEBACKDATE(day)

# không dùng @TaskMonitor vì hàm này đã có sẵn một lớp Monitor rồi
def DWHCoSo_InternetBanking(bank):
    from datawarehouse.DWH_CoSo import InternetBanking
    import datetime as dt
    today = dt.datetime.today()
    if today.hour >= 12: # chạy buổi tối (gửi mail CAPTCHA đến hiepdang@phs.vn -> debug mode ON)
        debug = True
    else: # chạy sáng sớm (gửi mail CAPTCHA đến all -> debug mode OFF)
        debug = False
    InternetBanking.run(bank,today-dt.timedelta(days=1),today-dt.timedelta(days=1),debug)

@TaskMonitor
def DWH_PhaiSinh_Update_Today():
    from datawarehouse.DWH_PhaiSinh import UPDATE
    UPDATE()

@TaskMonitor
def DWH_PhaiSinh_Update_BackDate():
    from datawarehouse.DWH_PhaiSinh import UPDATEBACKDATE
    import datetime as dt
    hour = dt.datetime.now().hour
    if 22 <= hour <= 24 or 0 <= hour <= 5:
        days = 5
    else:
        days = 1
    for day in range(1,days+1): # 1,2,3,...,day
        UPDATEBACKDATE(day)

@TaskMonitor
def DWH_ThiTruong_Update_DanhSachMa():
    from datawarehouse.DWH_ThiTruong.DanhSachMa import update as Update_DanhSachMa
    Update_DanhSachMa()

@TaskMonitor
def DWHThiTruongUpdate_DuLieuGiaoDichNgay():
    from datawarehouse.DWH_ThiTruong.DuLieuGiaoDichNgay import run as Update_DuLieuGiaoDichNgay
    import datetime as dt
    today = dt.datetime.now().replace(hour=0,minute=0,second=0,microsecond=0)
    Update_DuLieuGiaoDichNgay(today,today)

@TaskMonitor
def DWHThiTruongUpdate_TinChungKhoan():
    from datawarehouse.DWH_ThiTruong.TinChungKhoan import update as Update_TinChungKhoan
    Update_TinChungKhoan(24)

