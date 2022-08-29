from implementation import TaskMonitor


@TaskMonitor
def DWH_CoSo_Update_Today():
    from datawarehouse.DWH_CoSo import SYNCTODAY
    from datawarehouse import CHECKBATCH
    from request import connect_DWH_CoSo
    import datetime as dt
    import time
    now = dt.datetime.now()
    weekDay = now.weekday() + 2 # Thứ trong tuần
    if weekDay in (7,8): # Thứ 7, CN -> run as called
        SYNCTODAY()
    else: # ngày thường
        if now.hour < 13: # buổi trưa
            pass
        else:
            if now.hour < 19:
                batchType = 1
            else:
                batchType = 2
            while True:
                if CHECKBATCH(connect_DWH_CoSo,batchType):
                    break
                time.sleep(30)
        SYNCTODAY()

@TaskMonitor
def DWH_CoSo_Update_BackDate():
    from datawarehouse.DWH_CoSo import SYNCBACKDATE
    import datetime as dt
    hour = dt.datetime.now().hour
    if 22 <= hour <= 24 or 0 <= hour <= 5:
        days = 5
    else:
        days = 1
    for day in range(1,days+1): # 1,2,3,...,day
        SYNCBACKDATE(day)

# không dùng @TaskMonitor vì hàm này đã có sẵn một lớp Monitor rồi
def DWHCoSo_InternetBanking_EOD(bank):

    from datawarehouse.DWH_CoSo.InternetBanking import runBankDepositBalance
    from datawarehouse.DWH_CoSo.InternetBanking import runBankCurrentBalance
    from datawarehouse.DWH_CoSo.InternetBanking import runBankTransactionHistory
    from datawarehouse.DWH_CoSo.InternetBanking import runBankLoanBalance
    from datawarehouse.DWH_CoSo.InternetBanking import BankFailedException
    from automation.finance import BIDV,EIB,IVB,VTB,VCB,OCB,TCB,ESUN,FIRST,FUBON,HUANAN,MEGA,SINOPAC
    import datetime as dt
    now = dt.datetime.now()
    hour = now.hour
    today = dt.datetime.today().replace(hour=0,minute=0,second=0,microsecond=0)

    if hour >= 8: # thời điểm chạy không phải đầu giờ sáng (gửi mail CAPTCHA đến hiepdang@phs.vn -> debug mode ON)
        dataDate = today
        debug = True
    else: # chạy sáng sớm (gửi mail CAPTCHA đến all -> debug mode OFF)
        dataDate = today - dt.timedelta(days=1)
        debug = False

    # Tạo bankObject (Đã login sẵn)
    if bank == 'BIDV':
        bankObject = BIDV(debug).Login()
    elif bank == 'EIB':
        bankObject = EIB(debug).Login()
    elif bank == 'IVB':
        bankObject = IVB(debug).Login()
    elif bank == 'VTB':
        bankObject = VTB(debug).Login()
    elif bank == 'VCB':
        bankObject = VCB(debug).Login()
    elif bank == 'OCB':
        bankObject = OCB(debug).Login()
    elif bank == 'TCB':
        bankObject = TCB(debug).Login()
    elif bank == 'ESUN':
        bankObject = ESUN(debug).Login()
    elif bank == 'FIRST':
        bankObject = FIRST(debug).Login()
    elif bank == 'FUBON':
        bankObject = FUBON(debug).Login()
    elif bank == 'HUANAN':
        bankObject = HUANAN(debug).Login()
    elif bank == 'MEGA':
        bankObject = MEGA(debug).Login()
    elif bank == 'SINOPAC':
        bankObject = SINOPAC(debug).Login()
    else:
        raise ValueError(f'Invalid bank name: {bank}')

    # Chạy các hàm trên bankObject đã tạo
    for func in (runBankDepositBalance,runBankCurrentBalance,runBankTransactionHistory,runBankLoanBalance):
        for _ in range(2): # chạy lại tối đa 2 lần
            try:
                if func in (runBankDepositBalance,runBankLoanBalance):
                    bankObject = func(bankObject)
                else:
                    bankObject = func(bankObject,dataDate,dataDate)
                break
            except (BankFailedException,): # any error orcurs -> send traceback email and move on.
                pass

    # Terminate Object
    del bankObject

@TaskMonitor
def DWH_PhaiSinh_Update_Today():
    from datawarehouse.DWH_PhaiSinh import SYNCTODAY
    from datawarehouse import CHECKBATCH
    from request import connect_DWH_PhaiSinh
    import datetime as dt
    import time
    now = dt.datetime.now()
    weekDay = now.weekday() + 2 # Thứ trong tuần
    if weekDay in (7,8): # Thứ 7, CN -> run as called
        SYNCTODAY()
    else: # ngày thường
        if now.hour < 13: # buổi trưa
            pass
        else:
            while True:
                if CHECKBATCH(connect_DWH_PhaiSinh,2):
                    break
                time.sleep(30)
        SYNCTODAY()

@TaskMonitor
def DWH_PhaiSinh_Update_BackDate():
    from datawarehouse.DWH_PhaiSinh import SYNCBACKDATE
    import datetime as dt
    hour = dt.datetime.now().hour
    if 22 <= hour <= 24 or 0 <= hour <= 5:
        days = 5
    else:
        days = 1
    for day in range(1,days+1): # 1,2,3,...,day
        SYNCBACKDATE(day)

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

@TaskMonitor
def DWHThiTruongUpdate_SecuritiesInfoVSD(): # mỗi ngày update 1000 ID -> một tháng 30,000
    import datetime as dt
    from datawarehouse.DWH_ThiTruong.SecuritiesInfoVSD import run as Update_SecuritiesInfoVSD
    day = max(dt.datetime.now().day,30)
    startID = (day-1) * 4000
    endID = day * 4000
    Update_SecuritiesInfoVSD(startID,endID)

@TaskMonitor
def DWHCoSoUpdate_DanhMucChoVayMargin():
    import datetime as dt
    from datawarehouse.DWH_CoSo.DanhMucChoVayMargin import run as Update_DanhMucChoVayMargin
    Update_DanhMucChoVayMargin()

@TaskMonitor
def DWH_NotifySyncStatusToday(db):
    from datawarehouse import NOTIFYSYNCSTATUSTODAY
    NOTIFYSYNCSTATUSTODAY(db)

@TaskMonitor
def DWH_NotifySyncStatusBackDate(db):
    from datawarehouse import NOTIFYSYNCSTATUSBACKDATE
    NOTIFYSYNCSTATUSBACKDATE(db)

@TaskMonitor
def DWHBaseUpdate_Employee():
    from datawarehouse.DWH_Base.Employee import run
    run()

@TaskMonitor
def DWHThiTruongUpdate_GiaThanhToanPhaiSinhVSD():
    from datawarehouse.DWH_ThiTruong.GiaThanhToanPhaiSinhVSD import run
    import datetime as dt
    today = dt.datetime.now()
    run(today,today)


@TaskMonitor
def DWHThiTruongUpdate_TinToChucPhatHanhVSD():
    from datawarehouse.DWH_ThiTruong.TinToChucPhatHanhVSD import run
    run(100)


