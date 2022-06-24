from implementation import TaskMonitor


@TaskMonitor
def DWH_CoSo_Update_Today():
    from datawarehouse.DWH_CoSo import SYNCTODAY
    from datawarehouse import CHECKBATCH
    from request import connect_DWH_CoSo
    import datetime as dt
    now = dt.datetime.now()
    if now.hour < 15:
        SYNCTODAY() # cho phép chạy buổi trưa
    else:
        if now.hour >= 19:
            batchType = 2
        else:
            batchType = 1
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
def DWHCoSo_InternetBanking_EOD(bank,func='all'):

    from datawarehouse.DWH_CoSo import InternetBanking
    from automation.finance import BIDV,EIB,IVB,VTB,VCB,OCB,TCB
    import datetime as dt
    now = dt.datetime.now()
    hour = now.hour
    today = dt.datetime.today().replace(hour=0,minute=0,second=0,microsecond=0)

    if hour >= 12: # chạy buổi tối (gửi mail CAPTCHA đến hiepdang@phs.vn -> debug mode ON)
        dataDate = today
        debug = True
    else: # chạy sáng sớm (gửi mail CAPTCHA đến all -> debug mode OFF)
        dataDate = today - dt.timedelta(days=1)
        debug = False

    # Check arguments
    checkPairs = {
        'runBankCurrentBalance': ['BIDV','EIB','IVB','VTB','VCB','OCB','TCB',],
        'runBankDepositBalance': ['BIDV','IVB','VTB','VCB','OCB',],
    }
    if func == 'all':
        checkList = []
        for f in checkPairs.keys():
            checkList.extend(checkPairs[f])
        checkList = set(checkList)
        if bank not in checkList:
            raise ValueError(f'One of the function does not apply on {bank}')
    else:
        if func not in checkPairs.keys():
            raise ValueError(f'Invalid func: {func}')
        if bank not in checkPairs[func]:
            raise ValueError(f'Function {func} does not apply on {bank}')

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
    else:
        raise ValueError(f'Invalid bank name: {bank}')

    # Chạy các hàm trên bankObject đã tạo
    if func == 'runBankCurrentBalance':
        InternetBanking.runBankCurrentBalance(bankObject,dataDate,dataDate)
    elif func == 'runBankDepositBalance':
        InternetBanking.runBankDepositBalance(bankObject)
    elif func == 'all':
        bankObject = InternetBanking.runBankDepositBalance(bankObject) # chạy Deposit trước, giữ lại bankObject để dùng lại
        InternetBanking.runBankCurrentBalance(bankObject,dataDate,dataDate) # chạy Current sau
    else:
        raise ValueError(f'Invalid func {func}')
    # Terminate Object
    del bankObject

# không dùng @TaskMonitor vì hàm này đã có sẵn một lớp Monitor rồi
def DWHCoSo_InternetBanking_RT(bank,func='runBankTransactionHistory'):

    from automation.finance import BIDV,EIB,IVB,VTB,VCB,OCB,TCB
    from datawarehouse.DWH_CoSo import InternetBanking
    import datetime as dt

    # Tạo bankObject (Đã login sẵn)
    if bank == 'BIDV':
        bankObject = BIDV(True).Login()
    elif bank == 'EIB':
        bankObject = EIB(True).Login()
    elif bank == 'IVB':
        bankObject = IVB(True).Login()
    elif bank == 'VTB':
        bankObject = VTB(True).Login()
    elif bank == 'VCB':
        bankObject = VCB(True).Login()
    elif bank == 'OCB':
        bankObject = OCB(True).Login()
    elif bank == 'TCB':
        bankObject = TCB(True).Login()
    else:
        raise ValueError(f'Invalid bank name: {bank}')

    if func == 'runBankTransactionHistory':
        today = dt.datetime.today()
        while True:
            bankObject = InternetBanking.runBankTransactionHistory(bankObject,today,today,True)
            if dt.datetime.now().time() > dt.time(16,50,0):
                break
            time.sleep(60) # quét 1 phút/lần
    # Terminate Object
    del bankObject

@TaskMonitor
def DWH_PhaiSinh_Update_Today():
    from datawarehouse.DWH_CoSo import SYNCTODAY
    from datawarehouse import CHECKBATCH
    from request import connect_DWH_PhaiSinh
    import datetime as dt
    now = dt.datetime.now()
    if now.hour < 15:
        SYNCTODAY() # cho phép chạy buổi trưa
    else:
        while True:
            if CHECKBATCH(connect_DWH_PhaiSinh):
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
    startID = (day-1) * 1000
    endID = day * 1000
    Update_SecuritiesInfoVSD(startID,endID)

@TaskMonitor
def DWHCoSoUpdate_DanhMucChoVayMargin():
    import datetime as dt
    from datawarehouse.DWH_CoSo.DanhMucChoVayMargin import run as Update_DanhMucChoVayMargin
    Update_DanhMucChoVayMargin()
