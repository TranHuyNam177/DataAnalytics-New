from automation.finance import *
from automation.finance import BankTransferBalance

# không CAPTCHA
def runTCB(bankObject,fromDate,toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransferBalance.runTCB(bankObject,fromDate,toDate)
    except bankObject.ignored_exceptions:
        print(bankObject.ignored_exceptions)
        xpath = "//*[contains(text(),'vui lòng đăng nhập lại')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH,xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = TCB(True).Login()
            print('Done login')
    time.sleep(1)
    runTCB(bankObject,fromDate,toDate)

# không CAPTCHA
def runVTB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransferBalance.runVTB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        print(bankObject.ignored_exceptions)
        xpath = "//*[contains(text(), 'Lưu tên đăng nhập')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = VTB(True).Login()
            print('Done login')
    time.sleep(1)
    runVTB(bankObject, fromDate, toDate)

# không CAPTCHA
def runOCB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransferBalance.runOCB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        print(bankObject.ignored_exceptions)
        xpath = "//*[contains(text(),'Đăng nhập OMNI Doanh nghiệp')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = OCB(True).Login()
            print('Done login')
    time.sleep(1)
    runOCB(bankObject, fromDate, toDate)

# không CAPTCHA
def runEIB(bankObject, fromDate, toDate):
    """
        :param bankObject: Bank Object (đã login)
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngày kết thúc lấy dữ liệu
        """
    try:
        transactionTable = BankTransferBalance.runEIB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions or Exception:
        print(bankObject.ignored_exceptions, Exception)
        xpath = "//*[contains(text(),'Phiên làm việc hết hiệu lực')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        xpath = "//*[contains(text(), 'Đang tải dữ liệu...')]"
        reLoginNotices_2 = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices or reLoginNotices_2:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = EIB(True).Login()
            print('Done login')
    time.sleep(1)
    runEIB(bankObject, fromDate, toDate)

# có CAPTCHA
def runVCB(bankObject, fromDate, toDate):
    """
        :param bankObject: Bank Object (đã login)
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngày kết thúc lấy dữ liệu
        """
    try:
        transactionTable = BankTransferBalance.runVCB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        print(bankObject.ignored_exceptions)
        time.sleep(3)
        xpath = '//*[@value="Đăng nhập"]'
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        xpath = "//span[contains(text(), 'This site can’t be reached')]"
        reLoginNotices_2 = bankObject.driver.find_elements(By.XPATH, xpath) # trang web hiện This site can't be reached
        if reLoginNotices or reLoginNotices_2:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = VCB(True).Login()
            print('Done login')

    time.sleep(1)
    runVCB(bankObject, fromDate, toDate)

# có CAPTCHA
def runIVB(bankObject, fromDate, toDate):
    """
        :param bankObject: Bank Object (đã login)
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransferBalance.runIVB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        print(bankObject.ignored_exceptions)
        time.sleep(3)
        xpath = "//h2[contains(text(), 'Phiên giao dịch hết hạn')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = IVB(True).Login()
            print('Done login')

    time.sleep(1)
    runIVB(bankObject, fromDate, toDate)

# có CAPTCHA
def runBIDV(bankObject, fromDate, toDate):
    """
            :param bankObject: Bank Object (đã login)
            :param fromDate: Ngày bắt đầu lấy dữ liệu
            :param toDate: Ngày kết thúc lấy dữ liệu
        """
    try:
        transactionTable = BankTransferBalance.runBIDV(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        print(bankObject.ignored_exceptions)
        time.sleep(3)
        xpath = "//*[contains(text(), 'Tên đăng nhập')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = BIDV(True).Login()
            print('Done login')

    time.sleep(1)
    runBIDV(bankObject, fromDate, toDate)






