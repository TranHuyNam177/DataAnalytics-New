from automation.finance import *
from automation.finance import BankTransactionHistory

# không CAPTCHA
def runTCB(bankObject,fromDate,toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransactionHistory.runTCB(bankObject,fromDate,toDate)
    except bankObject.ignored_exceptions:
        time.sleep(1)
        xpath = "//*[contains(text(),'vui lòng đăng nhập lại')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH,xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = TCB(True).Login()
            print('Done login')
    time.sleep(3)
    runTCB(bankObject,fromDate,toDate)

# không CAPTCHA
def runVTB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransactionHistory.runVTB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        time.sleep(1)
        xpath = "//*[@type='submit']/span[text()='Đăng nhập']"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = VTB(True).Login()
            print('Done login')
    time.sleep(3)
    runVTB(bankObject, fromDate, toDate)

# không CAPTCHA
def runOCB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransactionHistory.runOCB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        time.sleep(1)
        xpath = "//*[contains(text(),'Đăng nhập OMNI Doanh nghiệp')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = OCB(True).Login()
            print('Done login')
    time.sleep(3)
    runOCB(bankObject, fromDate, toDate)

# không CAPTCHA
def runEIB(bankObject, fromDate, toDate):
    """
        :param bankObject: Bank Object (đã login)
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngày kết thúc lấy dữ liệu
        """
    try:
        transactionTable = BankTransactionHistory.runEIB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        time.sleep(1)
        xpath = "//*[contains(text(),'Phiên làm việc hết hiệu lực')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = EIB(True).Login()
            print('Done login')
    time.sleep(3)
    runEIB(bankObject, fromDate, toDate)

# có CAPTCHA
def runVCB(bankObject, fromDate, toDate):
    """
        :param bankObject: Bank Object (đã login)
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngày kết thúc lấy dữ liệu
        """
    try:
        transactionTable = BankTransactionHistory.runVCB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        time.sleep(1)
        xpath = '//*[@value="Đăng nhập"]'
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = VCB(True).Login()
            print('Done login')

    time.sleep(3)
    runVCB(bankObject, fromDate, toDate)

# có CAPTCHA
def runIVB(bankObject, fromDate, toDate):
    """
        :param bankObject: Bank Object (đã login)
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngày kết thúc lấy dữ liệu
    """
    try:
        transactionTable = BankTransactionHistory.runIVB(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        time.sleep(1)
        xpath = "//h2[contains(text(), 'Phiên giao dịch hết hạn')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = IVB(True).Login()
            print('Done login')

    time.sleep(3)
    runIVB(bankObject, fromDate, toDate)

# có CAPTCHA
def runBIDV(bankObject, fromDate, toDate):
    """
            :param bankObject: Bank Object (đã login)
            :param fromDate: Ngày bắt đầu lấy dữ liệu
            :param toDate: Ngày kết thúc lấy dữ liệu
        """
    try:
        transactionTable = BankTransactionHistory.runBIDV(bankObject, fromDate, toDate)
    except bankObject.ignored_exceptions:
        time.sleep(1)
        xpath = "//*[contains(text(), 'Tên đăng nhập')]"
        reLoginNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if reLoginNotices:  # Bị Log out
            print(reLoginNotices)
            bankObject.__del__()
            bankObject = BIDV(True).Login()
            print('Done login')

    time.sleep(3)
    runBIDV(bankObject, fromDate, toDate)






