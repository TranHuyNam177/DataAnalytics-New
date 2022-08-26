import time
import pandas as pd
import selenium
import selenium.webdriver
from automation.finance import *

def runOCB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    now = dt.datetime.now()
    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'DSKhoanVay' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Click main menu
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'main-menu-icon'))).click()
    # Click Khoản vay -> Danh sách khoản vay
    xpath = '//*[@class="ahref"]/*[@class="credits-icon"]'
    _, creditElement = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    creditElement.click()
    xpath = '//*[contains(text(),"Danh sách khoản vay") and ../@class="nav-submenu submenu_visible"]'
    _, listElement = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    listElement.click()
    time.sleep(5) # chờ chuyển trang
    # Check xem có khoản vay không
    xpath = '//*[contains(text(),"Không có khoản vay nào")]'
    notices = bankObject.driver.find_elements(By.XPATH,xpath)
    if notices: # không có khoản vay
        return pd.DataFrame()
    # Click Tải về chi tiết
    xpath = '//*[text()="Tải về"]'
    while True:
        try:
            bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
            break
        except (ElementNotInteractableException,):
            time.sleep(1)  # click bị fail thì chờ 1s rồi click lại
    time.sleep(1)  # chờ animation
    xpath = '//*[text()="Tập tin XLS"]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Đọc file download
    while True:
        downloadFile = first(listdir(bankObject.downloadFolder),lambda x:'DSKhoanVay' in x and 'download' not in x)
        if downloadFile: # download xong -> có file
            break
        time.sleep(1) # chưa download xong -> đợi thêm 1s nữa
    balanceTable = pd.read_excel(
        join(bankObject.downloadFolder,downloadFile),
        usecols='B:E,I:L',
        skiprows=19,
        names=['ContractNumber','Currency','ExpireDate','Remaining','Amount','IssueDate','InterestRate','Paid'],
        dtype=object
    )
    # Expire Date
    balanceTable['ExpireDate'] = pd.to_datetime(balanceTable['ExpireDate'],format='%d/%m/%Y')
    # Remaining
    balanceTable['Remaining'] = balanceTable['Remaining'].str.replace(',','').astype(np.float64)
    # Amount
    balanceTable['Amount'] = balanceTable['Amount'].str.replace(',','').astype(np.float64)
    # Issue Date
    balanceTable['IssueDate'] = pd.to_datetime(balanceTable['IssueDate'],format='%d/%m/%Y')
    # Interest Rate
    balanceTable['InterestRate'] = balanceTable['InterestRate'].str.replace('%','').str.strip().astype(float) / 100
    # Amount
    balanceTable['Paid'] = balanceTable['Paid'].str.replace(',','').astype(np.float64)
    # TermDays
    balanceTable['TermDays'] = (balanceTable['ExpireDate']-balanceTable['IssueDate']).dt.days
    # TermMonths
    balanceTable['TermMonths'] = (balanceTable['TermDays']/30).round().astype(np.int64)
    # Interest Amount
    balanceTable['InterestAmount'] = balanceTable['TermDays'] * balanceTable['InterestRate'] / 365 * balanceTable['Amount']
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0,'Date',d)
    # Bank
    balanceTable.insert(1,'Bank',bankObject.bank)

    return balanceTable


def runESUN(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    now = dt.datetime.now()
    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('mainFrame')
    time.sleep(1)
    # Check xem có Bank Announcements không
    xpath = '//*[contains(text(),"Bank Announcements")]'
    checkElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    if checkElement.is_displayed():
        xpath = '//*[contains(text(),"Confirm")]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        time.sleep(1)
    # Click Loan
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'menuIndex_2'))).click()
    # Click Loan Overview
    xpath = "//*[@id='menuIndex_2']//*[contains(text(),'Loan Overview')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    time.sleep(3)
    # Lấy danh sách records
    xpath = "//*[@class='ui-datatable-even']"
    rowElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    records = []
    for rowElement in rowElements:
        rowText = rowElement.text
        # ContractNumber
        contractNumber = rowText.split()[1]
        # IssueDate, ExpireDate
        issueDateString, expireDateString, _ = re.findall(r'\b\d{4}/\d{2}/\d{2}\b',rowText)
        issueDate = dt.datetime.strptime(issueDateString,'%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateString,'%Y/%m/%d')
        # Currency
        currency = re.search('VND|USD',rowText).group()
        # Remaining
        remainingString = re.search(r'\b(VND|USD)\s[\d,.]+\b',rowText).group()
        remainingString = re.sub(r'VND|USD','',remainingString).replace(',','').strip()
        remaining = float(remainingString)
        # InterestRate
        interestRateString = re.search(r'\s\d+\.\d+\b',rowText).group().strip()
        interestRate = float(interestRateString) / 100
        # Amount = Remaining
        amount = remaining
        # Paid = 0
        paid = 0
        # TermDays
        termDays = (expireDate-issueDate).days
        # TermMonths
        termMonths = round(termDays/30)
        # Interest Amount
        interestAmount = interestRate / 365 * termDays * amount
        # Append records
        records.append((contractNumber,termDays,termMonths,interestRate,issueDate,expireDate,amount,paid,remaining,interestAmount,currency))

    balanceTable = pd.DataFrame(
        records,
        columns=['ContractNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Amount','Paid','Remaining','InterestAmount','Currency']
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0,'Date',d)
    # Bank
    balanceTable.insert(1,'Bank',bankObject.bank)

    return balanceTable


def runIVB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    now = dt.datetime.now()
    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    mainPageURL = bankObject.driver.current_url
    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'logoimg'))).click()
    # Click tab "Tài khoản"
    xpath = '//*[@data-menu-id="1"]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    time.sleep(1)  # chờ animation
    # Click subtab "Thông tin tài khoản"
    bankObject.wait.until(EC.visibility_of_element_located((By.ID,'1_1'))).click()
    # Click "Thông tin vay"
    bankObject.driver.switch_to.frame('mainframe')
    while True:
        xpath = '//*[@id="fd_layer_pan"]/span'
        bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath))).click()
        xpath = '//*[contains(@data-th,"Số Hợp đồng")]/a'
        contractElements = bankObject.driver.find_elements(By.XPATH,xpath)
        URLs = [e.get_attribute('href') for e in contractElements]
        if contractElements[0].text:
            break

    frames = []
    for URL in URLs:
        bankObject.driver.get(URL)
        time.sleep(1)  # chờ để hiện số tài khoản (bắt buộc)
        # Danh sách các trường thông tin cần lấy:
        mapper = {
            'Số Hợp đồng': 'ContractNumber',
            'Số tiền vay': 'Amount',
            'Dư nợ hiện tại': 'Remaining',
            'Loại tiền': 'Currency',
            'Ngày giải ngân': 'IssueDate',
            'Ngày đến hạn': 'ExpireDate',
            'Lãi suất': 'InterestRate',
        }
        recordDict = dict()
        for subString, colName in mapper.items():
            xpath = f"//*[contains(text(),'{subString}')]//following-sibling::td"
            infoElements = bankObject.driver.find_elements(By.XPATH,xpath)
            if infoElements:
                recordDict[colName] = [infoElements[0].text]
            else:
                recordDict[colName] = [None]

        frame = pd.DataFrame(recordDict)
        frames.append(frame)
    
    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()
    balanceTable = pd.concat(frames)
    
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable['Date'] = d
    balanceTable['Bank'] = bankObject.bank
    balanceTable['Amount'] = balanceTable['Amount'].str.replace(',','').astype(np.float64)
    balanceTable['Remaining'] = balanceTable['Remaining'].str.replace(',','').astype(np.float64)
    balanceTable['Paid'] = balanceTable['Amount'] - balanceTable['Remaining']
    balanceTable['InterestRate'] = (balanceTable['InterestRate'].str.split('%').str.get(0)).astype(np.float64)/100
    balanceTable['IssueDate'] = pd.to_datetime(balanceTable['IssueDate'],format='%d/%m/%Y')
    balanceTable['ExpireDate'] = pd.to_datetime(balanceTable['ExpireDate'],format='%d/%m/%Y')
    # TermDays
    balanceTable['TermDays'] = (balanceTable['ExpireDate'] - balanceTable['IssueDate']).dt.days
    # Term months
    balanceTable['TermMonths'] = round(balanceTable['TermDays']/30)
    # Interest Amount
    balanceTable['InterestAmount'] = balanceTable['TermDays'] * balanceTable['InterestRate'] * balanceTable['Amount'] / 360
    # Về lại trang chủ
    bankObject.driver.get(mainPageURL)

    return balanceTable


def runFUBON(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    now = dt.datetime.now()
    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Dọn dẹp folder trước khi download
    today = now.strftime('%Y%m%d')
    for file in listdir(bankObject.downloadFolder):
        if file.startswith(today):
            os.remove(join(bankObject.downloadFolder,file))
    # Click "Financing"
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('topmenu')
    xpath = "//*[contains(text(),'Financing')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click "Lending Details Enquiry"
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('menu')
    bankObject.driver.switch_to.frame('area')
    bankObject.driver.switch_to.frame('info')
    bankObject.wait.until(EC.presence_of_element_located((By.NAME,'nodeIcon1'))).click()
    xpath = '//*[contains(text(),"Lending Details Enquiry") and @target="main"]'
    bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath))).click()
    # Click Enquiry
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('main')
    xpath = '//*[@name="query" and @type="button" and @class="fnct_btn"]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Check xem có dữ liệu ko
    time.sleep(1)
    bankObject.driver.switch_to.frame('Data2')
    xpath = '//*[contains(text(),"Information Not Exists")]'
    noticeStrings = bankObject.driver.find_elements(By.XPATH,xpath)
    if noticeStrings: # ko có dữ liệu
        return pd.DataFrame()
    # Contract Number
    xpath = '//*[contains(@class,"DataRow")]/td[2]' # could be either DataRowOdd or DataRowEven
    contractNumberElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    contractNumbers = [element.text for element in contractNumberElements]
    # Remaining
    def convertNumber(valueString):
        if valueString:
            return float(valueString.replace(',',''))
    xpath = '//*[contains(@class,"DataRow")]/td[4]'
    remainingElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    remainings = [convertNumber(element.text) for element in remainingElements]
    # Currency
    xpath = '//*[contains(@class,"DataRow")]/td[5]'
    currencyElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    currencies = [element.text for element in currencyElements]
    # ExpireDate
    def convertDatetime(valueString):
        if valueString:
            return dt.datetime.strptime(valueString,'%Y/%m/%d')
    xpath = '//*[contains(@class,"DataRow")]/td[6]'
    expireDateElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    expireDates = [convertDatetime(element.text) for element in expireDateElements]
    # Interest Rate
    def convertRate(valueString):
        if valueString:
            return float(valueString) / 100
    xpath = '//*[contains(@class,"DataRow")]/td[7]'
    interestRateElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    interestRates = [convertRate(element.text) for element in interestRateElements]

    # Click Sub-Function
    xpath = '//*[contains(@class,"DataRow")]//*[contains(@name,"select")]'
    loanButtons = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    parentWindow = bankObject.driver.current_window_handle

    frames = []
    for loanButton in loanButtons:
        bankObject.driver.switch_to.window(parentWindow)
        bankObject.driver.switch_to.default_content()
        bankObject.driver.switch_to.frame('main')
        bankObject.driver.switch_to.frame('Data2')
        select = Select(loanButton)
        select.select_by_index(1) # option thứ 2
        time.sleep(3) # chờ pop-up window
        # Switch sang pop-up window
        windows = bankObject.driver.window_handles # toàn bộ các window Google Chrome đang mở thuộc instance này (list 2 phần tử)
        windows.remove(parentWindow) # bỏ đi parent window ra khỏi list
        bankObject.driver.switch_to.window(windows[0]) # switch sang cái còn lại (child window)
        # Nhập ngày
        xpath = '//*[@class="TD_Input"]/input'
        fromDateInput, toDateInput = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
        fromDateString = (now-dt.timedelta(days=365)).strftime('%Y/%m/%d')
        toDateString = now.strftime('%Y/%m/%d')
        fromDateInput.clear()
        fromDateInput.send_keys(fromDateString)
        toDateInput.clear()
        toDateInput.send_keys(toDateString)
        # Click Enquiry
        xpath = '//*[@name="query" and @type="button" and @class="fnct_btn"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Amount
        bankObject.driver.switch_to.default_content()
        bankObject.driver.switch_to.frame('Data2')
        xpath = '//*[contains(text(),"Take-Down")]//following-sibling::td[1]'
        amountString = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        amount = convertNumber(amountString)
        # Issue Date
        xpath = '//*[contains(text(),"Take-Down")]//preceding-sibling::td[1]'
        issueDateString = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        issueDate = convertDatetime(issueDateString)
        # Click Close Window
        bankObject.driver.switch_to.default_content()
        bankObject.driver.switch_to.frame('Data3')
        xpath = '//*[contains(@value,"Close Window")]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Swith back to parent Window
        bankObject.driver.switch_to.window(parentWindow) # switch sang cái còn lại (child window)
        bankObject.driver.switch_to.default_content()
        bankObject.driver.switch_to.frame('main')
        bankObject.driver.switch_to.frame('Data2')
        # Append
        frame = pd.DataFrame(data=[(amount,issueDate)],columns=['Amount','IssueDate'])
        frames.append(frame)

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()
    balanceTable = pd.concat(frames)
    
    if now.hour >= 8:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable['Date'] = d
    balanceTable['Bank'] = bankObject.bank
    balanceTable['ContractNumber'] = contractNumbers
    balanceTable['Remaining'] = remainings
    balanceTable['Currency'] = currencies
    balanceTable['ExpireDate'] = expireDates
    balanceTable['InterestRate'] = interestRates
    balanceTable['TermDays'] = (balanceTable['ExpireDate'] - balanceTable['IssueDate']).dt.days
    balanceTable['TermMonths'] = (balanceTable['TermDays']/30).round()
    balanceTable['InterestAmount'] = balanceTable['TermDays'] * balanceTable['InterestRate'] * balanceTable['Amount'] / 360
    balanceTable['Paid'] = balanceTable['Amount'] - balanceTable['Remaining']

    return balanceTable


def runMEGA(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    now = dt.datetime.now()
    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    frameElement = bankObject.wait.until(EC.presence_of_element_located((By.ID,"ifrm")))
    bankObject.driver.switch_to.frame(frameElement)
    # Click Accounts
    xpath = '//*[contains(text(),"Accounts")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click Balance overview
    xpath = '//*[contains(text(),"Balance overview")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Switch frame
    bankObject.driver.switch_to.frame('frame1')
    # Get data in Balance overview -> Loan guarantee accounts
    xpath = '//*[contains(text(),"Loan Start Date")]//parent::tr//following-sibling::tr[*]'
    rowElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    records = []
    # Balance Overview
    for rowElement in rowElements:
        rowString = rowElement.text
        accountNumber = re.search(r'\b(VND|USD)\s[\d-]*\b',rowString).group()
        # Issue Date, Expire Date
        issueDateText, expireDateText = re.findall(r'\b\d{4}/\d{2}/\d{2}\b',rowString)
        issueDate = dt.datetime.strptime(issueDateText,'%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText,'%Y/%m/%d')
        # Term Days
        termDays = (expireDate-issueDate).days
        # Term Months
        termMonths = round(termDays/30)
        # Lãi suất
        interestRateString = re.search(r'\b\d{1,2}\.\d+\b',rowString).group()
        interestRate = float(interestRateString) / 100
        # Currency
        currency = re.search(r'VND|USD',rowString).group()
        # Remaining
        remainingString = re.search(r'\b\d+,.*,\d{3}\b',rowString).group()
        remaining = float(remainingString.replace(',', ''))
        # Append
        records.append((accountNumber,termDays,termMonths,interestRate,issueDate,expireDate,remaining,currency))

    balanceTable = pd.DataFrame(
        records,
        columns=['AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Remaining','Currency']
    )
    # Click tab Financing
    bankObject.driver.switch_to.default_content()
    # bankObject.driver.switch_to.frame(bankObject.driver.find_element(By.ID,'ifrm'))
    bankObject.driver.switch_to.frame(bankObject.wait.until(EC.presence_of_element_located((By.ID,'ifrm'))))
    xpath = "//*[contains(text(),'Financing')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click Loan details
    xpath = "//*[contains(text(),'Loan details')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Switch frame
    bankObject.driver.switch_to.frame('frame1')

    balanceTable['ContractNumber'] = None
    balanceTable['Amount'] = None
    balanceTable['Paid'] = None
    for rowLoc in balanceTable.index:
        # Select Account
        accountNumber = balanceTable.loc[rowLoc,'AccountNumber']
        accountNumber = accountNumber.rsplit('-',1)[0] + '-' + accountNumber.rsplit('-',1)[1].lstrip('0')
        xpath = '//*[contains(@id,"account")]'
        dropDownElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        Select(dropDownElement).select_by_visible_text(accountNumber)
        # Insert IssueDate into "Start Date" box
        issueDate = balanceTable.loc[rowLoc,'IssueDate'].strftime('%Y/%m/%d')
        xpath = '//*[contains(@id,"startDate")]'
        startDateInput = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        startDateInput.clear()
        startDateInput.send_keys(issueDate)
        # Click Inquire
        xpath = "//*[contains(text(),'Inquire')]"
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Contract Number
        xpath = f"//*[text()='{issueDate}']//following-sibling::td[1]"
        contractNumber = bankObject.driver.find_element(By.XPATH,xpath).text
        balanceTable.loc[rowLoc,'ContractNumber'] = contractNumber
        # Amount
        xpath = f"//*[text()='{issueDate}']//following-sibling::td[2]"
        amountString = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        amount = float(amountString.replace(',',''))
        balanceTable.loc[rowLoc,'Amount'] = amount
        # Paid
        remaining = balanceTable.loc[rowLoc,'Remaining']
        paid = amount - remaining
        balanceTable.loc[rowLoc,'Paid'] = paid

    # Interest Amount
    balanceTable['InterestAmount'] = balanceTable['InterestRate'] * balanceTable['TermDays'] * balanceTable['Amount'] / 360
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0,'Date',d)
    # Bank
    balanceTable.insert(1,'Bank',bankObject.bank)
    # Drop AccountNumber
    balanceTable = balanceTable.drop('AccountNumber',axis=1)
    return balanceTable


def runSINOPAC(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    now = dt.datetime.now()
    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if re.search(r'\bCOSLABAQU_\d+_\d+\b',file):
            os.remove(join(bankObject.downloadFolder,file))
    # Click Account Inquiry
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('indexFrame')
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'MENU_CAO'))).click()
    # Click Loan inquiry
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'MENU_CAO002'))).click()
    time.sleep(1)
    # Click Loan Balance Inquiry
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'MENU_COSLABAQU'))).click()
    # Chờ load xong data
    xpath = '//*[contains(text(),"Processing")]'
    while True:
        processingElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        if not processingElement.is_displayed(): # load data xong
            break
        time.sleep(1)
    # Click Search
    bankObject.driver.switch_to.frame('mainFrame')
    xpath = '//*[contains(text(),"Search")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click download file csv
    xpath = '//*[contains(@class,"download_csv")]'
    while True:
        downloadElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        if downloadElement.is_displayed():
            time.sleep(1)
            downloadElement.click()
            break
        time.sleep(1)
    # Đọc file download
    while True:
        checkFunc = lambda x: (re.search(r'\bCOSLABAQU_\d+_\d+\b', x) is not None) and ('download' not in x)
        downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
        if downloadFile:  # download xong -> có file
            break
        time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
    balanceTable = pd.read_csv(
        join(bankObject.downloadFolder,downloadFile),
        usecols=[1,3,5,6,7,8],
        names=['ContractNumber','IssueDate','Currency','Amount','Remaining','ExpireDate'],
        skiprows=1,
        dtype={
            'ContractNumber': object,
            'IssueDate': object,
            'Currency': object,
            'Amount': np.float64,
            'Remaining': np.float64,
            'ExpireDate': object
        }
    )
    balanceTable['IssueDate'] = pd.to_datetime(balanceTable['IssueDate'],format='%Y/%m/%d')
    balanceTable['ExpireDate'] = pd.to_datetime(balanceTable['ExpireDate'],format='%Y/%m/%d')
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable['Date'] = d
    balanceTable['Bank'] = bankObject.bank
    # TermDays
    balanceTable['TermDays'] = (balanceTable['ExpireDate'] - balanceTable['IssueDate']).dt.days
    # Term months
    balanceTable['TermMonths'] = (balanceTable['TermDays']/30).round()
    # Paid
    balanceTable['Paid'] = balanceTable['Amount'] - balanceTable['Remaining']
    return balanceTable


def runHUANAN(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    now = dt.datetime.now()
    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Click Loan Section
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('left')
    xpath = "//*[contains(text(),'Loan Section')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Reload main frame
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('main')
    # Show menu Inquiries
    xpath = "//*[contains(text(),'Inquiries')]"
    bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath))).click()
    # Click Outstanding loan inquiry
    xpath = "//*[contains(text(),'Outstanding')]"
    bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath))).click()
    # Select option of Unit
    xpath = "//*[contains(@name,'Unit')]/option"
    unitElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    unitOptions = [u.text for u in unitElements]
    # Select option of Customer ID
    xpath = "//*[contains(@name,'CoID_I')]/option"
    customerIDElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    customerIDOptions = [c.text for c in customerIDElements]

    for prefix in ['S','E']:
        for suffix in ['Year','Month','Date']:
            xpath = f"//*[@name='{prefix}_{suffix}']"
            inputElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            select = Select(inputElement)
            if suffix == 'Date':
                select.select_by_visible_text(str(now.day))
            elif suffix == 'Month':
                select.select_by_visible_text(str(now.month))
            else:
                if prefix == 'S':
                    select.select_by_visible_text(str(now.year-1))
                else: # prefix == 'E'
                    select.select_by_visible_text(str(now.year))

    rowElements = []
    for unitOption in unitOptions:
        for customerIDOption in customerIDOptions:
            xpath = "//select[contains(@name,'Unit')]"
            dropDownUnit = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            select = Select(dropDownUnit)
            select.select_by_visible_text(unitOption)
            xpath = "//select[contains(@name,'CoID_I')]"
            dropDownCustomerID = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            select = Select(dropDownCustomerID)
            select.select_by_visible_text(customerIDOption)
            # Click Submit
            xpath = "//a[contains(text(),'Submit')]"
            bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
            time.sleep(3) # chờ load data
            # Get data
            xpath = "//*[@class='Table_contentWt_C']"
            recordElements = bankObject.driver.find_elements(By.XPATH,xpath)
            if not recordElements:
                return pd.DataFrame()
            rowElements.extend(recordElements)

    records = []
    for rowElement in rowElements:
        rowString = rowElement.text
        # Contract Number
        contractNumber = re.search(r'[A-Z]+\d{3,}',rowString).group()
        # Issue Date, Expire Date
        issueDateString, expireDateString = re.findall(r'\b\d{4}/\d{2}/\d{2}\b',rowString)
        issueDate = dt.datetime.strptime(issueDateString,'%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateString,'%Y/%m/%d')
        # Term Days
        termDays = (expireDate-issueDate).days
        # Term Months
        termMonths = round(termDays/30)
        # Lãi suất
        interestRateString = re.search(r'\b\d{1,2}\.\d+\b',rowString).group()
        interestRate = float(interestRateString) / 100
        # Currency
        currency = re.search(r'VND|USD',rowString).group()
        # Amount and Remaining
        amountString, remainingString = re.findall(r'\b\d+,[\d,]+\.\d{2}\b',rowString)
        amount = float(amountString.replace(',',''))
        remaining = float(remainingString.replace(',',''))
        # Paid
        paid = amount - remaining
        # Interest Amount
        interestAmount = interestRate * termDays * amount / 360
        # Append data
        records.append((contractNumber,termDays,termMonths,interestRate,issueDate,expireDate,amount,paid,remaining,interestAmount,currency))

    balanceTable = pd.DataFrame(
        data=records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0,'Date',d)
    # Bank
    balanceTable.insert(1,'Bank',bankObject.bank)

    return balanceTable