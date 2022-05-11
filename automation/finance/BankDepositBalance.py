from automation.finance import *

def runBIDV(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'EBK_BC_TIENGOICOKYHAN' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Click Menu bar
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'menu-toggle-22'))).click()
    # Click "Tài khoản"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Vấn tin'))).click()
    # Click "Tiền gửi thanh toán"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Tiền gửi có kỳ hạn'))).click()
    # Click download
    bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'export'))).click()
    time.sleep(1) # chờ file download xong
    # Đọc file download để lấy danh sách tài khoản
    while True:
        checkFunc = lambda x:'EBK_BC_TIENGOICOKYHAN' in x and 'download' not in x
        downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
        if downloadFile:  # download xong -> có file
            break
        time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa

    accountNumbers = pd.read_excel(
        join(bankObject.downloadFolder,downloadFile),
        skiprows=7,
        usecols='C',
        skipfooter=1,
        dtype='object',
        squeeze=True,
    )
    # Xóa file
    os.remove(join(bankObject.downloadFolder,downloadFile))

    # Click từng tài khoản
    records = []
    for accountNumber in accountNumbers:
        xpath = f'//*[text()="{accountNumber}"]'
        accountElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        accountElement.click()
        time.sleep(0.5) # chờ animation
        # Ngày
        now = dt.datetime.now()
        if now.hour >= 12:
            d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        # Lãi suất
        xpath = f'//*[@aria-expanded="true"]/*[@class="information"]/div/div[2]/div[4]/p' # dòng 2 cột 4
        irateString, _ = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text.split()
        irate = float(irateString) / 100
        # Kỳ hạn
        xpath = f'//*[@aria-expanded="true"]/*[@class="information"]/div/div[2]/div[2]/p' # dòng 2 cột 2
        termString, _ = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text.split()
        term = float(termString) # số tháng
        # Ngày phát hành
        xpath = f'//*[@aria-expanded="true"]/*[@class="information"]/div/div[2]/div[1]/p' # dòng 2 cột 1
        issueDateString = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        issueDate = dt.datetime.strptime(issueDateString,'%d/%m/%Y')
        # Ngày đáo hạn
        xpath = f'//*[@aria-expanded="true"]/*[@class="information"]/div/div[4]/div[2]/p' # dòng 4 cột 2
        expireDateString = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        expireDate = dt.datetime.strptime(expireDateString,'%d/%m/%Y')
        # Số dư & currency
        xpath = f'//*[@aria-expanded="true"]/*[@class="information"]/div/div[3]/div[3]/p' # dòng 3 cột 3
        balanceString, currency = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text.split()
        balance = float(balanceString.replace(',',''))
        # Số ngày
        termDays = (expireDate-issueDate).days
        # Lãi
        interestAmount = irate/360*termDays*balance
        # Click để collapse dòng
        accountElement.click()
        # Append
        record = (d,bankObject.bank,accountNumber,termDays,term,irate,issueDate,expireDate,balance,interestAmount,currency)
        records.append(record)

    balanceTable = pd.DataFrame(
        data=records,
        columns=['Date','Bank','AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    )

    return balanceTable

def runVTB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'danh-sach-tai-khoan-tien-gui' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Bắt đầu từ trang chủ
    xpath = '//*[@href="/"]'
    _, MainMenu = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    MainMenu.click()
    # Click menu "Tài khoản"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Tài khoản'))).click()
    # Click sub-menu "Danh sách tài khoản"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Danh sách tài khoản'))).click()
    # Download File
    xpath = '//*[@src="/public/img/icon/icon-download.svg"]'
    _, downloadElement, _ = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    xpath = '//*[text()="Xuất Excel"]'
    downloadElement.click()
    _, exportElement, _ = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    exportElement.click()
    # Đọc file download
    while True:
        checkFunc = lambda x:'danh-sach-tai-khoan-tien-gui' in x and 'download' not in x
        downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
        if downloadFile:  # download xong -> có file
            break
        time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
    downloadTable = pd.read_excel(
        join(bankObject.downloadFolder,downloadFile),
        names=['AccountNumber','TermMonths','InterestRate','Currency','IssueDate','ExpireDate','Balance','InterestAmount'],
        skiprows=16,
        skipfooter=1,
        usecols='B,D:I,L',
        dtype={'AccountNumber':object,'Balance':np.float64,'InterestAmount':np.float64}
    )
    downloadTable[['IssueDate','ExpireDate']] = downloadTable[['IssueDate','ExpireDate']].applymap(lambda x: dt.datetime.strptime(x,'%d-%m-%Y'))
    downloadTable['TermMonths']  = np.int64(downloadTable['TermMonths'].str.replace('D','').str.split('M').str.get(0))
    downloadTable['TermDays'] = (downloadTable['ExpireDate']-downloadTable['IssueDate']).dt.days
    downloadTable['InterestRate'] /= 100
    downloadTable['Bank'] = bankObject.bank
    # Date
    now = dt.datetime.now()
    if now.hour >= 12:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    downloadTable['Date'] = d
    cols = ['Date','Bank','AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    balanceTable = downloadTable[cols]

    return balanceTable

def runIVB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'logoimg'))).click()
    # Click tab "Tài khoản"
    xpath = '//*[@data-menu-id="1"]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click subtab "Thông tin tài khoản"
    bankObject.wait.until(EC.visibility_of_element_located((By.ID,'1_1'))).click()
    # Click "Tài khoản tiền gửi có kỳ hạn"
    while True:
        bankObject.driver.switch_to.frame('mainframe')
        xpath = '//*[@id="sa_layer_pan"]/span'
        bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath))).click()
        # Lấy thông tin
        records = []
        rowElements = bankObject.driver.find_elements(By.XPATH,'//*[@id="SAccount"]/tbody/tr')[:-1] # bỏ dòng Tổng
        if rowElements[0].text != '':
            break
    for element in rowElements:
        elementString = element.text
        # Date
        now = dt.datetime.now()
        if now.hour >= 12:
            d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        # Số tài khoản
        account = re.search('^[0-9]{7}-[0-9]{3}',elementString).group()
        # Ngày hiệu lực, Ngày đáo hạn
        issueDateText, expireDateText = re.findall('[0-9]{2}.[0-9]{2}.[0-9]{4}',elementString)
        issueDate = dt.datetime.strptime(issueDateText,'%d/%m/%Y')
        expireDate = dt.datetime.strptime(expireDateText,'%d/%m/%Y')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termString = re.search('[0-9]* Tháng',elementString).group().replace(' Tháng','')
        term = float(termString) # Số tháng
        # Lãi suất
        iText = elementString.split(issueDateText)[0].split()[-1]
        iRate = round(float(iText)/100,5)
        # Balance
        balanceString = re.search('[0-9]+ VND',element.text.replace(',','')).group()
        balanceString, _ = balanceString.split()
        balance = float(balanceString)
        # Currency
        currency = re.search('VND|USD',element.text).group()
        # Lãi
        interest = float(re.findall('[0-9]+',element.text.replace(',',''))[-1])
        records.append((d,bankObject.bank,account,termDays,term,iRate,issueDate,expireDate,balance,interest,currency))

    balanceTable = pd.DataFrame(
        data=records,
        columns=['Date','Bank','AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    )

    return balanceTable

def runVCB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """
    
    # Bắt đầu từ trang chủ
    _, homeButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'icon-home')))
    homeButton.click()
    # Lấy danh sách đường dẫn vào tài khoản tiền gửi có kỳ hạn
    xpath = '//*[@id="dstkfd-tbody"]//td/a'
    accountElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    URLs = [e.get_attribute('href') for e in accountElems]

    records = []
    for URL in URLs:
        bankObject.driver.get(URL)
        time.sleep(1)  # chờ để hiện số tài khoản (bắt buộc)
        # Ngày
        now = dt.datetime.now()
        if now.hour >= 12:
            d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        # Tài khoản
        xpath = '//*[@id="Lbl_STK_Title" and text()!=""]'
        account = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        # Ngày hiệu lực
        xpath = '//*[@id="Lbl_NMS" and text()!=""]'
        issueDateText = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        issueDate = dt.datetime.strptime(issueDateText,'%d/%m/%Y')
        # Ngày đáo hạn
        xpath = '//*[@id="Lbl_NDH" and text()!=""]'
        expireDateText = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text
        expireDate = dt.datetime.strptime(expireDateText,'%d/%m/%Y')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        xpath = '//*[@id="Lbl_KH" and text()!=""]'
        termMonths = float(bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text.replace(' tháng',''))
        # Interest rate
        xpath = '//*[@id="Lbl_LSFD" and text()!=""]'
        iRate = float(bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text.replace('% /năm','')) / 100
        # Số dư, currency
        xpath = '//*[@id="Lbl_SDS" and text()!=""]'
        balanceString, currency = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).text.split()
        balance = float(balanceString.replace(',',''))
        records.append((d,bankObject.bank,account,termDays,termMonths,iRate,issueDate,expireDate,balance,iRate,currency))

    balanceTable = pd.DataFrame(
        data=records,
        columns=['Date','Bank','AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    )

    return balanceTable

def runOCB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'DSHDTG' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Click main menu
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'main-menu-icon'))).click()
    # Click Hợp đồng tiền gửi -> Danh sách hợp đồng tiền gửi
    xpath = '//*[@class="ahref"]/*[@class="deposits-icon"]'
    _, accountButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    accountButton.click()
    xpath = '//*[@id="side-nav"]/ul/li[7]/div[2]/span'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    time.sleep(5) # chờ chuyển trang

    # Click download
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@placeholder="Tải về"]'))).click()
    bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@ng-click="downloadFile(option)"]'))).click()
    # Đọc file download
    while True:
        downloadFile = first(listdir(bankObject.downloadFolder),lambda x:'DSHDTG' in x and 'download' not in x)
        if downloadFile: # download xong -> có file
            break
        time.sleep(1) # chưa download xong -> đợi thêm 1s nữa
    downloadTable = pd.read_excel(
        join(bankObject.downloadFolder,downloadFile),
        skiprows=15,
        usecols='B,E:I',
        names=['AccountNumber','Currency','IssueDate','ExpireDate','InterestRate','InterestAmount'],
        dtype={'AccountNumber':object}
    )
    downloadTable[['IssueDate','ExpireDate']] = downloadTable[['IssueDate','ExpireDate']].applymap(lambda x: dt.datetime.strptime(x,'%d/%m/%Y'))
    downloadTable['InterestRate'] = downloadTable['InterestRate'].map(lambda x: float(x.replace(' %',''))) / 100
    downloadTable['InterestAmount'] = downloadTable['InterestAmount'].map(lambda x: float(x.replace(',','')))
    downloadTable['TermDays'] = (downloadTable['ExpireDate']-downloadTable['IssueDate']).dt.days
    downloadTable['TermMonths'] = round(downloadTable['TermDays']/30)

    # Số dư
    xpath = '//*[@class="above"]/div/*[@class!="num"]'
    accountKeys = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    xpath = '//*[@amount="deposit.depositBalance"]/span[@class="bd-amount__value mg-right"]'
    balanceValues = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    mapper = {k.text: float(v.text.replace(',','')) for k,v in zip(accountKeys,balanceValues)}
    downloadTable['Balance'] = downloadTable['AccountNumber'].map(mapper)
    # Ngày
    now = dt.datetime.now()
    if now.hour >= 12:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    downloadTable['Date'] = d
    # Bank
    downloadTable['Bank'] = bankObject.bank

    cols = ['Date','Bank','AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    balanceTable = downloadTable[cols]
    # Xóa file
    os.remove(join(bankObject.downloadFolder,downloadFile))

    return balanceTable

