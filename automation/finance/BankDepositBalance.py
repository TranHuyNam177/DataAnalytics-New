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
    time.sleep(1)
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
    bankObject.driver.switch_to.frame('mainframe')
    while True:
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

    # Số dư: Tìm số lượng trang
    xpath = '//*[@class="bd-pagination__number"]'
    pageButtonElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    pageNumbers = [elem.text for elem in pageButtonElements if elem.text in '0123456789' and elem.text]
    mapper = dict()
    for pageNumber in pageNumbers:
        # Click chọn trang
        xpath = f'//*[@class="bd-pagination__number" and text()="{pageNumber}"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        time.sleep(5)
        xpath = '//*[@class="above"]/div/*[@class!="num"]'
        accountKeys = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
        xpath = '//*[@amount="deposit.depositBalance"]/span[@class="bd-amount__value mg-right"]'
        balanceValues = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
        for k,v in zip(accountKeys,balanceValues):
            mapper[k.text] = float(v.text.replace(',',''))
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

def runFUBON(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    # Dọn dẹp folder trước khi download
    today = dt.datetime.now().strftime('%Y%m%d')
    for file in listdir(bankObject.downloadFolder):
        if file.startswith(today):
            os.remove(join(bankObject.downloadFolder,file))
    # Click "Deposit Account"
    bankObject.driver.switch_to.frame('topmenu')
    xpath = "//*[contains(text(),'Deposit Account')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click "Account Overview"
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('menu')
    bankObject.driver.switch_to.frame('area')
    bankObject.driver.switch_to.frame('info')
    xpath = "//*[contains(text(),'Account Overview')]"
    _, clickObject = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    clickObject.click()
    # Click "EXCEL File Download"
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('main')
    bankObject.driver.switch_to.frame('Data3')
    xpath = '//*[contains(@value,"EXCEL File Download")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Đọc file
    while True:
        downloadFile = first(listdir(bankObject.downloadFolder),lambda x: x.startswith(today) and 'download' not in x)
        if downloadFile:  # download xong -> có file
            break
        time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
    downloadTable = pd.read_excel(
        join(bankObject.downloadFolder,downloadFile),
        usecols='A,C,E:J',
    )
    startRow = downloadTable.loc[downloadTable.iloc[:,0].map(lambda x: isinstance(x,str) and 'Time Deposit Account' in x)].index[0] + 2
    endRow = downloadTable.loc[downloadTable.iloc[:,0].map(lambda x: isinstance(x,str) and 'Time Deposit Total' in x)].index[0]
    balanceTable = balanceTable.iloc[startRow:endRow,1:]
    balanceTable.columns = [
        'AccountNumber',
        'Currency',
        'Balance',
        'InterestRate',
        'TermDays',
        'IssueDate',
        'ExpireDate',
    ]
    # Date
    now = dt.datetime.now()
    if now.hour >= 12:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0,'Date',d)
    # Bank
    balanceTable.insert(1,'Bank',bankObject.bank)
    # Balance
    balanceTable['Balance'] = balanceTable['Balance'].str.replace('(,|\.\d*)','',regex=True).astype(np.int64)
    # TermDays
    balanceTable['TermDays'] = balanceTable['TermDays'].str.replace('[^0-9]','',regex=True).astype(np.int64)
    # TermMonths
    balanceTable['TermMonths'] = round(balanceTable['TermDays']/30)
    # Interest Rate
    balanceTable['InterestRate'] = balanceTable['InterestRate'].astype(np.float64) / 100
    # Issue Date
    balanceTable['IssueDate'] = balanceTable['IssueDate'].map(lambda x: dt.datetime.strptime(x,'%Y/%m/%d'))
    # Expire Date
    balanceTable['ExpireDate'] = balanceTable['ExpireDate'].map(lambda x: dt.datetime.strptime(x,'%Y/%m/%d'))
    # Interest Amount
    balanceTable['InterestAmount'] = None
    # Xóa file
    os.remove(join(bankObject.downloadFolder,downloadFile))

    return balanceTable

def runFIRST(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'COSDATDQU' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Click "Account Inquiry"
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('iFrameID')
    xpath = '//*[contains(text(),"Account Inquiry")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click "Time Deposit Detail"
    time.sleep(0.5)
    xpath = '//*[contains(text(),"Time Deposit Detail")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Chọn từng tài khoản từ dropdown list
    bankObject.driver.switch_to.frame('mainFrame')
    xpath = '//*[contains(text(),"Deposit certificate number")]/following-sibling::td/*//select'
    selectObject = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    select = Select(selectObject)
    frames = []
    for option in select.options:
        select.select_by_visible_text(option.text)
        xpath = '//*[text()="Inquiry"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Download file (.csv)
        xpath = '//*[@class="dl_icons download_csv"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Đọc file
        while True:
            downloadFile = first(listdir(bankObject.downloadFolder),lambda x: 'COSDATDQU' in x and 'download' not in x)
            if downloadFile:  # download xong -> có file
                break
            time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
        downloadTable = pd.read_csv(
            join(bankObject.downloadFolder,downloadFile),
            usecols=[1,2,3,4,6,7,8],
            skiprows=1,
            names=['AccountNumber','Currency','Balance','TermMonths','InterestRate','IssueDate','ExpireDate']
        )
        # TermMonths
        downloadTable['TermMonths'] = downloadTable['TermMonths'].map(lambda x: int(x.replace('M','')))
        # TermDays
        downloadTable['TermDays'] = downloadTable['TermMonths'] * 30
        # Interest Rate
        downloadTable['InterestRate'] = downloadTable['InterestRate'].map(lambda x: float(x.replace('%',''))) / 100
        # Issue Date
        downloadTable['IssueDate'] = downloadTable['IssueDate'].map(lambda x: dt.datetime.strptime(x,'%Y/%m/%d'))
        # Expire Date
        downloadTable['ExpireDate'] = downloadTable['ExpireDate'].map(lambda x: dt.datetime.strptime(x,'%Y/%m/%d'))
        # Date
        now = dt.datetime.now()
        if now.hour >= 12:
            d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        downloadTable.insert(0,'Date',d)
        # Bank
        downloadTable.insert(1,'Bank',bankObject.bank)
        # Interest Amount
        downloadTable['InterestAmount'] = None
        # Append
        frames.append(downloadTable)

    balanceTable = pd.concat(frames)
    return balanceTable



