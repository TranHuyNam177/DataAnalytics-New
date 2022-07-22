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
    # Check tab "Tài khoản" có bung chưa (đã được click trước đó), phải bung rồi mới hiện tab "Danh sách tài khoản"
    xpath = '//*[text()="Danh sách tài khoản"]'
    queryElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    if not queryElement.is_displayed(): # nếu chưa bung
        # Click "Thông tin tài khoản"
        xpath = '//*[text()="Tài khoản"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        time.sleep(1) # chờ animation
    queryElement.click()
    time.sleep(1)
    # Download file
    while True:
        try:
            xpath = '//*[contains(text(),"Tài khoản tiền gửi có kỳ hạn")]//ancestor::div[1]//following-sibling::div//*/img[contains(@src,"icon-download")]'
            bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
            bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Xuất Excel'))).click()
            break
        except (Exception,):
            pass

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
        usecols='B,D:I,L',
        dtype={'AccountNumber':object,'Balance':np.float64,'InterestAmount':np.float64}
    )
    downloadTable = downloadTable.dropna(subset=['AccountNumber'])
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
    time.sleep(1) # chờ animation
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
        time.sleep(3)  # chờ để hiện số tài khoản (bắt buộc)
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
        termDays = (expireDate-issueDate).days
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
    bankObject.driver.switch_to.default_content()
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
    balanceTable = pd.read_excel(
        join(bankObject.downloadFolder,downloadFile),
        usecols='A,C,E:J',
    )
    checkTable = balanceTable.loc[balanceTable.iloc[:,0].map(lambda x: isinstance(x,str) and 'Time Deposit Account' in x)]
    if checkTable.empty:
        return pd.DataFrame()
    startRow =  checkTable.index[0] + 2
    checkTable = balanceTable.loc[balanceTable.iloc[:,0].map(lambda x: isinstance(x,str) and 'Time Deposit Total' in x)]
    if checkTable.empty:
        return pd.DataFrame()
    endRow = checkTable.index[0]
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
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
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
    balanceTable['InterestAmount'] = balanceTable['TermDays'] * balanceTable['InterestRate'] * balanceTable['Balance'] / 365
    # Xóa file
    os.remove(join(bankObject.downloadFolder,downloadFile))

    return balanceTable


def runFIRST(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """
    now = dt.datetime.now()
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if re.search(r'\bCOSDATDQU\d+\b',file):
            os.remove(join(bankObject.downloadFolder,file))
    # Click "Account Inquiry"
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('iFrameID')
    xpath = '//*[contains(text(),"Account Inquiry")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click "Time Deposit Detail"
    time.sleep(2)
    xpath = '//*[contains(text(),"Time Deposit Detail")]'
    bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath))).click()
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
            checkFunc = lambda x: (re.search(r'\bCOSDATDQU\d+\b',x) is not None) and ('download' not in x)
            downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
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
        if now.hour >= 12:
            d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        downloadTable.insert(0,'Date',d)
        # Bank
        downloadTable.insert(1,'Bank',bankObject.bank)
        # Interest Amount
        downloadTable['InterestAmount'] = downloadTable['TermDays'] * downloadTable['InterestRate'] * downloadTable['Balance'] / 365
        # Append
        frames.append(downloadTable)

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()
    balanceTable = pd.concat(frames)
    return balanceTable


def runTCB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    bankObject.driver.switch_to.default_content()

    # Check tab "Đầu tư" có bung chưa (đã được click trước đó), phải bung rồi mới hiện tab "Hợp đồng tiền gửi"
    xpath = '//*[contains(text(),"Hợp đồng tiền gửi")]'
    queryElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    if not queryElement.is_displayed(): # nếu chưa bung
        # Click Đầu tư
        xpath = '//*[contains(text(),"Đầu tư")]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click Hợp đồng tiền gửi
    queryElement.click()
    time.sleep(3)
    # Lấy toàn bộ dòng
    xpath = '//*[@class="enquirydata wrap_words"]/tbody/tr[*]'
    rowElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))[1:]  # bỏ dòng header
    records = []
    for element in rowElements:
        now = dt.datetime.now()
        if now.hour >= 12:
            d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        elementString = element.text
        # Số tài khoản
        account = re.search(r'\b[A-Z]{2}[0-9]{10}\b',elementString).group()
        # Ngày hiệu lực, Ngày đáo hạn
        issueDateText, expireDateText = re.findall(r'\b\d{2}/\d{2}/\d{4}\b',elementString)
        issueDate = dt.datetime.strptime(issueDateText,'%d/%m/%Y')
        expireDate = dt.datetime.strptime(expireDateText,'%d/%m/%Y')
        # Term Days
        termDays = (expireDate-issueDate).days
        # Term Months
        termMonths = (expireDate.year-issueDate.year)*12 + expireDate.month-issueDate.month
        # Lãi suất
        iText = re.search(r'\b\d+\.\d+\b',elementString).group()
        iRate = float(iText)/100
        # Currency
        currency = re.search(r'VND|USD',elementString).group()
        # Số dư tiền
        balanceString = re.search(r'(\bVND|USD)\s*.*,000\b',elementString).group()
        balanceString = re.sub('(VND|USD)','',balanceString).strip().replace(',','')
        balance = float(balanceString)
        # Số tiền Lãi
        interest = termDays * iRate * balance / 365
        records.append((d,bankObject.bank,account,termDays,termMonths,iRate,issueDate,expireDate,balance,interest,currency))

    balanceTable = pd.DataFrame(
        data=records,
        columns=['Date','Bank','AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    )
    return balanceTable


def runMEGA(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """
    now = dt.datetime.now()
    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    frameElement = bankObject.wait.until(EC.presence_of_element_located((By.ID,"ifrm")))
    bankObject.driver.switch_to.frame(frameElement)
    # Click Accounts
    xpath = '//*[contains(text(),"Accounts")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click Deposit
    xpath = '//*[contains(text(),"Deposit")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click Deposit balance
    xpath = '//*[contains(text(),"Deposit balance")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Switch frame
    bankObject.driver.switch_to.frame('frame1')
    # Get data in e-Time Deposit A/C
    xpath = '//*[@class="tb5" and contains(@id,"time_DataGridBody")]/tbody/tr[*]'
    rowElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))[1:]  # bỏ dòng header
    records = []
    for element in rowElements:
        if now.hour >= 12:
            d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        elementString = element.text
        # Số tài khoản
        account = re.search('\d{9}',elementString).group()
        # Ngày hiệu lực, Ngày đáo hạn
        expireDateText, issueDateText = re.findall('\d{4}/\d{2}/\d{2}',elementString)
        issueDate = dt.datetime.strptime(issueDateText,'%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText,'%Y/%m/%d')
        # Term Days
        termDays = (expireDate-issueDate).days
        # Term Months
        termMonths = (expireDate.year-issueDate.year)*12 + expireDate.month-issueDate.month
        # Lãi suất
        iText = elementString.split()[-1]
        iRate = float(iText) / 100
        # Currency
        currency = re.search('VND|USD',elementString).group()
        # Số dư tiền
        balanceString = re.search(r'\b\d+,.*,\d{3}\b',elementString).group()
        balance = float(balanceString.replace(',',''))
        # Số tiền Lãi
        interest = termDays * iRate * balance / 365
        records.append((d,bankObject.bank,account,termDays,termMonths,iRate,issueDate,expireDate,balance,interest,currency))

    balanceTable = pd.DataFrame(
        data=records,
        columns=['Date','Bank','AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    )
    return balanceTable


def runSINOPAC(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """
    now = dt.datetime.now()
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if re.search(r'\bCOSDATDQU_\d+_\d+\b',file):
            os.remove(join(bankObject.downloadFolder,file))
    # Click Account Inquiry
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('indexFrame')
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'MENU_CAO'))).click()
    # Click Deposit inquiry
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'MENU_CAO001'))).click()
    time.sleep(1)
    # Term Deposit Inquiry
    bankObject.wait.until(EC.visibility_of_element_located((By.ID,'MENU_COSDATDQU'))).click()

    bankObject.driver.switch_to.frame('mainFrame')
    xpath = '//*[contains(@id,"accountCombo_input")]/option'
    accountElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))[1:] # bỏ ===SEL===
    options = [l.text for l in accountElements]
    frames = []
    for option in options:
        xpath = '//*[contains(@id,"accountCombo_input")]'
        dropDown = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        select = Select(dropDown)
        select.select_by_visible_text(option)
        # Click "Search"
        xpath = '//*[@type="submit"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        time.sleep(3) # chờ load data
        # Click download file csv
        xpath = '//*[contains(@class,"download_csv")]'
        downloadButtons = bankObject.driver.find_elements(By.XPATH,xpath)
        if not downloadButtons:
            continue
        downloadButtons[0].click()
        # Đọc file download
        while True:
            checkFunc = lambda x: (re.search(r'\bCOSDATDQU_\d+_\d+\b',x) is not None) and ('download' not in x)
            downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
            if downloadFile:  # download xong -> có file
                break
            time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
        frame = pd.read_csv(
            join(bankObject.downloadFolder,downloadFile),
            usecols=[1,3,4,5,7],
            names=['Currency','Balance','IssueDate','ExpireDate','InterestRate'],
            skiprows=1,
            dtype={
                'Currency': object,
                'Balance': np.float64,
                'IssueDate': object,
                'ExpireDate': object,
                'InterestRate': np.float64,
            }
        )
        frame['IssueDate'] = pd.to_datetime(frame['IssueDate'],format='%Y/%m/%d')
        frame['ExpireDate'] = pd.to_datetime(frame['ExpireDate'],format='%Y/%m/%d')
        # Account
        account = option.split()[0]
        frame.insert(0,'AccountNumber',account)
        # TermDays
        frame['TermDays'] = (frame['ExpireDate']-frame['IssueDate']).dt.days
        # TermMonths
        frame['TermMonths'] = (frame['TermDays']/30).round()
        # InterestRate
        frame['InterestRate'] = frame['InterestRate'] / 100
        # InterestAmount
        frame['InterestAmount'] = frame['TermDays'] * frame['InterestRate'] * frame['Balance'] / 365
        # Append
        frames.append(frame)

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()
    balanceTable = pd.concat(frames)

    # Date
    if now.hour >= 12:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0,'Date',d)
    # Bank
    balanceTable.insert(1,'Bank',bankObject.bank)

    return balanceTable


def runESUN(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if re.search(r'\bCOSDATDQU_\d+\b',file):
            os.remove(join(bankObject.downloadFolder,file))

    # Show menu Deposits
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('mainFrame')
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'menuIndex_1'))).click()
    # Click Time Deposit Detail Enquiry
    xpath = '//*[contains(text(),"Time Deposit Detail Enquiry")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click Enquire
    xpath = '//*[contains(text(),"Enquire")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Download file csv
    bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'download_csv'))).click()
    # Đọc file download
    while True:
        checkFunc = lambda x: (re.search(r'\bCOSDATDQU_\d+\b',x) is not None) and ('download' not in x)
        downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
        if downloadFile:  # download xong -> có file
            break
        time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
    balanceTable = pd.read_csv(
        join(bankObject.downloadFolder, downloadFile),
        skiprows=1,
        names=['AccountNumber','IssueDate','ExpireDate','Currency','Balance','InterestRate'],
        usecols=[1,3,4,5,6,7],
        dtype={
            'AccountNumber': object,
            'IssueDate': object,
            'ExpireDate':object,
            'Currency': object,
            'Balance': object,
        }
    )
    balanceTable['IssueDate'] = pd.to_datetime(balanceTable['IssueDate'],format='%Y/%m/%d')
    balanceTable['ExpireDate'] = pd.to_datetime(balanceTable['ExpireDate'],format='%Y/%m/%d')
    balanceTable['Balance'] = balanceTable['Balance'].str.replace(',','').astype(float)
    balanceTable['Bank'] = bankObject.bank
    # Date
    now = dt.datetime.now()
    if now.hour >= 12:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable['Date'] = d
    # TermDays
    balanceTable['TermDays'] = (balanceTable['ExpireDate']-balanceTable['IssueDate']).dt.days
    # Term months
    balanceTable['TermMonths'] = (balanceTable['TermDays']/30).round()
    # InterestRate
    balanceTable['InterestRate'] = balanceTable['InterestRate'] / 100
    balanceTable['InterestAmount'] = balanceTable['TermDays'] * balanceTable['InterestRate'] * balanceTable['Balance'] / 365

    return balanceTable


def runHUANAN(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """
    now = dt.datetime.now()
    # Click Check Account Details
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('left')
    xpath = "//*[contains(text(),'Check account details')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click Check balance -> Time Depoist
    bankObject.driver.switch_to.default_content()
    bankObject.driver.switch_to.frame('main')
    xpath = '//*[@id="navbar"]/*/a[contains(@title,"Check balance")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    xpath = '//*[@id="navbar"]/*/a[contains(@title,"Check balance")]//following-sibling::ul/*/a[contains(@title,"Time deposit")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Lấy danh sách chi nhánh
    xpath = '//*[@name="UNIT"]/option'
    optionElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    optionStrings = [element.text for element in optionElements]

    rowStrings = []
    for optionString in optionStrings:
        # Chọn từng chi nhánh
        xpath = '//*[@name="UNIT"]'
        dropdownElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        select = Select(dropdownElement)
        select.select_by_visible_text(optionString)
        # Click Submit
        xpath = '//*[contains(@class,"BTN")]/*[contains(text(),"Submit")]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Chờ load data
        while True:
            xpath = '//*[text()="Interest Rate"]'
            if bankObject.driver.find_elements(By.XPATH,xpath):
                break
            time.sleep(1)
        # Lấy Row String
        xpath = '//td[@colspan>5]//*[contains(@class,"Table_content")]'
        recordElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
        recordStrings = [element.text for element in recordElements]
        # Extend Row String
        rowStrings.extend(recordStrings)
        # Back về trang chọn chi nhánh
        bankObject.driver.back()
        bankObject.driver.switch_to.default_content()
        bankObject.driver.switch_to.frame('main')
        while True:
            xpath = '//*[contains(@class,"BTN")]/*[contains(text(),"Submit")]'
            submitButton = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            if submitButton.is_displayed():
                break
            time.sleep(1)

    records = []
    for rowString in rowStrings:
        # Account Number
        accountNumber = re.search(r'^[\d*-]+\b',rowString).group()
        # Currency
        currency = re.search(r'\bVND|USD\b',rowString).group()
        # Balance
        balanceString = re.search(r'\b\d+,[\d,]+\.\d{2}\b',rowString).group()
        balance = float(balanceString.replace(',',''))
        # Issue Date, Expire Date
        dateStrings = re.findall(r'\b\d{4}-\d{2}-\d{2}',rowString)
        dates = sorted([dt.datetime.strptime(dateString,'%Y-%m-%d') for dateString in dateStrings])
        issueDate, expireDate = dates[-2:]
        # Term month
        termMonthsString = re.search(r'\b(\d+)(\smonth)',rowString).group(1)
        termMonths = int(termMonthsString)
        # Term day
        termDays = (expireDate-issueDate).days
        # Interest rate
        interestRateString = re.search(r'\b\d{1,2}\.\d+\b',rowString).group()
        interestRate = float(interestRateString) / 100
        # Interest amount
        interestAmount = termDays * interestRate * balance / 365
        # Append
        records.append((accountNumber,termDays,termMonths,interestRate,issueDate,expireDate,balance,interestAmount,currency))

    balanceTable = pd.DataFrame(
        records,
        columns=['AccountNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate','Balance','InterestAmount','Currency']
    )
    # Date
    if now.hour >= 12:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now-dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0,'Date',d)
    # Bank
    balanceTable.insert(1,'Bank',bankObject.bank)

    return balanceTable

