from automation.finance import *

def runBIDV(bankObject,fromDate,toDate):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'EBK_BC_LICHSUGIAODICH' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Click Menu bar
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'menu-toggle-22'))).click()
    # Click "Tài khoản"
    bankObject.wait.until(EC.visibility_of_element_located((By.LINK_TEXT,'Vấn tin'))).click()
    # Click "Tiền gửi thanh toán"
    bankObject.wait.until(EC.visibility_of_element_located((By.LINK_TEXT,'Tiền gửi thanh toán'))).click()
    # Lấy số lượng tài khoản
    Accounts = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'change')))
    accountNumber = len(Accounts)
    # Click vào tài khoản đầu tiên
    xpath = f'//*[@class="change"]'
    Account = bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath)))
    Account.click()
    time.sleep(1) # chờ animation
    # Click "Lịch sử giao dịch"
    xpath = f'//*[@data-action="btDetailTransaction"]'
    Button = bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath)))
    Button.click()
    # Chọn tab "Thời gian"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Thời gian'))).click()
    # Lấy dữ liệu từng tài khoản
    records = []
    for d in pd.date_range(fromDate,toDate):
        # Từ ngày
        fromDateInput = bankObject.wait.until(EC.presence_of_element_located((By.ID,'fromDate')))
        fromDateInput.clear()
        fromDateInput.send_keys((d-dt.timedelta(days=1)).strftime('%d/%m/%Y'))
        # Đến ngày
        toDateInput = bankObject.wait.until(EC.presence_of_element_located((By.ID,'toDate')))
        toDateInput.clear()
        toDateInput.send_keys(d.strftime('%d/%m/%Y'))
        # Chọn từng tài khoản
        xpath = '//*[@aria-owns="accountNo_listbox"]'
        accountInput = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        accountInput.clear()
        i = 0
        while True:
            # Bấm mũi tên xuống để lấy từng TK (làm cách này để tránh lỗi)
            accountInput.send_keys(Keys.DOWN)
            value = accountInput.get_attribute('value')
            print(value)
            account = re.search('[0-9]{14}',value).group()
            time.sleep(0.5) # chờ pop up (nếu có)
            # Đóng pop up nếu có
            xpath = '//*[@data-bb-handler="ok"]'
            popupButtons = bankObject.driver.find_elements(By.XPATH,xpath)
            if popupButtons:
                popupButtons[0].click()
            # Download file excel
            bankObject.wait.until(EC.presence_of_element_located((By.ID,'btnExportExcel01'))).click()
            # Đọc file, record data
            while True:
                checkFunc = lambda x: 'EBK_BC_LICHSUGIAODICH' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile: # download xong -> có file
                    break
                time.sleep(1) # chưa download xong -> đợi thêm 1s nữa

            downloadData = pd.read_excel(join(bankObject.downloadFolder,downloadFile))
            rowLocMask = downloadData.eq('Dư cuối ngày').any(axis=1)
            colLocMask = downloadData.eq('Số dư').any(axis=0)
            balanceString = downloadData.loc[rowLocMask,colLocMask].squeeze()
            balance = float(balanceString.replace(',',''))
            records.append((d,bankObject.bank,account,balance,'VND'))
            # Xóa file
            os.remove(join(bankObject.downloadFolder,downloadFile))
            i += 1
            if i == accountNumber:
                break

    balanceTable = pd.DataFrame(
        data=records,
        columns=['Date','Bank','AccountNumber','Balance','Currency']
    )

    return balanceTable

# không CAPTCHA
def runVTB(bankObject,fromDate,toDate):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'lich-su-giao-dich' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Bắt đầu từ trang chủ
    xpath = '//*[@href="/"]'
    _, MainMenu = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    MainMenu.click()
    # Click menu "Tài khoản"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Tài khoản'))).click()
    # Click sub-menu "Danh sách tài khoản"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Danh sách tài khoản'))).click()
    # table Element
    tableElement, _, _ = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'MuiTableBody-root')))
    tableElement.find_element(By.LINK_TEXT,'Xem thêm').click()

    # Create function to clear input box and send dates as string
    def sendDate(element,d):
        action = ActionChains(bankObject.driver)
        action.click(element)
        time.sleep(0.5)
        action.key_down(Keys.CONTROL,element)
        action.send_keys_to_element(element,'a')
        action.key_up(Keys.CONTROL,element)
        action.send_keys_to_element(element,Keys.BACKSPACE)
        action.send_keys_to_element(element,d.strftime('%d/%m/%Y'))
        action.send_keys_to_element(element,Keys.ENTER)
        action.perform()

    records = []
    accountNumbers = filter(lambda x: len(x)==12,tableElement.text.split('\n'))
    for x in accountNumbers:
        bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,x))).click()
        fromDateInput,toDateInput = bankObject.wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'ant-picker-input')))
        for d in pd.date_range(fromDate,toDate):
            # Điền ngày
            sendDate(fromDateInput,d-dt.timedelta(days=5)) # set days lớn quá dễ bị VTB log out
            sendDate(toDateInput,d)
            while True:
                try:
                    bankObject.driver.find_element(By.CLASS_NAME,'btn-submit').click()
                    break
                except (Exception,):
                    pass
            # time.sleep(1) # chờ load data
            # Download file
            while True:
                try:
                    xpath = '//*[@src="/public/img/icon/icon-download.svg"]'
                    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
                    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Xuất Excel'))).click()
                    break
                except (Exception,):
                    pass

            # Đọc file, record data
            while True:
                checkFunc = lambda x: 'lich-su-giao-dich' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile: # download xong -> có file
                    break
                time.sleep(1) # chưa download xong -> đợi thêm 1s nữa

            downloadData = pd.read_excel(join(bankObject.downloadFolder,downloadFile))
            # Số tài khoản
            rowMask = downloadData.eq('Số dư cuối kỳ/Closing Balance:').any(axis=1)
            balanceString = downloadData.loc[rowMask,downloadData.columns[2]].squeeze()
            balance = float(balanceString)
            # Currency
            rowMask = downloadData.eq('Loại tiền tệ/Currency:').any(axis=1)
            currency = downloadData.loc[rowMask,downloadData.columns[2]].squeeze()
            records.append((d,bankObject.bank,x,balance,currency))
            # Xóa file
            os.remove(join(bankObject.downloadFolder,downloadFile))

        bankObject.driver.back()
        bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Xem thêm'))).click()

    balanceTable = pd.DataFrame(
        data = records,
        columns=['Date','Bank','AccountNumber','Balance','Currency']
    )
    # Xử lý lỗi EIB: Ngày nào ko có giao dịch thì số dư cuối kỳ = 0, bằng cách: forward fill
    for account in balanceTable['AccountNumber'].unique():
        mask = balanceTable['AccountNumber']==account
        balanceTable.loc[mask] = balanceTable.loc[mask].fillna(method='ffill')
    # Xử lý những thằng thật sự bằng 0
    balanceTable = balanceTable.fillna(0)

    return balanceTable

# có CAPTCHA
def runIVB(bankObject,fromDate,toDate):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'AccountTransacionHistory' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@class="logoimg"]'))).click()
    # Click tab "Tài khoản"
    xpath = '//*[@data-menu-id="1"]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click subtab "Sao kê tài khoản"
    bankObject.wait.until(EC.visibility_of_element_located((By.ID,'2_2'))).click()
    # Chọn "Tài khoản thanh toán" từ dropdown list
    bankObject.driver.switch_to.frame('mainframe')
    accountTypeInput = Select(bankObject.wait.until(EC.presence_of_element_located((By.ID,'selectedAccType'))))
    accountTypeInput.select_by_visible_text('Tài khoản Thanh toán')

    # Điền số tài khoản:
    accountElems = bankObject.driver.find_elements(By.XPATH,'//*[@id="account_list"]/option')
    options = [a.text for a in accountElems]
    records = []
    for option in options:
        time.sleep(1)
        account = option.split()[0]
        currency = option.split()[-1].replace(']','')
        if account not in ('1017816-066','1017816-069','1017816-068'):
            continue
        accountInput = Select(bankObject.wait.until(EC.presence_of_element_located((By.ID,'account_list'))))
        accountInput.select_by_visible_text(option)

        # Điền ngày
        for d in pd.date_range(fromDate,toDate):
            # Từ ngày
            fromDateInput = bankObject.driver.find_element(By.ID,'beginDate')
            fromDateInput.clear()
            fromDateInput.send_keys((d-dt.timedelta(days=5)).strftime('%d/%m/%Y'))
            # Đến ngày
            toDateInput = bankObject.driver.find_element(By.ID,'endDate')
            toDateInput.clear()
            toDateInput.send_keys(d.strftime('%d/%m/%Y'))
            # Click "Truy vấn"
            bankObject.driver.find_element(By.ID,'btnQuery').click()
            # Click "In sao kê Excel"
            bankObject.driver.execute_script(f'window.scrollTo(0,100000)')
            xpath = """//*[@onclick="downloadReport('excel');"]"""
            bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
            # Đọc file download
            while True:
                checkFunc = lambda x:'AccountTransacionHistory' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile:  # download xong -> có file
                    break
                time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
            downloadTable = pd.read_excel(
                join(bankObject.downloadFolder,downloadFile),
                usecols='B:D',
                skiprows=2,
            )
            rowMask = downloadTable.eq('Số dư cuối kỳ').any(axis=1)
            balanceString = downloadTable.loc[rowMask,downloadTable.columns[-1]].squeeze().replace(',','')
            balance = float(balanceString)
            records.append((d,bankObject.bank,account,balance,currency))
            # Xóa file download
            os.remove(join(bankObject.downloadFolder,downloadFile))
            bankObject.driver.execute_script(f'window.scrollTo(0,0)')

    balanceTable = pd.DataFrame(
        records,
        columns=['Date','Bank','AccountNumber','Balance','Currency']
    )

    return balanceTable

# có CAPTCHA
def runVCB(bankObject,fromDate,toDate):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'Vietcombank_Account_Statement' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Bắt đầu từ trang chủ
    _, homeButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'icon-home')))
    homeButton.click()
    # Lấy danh sách đường dẫn vào tài khoản thanh toán
    xpath = '//*[@id="dstkdd-tbody"]//td/a'
    accountElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    URLs = [e.get_attribute('href') for e in accountElems]

    # Create function to clear input box and send dates as string
    def sendDate(element,d):
        action = ActionChains(bankObject.driver)
        action.click(element)
        time.sleep(0.5)
        action.key_down(Keys.CONTROL,element)
        action.send_keys_to_element(element,'a')
        action.key_up(Keys.CONTROL,element)
        action.send_keys_to_element(element,Keys.BACKSPACE)
        action.send_keys_to_element(element,d.strftime('%d/%m/%Y'))
        action.send_keys_to_element(element,Keys.ENTER)
        action.perform()

    records = []
    for URL in URLs:
        bankObject.driver.get(URL)
        time.sleep(1) # chờ để hiện số tài khoản (bắt buộc)
        for d in pd.date_range(fromDate,toDate):
            # Điền ngày
            startDateInput = bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'startDate')))
            bankObject.driver.execute_script(f'window.scrollTo(0,500)')
            sendDate(startDateInput,d)
            endDateInput = bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'endDate')))
            sendDate(endDateInput,d)
            # Xuống cuối trang
            bankObject.driver.execute_script(f'window.scrollTo(0,100000)')
            # Click "Xem sao kê"
            bankObject.wait.until(EC.visibility_of_element_located((By.ID,'TransByDate'))).click()
            # Click "Xuất file"
            exportButton = bankObject.wait.until(EC.visibility_of_element_located((By.ID,'ctl00_Content_TransactionDetail_ExportButton')))
            while True: # do nút button bị cover bởi một layer loading
                try:
                    exportButton.click()
                    break
                except (ElementClickInterceptedException,):
                    time.sleep(0.5)
            # Lên đầu trang
            bankObject.driver.execute_script(f'window.scrollTo(0,0)')

            # Đọc file download
            while True:
                checkFunc = lambda x: 'Vietcombank_Account_Statement' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile:  # download xong -> có file
                    renameFile = downloadFile.replace('xls','csv')
                    os.rename(join(bankObject.downloadFolder,downloadFile),join(bankObject.downloadFolder,renameFile))
                    downloadFile = renameFile
                    break
                time.sleep(0.5)  # chưa download xong -> đợi thêm 1s nữa

            with open(join(bankObject.downloadFolder,downloadFile),'rb') as f:
                htmlString = f.read()

            soup = BeautifulSoup(htmlString,'html5lib')
            account = soup.find(text='Số tài khoản:').find_next().text
            balanceString = soup.find(text='Số dư cuối kỳ').find_next().text.replace(',','')
            if balanceString != '':
                balance = float(balanceString)
            else:
                balance = None
            currency = soup.find(text='Loại tiền:').find_next().text
            records.append((d,bankObject.bank,account,balance,currency))
            # delete download file
            os.remove(join(bankObject.downloadFolder,downloadFile))

    balanceTable = pd.DataFrame(
        records,
        columns=['Date','Bank','AccountNumber','Balance','Currency']
    )

    return balanceTable

# không CAPTCHA
def runEIB(bankObject,fromDate,toDate):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'LichSuTaiKhoan' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Click Menu Tài khoản --> Tiền gửi thanh toán
    action = ActionChains(bankObject.driver)
    xpath = '//*[@class="navigation-menu"]/li[2]/a'
    menuAccount = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    action.move_to_element(menuAccount)
    xpath = '//*[@href="/KHDN/corp/account/payment"]'
    currentAccount = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    action.click(currentAccount)
    action.perform()
    time.sleep(3) # Chờ load xong
    # Lấy danh sách tài khoản
    xpath = '//tbody/tr/th/a'
    accountElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    accounts = [e.text for e in accountElems]
    # Lấy danh sách currency
    xpath = '//*[text()="VND" or text()="USD"]'
    currencyElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    currencies = [e.text for e in currencyElems]
    # Lấy số dư tài khoản
    records = []
    for account, currency in zip(accounts,currencies):
        Account = bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,account)))
        Account.click()
        # Click Xem lịch sử tài khoản
        xpath = '//*[@class="modal-footer"]/button'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click() # nút đầu tiên
        for d in pd.date_range(fromDate,toDate):
            xpath = '//*[@name="fromDate" or @name="toDate"]/div/input'
            fromDateInput,toDateInput = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
            # Từ ngày
            fromDateInput.clear()
            fromDateInput.send_keys(f"{(d-dt.timedelta(days=10)).strftime('%d/%m/%Y')} 00:00:00")
            # Đến ngày
            toDateInput.clear()
            toDateInput.send_keys(f"{d.strftime('%d/%m/%Y')} 23:59:59")
            # Click "Tìm kiếm"
            _, searchButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'btn-primary')))
            while True:
                try:
                    searchButton.click()
                    time.sleep(0.5)
                    break
                except (ElementClickInterceptedException,):
                    pass
            # Check xem có giao dịch không, vì không có giao dịch sẽ ko có file download
            flags = bankObject.driver.find_elements(By.XPATH,'//*[text()=" Không tìm thấy thông tin. "]')
            if flags:
                records.append((d,'EIB',account,np.nan,currency))
                continue
            # Download file
            while True:
                try:
                    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Tải Excel'))).click()
                    break
                except (ElementClickInterceptedException,StaleElementReferenceException):
                    pass
            # Đọc file download
            while True:
                checkFunc = lambda x:'LichSuTaiKhoan' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile:  # download xong -> có file
                    break
                time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa

            downloadData = pd.read_excel(join(bankObject.downloadFolder,downloadFile))
            # Số dư cuối kỳ
            rowMask = downloadData.eq('Số dư cuối kỳ / Current Balance ').any(axis=1)
            balance = float(downloadData.loc[rowMask,downloadData.columns[-2]].squeeze())
            records.append((d,bankObject.bank,account,balance,currency))
            # Xóa file
            os.remove(join(bankObject.downloadFolder,downloadFile))

        bankObject.driver.back()

    balanceTable = pd.DataFrame(
        records,
        columns=['Date','Bank','AccountNumber','Balance','Currency']
    )

    # Xử lý lỗi EIB: Ngày nào ko có giao dịch thì số dư cuối kỳ = 0, bằng cách: forward fill
    for account in balanceTable['AccountNumber'].unique():
        mask = balanceTable['AccountNumber']==account
        balanceTable.loc[mask] = balanceTable.loc[mask].fillna(method='ffill')
    # Xử lý những thằng thật sự bằng 0
    balanceTable = balanceTable.fillna(0)

    return balanceTable

# không CAPTCHA
def runOCB(bankObject,fromDate,toDate):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """
    
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'TransactionHistory' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Bắt đầu từ main menu
    bankObject.wait.until(EC.presence_of_element_located((By.ID,'main-menu-icon'))).click()
    # Click Thông tin tài khoản --> Sao kê tài khoản
    xpath = '//*[@class="ahref"]/*[@class="accounts-icon"]'
    _, accountButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    accountButton.click()
    xpath = '//*[@id="side-nav"]/ul/li[2]/div[3]/span'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    time.sleep(10) # chờ chuyển trang
    # Lấy danh sách tài khoản
    i = 0
    records = []
    while True:
        # Click dropdown button
        bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'indicator'))).click()  # dropdown button
        options = bankObject.wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'rb-account-select__account-desc-row')))
        if i == len(options):
            break
        # Nhập số tài khoản
        Entry = options[i]; i += 1
        accountNumber = Entry.text.split('\n')[0].replace(' ','')
        Entry.click()
        # Click "Tìm kiếm nâng cao"
        xpath = '//*[@ng-click="toggleAdvancedSearch()"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        time.sleep(1)
        # Tick chọn tính năng khoảng thời gian
        xpath = '//*[@class="bd-radio-option__marker"]'
        _, rangeButton = bankObject.wait.until(EC.visibility_of_all_elements_located((By.XPATH,xpath)))
        rangeButton.click()
        for d in pd.date_range(fromDate,toDate):
            # Từ ngày
            fromDateInput = bankObject.wait.until(EC.presence_of_element_located((By.NAME,'dateFromInput')))
            fromDateInput.clear()
            fromDateInput.send_keys((d-dt.timedelta(days=10)).strftime('%d.%m.%Y'))
            # Đến ngày
            toDateInput = bankObject.wait.until(EC.presence_of_element_located((By.NAME,'dateToInput')))
            toDateInput.clear()
            toDateInput.send_keys(d.strftime('%d.%m.%Y'))
            # Lấy số dư cuối kỳ
            while True: # click đến khi được thì thôi
                try:
                    searchButton = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@bd-id="search_button_mobile"]')))
                    searchButton.click()
                    break
                except (ElementClickInterceptedException,):
                    pass
            time.sleep(1) # chờ load data
            # Click Tải về chi tiết
            xpath = '//*[text()="Tải về"]'
            while True:
                try:
                    bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
                    break
                except (ElementNotInteractableException,):
                    time.sleep(1) # click bị fail thì chờ 1s rồi click lại
            time.sleep(1) # chờ animation
            xpath = '//*[text()="Tập tin XLS"]'
            bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
            # Đọc file download
            while True:
                checkFunc = lambda x:'TransactionHistory' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile:  # download xong -> có file
                    break
                time.sleep(0.5)  # chưa download xong -> đợi thêm 0.5s nữa
            downloadTable = pd.read_excel(
                join(bankObject.downloadFolder,downloadFile),
                usecols='C:D',
                skiprows=3,
            )
            # Số dư
            rowMask = downloadTable.eq('Số dư cuối kỳ/ Ending Balance').any(axis=1)
            balanceString = downloadTable.loc[rowMask,downloadTable.columns[-1]].squeeze()
            balance = float(balanceString.replace(',',''))
            # Currency
            rowMask = downloadTable.eq('Loại tiền/ Currency').any(axis=1)
            currency = downloadTable.loc[rowMask,downloadTable.columns[-1]].squeeze()
            records.append((d,bankObject.bank,accountNumber,balance,currency))
            # Xóa file
            os.remove(join(bankObject.downloadFolder,downloadFile))
            # Scroll lên đầu trang
            bankObject.driver.execute_script(f'window.scrollTo(0,0)')
            time.sleep(0.5) # tránh click khi chưa scroll xong

        # Click "Ẩn tìm kiếm nâng cao"
        xpath = '//*[@ng-click="toggleAdvancedSearch()"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

    balanceTable  = pd.DataFrame(
        records,
        columns=['Date','Bank','AccountNumber','Balance','Currency'],
    )

    return balanceTable

# không CAPTCHA
def runTCB(bankObject,fromDate,toDate):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'enquiry' in file:
            os.remove(join(bankObject.downloadFolder,file))
    # Check tab "Tông tin tài khoản" có bung chưa (đã được click trước đó), phải bung rồi mới hiện tab "Truy vấn giao dịch tài khoản"
    xpath = '//*[contains(text(),"Truy vấn giao dịch tài khoản")]'
    queryElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    if not queryElement.is_displayed(): # nếu chưa bung
        # Click "Thông tin tài khoản"
        xpath = '//*[contains(text(),"Thông tin tài khoản")]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    # Click "Truy vấn giao dịch tài khoản"
    queryElement.click()
    time.sleep(3)
    # Click dropdown button
    bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'dropdown_button'))).click()
    # Lấy danh sách tài khoản
    xpath = '//*[contains(@id,"dropDownRow")]'
    accountElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    infoTexts = [e.text.split('-')[0] for e in accountElems]
    infoZip = [info.split() for info in infoTexts]
    records = []
    for account, currency in infoZip:
        for d in pd.date_range(fromDate,toDate):
            # Điền số tài khoản
            accountBox = bankObject.wait.until(EC.presence_of_element_located((By.ID,'fieldName:FIELD.1')))
            accountBox.clear()
            accountBox.send_keys(account)
            # Từ ngày
            fromDateBox = bankObject.wait.until(EC.presence_of_element_located((By.ID,'fieldName:DATE.FIELD.1')))
            fromDateBox.clear()
            fromDateBox.send_keys((d-dt.timedelta(days=10)).strftime('%d/%m/%Y'))
            # Đến ngày
            toDateBox = bankObject.wait.until(EC.presence_of_element_located((By.ID,'fieldName:DATE.FIELD.2')))
            toDateBox.clear()
            toDateBox.send_keys(d.strftime('%d/%m/%Y'))
            # Click "Thực hiện"
            xpath = '//*[@title="Thực hiện"]'
            bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
            # Tải file
            bankObject.wait.until(EC.presence_of_element_located((By.ID,'nsavebtn'))).click()
            # Đọc file download
            while True:
                checkFunc = lambda x: 'enquiry' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile:  # download xong -> có file
                    break
                time.sleep(0.5)  # chưa download xong -> đợi thêm 0.5s nữa
            balanceString = pd.read_csv(
                join(bankObject.downloadFolder,downloadFile),
                skiprows=4,
                usecols=['So du'],
                nrows=1,
            ).squeeze() # Số dư của dòng sao kê đầu tiên
            if balanceString == balanceString: # not nan -> có phát sinh sao kê
                balance = float(balanceString.replace(',',''))
            else:
                balance = None # cho compatible với database NULL
            # Xóa file
            os.remove(join(bankObject.downloadFolder,downloadFile))
            records.append((d,bankObject.bank,account,balance,currency))
            # Back ra trang tài khoản
            bankObject.wait.until(EC.presence_of_element_located((By.ID,'nbackbtn'))).click()

    balanceTable = pd.DataFrame(
        records,
        columns=['Date','Bank','AccountNumber','Balance','Currency'],
    )
    # Xử lý: Ngày nào ko có giao dịch thì số dư cuối kỳ = 0, bằng cách: forward fill
    for account in balanceTable['AccountNumber'].unique():
        mask = balanceTable['AccountNumber']==account
        balanceTable.loc[mask] = balanceTable.loc[mask].fillna(method='ffill')
    # Xử lý những thằng thật sự bằng 0
    balanceTable = balanceTable.fillna(0)

    return balanceTable