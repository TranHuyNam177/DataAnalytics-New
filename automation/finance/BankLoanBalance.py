from automation.finance import *

def runIVB(bankObject):
    """
    :param bankObject: Bank Object (đã login)
    """
    now = dt.datetime.now()
    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'logoimg'))).click()
    # Click tab "Tài khoản"
    xpath = '//*[@data-menu-id="1"]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    # Click subtab "Thông tin tài khoản"
    bankObject.wait.until(EC.visibility_of_element_located((By.ID, '1_1'))).click()
    # Click "Thông tin vay"
    bankObject.driver.switch_to.frame('mainframe')
    while True:
        xpath = '//*[@id="fd_layer_pan"]/span'
        bankObject.wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
        xpath = '//*[@id="informAccount"]/tbody/tr[*]/td/a'
        contractElements = bankObject.driver.find_elements(By.XPATH, xpath)
        URLs = [e.get_attribute('href') for e in contractElements]
        if contractElements[0].text != '':
            break
    records = []
    for URL in URLs:
        bankObject.driver.get(URL)
        time.sleep(1)  # chờ để hiện số tài khoản (bắt buộc)
        # Danh sách các trường thông tin cần lấy:
        infoList = ['Số Hợp đồng', 'Số tiền vay', 'Dư nợ hiện tại', 'Loại tiền', 'Ngày giải ngân',
                    'Ngày đến hạn', 'Lãi suất']
        # Tạo list chứa tên cột tương đương với các cột trong Database
        infoColumnName = ['ContractNumber', 'Amount', 'Remaining', 'Currency', 'IssueDate', 'ExpireDate',
                          'InterestRate']
        finalDict = dict()
        for info, col_name in zip(infoList, infoColumnName):
            xpath = f"//*[contains(text(),'{info}')]//following-sibling::td"
            infoElements = bankObject.driver.find_elements(By.XPATH, xpath)
            if infoElements:
                finalDict[col_name] = infoElements[0].text
        df = pd.DataFrame().append(finalDict, ignore_index=True)
        # Date
        if now.hour >= 12:
            d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
        else:
            d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                     microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
        df['Date'] = d
        df['Bank'] = bankObject.bank
        df['Amount'] = float(df['Amount'].str.replace(',', ''))
        df['Remaining'] = float(df['Remaining'].str.replace(',', ''))
        df['Paid'] = df['Amount'] - df['Remaining']
        df['InterestRate'] = round(float(df['InterestRate'].str.split('%').str[0]) / 100, 5)
        df['IssueDate'] = pd.to_datetime(df['IssueDate'], format='%d/%m/%Y')
        df['ExpireDate'] = pd.to_datetime(df['ExpireDate'], format='%d/%m/%Y')
        # TermDays
        df['TermDays'] = (df['ExpireDate'] - df['IssueDate']).dt.days
        # Term months
        month = df['ExpireDate'].dt.month - df['IssueDate'].dt.month
        year = df['ExpireDate'].dt.year - df['IssueDate'].dt.year
        df['TermMonths'] = year * 12 + month
        df['InterestAmount'] = df['TermDays'] * (df['InterestRate'] / 360) * df['Amount']
        cols = ['Date', 'Bank', 'ContractNumber', 'TermDays', 'TermMonths', 'InterestRate', 'IssueDate', 'ExpireDate',
                'Amount', 'Paid', 'Remaining', 'InterestAmount', 'Currency']
        table = df[cols]
        records.append(table)
    balanceTable = pd.concat(records, ignore_index=True)
    return balanceTable

def runMEGA(bankObject):
    now = dt.datetime.now()
    # Click Accounts
    xpath = '//*[contains(text(),"Accounts")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    # Click Balance overview
    xpath = '//*[contains(text(),"Balance overview")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    # Reload frame
    bankObject.driver.switch_to.frame('frame1')
    # Get data in Balance overview -> Loan guarantee accounts
    xpath = '//*[contains(@id,"load_DataGridBody")]/tbody/tr'
    rowElements = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))[
                  1:]  # bỏ dòng đầu tiên vì là header
    records = []
    # Xử lý data trong tab Balance Overview
    if now.hour >= 12:
        d = now.replace(hour=0,minute=0,second=0,microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0,minute=0,second=0,microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    for element in rowElements:
        elementString = element.text
        accountNumber = re.search(r'[A-Z]{3} [0-9]{4}-[0-9]{4}-[0-9]{10}', elementString).group().replace('000000', '')
        # Ngày hiệu lực, Ngày đáo hạn
        issueDateText, expireDateText = re.findall(r'\b\d{4}/\d{2}/\d{2}\b', elementString)
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = int((expireDate.year - issueDate.year) * 12 + (expireDate.month - issueDate.month))
        # Lãi suất
        iText = elementString.split(expireDateText)[1].split()[0]
        iRate = round(float(iText) / 100, 5)
        # Currency
        currency = re.search(r'VND|USD', elementString).group()
        # Số tiền vay
        amountString = elementString.split(iText)[1].split()[0]
        amount = float(amountString.replace(',', ''))
        # Số tiền Lãi
        interest = termDays * (iRate / 360) * amount
        records.append([d, bankObject.bank, accountNumber, termDays, termMonths, iRate, issueDate,
                        expireDate, amount, interest, currency])
    # Chuyển sang tab Financing
    # Reload frame
    while True:
        try:
            bankObject.driver.switch_to.parent_frame()
            bankObject.driver.switch_to.frame(bankObject.driver.find_element(By.XPATH, '//*[@id="ifrm"]'))
            break
        except bankObject.ignored_exceptions:
            continue
    # Click Financing
    xpath = "//*[contains(text(), 'Financing')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    # Click Loan details
    xpath = "//*[contains(text(), 'Loan details')]"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    # Reload frame
    bankObject.driver.switch_to.frame('frame1')
    new_list = []
    for element in records:
        xpath = '//*[contains(@id,"account")]'
        # select Account
        dropDownList = bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        # phần tử thứ 2 trong element là account
        Select(dropDownList).select_by_visible_text(element[2])
        # Input start date in Inquiry Period
        xpath = '//*[contains(@id,"startDate")]'
        startDateInput = bankObject.driver.find_element(By.XPATH, xpath)
        startDateInput.clear()
        startDate = element[6].strftime('%Y/%m/%d')
        startDateInput.send_keys(startDate)
        # Click Inquire
        xpath = "//*[contains(text(), 'Inquire')]"
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
        # Get Contract Number
        xpath = f"//*[text()='{startDate}']//following-sibling::td[1]"
        contractNumber = bankObject.driver.find_element(By.XPATH, xpath).text
        # Get Remaining
        xpath = f"//*[text()='{contractNumber}']//parent::tr"
        remaining = float(bankObject.driver.find_elements(By.XPATH, xpath)[-1].text.split(' ')[-1].replace(',', ''))
        # Calculate Paid
        paid = element[8] - remaining
        # Delete account number in list
        del element[2]
        # Add contract number, paid, remaining to list
        element.extend([contractNumber, paid, remaining])
        new_list.append(element)
    balanceTable = pd.DataFrame(
        data=new_list,
        columns=['Date', 'Bank', 'TermDays', 'TermMonths', 'InterestRate', 'IssueDate', 'ExpireDate',
                 'Amount', 'InterestAmount', 'Currency', 'ContractNumber', 'Paid', 'Remaining']
    )
    return balanceTable

def runSINOPAC(bankObject):
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if re.search(r'\bCOSLABAQU_\d+_\d+\b', file):
            os.remove(join(bankObject.downloadFolder, file))
    # Click Account Inquiry
    bankObject.wait.until(EC.presence_of_element_located((By.ID, 'MENU_CAO'))).click()
    # Click Loan inquiry
    bankObject.wait.until(EC.presence_of_element_located((By.ID, 'MENU_CAO002'))).click()
    # Click Loan Balance Inquiry
    bankObject.wait.until(EC.visibility_of_element_located((By.ID, 'MENU_COSLABAQU'))).click()
    # Swtich frame
    bankObject.driver.switch_to.frame('mainFrame')
    # Click Search
    xpath = '//*[contains(text(),"Search")]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    time.sleep(3)  # chờ load data
    # Click download file csv
    xpath = '//*[contains(@class,"download_csv")]'
    bankObject.driver.find_element(By.XPATH, xpath).click()
    # Đọc file download
    while True:
        checkFunc = lambda x: (re.search(r'\bCOSLABAQU_\d+_\d+\b', x) is not None) and ('download' not in x)
        downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
        if downloadFile:  # download xong -> có file
            break
        time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
    balanceTable = pd.read_csv(
        join(bankObject.downloadFolder, downloadFile),
        usecols=[1, 3, 5, 6, 7, 8],
        names=['ContractNumber', 'IssueDate', 'Currency', 'Amount', 'Remaining', 'ExpireDate'],
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
    balanceTable['IssueDate'] = pd.to_datetime(balanceTable['IssueDate'], format='%Y/%m/%d')
    balanceTable['ExpireDate'] = pd.to_datetime(balanceTable['ExpireDate'], format='%Y/%m/%d')
    # Date
    now = dt.datetime.now()
    if now.hour >= 12:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy cuối ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable['Date'] = d
    balanceTable['Bank'] = bankObject.bank
    # TermDays
    balanceTable['TermDays'] = (balanceTable['ExpireDate'] - balanceTable['IssueDate']).dt.days
    # Term months
    balanceTable['TermMonths'] = (balanceTable['TermDays'] / 30).round()
    # Paid
    balanceTable['Paid'] = balanceTable['Amount'] - balanceTable['Remaining']

    return balanceTable
