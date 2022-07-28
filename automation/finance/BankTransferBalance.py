from automation.finance import *

# có CAPTCHA
def runBIDV(bankObject,fromDate,toDate):
    """
        :param bankObject: Bank Object (đã login)
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngày kết thúc lấy dữ liệu
    """

    dayLimit = 100
    if (toDate - fromDate).days > dayLimit:
        raise ValueError(f'{bankObject.bank} không cho phép query quá {dayLimit} ngày một lần')

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'EBK_BC_LICHSUGIAODICH' in file:
            os.remove(join(bankObject.downloadFolder, file))
    # Click Menu bar
    bankObject.wait.until(EC.presence_of_element_located((By.ID, 'menu-toggle-22'))).click()
    # Click "Vấn tin"
    bankObject.wait.until(EC.visibility_of_element_located((By.LINK_TEXT, 'Vấn tin'))).click()
    # Click "Tiền gửi thanh toán"
    bankObject.wait.until(EC.visibility_of_element_located((By.LINK_TEXT, 'Tiền gửi thanh toán'))).click()
    # Lấy số lượng tài khoản
    Accounts = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'change')))
    accountNumber = len(Accounts)
    # Danh sách tài khoản
    bankAccounts = ['26110002677688']
    # Click vào tài khoản đầu tiên
    xpath = f'//*[@class="change"]'
    Account = bankObject.wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
    Account.click()
    time.sleep(0.5)  # chờ animation
    # Click "Lịch sử giao dịch"
    xpath = f'//*[@data-action="btDetailTransaction"]'
    Button = bankObject.wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
    Button.click()
    # Chọn tab "Thời gian"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT, 'Thời gian'))).click()
    # Lấy dữ liệu từng tài khoản
    frames = []
    # Từ ngày
    fromDateInput = bankObject.wait.until(EC.presence_of_element_located((By.ID, 'fromDate')))
    fromDateInput.clear()
    fromDateInput.send_keys(fromDate.strftime('%d/%m/%Y'))
    # Đến ngày
    toDateInput = bankObject.wait.until(EC.presence_of_element_located((By.ID, 'toDate')))
    toDateInput.clear()
    toDateInput.send_keys(toDate.strftime('%d/%m/%Y'))
    # Chọn ô chứa tài khoản
    xpath = '//*[@aria-owns="accountNo_listbox"]'
    accountInput = bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    accountInput.clear()
    # Đóng pop up nếu có
    time.sleep(1)
    xpath = '//*[@data-bb-handler="ok"]'
    popupButtons = bankObject.driver.find_elements(By.XPATH, xpath)
    if popupButtons:
        popupButtons[0].click()
    i = 0
    while i < accountNumber:
        # Bấm mũi tên xuống để lấy từng TK (làm cách này để tránh lỗi)
        accountInput.send_keys(Keys.DOWN)
        # Đóng pop up nếu có
        time.sleep(1)  # chờ pop up (nếu có)
        xpath = '//*[@data-bb-handler="ok"]'
        popupButtons = bankObject.driver.find_elements(By.XPATH, xpath)
        if popupButtons:
            popupButtons[0].click()
        # Lấy số tài khoản
        value = accountInput.get_attribute('value')
        print(value)
        account = re.search('\d{14}', value).group()
        if account in bankAccounts:
            # Download file excel
            bankObject.wait.until(EC.presence_of_element_located((By.ID, 'btnExportExcel01'))).click()
            # Đọc file, record data
            while True:
                checkFunc = lambda x: 'EBK_BC_LICHSUGIAODICH' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
                if downloadFile:  # download xong -> có file
                    break
                time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa

            frame = pd.read_excel(
                join(bankObject.downloadFolder, downloadFile),
                skiprows=12,
                skipfooter=5,
                usecols='B,D,G',
                names=['Time', 'Debit', 'Content'],
            )
            # Account Number
            frame.insert(1, 'AccountNumber', account)
            frames.append(frame)
            # Xóa file
            os.remove(join(bankObject.downloadFolder, downloadFile))
            # Remove tài khoản đã duyệt
            bankAccounts.remove(account)
            if not bankAccounts:
                break
        i += 1
    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    transactionTable = transactionTable.loc[transactionTable['Content'].str.contains('chuyen tien sang TKCN')]
    transactionTable = transactionTable.reset_index(drop=True)
    # Time
    transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], format='%d/%m/%Y %H:%M:%S')
    # Debit
    transactionTable['Debit'] = transactionTable['Debit'].fillna(0)
    transactionTable['Debit'] = transactionTable['Debit'].map(lambda x: float(x.replace(',', '')) if isinstance(x, str) else x)
    # Bank
    transactionTable.insert(1, 'Bank', bankObject.bank)

    return transactionTable



