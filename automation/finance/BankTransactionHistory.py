from automation.finance import *

# có CAPTCHA
def runBIDV(bankObject,fromDate,toDate,cBankAccountsOnly:bool):

    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    :param cBankAccountsOnly: Chỉ lấy các tài khoản giao dịch với KH
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
    time.sleep(0.5) # chờ animation
    # Click "Lịch sử giao dịch"
    xpath = f'//*[@data-action="btDetailTransaction"]'
    Button = bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath)))
    Button.click()
    # Chọn tab "Thời gian"
    bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Thời gian'))).click()
    # Lấy dữ liệu từng tài khoản
    frames = []
    for d in pd.date_range(fromDate,toDate):
        # Từ ngày
        fromDateInput = bankObject.wait.until(EC.presence_of_element_located((By.ID,'fromDate')))
        fromDateInput.clear()
        fromDateInput.send_keys(d.strftime('%d/%m/%Y'))
        # Đến ngày
        toDateInput = bankObject.wait.until(EC.presence_of_element_located((By.ID,'toDate')))
        toDateInput.clear()
        toDateInput.send_keys(d.strftime('%d/%m/%Y'))
        # Chọn từng tài khoản
        xpath = '//*[@aria-owns="accountNo_listbox"]'
        accountInput = bankObject.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        accountInput.clear()
        i = 0
        while i < accountNumber:
            # Bấm mũi tên xuống để lấy từng TK (làm cách này để tránh lỗi)
            accountInput.send_keys(Keys.DOWN)
            value = accountInput.get_attribute('value')
            print(value)
            account = re.search('[0-9]{14}',value).group()
            time.sleep(1) # chờ pop up (nếu có)
            # Đóng pop up nếu có
            xpath = '//*[@data-bb-handler="ok"]'
            popupButtons = bankObject.driver.find_elements(By.XPATH,xpath)
            if popupButtons:
                popupButtons[0].click()
            if cBankAccountsOnly and (account not in bankObject.cBankAccounts[bankObject.bank]):
                i += 1
                continue
            # Download file excel
            bankObject.wait.until(EC.presence_of_element_located((By.ID,'btnExportExcel01'))).click()
            # Đọc file, record data
            while True:
                checkFunc = lambda x: 'EBK_BC_LICHSUGIAODICH' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder),checkFunc)
                if downloadFile: # download xong -> có file
                    break
                time.sleep(1) # chưa download xong -> đợi thêm 1s nữa

            downloadData = pd.read_excel(
                join(bankObject.downloadFolder,downloadFile),
                usecols='B,D:G',
                names=['Time','Debit','Credit','Balance','Content'],
            )
            startRow = downloadData.eq('Dư đầu').any(axis=1).argmax() + 1
            endRow = downloadData.eq('Cộng phát sinh').any(axis=1).argmax() - 1
            frame = downloadData.iloc[startRow:endRow+1].copy() # pandas treats endpoint exclusively
            # Date
            frame['Time'] = pd.to_datetime(frame['Time'],format='%d/%m/%Y %H:%M:%S')
            # Debit, Credit, Balance
            for colName in ('Debit','Credit','Balance'):
                frame[colName] = frame[colName].str.replace(',','').astype(float)
            # Account Number
            frame.insert(1,'AccountNumber',account)
            frames.append(frame)
            # Xóa file
            os.remove(join(bankObject.downloadFolder,downloadFile))
            i += 1

    transactionTable = pd.concat(frames)
    transactionTable = transactionTable.loc[(transactionTable['Debit']!=0)|(transactionTable['Credit']!=0)]
    transactionTable.insert(1,'Bank',bankObject.bank)
    transactionTable = transactionTable.replace({np.nan: None}) # to be compatible with SQL

    return transactionTable
