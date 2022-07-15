from automation.finance import *

def runIVB(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """
    now = dt.datetime.now()
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
        xpath = '//*[@id="fd_layer_pan"]/span'
        bankObject.wait.until(EC.visibility_of_element_located((By.XPATH,xpath))).click()
        xpath = '//*[@id="informAccount"]/tbody/tr[*]/td/a'
        contractElements = bankObject.driver.find_elements(By.XPATH,xpath)
        URLs = [e.get_attribute('href') for e in contractElements]
        if contractElements[0].text != '':
            break
    records = []
    for URL in URLs:
        bankObject.driver.get(URL)
        time.sleep(1)  # chờ để hiện số tài khoản (bắt buộc)
        # Danh sách các trường thông tin cần lấy:
        infoList = ['Số Hợp đồng','Số tiền vay','Dư nợ hiện tại','Loại tiền','Ngày giải ngân',
                    'Ngày đến hạn','Lãi suất']
        # Tạo list chứa tên cột tương đương với các cột trong Database
        infoColumnName = ['ContractNumber','Amount','Remaining','Currency','IssueDate','ExpireDate','InterestRate']
        finalDict = dict()
        for info, col_name in zip(infoList,infoColumnName):
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
        df['Amount'] = float(df['Amount'].str.replace(',',''))
        df['Remaining'] = float(df['Remaining'].str.replace(',', ''))
        df['Paid'] = df['Amount'] - df['Remaining']
        df['InterestRate'] = float(df['InterestRate'].str.split('%').str[0])/100
        df['IssueDate'] = pd.to_datetime(df['IssueDate'], format='%d/%m/%Y')
        df['ExpireDate'] = pd.to_datetime(df['ExpireDate'], format='%d/%m/%Y')
        # TermDays
        df['TermDays'] = (df['ExpireDate'] - df['IssueDate']).dt.days
        # Term months
        month = df['ExpireDate'].dt.month - df['IssueDate'].dt.month
        year = df['ExpireDate'].dt.year - df['IssueDate'].dt.year
        df['TermMonths'] = year * 12 + month
        df['InterestAmount'] = df['TermDays'] * (df['InterestRate'] / 360) * df['Amount']
        cols = ['Date','Bank','ContractNumber','TermDays','TermMonths','InterestRate','IssueDate','ExpireDate',
                'Amount','Paid','Remaining','InterestAmount','Currency']
        table = df[cols]
        records.append(table)
    balanceTable = pd.concat(records, ignore_index=True)
    return balanceTable
