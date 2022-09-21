import numpy as np
import pandas as pd
import os
from os import listdir
from os.path import join
import time
import re

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from function import first


# có CAPTCHA
def runBIDV(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    dayLimit = 100
    if (toDate - fromDate).days > dayLimit:
        raise ValueError(f'{bankObject.bank} không cho phép query quá {dayLimit} ngày một lần')

    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'EBK_BC_LICHSUGIAODICH' in file:
            os.remove(join(bankObject.downloadFolder, file))
    # Click Menu bar
    bankObject.wait.until(EC.presence_of_element_located((By.ID, 'menu-toggle-22'))).click()
    # Click "Tài khoản"
    bankObject.wait.until(EC.visibility_of_element_located((By.LINK_TEXT, 'Vấn tin'))).click()
    # Click "Tiền gửi thanh toán"
    bankObject.wait.until(EC.visibility_of_element_located((By.LINK_TEXT, 'Tiền gửi thanh toán'))).click()
    # Lấy số lượng tài khoản
    Accounts = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'change')))
    accountNumber = len(Accounts)
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
    # Chọn từng tài khoản
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
        time.sleep(1)  # tránh nhanh quá -> click mà ko download được file
        # Lấy số tài khoản
        value = accountInput.get_attribute('value')
        print(value)
        account = re.search('\d{14}', value).group()
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
            usecols='B,D:G',
            names=['Time', 'Debit', 'Credit', 'Balance', 'Content'],
        )
        # Account Number
        frame.insert(1, 'AccountNumber', account)
        frames.append(frame)
        # Trading Account
        if account in bankObject.cBankAccounts[bankObject.bank]:
            frame.insert(frame.shape[1], 'TradingAccount', True)
        else:
            frame.insert(frame.shape[1], 'TradingAccount', False)
        # Xóa file
        os.remove(join(bankObject.downloadFolder, downloadFile))
        i += 1

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    transactionTable = transactionTable.loc[~transactionTable['Time'].isna()]
    # Time
    transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], format='%d/%m/%Y %H:%M:%S')
    # Debit
    transactionTable['Debit'] = transactionTable['Debit'].fillna(0)
    transactionTable['Debit'] = transactionTable['Debit'].map(
        lambda x: float(x.replace(',', '')) if isinstance(x, str) else x)
    # Credit
    transactionTable['Credit'] = transactionTable['Credit'].fillna(0)
    transactionTable['Credit'] = transactionTable['Credit'].map(
        lambda x: float(x.replace(',', '')) if isinstance(x, str) else x)
    # Balance
    transactionTable['Balance'] = transactionTable['Balance'].map(
        lambda x: float(x.replace(',', '')) if isinstance(x, str) else x)
    transactionTable['Balance'] = transactionTable['Balance'].replace({np.nan: None})
    # Bank
    transactionTable.insert(1, 'Bank', bankObject.bank)

    return transactionTable


# không CAPTCHA
def runVTB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    dayLimit = 31
    if (toDate - fromDate).days > dayLimit:
        raise ValueError(f'{bankObject.bank} không cho phép query quá {dayLimit} ngày một lần')

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'lich-su-giao-dich' in file:
            os.remove(join(bankObject.downloadFolder, file))

    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Bắt đầu từ trang chủ
    xpath = '//*[@href="/"]'
    _, MainMenu = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    MainMenu.click()
    # Check tab "Tài khoản" có bung chưa (đã được click trước đó), phải bung rồi mới hiện tab "Danh sách tài khoản"
    xpath = '//*[text()="Danh sách tài khoản"]'
    queryElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    if not queryElement.is_displayed():  # nếu chưa bung
        # Click "Thông tin tài khoản"
        xpath = '//*[text()="Tài khoản"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
        time.sleep(2)  # chờ animation
    queryElement.click()
    time.sleep(1)
    # table Element
    # tableElement = bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'MuiTableBody-root')))
    xpath = '//*[contains(@id,"DDAAcount")]//table'
    tableElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    tableElement.find_element(By.LINK_TEXT, 'Xem thêm').click()

    # Create function to clear input box and send dates as string
    def sendDate(element, d):
        action = ActionChains(bankObject.driver)
        action.click(element)
        time.sleep(0.5)
        action.key_down(Keys.CONTROL, element)
        action.send_keys_to_element(element, 'a')
        action.key_up(Keys.CONTROL, element)
        action.send_keys_to_element(element, Keys.BACKSPACE)
        action.send_keys_to_element(element, d.strftime('%d/%m/%Y'))
        action.send_keys_to_element(element, Keys.ENTER)
        action.perform()

    frames = []
    accountNumbers = filter(lambda x: len(x) == 12, tableElement.text.split('\n'))
    for x in accountNumbers:
        bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT, x))).click()
        time.sleep(1)
        fromDateInput, toDateInput = bankObject.wait.until(
            EC.visibility_of_all_elements_located((By.CLASS_NAME, 'ant-picker-input')))
        # Điền ngày
        sendDate(fromDateInput, fromDate)
        sendDate(toDateInput, toDate)
        while True:
            try:
                bankObject.driver.find_element(By.CLASS_NAME, 'btn-submit').click()
                break
            except (Exception,):
                pass
        # Download file
        while True:
            try:
                xpath = '//*[@src="/public/img/icon/icon-download.svg"]'
                bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
                bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT, 'Xuất Excel'))).click()
                break
            except (Exception,):
                pass

        # Đọc file, record data
        while True:
            checkFunc = lambda x: 'lich-su-giao-dich' in x and 'download' not in x
            downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
            if downloadFile:  # download xong -> có file
                break
            time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa

        frame = pd.read_excel(
            join(bankObject.downloadFolder, downloadFile),
            skiprows=24,
            usecols='B:F',
            names=['Time', 'Content', 'Debit', 'Credit', 'Balance'],
        )
        # Account Number
        frame.insert(1, 'AccountNumber', x)
        # Trading Account
        if x in bankObject.cBankAccounts[bankObject.bank]:
            frame.insert(frame.shape[1], 'TradingAccount', True)
        else:
            frame.insert(frame.shape[1], 'TradingAccount', False)
        frames.append(frame)
        # Xóa file
        os.remove(join(bankObject.downloadFolder, downloadFile))

        bankObject.driver.back()
        bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT, 'Xem thêm'))).click()

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    # Bank
    transactionTable.insert(1, 'Bank', bankObject.bank)
    # Time
    transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], format='%d-%m-%Y %H:%M:%S')

    return transactionTable


# có CAPTCHA
def runIVB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'AccountTransacionHistory' in file:
            os.remove(join(bankObject.downloadFolder, file))

    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Bắt đầu từ trang chủ
    bankObject.driver.switch_to.default_content()
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@class="logoimg"]'))).click()
    # Click tab "Tài khoản"
    xpath = '//*[@data-menu-id="1"]'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    time.sleep(1)  # chờ animation
    # Click subtab "Sao kê tài khoản"
    bankObject.wait.until(EC.visibility_of_element_located((By.ID, '2_2'))).click()
    # Chọn "Tài khoản thanh toán" từ dropdown list
    bankObject.driver.switch_to.frame('mainframe')
    accountTypeInput = Select(bankObject.wait.until(EC.presence_of_element_located((By.ID, 'selectedAccType'))))
    accountTypeInput.select_by_visible_text('Tài khoản Thanh toán')

    # Điền số tài khoản:
    accountElems = bankObject.driver.find_elements(By.XPATH, '//*[@id="account_list"]/option')
    options = [a.text for a in accountElems]
    frames = []
    for option in options:
        time.sleep(1)
        account = option.split()[0]
        if account not in ('1017816-066', '1017816-069', '1017816-068'):
            continue
        accountInput = Select(bankObject.wait.until(EC.presence_of_element_located((By.ID, 'account_list'))))
        accountInput.select_by_visible_text(option)
        # Từ ngày
        fromDateInput = bankObject.driver.find_element(By.ID, 'beginDate')
        fromDateInput.clear()
        fromDateInput.send_keys(fromDate.strftime('%d/%m/%Y'))
        # Đến ngày
        toDateInput = bankObject.driver.find_element(By.ID, 'endDate')
        toDateInput.clear()
        toDateInput.send_keys(toDate.strftime('%d/%m/%Y'))
        # Click "Truy vấn"
        bankObject.driver.find_element(By.ID, 'btnQuery').click()
        # Click "In sao kê Excel"
        bankObject.driver.execute_script(f'window.scrollTo(0,100000)')
        xpath = """//*[@onclick="downloadReport('excel');"]"""
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
        # Đọc file download
        while True:
            checkFunc = lambda x: 'AccountTransacionHistory' in x and 'download' not in x
            downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
            if downloadFile:  # download xong -> có file
                break
            time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa
        frame = pd.read_excel(
            join(bankObject.downloadFolder, downloadFile),
            skiprows=7,
            skipfooter=1,
            usecols='C,E,F,H,J',
            names=['AccountNumber', 'Debit', 'Credit', 'Time', 'Content'],
        )
        # Trading Account
        if account in bankObject.cBankAccounts[bankObject.bank]:
            frame.insert(frame.shape[1], 'TradingAccount', True)
        else:
            frame.insert(frame.shape[1], 'TradingAccount', False)
        frames.append(frame)
        # Delete download file
        os.remove(join(bankObject.downloadFolder, downloadFile))

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    # Bank
    transactionTable.insert(0, 'Bank', bankObject.bank)
    # Balance
    transactionTable.insert(5, 'Balance', None)
    if not transactionTable.empty:
        # Time
        transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], '%Y-%m-%d %H:%M:%S')

    return transactionTable


# có CAPTCHA
def runVCB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    dayLimit = 90
    if (toDate - fromDate).days > dayLimit:
        raise ValueError(f'{bankObject.bank} không cho phép query quá {dayLimit} ngày một lần')

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'Vietcombank_Account_Statement' in file:
            os.remove(join(bankObject.downloadFolder, file))

    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Bắt đầu từ trang chủ
    _, homeButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'icon-home')))
    homeButton.click()
    # Lấy danh sách đường dẫn vào tài khoản thanh toán
    xpath = '//*[@id="dstkdd-tbody"]//td/a'
    accountElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    URLs = [e.get_attribute('href') for e in accountElems]

    # Create function to clear input box and send dates as string
    def sendDate(element, d):
        action = ActionChains(bankObject.driver)
        action.click(element)
        time.sleep(0.5)
        action.key_down(Keys.CONTROL, element)
        action.send_keys_to_element(element, 'a')
        action.key_up(Keys.CONTROL, element)
        action.send_keys_to_element(element, Keys.BACKSPACE)
        action.send_keys_to_element(element, d.strftime('%d/%m/%Y'))
        action.send_keys_to_element(element, Keys.ENTER)
        action.perform()

    frames = []
    for URL in URLs:
        bankObject.driver.get(URL)
        time.sleep(3)  # chờ để hiện số tài khoản (bắt buộc)
        # Điền ngày
        startDateInput = bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'startDate')))
        bankObject.driver.execute_script(f'window.scrollTo(0,500)')
        sendDate(startDateInput, fromDate)
        time.sleep(2)
        endDateInput = bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'endDate')))
        sendDate(endDateInput, toDate)
        # Xuống cuối trang
        bankObject.driver.execute_script(f'window.scrollTo(0,100000)')
        # Click "Xem sao kê"
        bankObject.wait.until(EC.visibility_of_element_located((By.ID, 'TransByDate'))).click()
        # Click "Xuất file"
        exportButton = bankObject.wait.until(
            EC.visibility_of_element_located((By.ID, 'ctl00_Content_TransactionDetail_ExportButton')))
        while True:  # do nút button bị cover bởi một layer loading
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
            downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
            if downloadFile:  # download xong -> có file
                renameFile = downloadFile.replace('xls', 'csv')
                os.rename(join(bankObject.downloadFolder, downloadFile), join(bankObject.downloadFolder, renameFile))
                downloadFile = renameFile
                break
            time.sleep(0.5)  # chưa download xong -> đợi thêm 1s nữa

        with open(join(bankObject.downloadFolder, downloadFile), 'rb') as f:
            htmlString = f.read()

        soup = BeautifulSoup(htmlString, 'html5lib')
        account = soup.find(text='Số tài khoản:').find_next().text
        frame = pd.read_html(htmlString, skiprows=11)[0]
        frame = frame.iloc[:frame.shape[0] - 18, [0, 2, 3, 4]]
        frame.columns = ['Time', 'Debit', 'Credit', 'Content']
        frame.insert(1, 'AccountNumber', account)
        if frame.empty:
            continue
        # Trading Account
        if account in bankObject.cBankAccounts[bankObject.bank]:
            frame.insert(frame.shape[1], 'TradingAccount', True)
        else:
            frame.insert(frame.shape[1], 'TradingAccount', False)
        frames.append(frame)
        # Delete download file
        os.remove(join(bankObject.downloadFolder, downloadFile))

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    # Bank
    transactionTable.insert(0, 'Bank', bankObject.bank)
    # Balance
    transactionTable.insert(5, 'Balance', None)
    # Xử lý Debit, Credit, Time
    transactionTable['Debit'] = transactionTable['Debit'].fillna(0)
    transactionTable['Credit'] = transactionTable['Credit'].fillna(0)
    if not transactionTable.empty:
        # Time
        transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], format='%Y-%m-%d')

    return transactionTable


# không CAPTCHA
def runEIB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    dayLimit = 90
    if (toDate - fromDate).days > dayLimit:
        raise ValueError(f'{bankObject.bank} không cho phép query quá {dayLimit} ngày một lần')

    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'LichSuTaiKhoan' in file:
            os.remove(join(bankObject.downloadFolder, file))
    # Click Trang chủ
    xpath = "//a[@href='/KHDN/home']"
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    # Click Menu Tài khoản --> Tiền gửi thanh toán
    action = ActionChains(bankObject.driver)
    xpath = '//*[@class="navigation-menu"]/li[2]/a'
    menuAccount = bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    action.move_to_element(menuAccount)
    xpath = '//*[@href="/KHDN/corp/account/payment"]'
    currentAccount = bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    action.click(currentAccount)
    action.perform()
    time.sleep(3)  # Chờ load xong
    # Đóng cửa sổ yêu cầu kích hoạt smart OTP (nếu có)
    xpath = '//*[contains(text(),"Đồng ý")]'
    Buttons = bankObject.driver.find_elements(By.XPATH, xpath)
    if Buttons:
        Buttons[0].click()
    xpath = '//*[contains(text(),"Đóng")]'
    Buttons = bankObject.driver.find_elements(By.XPATH, xpath)
    if Buttons:
        Buttons[0].click()
    # Lấy danh sách tài khoản
    xpath = '//tbody/tr/th/a'
    accountElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    accounts = [e.text for e in accountElems]
    # Lấy số dư tài khoản
    frames = []
    for account in accounts:
        Account = bankObject.wait.until(EC.visibility_of_element_located((By.LINK_TEXT, account)))
        Account.click()
        # Click Xem lịch sử tài khoản
        xpath = '//*[@class="modal-footer"]/button'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()  # nút đầu tiên
        xpath = '//*[@name="fromDate" or @name="toDate"]/div/input'
        fromDateInput, toDateInput = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        # Từ ngày
        fromDateInput.clear()
        fromDateInput.send_keys(f"{fromDate.strftime('%d/%m/%Y')} 00:00:00")
        # Đến ngày
        toDateInput.clear()
        toDateInput.send_keys(f"{toDate.strftime('%d/%m/%Y')} 23:59:59")
        # Click "Tìm kiếm"
        _, searchButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'btn-primary')))
        while True:
            try:
                searchButton.click()
                time.sleep(0.5)
                break
            except (ElementClickInterceptedException,):
                pass
        # Check xem có giao dịch không, vì không có giao dịch sẽ ko có file download
        flags = bankObject.driver.find_elements(By.XPATH, '//*[text()=" Không tìm thấy thông tin. "]')
        if flags:
            bankObject.driver.back()
            continue
        # Download file
        while True:
            try:
                bankObject.wait.until(EC.presence_of_element_located((By.LINK_TEXT, 'Tải Excel'))).click()
                break
            except (ElementClickInterceptedException, StaleElementReferenceException):
                pass
        # Đọc file download
        while True:
            checkFunc = lambda x: 'LichSuTaiKhoan' in x and 'download' not in x
            downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
            if downloadFile:  # download xong -> có file
                break
            time.sleep(1)  # chưa download xong -> đợi thêm 1s nữa

        frame = pd.read_excel(
            join(bankObject.downloadFolder, downloadFile),
            skiprows=6,
            skipfooter=4,
            usecols='B,D:G',
            names=['Time', 'Debit', 'Credit', 'Balance', 'Content'],
        )
        # Account Number
        frame.insert(1, 'AccountNumber', account)
        # Trading Account
        if account in bankObject.cBankAccounts[bankObject.bank]:
            frame.insert(frame.shape[1], 'TradingAccount', True)
        else:
            frame.insert(frame.shape[1], 'TradingAccount', False)
        frames.append(frame)
        # Delete download file
        os.remove(join(bankObject.downloadFolder, downloadFile))
        bankObject.driver.back()

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    transactionTable.insert(1, 'Bank', bankObject.bank)
    transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], format='%d/%m/%Y %H:%M:%S')

    return transactionTable


# không CAPTCHA
def runOCB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'TransactionHistory' in file:
            os.remove(join(bankObject.downloadFolder, file))

    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Bắt đầu từ main menu
    bankObject.wait.until(EC.presence_of_element_located((By.ID, 'main-menu-icon'))).click()
    # Click Thông tin tài khoản --> Sao kê tài khoản
    xpath = '//*[@class="ahref"]/*[@class="accounts-icon"]'
    _, accountButton = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    accountButton.click()
    xpath = '//*[@id="side-nav"]/ul/li[2]/div[3]/span'
    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    time.sleep(10)  # chờ chuyển trang
    # Lấy danh sách tài khoản
    i = 0
    frames = []
    while True:
        # Click dropdown button
        bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'indicator'))).click()  # dropdown button
        options = bankObject.wait.until(
            EC.visibility_of_all_elements_located((By.CLASS_NAME, 'rb-account-select__account-desc-row')))
        if i == len(options):
            break
        # Nhập số tài khoản
        Entry = options[i]
        i += 1
        accountNumber = Entry.text.split('\n')[0].replace(' ', '')
        Entry.click()
        # Click "Tìm kiếm nâng cao"
        xpath = '//*[@ng-click="toggleAdvancedSearch()"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
        time.sleep(1)
        # Tick chọn tính năng khoảng thời gian
        xpath = '//*[@class="bd-radio-option__marker"]'
        _, rangeButton = bankObject.wait.until(EC.visibility_of_all_elements_located((By.XPATH, xpath)))
        rangeButton.click()
        # Từ ngày
        fromDateInput = bankObject.wait.until(EC.presence_of_element_located((By.NAME, 'dateFromInput')))
        fromDateInput.clear()
        fromDateInput.send_keys(fromDate.strftime('%d.%m.%Y'))
        # Đến ngày
        toDateInput = bankObject.wait.until(EC.presence_of_element_located((By.NAME, 'dateToInput')))
        toDateInput.clear()
        toDateInput.send_keys(toDate.strftime('%d.%m.%Y'))
        # Lấy file
        while True:  # click đến khi được thì thôi
            try:
                searchButton = bankObject.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@bd-id="search_button_mobile"]')))
                searchButton.click()
                break
            except (ElementClickInterceptedException,):
                pass
        time.sleep(1)  # chờ load data
        # Check xem có data không
        xpath = '//*[contains(text(),"Không giao dịch nào phù hợp với tiêu chuẩn tìm kiếm")]'
        NoDataNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if not NoDataNotices:  # có data
            # Click Tải về chi tiết
            xpath = '//*[text()="Tải về"]'
            while True:
                try:
                    bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
                    break
                except (ElementNotInteractableException,):
                    time.sleep(1)  # click bị fail thì chờ 1s rồi click lại
            time.sleep(1)  # chờ animation
            xpath = '//*[text()="Tập tin XLS"]'
            bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
            # Đọc file download
            while True:
                checkFunc = lambda x: 'TransactionHistory' in x and 'download' not in x
                downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
                if downloadFile:  # download xong -> có file
                    break
                time.sleep(0.5)  # chưa download xong -> đợi thêm 0.5s nữa

            frame = pd.read_excel(
                join(bankObject.downloadFolder, downloadFile),
                skiprows=13,
                usecols='B,G:I',
                names=['Time', 'Content', 'Debit', 'Credit'],
            )
            # Account Number
            frame.insert(1, 'AccountNumber', accountNumber)
            # Trading Account
            if accountNumber in bankObject.cBankAccounts[bankObject.bank]:
                frame.insert(frame.shape[1], 'TradingAccount', True)
            else:
                frame.insert(frame.shape[1], 'TradingAccount', False)
            frames.append(frame)
            # Xóa file
            os.remove(join(bankObject.downloadFolder, downloadFile))

        # Scroll lên đầu trang
        bankObject.driver.execute_script(f'window.scrollTo(0,0)')
        time.sleep(0.5)  # tránh click khi chưa scroll xong

        # Click "Ẩn tìm kiếm nâng cao"
        xpath = '//*[@ng-click="toggleAdvancedSearch()"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    transactionTable.insert(1, 'Bank', bankObject.bank)
    transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], format='%d/%m/%Y')
    convertStrToInt = lambda x: abs(float(x.replace(',', ''))) if isinstance(x, str) else abs(float(x))
    transactionTable['Debit'] = transactionTable['Debit'].map(convertStrToInt)
    transactionTable['Credit'] = transactionTable['Credit'].map(convertStrToInt)
    transactionTable['Balance'] = None

    return transactionTable


# không CAPTCHA
def runTCB(bankObject, fromDate, toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    # Dọn dẹp folder trước khi download
    for file in listdir(bankObject.downloadFolder):
        if 'enquiry' in file:
            os.remove(join(bankObject.downloadFolder, file))

    time.sleep(3)  # nghỉ 3s giữa mỗi hàm để bankObject kịp update
    # Check tab "Tông tin tài khoản" có bung chưa (đã được click trước đó), phải bung rồi mới hiện tab "Truy vấn giao dịch tài khoản"
    xpath = '//*[contains(text(),"Truy vấn giao dịch tài khoản")]'
    queryElement = bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    if not queryElement.is_displayed():  # nếu chưa bung
        # Click "Thông tin tài khoản"
        xpath = '//*[contains(text(),"Thông tin tài khoản")]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
    # Click "Truy vấn giao dịch tài khoản"
    queryElement.click()
    time.sleep(3)
    # Click dropdown button
    bankObject.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'dropdown_button'))).click()
    # Lấy danh sách tài khoản
    xpath = '//*[contains(@id,"dropDownRow")]'
    accountElems = bankObject.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
    infoTexts = [e.text.split('-')[0] for e in accountElems]
    accounts = [info.split()[0] for info in infoTexts]
    frames = []
    for account in accounts:
        # Điền số tài khoản
        accountBox = bankObject.wait.until(EC.presence_of_element_located((By.ID, 'fieldName:FIELD.1')))
        accountBox.clear()
        accountBox.send_keys(account)
        # Từ ngày
        fromDateBox = bankObject.wait.until(EC.presence_of_element_located((By.ID, 'fieldName:DATE.FIELD.1')))
        fromDateBox.clear()
        fromDateBox.send_keys(fromDate.strftime('%d/%m/%Y'))
        # Đến ngày
        toDateBox = bankObject.wait.until(EC.presence_of_element_located((By.ID, 'fieldName:DATE.FIELD.2')))
        toDateBox.clear()
        toDateBox.send_keys(toDate.strftime('%d/%m/%Y'))
        # Click "Thực hiện"
        xpath = '//*[@title="Thực hiện"]'
        bankObject.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
        # Check xem có giao dịch nào không
        time.sleep(5)
        xpath = '//*[contains(text(),"Quý khách không có giao dịch nào")]'
        NoDataNotices = bankObject.driver.find_elements(By.XPATH, xpath)
        if NoDataNotices:
            # Back ra trang tài khoản
            bankObject.wait.until(EC.presence_of_element_located((By.ID, 'nbackbtn'))).click()
            continue
        # Tải file
        while True:
            Buttons = bankObject.driver.find_elements(By.ID, 'nsavebtn')
            if Buttons:
                Buttons[0].click()
                break
            time.sleep(1)
        # Đọc file download
        while True:
            checkFunc = lambda x: 'enquiry' in x and 'download' not in x
            downloadFile = first(listdir(bankObject.downloadFolder), checkFunc)
            if downloadFile:  # download xong -> có file
                break
            time.sleep(0.5)  # chưa download xong -> đợi thêm 0.5s nữa

        frame = pd.read_csv(
            join(bankObject.downloadFolder, downloadFile),
            skiprows=5,
            usecols=[0, 3, 4, 5, 6],
            names=['Time', 'Content', 'Debit', 'Credit', 'Balance'],
        )
        # Account Number
        frame.insert(1, 'AccountNumber', account)
        # Trading Account
        if account in bankObject.cBankAccounts[bankObject.bank]:
            frame.insert(frame.shape[1], 'TradingAccount', True)
        else:
            frame.insert(frame.shape[1], 'TradingAccount', False)
        frames.append(frame)
        # Xóa file
        os.remove(join(bankObject.downloadFolder, downloadFile))
        # Back ra trang tài khoản
        bankObject.wait.until(EC.presence_of_element_located((By.ID, 'nbackbtn'))).click()

    # Catch trường hợp không có data
    if not frames:
        return pd.DataFrame()

    transactionTable = pd.concat(frames)
    transactionTable.insert(1, 'Bank', bankObject.bank)
    transactionTable['Time'] = pd.to_datetime(transactionTable['Time'], format='%d/%m/%Y')

    def convertStrToInt(x):
        if isinstance(x, str):
            return abs(float(x.replace(',', '')))
        else:
            if x != x:  # np.nan
                return 0
            else:
                return abs(float(x))

    transactionTable['Debit'] = transactionTable['Debit'].map(convertStrToInt)
    transactionTable['Credit'] = transactionTable['Credit'].map(convertStrToInt)
    transactionTable['Balance'] = transactionTable['Balance'].map(convertStrToInt)

    return transactionTable
