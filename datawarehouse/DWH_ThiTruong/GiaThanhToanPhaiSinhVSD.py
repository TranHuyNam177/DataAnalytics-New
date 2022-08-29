import time
from os.path import join, dirname
import pandas as pd
from request import *
from datawarehouse import DELETE, BATCHINSERT, connect_DWH_ThiTruong

def run(
    fromDate: dt.datetime,
    toDate: dt.datetime,
):

    fromDate = fromDate.replace(hour=0,minute=0,second=0,microsecond=0)
    toDate = toDate.replace(hour=0,minute=0,second=0,microsecond=0)
    PATH = join(dirname(dirname(dirname(__file__))),'dependency','chromedriver')
    ignored_exceptions = (
        ValueError,
        IndexError,
        NoSuchElementException,
        StaleElementReferenceException,
        TimeoutException,
        ElementNotInteractableException
    )
    options = Options()
    options.headless = False
    driver = webdriver.Chrome(options=options,executable_path=PATH)
    driver.maximize_window()
    wait = WebDriverWait(driver,10,ignored_exceptions=ignored_exceptions)

    URL = r'https://www.vsd.vn/vi/thong-tin-san-pham'
    driver.get(URL)

    # Click tab "Giá thanh toán cuối ngày (DSP), giá thanh toán cuối cùng (FSP)"
    xpath = "//a[contains(text(),'Giá thanh toán cuối ngày')]"
    tabElement = wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    tabElement.click()
    # Điền ngày
    xpath = "//input[contains(@id,'txtSearchPriceDate')]"
    dateBox = wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
    xpath = "//button[contains(@class,'btn-sl-search')]"
    searchButton = wait.until(EC.presence_of_element_located((By.XPATH,xpath)))

    records = []
    for d in pd.date_range(fromDate,toDate):
        # Clear ô ngày
        dateBox.clear()
        # Ngày ngày
        dateString = d.strftime('%d%m%Y')
        dateBox.click()
        for _ in range(10):
            dateBox.send_keys(Keys.BACKSPACE)
        dateBox.send_keys(dateString)
        # Bấm search
        searchButton.click()
        time.sleep(2) # chờ load data
        # Lấy record
        xpath = "//*[@id='divTab_Fill_GiaThanhToan']//tbody/tr"
        rowElements = driver.find_elements(By.XPATH,xpath)
        if not rowElements: # ngày này không có data
            continue
        rowStrings = [element.text for element in rowElements]
        for rowString in rowStrings:
            # Mã sản phẩm
            ticker = re.search(r'\b\w+\d{4}\b',rowString).group()
            # Ngày thanh toán
            paymentDateString = re.search(r'\b\d{2}/\d{2}/\d{4}\b',rowString).group()
            paymentDate = dt.datetime.strptime(paymentDateString,'%d/%m/%Y')
            # Giá thanh toán ngày (DSP)
            priceString = re.search(r'\d+\.\d{3},?\d*\b',rowString).group()
            price = float(priceString.replace('.','').replace(',','.'))
            # Append
            records.append((d,ticker,paymentDate,price))

    driver.quit()
    if not records: # không có record
        print('No data to insert')
        return

    table = pd.DataFrame(records,columns=['Ngay','MaSanPham','NgayThanhToan','GiaThanhToanNgayDSP'])
    # Insert vào Database
    fromDateString = fromDate.strftime('%Y-%m-%d')
    toDateString = toDate.strftime('%Y-%m-%d')
    whereClause = f"WHERE [Ngay] BETWEEN '{fromDateString}' AND '{toDateString}'"
    DELETE(connect_DWH_ThiTruong,"GiaThanhToanPhaiSinhVSD",whereClause)
    BATCHINSERT(connect_DWH_ThiTruong,'GiaThanhToanPhaiSinhVSD',table)

