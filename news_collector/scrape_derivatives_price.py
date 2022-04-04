from request.stock import *


PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException
)

def run(
    tickers:list,
    atdate:dt.datetime,
) -> pd.DataFrame:

    url_hose = 'https://finance.vietstock.vn/chung-khoan-phai-sinh/thong-ke-giao-dich.htm'
    driver = webdriver.Chrome(executable_path=PATH)
    wait = WebDriverWait(driver,60,ignored_exceptions=ignored_exceptions)
    driver.get(url_hose)

    # Điền Từ ngày Đến ngày
    atdateText = atdate.strftime('%d/%m/%Y')
    fromDateElem = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="txtFromDate"]/input')))
    fromDateElem.clear()
    fromDateElem.send_keys('01/01/2021')
    toDateElem = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="txtToDate"]/input')))
    toDateElem.clear()
    toDateElem.send_keys(atdateText)
    
    priceSeries = pd.Series(dtype=np.float64)
    for ticker in tickers:
        # Điền mã chứng khoán
        wait.until(EC.presence_of_element_located((By.CLASS_NAME,'selection'))).click()
        inputElem = wait.until(EC.presence_of_element_located((By.CLASS_NAME,'select2-search__field')))
        inputElem.send_keys(ticker)
        inputElem.send_keys(Keys.RETURN)
        # Xem
        possibleElems = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'m-b')))
        viewElem = [e for e in possibleElems if e.text=='Xem'][0]
        viewElem.click()
        # Lấy dòng đầu tiên
        rowElem = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="statistic-price"]/table/tbody/tr')))
        rowContent = rowElem.text.split()
        priceSeries[ticker] = float(rowContent[4].replace(',',''))

    return priceSeries