from request.stock import *

def run(
    hide_window=True
) -> pd.DataFrame:


    PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
    ignored_exceptions = (
        ValueError,
        IndexError,
        NoSuchElementException,
        StaleElementReferenceException,
        TimeoutException,
        ElementNotInteractableException
    )
    options = Options()
    if hide_window:
        options.headless = True

    # HOSE
    url_hose = 'https://iboard.ssi.com.vn/bang-gia/hose'
    driver = webdriver.Chrome(executable_path=PATH,options=options)
    wait = WebDriverWait(driver,60,ignored_exceptions=ignored_exceptions)
    driver.get(url_hose)
    print('Getting tickers in HOSE')
    ticker_elems_hose = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@id!=""]'))
    )
    tickers_hose = list(map(lambda x:x.get_attribute('id'),ticker_elems_hose))
    table_hose = pd.DataFrame(index=pd.Index(tickers_hose,name='ticker'))
    table_hose['exchange'] = 'HOSE'
    driver.quit()

    # HNX
    url_hnx = 'https://iboard.ssi.com.vn/bang-gia/hnx'
    driver = webdriver.Chrome(executable_path=PATH,options=options)
    wait = WebDriverWait(driver,60,ignored_exceptions=ignored_exceptions)
    driver.get(url_hnx)
    print('Getting tickers in HNX')
    ticker_elems_hnx = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@id!=""]'))
    )
    tickers_hnx = list(map(lambda x:x.get_attribute('id'),ticker_elems_hnx))
    table_hnx = pd.DataFrame(index=pd.Index(tickers_hnx,name='ticker'))
    table_hnx['exchange'] = 'HNX'
    driver.quit()

    # UPCOM
    url_upcom = 'https://iboard.ssi.com.vn/bang-gia/upcom'
    driver = webdriver.Chrome(executable_path=PATH,options=options)
    wait = WebDriverWait(driver,60,ignored_exceptions=ignored_exceptions)
    driver.get(url_upcom)
    print('Getting tickers in UPCOM')
    ticker_elems_upcom = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@id!=""]'))
    )
    tickers_upcom = list(map(lambda x:x.get_attribute('id'),ticker_elems_upcom))
    table_upcom = pd.DataFrame(index=pd.Index(tickers_upcom,name='ticker'))
    table_upcom['exchange'] = 'UPCOM'
    driver.quit()

    result = pd.concat([table_hose,table_hnx,table_upcom])

    return result
