import time

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

    URL = 'https://priceboard.vcbs.com.vn/Priceboard/'
    driver = webdriver.Chrome(executable_path=PATH,options=options)
    wait = WebDriverWait(driver,60,ignored_exceptions=ignored_exceptions)
    driver.get(URL)

    # HOSE
    print('Getting tickers in HOSE')
    action = ActionChains(driver)
    action.move_to_element(driver.find_element(By.XPATH,'//*[text()="HOSE"]'))
    action.click(driver.find_element(By.XPATH,'//*[text()="Bảng giá HOSE"]'))
    action.perform()
    time.sleep(3)
    ticker_elems_hose = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@name!=""]'))
    )[1:]
    tickers_hose = list(map(lambda x:x.get_attribute('name'),ticker_elems_hose))
    table_hose = pd.DataFrame(index=pd.Index(tickers_hose,name='ticker'))
    table_hose['exchange'] = 'HOSE'

    # HNX
    print('Getting tickers in HNX')
    action = ActionChains(driver)
    action.move_to_element(driver.find_element(By.XPATH,'//*[text()="HNX"]'))
    action.click(driver.find_element(By.XPATH,'//*[text()="Bảng giá HNX"]'))
    action.perform()
    time.sleep(3)
    ticker_elems_hnx = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@name!=""]'))
    )[1:]
    tickers_hnx = list(map(lambda x:x.get_attribute('name'),ticker_elems_hnx))
    table_hnx = pd.DataFrame(index=pd.Index(tickers_hnx,name='ticker'))
    table_hnx['exchange'] = 'HNX'

    # UPCOM
    print('Getting tickers in UPCOM')
    action = ActionChains(driver)
    action.move_to_element(driver.find_element(By.XPATH,'//*[text()="UPCOM"]'))
    action.click(driver.find_element(By.XPATH,'//*[text()="Bảng giá UPCOM"]'))
    action.perform()
    time.sleep(3)
    ticker_elems_upcom = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@name!=""]'))
    )[1:]
    tickers_upcom = list(map(lambda x:x.get_attribute('name'),ticker_elems_upcom))
    table_upcom = pd.DataFrame(index=pd.Index(tickers_upcom,name='ticker'))
    table_upcom['exchange'] = 'UPCOM'

    # CW
    print('Getting tickers in CW')
    wait.until(EC.presence_of_element_located((By.XPATH,'//*[text()="Chứng quyền"]'))).click()
    time.sleep(3)
    ticker_elems_cw = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@name!=""]'))
    )[1:]
    tickers_cw = list(map(lambda x:x.get_attribute('name'),ticker_elems_cw))
    table_cw = pd.DataFrame(index=pd.Index(tickers_cw,name='ticker'))
    table_cw['exchange'] = 'CW'

    # BOND
    print('Getting tickers in BOND')
    wait.until(EC.presence_of_element_located((By.XPATH,'//*[text()="Trái phiếu doanh nghiệp"]'))).click()
    time.sleep(3)
    ticker_elems_bond = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@name!=""]'))
    )[1:]
    tickers_bond = list(map(lambda x:x.get_attribute('name'),ticker_elems_bond))
    table_bond = pd.DataFrame(index=pd.Index(tickers_bond,name='ticker'))
    table_bond['exchange'] = 'BOND'

    driver.quit()
    result = pd.concat([table_hose,table_hnx,table_upcom,table_cw,table_bond])

    return result
