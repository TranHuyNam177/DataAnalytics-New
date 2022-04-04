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
    hide_window=True
) \
        -> pd.DataFrame:
    options = Options()
    if hide_window:
        options.headless = True

    url = 'https://www.hsx.vn/Modules/Listed/Web/TSCSReport?fid=0a6b7b2dbe234923b690924ed36283d8'
    driver = webdriver.Chrome(executable_path=PATH,options=options)
    wait = WebDriverWait(driver,5,ignored_exceptions=ignored_exceptions)
    driver.get(url)
    print('Getting Ligit Underlying Stocks')
    ticker_elems = wait.until(
        EC.presence_of_all_elements_located(
            (By.XPATH,'//*[@id!=""]/td[2]/a')
        )
    )
    tickers = list(map(lambda x:x.text,ticker_elems))
    driver.quit()

    return tickers
