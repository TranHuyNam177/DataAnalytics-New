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
    """
    This function returns a DataFrame and export an excel file containing date/price of issuancepublished in
    'https://finance.vietstock.vn/chung-khoan-phai-sinh/chung-quyen.htm'

    """
    start_time = time.time()

    options = Options()
    if hide_window:
        options.headless = True
    url = 'https://finance.vietstock.vn/chung-khoan-phai-sinh/chung-quyen.htm'
    driver = webdriver.Chrome(executable_path=PATH,options=options)
    driver.get(url)
    driver.execute_script("window.scrollTo(0,window.scrollY+1500)")
    wait = WebDriverWait(driver,1,ignored_exceptions=ignored_exceptions)

    cw_names = []
    cw_dates_of_issuance = []
    cw_prices_of_issuance = []
    status_list = []
    first_trading_dates = []
    while True:
        status = ''
        for row in np.arange(10)+1:
            print(row)
            while True:
                try:
                    print(f'Getting Status of row {row}')
                    status = driver.find_element_by_xpath(f'//*[@id="cw-list"]/table/tbody/tr[{row}]/td[13]').text
                    break
                except (Exception,):
                    continue
            if status!='Bình thường':
                break
            else:
                status_list.append(status)

                def f():
                    cw_element = wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH,f'//*[@id="cw-list"]/table/tbody/tr[{row}]/td[2]/a')
                        )
                    )
                    return cw_element

                cw_name = ''
                cw_url = ''
                d = ''
                while True:
                    try:
                        print(f'Getting CW Element of row {row}')
                        d = driver.find_element_by_xpath(f'//*[@id="cw-list"]/table/tbody/tr[{row}]/td[11]').text
                        cw_element = f()
                        cw_name = cw_element.text
                        cw_url = cw_element.get_attribute('href')
                        break
                    except (Exception,):
                        continue
                cw_names.append(cw_name)
                first_trading_dates.append(dt.datetime.strptime(d,'%d/%m/%Y'))
                sub_driver = webdriver.Chrome(executable_path=PATH,options=options)
                sub_driver.get(cw_url)
                # Ngay phat hanh
                cw_date_of_issuance = ''
                while True:
                    try:
                        print(f'Getting Issuane Date of row {row}')
                        cw_date_of_issuance = sub_driver.find_element_by_xpath(
                            "//*[@id='view-content']/div[2]/div[2]/div[2]/table/tbody/tr[8]/td[2]"
                        ).text
                        break
                    except (Exception,):
                        continue
                cw_dates_of_issuance.append(cw_date_of_issuance)
                value_names = []
                values = []
                # Tim gia phat hanh
                while True:
                    try:
                        print(f'Getting Issuance Price of row {row}')
                        value_name_elements = sub_driver.find_elements_by_xpath(
                            '//*[@id="view-content"]/div[2]/div[2]/div[2]/table/tbody/*/td[1]'
                        )
                        value_names = [elem.text for elem in value_name_elements]
                        value_elements = sub_driver.find_elements_by_xpath(
                            '//*[@id="view-content"]/div[2]/div[2]/div[2]/table/tbody/*/td[2]'
                        )
                        values = [elem.text for elem in value_elements]
                        break
                    except (Exception,):
                        continue
                idx = value_names.index('Giá phát hành:')
                cw_price_of_issuance = int(values[idx].replace(',',''))
                cw_prices_of_issuance.append(cw_price_of_issuance)
                sub_driver.quit()

        else:
            # Turn Page and continue if the inner loop wasn't broken.
            while True:
                try:
                    print(f'Turning page...')
                    wait.until(EC.element_to_be_clickable(
                        (By.XPATH,'//*[@aria-label="next"]'))
                    ).click()
                    break
                except (Exception,):
                    continue
            continue
        # Inner loop was broken, break the outer.
        break

    output_table = pd.DataFrame(
        {
            'ISSUANCE_DATE':cw_dates_of_issuance,
            'ISSUANCE_PRICE':cw_prices_of_issuance,
            'FIRST_TRADING_DATE':first_trading_dates,
            'STATUS':status_list,
        },
        index=pd.Index(cw_names,name='CW')
    )
    output_table.drop_duplicates(inplace=True)
    output_table.to_excel("date_price_of_issuance.xlsx")

    print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

    return output_table
