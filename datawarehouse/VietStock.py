from datawarehouse import *


class ServerError(Exception):
    """Raise when the server does not respond"""
    pass

class MaximumRecords(Exception):
    """Raise when number of records exceeds 10,000 and prone to dropping data"""
    pass


class API(ServerError,MaximumRecords):

    """
    This class is given up due to unstability of Vietstock's Web Service
    """

    URL = 'http://s1.vietstock.vn/DataWS/DataWebService.svc?wsdl'
    client = Client(URL)

    def __init__(self,*args:object) -> None:
        super().__init__(*args)()

    def KetQuaGiaoDichCoPhieu(
        self,
        FromDate:dt.datetime,
        ToDate:dt.datetime,
    ):

        """
        Recommended diff days: 1 day
        """

        n = 0
        while n <= 5: # try upto 5 times
            try:
                # Call API from Web Service
                response = self.client.service.GetData(
                    _GroupType=4,
                    _DataType=1,
                    _FromDate=FromDate,
                    _ToDate=ToDate,
                )
                SerializedDictObject = zeep.helpers.serialize_object(response)
                SelectedNest = SerializedDictObject['_value_1'][1]['_value_1']
                break
            except (ValueError,KeyError,TypeError):
                time.sleep(30)
                n += 1
                continue
        else:
            raise ServerError(f'Retried {n-1} times and all failed')
            
        RowsNumber = len(SelectedNest)
        if RowsNumber == 10000:
            raise MaximumRecords

        tables = ((
                SelectedNest[row]['Table'],
                ) for row in range(RowsNumber)
            ) # generator
        table = pd.concat(tables)
        table['TrID'] = table['TrID'].astype(str)
        table['StockID'] = table['StockID'].astype(str)
        table['MarketState'] = table['MarketState'].astype(str)
        table['TradingDate'] = table['TradingDate'].map(lambda x: x.tz_localize(None))
        table['TradingTime'] = table['TradingTime'].map(lambda x: x.tz_localize(None))
        table['LastUpdate'] = table['LastUpdate'].map(lambda x: x.tz_localize(None))

        for col in table.columns:
            if col in ['TrID','StockID','TradingDate','TradingTime','LastUpdate','MarketState']:
                continue
            elif col in ['PerChange','TotalAdjustRate']:
                table[col] = table[col].astype(float)
            else:
                table[col] = table[col].astype(np.int64)

        # Insert to DWH-ThiTruong
        BATCHINSERT(connect_DWH_ThiTruong,'KetQuaGiaoDichCoPhieu',table)
        

class Web:

    """
    This class collect data using web scrapping
    """

    def __init__(
        self,
        groups:list
    ):

        """
        Constructor

        :param groups: list of one or many in ('VN30', 'HOSE', 'HNX30', 'HNX', 'UPCOM',
        'OilGas', 'BasicResource', 'IndustrialGoodsServices', 'FoodBeverage', 'HealthCare',
        'Media', 'Telecommunications', 'Banks', 'RealEstate', 'Technology', 'Chemicals',
        'ConstructionMaterials', 'AutomobilesParts', 'PersonalHouseholdGoods', 'Retail',
        'TravelLeisure', 'Utilities', 'Insurance', 'FinancialServices')
        """

        self.PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException
        )
        self.options = Options()
        self.options.headless = True

        getURL = {
            'VN30': 'https://iboard.ssi.com.vn/bang-gia/vn30',
            'HOSE': 'https://iboard.ssi.com.vn/bang-gia/hose',
            'HNX30': 'https://iboard.ssi.com.vn/bang-gia/hnx30',
            'HNX': 'https://iboard.ssi.com.vn/bang-gia/hnx',
            'UPCOM': 'https://iboard.ssi.com.vn/bang-gia/upcom',
            'OilGas': 'https://iboard.ssi.com.vn/bang-gia/industry-0500',
            'BasicResource': 'https://iboard.ssi.com.vn/bang-gia/industry-1700',
            'IndustrialGoodsServices': 'https://iboard.ssi.com.vn/bang-gia/industry-2700',
            'FoodBeverage': 'https://iboard.ssi.com.vn/bang-gia/industry-3500',
            'HealthCare': 'https://iboard.ssi.com.vn/bang-gia/industry-4500',
            'Media': 'https://iboard.ssi.com.vn/bang-gia/industry-5500',
            'Telecommunications': 'https://iboard.ssi.com.vn/bang-gia/industry-6500',
            'Banks': 'https://iboard.ssi.com.vn/bang-gia/industry-8300',
            'RealEstate': 'https://iboard.ssi.com.vn/bang-gia/industry-8600',
            'Technology': 'https://iboard.ssi.com.vn/bang-gia/industry-9500',
            'Chemicals': 'https://iboard.ssi.com.vn/bang-gia/industry-1300',
            'ConstructionMaterials': 'https://iboard.ssi.com.vn/bang-gia/industry-2300',
            'AutomobilesParts': 'https://iboard.ssi.com.vn/bang-gia/industry-3300',
            'PersonalHouseholdGoods': 'https://iboard.ssi.com.vn/bang-gia/industry-3700',
            'Retail': 'https://iboard.ssi.com.vn/bang-gia/industry-5300',
            'TravelLeisure': 'https://iboard.ssi.com.vn/bang-gia/industry-5700',
            'Utilities': 'https://iboard.ssi.com.vn/bang-gia/industry-7500',
            'Insurance': 'https://iboard.ssi.com.vn/bang-gia/industry-8500',
            'FinancialServices': 'https://iboard.ssi.com.vn/bang-gia/industry-8700',
        }

        driver = webdriver.Chrome(executable_path=self.PATH,options=self.options)
        wait = WebDriverWait(driver,30,ignored_exceptions=self.ignored_exceptions)

        self.tickers = []
        for group in groups:
            driver.get(getURL[group])
            elems = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,"stockSymbol")))
            self.tickers.extend([e.text for e in elems])

    
    def IntradayOrder(self):
        
        driver = webdriver.Chrome(executable_path=self.PATH)
        wait = WebDriverWait(driver,60,ignored_exceptions=self.ignored_exceptions)
        URL = 'https://finance.vietstock.vn/AAA/thong-ke-giao-dich.htm'
        driver.get(URL)
        driver.maximize_window()

        # Đăng nhập
        loginElem = wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Đăng nhập')))
        loginElem.click()
        emailElem = wait.until(EC.presence_of_element_located((By.ID,'txtEmailLogin')))
        emailElem.clear()
        emailElem.send_keys('hiepdang@phs.vn')
        passwordElem = wait.until(EC.presence_of_element_located((By.ID,'txtPassword')))
        passwordElem.clear()
        passwordElem.send_keys('123456')
        loginElem = wait.until(EC.presence_of_element_located((By.ID,'btnLoginAccount')))
        loginElem.click()

        d = {
            'Ticker': [],
            'Time': [],
            'Price': [],
            'Volume': [],
        }

        for ticker in self.tickers:
            # Tìm mã cổ phiếu
            tickerCell = wait.until(EC.presence_of_element_located((By.ID,'txt-filter-company')))
            tickerCell.clear()
            tickerCell.send_keys(ticker)
            tickerCell.send_keys(Keys.ARROW_DOWN)
            tickerCell.send_keys(Keys.RETURN)
            # Scroll down to data table
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight/4);")
            # Get number of pages
            pageElem = wait.until(EC.presence_of_element_located((By.CLASS_NAME,'m-r-xs')))
            numPages = int(pageElem.text.split('/')[1])

            page = 1
            while page <= numPages:
                # Time
                timeElems = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'col-xs-4')))
                d['Time'].extend(e.text for e in timeElems if ':' in e.text)
                # Price
                priceElems = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'p-r-xs')))
                d['Price'].extend(e.text for e in priceElems if len(e.text)>0 and e.text[0] in '0123456789')
                # Volume
                volumeElems = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'col-xs-3')))
                d['Volume'].extend(e.text for e in volumeElems if len(e.text)>0 and e.text[0] in '0123456789')
                # Next page
                wait.until(EC.presence_of_element_located((By.CLASS_NAME,'btn-page-next'))).click()
                page += 1