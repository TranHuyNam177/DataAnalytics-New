import requests

from request.stock import *
from request import *


def PotentialRequestRefused(hours):

    """
    Warn when hours is too large
    """

    message = f"hours = {hours} is too large, choose <hours> <= 48 to reduce the risk of being rejected by the web server"
    warnings.warn(message,RuntimeWarning)


class __Base__:

    def __init__(
        self,
        hours:int,
    ):

        # Lấy danh sách mã chứng khoán gần nhất
        self.tickers = pd.read_sql(
            f"""
            SELECT [Ticker]
            FROM [DanhSachMa]
            WHERE [Date] = (SELECT MAX(Date) FROM [DanhSachMa])
            """,
            connect_DWH_ThiTruong,
        ).squeeze().to_list()

        self.PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException,
        )

        if hours > 120:
            PotentialRequestRefused(hours)

        self.hours = hours


    @staticmethod
    def __processTime__(x):
        now = dt.datetime.now()
        bmkTime = dt.datetime(now.year,now.month,now.day,now.hour,now.minute)
        if ' phút trước' in x:
            minutes = int(x.split()[0])
            result = bmkTime - dt.timedelta(minutes=minutes)
        elif ' giờ trước' in x:
            hours = int(x.split()[0])
            result = bmkTime - dt.timedelta(hours=hours)
        elif x == '':
            result = np.nan
        elif x == 'Vừa xong':
            result = now
        else:
            result = dt.datetime.strptime(x,'%d/%m/%Y %H:%M')
        return result


class cafef(__Base__):

    def __init__(
        self,
        hours:int,
    ):

        """
        :param hours: Số giờ quá khứ cần quét
        """

        __Base__.__init__(self,hours)
        self.urls = [
            f'https://cafef.vn/timeline/31/trang-1.chn',
            f'https://cafef.vn/timeline/36/trang-1.chn',
            f'https://cafef.vn/timeline/35/trang-1.chn',
            f'https://cafef.vn/timeline/34/trang-1.chn'
        ]

    def run(self):

        start = time.time()
        run_time = dt.datetime.now()
        bmk_time = run_time - dt.timedelta(hours=self.hours)

        saved_timestamps = []
        saved_tickers = []
        saved_titles = []
        saved_descriptions = []
        saved_bodies = []
        saved_urls = []

        print('Getting News from https://www.cafef.vn/')

        with requests.Session() as session:
            retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
            session.mount('https://',adapter)
            headers = {
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
            }
            for u in self.urls:
                print(u)
                i = 1
                articleTimestamp = run_time
                while articleTimestamp > bmk_time:
                    print(articleTimestamp)
                    url = u.replace('trang-1',f'trang-{i}')
                    html = session.get(url,headers=headers,timeout=30).text
                    bs = BeautifulSoup(html,'html5lib')
                    h3_tags = bs.find_all(name='h3')
                    URLs = ['https://www.cafef.vn' + tag.find(name='a').get('href') for tag in h3_tags]
                    for articleURL in URLs:
                        try:
                            print(articleURL)
                            session = requests.Session()
                            retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
                            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
                            session.mount('https://',adapter)
                            print(articleURL)
                            articleHTML = session.get(articleURL,headers=headers,timeout=30).text
                            articleBS = BeautifulSoup(articleHTML,'html5lib')
                            # Title
                            articleTitle = articleBS.find(class_='title').get_text(strip=True)
                            print(articleTitle)
                            # Timestamp
                            articleTimestamp = dt.datetime.strptime(articleBS.find(class_='pdate').get_text(strip=True)[:-3],'%d-%m-%Y - %H:%M')
                            print(articleTimestamp)
                            # Description
                            articleDescription = articleBS.find(class_='sapo').get_text(strip=True)
                            print(articleDescription)
                            # Body
                            articleBody = '\n'.join(tag.get_text(strip=True) for tag in articleBS.find(id='mainContent').find_all(name='p'))
                            print(articleBody)
                            # Tickers
                            pattern = r'\b[A-Z]{1}[A-Z0-9]{2}\b'
                            articleTickers = set(re.findall(pattern,f'{articleTitle}\n{articleDescription}\n{articleBody}'))
                            print(articleTickers)

                            saved_timestamps.append(articleTimestamp)
                            saved_tickers.append(','.join(articleTickers))
                            saved_titles.append(articleTitle)
                            saved_descriptions.append(articleDescription)
                            saved_bodies.append(articleBody)
                            saved_urls.append(articleURL)

                        except AttributeError:
                            print(f'{articleURL} đã bị gỡ hoặc là một Landing Page')

                    i += 1

        result_dict = {
            'Time': saved_timestamps,
            'Ticker': saved_tickers,
            'Title': saved_titles,
            'Description': saved_descriptions,
            'Body': saved_bodies,
            'URL': saved_urls,
        }
        result = pd.DataFrame(result_dict)

        if __name__=='__main__':
            print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
        else:
            print(f"{__name__.split('.')[-1]} ::: Finished")
        print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

        return result


class ndh(__Base__):

    def __init__(
        self,
        hours:int,
    ):

        __Base__.__init__(self,hours)
        self.urls = [
            f'https://ndh.vn/chung-khoan',
            f'https://ndh.vn/doanh-nghiep',
            f'https://ndh.vn/tai-chinh',
            f'https://ndh.vn/bat-dong-san',
        ]

    def run(self):

        start = time.time()
        print('Getting News from https://www.ndh.vn/')

        """
        Với một giá trị <hours> cho trước, vì không thể xác định cần phải lấy bao nhiêu tin (scroll bao nhiêu lần) 
        để đủ tin trong số <hours> giờ vừa qua, nên cần make assumption:
        :: Với <hours> đủ lớn (lớn hơn 12) thì số giờ > số tin. Vì thế, nếu cần lấy 12 tiếng trước thì chỉ cần lấy 12 tin
        là chắc chắn đủ (hoặc dư). Khi insert vào database, cần cơ chế DELETE các tin dư đã lấy dựa theo đường dẫn URL
        --> Cần lấy tối thiểu 12 giờ --> <hours> = MAX(<hours>,12)
        """

        # Scroll bằng Selenium
        options = Options()
        options.headless = False
        driver = webdriver.Chrome(options=options,executable_path=self.PATH)

        URLs = []
        for u in self.urls:
            driver.get(u)
            while True:
                articleElemsOfTab = driver.find_elements(By.XPATH,'//*[@class="title-news"]/a')
                driver.execute_script(f"window.scrollTo(0,1000000000000);") # nhập 1 số đủ lớn để scroll tới cuối màn hình
                buttonList = driver.find_elements(By.XPATH,'//*[@id="btnLoadmore"]/a')
                if len(articleElemsOfTab) > self.hours: # break khi đủ số lượng tin
                    break
                if len(buttonList) > 0: # xuất hiện nút "Bấm để xem thêm"
                    time.sleep(1) # chờ scroll xong để tránh click nhầm vào tiêu đề tin
                    try:
                        showMoreElem = buttonList[0]
                        actions = ActionChains(driver)
                        actions.move_to_element(showMoreElem).click().perform()
                    except StaleElementReferenceException: # chạm cuối trang, nút "Bấm để xem thêm biến mất --> dừng"
                        break
            URLs.extend(e.get_attribute('href') for e in articleElemsOfTab)

        driver.quit()

        saved_timestamps = []
        saved_tickers = []
        saved_titles = []
        saved_descriptions = []
        saved_bodies = []
        saved_urls = []

        # Làm việc trên HTML bằng BeautifulSoup
        with requests.Session() as session:
            retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
            session.mount('https://',adapter)
            headers = {
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
            }
            for URL in URLs:
                try:
                    articleHTML = session.get(URL,headers=headers,timeout=30).text
                    articleBS = BeautifulSoup(articleHTML,'html5lib')
                    # Title
                    articleTitle = articleBS.find(class_='title-detail').get_text(strip=True)
                    print(articleTitle)
                    # Timmestamp
                    rawTimestamp = articleBS.find(class_='date').get_text(strip=True)
                    splitTimestamp = rawTimestamp.replace(' (GMT+7)','').split(', ')
                    _, dateString, timeString = tuple(splitTimestamp)
                    splitdateString = dateString.split('/') ; splitdateString = [int(x) for x in splitdateString]
                    splittimeString = timeString.split(':') ; splittimeString = [int(x) for x in splittimeString]
                    articleTimestamp = dt.datetime(splitdateString[2],splitdateString[1],splitdateString[0],splittimeString[0],splittimeString[1])
                    print(articleTimestamp)
                    # Description
                    articleDescriptionTags = articleBS.find(class_='related-news').find_all(name='a')
                    articleDescription = '\n'.join(tag.get_text(strip=True) for tag in articleDescriptionTags)
                    print(articleDescription)
                    # Body
                    articleBody = '\n'.join(paragraph.get_text(strip=True) for paragraph in articleBS.find(class_='fck_detail').find_all(name='p'))
                    print(articleDescription)
                    # Tickers
                    pattern = r'\b[A-Z]{1}[A-Z0-9]{2}\b'
                    articleTickers = set(re.findall(pattern,f'{articleTitle}\n{articleDescription}\n{articleBody}'))
                    print(articleTickers)

                    saved_timestamps.append(articleTimestamp)
                    saved_tickers.append(','.join(articleTickers))
                    saved_titles.append(articleTitle)
                    saved_descriptions.append(articleDescription)
                    saved_bodies.append(articleBody)
                    saved_urls.append(URL)

                except AttributeError:
                    print(f'{URL} không phải là bài báo') # loại topic

        result_dict = {
            'Time': saved_timestamps,
            'Ticker': saved_tickers,
            'Title': saved_titles,
            'Description': saved_descriptions,
            'Body': saved_bodies,
            'URL': saved_urls,
        }
        result = pd.DataFrame(result_dict)

        if __name__=='__main__':
            print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
        else:
            print(f"{__name__.split('.')[-1]} ::: Finished")
        print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

        return result


class vietstock(__Base__):

    def __init__(
        self,
        hours: int,
    ):
        __Base__.__init__(self,hours)
        self.urls = [
            'https://vietstock.vn/chung-khoan.htm',
            'https://vietstock.vn/doanh-nghiep.htm',
            'https://vietstock.vn/bat-dong-san.htm',
            'https://vietstock.vn/tai-chinh.htm',
            'https://vietstock.vn/kinh-te.htm'
        ]

    def run(self):

        start = time.time()
        print('Getting News from https://www.vietstock.vn/')

        """
        Với một giá trị <hours> cho trước, vì không thể xác định cần phải lấy bao nhiêu tin (scroll bao nhiêu lần) 
        để đủ tin trong số <hours> giờ vừa qua, nên cần make assumption:
        :: Với <hours> đủ lớn (lớn hơn 12) thì số giờ > số tin. Vì thế, nếu cần lấy 12 tiếng trước thì chỉ cần lấy 12 tin
        là chắc chắn đủ (hoặc dư). Khi insert vào database, cần cơ chế DELETE các tin dư đã lấy dựa theo đường dẫn URL
        --> Cần lấy tối thiểu 12 giờ --> <hours> = MAX(<hours>,12)
        """

        # Qua trang bằng Selenium
        options = Options()
        options.headless = False
        driver = webdriver.Chrome(options=options,executable_path=self.PATH)
        driver.maximize_window()
        wait = WebDriverWait(driver,20,ignored_exceptions=self.ignored_exceptions)

        URLs = []
        for u in self.urls:
            driver.get(u)
            pageURLs = []
            while True:
                pageElems = wait.until(EC.presence_of_all_elements_located((By.XPATH,"//*[@class='thumb']/a")))[:10]
                pageURLs.extend(elem.get_attribute('href') for elem in pageElems)
                if len(pageURLs) > self.hours: # break khi đủ tin
                    break
                # Scroll & Next page
                nextPageElem = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="page-next "]')))
                actions = ActionChains(driver)
                actions.move_to_element(nextPageElem).click().perform()
                time.sleep(1)
            URLs.extend(pageURLs)

        driver.quit()

        # Làm việc với HTMl bằng Beautiful Soup
        saved_timestamps = []
        saved_tickers = []
        saved_titles = []
        saved_descriptions = []
        saved_bodies = []

        with requests.Session() as session:
            retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
            session.mount('https://',adapter)
            headers = {
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
            }
            for url in URLs:
                articleHTML = session.get(url,headers=headers,timeout=30).text
                articleBS = BeautifulSoup(articleHTML,'html5lib')
                # Title
                articleTitle = articleBS.find(class_='pTitle').get_text(strip=True)
                print(articleTitle)
                # Time
                articleTime = self.__processTime__(articleBS.find(class_='date').get_text(strip=True))
                print(articleTime)
                # Description
                articleDescription = articleBS.find(class_='pHead',name='p').get_text(strip=True)
                print(articleDescription)
                # Body
                articleBody = '\n'.join(tag.get_text() for tag in articleBS.find_all(class_='pBody'))
                print(articleBody)
                # Tickers
                pattern = r'\b[A-Z]{1}[A-Z0-9]{2}\b'
                articleTickers = set(re.findall(pattern,f'{articleTitle}\n{articleDescription}\n{articleBody}'))
                print(articleTickers)

                saved_timestamps.append(articleTime)
                saved_tickers.append(','.join(articleTickers))
                saved_titles.append(articleTitle)
                saved_descriptions.append(articleDescription)
                saved_bodies.append(articleBody)

        saved_urls = URLs

        result_dict = {
            'Time':saved_timestamps,
            'Ticker':saved_tickers,
            'Title':saved_titles,
            'Description':saved_descriptions,
            'Body':saved_bodies,
            'URL':saved_urls,
        }
        result = pd.DataFrame(result_dict)
        result['Time'] = result['Time'].fillna(method='ffill') # lấy giờ của tin trước

        if __name__=='__main__':
            print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
        else:
            print(f"{__name__.split('.')[-1]} ::: Finished")
        print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

        return result


class tinnhanhchungkhoan(__Base__):

    def __init__(
        self,
        hours: int,
    ):
        __Base__.__init__(self,hours)
        self.urls = [
            'https://www.tinnhanhchungkhoan.vn/chung-khoan/',
            'https://www.tinnhanhchungkhoan.vn/doanh-nghiep/',
        ]


    def run(self):

        start = time.time()
        print('Getting News from https://www.tinnhanhchungkhoan.vn/')

        # Click "Xem thêm" bằng Selenium
        options = Options()
        options.headless = False
        driver = webdriver.Chrome(options=options,executable_path=self.PATH)
        driver.maximize_window()
        wait = WebDriverWait(driver,10,ignored_exceptions=self.ignored_exceptions)

        """
        Với một giá trị <hours> cho trước, vì không thể xác định cần phải lấy bao nhiêu tin (scroll bao nhiêu lần) 
        để đủ tin trong số <hours> giờ vừa qua, nên cần make assumption:
        :: Với <hours> đủ lớn (lớn hơn 12) thì số giờ > số tin. Vì thế, nếu cần lấy 12 tiếng trước thì chỉ cần lấy 12 tin
        là chắc chắn đủ (hoặc dư). Khi insert vào database, cần cơ chế DELETE các tin dư đã lấy dựa theo đường dẫn URL
        --> Cần lấy tối thiểu 12 giờ --> <hours> = MAX(<hours>,12)
        """

        URLs = []
        for u in self.urls:
            driver.get(u)
            pageURLs = []
            while True:
                pageElems = wait.until(EC.presence_of_all_elements_located((By.XPATH,'//*[@class="story "]/figure/a')))
                pageURLs.extend(e.get_attribute('href') for e in pageElems)
                if len(pageURLs) > self.hours: # break khi đủ số lượng tin
                    break
                # Click "Xem thêm"
                showMoreElem = wait.until(EC.presence_of_element_located((By.ID,'viewmore'))) # Click "Xem thêm"
                actions = ActionChains(driver)
                actions.move_to_element(showMoreElem).click().perform()
                time.sleep(1) # wait to load new articles
            URLs.extend(pageURLs)

        driver.quit()

        # Làm việc với HTMl bằng Beautiful Soup
        saved_timestamps = []
        saved_tickers = []
        saved_titles = []
        saved_descriptions = []
        saved_bodies = []
        saved_urls = []

        with requests.Session() as session:
            retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
            session.mount('https://',adapter)
            headers = {
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
            }
            for url in URLs:
                try:
                    articleHTML = session.get(url,headers=headers,timeout=30).text
                    articleBS = BeautifulSoup(articleHTML,'html5lib')
                    # Title
                    articleTitle = articleBS.find(class_='article__header').get_text(strip=True)
                    print(articleTitle)
                    # Time
                    articleTime = self.__processTime__(articleBS.find(name='time').get_text(strip=True))
                    print(articleTime)
                    # Description
                    articleDescription = articleBS.find(class_='article__sapo').get_text(strip=True)
                    print(articleDescription)
                    # Body
                    articleBody = '\n'.join(tag.get_text(strip=True) for tag in articleBS.find(class_='article__body').find_all('p'))
                    print(articleBody)
                    # Tickers
                    pattern = r'\b[A-Z]{1}[A-Z0-9]{2}\b'
                    articleTickers = set(re.findall(pattern,f'{articleTitle}\n{articleDescription}\n{articleBody}'))
                    print(articleTickers)

                    saved_timestamps.append(articleTime)
                    saved_tickers.append(','.join(articleTickers))
                    saved_titles.append(articleTitle)
                    saved_descriptions.append(articleDescription)
                    saved_bodies.append(articleBody)
                    saved_urls.append(url)

                except AttributeError:
                    print(f'{url} là một interactive article --> bỏ qua')

        result_dict = {
            'Time':saved_timestamps,
            'Ticker':saved_tickers,
            'Title':saved_titles,
            'Description':saved_descriptions,
            'Body':saved_bodies,
            'URL':saved_urls,
        }
        result = pd.DataFrame(result_dict)
        result['Time'] = result['Time'].fillna(method='ffill') # lấy giờ của tin trước

        if __name__=='__main__':
            print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
        else:
            print(f"{__name__.split('.')[-1]} ::: Finished")
        print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

        return result

