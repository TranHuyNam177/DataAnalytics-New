from request.stock import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
)
n_max_try = 100
max_wait_time = 10
num_hours = 120
bmk_time = btime(dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
fixedmp_list = internal.fixedmp_list
margin_list = internal.mlist()


def check_margin(ticker):
    if ticker in (margin_list+['HOSE','HNX','UPCOM']):
        result = 'Yes'
    elif len(ticker)>20:  # do có thể đăng tin sai format
        result = 'Unknown'
    else:
        result = 'No'
    return result


def check_fixed_mp(ticker):
    if ticker in fixedmp_list:
        result = 'Yes'
    elif len(ticker)>20:  # do có thể đăng tin sai format
        result = 'Unknown'
    else:
        result = 'No'
    return result


class vsd:

    @staticmethod
    def tinTCPH() -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        (tin từ tổ chức phát hành).
    
        :return: summary report of news update
        """

        start_time = time.time()
        now = dt.datetime.now()
        fromtime = now
        url = 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        driver = webdriver.Chrome(executable_path=PATH)
        driver.get(url)
        keywords = [
            'Cổ tức',
            'Tạm ứng cổ tức',
            'Tạm ứng',
            'Chi trả',
            'Bằng tiền',
            'Cổ phiếu',
            'Quyền mua',
            'Chuyển dữ liệu đăng ký',
            'cổ tức',
            'tạm ứng cổ tức',
            'tạm ứng',
            'chi trả',
            'bằng tiền',
            'cổ phiếu',
            'quyền mua',
            'chuyển dữ liệu đăng ký'
        ]
        frames = []
        while fromtime>=dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):
            try:  # catch all exceptions and retry
                news_times = []
                news_headlines = []
                news_urls = []
                tags = driver.find_elements_by_xpath('//*[@id="d_list_news"]/ul/li')
                for tag in tags:
                    h3_tag = tag.find_element_by_tag_name('h3')
                    txt = h3_tag.text
                    check = [word in txt for word in keywords]
                    if any(check):
                        news_headlines += [txt]
                        sub_url = h3_tag.find_element_by_tag_name('a').get_attribute('href')
                        news_urls += [f'=HYPERLINK("{sub_url}","Link")']
                        news_time = tag.find_element_by_tag_name('div').text
                        news_time = dt.datetime.strptime(news_time[-21:],'%d/%m/%Y - %H:%M:%S')
                        news_times += [news_time]

                frame = pd.DataFrame({
                    'Thời gian':news_times,
                    'Tiêu đề':news_headlines,
                    'Link':news_urls
                })
                frames.append(frame)
                # Turn Page
                nextpage_button = driver.find_elements_by_xpath(
                    "//button[substring(@onclick,1,string-length('changePage'))='changePage']"
                )[-2]
                nextpage_button.click()
                # Check time
                last_tag = driver.find_elements_by_xpath('//*[@id="d_list_news"]/ul/li')[-1]
                last_time = last_tag.find_element_by_tag_name('div').text
                fromtime = dt.datetime.strptime(last_time[-21:],'%d/%m/%Y - %H:%M:%S')
            except (Exception,):
                continue

        driver.quit()
        output_table = pd.concat(frames,ignore_index=True)
        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        def f(frame):
            series = frame['Tiêu đề'].str.split(': ',n=1)
            table = pd.DataFrame(
                series.to_list(),
                columns=['cophieu','tieude']
            )
            return table

        output_table[['Mã cổ phiếu','Lý do, mục đích']] = output_table.transform(f)
        output_table.drop('Tiêu đề',axis=1,inplace=True)
        output_table['Đang cho vay'] = output_table['Mã cổ phiếu'].map(check_margin)
        output_table['Fixed max price'] = output_table['Mã cổ phiếu'].map(check_fixed_mp)
        output_table = output_table[[
            'Thời gian',
            'Mã cổ phiếu',
            'Đang cho vay',
            'Fixed max price',
            'Lý do, mục đích',
            'Link'
        ]]
        print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

        return output_table

    @staticmethod
    def tinTVBT() -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://vsd.vn/vi/tin-thi-truong-phai-sinh'
        (tin từ thành viên bù trừ).
    
        :return: summary report of news update
        """

        start_time = time.time()
        now = dt.datetime.now()
        fromtime = now
        url = 'https://vsd.vn/vi/tin-thi-truong-phai-sinh'
        driver = webdriver.Chrome(executable_path=PATH)
        driver.get(url)
        keywords = [
            'tỷ lệ ký quỹ ban đầu',
            'Tỷ lệ ký quỹ ban đầu',
            'hợp đồng tương lai',
            'Hợp đồng tương lai',
            'HĐTL'
        ]
        frames = []
        while fromtime>=dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):
            news_times = []
            news_headlines = []
            news_urls = []
            tags = driver.find_elements_by_xpath('//*[@id="tab1"]/ul/li')
            for tag in tags:
                h3_tag = tag.find_element_by_tag_name('h3')
                if not h3_tag.text[:3].isupper():
                    txt = h3_tag.text
                    check = [word in txt for word in keywords]
                    if any(check):
                        news_headlines += [txt]
                        sub_url = h3_tag.find_element_by_tag_name('a').get_attribute('href')
                        news_urls += [f'=HYPERLINK("{sub_url}","Link")']
                        news_time = tag.find_element_by_tag_name('div').text
                        news_time = dt.datetime.strptime(news_time[-21:],'%d/%m/%Y - %H:%M:%S')
                        news_times += [news_time]
            frame = pd.DataFrame({
                'Thời gian':news_times,
                'Tiêu đề':news_headlines,
                'Link':news_urls
            })
            frames.append(frame)
            # Turn Page
            nextpage_button = driver.find_elements_by_xpath("//*[@id='d_number_of_page']/button")[-2]
            nextpage_button.click()
            # Check time
            last_tag = driver.find_elements_by_xpath('//*[@id="tab1"]/ul/li')[-1]
            last_time = last_tag.find_element_by_tag_name('div').text
            fromtime = dt.datetime.strptime(last_time[-21:],'%d/%m/%Y - %H:%M:%S')

        driver.quit()
        output_table = pd.concat(frames,ignore_index=True)
        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

        return output_table


vsd = vsd()


class hnx:

    @staticmethod
    def tinTCPH() -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://www.hnx.vn/thong-tin-cong-bo-ny-tcph.html'
        (tin từ tổ chức phát hành).
    
        :return: summary report of news update
        """

        start_time = time.time()
        now = dt.datetime.now()
        fromtime = now
        url = 'https://www.hnx.vn/thong-tin-cong-bo-ny-tcph.html'
        driver = webdriver.Chrome(executable_path=PATH)
        driver.get(url)
        driver.maximize_window()
        wait = WebDriverWait(driver,max_wait_time,ignored_exceptions=ignored_exceptions)
        key_words = [
            'Cổ tức',
            'Tạm ứng cổ tức',
            'Tạm ứng',
            'Chi trả',
            'Bằng tiền',
            'Quyền mua',
            'cổ tức',
            'tạm ứng cổ tức',
            'tạm ứng',
            'chi trả',
            'bằng tiền',
            'quyền mua'
        ]
        links = []
        box_text = []
        titles = []
        times = []
        tickers = []
        while fromtime>=dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):
            n_try = 1
            while n_try<n_max_try:  # try certain number of times, if fail throw an error and the function rerun under implementation
                try:
                    ticker_elems = wait.until(
                        EC.presence_of_all_elements_located((By.XPATH,"//*[@id='_tableDatas']/tbody/*/td[3]/a"))
                    )
                    title_elems = wait.until(
                        EC.presence_of_all_elements_located((By.XPATH,"//*[@id='_tableDatas']/tbody/*/td[5]/a"))
                    )
                    tickers_inpage = [t.text for t in ticker_elems if t.text!='']
                    titles_inpage = [t.text for t in title_elems if t.text!='']
                    title_elems = [t for t in title_elems if t.text!='']
                    ticker_title_inpage = [
                        f'{ticker} {title}' for ticker,title in zip(tickers_inpage,titles_inpage)
                    ]
                    break
                except (Exception,):
                    n_try += 1
                    time.sleep(1)
            else:
                raise PageFailToLoad
            for ticker_title in ticker_title_inpage:
                sub_check = [key_word in ticker_title for key_word in key_words]
                if any(sub_check):
                    n_try = 1
                    while n_try<n_max_try:  # try certain number of times, if fail throw an error and the function rerun under implementation
                        try:
                            title_no = ticker_title_inpage.index(ticker_title)+1
                            ticker = driver.find_elements_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[3]")[-1].text
                            title = driver.find_elements_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[5]")[-1].text
                            t = driver.find_elements_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[2]")[-1].text
                            titles += [title]
                            times += [dt.datetime.strptime(t,'%d/%m/%Y %H:%M')]
                            tickers += [ticker]
                            # open popup window
                            click_obj = title_elems[title_no-1]
                            click_obj.click()
                            break
                        except (Exception,):
                            n_try += 1
                            time.sleep(1)
                    else:
                        raise PageFailToLoad
                    time.sleep(1)  # wait for popup window appears
                    popup_content = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="divViewDetailArticles"]/div[2]/div[2]')))
                    content = popup_content.text
                    if content in ['.','',' ']:  # nếu có link, ko có nội dung
                        box_text += ['']
                        link_elems = driver.find_elements_by_xpath("//div[@class='divLstFileAttach']/p/a")
                        links += [[link.get_attribute('href') for link in link_elems]]
                    else:  # nếu có nội dung, ko có link
                        links += ['']
                        box_text += [content]
                    # exit pop-up window
                    wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@value='Thoát']"))).click()
            # check time
            last_time = driver.find_elements_by_xpath("//*[@id='_tableDatas']/tbody/tr[10]/td[2]")[-1].text
            fromtime = dt.datetime.strptime(last_time,'%d/%m/%Y %H:%M')
            # turn page
            n_try = 1
            while n_try<n_max_try:  # try certain number of times, if fail throw an error and the function rerun under implementation
                try:
                    driver.find_elements_by_xpath("//*[@id='d_number_of_page']/li")[-2].click()
                    break
                except (Exception,):
                    n_try += 1
            else:
                raise PageFailToLoad

        driver.quit()
        records = list(zip(times,tickers,titles,box_text,links))
        df = pd.DataFrame(
            records,
            columns=[
                'Thời gian',
                'Mã CK',
                'Tiêu đề',
                'Nội dung',
                'File đính kèm',
            ]
        )
        if df.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        def convert_link(full_list):  # tách list có nhiều link
            if len(full_list)==1:
                result = f'=HYPERLINK("{full_list[0]}","Link")'
            elif len(full_list)>1:
                result = ''
                result += f'=HYPERLINK("{full_list[0]}","Link")'
                for i in range(1,len(full_list)):
                    result += f'&" "&HYPERLINK("{full_list[i]}","Link")'
            else:
                result = None
            return result

        df['File đính kèm'] = df['File đính kèm'].map(convert_link)
        df['Đang cho vay'] = df['Mã CK'].map(check_margin)
        df['Fixed max price'] = df['Mã CK'].map(check_fixed_mp)
        df = df[[
            'Thời gian',
            'Mã CK',
            'Đang cho vay',
            'Fixed max price',
            'Tiêu đề',
            'Nội dung',
            'File đính kèm'
        ]]

        print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

        return df

    @staticmethod
    def tintuso() -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html'
        (tin từ sở). If upcom = True, include news update published in
        'https://www.hnx.vn/vi-vn/thong-tin-cong-bo-up-hnx.html' (tin từ sở)
    
        :return: summary report of news update
        """

        start_time = time.time()
        now = dt.datetime.now()
        from_time = now
        driver = webdriver.Chrome(executable_path=PATH)
        url = 'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html'
        driver.maximize_window()
        driver.get(url)
        wait = WebDriverWait(driver,max_wait_time,ignored_exceptions=ignored_exceptions)
        key_words = [
            'tạm ngừng giao dịch',
            'vào diện',
            'hủy niêm yết',
            'ra khỏi diện',
            'hủy bỏ niêm yết',
            'kiểm soát',
            'cảnh cáo',
            'cảnh báo',
            'không đủ điều kiện giao dịch ký quỹ',
            'ra khỏi',
            'vào',
            'hạn chế giao dịch',
            'chuyển giao dịch',
        ]
        links = []
        titles = []
        times = []
        tickers = []
        while from_time>=dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):
            n_try = 1
            while n_try<n_max_try:  # try certain number of times, if fail throw an error and the function rerun under implementation
                try:
                    ticker_elems = wait.until(EC.presence_of_all_elements_located((By.XPATH,"//*[@id='_tableDatas']/tbody/*/td[3]/a")))
                    title_elems = wait.until(EC.presence_of_all_elements_located((By.XPATH,'//*[@id="_tableDatas"]/tbody/*/td[4]/a')))
                    tickers_inpage = [t.text for t in ticker_elems if t.text!='']
                    titles_inpage = [t.text for t in title_elems if t.text!='']
                    title_elems = [t for t in title_elems if t.text!='']
                    ticker_title_inpage = [f'{ticker} {title}' for ticker,title in zip(tickers_inpage,titles_inpage)]
                    break
                except (Exception,):
                    n_try += 1
                    time.sleep(1)
            else:
                raise PageFailToLoad
            for ticker_title in ticker_title_inpage:
                check = [key_word in ticker_title for key_word in key_words]
                if any(check):
                    n_try = 1
                    while n_try<n_max_try:  # try certain number of times, if fail throw an error and the function rerun under implementation
                        try:
                            title_no = ticker_title_inpage.index(ticker_title)+1
                            titles += [
                                driver.find_elements_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[4]")[0].text
                            ]
                            times += [
                                dt.datetime.strptime(
                                    driver.find_elements_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[2]")[0].text,'%d/%m/%Y %H:%M'
                                )
                            ]
                            tickers += [
                                driver.find_elements_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[3]")[0].text
                            ]
                            # open popup window:
                            click_obj = title_elems[title_no-1]
                            click_obj.click()
                            time.sleep(1)  # wait for popup window appears
                            link_elems = driver.find_elements_by_xpath(
                                "//div[@class='divLstFileAttach']/p/a"
                            )
                            links += [[link.get_attribute('href') for link in link_elems]]
                            # exit popup window
                            wait.until(EC.element_to_be_clickable
                                       ((By.XPATH,"//*[@id='divViewDetailArticles']/*/input"))
                                       ).click()
                            break
                        except (Exception,):
                            n_try += 1
                            time.sleep(1)
                    else:
                        raise PageFailToLoad
            # check time
            last_time = driver.find_element_by_xpath(
                "//*[@id='_tableDatas']/tbody/tr[10]/td[2]"
            ).text
            from_time = dt.datetime.strptime(last_time,'%d/%m/%Y %H:%M')
            # turn page
            n_try = 1
            while n_try<n_max_try:  # try certain number of times, if fail throw an error and the function rerun under implementation
                try:
                    driver.find_element_by_xpath("//*[@id='next']").click()
                    break
                except (Exception,):
                    n_try += 1
            else:
                raise PageFailToLoad

        driver.quit()

        # export to DataFrame
        df = pd.DataFrame(
            list(zip(times,tickers,titles,links)),
            columns=[
                'Thời gian',
                'Mã CK',
                'Tiêu đề',
                'File đính kèm'
            ]
        )
        df.insert(2,'Đang cho vay',df['Mã CK'].map(check_margin))
        df.insert(3,'Fixed max price',df['Mã CK'].map(check_fixed_mp))

        if df.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        def f(full_list):
            if len(full_list)==1:
                result = f'=HYPERLINK("{full_list[0]}","Link")'
            elif len(full_list)>1:
                result = ''
                result += f'=HYPERLINK("{full_list[0]}","Link")'
                for i in range(1,len(full_list)):
                    result += f'&" "&HYPERLINK("{full_list[i]}","Link")'
            else:
                result = None
            return result

        df['File đính kèm'] = df['File đính kèm'].map(f)

        print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

        return df


hnx = hnx()


class hose:

    @staticmethod
    def tintonghop() -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in
        'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/822d8a8c-fd19-4358-9fc9-d0b27a666611?fid=0318d64750264e31b5d57c619ed6b338'
        (tin từ sở).
    
        :return: summary report of news update
        """

        start_time = time.time()
        # Convert str to datetime
        def f(s):
            if s.endswith('SA'):
                s = s.rstrip(' SA')
                return dt.datetime.strptime(s, '%d/%m/%Y %H:%M:%S')
            if s.endswith('CH'):
                s = s.rstrip(' CH')
                return dt.datetime.strptime(s, '%d/%m/%Y %H:%M:%S') + dt.timedelta(hours=12)

        def f1(full_list):
            if len(full_list) == 1:
                result = f'=HYPERLINK("{full_list[0]}","Link")'
            elif len(full_list) > 1:
                result = ''
                result += f'=HYPERLINK("{full_list[0]}","Link")'
                for i in range(1, len(full_list)):
                    result += f'&" "&HYPERLINK("{full_list[i]}","Link")'
            else:
                result = None
            return result

        now = dt.datetime.now()
        keywords = [
            'tạm ngừng giao dịch',
            'vào diện',
            'hủy niêm yết',
            'ra khỏi diện',
            'kiểm soát',
            'cảnh cáo',
            'cảnh báo',
            'hủy đăng ký',
            'hủy bỏ',
            'không đủ điều kiện giao dịch ký quỹ',
            'hạn chế giao dịch',
            'chuyển giao dịch',
            'giao dịch đầu tiên cổ phiếu niêm yết',
            'giao dịch đầu tiên của cổ phiếu niêm yết',
        ]
        excl_keywords = ['ban', 'Ban', 'bầu', 'Bầu', 'nhắc nhở', 'Nhắc nhở', 'họp', 'Họp']

        frames = []
        page = 1
        while True:
            pageFieldValue2 = now.strftime("%d.%m.%Y")
            pageFieldValue4 = '822d8a8c-fd19-4358-9fc9-d0b27a666611'
            web_url = 'https://www.hsx.vn/Modules/CMS/Web/ArticleInCategory/822d8a8c-fd19-4358-9fc9-d0b27a666611'
            f_url = f'{web_url}?exclude=00000000-0000-0000-0000-000000000000&lim=True&pageFieldName1=FromDate&' \
                    f'pageFieldValue1={pageFieldValue2[0:2]}.10.2021&pageFieldOperator1=eq&pageFieldName2=ToDate&' \
                    f'pageFieldValue2={pageFieldValue2}&pageFieldOperator2=eq&pageFieldName3=TokenCode&' \
                    f'pageFieldValue3=&pageFieldOperator3=eq&pageFieldName4=CategoryId&' \
                    f'pageFieldValue4={pageFieldValue4}&pageFieldOperator4=eq&pageCriteriaLength=4&' \
                    f'_search=false&nd=1649931964869&rows=30&page={page}&sidx=id&sord=desc'

            with requests.Session() as session:
                retry = requests.packages.urllib3.util.retry.Retry(connect=5, backoff_factor=1)
                adapter = requests.adapters.HTTPAdapter(max_retries=retry)
                session.mount('https://', adapter)
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36'
                }
                html = session.get(f_url,headers=headers,timeout=10).content
                soup = BeautifulSoup(html,'html5lib')
                news = json.loads(soup.text)

                def get_content_from_row(n):
                    ID, timeString, rawTitle = news['rows'][n]['cell']
                    title = BeautifulSoup(rawTitle,'html5lib').find('a',href=True).text
                    articleTime = f(timeString)
                    return ID, articleTime, title

                get_cell = [d.get('cell') for d in news['rows']]
                for row in range(len(get_cell)):
                    id_url,t,sub_title = get_content_from_row(row)
                    nxt_url = f'https://www.hsx.vn/Modules/Cms/Web/ViewArticle/{id_url}'
                    check_1 = [word in sub_title for word in keywords]
                    check_2 = [word not in sub_title for word in excl_keywords]
                    if any(check_1) and all(check_2):
                        print(t,sub_title,nxt_url)
                        # report and merge to parent table
                        frame = pd.DataFrame(
                            {
                                'Thời gian': t,
                                'Tiêu đề': sub_title,
                                'Nội dung': '',
                                'Link': f'=HYPERLINK("{nxt_url}","Link")'
                            },
                            index=[0],
                        )
                        frames.append(frame)
                    format_time = t
                if format_time < dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
                    break
            page += 1

        #######################################################################

        page = 1
        while True:
            pageFieldValue2 = now.strftime("%d.%m.%Y")
            pageFieldValue4 = 'dca0933e-a578-4eaf-8b29-beb4575052c5'
            web_url = 'https://www.hsx.vn/Modules/CMS/Web/ArticleInCategory/dca0933e-a578-4eaf-8b29-beb4575052c5'
            f_url = f'{web_url}?exclude=00000000-0000-0000-0000-000000000000&lim=True&pageFieldName1=FromDate&' \
                    f'pageFieldValue1={pageFieldValue2[0:2]}.10.2021&pageFieldOperator1=eq&pageFieldName2=ToDate&' \
                    f'pageFieldValue2={pageFieldValue2}&pageFieldOperator2=eq&pageFieldName3=TokenCode&' \
                    f'pageFieldValue3=&pageFieldOperator3=eq&pageFieldName4=CategoryId&' \
                    f'pageFieldValue4={pageFieldValue4}&pageFieldOperator4=eq&pageCriteriaLength=4&' \
                    f'_search=false&nd=1649990058235&rows=30&page={page}&sidx=id&sord=desc'

            with requests.Session() as session:
                retry = requests.packages.urllib3.util.retry.Retry(connect=5, backoff_factor=1)
                adapter = requests.adapters.HTTPAdapter(max_retries=retry)
                session.mount('https://', adapter)
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36'
                }
                html = session.get(f_url,headers=headers,timeout=10).content
                soup = BeautifulSoup(html,'html5lib')
                news = json.loads(soup.text)

                def get_content_from_row(n):
                    ID, timeString, rawTitle = news['rows'][n]['cell']
                    title = BeautifulSoup(rawTitle,'html5lib').find('span').text
                    articleTime = f(timeString)
                    return ID, articleTime, title

                get_cell = [d.get('cell') for d in news['rows']]
                for row in range(len(get_cell)):
                    sub_pdfs = []
                    id_url,t,sub_title = get_content_from_row(row)
                    print(sub_title)
                    check_1 = [word in sub_title for word in keywords]
                    check_2 = [word not in sub_title for word in excl_keywords]
                    if any(check_1) and all(check_2):
                        nxt_rq = f'https://www.hsx.vn/Modules/Cms/Web/LoadArticle?id={id_url}&objectType=1'
                        html = session.get(nxt_rq,headers=headers,timeout=10).content
                        soup = BeautifulSoup(html,'lxml')
                        get_content = soup.find_all('p')
                        paragraph = '\n'.join([con.text.replace('\xa0',' ') for con in get_content])
                        urls = soup.find_all(href=True)
                        get_pdf = ['https://www.hsx.vn'+pdf.get('href') for pdf in urls]
                        sub_pdfs += [f1(get_pdf)]
                        print(t,sub_title,paragraph,sub_pdfs)
                        # report and merge to parent table
                        frame = pd.DataFrame(
                            {
                                'Thời gian': [t],
                                'Tiêu đề': [sub_title],
                                'Nội dung': paragraph,
                                'Link': sub_pdfs
                            },
                            index=[0],
                        )
                        frames.append(frame)
                    format_time = t
                if format_time < dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
                    break
            page += 1

        if not frames:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')
        output_table = pd.concat(frames,ignore_index=True)
        # select out tickers from headline
        output_table.insert(1,'Mã cổ phiếu',output_table['Tiêu đề'].str.split(': ').str.get(0))
        # check if ticker is in margin list
        output_table.insert(2,'Đang cho vay',output_table['Mã cổ phiếu'].map(check_margin))
        output_table.insert(3,'Fixed max price',output_table['Mã cổ phiếu'].map(check_fixed_mp))
        output_table = output_table.drop_duplicates(subset=['Thời gian','Mã cổ phiếu','Tiêu đề'])

        print(f'Finished hose-tintonghop ::: Total execution time: {int(time.time() - start_time)}s\n')

        return output_table


hose = hose()