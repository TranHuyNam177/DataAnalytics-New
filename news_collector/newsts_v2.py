from request.stock import *


class NoNewsFound(Exception):
    pass


class vsd:

    def __init__(self):

        self.PATH = join(dirname(dirname(realpath(__file__))), 'dependency', 'chromedriver')
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException
        )

    def tinTCPH(self, num_hours: int = 48) -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        (tin từ tổ chức phát hành).
        :param num_hours: number of hours in the past that's in our concern
        :return: summary report of news update
        """

        start_time = time.time()
        now = dt.datetime.now()
        fromtime = now
        url = 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        driver = webdriver.Chrome(executable_path=self.PATH)
        driver.get(url)
        driver.maximize_window()
        keywords = [
            'Hủy đăng ký chứng khoán',
            'chứng nhận đăng ký chứng khoán lần đầu',
            'chứng nhận đăng ký chứng quyền lần đầu'
        ]
        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        frames = []
        while fromtime >= dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
            news_time = []
            news_headlines = []
            news_urls = []
            news_noidung = []

            time.sleep(1)
            driver_wait = WebDriverWait(driver, 5, ignored_exceptions=self.ignored_exceptions)
            list_news_elem = driver_wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'list-news')))
            tags = list_news_elem.find_elements_by_tag_name('li')
            for tag_ in tags:
                tag_wait = WebDriverWait(tag_, 5, ignored_exceptions=self.ignored_exceptions)
                h3_tag = tag_wait.until(EC.presence_of_element_located((By.TAG_NAME, 'h3')))
                txt = h3_tag.text
                check = [word in txt for word in keywords]
                if any(check):
                    news_headlines += [txt]
                    tag_wait = WebDriverWait(h3_tag, 5, ignored_exceptions=self.ignored_exceptions)
                    sub_url = tag_wait.until(EC.presence_of_element_located((By.TAG_NAME, 'a'))).get_attribute('href')
                    news_urls += [f'=HYPERLINK("{sub_url}","Link")']
                    news_time_ = tag_.find_element_by_tag_name('div').text
                    news_time_ = dt.datetime.strptime(news_time_[-21:], '%d/%m/%Y - %H:%M:%S')
                    news_time += [news_time_]
                    sub_driver = webdriver.Chrome(executable_path=self.PATH)
                    sub_driver.get(sub_url)
                    sub_driver_wait = WebDriverWait(sub_driver, 5, ignored_exceptions=self.ignored_exceptions)
                    noidung_elem = sub_driver_wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'col-md-12')))
                    noidung = noidung_elem.text
                    news_noidung += [noidung]
                    sub_driver.quit()

            frame = pd.DataFrame({
                'Thời gian': news_time,
                'Tiêu đề': news_headlines,
                'Nội dung': news_noidung,
                'Link': news_urls
            })
            frames.append(frame)

            # Turn Page
            time.sleep(1)
            button_row = driver_wait.until(EC.presence_of_element_located((By.ID, 'd_number_of_page')))
            button_elems = button_row.find_elements_by_xpath(".//*")
            nextpage_button = button_elems[-3]
            nextpage_button.click()

            # Check time
            last_tag = \
            driver_wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="d_list_news"]/ul/li')))[-1]
            fromtime = last_tag.find_element_by_tag_name('div').text
            fromtime = dt.datetime.strptime(fromtime[-21:], '%d/%m/%Y - %H:%M:%S')

        output_table = pd.concat(frames, ignore_index=True)
        driver.quit()

        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        def f(frame):
            series = frame['Tiêu đề'].str.split(': ')
            table = pd.DataFrame(series.tolist(), columns=['cophieu', 'tieude'])
            return table

        output_table[['Mã cổ phiếu', 'Lý do']] = output_table.transform(f)
        output_table.drop(['Tiêu đề'], axis=1, inplace=True)
        output_table = output_table[[
            'Thời gian',
            'Mã cổ phiếu',
            'Lý do',
            'Nội dung',
            'Link'
        ]]
        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return output_table

    def tinTVLK(self, num_hours: int = 48) -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://www.vsd.vn/vi/alc/4'
        (tin từ thành viên lưu ký).
        :param num_hours: number of hours in the past that's in our concern
        :return: summary report of news update
        """

        start_time = time.time()
        now = dt.datetime.now()
        fromtime = now

        url = 'https://www.vsd.vn/vi/alc/4'
        driver = webdriver.Chrome(executable_path=self.PATH)
        driver.get(url)
        driver.maximize_window()

        keywords = [
            'ngày hạch toán của cổ phiếu',
            'ngày hạch toán cổ phiếu',
            'ngày hạch toán của chứng quyền',
            'ngày hạch toán chứng quyền'
        ]
        excluded_words = ['bổ sung']
        driver_wait = WebDriverWait(driver, 5, ignored_exceptions=self.ignored_exceptions)
        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        frames = []
        while fromtime >= dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
            news_time = []
            news_headlines = []
            news_urls = []
            news_content = []
            time.sleep(1)
            tags = driver_wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='d_list_news']/ul/li")
            ))
            for tag_ in tags:
                tag_wait = WebDriverWait(tag_, 5, ignored_exceptions=self.ignored_exceptions)
                h3_tag = tag_wait.until(EC.presence_of_element_located((By.TAG_NAME, 'h3')))
                txt = h3_tag.text
                check_1 = [word not in txt for word in excluded_words]
                check_2 = [word in txt for word in keywords]
                if all(check_1) and any(check_2):
                    news_headlines.append(txt)
                    sub_url = h3_tag.find_element_by_tag_name('a').get_attribute('href')
                    news_urls.append(f'=HYPERLINK("{sub_url}","Link")')
                    news_time_ = tag_.find_element_by_tag_name('div').text
                    news_time_ = dt.datetime.strptime(news_time_[-21:], '%d/%m/%Y - %H:%M:%S')
                    news_time.append(news_time_)
                    sub_driver = webdriver.Chrome(executable_path=self.PATH)
                    sub_driver.get(sub_url)
                    sub_driver_wait = WebDriverWait(sub_driver, 5, ignored_exceptions=self.ignored_exceptions)
                    content = sub_driver_wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'col-md-12'))).text
                    news_content.append(content)
                    sub_driver.quit()

            frame = pd.DataFrame({
                'Thời gian': news_time,
                'Tiêu đề': news_headlines,
                'Nội dung': news_content,
                'Link': news_urls
            })
            frames.append(frame)

            # Turn Page
            nextpage_button = driver_wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='d_number_of_page']/button")
            ))[-2]
            nextpage_button.click()

            # Check time
            last_tag = driver_wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, '//*[@id="d_list_news"]/ul/li')
            ))[-1]
            fromtime = last_tag.find_element_by_tag_name('div').text
            fromtime = dt.datetime.strptime(fromtime[-21:], '%d/%m/%Y - %H:%M:%S')

        output_table = pd.concat(frames, ignore_index=True)

        def f(frame):
            series = frame['Tiêu đề'].str.split(': ')
            table = pd.DataFrame(series.tolist(), columns=['cophieu/chungquyen', 'tieude'])
            return table

        if output_table['Tiêu đề'].empty:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')
        else:
            output_table[['Mã cổ phiếu / chứng quyền', 'Lý do, mục đích']] = output_table.transform(f)

        driver.quit()

        output_table.drop(['Tiêu đề'], axis=1, inplace=True)
        output_table = output_table[[
            'Thời gian',
            'Mã cổ phiếu / chứng quyền',
            'Lý do, mục đích',
            'Nội dung',
            'Link'
        ]]
        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return output_table


class hnx:

    def __init__(self):

        self.PATH = join(dirname(dirname(realpath(__file__))), 'dependency', 'chromedriver')
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException
        )

    def tintuso(self, num_hours: int = 48):

        """
        This function returns a DataFrame and export an excel file containing
        news update published in
        'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html' (tin từ sở) and
        'https://www.hnx.vn/vi-vn/thong-tin-cong-bo-up-hnx.html' (tin từ sở)
        :param num_hours: number of hours in the past that's in our concern
        :return: summary report of news update
        """

        start_time = time.time()
        now = dt.datetime.now()
        from_time = now

        driver = webdriver.Chrome(executable_path=self.PATH)

        url = 'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html'
        driver.get(url)
        driver.maximize_window()

        key_words = [
            'ngày giao dịch đầu tiên cổ phiếu niêm yết',
            'ngày giao dịch đầu tiên của cổ phiếu niêm yết',
            'ngày giao dịch đầu tiên cổ phiếu chuyển giao dịch',
            'ngày giao dịch đầu tiên của cổ phiếu chuyển giao dịch',
            'ngày hủy niêm yết',
            'ngày giao dịch đầu tiên cổ phiếu đăng ký giao dịch',
            'ngày giao dịch đầu tiên của cổ phiếu đăng ký giao dịch',
            'ngày hủy đăng ký giao dịch cổ phiếu',
            'ngày hủy đăng ký giao dịch của cổ phiếu',
            'ngày hủy ĐKGD cổ phiếu',
            'ngày hủy ĐKGD của cổ phiếu'
        ]
        excluded_words = ['bổ sung']

        links = []
        box_text = []
        titles = []
        times = []
        tickers = []

        wait = WebDriverWait(driver, timeout=5, poll_frequency=1)
        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        while from_time >= dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
            # wait for the element to appear, avoid stale element reference
            ticker_elems = wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[3]/a")
            ))
            title_elems = wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[4]/a")
            ))
            tickers_inpage = [t.text for t in ticker_elems if t.text != '']
            titles_inpage = [t.text for t in title_elems if t.text != '']
            title_elems = [t for t in title_elems if t.text != '']

            ticker_title_inpage = [f'{ticker} {title}' for ticker, title in zip(tickers_inpage, titles_inpage)]

            for ticker_title in ticker_title_inpage:

                subcheck_1 = [key_word in ticker_title for key_word in key_words]
                subcheck_2 = [excl_word not in ticker_title for excl_word in excluded_words]

                if any(subcheck_1) and all(subcheck_2):
                    pass
                else:
                    continue

                title_no = ticker_title_inpage.index(ticker_title) + 1
                titles.append(driver.find_element_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[4]").text)
                times.append(driver.find_element_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[2]").text)
                tickers.append(driver.find_element_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[3]").text)
                # open popup window:
                click_obj = title_elems[title_no - 1]
                click_obj.click()
                time.sleep(1)
                # evaluate popup content
                popup_content = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='Box-Noidung']")))
                content = popup_content.text
                box_text.append(content)
                link_elems = driver.find_elements_by_xpath("//div[@class='divLstFileAttach']/p/a")
                links.append([link.get_attribute('href') for link in link_elems])

                # close popup windows
                wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divViewDetailArticles"]/*/input'))).click()
                time.sleep(1)
                # check time
                from_time = driver.find_element_by_xpath("//*[@id='_tableDatas']/tbody/tr[10]/td[2]").text
                from_time = dt.datetime.strptime(from_time, '%d/%m/%Y %H:%M')

            # next page
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='next']"))).click()
            time.sleep(1)

        url = 'https://www.hnx.vn/vi-vn/thong-tin-cong-bo-up-hnx.html'
        driver.get(url)
        driver.maximize_window()
        fromtime = now
        wait = WebDriverWait(driver, timeout=5, poll_frequency=1)
        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        while fromtime >= dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
            ticker_elems = wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[3]/a")
            ))
            title_elems \
                = wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[4]/a")
            ))
            tickers_inpage = [t.text for t in ticker_elems if t.text != '']
            titles_inpage = [t.text for t in title_elems if t.text != '']
            title_elems = [t for t in title_elems if t.text != '']
            ticker_title_inpage = [f'{ticker} {title}' for ticker, title in zip(tickers_inpage, titles_inpage)]

            for ticker_title in ticker_title_inpage:
                sub_check = [key_word in ticker_title for key_word in key_words]
                if any(sub_check):
                    pass
                else:
                    continue
                title_no = ticker_title_inpage.index(ticker_title) + 1
                titles.append(driver.find_element_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[4]").text)
                times.append(driver.find_element_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[2]").text)
                tickers.append(driver.find_element_by_xpath(f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[3]").text)

                # open popup window:
                click_obj = title_elems[title_no - 1]
                click_obj.click()
                time.sleep(1)
                popup_content = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//div[@class='Box-Noidung']")
                ))
                content = popup_content.text
                if content in ['.', '', ' ']:  # nếu có link, ko có nội dung
                    box_text.append('')
                    link_elems = driver.find_elements_by_xpath("//div[@class='divLstFileAttach']/p/a")
                    links.append([link.get_attribute('href') for link in link_elems])
                else:  # nếu có nội dung, ko có link
                    links.append('')
                    box_text.append(content)
                wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='divViewDetailArticles']/*/input"))).click()
                time.sleep(1)
                # check time
                fromtime = driver.find_element_by_xpath("//*[@id='_tableDatas']/tbody/tr[10]/td[2]").text
                fromtime = dt.datetime.strptime(fromtime, '%d/%m/%Y %H:%M')

            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='next']"))).click()
            time.sleep(1)

        driver.quit()

        # export to DataFrame
        df = pd.DataFrame(
            data=list(zip(times, tickers, titles, box_text, links)),
            columns=[
                'Thời gian',
                'Mã CK',
                'Tiêu đề',
                'Nội dung',
                'File đính kèm'
            ]
        )
        df['Thời gian'] = pd.to_datetime(df['Thời gian'], format='%d/%m/%Y %H:%M')
        df.sort_values('Thời gian', ascending=False, inplace=True)

        if df.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        def f(full_list):
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

        df['File đính kèm'] = df['File đính kèm'].map(f)
        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return df


class hose:
    def tinTCNY(self, num_hours: int = 48):

        """
        This function returns a DataFrame and export an excel file containing
        news update published in
        'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/dca0933e-a578-4eaf-8b29-beb4575052c5?fid=6d1f1d5e9e6c4fb593077d461e5155e7'
        (tin về hoạt động của tổ chức niêm yết)
        :param num_hours: number of hours in the past that's in our concern
        :return: summary report of news update
        """

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

        start_time = time.time()
        now = dt.datetime.now()
        from_time = now
        keywords = [
            'niêm yết và ngày giao dịch đầu tiên',
            'hủy niêm yết',
            'hủy đăng ký',
            'chuyển giao dịch'
        ]
        excl_keywords = ['bổ sung']
        frames = []
        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        page = 1
        while from_time >= dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
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
                html = session.get(f_url, headers=headers, timeout=10).content
                soup = BeautifulSoup(html, 'html5lib')
                news = json.loads(soup.text)

                def get_content_from_row(n):
                    ID, timeString, rawTitle = news['rows'][n]['cell']
                    title = BeautifulSoup(rawTitle,'html5lib').find('span').text
                    articleTime = f(timeString)
                    return ID, articleTime, title

                get_cell = [d.get('cell') for d in news['rows']]
                for row in range(len(get_cell)):
                    sub_pdfs = []
                    id_url, time_, sub_title = get_content_from_row(row)

                    nxt_rq = f'https://www.hsx.vn/Modules/Cms/Web/LoadArticle?id={id_url}&objectType=1'
                    html = session.get(nxt_rq, headers=headers, timeout=10).content
                    soup = BeautifulSoup(html, 'lxml')
                    get_content = soup.find_all('p')
                    paragraph = '\n'.join([con.text.replace('\xa0', ' ') for con in get_content])
                    urls = soup.find_all(href=True)
                    get_pdf = ['https://www.hsx.vn' + pdf.get('href') for pdf in urls]

                    check_1 = [word in sub_title for word in keywords]
                    check_2 = [word not in sub_title for word in excl_keywords]
                    if any(check_1) and all(check_2):
                        sub_pdfs += [f1(get_pdf)]
                        print(time_, sub_title, paragraph, sub_pdfs)
                        # report and merge to parent table
                        frame = pd.DataFrame(
                            {
                                'Thời gian': [time_],
                                'Tiêu đề': [sub_title],
                                'Nội dung': paragraph,
                                'Link': sub_pdfs
                            },
                            index=[0],
                        )
                        frames.append(frame)
                    format_time = time_
                if format_time < dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
                    break
            page += 1

        output_table = pd.concat(frames, ignore_index=True)
        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        # select out tickers from headline
        output_table.insert(
            1,
            'Mã cổ phiếu',
            output_table['Tiêu đề'].str.split(': ').str.get(0)
        )
        output_table = output_table.drop_duplicates(subset=['Thời gian', 'Mã cổ phiếu', 'Tiêu đề'])
        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return output_table

    def tinCW(self, num_hours: int = 48):

        """
        This function returns a DataFrame and export an excel file containing
        news update published in
        'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/95cd3266-e6d1-42a3-beb5-20ed010aea4a?fid=f91940eef3384bcdbe96ee9aa3eefa04'
        (tin chứng quyền).
        :param num_hours: number of hours in the past that's in our concern
        :return: summary report of news update
        """

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

        start_time = time.time()
        now = dt.datetime.now()
        from_time = now
        keywords = ['chấp thuận niêm yết chứng quyền']
        frames = []
        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        page = 1
        while from_time >= dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
            pageFieldValue2 = now.strftime("%d.%m.%Y")
            pageFieldValue4 = '95cd3266-e6d1-42a3-beb5-20ed010aea4a'
            web_url = 'https://www.hsx.vn/Modules/CMS/Web/ArticleInCategory/95cd3266-e6d1-42a3-beb5-20ed010aea4a'
            f_url = f'{web_url}?exclude=00000000-0000-0000-0000-000000000000&lim=True&pageFieldName1=FromDate&' \
                    f'pageFieldValue1={pageFieldValue2[0:2]}.10.2021&pageFieldOperator1=eq&pageFieldName2=ToDate&' \
                    f'pageFieldValue2={pageFieldValue2}&pageFieldOperator2=eq&pageFieldName3=TokenCode&' \
                    f'pageFieldValue3=&pageFieldOperator3=eq&pageFieldName4=CategoryId&' \
                    f'pageFieldValue4={pageFieldValue4}&pageFieldOperator4=eq&pageCriteriaLength=4&_search=false&' \
                    f'nd=1650593946103&rows=30&page={page}&sidx=id&sord=desc'

            with requests.Session() as session:
                retry = requests.packages.urllib3.util.retry.Retry(connect=5, backoff_factor=1)
                adapter = requests.adapters.HTTPAdapter(max_retries=retry)
                session.mount('https://', adapter)
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36'
                }
                html = session.get(f_url, headers=headers, timeout=10).content
                soup = BeautifulSoup(html, 'lxml')
                news = json.loads(soup.text)

                def get_content_from_row(n):
                    ID, timeString, rawTitle = news['rows'][n]['cell']
                    title = BeautifulSoup(rawTitle,'html5lib').find('a', href=True).text
                    articleTime = f(timeString)
                    return ID, articleTime, title

                get_cell = [d.get('cell') for d in news['rows']]
                for row in range(len(get_cell)):
                    sub_pdfs = []
                    id_url, time_, sub_title = get_content_from_row(row)
                    sub_url = f'https://www.hsx.vn/Modules/Cms/Web/ViewArticle/{id_url}'

                    nxt_rq = f'https://www.hsx.vn/Modules/Cms/Web/LoadArticle?id={id_url}&objectType=1'
                    html = session.get(nxt_rq, headers=headers, timeout=10).content
                    soup = BeautifulSoup(html, 'lxml')
                    get_content = soup.find_all('p')
                    paragraph = '\n'.join([con.text.replace('\xa0', ' ') for con in get_content])
                    pdf_elements = soup.find_all(href=True)
                    pdf_files = ['https://www.hsx.vn' + pdf.get('href') for pdf in pdf_elements]

                    check = [word in sub_title for word in keywords]
                    if any(check):
                        sub_pdfs.append(f1(pdf_files))
                        print(time_, sub_title, paragraph, sub_pdfs, sub_url)
                        # report and merge to parent table
                        frame = pd.DataFrame({
                            'Thời gian': [time_],
                            'Tiêu đề': [sub_title],
                            'Nội dung': paragraph,
                            'File đính kèm': sub_pdfs,
                            'Link': f'=HYPERLINK("{sub_url}","Link")'
                        })
                        frames.append(frame)
                    format_time = time_
                if format_time < dt.datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):
                    break
            page += 1
        if len(frames) == 0:
            output_table = pd.DataFrame()
        output_table = pd.concat(frames, ignore_index=True)
        output_table = output_table.drop_duplicates(subset=['Thời gian', 'Tiêu đề'])

        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return output_table


vsd = vsd()
hnx = hnx()
hose = hose()
