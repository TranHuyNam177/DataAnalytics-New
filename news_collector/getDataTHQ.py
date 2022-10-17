from bs4 import BeautifulSoup
import requests
import re

class thucHienQuyen:
    def __init__(self, url):
        self.url = url
    def readContentInLink(self, link):
        with requests.Session() as session:
            retry = requests.packages.urllib3.util.retry.Retry(connect=5, backoff_factor=1)
            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
            session.mount('https://', adapter)
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
            }
            html = session.get(link, headers=headers, timeout=10).text
            soup = BeautifulSoup(html, 'html5lib')
            news = soup.find(class_='container-small')
            titleNews = news.find(class_="title-category").get_text()
            timeNews = news.find(class_="time-newstcph").get_text()
            contents = news.find(class_='col-md-12').get_text()
            contents = re.sub(r'\s{2,}', ' ', contents).replace('\n', ' ')