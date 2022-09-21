import pandas as pd
import re
import requests
from bs4 import BeautifulSoup
from datawarehouse import BATCHINSERT, DELETE
from request import connect_DWH_ThiTruong


def __cleanText__(x):
    for char in """~`!@#$%^&*()+=[]\:;'{}|"<>?""": # không tính dấu - và _
        x = x.replace(char,'')
    return x.replace('\n','').strip()

def run(nrequest,startID=None):

    """
    This function updates data to table [DWH-ThiTruong].[dbo].[TinToChucPhatHanhVSD]
    """

    def _getLastID():
        sql = """
            SELECT MAX([VSDID]) FROM [TinToChucPhatHanhVSD]
        """
        lastID = pd.read_sql(sql,connect_DWH_ThiTruong).squeeze()
        if lastID is None:
            lastID = 0
        return lastID

    def _insert(ticker,id_,attr,value):
        deleteCondition = f"WHERE [Ticker] = '{ticker}' AND [Language] = 'VI' AND [Attribute] = N'{attr}'"
        DELETE(connect_DWH_ThiTruong,'TinToChucPhatHanhVSD',deleteCondition)
        resultRecord = [(ticker,id_,'VI',attr,value)]
        resultFrame = pd.DataFrame(resultRecord,columns=['Ticker','VSDID','Language','Attribute','Value'])
        BATCHINSERT(connect_DWH_ThiTruong,'TinToChucPhatHanhVSD',resultFrame)

    with requests.Session() as session:
        retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
        adapter = requests.adapters.HTTPAdapter(max_retries=retry)
        session.mount('https://',adapter)
        headers = {
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
        }
        if startID is None:
            startID = _getLastID()
        for id_ in range(startID,startID+nrequest):
            URL = f'https://www.vsd.vn/vi/ad/{id_}'
            print(URL)
            res = session.get(URL,headers=headers,timeout=30)
            soup = BeautifulSoup(res.content,'html5lib')
            tickerTag = soup.find(class_='title-category')
            if tickerTag is None or tickerTag.text == '': # Bài viết không có hoặc đã bị gỡ bỏ!
                continue
            ticker = __cleanText__(tickerTag.text).split()[0]
            if re.search(r'\b[A-Z\d]{3,}\b',ticker) is None: # tin đăng bị lỗi
                continue
            if len(ticker) < 3: # id không được gán vào chứng khoán nào cả
                continue
            # Lấy raw content
            content = soup.find(class_='col-md-12')
            if content is None: # Bài viết không có hoặc đã bị gỡ bỏ!
                continue
            # Lấy table
            tableRows = content.find_all(class_='row')
            for tableRow in tableRows:
                recordTags = tableRow.find_all('div')
                if len(recordTags) < 2: # dòng bị gộp -> chứa thông tin ko liên quan
                    continue
                attr, tableValue = [__cleanText__(tag.text) for tag in recordTags]
                _insert(ticker,id_,attr,tableValue)
            # Lấy text
            paraTexts = []
            for tagName in ('b','p'):
                paraTexts += content.find_all(name=tagName)
            textValue = ''.join([paraText.text.strip() for paraText in paraTexts])
            _insert(ticker,id_,'Văn bản',textValue)