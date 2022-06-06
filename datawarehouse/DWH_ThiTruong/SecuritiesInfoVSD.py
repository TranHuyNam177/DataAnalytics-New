from request import *
from datawarehouse import BATCHINSERT,DELETE


def __cleanText__(x):
    for char in """~`!@#$%^&*()_+=[]\;',./{}|:"<>?""": # không tính dấu -
        x = x.replace(char,'')
    return x.replace('\n','').strip()

def run(idStart=1,idEnd=20000):

    """
    This function updates data to table [DWH-ThiTruong].[dbo].[SecuritiesInFoVSD]
    """

    for lang in ['en','vi']:
        with requests.Session() as session:
            retry = requests.packages.urllib3.util.retry.Retry(connect=5,backoff_factor=1)
            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
            session.mount('https://',adapter)
            headers = {
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
            }
            for id_ in range(idStart,idEnd+1):
                URL = f'https://www.vsd.vn/{lang}/s-detail/{id_}'
                print(URL)
                res = session.get(URL,headers=headers,timeout=30)
                soup = BeautifulSoup(res.content,'html5lib')
                ticker = __cleanText__(soup.find(class_='title-category').text).split()[0]
                if len(ticker) < 3: # id không được gán vào chứng khoán nào cả
                    continue
                table = soup.find(class_='news-issuers')
                rows = table.find_all(class_='row')
                for row in rows[:-1]: # bỏ dòng ghi chú cuối cùng
                    attr, value = [__cleanText__(tag.text) for tag in row.find_all('div')]
                    deleteCondition = f"WHERE [Ticker] = '{ticker}' AND [Language] = '{lang}' AND [Attribute] = '{attr}'"
                    DELETE(connect_DWH_ThiTruong,'SecuritiesInFoVSD',deleteCondition)
                    resultRecord = [(ticker,id_,lang.upper(),attr,value)]
                    resultFrame = pd.DataFrame(resultRecord,columns=['Ticker','VSDID','Language','Attribute','Value'])
                    BATCHINSERT(connect_DWH_ThiTruong,'SecuritiesInFoVSD',resultFrame)

