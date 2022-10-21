from os.path import join, dirname, realpath
import time
import pandas as pd
import datetime as dt
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

class NoNewsFound(Exception):
    pass

class vsdTHQ:
    def __init__(self, fromDate, toDate):
        self.__PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
        self.__ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException
        )
        self.__url = 'https://vsd.vn/vi/lich-giao-dich?tab=LICH_THQ&date=10/2022'
        self.__driver = None
        self.__wait = None
        self.fromDate = fromDate
        self.toDate = toDate

    def getNewsThucHienQuyen(self) -> pd.DataFrame:
        """
        This function returns a DataFrame,
        get data from 'https://vsd.vn/vi/lich-giao-dich?tab=LICH_THQ&date=10/2022'
        and add data to database SQL server
        (lịch thực hiện quyền).
        :return: dataframe
        """

        start_time = time.time()

        # Create function to clear input box and send dates as string
        def sendDate(element, d):
            element.click()  # Focus input field
            time.sleep(0.5)
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(Keys.BACKSPACE)
            element.send_keys(d.strftime('%d%m%Y'))

        self.__driver = webdriver.Chrome(service=Service(self.__PATH), options=Options())
        self.__driver.maximize_window()
        self.__driver.get(self.__url)
        self.__wait = WebDriverWait(self.__driver, 10, ignored_exceptions=self.__ignored_exceptions)

        # Nhập 2 thông tin Từ ngày và Đến ngày
        fromDateInput = self.__wait.until(EC.presence_of_element_located((By.ID, 'txtSearchLichTHQ_TuNgay')))
        sendDate(fromDateInput, self.fromDate)
        time.sleep(0.5)  # khoảng nghỉ khi click giữa 2 ngày
        toDateInput = self.__wait.until(EC.presence_of_element_located((By.ID, 'txtSearchLichTHQ_DenNgay')))
        sendDate(toDateInput, self.toDate)
        # click search
        xpath = "//*[contains(@onclick, 'btnSearchLichTHQ();')]"
        searchButton = self.__wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        searchButton.click()
        time.sleep(2)  # đợi 2s để load sau khi nhấn search
        # Kiểm tra xem có data hay không
        xpath = "//*[contains(@id, 'd_total_rec')]"
        checkNumberRecords = self.__wait.until(EC.presence_of_element_located((By.XPATH, xpath))).text  # trả ra chuỗi Hiển thị: ... - ... / ... bản ghi
        if '0 - 0 / 0' in checkNumberRecords:  # không có records
            raise NoNewsFound(f'Không có tin từ {self.fromDate.strftime("%d/%m/%Y")} đến {self.toDate.strftime("%d/%m/%Y")}')

        numberRecords = int(checkNumberRecords.split('/')[-1].replace('bản ghi', ''))

        # Tìm đúng vị trí của từng cột ứng với tiêu đề
        xpath = '//*[@id="tblLichTHQ"]/thead/tr/th'
        headers = self.__wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))[1:]  # bỏ STT
        headerDict = dict()
        for row_num, header in enumerate(headers):
            headerDict[header.text] = row_num + 2  # cộng 2 lên mới ra đúng thứ tự cột

        output_table = pd.DataFrame()
        while True:
            # lấy table chứa data
            xpath = "//*[contains(@id, 'tblLichTHQ')]/tbody"
            tableElement = self.__wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            # tạo dictionary có value là 8 list rỗng tương ứng với 8 keys trong headerDict
            colMapper = {k: v for k, v in zip(headerDict.keys(), [[]] * len(headerDict))}
            for key in headerDict.keys():
                xpath = f'tr[*]/td[{headerDict[key]}]'
                colElements = tableElement.find_elements(By.XPATH, xpath)
                colStrings = [col.text for col in colElements]
                colMapper[key] = colMapper[key] + colStrings
            # Get URLs
            xpath = f'tr[*]/*/a'
            linkElements = tableElement.find_elements(By.XPATH, xpath)
            colMapper['URL'] = [link.get_attribute('href') for link in linkElements]
            # save data to dataframe
            output_table = pd.concat([output_table, pd.DataFrame(colMapper)], ignore_index=True)
            if output_table.shape[0] == numberRecords:
                break
            # Turn page
            self.__driver.execute_script(f'window.scrollTo(0,700)')
            xpath = "//*[contains(@id, 'next')]"
            self.__wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
            time.sleep(1)  # đợi 1s để load dữ liệu sau khi nhấn chuyển trang
        # Rename columns
        output_table.columns = ['NgayDKCC','MaCK','MaISIN','TieuDe','LoaiCK','ThiTruong','NoiQuanLy','URL']
        # Change date in NgayDKCC from str to datetime
        output_table['NgayDKCC'] = pd.to_datetime(output_table['NgayDKCC'], format='%d/%m/%Y')
        # close driver
        self.__driver.quit()
        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return output_table


vsd = vsdTHQ(dt.datetime(2022,10,10), dt.datetime(2022,10,17))
df = vsd.getNewsThucHienQuyen()
