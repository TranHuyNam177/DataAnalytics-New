from os.path import join, dirname, realpath
import time
import pandas as pd


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
        self.fromDate = fromDate
        self.toDate = toDate

    def getNewsThucHienQuyen(self) -> pd.DataFrame:
        """
        This function returns a DataFrame, get data from 'https://vsd.vn/vi/lich-giao-dich?tab=LICH_THQ&date=10/2022'
        and add data to database SQL server
        (lịch thực hiện quyền).
        :return: dataframe
        """

        start_time = time.time()
        # Create function to clear input box and send dates as string

        def sendDate(element, d):
            element.click()  # Focus input field
            time.sleep(0.5)
            element.send_keys(Keys.CONTROL, "a")  # Select all pre-existing text/input value
            element.send_keys(Keys.BACKSPACE)  # Remove that text
            element.send_keys(d.strftime('%d%m%Y'))

        url = 'https://vsd.vn/vi/lich-giao-dich?tab=LICH_THQ&date=10/2022'
        driver = webdriver.Chrome(service=Service(self.__PATH), options=Options())
        driver.get(url)
        driver.maximize_window()
        wait = WebDriverWait(driver, 10, ignored_exceptions=self.__ignored_exceptions)
        # Nhập 2 thông tin Từ ngày và Đến ngày
        id = 'txtSearchLichTHQ_TuNgay'
        fromDateInput = wait.until(EC.presence_of_element_located((By.ID, id)))
        sendDate(fromDateInput, self.fromDate)
        time.sleep(0.5)  # khoảng nghỉ giữa 2 ngày
        id = 'txtSearchLichTHQ_DenNgay'
        toDateInput = wait.until(EC.presence_of_element_located((By.ID, id)))
        sendDate(toDateInput, self.toDate)
        # click search
        xpath = "//*[contains(@onclick, 'btnSearchLichTHQ();')]"
        searchButton = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        searchButton.click()
        time.sleep(2)  # đợi 2s để load ra data
        # Kiểm tra xem có data hay không
        xpath = "//*[contains(@id, 'd_total_rec')]"
        # trả ra chuỗi Hiển thị: ... - ... / ... bản ghi
        checkNumberRecords = wait.until(EC.presence_of_element_located((By.XPATH, xpath))).text
        if '0 - 0 / 0' in checkNumberRecords:
            raise NoNewsFound(f'Không có tin từ {self.fromDate.strftime("%d/%m/%Y")} đến {self.toDate.strftime("%d/%m/%Y")}')

        numberRecords = int(checkNumberRecords.split('/')[-1].replace('bản ghi', ''))

        # Tìm đúng vị trí của từng cột trong tiêu đề
        xpath = '//*[@id="tblLichTHQ"]/thead/tr/th'
        headers = wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))[1:]
        headerDict = dict()
        for row_num, header in enumerate(headers):
            headerDict[header.text] = row_num + 2

        table = pd.DataFrame()
        while True:
            # lấy data
            xpath = "//*[contains(@id, 'tblLichTHQ')]/tbody"
            tableElement = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            colMapper = {k: v for k, v in zip(headerDict.keys(), [[]] * len(headerDict))}
            for key in headerDict.keys():
                xpath = f'tr[*]/td[{headerDict[key]}]'
                colElements = tableElement.find_elements(By.XPATH, xpath)
                colStrings = [col.text for col in colElements]
                colMapper[key] = colMapper[key] + colStrings
            # Get URL
            xpath = f'tr[*]/*/a'
            linkElements = tableElement.find_elements(By.XPATH, xpath)
            colMapper['URL'] = [link.get_attribute('href') for link in linkElements]
            # save data to dataframe
            table = pd.concat([table, pd.DataFrame(colMapper)], ignore_index=True)
            if table.shape[0] == numberRecords:
                break
            # turn page
            driver.execute_script(f'window.scrollTo(0,700)')
            xpath = "//*[contains(@id, 'next')]"
            wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()
            time.sleep(1)

        driver.quit()

        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')
        return table
