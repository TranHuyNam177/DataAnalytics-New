import re
import itertools
import datetime as dt
import time
import pandas as pd
from datawarehouse import connect_DWH_Base, DELETE, SEQUENTIALINSERT
from os.path import join, dirname, realpath

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

def run():

    PATH = join(dirname(dirname(dirname(realpath(__file__)))),'dependency','chromedriver')
    options = Options()
    options.headless = False
    driver = webdriver.Chrome(options=options,executable_path=PATH)
    driver.get('https://account.base.vn/company')
    driver.maximize_window()
    ignored_exceptions = (
        ValueError,
        IndexError,
        NoSuchElementException,
        StaleElementReferenceException,
        ElementNotInteractableException,
    )
    wait = WebDriverWait(driver,10,ignored_exceptions=ignored_exceptions)

    with open(r"C:\Users\hiepdang\Desktop\Passwords\Base\hiepdang.txt") as file:
        email, password = file.readlines()
        email = email.replace('\n','')
        password = password.replace('\n','')

    # Input Email
    emailElement = wait.until(EC.presence_of_element_located((By.NAME,'email')))
    emailElement.clear()
    emailElement.send_keys(email)
    # Input Password
    passwordElement = wait.until(EC.presence_of_element_located((By.NAME,'password')))
    passwordElement.clear()
    passwordElement.send_keys(password)
    # Click "Login to start working"
    wait.until(EC.presence_of_element_located((By.CLASS_NAME,'submit'))).click()
    # Click "Tiếp tục" trên pop-up (nếu có)
    time.sleep(5) # chờ load xong
    popUps = driver.find_elements(By.CLASS_NAME,'ok')
    if popUps:
        popUps[0].click()
    # Chọn tab "Thành viên"
    xpath = '//*[@class="-ap icon-users"]'
    wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
    driver.switch_to.default_content()
    # Column "Họ và Tên"
    time.sleep(5) # chờ load xong
    xpath = '//*[contains(@data-url,"user/")]'
    nameElements = wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    names = [elem.text.title() for elem in nameElements]
    xpath = '//*[contains(@class,"sub ap-xdot")]'
    infoElements = wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    infos = [elem.text.strip() for elem in infoElements if len(elem.text) > 10]
    users = [re.search('@\w{0,}',i).group() for i in infos if len(i) > 10]
    positions = []
    for i in infos:
        splitList = i.split('·')[1].split('-')
        if len(splitList) > 1:
            positions.append(splitList[0].strip())
        else:
            positions.append(None)
    branchs = []
    for i in infos:
        splitList = i.split('-')
        if len(splitList) > 1:
            branchs.append(splitList[1])
        else:
            branchs.append(None)
    # Column "Thông tin liên lạc"
    xpath = '//*[@class="minfo list-info"]/*[@class="li ap-xdot"]'
    infoElements = wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    infos = (elem.text.strip() for elem in infoElements) # iterator
    emails = []
    phoneNumbers = []
    birthDates = []
    for d1,d2,d3 in zip(infos,infos,infos):
        emails.append(d1)
        pattern = '0\d{1,}'
        phoneNumberMatch = re.search(pattern,d2)
        if not phoneNumberMatch:
            phoneNumbers.append(None)
        else:
            phoneNumbers.append(phoneNumberMatch.group())
        if ('/' not in d3) or ('--' in d3):
            birthDates.append(None)
        else:
            birthDates.append(dt.datetime.strptime(d3,'%d/%m/%Y'))
    # Column "Quản lý trực tiếp"
    xpath = '//*[@id="people-list"]/tbody/tr[*]/td[3]'
    infoElements = wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    infos = [elem.text for elem in infoElements]
    managerUsers = []
    for i in infos:
        if i:
            managerUsers.append(i.split('·')[0].split('\n')[1].strip())
        else:
            managerUsers.append(None)
    xpath = '//*[@class="ms"]'
    infoElements = wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
    joinDates = [dt.datetime.strptime(elem.text.strip(),'%d/%m/%y') for elem in infoElements]
    # Xử lý bằng Pandas
    table = pd.DataFrame({
        'Name': names,
        'User': users,
        'Position': positions,
        'BranchName': branchs,
        'Email': emails,
        'PhoneNumber': phoneNumbers,
        'BirthDate': birthDates,
        'ManagerUser': managerUsers,
        'JoinDate': joinDates,
    })
    DELETE(connect_DWH_Base,'Employee','')
    SEQUENTIALINSERT(connect_DWH_Base,'Employee',table)
