import time

from automation import *
import cv2

dept_folder = r'C:\Users\namtran\Share Folder\Finance\Report'

def get_info(
    periodicity:str,
    run_time=None,
):
    if run_time=='now' or run_time is None:
        run_time = dt.datetime.now()

    run_year = run_time.year
    run_month = run_time.month

    """
    Mon >> run_weekday = 2
    Tue >> run_weekday = 3
    Wed >> run_weekday = 4
    Thu >> run_weekday = 5
    Fri >> run_weekday = 6
    Sat >> run_weekday = 7
    Sun >> run_weekday = 8
    """
    run_weekday = run_time.weekday()+2

    # Calculate time for quarterly report
    if run_month in [1,2,3]:
        soq = dt.datetime(run_year-1,10,1)
        eoq = dt.datetime(run_year-1,12,31)
    elif run_month in [4,5,6]:
        soq = dt.datetime(run_year,1,1)
        eoq = dt.datetime(run_year,3,31)
    elif run_month in [7,8,9]:
        soq = dt.datetime(run_year,4,1)
        eoq = dt.datetime(run_year,6,30)
    else:
        soq = dt.datetime(run_year,7,1)
        eoq = dt.datetime(run_year,9,30)

    # Calculate time for monthly report
    if run_month==1:
        mreport_year = run_year-1
        mreport_month = 12
    else:
        mreport_year = run_year
        mreport_month = run_month-1
    som = dt.datetime(mreport_year,mreport_month,1)
    eom = dt.datetime(run_year,run_month,1)-dt.timedelta(days=1)

    # Calculate time for weekly report
    if run_weekday in [2,3,4,5]:
        sow = run_time-dt.timedelta(days=run_weekday+5)
        eow = run_time-dt.timedelta(days=run_weekday+1)
    else:  # run_weekday in [6,7,8]
        sow = run_time-dt.timedelta(days=run_weekday-2)
        eow = run_time-dt.timedelta(days=run_weekday-6)

    # select name of the folder
    folder_mapper = {
        'daily':'BaoCaoNgay',
        'weekly':'BaoCaoTuan',
        'monthly':'BaoCaoThang',
        'quarterly':'BaoCaoQuy',
    }
    folder_name = folder_mapper[periodicity]

    # choose dates and period
    if periodicity.lower()=='daily':
        start_date = run_time.strftime('%Y.%m.%d')
        end_date = start_date
        period = run_time.strftime('%Y.%m.%d')
    elif periodicity.lower()=='weekly':
        start_date = sow.strftime('%Y.%m.%d')
        end_date = eow.strftime('%Y.%m.%d')
        start_str = convert_int(sow.day)
        end_str = f'{convert_int(eow.day)}.{convert_int(eow.month)}.{eow.year}'
        period = f'{start_str}-{end_str}'
    elif periodicity.lower()=='monthly':
        start_date = som.strftime('%Y.%m.%d')
        end_date = eom.strftime('%Y.%m.%d')
        period = f'{convert_int(eom.month)}.{eom.year}'
    elif periodicity.lower()=='quarterly':
        start_date = soq.strftime('%Y.%m.%d')
        end_date = eoq.strftime('%Y.%m.%d')
        quarter_mapper = {
            3:'Q1',
            6:'Q2',
            9:'Q3',
            12:'Q4',
        }
        quarter = quarter_mapper[eoq.month]
        period = f'{quarter}.{eoq.year}'
    else:
        raise ValueError('Invalid periodicity')

    result_as_dict = {
        'run_time':run_time,
        'start_date':start_date,
        'end_date':end_date,
        'period':period,
        'folder_name':folder_name,
    }

    return result_as_dict


def get_bank_authentication(
    bank:str
) -> dict:

    """
    This function returns user and password to login to internet banking of the company

    :param bank: Name of target bank. Accept either: {'BIDV','VTB','IVB'}
    """

    resultDict = dict()
    with open(fr'C:\Users\namtran\Desktop\Passwords\Bank\{bank}.txt') as file:
        if bank in ['BIDV','IVB','VCB','VTB','EIB','OCB','TCB']:
            resultDict['id'] = ''
            resultDict['user'], resultDict['password'], resultDict['URL'] = file.readlines()
        elif bank in ['FUBON','SCSB','FIRST','MEGA','SINOPAC','ESUN','HUANAN']:
            resultDict['id'], resultDict['user'], resultDict['password'], resultDict['URL'] = file.readlines()
        else:
            raise ValueError(f'Invalid bank name: {bank}')
        for key,value in resultDict.items():
            resultDict[key] = value.replace('\n','')

    return resultDict


class Base:

    def __init__(self,bank:str,debug:bool=True):

        """
        Base class for all Bank classes
        :param bank: bank name
        """

        self.bank = bank
        self.debug = debug
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException,
            ElementClickInterceptedException,
        )
        self.PATH = join(dirname(dirname(dirname(realpath(__file__)))),'dependency','chromedriver')
        infoDict = get_bank_authentication(self.bank)
        self.id = infoDict['id']
        self.user = infoDict['user']
        self.password = infoDict['password']
        self.URL = infoDict['URL']
        self.downloadFolder = r'C:\Users\namtran\Downloads'

        self.cBankAccounts = {
            'EIB': ['140114851002285','160314851020212'],
            'OCB': ['0021100002115004'],
            'BIDV': ['11910000132943','26110002677688'],
            'VCB': ['0071001264078'],
            'VTB': ['147001536591'],
            'TCB': ['19038382442022'],
            'IVB': [],
        }

    def __repr__(self):
        return '<BaseObjectOfBank>'

    def __GetCaptchaFromMail__(self,captcha_element):

        """
        This function get CAPTCHA image and send it to mail for user input
        :param captcha_element: Selenium's Web Driver of CAPTCHA image
        """

        if self.bank == 'IVB':
            regexPattern = '^[0-9a-z]{5}$'
        elif self.bank == 'BIDV':
            regexPattern = '^[0-9A-Za-z]{6}$'
        elif self.bank == 'VCB':
            regexPattern = '^[0-9A-Z]{6}$'
        elif self.bank == 'SCSB':
            regexPattern = '^[0-9]{5}$'
        elif self.bank == 'FIRST':
            regexPattern = '^[0-9]{5}$'
        elif self.bank == 'SINOPAC':
            regexPattern = '^[0-9A-Z]{5}$'
        else:
            raise ValueError('Invalid Bank Name')

        outlook = Dispatch('outlook.application')
        # Dọn dẹp folder mail
        mapi = outlook.GetNamespace("MAPI")
        inbox = mapi.Folders.Item(1).Folders['Inbox']
        email_ids = []
        for i in range(len(inbox.Items)):
            if f're: captcha required: {self.bank.lower()}' in inbox.Items[i].Subject.lower():
                email_ids.append(inbox.Items[i].EntryID)
        folder_id = inbox.StoreID
        for email_id in email_ids:
            mapi.GetItemFromID(email_id,folder_id).Delete()

        # Lấy CAPTCHA từ IB
        imgPATH = join(dirname(__file__),'CAPTCHA',f'{self.bank}.png')
        captcha_element.screenshot(imgPATH)
        # Decode image sang UTF-8
        base64Img = base64.b64encode(open(imgPATH,'rb').read()).decode('utf-8')
        # Gửi CAPTCHA qua bank
        mail = outlook.CreateItem(0)
        if self.debug: # Development Run
            mail.To = 'namtran@phs.vn'
        else: # Production Run
            # mail.To = 'namtran@phs.vn; duynguyen@phs.vn; thaonguyenthanh@phs.vn'
            mail.To = 'namtran@phs.vn'
        mail.Subject = f"CAPTCHA Required: {self.bank}"
        mail.attachments.Add(imgPATH)
        html = f"""
        <html>
            <head></head>
            <body>
                <img alt="" src="data:image/png;base64,{base64Img}"/><br>
                <p style="font-family:Times New Roman; font-size:100%"><i>
                    Machine can't read CAPTCHA <br>
                    Reply this email with exact result to make the application continue
                </i></p>
                <p style="font-family:Times New Roman; font-size:90%"><i>
                    -- Generated by Reporting System
                </i></p>
            </body>
        </html>
        """
        mail.HTMLBody = html
        mail.Send()
        # Chờ phản hồi để nhận CAPTCHA
        while True:
            inbox = mapi.Folders.Item(1).Folders['Inbox']
            messages = inbox.Items
            for message in messages:
                if f're: captcha required: {self.bank.lower()}' in message.Subject.lower():
                    subBody = message.Body.split()[0]
                    regex = re.compile(regexPattern)
                    match = regex.search(subBody)
                    if match:
                        return match.group()
            time.sleep(5)  # 5s quét 1 lần


class BIDV(Base):

    def __init__(self,debug=True):
        super().__init__('BIDV',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_BIDV>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        def GO(CAPTCHA):
            """
            Procedure login với CAPTCHA cho trước
            """
            # Input CAPTCHA
            captchaInput = self.driver.find_element(By.ID,'captcha')
            captchaInput.clear()
            captchaInput.send_keys(CAPTCHA)
            # Input Username
            userInput = self.driver.find_element(By.ID,'username')
            userInput.clear()
            userInput.send_keys(self.user)
            # Input Password
            userInput = self.driver.find_element(By.ID,'password')
            userInput.clear()
            userInput.send_keys(self.password)
            # Click "Đăng nhập"
            loginButton = self.driver.find_element(By.ID,'btLogin')
            loginButton.click()

        # CAPTCHA
        while True:
            captchaElement = self.wait.until(EC.presence_of_element_located((By.ID,'idImageCap')))
            imgPATH = join(realpath(dirname(__file__)),'CAPTCHA',f'{self.bank}.png')
            captchaElement.screenshot(imgPATH) # download CAPTCHA về dưới dạng .png
            image = cv2.imread(imgPATH)
            crop = image[5:35,5:90,:]
            cv2.imwrite(imgPATH,crop)
            predictedCAPTCHA = pytesseract.image_to_string(imgPATH).replace('\n','').replace(' ','')
            print(predictedCAPTCHA)
            condition1 = len(predictedCAPTCHA) == 6
            condition2 = not re.findall('[^a-zA-Z0-9]',predictedCAPTCHA)
            condition3 = not re.findall('[LSWYiklsvyu]',predictedCAPTCHA)
            if condition1 and condition2 and condition3: # Cases do not need refresh
                break
            self.driver.find_element(By.CLASS_NAME,'btnRefresh').click() # Cases need refresh
            time.sleep(0.5)

        GO(predictedCAPTCHA)

        # Check xem CAPTCHA đúng chưa, nếu chưa đúng -> gửi mail đọc CAPTCHA bằng tay
        errorMessages = self.driver.find_elements(By.CLASS_NAME,'errorMessage')
        if errorMessages:  # Nếu chưa đúng
            captchaElement = self.wait.until(EC.presence_of_element_located((By.ID,'idImageCap')))
            manualCAPTCHA = self.__GetCaptchaFromMail__(captchaElement)
            GO(manualCAPTCHA)

        return self

class VTB(Base):

    def __init__(self,debug=True):
        super().__init__('VTB',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_VTB>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Username
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@placeholder="Tên đăng nhập"]')))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@placeholder="Mật khẩu"]')))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click đăng nhập
        loginButton = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@type="submit"]')))
        loginButton.click()
        time.sleep(1)

        return self

class IVB(Base):

    def __init__(self,debug=True):
        super().__init__('IVB',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_IVB>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Username
        xpath = '//*[@placeholder="Tên truy cập"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        xpath = '//*[@placeholder="Mật khẩu"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # CAPTCHA
        captcha_element = self.wait.until(EC.presence_of_element_located((By.ID,'safecode')))
        CAPTCHA = self.__GetCaptchaFromMail__(captcha_element)
        xpath = '//*[@placeholder="Mã xác thực"]'
        captchaInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        captchaInput.clear()
        captchaInput.send_keys(CAPTCHA)
        # Click Đăng nhập
        self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@onclick="logon()"]'))).click()
        # Nếu xuất hiện màn hình xác nhận
        xpath = '//*[@onclick="forceSubmit()"]'
        possibleButtons = self.driver.find_elements(By.XPATH,xpath)
        if possibleButtons:
            possibleButtons[0].click()

        return self

class VCB(Base):

    def __init__(self,debug=True):
        super().__init__('VCB',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_VCB>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        def GO(CAPTCHA):
            """
            Procedure login với CAPTCHA cho trước
            """
            # Nhập CAPTCHA
            xpath = '//*[@placeholder="Nhập số bên"]'
            captchaInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            captchaInput.clear()
            captchaInput.send_keys(CAPTCHA)
            # Username
            xpath = '//*[@placeholder="Tên truy cập"]'
            userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            userInput.clear()
            userInput.send_keys(self.user)
            # Password
            xpath = '//*[@placeholder="Mật khẩu"]'
            passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            passwordInput.clear()
            passwordInput.send_keys(self.password)
            # Click 'Đăng nhập'
            xpath = '//*[@value="Đăng nhập"]'
            self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        # CAPTCHA
        while True:
            xpath = '//*[@alt="Chuỗi bảo mật"]'
            CAPTCHA = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            imgPATH = join(dirname(__file__),'CAPTCHA',f'{self.bank}.png')
            CAPTCHA.screenshot(imgPATH)  # download CAPTCHA về dưới dạng .png
            predictedCAPTCHA = pytesseract.image_to_string(imgPATH).replace('\n','')
            print(predictedCAPTCHA)
            condition1 = len(predictedCAPTCHA) == 6
            condition2 = not re.findall('[^0-9A-Z]',predictedCAPTCHA)
            condition3 = not re.findall('[048AGIOZRS]',predictedCAPTCHA)
            if condition1 and condition2 and condition3: # case không cần refresh
                break
            self.driver.refresh() # case cần refresh

        GO(predictedCAPTCHA)

        # Check xem CAPTCHA đúng chưa, nếu chưa đúng -> gửi mail đọc CAPTCHA bằng tay
        errorMessages = self.driver.find_elements(By.ID,'ctl00_Content_Login_CaptchaValidator')
        if errorMessages: # Nếu chưa đúng
            xpath = '//*[@alt="Chuỗi bảo mật"]'
            captchaElement = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            imgPATH = join(dirname(__file__),'CAPTCHA',f'{self.bank}.png')
            captchaElement.screenshot(imgPATH)  # download CAPTCHA về dưới dạng .png
            manualCAPTCHA = self.__GetCaptchaFromMail__(captchaElement)
            GO(manualCAPTCHA)

        return self

class EIB(Base):

    def __init__(self,debug=True):
        super().__init__('EIB',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_EIB>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Tên đâng nhập
        xpath = '//*[@placeholder="Tên đăng nhập"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Mật khẩu
        xpath = '//*[@placeholder="Mật khẩu"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click "Đăng nhập"
        self.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'btn-primary'))).click()
        # Tắt popup yêu cầu đăng ký SMS (nếu có)
        time.sleep(5) # chờ đăng nhập xong
        xpath = '//*[contains(text(),"Thông báo")]'
        noticeWindow = self.driver.find_elements(By.XPATH,xpath)
        if noticeWindow:  # có popup
            xpath = '//*[contains(text(),"Đồng ý")]'
            Buttons = self.driver.find_elements(By.XPATH,xpath)
            if Buttons:
                Buttons[0].click()
            xpath = '//*[contains(text(),"Đóng")]'
            Buttons = self.driver.find_elements(By.XPATH,xpath)
            if Buttons:
                Buttons[0].click()
        return self

class OCB(Base):

    def __init__(self,debug=True):
        super().__init__('OCB',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_OCB>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Trong trường hợp có pop-up window
        exitPopup = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'x-button-popup-login')))
        if exitPopup.is_displayed():  # có pop-up window
            exitPopup.click()
        # Tên đâng nhập
        xpath = '//*[@placeholder="Tên đăng nhập"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Click "Tiếp tục"
        self.wait.until(EC.presence_of_element_located((By.ID,'loginSubmitButton'))).click()
        # Mật khẩu
        xpath = '//*[@placeholder="Mật khẩu"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click "Đăng nhập"
        self.wait.until(EC.presence_of_element_located((By.ID,'loginSubmitButton'))).click()

        return self

class TCB(Base):

    def __init__(self,debug=True):
        super().__init__('TCB',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_TCB>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        # self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Tên đăng nhập
        userInput = self.wait.until(EC.presence_of_element_located((By.NAME,'signOnName')))
        userInput.clear()
        userInput.send_keys(self.user)
        # Mật khẩu
        passwordInput = self.wait.until(EC.presence_of_element_located((By.NAME,'password')))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click "Đăng nhập"
        xpath = '//*[@value="Đăng Nhập"]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        return self

class FUBON(Base):

    def __init__(self,debug=True):
        super().__init__('FUBON',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_FUBON>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)
        # Switch frame
        self.driver.switch_to.frame(0)
        # Customer Code
        xpath = '//*[contains(text(),"Customer Code")]/../*/input'
        idInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        idInput.clear()
        idInput.send_keys(self.id)
        # User Code
        xpath = '//*[contains(text(),"User Code")]/../*/input'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        xpath = '//*[contains(text(),"Password")]/../*/input'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click "Submit"
        xpath = '//*[@value="Submit"]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        return self

class SCSB(Base):

    def __init__(self,debug=True):
        super().__init__('SCSB',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_SCSB>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Switch frame
        self.driver.switch_to.frame(1)

        def firstGO(CAPTCHA):
            """
            Procedure login màn hình đầu tiên với CAPTCHA cho trước
            """
            # Nhập CAPTCHA
            xpath = '//*[@id="cpatchaTextBox"]'
            captchaInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            captchaInput.clear()
            captchaInput.send_keys(CAPTCHA)
            # User ID
            xpath = '//*[@id="userID"]'
            idInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            idInput.clear()
            idInput.send_keys(self.id)
            # Username
            xpath = '//*[@name="loginUID"]'
            userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            userInput.clear()
            userInput.send_keys(self.user)
            # Click Đăng nhập
            xpath = '//*[@value="Submit"]'
            self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        def secondGO():
            """
            Procedure login màn hình thứ hai với CAPTCHA cho trước
            """
            # Điền password
            xpath = '//*[@name="password"]'
            passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            passwordInput.clear()
            passwordInput.send_keys(self.password)
            # Click đăng nhập
            self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@value="Submit"]'))).click()

        # CAPTCHA
        while True:
            xpath = '//*[@id="captcha"]'
            CAPTCHA = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            imgPATH = join(dirname(__file__),'CAPTCHA',f'{self.bank}.png')
            CAPTCHA.screenshot(imgPATH)  # download CAPTCHA về dưới dạng .png
            predictedCAPTCHA = re.sub('[^0-9]','',pytesseract.image_to_string(imgPATH))
            print(predictedCAPTCHA)
            condition = len(predictedCAPTCHA) == 5
            if condition: # case không cần refresh
                break
            self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@name="Image40"]'))).click() # case cần refresh

        firstGO(predictedCAPTCHA)

        # Check xem CAPTCHA đúng chưa, nếu chưa đúng -> gửi mail đọc CAPTCHA bằng tay
        try:
            self.driver.switch_to.alert.accept()
            # Màn hình login đầu tiên thất bại -> gửi mail, đọc CAPTCHA từ mail
            xpath = '//*[@id="captcha"]'
            captchaElement = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            imgPATH = join(dirname(__file__),'CAPTCHA',f'{self.bank}.png')
            captchaElement.screenshot(imgPATH)  # download CAPTCHA về dưới dạng .png
            manualCAPTCHA = self.__GetCaptchaFromMail__(captchaElement)
            firstGO(manualCAPTCHA)
            secondGO()
        except (NoAlertPresentException,): # Màn hình login đầu tiên thành công
            secondGO()

        return self

class FIRST(Base):

    def __init__(self,debug=True):
        super().__init__('FIRST',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_FIRST>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        # self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Chờ để các frame xuất hiện
        time.sleep(3)
        # Chuyển frame
        self.driver.switch_to.frame("iFrameID")
        # Click "Overseas Branch User"
        xpath = '//*[@class="login"]/*[@id="tabs"]/li[2]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Chọn chi nhánh Ho Chi Minh
        xpath = '//*[@id="form:branchCombo_label"]'
        dropDownList = self.driver.find_element(By.XPATH,xpath)
        dropDownList.click()
        xpath = '//*[@data-label="Ho Chi Minh City Branch - 920"]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        def GO(CAPTCHA):
            # ID
            xpath = '//*[@id="form:custId"]'
            idInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            idInput.clear()
            idInput.send_keys(self.id)
            # Usernam
            xpath = '//*[@id="form:userUuid"]'
            userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            userInput.clear()
            userInput.send_keys(self.user)
            # Password
            xpath = '//*[@id="form:pwd"]'
            passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            passwordInput.clear()
            passwordInput.send_keys(self.password)
            # CAPTCHA
            xpath= '//*[@id="form:verifyCode"]'
            captchaInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            captchaInput.clear()
            captchaInput.send_keys(CAPTCHA)
            # Click login
            xpath = '//*[@id="form:submitLoginBtn"]'
            self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        # CAPTCHA
        while True:
            xpath = '//*[@id="captchaImg1"]'
            captchaElement = self.wait.until(EC.visibility_of_element_located((By.XPATH,xpath)))
            imgPATH = join(realpath(dirname(__file__)),'CAPTCHA',f'{self.bank}.png')
            captchaElement.screenshot(imgPATH)  # download CAPTCHA về dưới dạng .png
            image = cv2.imread(imgPATH)
            image = image[:,5:100,:] # crop image
            cv2.imwrite(imgPATH,image)
            tor = 0.05 # tolerance level
            correctRGB = (0,88,61)
            blueMatch = (correctRGB[2]*(1-tor)<=image[:,:,0]) & (image[:,:,0]<=correctRGB[2]*(1+tor))
            greenMatch = (correctRGB[1]*(1-tor)<=image[:,:,1]) & (image[:,:,1]<=correctRGB[1]*(1+tor))
            redMatch = (correctRGB[0]*(1-tor)<=image[:,:,2]) & (image[:,:,2]<=correctRGB[0]*(1+tor))
            rgbMatch = redMatch & greenMatch & blueMatch
            imageArray = np.stack([rgbMatch,rgbMatch,rgbMatch],axis=2)
            imageArray = imageArray.astype(np.uint8)
            imageArray *= 255
            outputTable = pytesseract.pytesseract.image_to_data(
                imageArray,
                output_type='data.frame',
                pandas_config={'dtype':{'text':str}}
            )
            outputTable = outputTable.loc[outputTable['conf']!=-1]
            outputSeries = outputTable.loc[outputTable.index[-1],['conf','text']]
            confidenceLevel = outputSeries['conf']
            predictedCAPTCHA = re.sub('[^0-9]','',outputSeries['text'])
            condition1 = confidenceLevel >= 90 # confidence level >= 90%
            condition2 = len(predictedCAPTCHA) == 5
            if condition1 and condition2: # case không cần refresh
                break
            else: # case cần refresh
                xpath = '//*[@class="refresh"]'
                self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        GO(predictedCAPTCHA)

        # Check xem đọc CAPTCHA đúng chưa
        time.sleep(3)
        self.driver.switch_to.frame(0)
        xpath = '//*[contains(text(),"please try again")]'
        errorMessage = self.driver.find_elements(By.XPATH,xpath)
        if errorMessage: # Nếu đọc sai CAPTCHA
            # Đóng pop-up báo lỗi
            xpath = '//*[@id="form:j_idt13"]'
            self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
            while True:
                try:
                    self.driver.switch_to.frame(0)
                    break
                except (NoSuchFrameException,NoSuchWindowException):
                    time.sleep(0.5)
            # Lấy hình ảnh + gửi mail
            xpath = '//*[@id="captchaImg1"]'
            captchaElement = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            manualCAPTCHA = self.__GetCaptchaFromMail__(captchaElement)
            GO(manualCAPTCHA)

        # Đóng popup window báo repeated login nếu có
        time.sleep(2)
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(self.driver.find_element_by_tag_name("iframe"))
        self.driver.switch_to.frame(self.driver.find_element_by_tag_name("iframe"))
        xpath = '//*[text()="Confirm"]'
        confirmButtons = self.driver.find_elements(By.XPATH,xpath)
        if confirmButtons:
            confirmButtons[0].click()

        return self

class MEGA(Base):

    def __init__(self,debug=True):
        super().__init__('MEGA',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_MEGA>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        # self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)
        # Switch frame
        self.driver.switch_to.frame(0)
        # Click "Selected Location Services"
        xpath = '//*[contains(text(),"Selected Location Services")]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Chọn "HoChiMinh City Branch"
        xpath = '//*[@id="form1:branchCode"]'
        dropDownList = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        Select(dropDownList).select_by_visible_text('Ho Chi Minh City Branch')
        # Company ID
        xpath = '//*[@id="form1:custId"]'
        idInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        idInput.clear()
        idInput.send_keys(self.id)
        # Username
        xpath = '//*[@id="form1:loginId"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        xpath = '//*[@id="form1:password"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click "Login"
        xpath = '//*[@id="form1:login3"]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        return self

class SINOPAC(Base):

    def __init__(self,debug=True):
        super().__init__('SINOPAC',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_SINOPAC>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        # self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Switch frame
        self.driver.switch_to.frame(0)
        # e-Banking ID
        xpath = '//*[@id="form:txtCustId"]'
        idInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        idInput.clear()
        idInput.send_keys(self.id)
        # User ID
        xpath = '//*[@id="form:txtUserId"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        xpath = '//*[@id="form:txtUserPwd"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # CAPTCHA
        xpath = '//*[@id="form:captchaImg"]'
        captchaElement = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        CAPTCHA = self.__GetCaptchaFromMail__(captchaElement)
        xpath = '//*[@id="form:captchaInput"]'
        captchaInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        captchaInput.clear()
        captchaInput.send_keys(CAPTCHA)
        # Click Login
        xpath = '//*[@id="form:submitLoginBtn"]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        return self

class ESUN(Base):

    def __init__(self,debug=True):
        super().__init__('ESUN',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_ESUN>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(executable_path=self.PATH)
        # self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Switch frame
        self.driver.switch_to.frame('mainFrame')
        # Đổi ngôn ngữ sang tiếng anh
        xpath = '//*[contains(@class,"ui-dropdown")]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        xpath = '//*[contains(span,"English")]'
        englishElement = self.wait.until(EC.visibility_of_any_elements_located((By.XPATH,xpath)))[0]
        englishElement.click()
        # Switch frame (do load lại trang)
        try:
            self.driver.switch_to.frame('mainFrame')
        except (NoSuchFrameException,NoSuchWindowException):
            pass
        # Chọn Vietnam
        xpath = '//*[contains(@class,"ui-dropdown-trigger")]'
        _, countryDropDown = self.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
        countryDropDown.click()
        xpath = '//*[span="Vietnam" and contains(@class,"ui-dropdown-item")]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Customer ID
        xpath = '//*[@placeholder="Customer ID/No"]'
        idInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        idInput.clear()
        idInput.send_keys(self.id)
        # User Name
        xpath = '//*[@placeholder="User Name"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        xpath = '//*[@placeholder="User Password"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click "Login"
        xpath = '//*[contains(button,"Log In")]/button'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        time.sleep(1) # chờ animation (trong TH popup "Repeated Login" window)
        # Check xem có hiện "Repeated Login" không
        xpath = '//*[text()="Repeated Login"]'
        errorMessage = self.driver.find_elements(By.XPATH,xpath)
        if errorMessage: # trường hợp có lỗi
            xpath = '//*[@class="submit_btn"]'
            _, confirmButton = self.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
            confirmButton.click()

        return self

class HUANAN(Base):

    def __init__(self,debug=True):
        super().__init__('HUANAN',debug)
        self.driver = None
        self.wait = None

    def __repr__(self):
        return f'<BankObject_HUANAN>'

    def __del__(self):
        self.driver.quit()
        print('Destructor: Chrome Driver has quit')

    def Login(self):

        self.driver = webdriver.Chrome(service=Service(self.PATH),options=Options())
        # self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,30,ignored_exceptions=self.ignored_exceptions)

        # Đổi ngôn ngữ sang tiếng anh
        xpath = "//*[@class='language']//*[contains(text(), 'English')]"
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        time.sleep(1)
        # Customer ID
        idInput = self.wait.until(EC.element_to_be_clickable((By.ID,'USERID')))
        idInput.click()
        idInput.send_keys(self.id)
        # User Name
        userInput = self.wait.until(EC.element_to_be_clickable((By.ID,'NICKNAME')))
        userInput.click()
        userInput.send_keys(self.user)
        # Password
        passwordInput = self.wait.until(EC.element_to_be_clickable((By.ID,'pwdText')))
        passwordInput.click()
        self.wait.until(EC.element_to_be_clickable((By.ID,'password'))).send_keys(self.password)
        # Click "Login"
        self.wait.until(EC.presence_of_element_located((By.ID,'WannaLogin'))).click()
        time.sleep(3)  # chờ animation

        return self
