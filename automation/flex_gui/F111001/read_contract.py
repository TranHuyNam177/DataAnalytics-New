import datetime as dt
import os
import pickle
import re
import time
import traceback
import logging
from functools import wraps

import PyPDF2
import cv2 as cv
import numpy as np
import pandas as pd
import pyodbc
import pyperclip
import unidecode
from pdf2image import convert_from_path
from pywinauto.application import Application
from win32com.client import Dispatch

from automation.flex_gui.base import Flex, setFocus
from datawarehouse import SYNC

# DWH-Base Database Information
driver_DWH_CoSo = '{SQL Server}'
server_DWH_CoSo = 'SRV-RPT'
db_DWH_CoSo = 'DWH-CoSo'
connect_DWH_CoSo = pyodbc.connect(
    f'Driver={driver_DWH_CoSo};'
    f'Server={server_DWH_CoSo};'
    f'Database={db_DWH_CoSo};'
    f'uid=hiep;'
    f'pwd=5B7Cv6huj2FcGEM4'
)

SLEEP_TIME = 3 # số giây nghỉ sau mỗi hợp đồng

def _returnEmptyStringIfNoMatch(func):
    @wraps(func)
    def wrapper(*args,**kwargs):
        func.__name__ = ''
        try:
            return func(*args,**kwargs)
        except (AttributeError,): # AttributeError: 'NoneType' object has no attribute 'group'
            return ''
    return wrapper

class NothingToInsert(Exception):
    pass

class Contract:

    PATH = 'W:'
    def __init__(
        self,
        date:dt.datetime=dt.datetime.now()
    ):
        self.date = date
        self.__history = './history.pickle'
        if not os.path.isfile(self.__history): # nếu chưa có file history -> tạo
            with open(self.__history,'wb') as file:
                pickle.dump(set(),file)
        self.__insertedSet = self.readHistory()
        self.__fullSet= set(filter(self.__filterFiles,os.listdir(self.PATH)))
        self.__remainingSet = self.__fullSet.difference(self.__insertedSet)
        if not self.__remainingSet: # rỗng
            raise NothingToInsert('Không có hợp đồng cần phải nhập')

        self.__fileNames = list(self.__remainingSet) # làm việc trên list để preserve order
        self.__contents = self.__pdf2Text()
        self.__images = self.__pdf2Image() # BGR
        self.pointer = 0 # default
        self.selectedFile = self.__fileNames[self.pointer]
        self.selectedContent = self.__contents[self.pointer]
        self.selectedImage = self.__images[self.pointer]

    def __filterFiles(self,file):
        pattern = r'(\d{4})\.(\d{2}-\d{2}-\d{4})\.[A-Za-z]+'
        regexSearch = re.search(pattern,file)
        if regexSearch:
            return regexSearch.group(2)==self.date.strftime('%d-%m-%Y')
        else:
            return False

    def __pdf2Text(self):
        contents = []
        for pdfFileName in self.__fileNames:
            filePDFPath = os.path.join(self.PATH,pdfFileName)
            # Chỉ đọc trang 0 và trang 2, là 2 trang chứa content
            page0RawContent = PyPDF2.PdfReader(filePDFPath).pages[0].extract_text()
            page2RawContent = PyPDF2.PdfReader(filePDFPath).pages[2].extract_text()
            rawContent = page0RawContent + page2RawContent
            rawContent = re.sub(r':\s*\.+',': ',rawContent)
            rawContent = re.sub(r'…','',rawContent)
            contents.append(rawContent)
        return contents

    def __pdf2Image(self):
        arrays = []
        for pdfFileName in self.__fileNames:
            image = convert_from_path(
                pdf_path=os.path.join(self.PATH,pdfFileName),
                poppler_path=r"../../poppler-22.04.0/Library/bin",
                single_file=True,
                size=(1000,None)
            )[0]
            array = cv.cvtColor(np.array(image),cv.COLOR_RGB2BGR)
            arrays.append(array)
        return arrays

    def __findGioiTinh(self):
        # convert to binary image
        grayscaleImage = cv.cvtColor(self.selectedImage,cv.COLOR_BGR2GRAY)
        _, binaryImage = cv.threshold(grayscaleImage,200,255,cv.THRESH_BINARY_INV)
        # tìm 2 ô giới tính (tọa độ cố định vì đã cố định image width = 1000 và aspect ratio ko đổi)
        maleBox = binaryImage[534:547,724:736]
        femaleBox = binaryImage[534:547,808:820]
        maleScore = maleBox.sum()
        femaleScore = femaleBox.sum()
        if maleScore > femaleScore:
            return 'Nam'
        elif maleScore < femaleScore:
            return 'Nữ'
        else:
            return ''

    def __findNgayHetHan(self):
        issueDateString = self.getNgayCap()
        expireDateString = self.getNgaySinh()
        if issueDateString == '' or expireDateString == '':
            return dt.datetime(1900,1,1).strftime('%d/%m/%Y')
        issueDate = dt.datetime.strptime(issueDateString,'%d/%m/%Y')
        birthDate = dt.datetime.strptime(expireDateString,'%d/%m/%Y')
        age = (issueDate-birthDate).days//365
        if len(self.getID()) == 12: # CCCD
            if 14 <= age < 23:
                return birthDate.replace(year=birthDate.year+25).strftime('%d/%m/%Y')
            elif 23 <= age < 38:
                return birthDate.replace(year=birthDate.year+40).strftime('%d/%m/%Y')
            elif 38 <= age < 58:
                return birthDate.replace(year=birthDate.year+60).strftime('%d/%m/%Y')
            else:
                return issueDate.replace(year=issueDate.year+15).strftime('%d/%m/%Y')
        return issueDate.replace(year=issueDate.year+15).strftime('%d/%m/%Y')

    def writeHistory(self,fileName):
        data = self.readHistory()
        data.add(fileName)
        with open(self.__history,'wb') as file:
            pickle.dump(data,file)

    def readHistory(self):
        with open(self.__history,'rb') as file:
            return pickle.load(file)

    def getLength(self):
        return len(self.__contents)
    
    def select(self,pointer:int):
        self.selectedFile = self.__fileNames[pointer]
        self.selectedContent = self.__contents[pointer]
        self.selectedImage = self.__images[pointer]
        self.pointer = pointer
    
    def drop(self):
        self.__fileNames.pop(self.pointer)
        self.__contents.pop(self.pointer)
        self.__images.pop(self.pointer)

    @_returnEmptyStringIfNoMatch
    def getSoTK(self):
        pattern = r'(Số tài khoản\*?:\s*0 2 2)([FC]\d{6})\b'
        return '022' + re.search(pattern,self.selectedContent).group(2).strip()

    def getMaChiNhanh(self):
        return self.selectedFile.split('.')[0]

    def getTellerEmails(self):
        emailSeries = pd.read_sql(
            f"""
            SELECT
                [t2].[Email]
            FROM [DWH-CoSo].[dbo].[010005] [t1]
            LEFT JOIN [DWH-Base].[dbo].[Employee] [t2]
                ON LEFT([t1].[TenNguoiSuDung],4) = RIGHT([t2].[User],4)
            LEFT JOIN [DWH-CoSo].[dbo].[branch] [t3] 
                ON [t1].[ChiNhanh] = [t3].[branch_name]
            WHERE [t1].[TrangThai] = N'Hoạt động'
                AND [t1].[PhongBanTrucThuoc] = 'Trading Support'
                AND [t3].[branch_id] = '{self.getMaChiNhanh()}'
            """,
            connect_DWH_CoSo
        ).squeeze(axis=1)
        # return '; '.join(emailSeries)
        return 'hiepdang@phs.vn; namtran@phs.vn'

    @_returnEmptyStringIfNoMatch
    def getTenKhachHang(self):
        pattern = r'(KHÁCH HÀNG\*?:\s*)(.*)(\sNam)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    def getGioiTinh(self):
        return self.__findGioiTinh()

    @_returnEmptyStringIfNoMatch
    def getNgaySinh(self):
        pattern = r'(Ngày sinh\*?:\s*)(\d{2}/\d{2}/\d{4})(\s+Nơi sinh)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    @_returnEmptyStringIfNoMatch
    def getNoiSinh(self):
        pattern = r'(Nơi sinh\*?:\s*)(.*)(\s+Quốc tịch)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    @_returnEmptyStringIfNoMatch
    def getQuocTich(self):
        pattern = r'(Quốc tịch\*?:\s*)(.*)(\nSố CMND)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    @_returnEmptyStringIfNoMatch
    def getID(self):
        pattern = r'(Số CMND/CCCD/Hộ chiếu\*:\s*)(.*)(\nNgày cấp)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    @_returnEmptyStringIfNoMatch
    def getNgayCap(self):
        pattern = r'(Ngày cấp\*?:\s*)(\d{2}/\d{2}/\d{4})(\s+Nơi cấp)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getNoiCap(self):
        pattern = r'(Nơi cấp\*?:\s*)(.*)(\nĐịa chỉ thường trú)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    @_returnEmptyStringIfNoMatch
    def getDiaChiLienLac(self):
        pattern = r'(Địa chỉ liên lạc\*?:\s*)(.*)(\nĐiện thoại)'
        value = re.search(pattern,self.selectedContent).group(2).strip()
        if len(value) < 15: # địa chỉ không được ít hơn 15 ký tự
            return value + ' (ĐỊA CHỈ KHÔNG HỢP LỆ, VUI LÒNG ĐIỀU CHỈNH)'
        return value

    @_returnEmptyStringIfNoMatch
    def getDienThoaiCoDinh(self):
        pattern = r'(Điện thoại cố định\*?:\s*)(\d*)(\s+Di động)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    @_returnEmptyStringIfNoMatch
    def getDiDong(self):
        pattern = r'(Di động\*?:\s*)(\d*)(\s+Email)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    @_returnEmptyStringIfNoMatch
    def getEmail(self):
        pattern = r'(Email\*?:\s*)(.*)(\nNơi công tác)'
        return re.search(pattern,self.selectedContent).group(2).strip()
    
    def getNgayHetHan(self):
        return self.__findNgayHetHan()
    
    @_returnEmptyStringIfNoMatch
    def getDaiDien(self):
        pattern = r'(Đại diện\*?:\s*)(.*)(\s*Chức)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getTinhThanh(self):
        pattern = r'(Chức vụ\*?:\s*)\s*(.*)\b'
        result = re.search(pattern,self.selectedContent).group(2).strip()
        if 'Hà Nội' in result or 'Thanh Xuân' in result:
            return 'Ha noi'
        elif 'Hải Phòng' in result:
            return 'Haiphong'
        else:
            return 'Ho Chi Minh'

    @_returnEmptyStringIfNoMatch
    def getMaPIN(self):
        pattern = r'(qua điện \nthoại\s{,2})(.*)\b'
        return re.search(pattern,self.selectedContent).group(2).strip()

    def getGDDienThoai(self):
        if self.getMaPIN():
            return "Có"
        else:
            return "Không"


os.chdir(r'C:\Users\hiepdang\PycharmProjects\DataAnalytics\automation\flex_gui\F111001\dist\F111001')
class F111001(Flex):

    def __init__(self,username,password):
        super().__init__()
        self.start(existing=False) # Tạo Flex instance mới
        self.login(username,password)
        self.insertFuncCode('111001')
        self.funcWindow = self.app.window(auto_id='frmSearch')
        self.insertWindow = self.app.window(auto_id='frmCFMAST') # only exists after __clickCreateButton()
        self.__contract = None
        self.__outlook = Dispatch('outlook.application')
        self.__mail = None

    def openContract(self,openDate:dt.datetime):

        def __takeFlexScreenshot(window):
            setFocus(window)
            return cv.cvtColor(np.array(window.capture_as_image()),cv.COLOR_RGB2BGR)

        def __clickDate(dateBox): # click at day box, 3 pixel from the left is safe
            dateBox.click_input(coords=(3,dateBox.rectangle().height()//2))

        def __clickCreateButton():
            setFocus(self.funcWindow)
            while True: # đợi maximize window xong
                if self.funcWindow.rectangle() == self.mainWindow.rectangle():
                    break
                time.sleep(0.5)
            actionWindow = self.funcWindow.child_window(title='SMS')
            actionImage = __takeFlexScreenshot(actionWindow)
            actionImage = actionImage[:,:-10,:]
            unique, count = np.unique(actionImage,return_counts=True,axis=1)
            mostFrequentColumn = unique[:,np.argmax(count),:]
            columnMask = ~(actionImage==mostFrequentColumn[:,np.newaxis,:]).all(axis=(0,2))
            lastColumn = np.argwhere(columnMask).max()
            croppedImage = actionImage[:,:lastColumn,:]
            midPoint = croppedImage.shape[1]//7, croppedImage.shape[0]//2
            actionWindow.click_input(coords=midPoint)
            self.insertWindow.wait('exists',timeout=30)

        def __insertFromClipboard(textBox,textString:str):
            setFocus(self.insertWindow)
            textBox.click_input()
            pyperclip.copy(textString)
            textBox.type_keys('^a{DELETE}')
            textBox.type_keys('^v')

        def __insertFromKeyboard(textBox,textString:str):
            setFocus(self.insertWindow)
            textBox.click_input()
            textBox.type_keys('^a{DELETE}')
            textBox.type_keys(textString,with_spaces=True)

        def __insertDate(dateBox,dateString:str):
            setFocus(self.insertWindow)
            __clickDate(dateBox)
            dateBox.type_keys('{RIGHT}'*2)
            dayString, monthString, yearString = dateString.split('/')
            for valueString in [yearString,monthString,dayString]:
                dateBox.type_keys(valueString+'{LEFT}')

        def __insertSoTaiKhoan(textBox,value):
            setFocus(self.insertWindow)
            # Xóa nội dung trong box
            textBox.click_input()
            textBox.type_keys('^a{DELETE}')
            if value: # có số tài khoản trên hợp đồng -> nhập
                textBox.type_keys(value)
            else: # không có số tài khoản trên hợp đồng -> click cho hệ thống tự sinh
                autoGenButton = self.insertWindow.child_window(auto_id='btnGenCheckCUSTODYCD')
                autoGenButton.click_input()

        def __createMaKhachHang(): # click cho hệ thống tự sinh
            autoGetMaKHButton = self.insertWindow.child_window(auto_id='btnGenCheckCustID')
            autoGetMaKHButton.click_input()

        def __clickAcceptButton():
            setFocus(self.insertWindow)
            acceptButton = self.insertWindow.child_window(auto_id='btnOK')
            acceptButton.click_input()

        def __clickSuccess():
            successWindow = self.app.window(title='FlexCustodian')
            successWindow.wait('exists',timeout=30)
            okButton = successWindow.child_window(title='OK')
            okButton.click_input()

        def __closePopUps():
            # Đóng toàn bộ Pop Up nếu có
            while True:
                popUpWindow = self.app.window(best_match='Dialog')
                if popUpWindow.exists():
                    btnOK = popUpWindow.child_window(title='OK')
                    btnOK.click()
                else: break
            popUpWindow.wait_not('exists',timeout=30)

        def __attachFile(attachmentName):
            cv.imwrite(attachmentName,self.__contract.selectedImage)
            cwd = os.getcwd()
            self.__mail.Attachments.Add(os.path.join(cwd,f'{attachmentName}'))

        def __removeAttachedFile(attachmentName):
            os.remove(attachmentName)

        def __emailMissingInfo(missingFields): # Thiếu thông tin KH
            self.__mail = self.__outlook.CreateItem(0)
            self.__mail.To = self.__contract.getTellerEmails()
            self.__mail.Subject = f"Thiếu thông tin hợp đồng mở tài khoản"
            attachmentName = f"{self.__contract.selectedFile.replace('.pdf','.png')}"
            __attachFile(attachmentName)
            htmlBody = f"""
            <html>
                <head></head>
                <body>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Hợp đồng trong file đính kèm thiếu các thông tin sau:
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        {missingFields}
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Vui lòng tạo hợp đồng mới với đầy đủ các thông tin trên.
                    </p>
                    <p style="font-family:Times New Roman; font-size:80%"><i>
                        -- Generated by Reporting System
                    </i></p>
                </body>
            </html>
            """
            self.__mail.HTMLBody = htmlBody
            self.__mail.Send()
            __removeAttachedFile(attachmentName)

        def __emailMultipleAccounts(existingAccount): # Trùng thông tin KH
            self.__mail = self.__outlook.CreateItem(0)
            self.__mail.To = self.__contract.getTellerEmails()
            self.__mail.Subject = f"Trùng thông tin khách hàng trên hệ thống"
            attachmentName = f"{self.__contract.selectedFile.replace('.pdf','.png')}"
            __attachFile(attachmentName)
            htmlBody = f"""
            <html>
                <head></head>
                <body>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Hợp đồng trong file đính kèm trùng thông tin với tài khoản sau:
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        {existingAccount}
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Vui lòng kiểm tra lại với khách hàng.
                    </p>
                    <p style="font-family:Times New Roman; font-size:80%"><i>
                        -- Generated by Reporting System
                    </i></p>
                </body>
            </html>
            """
            self.__mail.HTMLBody = htmlBody
            self.__mail.Send()
            __removeAttachedFile(attachmentName)

        def __emailDuplicatedID(existingID): # Trùng số ID
            self.__mail = self.__outlook.CreateItem(0)
            self.__mail.To = self.__contract.getTellerEmails()
            self.__mail.Subject = f"Mã giấy tờ đã tồn tại"
            attachmentName = f"{self.__contract.selectedFile.replace('.pdf','.png')}"
            __attachFile(attachmentName)
            htmlBody = f"""
            <html>
                <head></head>
                <body>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Hợp đồng trong file đính kèm trùng Mã Giấy Tờ với tài khoản sau:
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        {existingID}
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Vui lòng kiểm tra lại với khách hàng.
                    </p>
                    <p style="font-family:Times New Roman; font-size:80%"><i>
                        -- Generated by Reporting System
                    </i></p>
                </body>
            </html>
            """
            self.__mail.HTMLBody = htmlBody
            self.__mail.Send()
            __removeAttachedFile(attachmentName)

        def __emailNotEnoughAge(): # KH không đủ tuổi
            self.__mail = self.__outlook.CreateItem(0)
            self.__mail.To = self.__contract.getTellerEmails()
            self.__mail.Subject = f"Khách hàng dưới 16 tuổi"
            attachmentName = f"{self.__contract.selectedFile.replace('.pdf','.png')}"
            __attachFile(attachmentName)
            htmlBody = f"""
            <html>
                <head></head>
                <body>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Khách hàng trong file đính kèm nhỏ hơn 16 tuổi:
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Vui lòng kiểm tra lại với khách hàng.
                    </p>
                    <p style="font-family:Times New Roman; font-size:80%"><i>
                        -- Generated by Reporting System
                    </i></p>
                </body>
            </html>
            """
            self.__mail.HTMLBody = htmlBody
            self.__mail.Send()
            __removeAttachedFile(attachmentName)

        def __emailExpiredID(): # ID hết hạn
            self.__mail = self.__outlook.CreateItem(0)
            self.__mail.To = self.__contract.getTellerEmails()
            self.__mail.Subject = f"Giấy tờ khách hàng hết hạn"
            attachmentName = f"{self.__contract.selectedFile.replace('.pdf','.png')}"
            __attachFile(attachmentName)
            htmlBody = f"""
            <html>
                <head></head>
                <body>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Khách hàng trong file đính kèm có ngày hết hạn giấy tờ nhỏ hơn ngày hệ thống:
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Vui lòng kiểm tra lại với khách hàng.
                    </p>
                    <p style="font-family:Times New Roman; font-size:80%"><i>
                        -- Generated by Reporting System
                    </i></p>
                </body>
            </html>
            """
            self.__mail.HTMLBody = htmlBody
            self.__mail.Send()
            __removeAttachedFile(attachmentName)

        def __checkRun(runTime:dt.datetime):
            if runTime.weekday() in (5,6): # thứ 7, CN
                self.app.kill(soft=False) # đóng app
                return False
            if dt.time(7,0,0) <= runTime.time() <= dt.time(19,0,0):
                return True
            # Note: thêm SLEEP_TIME tránh bug
            return False

        print('::: WELCOME :::')
        print('Flex GUI Automation - Author: Hiep Dang')
        print('===========================================\n')

        logging.basicConfig(
            filename=f'.\main.log',
            format='%(asctime)s: %(levelname)s: %(message)s',
            level=logging.DEBUG,
        )

        # Lấy danh sách toàn bộ tài khoản để kiểm tra
        dateString = openDate.strftime('%Y-%m-%d')
        SYNC(connect_DWH_CoSo,'spaccount',dateString,dateString)
        allCustomerTable = pd.read_sql(
            """
            SELECT
                [account_code] [SoTaiKhoan],
                [customer_name] [TenKhachHang],
                [date_of_birth] [NgaySinh],
                [id_type] [LoaiGiayTo],
                [customer_id_number] [MaGiayTo],
                [date_of_issue] [NgayCap],
                [place_of_issue] [NoiCap]
            FROM [account]
            """,
            connect_DWH_CoSo
        )
        allCustomerTable['NgaySinh'] = allCustomerTable['NgaySinh'].fillna(dt.datetime(1900,1,1))
        allCustomerTable['NgayCap'] = allCustomerTable['NgayCap'].fillna(dt.datetime(1900,1,1))
        allCustomerTable = allCustomerTable.fillna('*****')
        allCustomerTable['NgaySinh'] = allCustomerTable['NgaySinh'].dt.strftime('%d/%m/%Y')
        allCustomerTable['NgayCap'] = allCustomerTable['NgayCap'].dt.strftime('%d/%m/%Y')

        def __processFunc(name,dob):
            return unidecode.unidecode(name).replace(' ','').lower() + dob.replace('/','')

        allCustomerTable['NameAndBOD'] = allCustomerTable.apply(
            func=lambda x:__processFunc(x['TenKhachHang'],x['NgaySinh']),
            axis=1
        )

        nameBODSet = set(allCustomerTable['NameAndBOD']) # Để kiểm tra một khách hàng mở nhiều tài khoản
        idSet = set(allCustomerTable['MaGiayTo']) # Để kiểm tra trùng Mã Giấy Tờ

        while __checkRun(dt.datetime.now()):
            try:
                # instantiate Contract object
                self.__contract = Contract(openDate)
                # kiểm tra xem có phát sinh mới không
                if self.__contract.getLength() == 0:
                    # Ghi log
                    message = "No contract ::"
                    print(message)
                    logging.info(message)
                    # sleep SLEEP_TIME giây sau mỗi vòng lặp
                    time.sleep(SLEEP_TIME)
                    continue

                # **** Kiểm tra các hợp đồng có đầy đủ các thông tin bắt buộc chưa ****
                fieldMapper = {
                    self.__contract.getTenKhachHang: 'Tên Khách Hàng',
                    self.__contract.getQuocTich: 'Quốc Tịch',
                    self.__contract.getGioiTinh: 'Giới Tính',
                    self.__contract.getTinhThanh: 'Tỉnh Thành',
                    self.__contract.getID: 'Số CCCD / Passport',
                    self.__contract.getNgayCap: 'Ngày Cấp',
                    self.__contract.getNoiCap: 'Nơi Cấp',
                    self.__contract.getNgaySinh: 'Ngày Sinh',
                    self.__contract.getDiDong: 'Di Động',
                    self.__contract.getEmail: 'Email',
                    self.__contract.getDiaChiLienLac: 'Địa chỉ liên lạc',
                } # các thông tin bắt buộc

                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    missingList = []
                    for getField in fieldMapper.keys():
                        if getField() == '':
                            missingList.append(fieldMapper[getField])

                    if missingList: # thiếu ít nhất 1 giá trị bắt buộc
                        # gửi mail cho teller báo thiếu thông tin hợp đồng
                        __emailMissingInfo('; '.join(missingList))
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # Ghi log
                        message = f"Not enough information {self.__contract.getTenKhachHang()} ::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1
                # **** Kiểm tra một khách hàng mở nhiều tài khoản ****
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    customerName = self.__contract.getTenKhachHang()
                    customerDOB = self.__contract.getNgaySinh()
                    checkValue = __processFunc(customerName,customerDOB)
                    if checkValue in nameBODSet:
                        # gửi mail cho teller báo trùng khách hàng
                        existingRecord = allCustomerTable.loc[
                            allCustomerTable['NameAndBOD']==checkValue,
                            ['SoTaiKhoan','TenKhachHang','NgaySinh']
                        ].head(1)
                        existingAccount = (' '*5).join(existingRecord.squeeze(axis=0))
                        __emailMultipleAccounts(existingAccount)
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # Ghi log
                        message = f"Customer exists {self.__contract.getTenKhachHang()} {self.__contract.getNgaySinh()} {existingAccount}::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1
                # **** Kiểm tra trùng Mã Giấy Tờ ****
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    checkValue = self.__contract.getID()
                    if checkValue in idSet:
                        # gửi mail cho teller báo trùng Mã Giấy Tờ
                        existingRecord = allCustomerTable.loc[
                            allCustomerTable['MaGiayTo']==checkValue,
                            ['SoTaiKhoan','TenKhachHang','LoaiGiayTo','MaGiayTo','NgayCap','NoiCap']
                        ].head(1)
                        existingID = (' '*5).join(existingRecord.squeeze(axis=0))
                        __emailDuplicatedID(existingID)
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # Ghi log
                        message = f"Duplicated customer ID {self.__contract.getTenKhachHang()} {existingID}::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1
                # Kiểm tra KH đủ tuổi chưa
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    date, month, year = self.__contract.getNgaySinh().split('/')
                    checkValue = dt.datetime(int(year),int(month),int(date))
                    if checkValue < dt.datetime(int(year)-16,int(month),int(date)): # KH dưới 16 tuổi
                        # Gửi mail cho teller báo KH không đủ tuổi
                        __emailNotEnoughAge()
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # Ghi log
                        message = f"Customer not enough age {self.__contract.getTenKhachHang()} {self.__contract.getNgaySinh()}::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1
                # Kiểm tra ID còn hạn không
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    date, month, year = self.__contract.getNgayHetHan().split('/')
                    checkValue = dt.datetime(int(year),int(month),int(date))
                    today  = dt.datetime.now().replace(hour=0,minute=0,second=0,microsecond=0)
                    if checkValue <= today: # ID hết hạn
                        # Gửi mail cho teller báo ID hết hạn
                        __emailExpiredID()
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # Ghi log
                        message = f"Customer expired documents {self.__contract.getTenKhachHang()} {self.__contract.getNgayHetHan()}::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1

                length = self.__contract.getLength()
                for contractIndex in range(length):
                    # Trỏ chọn hợp đồng
                    self.__contract.select(contractIndex)
                    # Click "Thêm mới" để tạo hợp đồng
                    __clickCreateButton()
                    # Điền từ trái qua phải, từ trên xuống
                    # Điền Quốc Gia
                    textBox = self.insertWindow.child_window(auto_id='cboCOUNTRY')
                    __insertFromClipboard(textBox,self.__contract.getQuocTich())
                    # Điền Tỉnh Thành
                    textBox = self.insertWindow.child_window(auto_id='cboPROVINCE')
                    __insertFromClipboard(textBox,self.__contract.getTinhThanh())
                    # Điền Mã Khách Hàng
                    __createMaKhachHang()
                    # Điền Số Tài Khoản
                    textBox = self.insertWindow.child_window(auto_id='txtCUSTODYCD')
                    __insertSoTaiKhoan(textBox,self.__contract.getSoTK())
                    flexAccountCode = textBox.window_text()
                    # Điền Tên Khách Hàng
                    textBox = self.insertWindow.child_window(auto_id='txtFULLNAME')
                    __insertFromKeyboard(textBox,self.__contract.getTenKhachHang())
                    # Điền Loại Giấy Tờ
                    textBox = self.insertWindow.child_window(auto_id='cboIDTYPE')
                    if 'F' in flexAccountCode:
                        textString = 'Trading code'
                    else:
                        textString = 'CMND'
                    __insertFromClipboard(textBox,textString)
                    # Điền Trading Code
                    if 'F' in flexAccountCode:
                        textBox = self.insertWindow.child_window(auto_id='txtTRADINGCODE')
                        __insertFromKeyboard(textBox,flexAccountCode[-6:]) # 6 giá trị cuối của Accound Code
                    # Điền CCCD / Passport
                    textBox = self.insertWindow.child_window(auto_id='txtIDCODE')
                    __insertFromKeyboard(textBox,self.__contract.getID())
                    # Điền Ngày Cấp
                    dateBox = self.insertWindow.child_window(auto_id='dtpIDDATE')
                    __insertDate(dateBox,self.__contract.getNgayCap())
                    # Điền Ngày Hết Hạn
                    dateBox = self.insertWindow.child_window(auto_id='dtpIDEXPIRED')
                    __insertDate(dateBox,self.__contract.getNgayHetHan())
                    # Điền Nơi Cấp
                    textBox = self.insertWindow.child_window(auto_id='txtIDPLACE')
                    __insertFromKeyboard(textBox,self.__contract.getNoiCap())
                    # Điền Địa Chỉ Liên Lạc
                    textBox = self.insertWindow.child_window(auto_id='txtADDRESS')
                    __insertFromKeyboard(textBox,self.__contract.getDiaChiLienLac())
                    # Điền Giao Dịch Điện Thoại
                    textBox = self.insertWindow.child_window(auto_id='cboTRADETELEPHONE')
                    __insertFromClipboard(textBox,self.__contract.getGDDienThoai())
                    flexGiaoDichDienThoai = textBox.window_text()
                    # Điền Số Di Động
                    textBox = self.insertWindow.child_window(auto_id='txtMOBILESMS')
                    __insertFromKeyboard(textBox,self.__contract.getDiDong())
                    # Điền Email
                    textBox = self.insertWindow.child_window(auto_id='txtEMAIL')
                    __insertFromKeyboard(textBox,self.__contract.getEmail())
                    # Điền Ngày Sinh
                    dateBox = self.insertWindow.child_window(auto_id='dtpDATEOFBIRTH')
                    __insertDate(dateBox,self.__contract.getNgaySinh())
                    # Điền Nơi Sinh
                    textBox = self.insertWindow.child_window(auto_id='txtPLACEOFBIRTH')
                    __insertFromKeyboard(textBox,self.__contract.getNoiSinh())
                    # Điền Mã PIN
                    if flexGiaoDichDienThoai == 'Có':
                        textBox = self.insertWindow.child_window(auto_id='txtPLACEOFBIRTH')
                        __insertFromKeyboard(textBox,self.__contract.getMaPIN())
                    # Điền Giới Tính
                    textBox = self.insertWindow.child_window(auto_id='cboSEX')
                    __insertFromClipboard(textBox,self.__contract.getGioiTinh())
                    # Click Chấp Nhận
                    __clickAcceptButton()
                    # Click OK "Thêm dữ liệu thành công"
                    __clickSuccess()
                    # Đóng toàn bộ pop up (nếu có)
                    __closePopUps()
                    # Ghi History
                    self.__contract.writeHistory(self.__contract.selectedFile)
                    # Ghi log
                    message = f"Open contract :: {self.__contract.getTenKhachHang()} {self.__contract.getNgaySinh()} {self.__contract.getID()}"
                    print(message)
                    logging.info(message)

            except (Exception,):
                # Đóng toàn bộ pop up đang mở, nhập contract tiếp theo (nhưng không tắt màn hình nhập)
                __closePopUps()
                # Ghi log
                message = traceback.format_exc()
                print(message)
                logging.critical(message)
            # Sleep SLEEP_TIME giây sau mỗi vòng lặp
            time.sleep(SLEEP_TIME)



if __name__ == '__main__':
    try:
        # flexObject = F111001('2008','Ly281000@')
        flexObject = F111001('admin','123456')
        flexObject.openContract(dt.datetime.now())
    except (Exception,): # để debug
        print(traceback.format_exc())
        input('Press any key to quit: ')
        try: # Khi cửa sổ Flex còn mở
            app = Application().connect(title_re='^\.::.*Flex.*',timeout=10)
            app.kill()
        except (Exception,): # Khi cửa sổ Flex đã đóng sẵn
            pass

