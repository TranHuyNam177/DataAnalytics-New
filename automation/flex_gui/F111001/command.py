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

# When run on dev mode
os.chdir(r'C:\Users\hiepdang\PycharmProjects\DataAnalytics\automation\flex_gui\F111001\dist\F111001')

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

SLEEP_TIME = 3 # số giây nghỉ sau mỗi vòng lặp

def _returnEmptyStringIfNoMatch(func):
    @wraps(func)
    def wrapper(*args,**kwargs):
        try:
            return func(*args,**kwargs)
        except (AttributeError,): # AttributeError: 'NoneType' object has no attribute 'group'
            return ''
    return wrapper

def _doNotInsertEmptyString(func):
    @wraps(func)
    def wrapper(*args):
        if args[1] != '':
            func(*args)
    return wrapper

class NothingToInsert(Exception):
    pass

class Contract:

    PATH = r'\\192.168.66.150\du lieu hop dong tu link hđ'
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
        self.__contractTypes = self.__findLoaiHopDong()
        self.pointer = 0 # default
        self.selectedFile = self.__fileNames[self.pointer]
        self.selectedContent = self.__contents[self.pointer]
        self.selectedImage = self.__images[self.pointer]
        self.selectedContractType = self.__contractTypes[self.pointer]

    def __filterFiles(self,file):
        pattern = r'(\d{4})\.(\d{2}-\d{2}-\d{4})\.[A-Za-z,]+'
        regexSearch = re.search(pattern,file)
        if regexSearch:
            return regexSearch.group(2)==self.date.strftime('%d-%m-%Y')
        else:
            return False

    def __pdf2Text(self):
        contents = []
        for pdfFileName in self.__fileNames:
            filePDFPath = os.path.join(self.PATH,pdfFileName)
            rawContent = ''
            for pageNumber in range(2): # Đọc 2 trang đầu
                rawContent += '\n ' + PyPDF2.PdfReader(filePDFPath).pages[pageNumber].extract_text()
            rawContent = re.sub(r':\s*\.+',': ',rawContent)
            rawContent = re.sub(r'…','',rawContent)
            rawContent = rawContent.replace('\x05',' ').replace('\x01',' ')
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

    def __findLoaiHopDong(self):
        contractTypes = []
        for content in self.__contents:
            if re.search(r'Loại tài khoản\s*:\s*Individual',content):
                contractTypes.append('CaNhanNuocNgoai')
            elif re.search(r'Loại tài khoản\s*:\s*Cá nhân',content):
                contractTypes.append('CaNhanTrongNuoc')
            elif re.search(r'Loại tài khoản\s*:\s*Organization',content):
                contractTypes.append('ToChucNuocNgoai')
            else:
                contractTypes.append('ToChucTrongNuoc')
        return contractTypes

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
        self.selectedContractType = self.__contractTypes[pointer]
        self.pointer = pointer

    def drop(self):
        self.__fileNames.pop(self.pointer)
        self.__contents.pop(self.pointer)
        self.__images.pop(self.pointer)
        self.__contractTypes.pop(self.pointer)

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
        # return 'hiepdang@phs.vn; namtran@phs.vn'
        return 'lytran@phs.vn; hathai@phs.vn'

    def getCareBy(self):
        carebyTable = pd.read_sql(
            f"""
            SELECT TOP 1 [TenNhomQuanLy] [CareByGroup] FROM [010002]
            WHERE [TenNhomQuanLy] LIKE 'CareBy%' AND [MaChiNhanh] = '{self.getMaChiNhanh()}'
            """,
            connect_DWH_CoSo,
        )
        if carebyTable.empty:
            return 'CAREBY_COMMON'
        else:
            return carebyTable.squeeze()

    @_returnEmptyStringIfNoMatch
    def getTenKhachHang(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(KHÁCH HÀNG\*?:\s*)(.*)(\sNam)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(KHÁCH HÀNG\*?:\s*)(.*)(\sMale)'
        elif self.selectedContractType == 'ToChucTrongNuoc':
            pattern = r'(Tên Công ty\*?:\s*)(.*)(\sĐịa chỉ)'
        else:
            pattern = r'(Tên Công ty\*?:\s*)(.*)(\sAddress)'
        return re.search(pattern,self.selectedContent).group(2).strip().upper()

    def getGioiTinh(self):
        # convert to binary image
        grayscaleImage = cv.cvtColor(self.selectedImage,cv.COLOR_BGR2GRAY)
        _,binaryImage = cv.threshold(grayscaleImage,200,255,cv.THRESH_BINARY_INV)
        # tìm 2 ô giới tính (tọa độ cố định vì đã cố định image width = 1000 và aspect ratio ko đổi)
        if self.selectedContractType in ('ToChucNuocNgoai','ToChucTrongNuoc'):
            return ''
        elif self.selectedContractType == 'CaNhanTrongNuoc':
            maleBox = binaryImage[534:547,724:736]
            femaleBox = binaryImage[534:547,808:820]
        else:  # 'CaNhanNuocNgoai'
            maleBox = binaryImage[686:699,734:746]
            femaleBox = binaryImage[686:699,876:888]
        maleScore = maleBox.sum()
        femaleScore = femaleBox.sum()
        if maleScore > femaleScore:
            return 'Nam'
        elif maleScore < femaleScore:
            return 'Nữ'
        else:
            return ''

    @_returnEmptyStringIfNoMatch
    def getNgaySinh(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(Ngày sinh\*?:\s*)(\d{2}/\d{2}/\d{4})(\s+Nơi sinh)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Ngày sinh\*?:\s*)(\d{2}/\d{2}/\d{4})(\s+Place)'
        else:
            return ''
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getNoiSinh(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(Nơi sinh\*?:\s*)(.*)(\s+Quốc tịch)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Nơi sinh\*?:\s*)(.*)(\s+Nationality)'
        else:
            return ''
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getQuocTich(self):
        if self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Quốc tịch\*?:\s*)(.*)(\s+ID)'
            return re.search(pattern,self.selectedContent).group(2).strip()
        elif self.selectedContractType == 'ToChucNuocNgoai':
            return 'Alaska'
        return 'Viet Nam'

    @_returnEmptyStringIfNoMatch
    def getLoaiGiayTo(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            return 'CMND'
        elif self.selectedContractType == 'ToChucTrongNuoc':
            return 'Giấy phép KD'
        return 'Trading code'

    def getNgayCapTradingCode(self): # SS đang set tạm rule này, có thể thay đổi trong tương lai
        if self.getSoTK() and self.getLoaiGiayTo() == 'Trading code':  # Có số tài khoản & Loại giấy tờ = 'Trading code'
            return dt.datetime.now().strftime('%d/%m/%Y')
        return '' # Không nhập

    @_returnEmptyStringIfNoMatch
    def getMaGiayTo(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(Hộ chiếu\*?:\s*)(.*)(\s+Ngày cấp)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Hộ chiếu\*?:\s*)(.*)(\s+Date)'
        elif self.selectedContractType == 'ToChucTrongNuoc':
            pattern = r'(GCNĐKDN\*?:\s*)(.*)(\s+Ngày)'
        else:
            pattern = r'(GCNĐKDN\*?:\s*)(.*)(\s+Date)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getNgayCap(self):
        if self.selectedContractType in ('CaNhanTrongNuoc','ToChucTrongNuoc'):
            pattern = r'(Ngày cấp\*?:\s*)(\d{2}/\d{2}/\d{4})(\s+Nơi cấp)'
        else:
            pattern = r'(Ngày cấp\*?:\s*)(\d{2}/\d{2}/\d{4})(\s+Place)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getNoiCap(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(Nơi cấp\*?:\s*)(.*)(\s+Địa chỉ thường trú)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Nơi cấp\*?\s*:\s*)(.*)(\s+Permanent)'
        elif self.selectedContractType == 'ToChucTrongNuoc':
            pattern = r'(Nơi cấp\*?:\s*)(.*)(\s+Mã số thuế)'
        else:
            pattern = r'(Nơi cấp\*?:\s*)(.*)(\s+Tax code)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getDiaChi(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(Địa chỉ liên lạc\*?:\s*)(.*)(\s+Điện thoại)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Địa chỉ liên lạc\*?:\s*)(.*)(\s+Telephone)'
        elif self.selectedContractType == 'ToChucTrongNuoc':
            pattern = r'(Địa chỉ\*?:\s*)(.*)(\s+Điện thoại)'
        else:
            pattern = r'(Địa chỉ\*?:\s*)(.*)(\s+Telephone)'
        value = re.search(pattern,self.selectedContent.replace('\n',' ')).group(2).strip()
        if len(value) < 15: # địa chỉ không được ít hơn 15 ký tự
            return value + ' ----- ĐỊA CHỈ KHÔNG HỢP LỆ' # nếu không điền địa chỉ thì trả ra '' để gửi mail báo thiếu thông tin
        return value

    @_returnEmptyStringIfNoMatch
    def getDienThoaiCoDinh(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(Điện thoại cố định\*?:\s*)(\d*)(\s+Di động)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Điện thoại cố định\*?:\s*)(\d*)(\s+Mobile)'
        else:
            pattern = r'(Điện thoại\*?:\s*)(\d*)(\s+Fax)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getDienThoaiDiDong(self):
        pattern = r'(Di động\*?:\s*)(\d*)(\s+Email)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getEmail(self):
        if self.selectedContractType == 'CaNhanTrongNuoc':
            pattern = r'(Email\*?:\s*)(.*)(\s+Nơi công tác)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Email\*?:\s*)(.*)(\s+Employer)'
        elif self.selectedContractType == 'ToChucTrongNuoc':
            pattern = r'(Email\*?:\s*)(.*)(\s+Số Giấy phép thành lập)'
        else:
            pattern = r'(Email\*?:\s*)(.*)(\s+License)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    def getNgayHetHan(self):
        if self.selectedContractType.startswith('ToChuc'):
            return ''
        issueDateString = self.getNgayCap()
        birthDateString = self.getNgaySinh()
        if issueDateString == '' or birthDateString == '':  # thiếu các thông tin này là gửi mail và không nhập
            return ''
        issueDate = dt.datetime.strptime(issueDateString,'%d/%m/%Y')
        birthDate = dt.datetime.strptime(birthDateString,'%d/%m/%Y')
        age = (issueDate - birthDate).days // 365
        if self.selectedContractType == 'CaNhanTrongNuoc':
            if len(self.getMaGiayTo()) == 12: # CCCD
                if 14 <= age < 23:
                    return birthDate.replace(year=birthDate.year+25).strftime('%d/%m/%Y')
                elif 23 <= age < 38:
                    return birthDate.replace(year=birthDate.year+40).strftime('%d/%m/%Y')
                elif 38 <= age < 58:
                    return birthDate.replace(year=birthDate.year+60).strftime('%d/%m/%Y')
                else:
                    return issueDate.replace(year=issueDate.year+15).strftime('%d/%m/%Y')
            elif len(self.getMaGiayTo()) == 9: # CMND
                return issueDate.replace(year=issueDate.year+15).strftime('%d/%m/%Y')
            else: # Rule do SS set
                return '31/12/2100'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            return '31/12/2100'  # Không có ngày hết hạn -> không nhập

    @_returnEmptyStringIfNoMatch
    def getMaSoThue(self):
        if self.selectedContractType in ('CaNhanTrongNuoc','CaNhanNuocNgoai'):
            return '' # không nhập
        elif self.selectedContractType == 'ToChucTrongNuoc':
            pattern = r'(Mã số thuế\*?:\s*)(.*)(\s+Vốn điều lệ)'
        else:
            pattern = r'(Mã số thuế\*?:\s*)(.*)(\s+Charter capital)'
        return re.search(pattern,self.selectedContent).group(2).strip()

    @_returnEmptyStringIfNoMatch
    def getTinhThanh(self):
        if self.selectedContractType in ('CaNhanTrongNuoc','ToChucTrongNuoc'):
            pattern = r'(Chức vụ\*?:\s*)(.*)(\s+Theo ủy quyền số)'
        elif self.selectedContractType == 'CaNhanNuocNgoai':
            pattern = r'(Chức vụ\*?:\s*)(.*)(\s+Under Power of Attorney)'
        else:
            pattern = r'(Chức vụ\*?:\s*)(.*)(\s+Power of Attorney)'

        result = unidecode.unidecode(re.search(pattern,self.selectedContent).group(2)).replace(' ','').lower()
        if 'hanoi' in result or 'thanhxuan' in result:
            return 'Ha noi'
        elif 'haiphong' in result:
            return 'Haiphong'
        else:
            return 'Ho Chi Minh'


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

        @_doNotInsertEmptyString
        def __insertFromClipboard(textBox,textString:str):
            setFocus(self.insertWindow)
            textBox.click_input()
            pyperclip.copy(textString)
            textBox.type_keys('^a{DELETE}')
            textBox.type_keys('^v')

        @_doNotInsertEmptyString
        def __insertFromKeyboard(textBox,textString:str):
            setFocus(self.insertWindow)
            textBox.click_input()
            textBox.type_keys('^a{DELETE}')
            textBox.type_keys(textString,with_spaces=True)

        @_doNotInsertEmptyString
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
            successWindow.wait('exists',timeout=45) # chờ ghi nhận dữ liệu xong
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
            print(attachmentName)
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

        def __emailAddressTooShort(userAddress):
            self.__mail = self.__outlook.CreateItem(0)
            self.__mail.To = self.__contract.getTellerEmails()
            self.__mail.Subject = f"Địa chỉ liên lạc ít hơn 15 ký tự"
            attachmentName = f"{self.__contract.selectedFile.replace('.pdf','.png')}"
            print(attachmentName)
            __attachFile(attachmentName)
            htmlBody = f"""
            <html>
                <head></head>
                <body>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Hợp đồng trong file đính kèm có địa chỉ liên lạc ít hơn 15 ký tự:
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        {userAddress}
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        Vui lòng tạo hợp đồng mới với địa chỉ liên lạc đầy đủ.
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

        def __emailMultipleAccounts(existingValues,existingFields): # Trùng thông tin KH
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
                        Hợp đồng trong file đính kèm trùng {existingFields} với tài khoản sau:
                    </p>
                    <p style="font-family:Times New Roman; font-size:100%">
                        {existingValues}
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
            if dt.time(7,0,0) <= runTime.time() <= dt.time(22,0,0):
                return True
            return False

        print('::: WELCOME :::')
        print('Flex GUI Automation - Author: Hiep Dang')
        print('===========================================\n')

        logging.basicConfig(
            filename=f'.\main.log',
            format='%(asctime)s: %(levelname)s: %(message)s',
            level=logging.DEBUG,
        )

        dbSyncedTime = set()
        while __checkRun(dt.datetime.now()):
            try:
                # Quét lại DB 2 phút 1 lần
                now = dt.datetime.now()
                if now.minute % 2 == 0 and now.strftime('%H%M') not in dbSyncedTime:
                    dateString = openDate.strftime('%Y-%m-%d')
                    SYNC(connect_DWH_CoSo,'spaccount',dateString,dateString)
                # Instantiate Contract object
                self.__contract = Contract(openDate) # nếu không có phát sinh hợp đồng mới sẽ dừng ở đây, chờ vòng lặp tiếp theo
                # Lấy danh sách toàn bộ tài khoản để kiểm tra
                allCustomerTable = pd.read_sql(
                    """
                    SELECT
                        [account_code] [SoTaiKhoan],
                        [customer_name] [TenKhachHang],
                        [nationality] [QuocTich],
                        [date_of_birth] [NgaySinh],
                        [id_type] [LoaiGiayTo],
                        [customer_id_number] [MaGiayTo],
                        [date_of_issue] [NgayCap],
                        [place_of_issue] [NoiCap]
                    FROM [account_UAT]
                    -- FROM [account]
                    WHERE [date_of_close] IS NULL
                    """,
                    connect_DWH_CoSo
                )
                allCustomerTable['NgaySinh'] = allCustomerTable['NgaySinh'].fillna(dt.datetime(1900,1,1))
                allCustomerTable['NgayCap'] = allCustomerTable['NgayCap'].fillna(dt.datetime(1900,1,1))
                allCustomerTable = allCustomerTable.fillna('*****')
                allCustomerTable['NgaySinh'] = allCustomerTable['NgaySinh'].dt.strftime('%d/%m/%Y')
                allCustomerTable['NgayCap'] = allCustomerTable['NgayCap'].dt.strftime('%d/%m/%Y')

                def __processTwoInfo(name,dob):
                    return unidecode.unidecode(name).replace(' ','').lower() + dob.replace('/','')
                
                def __processThreeInfo(name,dob,issuedate):
                    return unidecode.unidecode(name).replace(' ','').lower() + dob.replace('/','') + issuedate.replace('/','')

                allCustomerTable['TwoInfoSet'] = allCustomerTable.apply(
                    func=lambda x:__processTwoInfo(x['TenKhachHang'],x['NgaySinh']),
                    axis=1
                )
                allCustomerTable['ThreeInfoSet'] = allCustomerTable.apply(
                    func=lambda x:__processThreeInfo(x['TenKhachHang'],x['NgaySinh'],x['NgayCap']),
                    axis=1
                )
                twoInfoSet = set(allCustomerTable['TwoInfoSet'])  # Để kiểm tra một khách hàng mở nhiều tài khoản
                threeInfoSet = set(allCustomerTable['ThreeInfoSet'])  # Để kiểm tra một khách hàng mở nhiều tài khoản
                idSet = set(allCustomerTable['MaGiayTo'])  # Để kiểm tra trùng Mã Giấy Tờ

                # Kiểm tra các hợp đồng có đầy đủ các thông tin bắt buộc chưa
                fieldMapper = {
                    self.__contract.getTenKhachHang: 'Tên Khách Hàng',
                    self.__contract.getQuocTich: 'Quốc Tịch',
                    self.__contract.getGioiTinh: 'Giới Tính',
                    self.__contract.getTinhThanh: 'Tỉnh Thành',
                    self.__contract.getMaGiayTo: 'Mã Giấy Tờ',
                    self.__contract.getNgayCap: 'Ngày Cấp',
                    self.__contract.getMaSoThue: 'Mã Số Thuế',
                    self.__contract.getNoiCap: 'Nơi Cấp',
                    self.__contract.getNgaySinh: 'Ngày Sinh',
                    self.__contract.getDienThoaiDiDong: 'Di Động',
                    self.__contract.getDienThoaiCoDinh:'Điện Thoại Cố Định',
                    self.__contract.getDiaChi: 'Địa chỉ liên lạc'
                } # các thông tin bắt buộc
                exceptionMapper = {
                    'CaNhanTrongNuoc': [
                        self.__contract.getMaSoThue,
                        self.__contract.getDienThoaiCoDinh,
                        self.__contract.getDiaChi,
                    ],
                    'CaNhanNuocNgoai': [
                        self.__contract.getMaSoThue,
                        self.__contract.getDienThoaiCoDinh,
                        self.__contract.getDiaChi,
                    ],
                    'ToChucTrongNuoc': [
                        self.__contract.getGioiTinh,
                        self.__contract.getNgaySinh,
                        self.__contract.getDienThoaiDiDong,
                    ],
                    'ToChucNuocNgoai': [
                        self.__contract.getGioiTinh,
                        self.__contract.getNgaySinh,
                        self.__contract.getDienThoaiDiDong,
                    ]
                }
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    missingFields = []
                    for getField in fieldMapper.keys():
                        if getField in exceptionMapper[self.__contract.selectedContractType]:
                            continue
                        if getField() == '':
                            missingFields.append(fieldMapper[getField])
                    if missingFields: # thiếu ít nhất 1 giá trị bắt buộc
                        # gửi mail cho teller báo thiếu thông tin hợp đồng
                        __emailMissingInfo('; '.join(missingFields))
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # Ghi log
                        customerNameLog = unidecode.unidecode(self.__contract.getTenKhachHang())
                        missingFieldsLog = ' '.join([unidecode.unidecode(field.replace(' ','')) for field in missingFields])
                        message = f"Not enough information {customerNameLog} {missingFieldsLog}::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1

                # Kiểm tra một khách hàng mở nhiều tài khoản (chỉ áp dụng cho CaNhanTrongNuoc CaNhanNuocNgoai)
                for contractIndex in range(self.__contract.getLength()):
                    self.__contract.select(contractIndex)
                    if self.__contract.selectedContractType.startswith('ToChuc'):
                        continue
                    customerName = self.__contract.getTenKhachHang()
                    customerDOB = self.__contract.getNgaySinh()
                    checkValue2Info = __processTwoInfo(customerName,customerDOB)
                    if checkValue2Info in twoInfoSet:
                        # gửi mail cho teller báo trùng khách hàng (vẫn cho nhập tự động)
                        customerIssueDate = self.__contract.getNgayCap()
                        checkValue3Info = __processThreeInfo(customerName,customerDOB,customerIssueDate)
                        if checkValue3Info not in threeInfoSet:
                            columnList = ['SoTaiKhoan','TenKhachHang','NgaySinh']
                            existingFields = 'Tên, Ngày Sinh'
                        else:
                            columnList = ['SoTaiKhoan','TenKhachHang','NgaySinh','NgayCap']
                            existingFields = 'Tên, Ngày Sinh, Ngày Cấp'
                        existingRecord = allCustomerTable.loc[
                            allCustomerTable['TwoInfoSet']==checkValue2Info,
                            columnList
                        ].head(1)
                        existingValues = (' '*5).join(existingRecord.squeeze(axis=0))
                        __emailMultipleAccounts(existingValues,existingFields)
                        # Ghi log
                        customerNameLog = unidecode.unidecode(self.__contract.getTenKhachHang())
                        message = f"Customer exists {existingValues.split()[0]} {customerNameLog} {existingValues.split()[2:]}::"
                        print(message)
                        logging.info(message)

                # Kiểm tra địa chỉ < 15 ký tự (áp dụng cho toàn bộ khách hàng)
                for contractIndex in range(self.__contract.getLength()):
                    self.__contract.select(contractIndex)
                    if 'ĐỊA CHỈ KHÔNG HỢP LỆ' in self.__contract.getDiaChi():
                        # Gửi mail cho teller báo địa chỉ < 15 ký tự (nhưng vẫn cho nhập với chuỗi " ----- ĐỊA CHỈ KHÔNG HỢP LỆ")
                        userAddress = self.__contract.getDiaChi().split('-')[0].strip()
                        __emailAddressTooShort(userAddress)
                        # Ghi log
                        customerNameLog = unidecode.unidecode(self.__contract.getTenKhachHang())
                        message = f"Address less than 15 chars {customerNameLog} {userAddress}::"
                        print(message)
                        logging.info(message)

                # Kiểm tra trùng Mã Giấy Tờ (áp dụng cho toàn bộ khách hàng)
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    checkValue = self.__contract.getMaGiayTo()
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
                        customerNameLog = unidecode.unidecode(self.__contract.getTenKhachHang())
                        message = f"Duplicated customer ID {customerNameLog} {existingID}::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1

                # Kiểm tra KH chưa đủ tuổi (áp dụng với CaNhanTrongNuoc CaNhanNuocNgoai)
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    if self.__contract.selectedContractType.startswith('ToChuc'):
                        contractIndex += 1
                        continue
                    birthDay, birthMonth, birthYear = self.__contract.getNgaySinh().split('/')
                    checkValue = dt.date(int(birthYear)+16,int(birthMonth),int(birthDay)) # thời điểm đủ 16 tuổi
                    if checkValue > dt.date.today(): # KH dưới 16 tuổi
                        # Gửi mail cho teller báo KH không đủ tuổi
                        __emailNotEnoughAge()
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # ghi log
                        customerNameLog = unidecode.unidecode(self.__contract.getTenKhachHang())
                        message = f"Customer not enough age {customerNameLog} {self.__contract.getNgaySinh()}::"
                        print(message)
                        logging.info(message)
                    else:
                        contractIndex += 1

                # Kiểm tra ID hết hạn (chỉ áp dụng cho CaNhanTrongNuoc)
                contractIndex = 0
                while contractIndex < self.__contract.getLength():
                    self.__contract.select(contractIndex)
                    if self.__contract.selectedContractType != 'CaNhanTrongNuoc':
                        contractIndex += 1
                        continue
                    print(self.__contract.getTenKhachHang())
                    print(self.__contract.getNgayHetHan())
                    expireDay, expireMonth, expireYear = self.__contract.getNgayHetHan().split('/')
                    expireDate = dt.date(int(expireYear),int(expireMonth),int(expireDay))
                    if expireDate <= dt.date.today(): # ID hết hạn
                        # Gửi mail cho teller báo ID hết hạn
                        __emailExpiredID()
                        # write history để không nhập nữa, teller tự nhập bằng tay
                        self.__contract.writeHistory(self.__contract.selectedFile)
                        # xóa khỏi danh sách hợp đồng cần insert
                        self.__contract.drop()
                        # Ghi log
                        customerNameLog = unidecode.unidecode(self.__contract.getTenKhachHang())
                        message = f"Customer expired documents {customerNameLog} {self.__contract.getNgayHetHan()}::"
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
                    # Điền Loại KH
                    textBox = self.insertWindow.child_window(auto_id='cboCUSTTYPE')
                    if self.__contract.selectedContractType.startswith('CaNhan'):
                        value = 'Cá nhân'
                    else:
                        value = 'Tổ chức'
                    __insertFromClipboard(textBox,value)
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
                    accountCodeFlex = textBox.window_text()
                    accountCodeFlex = '022FIW0003'
                    __insertFromClipboard(textBox,accountCodeFlex)
                    # Điền Tên Khách Hàng
                    textBox = self.insertWindow.child_window(auto_id='txtFULLNAME')
                    __insertFromKeyboard(textBox,self.__contract.getTenKhachHang())
                    # Điền Loại Giấy Tờ
                    textBox = self.insertWindow.child_window(auto_id='cboIDTYPE')
                    __insertFromClipboard(textBox,self.__contract.getLoaiGiayTo())
                    # Điền Trading Code
                    if self.__contract.selectedContractType.endswith('NuocNgoai'):
                        textBox = self.insertWindow.child_window(auto_id='txtTRADINGCODE')
                        __insertFromKeyboard(textBox,accountCodeFlex[-6:]) # nhập 6 ký tự cuối của số TK
                    # Điền Mã giấy tờ
                    textBox = self.insertWindow.child_window(auto_id='txtIDCODE')
                    __insertFromKeyboard(textBox,self.__contract.getMaGiayTo())
                    # Điền Ngày Cấp
                    dateBox = self.insertWindow.child_window(auto_id='dtpIDDATE')
                    __insertDate(dateBox,self.__contract.getNgayCap())
                    # Điền Ngày Hết Hạn
                    dateBox = self.insertWindow.child_window(auto_id='dtpIDEXPIRED')
                    __insertDate(dateBox,self.__contract.getNgayHetHan())
                    # Điền Nơi Cấp
                    textBox = self.insertWindow.child_window(auto_id='txtIDPLACE')
                    __insertFromKeyboard(textBox,self.__contract.getNoiCap())
                    # Điền Mã Số Thuế
                    textBox = self.insertWindow.child_window(auto_id='txtTAXCODE')
                    __insertFromKeyboard(textBox,self.__contract.getMaSoThue())
                    # Điền Địa Chỉ
                    textBox = self.insertWindow.child_window(auto_id='txtADDRESS')
                    __insertFromKeyboard(textBox,self.__contract.getDiaChi())
                    # Điền Giao Dịch Điện Thoại
                    textBox = self.insertWindow.child_window(auto_id='cboTRADETELEPHONE')
                    __insertFromClipboard(textBox,'Không') # Điền cứng theo rule SS
                    # Điền Số Di Động & Số Cố Định
                    mobile1Box = self.insertWindow.child_window(auto_id='txtMOBILESMS')
                    mobile2Box = self.insertWindow.child_window(auto_id='txtMOBILE')
                    if self.__contract.getDienThoaiCoDinh() and self.__contract.getDienThoaiDiDong(): # có cả 2
                        __insertFromKeyboard(mobile1Box,self.__contract.getDienThoaiDiDong())
                        __insertFromKeyboard(mobile2Box,self.__contract.getDienThoaiCoDinh())
                    else: # chỉ có 1 trong 2 thì ưu tiên nhập mobile 1, không có TH thiếu cả 2 vì đã chặn ở trên rồi
                        __insertFromKeyboard(mobile1Box,self.__contract.getDienThoaiDiDong()+self.__contract.getDienThoaiCoDinh())  # thiếu tức là ''
                    # Điền Email
                    textBox = self.insertWindow.child_window(auto_id='txtEMAIL')
                    __insertFromKeyboard(textBox,self.__contract.getEmail())
                    # Điền Ngày Sinh
                    dateBox = self.insertWindow.child_window(auto_id='dtpDATEOFBIRTH')
                    __insertDate(dateBox,self.__contract.getNgaySinh())
                    # Điền Nơi Sinh
                    textBox = self.insertWindow.child_window(auto_id='txtPLACEOFBIRTH')
                    __insertFromKeyboard(textBox,self.__contract.getNoiSinh())
                    # Điền Giới Tính
                    textBox = self.insertWindow.child_window(auto_id='cboSEX')
                    __insertFromClipboard(textBox,self.__contract.getGioiTinh())
                    # Điền CareBy
                    textBox = self.insertWindow.child_window(auto_id='cboCAREBY')
                    __insertFromClipboard(textBox,self.__contract.getCareBy())
                    # Click Chấp Nhận
                    __clickAcceptButton()
                    # Click OK "Thêm dữ liệu thành công"
                    __clickSuccess()
                    # Đóng toàn bộ pop up (nếu có)
                    __closePopUps()
                    # Ghi History
                    self.__contract.writeHistory(self.__contract.selectedFile)
                    # Ghi log
                    message = f"Open contract :: {unidecode.unidecode(self.__contract.getTenKhachHang())} {self.__contract.getNgaySinh()} {self.__contract.getMaGiayTo()}"
                    print(message)
                    logging.info(message)

            except (NothingToInsert,):
                message = "No contract ::"
                print(message)
                logging.info(message)
                # sleep SLEEP_TIME giây sau mỗi vòng lặp
                time.sleep(SLEEP_TIME)
                continue

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
        flexObject.openContract(dt.datetime(2022,10,15))
    except (Exception,): # để debug
        print(traceback.format_exc())
        input('Press any key to quit: ')
        try: # Khi cửa sổ Flex còn mở
            app = Application().connect(title_re='^\.::.*Flex.*',timeout=10)
            app.kill()
        except (Exception,): # Khi cửa sổ Flex đã đóng sẵn
            pass