import os
import re
import math
import cv2
import PyPDF2
import pickle
import numpy as np
from pdf2image import convert_from_path
import datetime as dt

class PDFHopDongMoTK:
    def __init__(self, date: str):
        self.__PATH = os.path.dirname(__file__)
        self.__historyPickle = os.path.join(self.__PATH, f'savedData.pkl')
        self.__date = date
        self.__dataSaved = self.readPickleFile()
        self.__listFilePDF = [
            fileName for fileName in os.listdir(os.path.join(self.__PATH, 'HopDongMoTKGDCKCoSo'))
            if fileName not in self.__dataSaved and (re.search(r'(\d+)*\.(\d{2}-\d{2}-\d{4})', fileName).group(2) == self.__date)
        ]
        self.__listPDFPathContent = self.__readPDF()
        self.__pdfContent = None
        self.__pdfPath = None
        self.__pdfCheckTradingPwdSMS = None

    def readPickleFile(self):
        with open(self.__historyPickle, 'rb') as file:
            return pickle.load(file)

    def writePickleFile(self, i):
        self.__dataSaved.add(self.__listFilePDF[i])
        with open(self.__historyPickle, 'wb') as file:
            pickle.dump(self.__dataSaved, file)

    def getLength(self):
        return len(self.__listFilePDF)

    def __readPDF(self):
        listPDFPathContent = []
        for i in range(self.getLength()):
            if self.__listFilePDF[i] in self.__dataSaved:
                continue
            filePDFPath = os.path.join(self.__PATH, 'HopDongMoTKGDCKCoSo', self.__listFilePDF[i])
            pdfInfo = PyPDF2.PdfReader(filePDFPath).pages[0].extract_text()
            pdfInfo = re.sub(r':\s*\.+', ': ', pdfInfo)
            pdfCheckTradingPwdSMS = PyPDF2.PdfReader(filePDFPath).pages[2].extract_text()
            listPDFPathContent.append((filePDFPath, pdfInfo, pdfCheckTradingPwdSMS))
        return listPDFPathContent

    def selectPDF(self, i):
        self.__pdfPath, self.__pdfContent, self.__pdfCheckTradingPwdSMS = self.__listPDFPathContent[i]

    def __findCoords(self, sex):
        pdfImage = convert_from_path(
            pdf_path=self.__pdfPath,
            poppler_path=r'C:\Users\namtran\poppler-0.68.0\bin'
            # poppler_path= r'D:\applications\poppler-0.68.0_x86\poppler-0.68.0\bin'
        )[0]
        pdfImage = cv2.cvtColor(np.array(pdfImage), cv2.COLOR_BGR2GRAY)
        _, pdfImage = cv2.threshold(pdfImage, 200, 255, cv2.THRESH_BINARY)
        smallImagePath = os.path.join(os.path.dirname(__file__), 'img', 'sex.png')
        smallImage = cv2.imread(smallImagePath, 0)  # hình trắng đen (array 2 chiều)
        w, h = smallImage.shape[::-1]
        matchResult = cv2.matchTemplate(pdfImage, smallImage, cv2.TM_CCOEFF)
        _, _, _, topLeft = cv2.minMaxLoc(matchResult)
        top, left = topLeft
        right = left + h
        if sex == 'male':
            bottom = int(top + w / 2 - 8)
            top = int(top + 0.3 * w)
        else:
            bottom = top + w
            top = int(top + 0.83 * w)
        return pdfImage[left:right, top:bottom]

    def __findGioiTinh(self):
        imgMale = self.__findCoords('male')
        imgFemale = self.__findCoords('female')
        if np.count_nonzero(imgMale == 0) > np.count_nonzero(imgFemale == 0):
            sex = 'Nam'
        elif np.count_nonzero(imgMale == 0) < np.count_nonzero(imgFemale == 0):
            sex = 'Nữ'
        else:
            sex = ''
        return sex

    def __findSoTK(self):
        soTK = re.search(r'(Số tài khoản :\s*)(.*)(\nHôm nay)', self.__pdfContent).group(2).replace(' ', '')
        if soTK == '022':
            soTK = ''
        return soTK

    def __findNgayHetHan(self):
        ngayCap = self.getNgayCap()
        ngaySinh = self.getNgaySinh()
        soCMNDCCCD = self.getCMNDCCCD()
        # soCMNDCCCD = '032514871254'
        age = math.floor((ngayCap - ngaySinh).days / 365)
        if len(soCMNDCCCD) == 12:
            if 14 <= age < 23:
                ngayHetHan = ngaySinh.replace(ngaySinh.year + 25)
            elif 23 <= age < 38:
                ngayHetHan = ngaySinh.replace(ngaySinh.year + 40)
            elif 38 <= age < 58:
                ngayHetHan = ngaySinh.replace(ngaySinh.year + 60)
            else:
                ngayHetHan = ngayCap.replace(ngayCap.year + 15)
        else:
            ngayHetHan = ngayCap.replace(ngayCap.year + 15)
        return ngayHetHan

    def __findTinhThanh(self):
        chucVu = re.search(r'(Chức vụ:\s*)(.*)(\s*\nTheo ủy quyền số)', self.__pdfContent).group(2)
        if 'Hà Nội' in chucVu or 'Thanh Xuân' in chucVu:
            tinhThanh = 'Ha noi'
        elif 'Hải Phòng' in chucVu:
            tinhThanh = 'Haiphong'
        else:
            tinhThanh = 'Ho Chi Minh'
        return tinhThanh

    def __findTradingPwdSMS(self):
        tradingPwdSMS = re.search(r'(Mật khẩu giao dịch qua điện \nthoại\s*)(.*)(\s*\n3\.\s+Dịch vụ)', self.__pdfCheckTradingPwdSMS).group(2)
        if not re.match('[A-Za-z0-9]+', tradingPwdSMS):
            tradingPwdSMS = ''
        return tradingPwdSMS

    def getTradingPwdSMS(self):
        return self.__findTradingPwdSMS()

    def getGDDienThoai(self):
        if self.__findTradingPwdSMS():
            return "Có"
        else:
            return "Không"

    def getSoTK(self):
        return self.__findSoTK()

    def getKhachHang(self):
        return re.search(r'(KHÁCH HÀNG:\s*)(.*)(\s+Nam Nữ)', self.__pdfContent).group(2)

    def getGioiTinh(self):
        return self.__findGioiTinh()

    def getNgaySinh(self):
        return dt.datetime.strptime(re.search(r'(Ngày sinh\*:\s*)(.*)(\s+Nơi sinh)', self.__pdfContent).group(2), '%d/%m/%Y')

    def getNoiSinh(self):
        return re.search(r'(Nơi sinh\*:\s*)(.*)(\s+Quốc tịch)', self.__pdfContent).group(2)

    def getQuocTich(self):
        return re.search(r'(Quốc tịch\*:\s*)(.*)(\nSố CMND)', self.__pdfContent).group(2)

    def getCMNDCCCD(self):
        return re.search(r'(Số CMND/CCCD/Hộ chiếu\*:\s*)(.*)(\nNgày cấp)', self.__pdfContent).group(2)

    def getNgayCap(self):
        return dt.datetime.strptime(re.search(r'(Ngày cấp\*:\s*)(.*)(\s+Nơi cấp)', self.__pdfContent).group(2), '%d/%m/%Y')

    def getNoiCap(self):
        return re.search(r'(Nơi cấp\*:\s*)(.*)(\nĐịa chỉ thường trú)', self.__pdfContent).group(2)

    def getDiaChiLienLac(self):
        return re.search(r'(Địa chỉ liên lạc\*:\s*)(.*)(\nĐiện thoại)', self.__pdfContent).group(2)

    def getDienThoaiCoDinh(self):
        return re.search(r'(Điện thoại cố định:\s*)(.*)(\s+Di động)', self.__pdfContent).group(2)

    def getDiDong(self):
        return re.search(r'(Di động\*:\s*)(.*)(\s+Email)', self.__pdfContent).group(2)

    def getEmail(self):
        return re.search(r'(Email\*:\s*)(.*)(\nNơi công tác)', self.__pdfContent).group(2)

    def getNgayHetHan(self):
        return self.__findNgayHetHan()

    def getDaiDien(self):
        daiDien = re.search(r'(Đại diện:\s*)(.*)(\s*Chức)', self.__pdfContent).group(2)
        return daiDien

    def getTinhThanh(self):
        return self.__findTinhThanh()
    