import os
import re
import cv2
import PyPDF2
import pickle
import numpy as np
from pdf2image import convert_from_path

class PDFHopDongMoTK:
    def __init__(self, date: str):
        self.__checkFileRun = set()
        self.__PATH = os.path.dirname(__file__)
        self.__date = date
        self.__listFilePDF = [
            fileName for fileName in os.listdir(os.path.join(self.__PATH, 'HopDongMoTKGDCKCoSo'))
            if re.search(r'(\d+)*\.(\d{2}-\d{2}-\d{4})', fileName).group(2) == self.__date
        ]
        self.__listPDFPathContent = self.__readPDF()
        self.__pdfContent = None
        self.__pdfPath = None

    def __del__(self):
        pickleFile = os.path.join(self.__PATH, 'savedFiles', f'{self.__date}.pkl')
        pickle.dump(self.__checkFileRun, open(pickleFile, 'wb'))

    def getLength(self):
        return len(self.__listFilePDF)

    def __readPDF(self):
        listPDFPathContent = []
        for i in range(self.getLength()):
            if self.__listFilePDF[i] in self.__checkFileRun:
                continue
            filePDFPath = os.path.join(self.__PATH, 'HopDongMoTKGDCKCoSo', self.__listFilePDF[i])
            pdfContent = PyPDF2.PdfReader(filePDFPath).pages[0].extract_text()
            pdfContent = re.sub(r':\s*\.+', ': ', pdfContent)
            listPDFPathContent.append((filePDFPath, pdfContent))
        return listPDFPathContent

    def selectPDF(self, i):
        self.__pdfPath, self.__pdfContent = self.__listPDFPathContent[i]
        self.__checkFileRun.add(self.__listFilePDF[i])

    def __findCoords(self, sex):
        pdfImage = convert_from_path(
            pdf_path=self.__pdfPath,
            poppler_path=r'C:\Users\namtran\poppler-0.68.0\bin'
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

    def getSoTK(self):
        return self.__findSoTK()

    def getKhachHang(self):
        return re.search(r'(KHÁCH HÀNG:\s*)(.*)(\s+Nam Nữ)', self.__pdfContent).group(2)

    def getGioiTinh(self):
        return self.__findGioiTinh()

    def getNgaySinh(self):
        return re.search(r'(Ngày sinh\*:\s*)(.*)(\s+Nơi sinh)', self.__pdfContent).group(2)

    def getNoiSinh(self):
        return re.search(r'(Nơi sinh\*:\s*)(.*)(\s+Quốc tịch)', self.__pdfContent).group(2)

    def getQuocTich(self):
        return re.search(r'(Quốc tịch\*:\s*)(.*)(\nSố CMND)', self.__pdfContent).group(2)

    def getCMNDCCCD(self):
        return re.search(r'(Số CMND/CCCD/Hộ chiếu\*:\s*)(.*)(\nNgày cấp)', self.__pdfContent).group(2)

    def getNgayCap(self):
        return re.search(r'(Ngày cấp\*:\s*)(.*)(\s+Nơi cấp)', self.__pdfContent).group(2)

    def getNoiCap(self):
        return re.search(r'(Nơi cấp\*:\s*)(.*)(\nĐịa chỉ thường trú)', self.__pdfContent).group(2)

    def getDiaChiThuongTru(self):
        return re.search(r'(Địa chỉ thường trú\*:\s*)(.*)(\nĐịa chỉ liên lạc)', self.__pdfContent).group(2)

    def getDiaChiLienLac(self):
        return re.search(r'(Địa chỉ liên lạc\*:\s*)(.*)(\nĐiện thoại)', self.__pdfContent).group(2)

    def getDienThoaiCoDinh(self):
        return re.search(r'(Điện thoại cố định:\s*)(.*)(\s+Di động)', self.__pdfContent).group(2)

    def getDiDong(self):
        return re.search(r'(Di động\*:\s*)(.*)(\s+Email)', self.__pdfContent).group(2)

    def getEmail(self):
        return re.search(r'(Email\*:\s*)(.*)(\nNơi công tác)', self.__pdfContent).group(2)

    def getDaiDien(self):
        return re.search(r'(Đại diện:\s*)(.*)(\s+Chức vụ)', self.__pdfContent).group(2)

    def getChucVu(self):
        return re.search(r'(Chức vụ:\s*)(.*)(\nTheo ủy quyền)', self.__pdfContent).group(2)

    