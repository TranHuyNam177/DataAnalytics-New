from automation import *
import PyPDF2
import pickle

class PDFHopDongMoTK:
    def __init__(self, date: str):
        self.__checkFileRun = set()
        self.__PATH = os.path.dirname(__file__)
        self.__date = date
        self.__listFilePDF = [
            fileName for fileName in os.listdir(os.path.join(self.__PATH, 'HopDongMoTKGDCKCoSo'))
            if re.search(r'(\d+)?\.(\d{2}-\d{2}-\d{4})', fileName).group(2) == self.__date
        ]
        self.__listPDF = self.__readPDF()
        self.__pdfContent = None
        self.__pdfPath = None

    def __del__(self):
        pickleFile = os.path.join(self.__PATH, 'savedFiles', f'{self.__date}.pkl')
        pickle.dump(self.__checkFileRun, open(pickleFile, 'wb'))

    def __getLength(self):
        return len(self.__listFilePDF)

    def __readPDF(self):
        listPDF = []
        for i in range(self.__getLength()):
            if self.__listFilePDF[i] in self.__checkFileRun:
                continue
            filePDFPath = os.path.join(self.__PATH, 'HopDongMoTKGDCKCoSo', self.__listFilePDF[i])
            pdfContent = PyPDF2.PdfReader(filePDFPath).pages[0].extract_text()
            listPDF.append((filePDFPath, pdfContent))
        return listPDF

    def selectPDF(self, i):
        self.__pdfPath, self.__pdfContent = self.__listPDF[i]
        self.__checkFileRun.add(self.__listFilePDF[i])

    def __findCoords(self, sex):
        pdfImage = convert_from_path(
            pdf_path=self.__pdfPath,
            poppler_path=r'D:\applications\poppler-0.68.0_x86\poppler-0.68.0\bin'
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

    def getSoTK(self):
        soTK = re.search(r'(Số tài khoản :) (.*)\n', self.__pdfContent).group(2)
        soTK = soTK.replace(' ', '')
        if soTK == '022':
            soTK = ''
        return soTK

    def getKhachHang(self):
        tenKH = re.search(r'(KHÁCH HÀNG:) (.*) (Nam Nữ)', self.__pdfContent).group(2)
        return tenKH

    def getGioiTinh(self):
        imgMale = self.__findCoords('male')
        imgFemale = self.__findCoords('female')
        if np.count_nonzero(imgMale == 0) > np.count_nonzero(imgFemale == 0):
            sex = 'Nam'
        elif np.count_nonzero(imgMale == 0) < np.count_nonzero(imgFemale == 0):
            sex = 'Nữ'
        else:
            sex = ''
        return sex

    def getNgaySinh(self):
        ngaySinh = re.search(r'(Ngày sinh\*:) (.*) (Nơi sinh)', self.__pdfContent).group(2)
        return ngaySinh

    def getNoiSinh(self):
        noiSinh = re.search(r'(Nơi sinh\*:) (.*) (Quốc tịch)', self.__pdfContent).group(2)
        return noiSinh

    def getQuocTich(self):
        quocTich = re.search(r'(Quốc tịch\*:) (.*)\n', self.__pdfContent).group(2)
        return quocTich

    def getCMNDCCCD(self):
        soCMNDCCCD = re.search(r'(Số CMND/CCCD/Hộ chiếu\*:) (.*)\n', self.__pdfContent).group(2)
        return soCMNDCCCD

    def getNgayCap(self):
        ngayCap = re.search(r'(Ngày cấp\*:) (.*) (Nơi)', self.__pdfContent).group(2)
        return ngayCap

    def getNoiCap(self):
        noiCap = re.search(r'(Nơi cấp\*:) (.*)\n', self.__pdfContent).group(2)
        return noiCap

    def getDiaChiThuongTru(self):
        diaChiThuongTru = re.search(r'(Địa chỉ thường trú\*:) (.*)\n', self.__pdfContent).group(2)
        return diaChiThuongTru

    def getDiaChiLienLac(self):
        diaChiLienLac = re.search(r'(Địa chỉ liên lạc\*:) (.*)\n', self.__pdfContent).group(2)
        return diaChiLienLac

    def getDienThoaiCoDinh(self):
        dienThoai = re.search(r'(Điện thoại cố định:) (.*) (Di động)', self.__pdfContent).group(2)
        return dienThoai

    def getDiDong(self):
        diDong = re.search(r'(Di động\*:) (.*) (Email)', self.__pdfContent).group(2)
        return diDong

    def getEmail(self):
        email = re.search(r'(Email\*:) (.*)\n', self.__pdfContent).group(2)
        return email

    def getDaiDien(self):
        daiDien = re.search(r'(Đại diện:) (.*) (Chức)', self.__pdfContent).group(2)
        return daiDien

    def getChucVu(self):
        chucVu = re.search(r'(Chức vụ:) (.*)\n', self.__pdfContent).group(2)
        return chucVu