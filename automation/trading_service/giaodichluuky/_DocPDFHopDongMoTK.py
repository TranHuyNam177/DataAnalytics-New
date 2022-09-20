from pdf2image import convert_from_path
import os
import re
import pandas as pd
import cv2
import numpy as np
import PyPDF2
from unidecode import unidecode


def convertPDFtoImage(pathPDF):
    """
    Hàm chuyển từng trang trong PDF sang image
    """
    images = convert_from_path(
        pdf_path=pathPDF,
        poppler_path=r'D:\applications\poppler-0.68.0_x86\poppler-0.68.0\bin'
    )
    return images

def _findCoords(pdfImage, sex):
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


def docFilePDF(customerName: str):
    customerName = unidecode(customerName).replace(' ', '').upper()
    directory = r"D:\Learn Python\Data Analysis\Contract\HopDongMoTKGDCKCoSo"
    records = []
    for fileName in os.listdir(directory):
        fileName = fileName.upper()
        if customerName in fileName:
            filePDFPath = os.path.join(directory, fileName)
            pdfContent = PyPDF2.PdfReader(filePDFPath).pages[0].extract_text()
            # Số tài khoản
            soTK = re.search(r'(Số tài khoản :) (.*)\n', pdfContent).group(2)
            soTK = soTK.replace(' ', '')
            if soTK == '022':
                soTK = ''
            # Khách hàng
            tenKH = re.search(r'(KHÁCH HÀNG:) (.*) (Nam Nữ)', pdfContent).group(2)
            # Ngày sinh
            ngaySinh = re.search(r'(Ngày sinh\*:) (.*) (Nơi sinh)', pdfContent).group(2)
            # Nơi sinh
            noiSinh = re.search(r'(Nơi sinh\*:) (.*) (Quốc tịch)', pdfContent).group(2)
            # Quốc tịch
            quocTich = re.search(r'(Quốc tịch\*:) (.*)\n', pdfContent).group(2)
            # CMND/CCCD/Hộ chiếu
            soCMNDCCCD = re.search(r'(Số CMND/CCCD/Hộ chiếu\*:) (.*)\n', pdfContent).group(2)
            # Ngày cấp
            ngayCap = re.search(r'(Ngày cấp\*:) (.*) (Nơi)', pdfContent).group(2)
            # Nơi cấp
            noiCap = re.search(r'(Nơi cấp\*:) (.*)\n', pdfContent).group(2)
            # Địa chỉ thường trú
            diaChiThuongTru = re.search(r'(Địa chỉ thường trú\*:) (.*)\n', pdfContent).group(2)
            # Địa chỉ liên lạc
            diaChiLienLac = re.search(r'(Địa chỉ liên lạc\*:) (.*)\n', pdfContent).group(2)
            # Điện thoại cố định
            dienThoai = re.search(r'(Điện thoại cố định:) (.*) (Di động)', pdfContent).group(2)
            # Di động
            diDong = re.search(r'(Di động\*:) (.*) (Email)', pdfContent).group(2)
            # Email
            email = re.search(r'(Email\*:) (.*)\n', pdfContent).group(2)
            # Đại diện
            daiDien = re.search(r'(Đại diện:) (.*) (Chức)', pdfContent).group(2)
            # Chức vụ
            chucVu = re.search(r'(Chức vụ:) (.*)\n', pdfContent).group(2)

            # Giới tính
            pdfImage = convertPDFtoImage(filePDFPath)[0]
            pdfImage = cv2.cvtColor(np.array(pdfImage), cv2.COLOR_BGR2GRAY)
            _, pdfImage = cv2.threshold(pdfImage, 200, 255, cv2.THRESH_BINARY)
            imgMale = _findCoords(pdfImage, 'male')
            imgFemale = _findCoords(pdfImage, 'female')
            if np.count_nonzero(imgMale == 0) > np.count_nonzero(imgFemale == 0):
                sex = 'Nam'
            elif np.count_nonzero(imgMale == 0) < np.count_nonzero(imgFemale == 0):
                sex = 'Nữ'
            else:
                sex = ''
            # Append data
            records.append((soTK, tenKH, sex, ngaySinh, noiSinh, quocTich, soCMNDCCCD, ngayCap, noiCap, diaChiThuongTru, diaChiLienLac, dienThoai, diDong, email, daiDien, chucVu))

    balanceTable = pd.DataFrame(
        data=records,
        columns=[
            'SoTaiKhoan',
            'KhachHang',
            'GioiTinh',
            'NgaySinh',
            'NoiSinh',
            'QuocTich',
            'SoCMND/CCCD/HoChieu',
            'NgayCap',
            'NoiCap',
            'DiaChiThuongTru',
            'DiaChiLienLac',
            'DienThoaiCoDinh',
            'DiDong',
            'Email',
            'DaiDien',
            'ChucVu'
        ]
    ).reset_index(drop=True)
    balanceTable = balanceTable.replace(to_replace=r'(\.{5,})', value='', regex=True)

    return balanceTable
