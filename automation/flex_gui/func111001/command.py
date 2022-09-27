import warnings
warnings.filterwarnings('ignore')
import numpy as np
import time
import traceback
import pytesseract
import cv2
import logging
import datetime as dt
from DocPDFHopDongMoTK import PDFHopDongMoTK
from automation.flex_gui.base import Flex, setFocus
from pywinauto.application import Application
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\namtran\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

class func111001(Flex):
    def __init__(self,username,password):
        super().__init__()
        self.start(existing=False) # Tạo Flex instance mới
        self.login(username,password)
        self.insertFuncCode(self.__class__.__name__)
        self.dataWindow = None

    def openContract(self, date: str):
        pdfHopDong = PDFHopDongMoTK(date)

        def __takeFlexScreenshot(window):
            setFocus(self.funcWindow)
            return cv2.cvtColor(np.array(window.capture_as_image()), cv2.COLOR_RGB2BGR)

        def __clickDate(window):
            window.click_input(coords=(3, window.rectangle().height() // 2))

        def __clickThemMoiButton():
            setFocus(self.funcWindow)
            while True:
                time.sleep(1)
                if self.funcWindow.client_rect().width() == self.mainWindow.client_rect().width():
                    break
            actionWindow = self.funcWindow.child_window(title='SMS')
            actionImage = __takeFlexScreenshot(actionWindow)
            actionImage = actionImage[:,:-10,:]
            unique, count = np.unique(actionImage, return_counts=True, axis=1)
            mostFrequentColumn = unique[:, np.argmax(count), :]
            columnMask = ~(actionImage == mostFrequentColumn[:, np.newaxis, :]).all(axis=(0, 2))
            lastColumn = np.argwhere(columnMask).max()
            croppedImage = actionImage[:, :lastColumn, :]
            midPoint = int(croppedImage.shape[1] / 6.5), croppedImage.shape[0] // 2
            actionWindow.click_input(coords=midPoint)

        def __inputSoTK(soTK):
            if soTK != '':
                soTKBox = self.dataWindow.child_window(auto_id='txtCUSTODYCD')
                soTKBox.click_input()
                soTKBox.type_keys(soTK)
            else:
                self.dataWindow.wait('exists', timeout=30)
                autoGetSoTKButton = self.dataWindow.child_window(auto_id='btnGenCheckCUSTODYCD')
                autoGetSoTKButton.click()

        def __inputTradingCode():
            soTK = self.dataWindow.child_window(auto_id='txtCUSTODYCD').window_text()
            if '022F' in soTK:
                tradingCode = soTK.replace('022F', '')
                tradingCodeBox = self.dataWindow.child_window(auto_id='txtTRADINGCODE')
                tradingCodeBox.click_input()
                tradingCodeBox.type_keys(tradingCode)

        def __inputTinhThanh(tinhThanh):
            tinhThanhBox = self.dataWindow.child_window(auto_id='cboPROVINCE')
            tinhThanhBox.click_input()
            while True:
                tinhThanhBox.type_keys('{DOWN}', with_spaces=True)
                tinhThanhString = tinhThanhBox.window_text()
                if tinhThanhString == tinhThanh:
                    tinhThanhBox.type_keys('{ENTER}')
                    break

        def __inputHoTen(hoTen):
            hoTenBox = self.dataWindow.child_window(auto_id='txtFULLNAME')
            hoTenBox.click_input()
            hoTenBox.type_keys(hoTen, with_spaces=True)

        def __inputGioiTinh(gioiTinh):
            gioiTinhBox = self.dataWindow.child_window(auto_id='cboSEX')
            gioiTinhBox.click_input()
            gioiTinhBox.type_keys(gioiTinh)

        def __inputNgaySinh(ngaySinh):
            ngaySinhBox = self.dataWindow.child_window(auto_id='dtpDATEOFBIRTH')
            __clickDate(ngaySinhBox)
            ngaySinhBox.type_keys(f'/{ngaySinh.month}/{ngaySinh.year}/{ngaySinh.day}')

        def __inputNoiSinh(noiSinh):
            noiSinhBox = self.dataWindow.child_window(auto_id='txtPLACEOFBIRTH')
            noiSinhBox.click_input()
            noiSinhBox.type_keys(noiSinh, with_spaces=True)

        def __inputQuocTich(quocTich):
            quocTichBox = self.dataWindow.child_window(auto_id='cboCOUNTRY')
            quocTichBox.click_input()
            quocTichBox.type_keys(quocTich, with_spaces=True)

        def __inputLoaiGiayTo(loaiGiayTo):
            loaiGiayToBox = self.dataWindow.child_window(auto_id='cboIDTYPE')
            loaiGiayToBox.click_input()
            loaiGiayToBox.type_keys(loaiGiayTo, with_spaces=True)

        def __inputMaGiayTo(soCMNDCCCD):
            maGiayToBox = self.dataWindow.child_window(auto_id='txtIDCODE')
            maGiayToBox.click_input()
            maGiayToBox.type_keys(soCMNDCCCD)
            maGiayToBox.type_keys('{ENTER}')

        def __inputNgayCap(ngayCap):
            ngayCapBox = self.dataWindow.child_window(auto_id='dtpIDDATE')
            __clickDate(ngayCapBox)
            ngayCapBox.type_keys(f'/{ngayCap.month}/{ngayCap.year}/{ngayCap.day}')

        def __inputNoiCap(noiCap):
            noiCapBox = self.dataWindow.child_window(auto_id='txtIDPLACE')
            noiCapBox.click_input()
            noiCapBox.type_keys(noiCap, with_spaces=True)

        def __inputDiaChi(diaChi):
            diaChiBox = self.dataWindow.child_window(auto_id='txtADDRESS')
            diaChiBox.click_input()
            diaChiBox.type_keys(diaChi, with_spaces=True)

        def __inputMobile2(dienThoaiCoDinh):
            mobile2Box = self.dataWindow.child_window(auto_id='txtMOBILE')
            mobile2Box.click_input()
            mobile2Box.type_keys(dienThoaiCoDinh)

        def __inputMobile1(diDong):
            mobile1Box = self.dataWindow.child_window(auto_id='txtMOBILESMS')
            mobile1Box.click_input()
            mobile1Box.type_keys(diDong)

        def __inputEmail(email):
            emailBox = self.dataWindow.child_window(auto_id='txtEMAIL')
            emailBox.click_input()
            emailBox.type_keys(email)

        def __inputNgayHetHan(ngayHetHan):
            ngayHetHanBox = self.dataWindow.child_window(auto_id='dtpIDEXPIRED')
            __clickDate(ngayHetHanBox)
            ngayHetHanBox.type_keys(f'/{ngayHetHan.month}/{ngayHetHan.year}/{ngayHetHan.day}')

        def __inputGDDienThoai(gdDienThoai):
            gdDienThoaiBox = self.dataWindow.child_window(auto_id='cboTRADETELEPHONE')
            gdDienThoaiBox.type_keys('{DELETE}')
            gdDienThoaiBox.type_keys(gdDienThoai)

        def __inputMaPin(maPin):
            if self.dataWindow.child_window(auto_id='cboTRADETELEPHONE').window_text() == 'Có':
                maPINBox = self.dataWindow.child_window(auto_id='txtPIN')
                maPINBox.click_input()
                maPINBox.type_keys(maPin)

        def __clickAutoGetMaKHButton():
            autoGetMaKHButton = self.dataWindow.child_window(auto_id='btnGenCheckCustID')
            autoGetMaKHButton.click()
            self.dataWindow.wait('exists', timeout=30)

        def __clickChapNhanButton():
            chapNhanButton = self.dataWindow.child_window(auto_id='btnOK')
            chapNhanButton.click()
            # chỗ này hiện bảng thông báo "Ngày hết hạn giấy tờ phải lớn hơn ngày hệ thống"
            dialogWindow = self.app.Dialog.child_window(title='Ngày hết hạn giấy tờ phải lớn hơn ngày hệ thống!')
            if dialogWindow.exists():
                btnOK = self.app.Dialog.child_window(title='OK')
                btnOK.click()

        def __clickWindowAfterClickChapNhanButton():
            dialogWindow = self.app.Dialog
            dialogWindow.wait('exists', timeout=30)
            while True:
                if dialogWindow.exists():
                    btnOK = dialogWindow.child_window(class_name='Button')
                    btnOK.click()
                    break

        def __checkRun(runTime:dt.datetime):
            if runTime.weekday() in (5,6): # thứ 7, CN
                self.app.kill(soft=False) # đóng app
                return False
            if dt.time(8,0,0) <= runTime.time() <= dt.time(16,1,0):
                return True
            # Note: thêm 1 phút để tránh bug (do cài sleep)
            return False

        ######### RUN #########
        print('::: WELCOME :::')
        print('Flex GUI Automation - Author: Hiep Dang')
        print('===========================================\n')

        # đọc file Pickle cũ
        pdfHopDong.readPickleFile()
        for i in range(0, pdfHopDong.getLength()):
            try:
                # Nhấn nút thêm mới
                __clickThemMoiButton()
                # nhập dữ liệu vào khung dữ liệu
                self.dataWindow = self.app.window(auto_id='frmCFMAST')
                self.dataWindow.wait('exists', timeout=30)
                setFocus(self.dataWindow)
                # chọn hợp đồng để extract ra thông tin
                pdfHopDong.selectPDF(i)
                # nhập thông tin của PDF vừa extract
                __inputQuocTich(pdfHopDong.getQuocTich())
                __inputTinhThanh(pdfHopDong.getTinhThanh())
                __clickAutoGetMaKHButton()
                __inputSoTK(pdfHopDong.getSoTK())
                __inputHoTen(pdfHopDong.getKhachHang())
                __inputLoaiGiayTo('CMND')
                __inputTradingCode()
                __inputMaGiayTo(pdfHopDong.getCMNDCCCD())
                # __inputMaGiayTo('032514871254')
                notificationWindow = self.app.window(title='FlexCustodian')
                errorWindow = notificationWindow.child_window(title='Mã giấy tờ đã tồn tại!')
                btnOK = notificationWindow.child_window(title='OK')
                if errorWindow.exists():
                    message = errorWindow.window_text()
                    btnOK.click()
                    self.dataWindow.close()
                    logging.critical(message)
                    continue
                __inputNgayCap(pdfHopDong.getNgayCap())
                __inputNgayHetHan(pdfHopDong.getNgayHetHan())
                __inputNoiCap(pdfHopDong.getNoiCap())
                __inputDiaChi(pdfHopDong.getDiaChiLienLac())
                __inputMobile1(pdfHopDong.getDiDong())
                __inputNgaySinh(pdfHopDong.getNgaySinh())
                __inputGioiTinh(pdfHopDong.getGioiTinh())
                __inputGDDienThoai(pdfHopDong.getGDDienThoai())
                __inputEmail(pdfHopDong.getEmail())
                __inputMaPin('wasfwasd')
                __inputMobile2(pdfHopDong.getDienThoaiCoDinh())
                __inputNoiSinh(pdfHopDong.getNoiSinh())
                # nhấn nút chấp nhận sau khi đã insert xong hết dữ liệu
                __clickChapNhanButton()
                # nhấn nút ok trong window hiện lên sau khi nhấn nút chấp nhận
                __clickWindowAfterClickChapNhanButton()

                # lưu đè lên file pickle cũ
                pdfHopDong.writePickleFile(i)
            except (Exception,):
                # Đóng toàn bộ cửa sổ đang mở ngoài self.mainWindow và self.funcWindow
                windowToCloseTitles = [
                    'Thông tin khách hàng'
                ]
                for window in self.app.windows():
                    if window.window_text() in windowToCloseTitles:
                        try:
                            window.close()
                        except (Exception,):
                            pass
                # Ghi log
                message = traceback.format_exc()
                logging.critical(message)
        self.app.kill(soft=False)  # chạy xong tự động đóng đóng app

if __name__ == '__main__':
    try:
        flexObject = func111001('admin','123456')
        flexObject.openContract('06-09-2022')
    except (Exception,):  # để debug
        print(traceback.format_exc())
        input('Press any key to quit: ')
        try: # Khi cửa sổ Flex còn mở
            app = Application().connect(title_re='^\.::.*Flex.*',timeout=10)
            app.kill()
        except (Exception,): # Khi cửa sổ Flex đã đóng sẵn
            pass