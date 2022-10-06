import warnings
warnings.filterwarnings('ignore')
import re
import numpy as np
import pandas as pd
import time
import datetime as dt
import unidecode
from automation.flex_gui.base import Flex, setFocus
from pywinauto.application import Application
import cv2 as cv
import logging
import traceback
import pyodbc
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'../../tesseract/tesseract.exe'

from datawarehouse.DWH_CoSo import SYNC
from datawarehouse import BDATE

# # when run on dev mode
# import os
# os.chdir(r"C:\Users\hiepdang\PycharmProjects\DataAnalytics\automation\flex_gui\VCI1104\dist\VCI1104")

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

class VCI1104(Flex):

    def __init__(self,username,password):
        super().__init__()
        self.start(existing=False) # Tạo Flex instance mới
        self.login(username,password)
        self.insertFuncCode('VCI1104')
        self.funcWindow = self.app.window(title_re=r'\b.*1104.*\b')
        self.funcWindow.wait('exists',timeout=10)
        self.funcWindow.maximize()
        self.dataWindow = None

    def pushCashOutOrders(self):

        """
        Xuất lệnh chuyển tiền ra ngoài (Realtime)
        """
        def __findBankCodeBankName(cashOut,rawBankName):

            # Nhỏ hơn 100tr, mặc định vào VTB
            if cashOut <= 100e6:
                return '0010', 'VTB'
            # Lớn hơn 100tr, vào tài khoản thụ hưởng
            # Nếu tài khoản thụ hưởng không phải OCB, TCB, VCB, EIB, BIDV, VTB thì mặc định vào VTB
            convertedBankName = unidecode.unidecode(rawBankName).lower().replace(' ','').replace('y','i').replace('nganhang','')
            if convertedBankName in ['tmcpphuongdong','ocb','phuongdong']:
                bankCode = '0003'
                bankName = 'OCB'
            elif convertedBankName in ['techcombank','kithuong','tcb']:
                bankCode = '0011'
                bankName = 'TCB'
            elif convertedBankName in ['vcb','tmcpngoaithuong','ngoaithuong','vietcombank']:
                bankCode = '0009'
                bankName = 'VCB'
            elif convertedBankName in ['eib','eximbank','xuatnhapkhau','tmcpxuatnhapkhauvietnam']:
                bankCode = '0002'
                bankName = 'EIB'
            elif convertedBankName in ['bidv','dautuvaphattrien','dautuvaphattrienvietnam']:
                bankCode = '0005'
                bankName = 'BIDV'
            else:
                bankCode = '0010'
                bankName = 'VTB'

            return bankCode, bankName


        def __findTopLeftPoint(containingImage,templateImage):
            containingImage = np.array(containingImage) # đảm bảo image để được đưa về numpy array
            templateImage = np.array(templateImage) # đảm bảo image để được đưa về numpy array
            matchResult = cv.matchTemplate(containingImage,templateImage,cv.TM_CCOEFF)
            _, _, _, topLeft = cv.minMaxLoc(matchResult)
            return topLeft[1], topLeft[0] # cho compatible với openCV

        def __findColumnCoords(colName):

            if colName == 'SoTienChuyen':
                fileName = 'SoTienChuyen.png'
            elif colName == 'SoTaiKhoanLuuKy':
                fileName = 'SoTaiKhoanLuuKy.png'
            elif colName == 'Select':
                fileName = 'Select.png'
            else:
                raise ValueError('colName must be either "SoTienChuyebn" or "SoTaiKhoanLuuKy" or "Select"')

            tempImage = cv.imread(fr".\{fileName}")
            tempImage = cv.cvtColor(tempImage,cv.COLOR_RGB2BGR)
            headerTop, headerLeft = __findTopLeftPoint(dataImage,tempImage)
            columnLeft = headerLeft # left của header và column là giống nhau
            columnRight = columnLeft + tempImage.shape[1]
            columnTop = headerTop + tempImage.shape[0] # top của của cột là bottom của header

            return (columnTop,columnLeft), (columnTop,columnRight)

        def __findRowHeight(binaryRecordImage):
            sumIntensity = binaryRecordImage.sum(axis=1)
            minLocArray = np.asarray(sumIntensity==sumIntensity.min()).nonzero()
            minLocArray = np.insert(minLocArray,0,0)
            rowHeightArray = np.diff(minLocArray)
            rowHeightArray = rowHeightArray[rowHeightArray<30] # row height ko thể lớn hơn 30 pixel
            rowHeight = np.median(rowHeightArray).astype(int)
            return rowHeight

        def __cropColumn(colName):
            topLeft, topRight = __findColumnCoords(colName)
            top = topLeft[0]
            left = topLeft[1]
            right = topRight[1]
            bottom = dataImage.shape[0]
            return dataImage[top:bottom,left:right,:]

        def __clickRow(rowNumber,rowHeight):
            topLeft, topRight = __findColumnCoords('Select')
            left = topLeft[1]
            right = topRight[1]
            top = topLeft[0]
            midPoint = (int(left*0.9+right*0.1),int(top+rowHeight/2*(2*rowNumber-1)))
            setFocus(self.funcWindow)
            self.dataWindow.click_input(coords=midPoint,absolute=False)

        def __readAccountNumber(image):
            content = pytesseract.pytesseract.image_to_string(
                image,
                config='--psm 4 tessedit_char_whitelist=0123456789C,.',
            )
            if not content:
                raise ValueError("Can't read Account Number")
            content = re.sub(r'\b022[,.].?[,.]','022C',content)
            content = re.sub(r'[^\dC]','',content)
            content = re.sub(r'022C',' 022C',content).strip()
            return content

        def __getFlexDataFrame(accountStrings):
            return pd.DataFrame(
                data={
                    'RowNumber':np.arange(len(accountStrings.split())) + 1,
                    'AccountNumber':accountStrings.split(),
                }
            )

        def __hasRecord(accountImage):
            accountBinaryImage = cv.cvtColor(accountImage,cv.COLOR_BGR2GRAY)
            trimmedImage = accountBinaryImage[:accountBinaryImage.shape[0]//2,:-10] # conservative approach
            emptyPixels = (trimmedImage>250).sum()
            return emptyPixels / trimmedImage.size < 0.95 # có record -> True, không có record -> False

        def __mapData(accountImage):
            dbTable = __queryVCI1104()
            thresholdRange = range(50,205,5) # xử lý ảnh với nhiều mức threshold khác nhau
            table = None
            for threshold in thresholdRange:
                # Số tài khoản
                accountBinaryImage = __processFlexImage(accountImage,threshold)
                accounts = __readAccountNumber(accountBinaryImage); print(accounts)
                # Số tiền chuyển
                flexTable  = __getFlexDataFrame(accounts)
                table = pd.merge(flexTable,dbTable,on='AccountNumber',how='left')
                table = table.drop_duplicates(subset=['RowNumber'])
                table = table.loc[table['AmountValue'].notna()]
                if not table.empty: # đọc được ít nhất 1 record
                    break
            return table.head(1) # chỉ lấy 1 record đọc được

        def __clickSearchButton():
            setFocus(self.funcWindow)
            searchButtonWindow = self.funcWindow.child_window(auto_id='btnSearch')
            searchButtonWindow.click_input()

        def __sendBankCode(bankCode):
            setFocus(self.funcWindow)
            # Xóa ký tự
            bankCodeWindow = self.funcWindow.child_window(auto_id="mskBANKID")
            existingChars = bankCodeWindow.window_text()
            for _ in range(len(existingChars)):
                bankCodeWindow.type_keys('{BACKSPACE}')
            # Nhập bank code
            bankCodeWindow.type_keys(bankCode)

        def __clickExecuteButton():
            setFocus(self.funcWindow)
            actionWindow = self.funcWindow.child_window(title='SMS')
            actionImage = __takeFlexScreenshot(actionWindow)
            actionImage = actionImage[10:,:,:]
            unique,count = np.unique(actionImage,return_counts=True,axis=1)
            mostFrequentColumn = unique[:,np.argmax(count),:]
            columnMask = ~(actionImage==mostFrequentColumn[:,np.newaxis,:]).all(axis=(0,2))
            lastColumn = np.argwhere(columnMask).max()
            croppedImage = actionImage[:,:lastColumn,:]
            midPoint = croppedImage.shape[1]//2, croppedImage.shape[0]//2
            actionWindow.click_input(coords=midPoint)

        def __takeFlexScreenshot(window):
            setFocus(self.funcWindow)
            return cv.cvtColor(np.array(window.capture_as_image()),cv.COLOR_RGB2BGR)

        def __processFlexImage(image,threshold):
            flexGrayImg = cv.cvtColor(image,cv.COLOR_RGB2GRAY)
            _, flexBinaryImg = cv.threshold(flexGrayImg,threshold,255,cv.THRESH_BINARY)
            return flexGrayImg

        def __queryVCI1104():
            dataDate = dt.datetime.now().strftime('%Y-%m-%d')
            SYNC(connect_DWH_CoSo,'spVCI1104',dataDate,dataDate)
            runTime = dt.datetime.now().time()
            if runTime <= dt.time(8,15,0): # các lần quét trước 8:15
                sinceDate = BDATE(dataDate,-1)
                sinceTime = f"'{sinceDate} 16:00:00'"
            else:
                sinceTime = 'DATEADD(SECOND,-300,GETDATE())'
            return pd.read_sql(
                f"""
                SELECT 
                    REPLACE([SoTaiKhoanLuuKy],'.','') [AccountNumber],
                    [TenKhachHang] [CustomerName],
                    [NganHangThuHuong] [ReceivingBank],
                    [SoTienChuyen] [AmountValue],
                    [ThoiGianGhiNhan] [RecordTime]
                FROM [VCI1104]
                WHERE [ThoiGianGhiNhan] >= {sinceTime}
                    AND [SoTienChuyen] < 10e9 -- Chỉ đẩy các lệnh dưới 10 tỷ
                ORDER BY [ThoiGianGhiNhan] -- Lệnh vào trước đẩy trước
                """,
                connect_DWH_CoSo,
            )

        def __getBankAutoID():
            if bankCode == '0010':
                auto_id = 'POVTB'
            elif bankCode == '0003':
                auto_id = 'POOCB'
            elif bankCode == '0002':
                auto_id = 'POEXB'
            elif bankCode == '0011':
                auto_id = 'POTCB'
            elif bankCode == '0009':
                auto_id = 'POVCB'
            elif bankCode in ('1','0005'):
                auto_id = 'POBIDV'
            else:
                raise ValueError(f'Invalid Bank Code: {bankCode}')
            return auto_id

        def __selectBankFromList(bankAutoID):
            bankSelectWindow = self.app.window(title='In bảng kê')
            time.sleep(7) # chỗ này Flex rất lag, để 7s là an toàn, ko dùng method exists được
            selectButton = bankSelectWindow.child_window(auto_id=bankAutoID)
            selectButton.click()
            bankSelectWindow.child_window(title='In',auto_id="btnPrint").click_input()

        def __clickExportPDF(bankOrderWindow):
            # Maximize window
            setFocus(bankOrderWindow)
            while bankOrderWindow.rectangle().width() != self.mainWindow.rectangle().width(): # chưa maximize xong
                time.sleep(1)
            time.sleep(1) # conservative, tránh click nhầm nút
            menuBar = bankOrderWindow.children()[-1]
            menuImage = np.array(menuBar.capture_as_image())
            clickPoint = 10, menuImage.shape[0]//2
            # (10 icon, chia thêm 2 để lấy midPoint -> chia 20)
            menuBar.click_input(coords=clickPoint,absolute=False)

        def __findFileName():
            nowString = dt.datetime.now().strftime('%H%M%S')
            if cashOut >= 100e6 and bankName == 'VTB':
                return f'UP_VTB_{customerName.title()}_{recordTime.strftime("%H%M%S")}_{nowString}'
            else:
                return f'{bankName}_{customerName.title()}_{recordTime.strftime("%H%M%S")}_{nowString}'

        def __savePDF(saveFileWindow):
            # Chọn định dạng xls
            setFocus(saveFileWindow)
            fileTypeBox = saveFileWindow.child_window(class_name='ComboBox',found_index=1)
            for _ in range(10): # Chọn tối đa 10 lần
                fileTypeBox.click_input()
                fileTypeBox.type_keys('{DOWN}{ENTER}')
                seenText = fileTypeBox.window_text()
                if 'Excel (*.xls)' in seenText:
                    break
            # Nhập tên file
            fileName = __findFileName()
            fileNameBox = saveFileWindow.child_window(class_name='ComboBox',found_index=0)
            fileNameBox.click_input()
            fileNameBox.type_keys('^a{DELETE}')
            savedFolder = r'C:\Users\Roger\Data\UyNhiemChi'
            fileNameBox.type_keys(fr'{savedFolder}\{fileName}.xls',with_spaces=True) # lưu ổ tài chính
            saveButton = saveFileWindow.child_window(title='&Save')
            # Click "Save"
            while saveButton.exists(timeout=1,retry_interval=0.5):
                saveButton.click_input()
            # Click "Ok" (Export Complete)
            confirmWindow = self.app.window(title='Export Report')
            confirmWindow.wait('exists',timeout=30)
            confirmWindow.type_keys('{ENTER}')

        def __quitPDFWindow():
            bankOrderWindow.close()

        def __checkRun(runTime:dt.datetime):
            if runTime.weekday() in (5,6): # thứ 7, CN
                self.app.kill(soft=False) # đóng app
                return False
            if dt.time(8,0,0) <= runTime.time() <= dt.time(16,1,0):
                return True
            # Note: thêm 1 phút để tránh bug (do cài sleep)
            return False

        logging.basicConfig(
            filename=f'.\pushCashOutOrders.log',
            format='%(asctime)s: %(levelname)s: %(message)s',
            level=logging.DEBUG,
        )

        # print hello lên terminal
        print('::: WELCOME :::')
        print('Flex GUI Automation - Author: Hiep Dang')
        print('===========================================\n')

        while __checkRun(dt.datetime.now()):
            try:
                # Click "Tìm kiếm" để kiểm tra các lệnh mới trên Flex
                __clickSearchButton()
                # Screen shot khung dữ liệu
                self.dataWindow = self.funcWindow.child_window(auto_id='pnlSearchResult')
                dataImage = __takeFlexScreenshot(self.dataWindow)
                # Check có dữ liệu không
                accountImage = __cropColumn('SoTaiKhoanLuuKy')
                if not __hasRecord(accountImage):
                    now = dt.datetime.now()
                    logMessage = f"""No orders at {now.strftime('%H:%M:%S')}"""
                    print(logMessage)
                    logging.info(logMessage)
                    time.sleep(15) # Nghỉ 15s sau mỗi lệnh
                    continue
                # Đọc dữ liệu
                table = __mapData(accountImage)
                table = table.set_index('RowNumber')
                table = table.sort_index(ascending=True)

                # Lấy các thông tin cơ bản
                _rowNumber = table.index[0]
                cashOut = table.loc[_rowNumber,'AmountValue']
                rawBankName = table.loc[_rowNumber,'ReceivingBank']
                bankCode, bankName = __findBankCodeBankName(cashOut,rawBankName)
                customerName = table.loc[_rowNumber,'CustomerName']
                recordTime = table.loc[_rowNumber,'RecordTime']
                # Click chọn từng lệnh
                accountBinaryImage = __processFlexImage(accountImage,100)
                rowHeight = __findRowHeight(accountBinaryImage)
                __clickRow(_rowNumber,rowHeight)
                # Nhập bank code vào khung "Mã NH UNC"
                __sendBankCode(bankCode)
                # Bấm thực hiện
                __clickExecuteButton()
                # Chọn ngân hàng
                _bankAutoID = __getBankAutoID()
                __selectBankFromList(_bankAutoID)
                # Export ủy nhiệm chi
                bankOrderWindow = self.app.window(title=_bankAutoID)
                bankOrderWindow.wait('exists',timeout=30)
                __clickExportPDF(bankOrderWindow)
                # Save ủy nhiệm chi
                saveFileWindow = self.app.window(title='Export Report')
                saveFileWindow.wait('visible',timeout=30)
                __savePDF(saveFileWindow)
                # Đóng PDF preview
                __quitPDFWindow()
                # Thông báo xử lý xong
                doneTime = dt.datetime.now()
                message = f"""
                *** Pushed at {doneTime.strftime('%H:%M:%S')} ::
                - Customer Name: {customerName}
                - Order Value: {int(cashOut):,d}
                - Placement Time: {recordTime.strftime('%H:%M:%S')}
                --------------------
                """
                print(message)
                logMessage = f'Pushed {unidecode.unidecode(customerName)} {int(cashOut):,d} VND at {doneTime}. Placed at {recordTime}'
                logging.info(logMessage)
            except (Exception,):
                # Đóng toàn bộ cửa sổ đang mở ngoài self.mainWindow và self.funcWindow
                windowToCloseTitles = [
                    'In bảng kê',
                    'Export Report',
                    'POVTB','POOCB','POEXB','POTCB','POVCB','POBIDV',
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
            # Nghỉ 15s sau mỗi lệnh
            time.sleep(10)

        self.app.kill(soft=False) # chạy xong tự động đóng đóng app


if __name__ == '__main__':
    try:
        flexObject = VCI1104('1950','ER45BY27')
        flexObject.pushCashOutOrders()
    except (Exception,): # để debug
        print(traceback.format_exc())
        input('Press any key to quit: ')
        try: # Khi cửa sổ Flex còn mở
            app = Application().connect(title_re='^\.::.*Flex.*',timeout=10)
            app.kill()
        except (Exception,): # Khi cửa sổ Flex đã đóng sẵn
            pass

r"""
command to bundle files:
cd C:\Users\hiepdang\PycharmProjects\DataAnalytics\automation\flex_gui\VCI1104\venv_vci1101\Scripts
activate
cd ../..
pyinstaller command.py -D -p C:\Users\hiepdang\PycharmProjects\DataAnalytics -n VCI1104 --add-data Select.PNG;. --add-data SoTaiKhoanLuuKy.PNG;.
"""


