import datetime as dt
import os
import os.path
import re
import time
import numpy as np
import pandas as pd
import pytesseract
import unidecode
from PIL import Image
from automation.finance import BankCurrentBalance
from automation.flex_gui.base import Flex, setFocus
from datawarehouse.DWH_CoSo import connect_DWH_CoSo
import cv2 as cv


class VCI1104(Flex):

    def __init__(self,env,username,password):
        super().__init__(env)
        self.start()
        self.login(username,password)
        self.insertFuncCode(self.__class__.__name__)
        self.dataWindow = None

    def exportCashOutOrders(self):

        """
        Xuất lệnh chuyển tiền ra ngoài (Realtime)
        """

        def _checkBalance(bankFunction,bankObject) -> float:
            now = dt.datetime.now()
            bankObject = bankObject.Login()
            balanceTable = bankFunction(bankObject,now,now)
            if bankObject.bank == 'OCB':
                accountNumber = '0021100002115004'
            elif bankObject.bank == 'TCB':
                accountNumber = '19038382442022'
            elif bankObject.bank == 'VCB':
                accountNumber = '0071001264078'
            elif bankObject.bank == 'EIB':
                accountNumber = '140114851002285'
            elif bankObject.bank == 'BIDV':
                accountNumber = '26110002677688'
            elif bankObject.bank == 'VTB':
                accountNumber = '147001536591'
            else:
                raise ValueError(f'Invalid bank {bankObject.bank}')
            return balanceTable.loc[balanceTable['AccountNumber'] == accountNumber,'Balance']

        def _findBankCodeBankName(cashOut,rawBankName):

            # Nhỏ hơn 100tr, mặc định vào VTB
            if cashOut <= 100e6:
                return '0010','VTB'
            # Lớn hơn 100tr, vào tài khoản thụ hưởng
            # Nếu tài khoản thụ hưởng không phải OCB, TCB, VCB, EIB, BIDV, VTB thì mặc định vào VTB
            convertedBankName = unidecode.unidecode(rawBankName).lower().replace(' ','').replace('y','i')
            if convertedBankName in ['nganhangtmcpphuongdong','ocb','phuongdong']:
                bankCode = '0003'
                _bankFunction = BankCurrentBalance.runOCB
                _bankObject = BankCurrentBalance.OCB(True)
            elif convertedBankName in ['techcombank','kithuong','tcb']:
                bankCode = '0011'
                _bankFunction = BankCurrentBalance.runTCB
                _bankObject = BankCurrentBalance.TCB(True)
            elif convertedBankName in ['vcb','nganhangtmcpngoaithuong','ngoaithuong','vietcombank']:
                bankCode = '0009'
                _bankFunction = BankCurrentBalance.runVCB
                _bankObject = BankCurrentBalance.VCB(True)
            elif convertedBankName in ['eib','eximbank','xuatnhapkhau','nganhangtmcpxuatnhapkhauvietnam']:
                bankCode = '0002'
                _bankFunction = BankCurrentBalance.runEIB
                _bankObject = BankCurrentBalance.EIB(True)
            elif convertedBankName in ['bidv','dautuvaphattrien','nganhangdautuvaphattrienvietnam']:
                bankCode = '0005'
                _bankFunction = BankCurrentBalance.runBIDV
                _bankObject = BankCurrentBalance.BIDV(True)
            else:
                bankCode = '0010'
                _bankFunction = BankCurrentBalance.runVTB
                _bankObject = BankCurrentBalance.VTB(True)

            # Kiểm tra nếu không phải lệnh lớn
            if cashOut < 20e9:
                return bankCode,_bankObject.bank
            # Nếu là lệnh lớn
            currentBalance = _checkBalance(_bankFunction,_bankObject)
            if currentBalance >= cashOut:  # Kiểm tra nếu đủ tiền thì chuyển
                return bankCode
            # Kiểm tra nếu không đủ tiền thì chuyển bằng BIDV NKKN (bankCode=1)
            return '1','BIDV'


        def _findTopLeftPoint(containingImage,templateImage):
            containingImage = np.array(containingImage) # đảm bảo image để được đưa về numpy array
            templateImage = np.array(templateImage) # đảm bảo image để được đưa về numpy array
            matchResult = cv.matchTemplate(containingImage,templateImage,cv.TM_CCOEFF)
            _, _, _, topLeft = cv.minMaxLoc(matchResult)
            return topLeft[1], topLeft[0] # cho compatible với openCV

        def _findColumnCoords(colName):
            if colName == 'SoTienChuyen':
                fileName = 'SoTienChuyen.png'
            elif colName == 'SoTaiKhoanLuuKy':
                fileName = 'SoTaiKhoanLuuKy.png'
            elif colName == 'Select':
                fileName = 'Select.png'
            else:
                raise ValueError('colName must be either "SoTienChuyebn" or "SoTaiKhoanLuuKy" or "Select"')
            tempPath = os.path.join(os.path.dirname(__file__),fileName)
            tempImage = cv.imread(tempPath)
            tempImage = cv.cvtColor(tempImage,cv.COLOR_RGB2BGR)
            headerTop, headerLeft = _findTopLeftPoint(dataImage,tempImage)
            columnLeft = headerLeft # left của header và column là giống nhau
            columnRight = columnLeft + tempImage.shape[1]
            columnTop = headerTop + tempImage.shape[0] # top của của cột là bottom của header
            return (columnTop,columnLeft), (columnTop,columnRight)

        def _cropColumn(colName):
            topLeft, topRight = _findColumnCoords(colName)
            top = topLeft[0]
            left = topLeft[1]
            right = topRight[1]
            return dataImage[top:,left:right,:]

        def _clickRow(rowNumber):
            ROW_HEIGHT = 17
            topLeft, topRight = _findColumnCoords('Select')
            left = topLeft[1]
            right = topRight[1]
        def _readFlexDataImage(image):
            # Đọc dữ liệu từ ảnh để tìm số dòng trên màn hình Flex
            content = pytesseract.pytesseract.image_to_string(
                image,
                config='--psm 6 tessedit_char_whitelist=0123456789,.',
            )
            # contentTable['Type'] =
            content = re.sub(r'\b022[,.].[,.]','022C',content)
            for char in ('\n',', ','. '):
                content = content.replace(char,' ')
            for char in content:
                if char not in '0123456789C,. ':
                    content.replace(char,'')
            return content

        def _getFlexDataFrame(content):
            # Số tài khoản
            accountNumers = re.findall(r'022C\d{6}',content)
            # Số tiền
            valueContent = content
            for accountNumer in accountNumers:
                valueContent = valueContent.replace(accountNumer,'')
            amountStrings = re.findall(r'[\d,]*\d{3}',valueContent)
            amountValues = [float(string.replace(',','')) for string in amountStrings]
            # Số thứ tự xuất hiện trên Flex
            _rowNumbers = np.arange(len(amountValues)) + 1
            return pd.DataFrame(
                data={
                    'RowNumber':_rowNumbers,
                    'AccountNumber':accountNumers,
                    'AmountValue':amountValues,
                }
            )

        def _clickSearchButton():
            setFocus(self.funcWindow)
            searchButtonWindow = self.funcWindow.child_window(auto_id='btnSearch')
            searchButtonWindow.click()

        def _clickFirstOrder():
            setFocus(self.funcWindow)
            self.funcWindow.click_input(coords=(45,465),absolute=True)

        def _sendBankCode():
            setFocus(self.funcWindow)
            # Xóa ký tự
            bankCodeWindow = self.funcWindow.child_window(auto_id="mskBANKID")
            existingChars = bankCodeWindow.window_text()
            for _ in range(len(existingChars)):
                bankCodeWindow.type_keys('{BACKSPACE}')
            # Nhập bank code
            bankCodeWindow.type_keys(bankCode)

        def _clickExecuteButton():
            setFocus(self.funcWindow)
            self.funcWindow.click_input(coords=(110,50),absolute=True)
            self.app.windows()

        def _closePopUp():
            popUpWindow = self.app.window(title='Nhập giao dịch')
            if popUpWindow.exists():
                setFocus(popUpWindow)
                popUpWindow.type_keys('{ENTER}')

        def _clickPrint():
            setFocus(bankSelectWindow)
            bankSelectWindow.child_window(title='In',auto_id="btnPrint").click()

        def _takeFlexScreenshot(window):
            setFocus(window)
            return cv.cvtColor(np.array(window.capture_as_image()),cv.COLOR_RGB2BGR)

        def _processFlexImage(image):
            flexGrayImg = cv.cvtColor(image,cv.COLOR_RGB2GRAY)
            _, flexBinaryImg = cv.threshold(flexGrayImg,130,255,cv.THRESH_BINARY)
            return flexBinaryImg

        def _queryVCI1104():
            # dataDate = dt.datetime.now().strftime('%Y-%m-%d')
            dataDate = '2022-08-09'
            return pd.read_sql(
                f"""
                SELECT 
                    REPLACE([SoTaiKhoanLuuKy],'.','') [AccountNumber],
                    [TenKhachHang] [CustomerName],
                    [NganHangThuHuong] [ReceivingBank],
                    [SoTienChuyen] [AmountValue],
                    [ThoiGianGhiNhan] [RecordTime]
                -- FROM [VCI1104]
                FROM [VCI1104_UAT]
                -- WHERE [NgayLapChungTu] = '{dataDate}'
                ORDER BY [ThoiGianGhiNhan] DESC
                """,
                connect_DWH_CoSo,
            )

        def _getBankAutoID():
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

        def _selectBankFromList(bankAutoID):
            selectButton = bankSelectWindow.child_window(auto_id=bankAutoID)
            selectButton.click()

        def _exportPDF():
            setFocus(bankOrderWindow)
            bankOrderWindow.maximize()
            bankOrderWindow.click_input(coords=(14,68),absolute=True)

        def _findFileName():
            if cashOut >= 100e6 and bankName == 'VTB':
                return f'UP_VTB_{customerName.title()}_{recordTime.strftime("%H%M%S")}'
            else:
                return f'{bankName}_{customerName.title()}_{recordTime.strftime("%H%M%S")}'

        def _savePDF():
            # Chọn định dạng pdf
            setFocus(saveFileWindow)
            fileTypeBox = saveFileWindow.child_window(class_name='ComboBox',found_index=1)
            fileTypeBox.click()
            fileTypeBox.type_keys('Adobe Acrobat (*.pdf)')
            # Nhập tên file
            filePath = r'C:\Users\hiepdang\Shared Folder\Finance\UyNhiemChi'
            fileName = _findFileName()
            fileNameBox = saveFileWindow.child_window(class_name='ComboBox',found_index=0)
            fileNameBox.click()
            fileNameBox.type_keys('^a{DELETE}')
            fileNameBox.type_keys(f'{filePath}\\{fileName}.pdf',with_spaces=True)
            print(f'{filePath}\\{fileName}.pdf')
            # Click "Save"
            setFocus(saveFileWindow)
            saveButton = saveFileWindow.child_window(title="&Save",class_name="Button")
            saveButton.click()
            # Click "Ok" (Export Complete)
            saveFileWindow.type_keys('{ENTER}')

        def _quitPDFWindow():
            bankOrderWindow.close()

        _lastNumberOfOrders = 0
        while True:
            # Ghi nhận thời gian quét
            runTime = dt.datetime.now().strftime('%H:%M:%S')
            # Check xem có lệnh mới không
            _tableDatabase = _queryVCI1104()
            _currentNumberOfOrder = _tableDatabase.shape[0]
            if _currentNumberOfOrder <= _lastNumberOfOrders:  # không có lệnh mới
                logMessage = f'*** Lần quét tại {runTime} -- Không có lệnh mới\n--------------------'
                print(logMessage)
                time.sleep(60)
                continue
            # Nếu có lệnh mới -> Lấy các lệnh mới
            _increaseNumberOfOrder = _currentNumberOfOrder - _lastNumberOfOrders
            _tableDatabase = _tableDatabase.head(_increaseNumberOfOrder)
            _lastNumberOfOrders = _currentNumberOfOrde
            # Click "Tìm kiếm" để hiện các lệnh mới trên Flex
            _clickSearchButton()
            # Screen shot toàn bộ màn hình chức năng
            funcImage = _takeFlexScreenshot(self.funcWindow)
            # Screen shot khung dữ liệu
            self.dataWindow = self.funcWindow.child_window(auto_id='pnlSearchResult')
            dataImage = _takeFlexScreenshot(self.dataWindow)
            # Xử lý ảnh cột Tài khoản
            accountImage = _cropColumn('SoTaiKhoanLuuKy')
            accountBinaryImage = _processFlexImage(accountImage)
            # Đọc dữ liệu từ ảnh để tìm số dòng trên màn hình Flex
            _content = _readFlexDataImage(dataImage)
            _tableFlex = _getFlexDataFrame(_content)
            # Map tableFlex và tableDatabase
            table = pd.merge(_tableDatabase,_tableFlex,how='right',on=['AccountNumber','AmountValue'])
            table = table.set_index('RowNumber')
            table = table.sort_index(ascending=True)

            for _rowNumber in table.index:
                # Lấy các thông tin cơ bản
                cashOut = table.loc[_rowNumber,'AmountValue']
                rawBankName = table.loc[_rowNumber,'ReceivingBank']
                bankCode,bankName = _findBankCodeBankName(cashOut,rawBankName)
                customerName = table.loc[_rowNumber,'CustomerName']
                recordTime = table.loc[_rowNumber,'RecordTime']
                # Click chọn từng lệnh
                _clickFirstOrder()
                # Nhập bank code vào khung "Mã NH UNC"
                _sendBankCode()
                # Bấm thực hiện
                _clickExecuteButton()
                # Đóng pop up (nếu có)
                _closePopUp()
                # Chọn ngân hàng
                bankSelectWindow = self.app.window(title='In bảng kê')
                _bankAutoID = _getBankAutoID()
                _selectBankFromList(_bankAutoID)
                # Bấm "In"
                _clickPrint()
                # Export ủy nhiệm chi
                bankOrderWindow = self.app.window(title=_bankAutoID)
                _exportPDF()
                # Save ủy nhiệm chi
                saveFileWindow = self.app.window(title='Export Report')
                _savePDF()
                # Đóng PDF preview
                _quitPDFWindow()
                # Thông báo xử lý xong
                logMessage = f"""
                *** Lần quét tại {runTime} -- Đã xử lý lệnh ::
                - Tên khách hàng: {customerName}
                - Giá trị lệnh: {int(cashOut):,d}
                - Thời gian ghi nhận lệnh: {recordTime.strftime('%H:%M:%S')}
                --------------------
                """
                print(logMessage)



