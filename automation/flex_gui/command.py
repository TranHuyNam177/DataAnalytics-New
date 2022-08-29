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
from datawarehouse.DWH_CoSo import connect_DWH_CoSo, SYNC
import cv2 as cv
import loggins


class VCI1104(Flex):

    def __init__(self,env,username,password):
        super().__init__(env)
        self.start()
        self.login(username,password)
        self.insertFuncCode(self.__class__.__name__)
        self.dataWindow = None

    def pushCashOutOrders(self):

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
            return balanceTable.loc[balanceTable['AccountNumber']==accountNumber,'Balance']

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

        def _findRowHeight(binaryRecordImage):
            sumIntensity = binaryRecordImage.sum(axis=1)
            minLocArray = np.asarray(sumIntensity==sumIntensity.min()).nonzero()
            minLocArray = np.insert(minLocArray,0,0)
            rowHeightArray = np.diff(minLocArray)
            rowHeightArray = rowHeightArray[rowHeightArray<30] # row height ko thể lớn hơn 30 pixel
            rowHeight = np.median(rowHeightArray).astype(int)
            return rowHeight

        def _cropColumn(colName):
            topLeft, topRight = _findColumnCoords(colName)
            top = topLeft[0]
            left = topLeft[1]
            right = topRight[1]
            bottom = dataImage.shape[0]
            return dataImage[top:bottom,left:right,:]

        def _clickRow(rowNumber,rowHeight):
            topLeft, topRight = _findColumnCoords('Select')
            left = topLeft[1]
            right = topRight[1]
            top = topLeft[0]
            midPoint = (int(left*0.9+right*0.1),int(top+rowHeight/2*(2*rowNumber-1)))
            setFocus(self.funcWindow)
            self.dataWindow.click_input(coords=midPoint,absolute=False)

        def _readAccountNumber(image):
            content = pytesseract.pytesseract.image_to_string(
                image,
                config='--psm 4 tessedit_char_whitelist=0123456789C,.',
            )
            content = re.sub(r'\b022[,.].?[,.]','022C',content)
            content = re.sub(r'[^\dC]','',content)
            content = re.sub(r'022C',' 022C',content).strip()
            return content

        def _readAmountValue(image):
            content = pytesseract.pytesseract.image_to_string(
                image,
                config='--psm 4 tessedit_char_whitelist=0123456789,.',
            )
            content = re.sub(r'\n+',' ',content)
            content = re.sub(r'[^\d,. ]','',content)
            return content

        def _getFlexDataFrame(accountStrings,amountStrings):
            return pd.DataFrame(
                data={
                    'RowNumber':np.arange(len(amountStrings.split())) + 1,
                    'AccountNumber':accountStrings.split(),
                    'AmountValue':[float(string.replace(',','')) for string in amountStrings.split()],
                }
            )

        def _hasRecord(accountImage):
            accountBinaryImage = cv.cvtColor(accountImage,cv.COLOR_BGR2GRAY)
            trimmedImage = accountBinaryImage[:accountBinaryImage.shape[0]//2,5:-5] # conservative approach
            emptyPixels = (trimmedImage>250).sum()
            return emptyPixels / trimmedImage.size < 0.99 # có record -> True, không có record -> False

        def _mapData(accountImage,amountImage):
            dbTable = _queryVCI1104()
            thresholdRange = range(100,155,5) # xử lý ảnh với nhiều mức threshold với nhau
            table = None
            for threshold in thresholdRange:
                # Số tài khoản
                accountBinaryImage = _processFlexImage(accountImage,threshold)
                accounts = _readAccountNumber(accountBinaryImage)
                # Số tiền chuyển
                amountBinaryImage = _processFlexImage(amountImage,threshold)
                amounts = _readAmountValue(amountBinaryImage)
                flexTable  = _getFlexDataFrame(accounts,amounts)
                table = pd.merge(flexTable,dbTable,on=['AccountNumber','AmountValue'],how='left')
                if table['ReceivingBank'].notna().any(): # đọc được ít nhất 1 record
                    break
            table = table.loc[table['ReceivingBank'].notna()] # chỉ lấy thằng nào đọc được
            return table.head(1) # chỉ lấy 1 record

        def _clickSearchButton():
            setFocus(self.funcWindow)
            searchButtonWindow = self.funcWindow.child_window(auto_id='btnSearch')
            searchButtonWindow.click_input()

        def _sendBankCode(bankCode):
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
            actionWindow = self.funcWindow.child_window(title='SMS')
            actionImage = _takeFlexScreenshot(actionWindow)
            actionImage = actionImage[10:,:,:]
            unique,count = np.unique(actionImage,return_counts=True,axis=1)
            mostFrequentColumn = unique[:,np.argmax(count),:]
            columnMask = ~(actionImage==mostFrequentColumn[:,np.newaxis,:]).all(axis=(0,2))
            lastColumn = np.argwhere(columnMask).max()
            croppedImage = actionImage[:,:lastColumn,:]
            midPoint = croppedImage.shape[1]//2, croppedImage.shape[0]//2
            actionWindow.click_input(coords=midPoint)

        def _closePopUp():
            popUpWindow = self.app.window(title='Nhập giao dịch')
            if popUpWindow.exists():
                setFocus(popUpWindow)
                popUpWindow.type_keys('{ENTER}')

        def _clickPrint():
            setFocus(bankSelectWindow)
            bankSelectWindow.child_window(title='In',auto_id="btnPrint").click()

        def _takeFlexScreenshot(window):
            setFocus(self.funcWindow)
            return cv.cvtColor(np.array(window.capture_as_image()),cv.COLOR_RGB2BGR)

        def _processFlexImage(image,threshold):
            flexGrayImg = cv.cvtColor(image,cv.COLOR_RGB2GRAY)
            _, flexBinaryImg = cv.threshold(flexGrayImg,threshold,255,cv.THRESH_BINARY)
            return flexGrayImg

        def _queryVCI1104():
            now = dt.datetime.now()
            past = now - dt.timedelta(minutes=30)
            dataDate = now.strftime('%Y-%m-%d')
            fromTime = past.strftime('%Y-%m-%d %H:%M:%S')
            # SYNC(connect_DWH_CoSo,'spVCI1104_UAT',dataDate,dataDate)
            SYNC(connect_DWH_CoSo,'spVCI1104',dataDate,dataDate)
            return pd.read_sql(
                f"""
                SELECT 
                    REPLACE([SoTaiKhoanLuuKy],'.','') [AccountNumber],
                    [TenKhachHang] [CustomerName],
                    [NganHangThuHuong] [ReceivingBank],
                    [SoTienChuyen] [AmountValue],
                    [ThoiGianGhiNhan] [RecordTime]
                FROM [VCI1104]
                -- FROM [VCI1104_UAT]
                WHERE [ThoiGianGhiNhan] > '{fromTime}'
                ORDER BY [ThoiGianGhiNhan] -- Lệnh vào trước đẩy trước
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

        def _clickExportPDF(bankOrderWindow):
            setFocus(bankOrderWindow)
            bankOrderWindow.maximize()
            menuBar = bankOrderWindow.children()[-1]
            menuImage = np.array(menuBar.capture_as_image())
            menuImage = menuImage[:,:-10,:] # bỏ 10 pixel cuối
            menuImage = cv.cvtColor(menuImage,cv.COLOR_RGB2BGR)

            unique,count = np.unique(menuImage,return_counts=True,axis=1)
            mostFrequentColumn = unique[:,np.argmax(count),:]
            columnMask = ~(menuImage==mostFrequentColumn[:,np.newaxis,:]).all(axis=0)
            lastColumn = np.argwhere(columnMask).max()
            croppedImage = menuImage[:,:lastColumn,:]
            midPoint = croppedImage.shape[1]//20, croppedImage.shape[0]//2
            # (10 icon, chia thêm 2 để lấy midPoint -> chia 20)
            menuBar.click_input(coords=midPoint,absolute=False)

        def _findFileName():
            nowString = dt.datetime.now().strftime('%H%M%S')
            if cashOut >= 100e6 and bankName == 'VTB':
                return f'UP_VTB_{customerName.title()}_{recordTime.strftime("%H%M%S")}_{nowString}'
            else:
                return f'{bankName}_{customerName.title()}_{recordTime.strftime("%H%M%S")}_{nowString}'

        def _savePDF(saveFileWindow):
            # Chọn định dạng pdf
            setFocus(saveFileWindow)
            fileTypeBox = saveFileWindow.child_window(class_name='ComboBox',found_index=1)
            fileTypeBox.click()
            fileTypeBox.type_keys('Adobe Acrobat (*.pdf)')
            # Nhập tên file
            filePath = r'\\192.168.8.63\Finance\UyNhiemChi'
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
            saveFileWindow.wait('visible',timeout=30)
            saveFileWindow.type_keys('{ENTER}')

        def _quitPDFWindow():
            bankOrderWindow.close()

        while True:
            # Click "Tìm kiếm" để kiểm tra các lệnh mới trên Flex
            _clickSearchButton()
            # Screen shot khung dữ liệu
            self.dataWindow = self.funcWindow.child_window(auto_id='pnlSearchResult')
            dataImage = _takeFlexScreenshot(self.dataWindow)
            # Check có dữ liệu không
            accountImage = _cropColumn('SoTaiKhoanLuuKy')
            amountImage = _cropColumn('SoTienChuyen')
            if not _hasRecord(accountImage):
                time.sleep(15)
                continue
            # Đọc dữ liệu
            table = _mapData(accountImage,amountImage)
            table = table.set_index('RowNumber')
            table = table.sort_index(ascending=True)
            # Tìm Row Height
            accountBinaryImage = _processFlexImage(accountImage,100)
            rowHeight = _findRowHeight(accountBinaryImage)

            for _rowNumber in table.index:
                # Lấy các thông tin cơ bản
                cashOut = table.loc[_rowNumber,'AmountValue']
                rawBankName = table.loc[_rowNumber,'ReceivingBank']
                bankCode, bankName = _findBankCodeBankName(cashOut,rawBankName)
                customerName = table.loc[_rowNumber,'CustomerName']
                recordTime = table.loc[_rowNumber,'RecordTime']
                # Click chọn từng lệnh
                _clickRow(_rowNumber,rowHeight)
                # Nhập bank code vào khung "Mã NH UNC"
                _sendBankCode(bankCode)
                # Bấm thực hiện
                _clickExecuteButton()
                # Đóng pop up (nếu có)
                _closePopUp()
                # Chọn ngân hàng
                bankSelectWindow = self.app.window(title='In bảng kê')
                bankSelectWindow.wait('visible',timeout=30)
                _bankAutoID = _getBankAutoID()
                _selectBankFromList(_bankAutoID)
                # Bấm "In"
                _clickPrint()
                # Export ủy nhiệm chi
                bankOrderWindow = self.app.window(title=_bankAutoID)
                bankOrderWindow.wait('visible',timeout=30)
                _clickExportPDF(bankOrderWindow)
                # Save ủy nhiệm chi
                saveFileWindow = self.app.window(title='Export Report')
                saveFileWindow.wait('visible',timeout=30)
                _savePDF(saveFileWindow)
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
                time.sleep(1) # nghỉ 1s sau mỗi lệnh
                break

