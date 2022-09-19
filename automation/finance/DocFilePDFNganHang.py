from automation import *
import cv2 as cv
import warnings

warnings.filterwarnings("ignore", 'This pattern has match groups')

"""
--------- RULES ---------
1. Số hợp đồng (ContractNumber):
    - Nếu có "Contract number/ Contract No." trong file thì lấy cái này làm contract number
    - Nếu có Ref No., Ref, Loan Ref thì lấy cái này làm contract number
    - Nếu có TXNT no. thì lấy cái này làm contract number
    * Note: tất cả các Bank, nếu không tìm thấy số hợp đồng trên PDF tương ứng với số hợp đồng trong file excel
    thì để số hợp đồng trống, bên Finance tự điền
2. Khi nào các ngân hàng có mẫu PDF mà bên PHS đã trả các khoản vay thì bên Finance thông báo
3. Theo dõi file loan nước ngoài từ trước tới giờ thì không có khoản vay nào thấp hơn 1,000,000
4. Các ngân hàng OBU (scan file PDF) trừ ESUN, các ngân hàng còn lại chỉ có 1 khoản vay trong mỗi file từng tháng
5. Rule của từng bank:
    5.1 CATHAY:
        - Nếu có CAP -> Paid = Amount
        - Tối đa 1 file sẽ có 2 dòng với INT nằm ở đầu dòng
        - Nếu có CAP và 2 INT thì INT gần nhất (nằm ngay phía trên CAP) -> dòng INT đó Paid sẽ bằng Amount, còn Int ở
        trên cùng paid = 0
    5.2 SHHK:
        - Trước tới giờ 1 loan account chỉ có đúng 1 khoản vay (Bảng đầu tiên trong file chỉ có 1 dòng)
        - Số hợp đồng (Loan Ref.) 6 kí tự đầu '888LN' cố định
    5.3 ESUN:
        - Số hợp đồng: 6 kí tự đầu cố định
        - Trang để lấy xử lý dữ liệu là trang có 'Payment Advice'
        - Amount lấy từ Loan Outstanding
    5.4 ENTIE:
        - Amount lấy từ Loan O/S Amount
        - Loan O/S Amount tương dương principal payable trong file
        - Contract number có thể bỏ USD1
        - 1 file có 2 page giống nhau, page nào PRINCIPAL PAYABLE khác 0 thì đọc
    5.5 UBOT:
        - contract number lấy Contract No.
        - Anh Duy: công ty có trả vay rồi nhưng không có file PDF, chỉ có 1 file tờ đã thanh toán thôi
    5.6 TAISHIN:
         - File tháng 1 có bảng gồm cả 3 tháng vì bên cty mình yêu cầu thống kê lãi vay 3 tháng gần nhất
         - Bên Finance sẽ báo cho bên TAISHIN tháng nào gửi file tháng đó, không gửi gộp
    5.7 YUANTA:
        - contract number: 4 chữ cái đầu set cứng
"""

def convertPDFtoImage(bankName: str, month: int):
    """
    Hàm chuyển từng trang trong PDF sang image
    """
    directory = join(realpath(dirname(__file__)), 'bank', f'THÁNG {month}')
    pathPDF = join(directory, f'{bankName}.pdf')
    images = convert_from_path(
        pdf_path=pathPDF,
        poppler_path=r'C:\Users\namtran\poppler-0.68.0\bin'
    )
    return images


def detect_table(inputImage):
    """
    Hàm nhận diện bảng trong 1 image
    """
    thresh = cv2.adaptiveThreshold(inputImage, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 51, 9)
    # Fill rectangular contours
    cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    for c in cnts:
        cv2.drawContours(thresh, [c], -1, (255, 255, 255), -1)
    # Morph open
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    opening = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=4)
    # Draw rectangles
    cnts = cv2.findContours(opening, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    img_list = []
    for c in reversed(cnts):  # dùng reversed chỗ này vì hình được đọc từ dưới lên trên
        x, y, w, h = cv2.boundingRect(c)
        if x > w and y > h:
            continue
        roi = inputImage[y:y + h, x:x + w]
        img_list.append(roi)
    return img_list  # return ra list chứa image là table (đã scale về array 2 chiều)


def readImgPytesseractToString(img, numConfig: int):
    """
    Hàm trả về 1 text theo dạng chuỗi từ image sử dụng pytesseract
    """
    string = pytesseract.image_to_string(
        image=img,
        config=f'--psm {numConfig}'
    )
    return string


def readImgPytesseractToDataframe(img, numConfig: int):
    df = pytesseract.image_to_data(
        image=img,
        config=f'--psm {numConfig}',
        output_type='data.frame'
    )
    df = df[df['conf'] != -1].reset_index(drop=True)
    return df

def groupByDataFrame(df):
    """
    Hàm trả về 1 dataframe sau khi đã groupby 2 cột text và conf
    """
    df['text'] = df['text'].astype(str)
    # group by dataframe theo block_num
    groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
    groupByConf = df.groupby(['block_num'])['conf'].mean()
    dfGroupBy = pd.concat([groupByText, groupByConf], axis=1)
    return dfGroupBy

def checkPatternRegexInDataFrame(df, dictKeys, confidence: int):
    if df.empty:
        return pd.DataFrame()
    df['text'] = df['text'].astype(str)
    df['check'] = df['text'].str.contains(dictKeys)
    df = df.loc[(df['check']) & (df['conf'] >= confidence)].copy()
    if df.empty:
        df = pd.DataFrame()
    else:
        df['text'] = df['text'].apply(lambda x: x.replace('..', '.').replace('.,', ','))
        df['regex'] = df['text'].str.extract(dictKeys)
    return df

def erosionAndDilation(img, f: str, array: int):
    kernel = np.ones((array, array), np.uint8)
    if f == 'e':
        newImg = cv2.erode(img, kernel, iterations=1)
    else:
        newImg = cv2.dilate(img, kernel, iterations=1)
    return newImg

def _findTopLeftPoint(pdfImage, smallImage):
    pdfImage = np.array(pdfImage)  # đảm bảo image được đưa về numpy array
    smallImage = np.array(smallImage)  # đảm bảo image được đưa về numpy array
    matchResult = cv.matchTemplate(pdfImage, smallImage, cv.TM_CCOEFF)
    _, _, _, topLeft = cv.minMaxLoc(matchResult)
    return topLeft[0], topLeft[1]  # cho compatible với openCV

def _findTopLeftPointCATHAY(pdfImage, smallImage):
    pdfImage = np.array(pdfImage)  # đảm bảo image để được đưa về numpy array
    smallImage = np.array(smallImage)  # đảm bảo image để được đưa về numpy array

    matchResult = cv.matchTemplate(pdfImage, smallImage, cv.TM_SQDIFF)
    # get all the matches:
    matchResult2 = np.reshape(matchResult, matchResult.shape[0] * matchResult.shape[1])
    sort = np.argsort(matchResult2)
    (y1, x1) = np.unravel_index(sort[0], matchResult.shape)  # best match
    (y2, x2) = np.unravel_index(sort[1], matchResult.shape)  # second best match
    if y1 - y2 > 5:
        topLeftList = [(x2, y2), (x1, y1)]
    elif y2 - y1 > 5:
        topLeftList = [(x1, y1), (x2, y2)]
    else:
        topLeftList = [(x1, y1)]
    return topLeftList

def _findCoords(pdfImage, name, bank):
    if (bank == 'ESUN' and name in ('interestRate', 'date')) or (bank == 'MEGA' and name in ('date', 'paid')) or (bank == 'ENTIE' and name in ('contractNumber', 'interestRate')):
        fileName = 'values.png'
    elif bank == 'FUBON' and name in ('issueDate', 'expireDate', 'amount', 'interestRate', 'paid', 'contractNumber'):
        fileName = 'values.png'
    elif name == 'amount':
        fileName = 'amount.png'
    elif name == 'interestRate':
        fileName = 'interestRate.png'
    elif name == 'date':
        fileName = 'date.png'
    elif name == 'issueDate':
        fileName = 'issueDate.png'
    elif name == 'expireDate':
        fileName = 'expireDate.png'
    elif name == 'contractNumber':
        fileName = 'contractNumber.png'
    elif name == 'paid':
        fileName = 'paid.png'
    elif name == 'condition':
        fileName = 'condition.png'
    elif name == 'values':
        fileName = 'values.png'
    else:
        raise ValueError(
            'colName must be either "amount" or "interestRate" or "date" or "contractNumber" or "paid" or "issueDate" '
            'or "expireDate" or "condition" or "values"'
        )
    smallImagePath = os.path.join(os.path.dirname(__file__), 'bank_img', f'{bank}', fileName)
    smallImage = cv.imread(smallImagePath, 0)  # hình trắng đen (array 2 chiều)
    w, h = smallImage.shape[::-1]
    top, left = _findTopLeftPoint(pdfImage, smallImage)
    if bank == 'MEGA':
        if name == 'date':
            right = int(left + h * 2.1 - 5)
            left = int(left + 1.7 * h)
            bottom = int(top + 0.58 * w)
            top = top + 20
        elif name == 'amount':
            right = int(left + h * 2.2)
            left = int(left + 1.2 * h)
            bottom = top + w + 15
            top = int(top + w * 0.66)
        elif name == 'interestRate':
            right = int(left + h * 2.3 - 10)
            left = int(left + 1.2 * h)
            bottom = int(top + w / 2)
            top = int(top + 0.1 * w)
        elif name == 'contractNumber':
            right = left + h + 7
            left = int(left + 2 * h / 3 - 15)
            top = top + w - 5
            bottom = int(top + w * 0.6)
        else:
            right = int(left + 0.7 * h)
            left = int(left + h / 3 - 5)
            top = int(top + 0.6 * w)
            bottom = top * 2 + 10
        return pdfImage[left:right, top:bottom]
    elif bank == 'SHINKONG':
        if name == 'contractNumber':
            top = top + w - 5
            right = left + h + 5
            left = left - 5
        elif name == 'date':
            right = left + h + 5
            left = left - 7
            top = top + w - 7
        else:
            top = top + w - 5
            right = left + h + 5
        return pdfImage[left:right, top:]
    elif bank == 'YUANTA':
        if name == 'contractNumber':
            top = int(top + w * 1.4 + 5)
            left = int(left + 1.29 * h - 3)
            right = int(left + h / 3 - 10)
            bottom = None
        else:
            top = top + w - 5
            right = left + h + 5
            bottom = None
        return pdfImage[left:right, top:bottom]
    elif bank == 'SHHK':
        if name == 'date':
            bottom = top + w * 2
            right = left + h * 2 + 5
            left = left + h
        elif name == 'paid':
            top = top + w
            right = left + h
            bottom = None
        else:
            top = top - 5
            bottom = top + w
            right = left + h * 2 + 5
            left = left + h
        return pdfImage[left:right, top:bottom]
    elif bank == 'SINOPAC':
        if name in ('contractNumber', 'paid'):
            top = top + w - 5
            bottom = top + w
            right = left + h
        else:
            bottom = top + w + 10
            top = top - 24
            right = left + h * 3 + 5
            left = left + h
        return pdfImage[left:right, top:bottom]
    elif bank == 'CHANG HWA':
        if name == 'interestRate':
            top = top + w - 5
            bottom = int(top + 2 * w / 3)
            right = left + h
        else:
            top = top + w - 5
            right = left + h
            bottom = None
        return pdfImage[left:right, top:bottom]
    elif bank == 'KGI':
        top = top - 5
        bottom = top + w + 15
        right = left + h * 2 + 5
        left = left + h
        return pdfImage[left:right, top:bottom]
    elif bank == 'ESUN':
        if name == 'amount':
            top = top + w - 20
            right = int(left + h / 3 - 5)
            left = left - 2
        elif name == 'interestRate':
            top = int(top + w * 3.9)
            right = int(left + h / 2)
            left = left - 10
        elif name == 'date':
            top = top + w
            right = left + h
            left = int(left + h / 2 - 10)
        elif name == 'paid':
            top = top + w - 5
            right = int(left + 0.4 * h - 4)
            left = left - 5
        else:
            top = top + w - 110
            right = left + h + 3
            left = left - 5
        return pdfImage[left:right, top:]
    elif bank == 'BOP':
        bottom = top + w
        right = left + h * 2
        left = left + h
        return pdfImage[left:right, top:bottom]
    elif bank == 'FIRST':
        if name in ('amount', 'date', 'interestRate', 'paid'):
            bottom = top + w
            right = left + h * 2
            left = left + h + 10
        else:
            top = top + w + 10
            bottom = top + w
            right = left + h
        return pdfImage[left:right, top:bottom]
    elif bank == 'UBOT':
        if name == 'date':
            top = top + w
            right = left + h + 5
            left = left - 5
        elif name == 'contractNumber':
            right = left + h * 2
            left = left + h - 10
        elif name == 'interestRate':
            top = top + w
            right = left + h + 5
            left = left - 3
        else:
            top = top + w
            right = left + h + 5
        return pdfImage[left:right, top:]
    elif bank == 'TCB':
        if name in ('contractNumber', 'interestRate', 'date'):
            top = top + w
            right = left + h
            bottom = None
        else:
            top = top + w
            bottom = top + w + 30
            right = left + h
        return pdfImage[left:right, top:bottom]
    elif bank == 'ENTIE':
        if name == 'amount':
            top = top + w
            bottom = top + w + 50
            right = left + h
            left = int(left + 2 * h / 3)
        elif name == 'contractNumber':
            top = top + w
            bottom = top + w
            right = int(left + 0.7 * h + 5)
            left = int(left + h / 3 - 7)
        elif name == 'date':
            bottom = None
            top = top + w + 10
            right = int(left + 0.7 * h - 2)
            left = int(left + 0.22 * h)
        elif name == 'interestRate':
            top = top + w
            bottom = top + w - 100
            right = left + h
            left = int(left + 0.7 * h)
        else:
            top = top + w * 2
            bottom = None
            right = int(left + 0.7 * h - 5)
            left = int(left + h / 3 + 5)
        return pdfImage[left:right, top:bottom]
    elif bank == 'TAISHIN':
        if name == 'values':
            top = int(top + 0.27 * w)
            left = left + h - 20
            right = left + h + 30
        else:
            right = left - 10
            left = int(left - h / 2)
            top = int(top + 0.32 * w)
        return pdfImage[left:right, top:]
    elif bank == 'FUBON':
        if name == 'issueDate':
            bottom = int(top + w * 0.3)
            top = int(top + w * 0.2)
            right = int(left + h * 1.5 + 19)
            left = left + h
        elif name == 'expireDate':
            bottom = int(top + w * 0.4)
            top = int(top + w * 0.3)
            right = int(left + h * 1.5 + 19)
            left = left + h
        elif name == 'amount':
            bottom = int(top + w * 0.55)
            top = int(top + w * 0.4)
            right = int(left + h * 1.5 + 19)
            left = left + h - 5
        elif name == 'contractNumber':
            bottom = int(top + 0.4 * w)
            top = int(top + 0.14 * w)
            left = int(left - 2.3 * h)
            right = left + h + 17
        elif name == 'paid':
            bottom = top + w + 10
            top = int(top + w * 0.8)
            right = int(left + h * 4 + 17)
            left = left + h * 2
        else:  # interestRate
            bottom = int(top + w * 0.73)
            top = int(top + w * 0.63 - 10)
            right = int(left + h * 1.5 + 19)
            left = left + h - 7
        return pdfImage[left:right, top:bottom]
    else:  # CATHAY
        right = int(left + h / 2)
        left = int(left - h / 4)
        bottom = top + w
        top = int(top + 0.3 * w)
        return pdfImage[left:right, top:bottom]

def runMEGA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    if bank != 'MEGA':
        bank = 'MEGA'
    patternDict = {
        'amount': r'(\d+[,.][\d,\d.]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[~-]+\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{14})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})',
    }
    for img in images:
        # scale image to gray
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        check = 'INTEREST PAYMENT NOTICE' not in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 6)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '').replace('.', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            imgDate = cv2.resize(imgDate, (405, 55))
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            dfDate = groupByDataFrame(dfDate)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            if '~' in dateText:
                issueDateText, expireDateText = dateText.split('~')
            else:
                issueDateText, expireDateText = dateText.split('-')
            issueDate = dt.datetime.strptime(issueDateText.replace('-', ''), '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText.replace('-', ''), '%Y/%m/%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays/30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            imgContractNumber = cv.GaussianBlur(imgContractNumber, (3, 3), 0)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 0)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            dfPaid = readImgPytesseractToDataframe(imgPaid, 6)
            dfPaid = checkPatternRegexInDataFrame(dfPaid, patternDict['paid'], 10)
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', ''))
                paid = round(paid - interestAmount, 2)
            # remaining
            remaining = amount - paid

            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

def runSHINKONG(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]\d[A-Z]{7}\d{5}|[A-Z]{9}\d{5})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})',
    }
    for img in images:
        # scale image to gray
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # using opening morphological
        kernel = np.ones((3, 3), np.uint8)
        fullImageScale = cv2.morphologyEx(fullImageScale, cv2.MORPH_OPEN, kernel)
        # check image with condition
        if 'Payment Advice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            if '~' in dateText:
                issueDateText, expireDateText = dateText.split('~')
            else:
                issueDateText, expireDateText = dateText.split('-')
            issueDate = dt.datetime.strptime(issueDateText.replace('-', ''), '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText.replace('-', ''), '%Y/%m/%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            if contractNumber.startswith('LI'):
                contractNumber = contractNumber.replace('LI', 'L1')
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            dfPaid = readImgPytesseractToDataframe(imgPaid, 11)
            dfPaid = checkPatternRegexInDataFrame(dfPaid, patternDict['paid'], 10)
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', ''))
                paid = round(paid - interestAmount, 2)
            # remaining
            remaining = amount - paid

            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runYUANTA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}-\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{4}[A-Z]{2}\d{5})'
    }
    images = convertPDFtoImage(bank, month)

    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        if 'Interest Payment Notice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # detect table in image
            tables = detect_table(fullImageScale)
            table = tables[0]  # vì chỉ có 1 bảng trong hình nên phần tử đầu tiên luôn là bảng đó
            # amount
            imgAmount = _findCoords(np.array(table), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(table), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(table), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 4)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            issueDateText, expireDateText = dateText.split('-')
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            contractNumber = 'AOBU' + contractNumber
            # paid
            paid = 0
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# Done
def runSHHK(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[-~]\d{4}/\d{2}/\d{2}|\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{2}\d{7})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})'
    }

    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            elif dfDate['check'].count() == 2:
                issueDateText = dfDate.loc[dfDate.index[0], 'regex']
                expireDateText = dfDate.loc[dfDate.index[1], 'regex']
                issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
                expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            else:
                dateText = dfDate.loc[dfDate.index[0], 'regex']
                if '~' in dateText:
                    issueDateText, expireDateText = dateText.split('~')
                else:
                    issueDateText, expireDateText = dateText.split('-')
                issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
                expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            contractNumber = contractNumber.replace(contractNumber[0:2], '888LN')
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            dfPaid = checkPatternRegexInDataFrame(dfPaid, patternDict['paid'], 10)
            if dfPaid.empty:
                paid = 0
            else:
                paid = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paid.replace(',', ''))
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# Done
def runSINOPAC(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{1,2}[.,;:]?[A-Z]{3}[.,;:]?\d{4}to\d{1,2}[.,;:]?[A-Z]{3}[.,;:]?\d{4})',
        'contractNumber': r'(\d{8})',
        'paid': r'(\d+[,][\d,]+\d{3}\.\d{2})'
    }
    if bank != 'SINOPAC':
        bank = 'SINOPAC'
    for img in images:
        img = img.rotate(-90, expand=1)
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)

            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            dfDate = groupByDataFrame(dfDate)
            dfDate['text'] = dfDate['text'].apply(lambda x: x.replace('..', '.'))
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            char = ',.:;'
            for c in char:
                if c in dateText:
                    dateText = dateText.replace(c, '')
            issueDateText, expireDateText = dateText.split('to')
            issueDate = dt.datetime.strptime(issueDateText, '%d%b%Y')
            expireDate = dt.datetime.strptime(expireDateText, '%d%b%Y')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # paid
            imgPaid = _findCoords(fullImageScale, 'paid', bank)
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            dfPaid = checkPatternRegexInDataFrame(dfPaid, patternDict['paid'], 10)
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',',''))
                paid = round(paid - interestAmount, 2)
            # remaining
            remaining = amount - paid
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# Done
def runCHANGHWA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[.,]\d+%)',
        'date': r'(\d{1,2}[A-Z]{3}\.\d{4}TO\d{1,2}[A-Z]{3}\.\d{4})',
        'contractNumber': r'(\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3})',
        'paid': r'(\d+,[\d,]+\d{3})'
    }
    if bank != 'CHANG HWA':
        bank = 'CHANG HWA'
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        lst_condition = ['INTEREST PAYMENT NOTICE', 'PAYMENT NOTICE']
        check = all(c not in readImgPytesseractToString(fullImageScale, 6) for c in lst_condition)
        if check:
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 6)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            dfDate = groupByDataFrame(dfDate)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            dateText = dateText.replace('.', '')
            issueDateText, expireDateText = dateText.split('TO')
            issueDate = dt.datetime.strptime(issueDateText, '%d%b%Y')
            expireDate = dt.datetime.strptime(expireDateText, '%d%b%Y')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            imgPaid = _findCoords(fullImageScale, 'paid', bank)
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            dfPaid = checkPatternRegexInDataFrame(dfPaid, patternDict['paid'], 10)
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', ''))
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runKGI(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[.,]\d+)',
        'date': r'(\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{3}\d{7}[A-Z]{2})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        if 'NOTICE OF LOAN REPAYMENT' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 0)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            if dfDate.loc[dfDate.index[0], 'regex'] < dfDate.loc[dfDate.index[1], 'regex']:
                issueDateText = dfDate.loc[dfDate.index[0], 'regex']
                expireDateText = dfDate.loc[dfDate.index[1], 'regex']
            else:
                issueDateText = dfDate.loc[dfDate.index[1], 'regex']
                expireDateText = dfDate.loc[dfDate.index[0], 'regex']
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            # Term Days
            termDays = (expireDate - issueDate).days + 1
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            paid = 0
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# Done
def runESUN(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    patternDict = {
        'amount': r'(\d+[,.][\d,\d.]+\d{3})',
        'interestRate': r'(\d+[,.]+\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[~-]\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{6})',
        'paid': r'(\d+[,.][\d,\d.]+\d{3})'
    }
    images = convertPDFtoImage(bank, month)
    if bank != 'ESUN':
        bank = 'ESUN'

    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        if 'Payment Advice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '').replace('.', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 4)
            dfDate = groupByDataFrame(dfDate)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            if '~' in dateText:
                issueDateText, expireDateText = dateText.split('~')
            else:
                issueDateText, expireDateText = dateText.split('-')
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            contractNumber = '22OBLN' + contractNumber
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 11)
            dfPaid = checkPatternRegexInDataFrame(dfPaid, patternDict['paid'], 10)
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', '').replace('.', ''))
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# viết lại CATHAY (có sử dụng opencv):
def runCATHAY(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[.,]+\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{1}[A-Z]{7}\d{7})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        conditionList = ['(LN4030)', 'N4030']
        check = any(c in readImgPytesseractToString(fullImageScale, 6) for c in conditionList)
        if not check:
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            imgContractNumber = erosionAndDilation(imgContractNumber, 'e', 3)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # các trường thông tin khác
            containingPath = os.path.join(os.path.dirname(__file__), 'bank_img', f'{bank}', 'check.png')
            containingImage = cv.imread(containingPath, 0)  # hình trắng đen (array 2 chiều)
            w, h = containingImage.shape[::-1]
            topLeftList = _findTopLeftPointCATHAY(containingImage, fullImageScale)
            for topLeft in topLeftList:
                top, left = topLeft
                # Ngày hiệu lực, ngày đáo hạn
                topDate = top + w * 2 - 10
                bottomDate = top + w * 8 + 10
                rightDate = left + h + 7
                # Date image
                imgDate = fullImageScale[left:rightDate, topDate:bottomDate]
                dfDate = readImgPytesseractToDataframe(imgDate, 11)
                dfDate = groupByDataFrame(dfDate)
                dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
                if dfDate.empty:
                    continue
                dateText = dfDate.loc[dfDate.index[0], 'regex']
                if '~' in dateText:
                    issueDateText, expireDateText = dateText.split('~')
                else:
                    issueDateText, expireDateText = dateText.split('-')
                issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
                expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
                if issueDate > expireDate:
                    continue
                # Term Days
                termDays = (expireDate - issueDate).days
                # Term Months
                termMonths = round(termDays / 30)
                # amount
                topAmount = top + w * 8
                bottomAmount = top + w * 14 - 44
                rightAmount = left + h + 6
                leftAmount = left - 3
                # Amount image
                imgAmount = fullImageScale[leftAmount:rightAmount, topAmount:bottomAmount]
                imgAmount = erosionAndDilation(imgAmount, 'e', 3)
                dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
                dfAmount = groupByDataFrame(dfAmount)
                dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 0)
                if dfAmount.empty:
                    continue
                amountText = dfAmount.loc[dfAmount.index[0], 'regex']
                amount = float(amountText.replace(',', ''))
                # interest rate
                bottomInterestRate = top + w * 20
                topInterestRate = top + w * 14
                rightInterestRate = int(left + h + 6.5)
                # interest image
                imgInterestRate = fullImageScale[left:rightInterestRate, topInterestRate:bottomInterestRate]
                dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
                dfInterestRate = groupByDataFrame(dfInterestRate)
                dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
                if dfInterestRate.empty:
                    continue
                interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
                interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
                # interest amount
                interestAmount = amount * termDays * interestRate / 360
                # paid
                bottomPaid = top + w
                leftPaid = left + h
                rightPaid = left + h * 2
                imgPaid = fullImageScale[leftPaid:rightPaid, top:bottomPaid]
                paidText = readImgPytesseractToString(imgPaid, 6)
                if 'CAP' in paidText:
                    paid = amount
                else:
                    paid = 0
                # remaining
                remaining = amount - paid
                # Append data
                records.append([contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount])
        except (Exception,):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# Mới có file của tháng 8 nên chưa có nhiều data để test, không thấy contractNumber trong file nên để là None
def runBOP(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}\d{2}\d{2})'
    }
    for img in images:
        img = img.rotate(-90, expand=1)
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        if 'Interest Rate Notice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            # ngày hiệu lực
            imgIssueDate = _findCoords(np.array(fullImageScale), 'issueDate', bank)
            dfIssueDate = readImgPytesseractToDataframe(imgIssueDate, 4)
            dfIssueDate = checkPatternRegexInDataFrame(dfIssueDate, patternDict['date'], 10)
            if dfIssueDate.empty:
                continue
            issueDateText = dfIssueDate.loc[dfIssueDate.index[0], 'regex']
            issueDate = dt.datetime.strptime(issueDateText, '%Y%m%d')
            # ngày đáo hạn
            imgExpireDate = _findCoords(np.array(fullImageScale), 'expireDate', bank)
            dfExpireDate = readImgPytesseractToDataframe(imgExpireDate, 4)
            dfExpireDate = checkPatternRegexInDataFrame(dfExpireDate, patternDict['date'], 10)
            if dfExpireDate.empty:
                continue
            expireDateText = dfExpireDate.loc[dfExpireDate.index[0], 'regex']
            expireDate = dt.datetime.strptime(expireDateText, '%Y%m%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            contractNumber = None
            # paid
            paid = 0
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# Done
def runFIRST(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}\.\d{1,2}\.\d{1,2}~\d{4}\.\d{1,2}\.\d{1,2})',
        'contractNumber': r'([A-Z]\d[A-Z]\d{7})',
        'paid': r'(\d+,[\d,]+\d{3})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        check = 'LOAN REPAYMENT CONFIRMATION' in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            dfDate = groupByDataFrame(dfDate)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            issueDateText, expireDateText = dateText.split('~')
            issueDate = dt.datetime.strptime(issueDateText, '%Y.%m.%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y.%m.%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            imgPaid = _findCoords(fullImageScale, 'paid', bank)
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            dfPaid = checkPatternRegexInDataFrame(dfPaid, patternDict['paid'], 10)
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', ''))
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# Không có file trả khoản vay, chỉ có file thanh toán đã trả -> chưa xác định được paid -> để paid = 0
def runUBOT(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}\.\d{1,2}\.\d{1,2}-\d{4}\.\d{1,2}\.\d{1,2})',
        'contractNumber': r'(\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3}[-]+[A-Z]{5})'
    }
    for img in images:
        img = img.rotate(-90, expand=1)
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()

        # check image with condition
        lst_condition = ['Interest Payment Notice', 'Payment Notice', 'Interest Payment']
        check = all(c not in readImgPytesseractToString(fullImageScale, 6) for c in lst_condition)
        if check:
            continue
        try:
            # amount & currency
            imgAmountCurrency = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmountCurrency, 4)
            dfAmount = groupByDataFrame(dfAmount)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 0)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate = groupByDataFrame(dfInterestRate)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            imgDate = cv2.resize(imgDate, (500, 32), interpolation=cv2.INTER_NEAREST)
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            dfDate = groupByDataFrame(dfDate)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            issueDateText, expireDateText = dateText.split('-')
            issueDate = dt.datetime.strptime(issueDateText, '%Y.%m.%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y.%m.%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays/30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            paid = 0
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# dữ liệu ít, mới chỉ có file 1 tháng, chưa có file trả vay nên tạm set paid = 0
def runTCB(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{2}/\d{2}/\d{4}-\d{2}/\d{2}/\d{4})',
        'contractNumber': r'([A-Z]{4}\d[A-Z]{2}\d{5})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check image with condition
        check = 'INTEREST PAYMENT NOTICE' not in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 0)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            dfDate = readImgPytesseractToDataframe(imgDate, 4)
            dfDate = groupByDataFrame(dfDate)
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            issueDateText, expireDateText = dateText.split('-')
            issueDate = dt.datetime.strptime(issueDateText, '%d/%m/%Y')
            expireDate = dt.datetime.strptime(expireDateText, '%d/%m/%Y')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            paid = 0
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runENTIE(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+[,.][\d,\d.]+\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}TO\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{11}[-]+\d{1,2})',
        'termDays': r'(\d{1,2}[A-Z]{4})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check condition in page to read
        imgCondition = _findCoords(np.array(fullImageScale), 'condition', bank)
        condition = readImgPytesseractToString(imgCondition, 11)
        condition = re.sub(r'[:A-Z\n\s\']', '', condition)
        if len(condition) < 12:
            Image.fromarray(fullImageScale).show()
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 6)
            dfAmount = groupByDataFrame(dfAmount)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 0)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '').replace('.', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            imgInterestRate = cv.morphologyEx(imgInterestRate, cv.MORPH_OPEN, (3, 3))
            imgInterestRate = cv.GaussianBlur(imgInterestRate, (5, 5), 0)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate = groupByDataFrame(dfInterestRate)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            imgDate = cv.GaussianBlur(imgDate, (3, 3), 0)
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            dfDate = groupByDataFrame(dfDate)
            # termDaysString in image
            dfTermDays = checkPatternRegexInDataFrame(dfDate, patternDict['termDays'], 10)
            termDaysText = dfTermDays.loc[dfTermDays.index[0], 'regex']
            termDaysInImage = int(re.sub('[A-Z]', '', termDaysText))
            # dateString in image
            dfDate = checkPatternRegexInDataFrame(dfDate, patternDict['date'], 10)
            if dfDate.empty:
                continue
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            issueDateText, expireDateText = dateText.split('TO')
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            # Term Days
            termDays = (expireDate - issueDate).days
            if (termDays < 0) or (termDaysInImage != termDays):
                continue
            # Term Months
            termMonths = round(termDays/30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            imgContractNumber = cv.morphologyEx(imgContractNumber, cv.MORPH_CLOSE, (5, 5))
            imgContractNumber = cv.GaussianBlur(imgContractNumber, (3, 3), 0)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            if dfContractNumber['text'].dtype == 'float64':
                dfContractNumber['text'] = dfContractNumber['text'].astype('Int64').astype('str')
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            paid = 0
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runTAISHIN(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,\d{3})',
        'interestRate': r'(\d+[,.]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{4}\d{2}[A-Z]{2}\d{8})'
    }
    if bank != 'TAISHIN':
        bank = 'TAISHIN'
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check condition in page to read
        check = 'Interest Notice' not in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        Image.fromarray(fullImageScale).show()
        valueImage = _findCoords(np.array(fullImageScale), 'values', bank)
        dfValue = readImgPytesseractToDataframe(valueImage, 11)
        dfValue = dfValue.loc[
            (dfValue['text'].str.contains(r'[0-9]+')) &
            (dfValue['conf'] > 20)
        ]
        try:
            # amount
            dfAmount = checkPatternRegexInDataFrame(dfValue, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '')) * 1000
            # interest rate
            dfInterestRate = checkPatternRegexInDataFrame(dfValue, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            dfDate = checkPatternRegexInDataFrame(dfValue, patternDict['date'], 10)
            if dfDate.empty:
                continue
            issueDateText = dfDate.loc[dfDate.index[0], 'regex']
            expireDateText = dfDate.loc[dfDate.index[1], 'regex']
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays/30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            paid = 0
            # remaining
            remaining = amount - paid
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception, ):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable

def runFUBON(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+[,.][\d,\d.]+\d{3})',
        'interestRate': r'(\d+[,.]+\d+)',
        'date': r'(\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{14})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # check condition in page to read
        check = 'Notice of Interest Payment' not in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            dfAmount = readImgPytesseractToDataframe(imgAmount, 6)
            dfAmount = groupByDataFrame(dfAmount)
            dfAmount = checkPatternRegexInDataFrame(dfAmount, patternDict['amount'], 10)
            if dfAmount.empty:
                continue
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '').replace('.', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate = groupByDataFrame(dfInterestRate)
            dfInterestRate = checkPatternRegexInDataFrame(dfInterestRate, patternDict['interestRate'], 10)
            if dfInterestRate.empty:
                continue
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.')) / 100
            # ngày hiệu lực, ngày đáo hạn
            # ngày hiệu lực
            imgIssueDate = _findCoords(np.array(fullImageScale), 'issueDate', bank)
            dfIssueDate = readImgPytesseractToDataframe(imgIssueDate, 6)
            dfIssueDate = groupByDataFrame(dfIssueDate)
            dfIssueDate = checkPatternRegexInDataFrame(dfIssueDate, patternDict['date'], 10)
            if dfIssueDate.empty:
                continue
            issueDateText = dfIssueDate.loc[dfIssueDate.index[0], 'regex']
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            # ngày đáo hạn
            imgExpireDate = _findCoords(np.array(fullImageScale), 'expireDate', bank)
            dfExpireDate = readImgPytesseractToDataframe(imgExpireDate, 6)
            dfExpireDate = groupByDataFrame(dfExpireDate)
            dfExpireDate = checkPatternRegexInDataFrame(dfExpireDate, patternDict['date'], 10)
            if dfExpireDate.empty:
                continue
            expireDateText = dfExpireDate.loc[dfExpireDate.index[0], 'regex']
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            if issueDate > expireDate:
                continue
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays / 30)
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber = checkPatternRegexInDataFrame(dfContractNumber, patternDict['contractNumber'], 10)
            if dfContractNumber.empty:
                continue
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            dfPaid = readImgPytesseractToDataframe(imgPaid, 6)
            dfPaid = groupByDataFrame(dfPaid)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', ''))
                paid = round(paid - interestAmount, 2)
            # remaining
            remaining = amount - paid
            # Append data
            records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount))
        except (Exception,):
            continue
    balanceTable = pd.DataFrame(
        records,
        columns=[
            'ContractNumber',
            'TermDays',
            'TermMonths',
            'InterestRate',
            'IssueDate',
            'ExpireDate',
            'Amount',
            'Paid',
            'Remaining',
            'InterestAmount'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', bank)
    # Currency
    balanceTable['Currency'] = 'USD'

    return balanceTable
