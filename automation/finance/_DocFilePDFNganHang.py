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
    # group by dataframe theo block_num
    groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
    groupByConf = df.groupby(['block_num'])['conf'].mean()
    dfGroupBy = pd.concat([groupByText, groupByConf], axis=1)
    return dfGroupBy

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
            right = int(left + 0.4 * h - 10)
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
            top = int(top + w + 10)
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
            right = int(left + h * 1.5 + 20)
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
            right = int(left + h * 1.5 + 16)
            left = left + h - 7
        return pdfImage[left:right, top:bottom]
    else:  # CATHAY
        right = int(left + h / 2)
        left = int(left - h / 4)
        bottom = top + w
        top = int(top + 0.3 * w)
        return pdfImage[left:right, top:bottom]

# Done
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
        # show image from pdf
        Image.fromarray(fullImageScale).show()
        # check image with condition
        check = 'INTEREST PAYMENT NOTICE' not in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 6)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '').replace('.', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            imgDate = cv2.resize(imgDate, (405, 55))
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            # dùng hàm groupByDataFrame
            dfDate = groupByDataFrame(dfDate)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 0)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 6)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[dfPaid['check']]
            dfPaid['regex'] = dfPaid['text'].str.extract(patternDict['paid'])
            if dfPaid.empty:
                paid = 0
            else:
                paid = amount
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

# done
def runSHINKONG(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3}\.\d{2})',
        'interestRate': r'(\d+[.,]\d+%)',
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
        Image.fromarray(fullImageScale).show()
        # check image with condition
        if 'Payment Advice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            if contractNumber.startswith('LI'):
                contractNumber = contractNumber.replace('LI', 'L1')
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 11)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            if dfPaid.empty:
                paid = 0
            else:
                paid = amount
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
def runYUANTA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[.,]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}-\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{4}[A-Z]{2}\d{5})'
    }
    images = convertPDFtoImage(bank, month)

    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        # check image with condition
        if 'Interest Payment Notice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        try:
            # detect table in image
            tables = detect_table(fullImageScale)
            table = tables[0]  # vì chỉ có 1 bảng trong hình nên phần tử đầu tiên luôn là bảng đó
            Image.fromarray(table).show()
            # amount
            imgAmount = _findCoords(np.array(table), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(table), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(table), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 4)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
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
        'amount': r'(\d+,[\d,]+\d{3}\.\d{2})',
        'interestRate': r'(\d+[.,]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[-~]\d{4}/\d{2}/\d{2}|\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{2}\d{7})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2}|\d\.\d{2})'
    }

    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            elif dfDate['check'].count() == 2:
                dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
                issueDateText = dfDate.loc[dfDate.index[0], 'regex']
                expireDateText = dfDate.loc[dfDate.index[1], 'regex']
                issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
                expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            else:
                dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            contractNumber = contractNumber.replace(contractNumber[0:2], '888LN')
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            if dfPaid.empty:
                continue
            dfPaid['regex'] = dfPaid['text'].str.extract(patternDict['paid'])
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
        'interestRate': r'(\d+[.|,]\d+%)',
        'date': r'(\d{1,2}[.,;:]?[A-Z]{3}[.,;:]?\d{4}to\d{1,2}[.,;:]?[A-Z]{3}[.,;:]?\d{4})',
        'contractNumber': r'(\d{8})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})'
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
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            # dùng hàm groupByDataFrame
            dfDate = groupByDataFrame(dfDate)
            dfDate['text'] = dfDate['text'].apply(lambda x: x.replace('..', '.'))
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber['text'] = dfContractNumber['text'].astype(str)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid & currency
            imgPaid = _findCoords(fullImageScale, 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            if dfPaid.empty:
                paid = 0
            else:
                dfPaid['regex'] = dfPaid['text'].str.extract(patternDict['paid'])
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', ''))
            if paid > amount:
                paid = amount
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
def runCHANGHWA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+[.|,]\d+%)',
        'date': r'(\d{1,2}[A-Z]{3}\.\d{4}TO\d{1,2}[A-Z]{3}\.\d{4})',
        'contractNumber': r'(\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3})',
        'paid': r'(\d+,[\d,]+\d{3})'
    }
    if bank != 'CHANG HWA':
        bank = 'CHANG HWA'
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        # check image with condition
        lst_condition = ['INTEREST PAYMENT NOTICE', 'PAYMENT NOTICE']
        check = all(c not in readImgPytesseractToString(fullImageScale, 6) for c in lst_condition)
        if check:
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 6)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            # dùng hàm groupByDataFrame
            dfDate = groupByDataFrame(dfDate)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber['text'] = dfContractNumber['text'].astype(str)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            imgPaid = _findCoords(fullImageScale, 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            try:
                # dùng try except chỗ này vì có trường hợp ảnh trả ra là 1 ảnh trắng, nên dfPaid rỗng
                dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
                dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            except (Exception, ):
                dfPaid = pd.DataFrame()
            if dfPaid.empty:
                paid = 0
            else:
                dfPaid['regex'] = dfPaid['text'].str.extract(patternDict['paid'])
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
        'amount': r'(\d+,[\d,]+\d{3}\.\d{2}|\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+\.\d+)',
        'date': r'(\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{3}\d{7}[A-Z]{2})'
    }
    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # grayscale full image
        fullImageScale = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        # check image with condition
        if 'NOTICE OF LOAN REPAYMENT' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] >= 0)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate['text'] = dfInterestRate['text'].astype(str)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[
                (dfContractNumber['check']) & (dfContractNumber['conf'] > 10)
            ]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
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
def runESUN(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    patternDict = {
        'amount': r'(\d+[,.][\d,\d.]+\d{3})',
        'interestRate': r'(\d+[,.]+\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[~-]\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{6})',
        'paid': r'(\d+[,.][\d,\d.]+\d{3}[,.]\d{2})'
    }
    images = convertPDFtoImage(bank, month)
    if bank != 'ESUN':
        bank = 'ESUN'

    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()

        # check image with condition
        if 'Payment Advice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace('.', '').replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[
                (dfInterestRate['check']) & (dfInterestRate['conf'] > 10)
            ]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 4)
            # dùng hàm groupByDataFrame
            dfDate = groupByDataFrame(dfDate)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[
                (dfContractNumber['check']) & (dfContractNumber['conf'] > 10)
            ]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            contractNumber = '22OBLN' + contractNumber
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 11)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            if dfPaid.empty:
                paid = 0
            else:
                paid = amount
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
        'amount': r'(\d+,[\d,]+\d{3}\.\d{2})',
        'interestRate': r'(\d+[.,]+\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{1}[A-Z]{7}\d{7})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        conditionList = ['(LN4030)', 'N4030']
        check = any(c in readImgPytesseractToString(fullImageScale, 6) for c in conditionList)
        if not check:
            continue
        try:
            # contract number
            imgContractNumber = _findCoords(np.array(fullImageScale), 'contractNumber', bank)
            imgContractNumber = erosionAndDilation(imgContractNumber, 'e', 3)
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            # dùng hàm groupByDataFrame
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[
                (dfContractNumber['check']) & (dfContractNumber['conf'] >= 0)
            ]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
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
                Image.fromarray(imgDate).show()
                dfDate = readImgPytesseractToDataframe(imgDate, 11)
                # dùng hàm groupByDataFrame
                dfDate = groupByDataFrame(dfDate)
                dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
                dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
                if dfDate.empty:
                    continue
                dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
                Image.fromarray(imgAmount).show()
                dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
                # dùng hàm groupByDataFrame
                dfAmount = groupByDataFrame(dfAmount)
                dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
                dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 0)]
                if dfAmount.empty:
                    continue
                dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
                print("amount confidence: ", dfAmount.loc[dfAmount.index[0], 'conf'])
                amountText = dfAmount.loc[dfAmount.index[0], 'regex']
                amount = float(amountText.replace(',', ''))
                # interest rate
                bottomInterestRate = top + w * 20
                topInterestRate = top + w * 14
                rightInterestRate = int(left + h + 6.5)
                # interest image
                imgInterestRate = fullImageScale[left:rightInterestRate, topInterestRate:bottomInterestRate]
                Image.fromarray(imgInterestRate).show()
                dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
                # dùng hàm groupByDataFrame
                dfInterestRate = groupByDataFrame(dfInterestRate)
                dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
                dfInterestRate = dfInterestRate.loc[
                    (dfInterestRate['check']) & (dfInterestRate['conf'] > 10)
                ]
                if dfInterestRate.empty:
                    continue
                dfInterestRate['text'] = dfInterestRate['text'].apply(lambda x: x.replace('..', '.'))
                dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
                interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
                interestRate = float(interestRateText.replace('%', '')) / 100
                # interest amount
                interestAmount = amount * termDays * interestRate / 360
                # paid
                bottomPaid = top + w
                leftPaid = left + h
                rightPaid = left + h * 2
                imgPaid = fullImageScale[leftPaid:rightPaid, top:bottomPaid]
                Image.fromarray(imgPaid).show()
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

# Mới có file của tháng 8 nên chưa có nhiều data để test, tạm xong
def runBOP(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+\.\d+%)',
        'date': r'(\d{4}\d{2}\d{2})'
    }
    for img in images:
        img = img.rotate(-90, expand=1)
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()

        # check image with condition
        if 'Interest Rate Notice' not in readImgPytesseractToString(fullImageScale, 6):
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] >= 0)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[
                (dfInterestRate['check']) & (dfInterestRate['conf'] > 10)
            ]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            # ngày hiệu lực
            imgIssueDate = _findCoords(np.array(fullImageScale), 'issueDate', bank)
            Image.fromarray(imgIssueDate).show()
            dfIssueDate = readImgPytesseractToDataframe(imgIssueDate, 4)
            dfIssueDate['text'] = dfIssueDate['text'].astype(str)
            dfIssueDate['check'] = dfIssueDate['text'].str.contains(patternDict['date'])
            dfIssueDate = dfIssueDate.loc[(dfIssueDate['check']) & (dfIssueDate['conf'] > 10)]
            if dfIssueDate.empty:
                continue
            dfIssueDate['regex'] = dfIssueDate['text'].str.extract(patternDict['date'])
            issueDateText = dfIssueDate.loc[dfIssueDate.index[0], 'regex']
            issueDate = dt.datetime.strptime(issueDateText, '%Y%m%d')
            # ngày đáo hạn
            imgExpireDate = _findCoords(np.array(fullImageScale), 'expireDate', bank)
            Image.fromarray(imgExpireDate).show()
            dfExpireDate = readImgPytesseractToDataframe(imgExpireDate, 4)
            dfExpireDate['text'] = dfExpireDate['text'].astype(str)
            dfExpireDate['check'] = dfExpireDate['text'].str.contains(patternDict['date'])
            dfExpireDate = dfExpireDate.loc[(dfExpireDate['check']) & (dfExpireDate['conf'] > 10)]
            if dfExpireDate.empty:
                continue
            dfExpireDate['regex'] = dfExpireDate['text'].str.extract(patternDict['date'])
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
        'interestRate': r'(\d+\.\d+%)',
        'date': r'(\d{4}\.\d{1,2}\.\d{1,2}~\d{4}\.\d{1,2}\.\d{1,2})',
        'contractNumber': r'([A-Z]\d[A-Z]\d{7})',
        'paid': r'(\d+,[\d,]+\d{3})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        check = 'LOAN REPAYMENT CONFIRMATION' in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            # dùng hàm groupByDataFrame
            dfDate = groupByDataFrame(dfDate)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                # dfDate = pd.DataFrame()
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # paid
            imgPaid = _findCoords(fullImageScale, 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            if dfPaid.empty:
                paid = 0
            else:
                dfPaid['regex'] = dfPaid['text'].str.extract(patternDict['paid'])
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
def runUBOT(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+\.\d+%)',
        'date': r'(\d{4}\.\d{1,2}\.\d{1,2}-\d{4}\.\d{1,2}\.\d{1,2})',
        'contractNumber': r'(\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3}[-]+[A-Z]{5})',
        'paid': r'(\d+,[\d,]+\d{3})'
    }
    for img in images:
        img = img.rotate(-90, expand=1)
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # _, fullImageScale = cv.threshold(fullImageScale, 200, 255, cv.THRESH_BINARY)
        Image.fromarray(fullImageScale).show()

        # check image with condition
        lst_condition = ['Interest Payment Notice', 'Payment Notice', 'Interest Payment']
        check = all(c not in readImgPytesseractToString(fullImageScale, 6) for c in lst_condition)
        if check:
            continue
        try:
            # amount & currency
            imgAmountCurrency = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmountCurrency).show()
            dfAmount = readImgPytesseractToDataframe(imgAmountCurrency, 4)
            # dùng hàm groupByDataFrame
            dfAmount = groupByDataFrame(dfAmount)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] >= 0)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            # dùng hàm groupByDataFrame
            dfInterestRate = groupByDataFrame(dfInterestRate)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            imgDate = cv2.resize(imgDate, (500, 32), interpolation=cv2.INTER_NEAREST)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            # dùng hàm groupByDataFrame
            dfDate = groupByDataFrame(dfDate)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
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

def runTCB(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d+\.\d+%)',
        'date': r'(\d{2}/\d{2}/\d{4}-\d{2}/\d{2}/\d{4})',
        'contractNumber': r'([A-Z]{4}\d[A-Z]{2}\d{5})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()

        # check image with condition
        check = 'INTEREST PAYMENT NOTICE' not in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] >= 0)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 4)
            # dùng hàm groupByDataFrame
            dfDate = groupByDataFrame(dfDate)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
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
        'interestRate': r'(\d+[.,]\d+)',
        'date': r'(\d{4}/\d{2}/\d{2}TO\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{11}[-]+\d{1,2})',
        'termDays': r'(\d{1,2}[A-Z]{4})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        # check condition in page to read
        imgCondition = _findCoords(np.array(fullImageScale), 'condition', bank)
        Image.fromarray(imgCondition).show()
        condition = readImgPytesseractToString(imgCondition, 11)
        condition = re.sub(r'[:A-Z\n\s\']', '', condition)
        if len(condition) < 12:
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 6)
            dfAmount = groupByDataFrame(dfAmount)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] >= 0)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '').replace('.', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            imgInterestRate = cv.morphologyEx(imgInterestRate, cv.MORPH_OPEN, (3, 3))
            imgInterestRate = cv.GaussianBlur(imgInterestRate, (5, 5), 0)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate = groupByDataFrame(dfInterestRate)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.')) / 100
            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(fullImageScale), 'date', bank)
            imgDate = cv.GaussianBlur(imgDate, (3, 3), 0)
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 6)
            dfDate = groupByDataFrame(dfDate)
            # termDaysString in image
            dfTermDays = dfDate.loc[dfDate['text'].str.contains(patternDict['termDays'])]
            dfTermDays['regex'] = dfTermDays['text'].str.extract(patternDict['termDays'])
            termDaysText = dfTermDays.loc[dfTermDays.index[0], 'regex']
            termDaysInImage = int(re.sub('[A-Z]', '', termDaysText))
            # dateString in image
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            if dfContractNumber['text'].dtype == 'float64':
                dfContractNumber['text'] = dfContractNumber['text'].astype('Int64').astype('str')
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            contractNumber = contractNumber.replace('--', '-')
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
        'interestRate': r'(\d+\.\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{4}\d{2}[A-Z]{2}\d{8})'
    }
    if bank != 'TAISHIN':
        bank = 'TAISHIN'
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        # check condition in page to read
        check = 'Interest Notice' not in readImgPytesseractToString(fullImageScale, 6)
        if check:
            continue
        valueImage = _findCoords(np.array(fullImageScale), 'values', bank)
        Image.fromarray(valueImage).show()
        dfValue = readImgPytesseractToDataframe(valueImage, 11)
        dfValue = dfValue.loc[
            (dfValue['text'].str.contains(r'[0-9]+')) &
            (dfValue['conf'] > 20)
        ]
        try:
            # amount
            dfAmount = dfValue.loc[
                (dfValue['text'].str.contains(patternDict['amount'])) &
                (dfValue['conf'] > 10)
            ].copy()
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '')) * 1000
            # interest rate
            dfInterestRate = dfValue.loc[
                (dfValue['text'].str.contains(patternDict['interestRate'])) &
                (dfValue['conf'] > 10)
            ].copy()
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace('%', '')) / 100
            # ngày hiệu lực, ngày đáo hạn
            dfDate = dfValue.loc[
                (dfValue['text'].str.contains(patternDict['date'])) &
                (dfValue['conf'] > 10)
            ].copy()
            if dfDate.empty:
                continue
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
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
        'interestRate': r'(\d+[.,]+\d+)',
        'date': r'(\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{14})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})'
    }
    for img in images:
        # grayscale full image
        fullImageScale = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(fullImageScale).show()
        # check condition in page to read
        lst_condition = ['Notice of Interest Payment', 'Interest Payment']
        check = all(c not in readImgPytesseractToString(fullImageScale, 6) for c in lst_condition)
        if check:
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(fullImageScale), 'amount', bank)
            imgAmount = cv.morphologyEx(imgAmount, cv.MORPH_OPEN, (151, 151))
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            # dùng hàm groupByDataFrame
            dfAmount = groupByDataFrame(dfAmount)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', '').replace('.', ''))
            # interest rate
            imgInterestRate = _findCoords(np.array(fullImageScale), 'interestRate', bank)
            imgInterestRate = cv.morphologyEx(imgInterestRate, cv.MORPH_OPEN, (99, 99))
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 6)
            dfInterestRate['text'] = dfInterestRate['text'].astype(str)
            # dùng hàm groupByDataFrame
            dfInterestRate = groupByDataFrame(dfInterestRate)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 0)]
            if dfInterestRate.empty:
                continue
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.')) / 100
            # ngày hiệu lực, ngày đáo hạn
            # ngày hiệu lực
            imgIssueDate = _findCoords(np.array(fullImageScale), 'issueDate', bank)
            imgIssueDate = erosionAndDilation(imgIssueDate, 'e', 3)
            Image.fromarray(imgIssueDate).show()
            dfIssueDate = readImgPytesseractToDataframe(imgIssueDate, 6)
            dfIssueDate = groupByDataFrame(dfIssueDate)
            dfIssueDate['check'] = dfIssueDate['text'].str.contains(patternDict['date'])
            dfIssueDate = dfIssueDate.loc[(dfIssueDate['check']) & (dfIssueDate['conf'] > 10)]
            if dfIssueDate.empty:
                continue
            dfIssueDate['regex'] = dfIssueDate['text'].str.extract(patternDict['date'])
            issueDateText = dfIssueDate.loc[dfIssueDate.index[0], 'regex']
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            # ngày đáo hạn
            imgExpireDate = _findCoords(np.array(fullImageScale), 'expireDate', bank)
            imgExpireDate = erosionAndDilation(imgExpireDate, 'e', 3)
            Image.fromarray(imgExpireDate).show()
            dfExpireDate = readImgPytesseractToDataframe(imgExpireDate, 6)
            # Dùng hàm groupByDataFrame
            dfExpireDate = groupByDataFrame(dfExpireDate)
            dfExpireDate['check'] = dfExpireDate['text'].str.contains(patternDict['date'])
            dfExpireDate = dfExpireDate.loc[(dfExpireDate['check']) & (dfExpireDate['conf'] > 10)]
            if dfExpireDate.empty:
                continue
            dfExpireDate['regex'] = dfExpireDate['text'].str.extract(patternDict['date'])
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
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 6)
            dfContractNumber['text'] = dfContractNumber['text'].astype(str)
            # dùng hàm groupByDataFrame
            dfContractNumber = groupByDataFrame(dfContractNumber)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 0)]
            if dfContractNumber.empty:
                continue
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
            # interest amount
            interestAmount = amount * termDays * interestRate / 360
            # paid
            imgPaid = _findCoords(np.array(fullImageScale), 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid, 6)
            # dùng hàm groupByDataFrame
            dfPaid = groupByDataFrame(dfPaid)
            dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
            dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
            if dfPaid.empty:
                paid = 0
            else:
                paidText = dfPaid.loc[dfPaid.index[0], 'regex']
                paid = float(paidText.replace(',', ''))
                paid = round(paid - interestAmount)
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