from automation import *
import cv2 as cv
import warnings

warnings.filterwarnings("ignore", 'This pattern has match groups')

"""
--------- RULES ---------
1. Số hợp đồng (ContractNumber):
    - Nếu có "contract number" trong file thì lấy cái này làm contract number
    - Nếu có Ref No., Ref, Loan Ref thì lấy cái này làm contract number
    - Nếu có TXNT thì lấy cái này làm contract number
2. Khi nào các ngân hàng có mẫu PDF mà bên PHS đã trả các khoản vay thì bên Finance thông báo
3. Theo dõi file loan nước ngoài từ trước tới giờ thì không có khoản vay nào thấp hơn 1,000,000
4. Các ngân hàng OBU (scan file PDF) trừ ESUN, các ngân hàng còn lại chỉ có 1 khoản vay trong mỗi file từng tháng
5. Rule của từng bank:
    5.1 CATHAY:
        - Nếu có CAP nằm ở đầu dòng -> Paid = Amount
        - Tối đa 1 file sẽ có 2 dòng với INT nằm ở đầu dòng
        - Nếu có CAP và 2 INT thì INT gần nhất (nằm ngay phía trên CAP) -> dòng INT đó Paid sẽ bằng Amount, còn Int ở
        trên cùng paid = 0
    5.2 SHHK:
        - Trước tới giờ 1 loan account chỉ có đúng 1 khoản vay (Bảng đầu tiên trong file chỉ có 1 dòng)
        - Số hợp đồng (Loan Ref.) 6 kí tự đầu '888LN' cố định
    5.3 ESUN:
        - Số hợp đồng: 6 kí tự đầu cố định
        - Trang để lấy xử lý dữ liệu là trang có 'Payment Advice'
        - Amount luôn lấy từ trường thông tin Loan Outstanding
    5.4 ENTIE:
        - Amount = Loan O/S Amount
        - Loan O/S Amount tương dương principal payable trong file
"""


def rotate_pdf(dir: str, bank: str):
    """
    Hàm xoay từng trang trong file PDF
    """
    pdf_in = open(join(dir, f'{bank}.pdf'), 'rb')
    reader = PdfFileReader(pdf_in)
    writer = PdfFileWriter()
    for pagenum in range(reader.numPages):
        page = reader.getPage(pagenum)
        page.rotateClockwise(90)
        writer.addPage(page)
    pdf_out = open(join(dir, f'{bank}_rotate.pdf'), 'wb')
    writer.write(pdf_out)
    pdf_out.close()
    pdf_in.close()

    return pdf_out.name


def convertPDFtoImage(bankName: str, month: int):
    """
    Hàm chuyển từng trang trong PDF sang image
    """
    directory = join(realpath(dirname(__file__)), 'bank', 'pdf', f'THÁNG {month}')
    bank_rotate = ['BOP', 'SINOPAC', 'UBOT']
    check = any(bank in bankName for bank in bank_rotate)
    if check:
        pathPDF = rotate_pdf(directory, bankName)
    else:
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
    # Load iamge, grayscale, adaptive threshold
    image = np.array(inputImage)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 51, 9)
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
    for c in reversed(cnts):
        x, y, w, h = cv2.boundingRect(c)
        if x > w and y > h:
            continue
        roi = image[y:y + h, x:x + w]
        img_list.append(roi)
    return img_list


def readImgPytesseractToString(img, numConfig: int):
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
    df = df[df['conf'] != -1]

    return df

def _findTopLeftPoint(containingImage, pdfImage):
    containingImage = np.array(containingImage)  # đảm bảo image để được đưa về numpy array
    pdfImage = np.array(pdfImage)  # đảm bảo image để được đưa về numpy array
    matchResult = cv.matchTemplate(containingImage, pdfImage, cv.TM_CCOEFF)
    _, _, _, topLeft = cv.minMaxLoc(matchResult)
    return topLeft[0], topLeft[1]  # cho compatible với openCV

def _findTopLeftPointCATHAY(containingImage, pdfImage):
    containingImage = np.array(containingImage)  # đảm bảo image để được đưa về numpy array
    pdfImage = np.array(pdfImage)  # đảm bảo image để được đưa về numpy array

    matchResult = cv.matchTemplate(containingImage, pdfImage, cv.TM_SQDIFF)
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
    if name == 'amount':
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
    elif name == 'currency':
        fileName = 'currency.png'
    else:
        raise ValueError(
            'colName must be either "amount" or '
            '"interestRate" or '
            '"period" or '
            '"contractNumber" or '
            '"paid" or '
            '"currency" or '
            '"issueDate" or '
            '"expireDate" '
        )
    containingPath = os.path.join(os.path.dirname(__file__), 'bank_img', f'{bank}', fileName)
    containingImage = cv.imread(containingPath, 0)  # hình trắng đen (array 2 chiều)
    w, h = containingImage.shape[::-1]
    top, left = _findTopLeftPoint(pdfImage, containingImage)
    if bank == 'MEGA':
        if name in ['contractNumber', 'currency']:
            top = top + w - 5
            right = left + h + 5
            return pdfImage[left:right, top:]
        elif name == 'paid':
            top = top + w - 5
            bottom = top + w
            right = left + h + 5
            return pdfImage[left:right, top:bottom]
        else:
            bottom = top + w
            right = left + h * 2
            left = left + h - 5
            return pdfImage[left:right, top:bottom]
    elif bank == 'SHINKONG':
        top = top + w - 5
        right = left + h
        return pdfImage[left:right, top:]
    elif bank == 'YUANTA':
        if name == 'currency':
            right = left + h
            bottom = top + w
            return pdfImage[left:right, top:bottom]
        else:
            top = top + w - 5
            right = left + h + 5
            return pdfImage[left:right, top:]
    elif bank == 'SHHK':
        if name == 'date':
            bottom = top + w * 2
            right = left + h * 2 + 5
            left = left + h
            return pdfImage[left:right, top:bottom]
        elif name == 'paid':
            top = top + w
            right = left + h
            return pdfImage[left:right, top:]
        else:
            top = top - 5
            bottom = top + w
            right = left + h * 2 + 5
            left = left + h
            return pdfImage[left:right, top:bottom]
    elif bank == 'SINOPAC':
        if name in ['contractNumber', 'paid']:
            top = top + w - 5
            bottom = top + w
            right = left + h
            return pdfImage[left:right, top:bottom]
        else:
            bottom = top + w + 10
            top = top - 24
            right = left + h * 3 + 5
            left = left + h
            return pdfImage[left:right, top:bottom]
    elif bank == 'CHANG HWA':
        top = top + w - 5
        bottom = top + w
        right = left + h
        if name == 'interestRate':
            return pdfImage[left:right, top:bottom]
        return pdfImage[left:right, top:]
    elif bank == 'KGI':
        if name == 'currency':
            top = top - w
            right = left + h
            bottom = top + w
        else:
            top = top - 5
            bottom = top + w + 15
            right = left + h * 2 + 5
            left = left + h
        return pdfImage[left:right, top:bottom]
    elif bank == 'ESUN':
        if name == 'interestRate':
            top = top + w * 5 - 10
            right = left + h
            left = left - 20
        elif name == 'date':
            top = top + w
            right = left + h
        else:
            left = left - 5
            top = top + w
            right = left + h
        return pdfImage[left:right, top:]
    elif bank == 'BOP':
        bottom = top + w
        right = left + h * 2
        left = left + h
        return pdfImage[left:right, top:bottom]
    else:  # CATHAY
        if name == 'contractNumber':
            top = top + w
            bottom = top + w * 4
            right = left + h
        else:
            top = top + w * 2
            bottom = top + w
            right = left + h
            left = left - 10
        return pdfImage[left:right, top:bottom]

# Done
def runMEGA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    if bank != 'MEGA':
        bank = 'MEGA'

    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d{1,2}[.,]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[~-]+?\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{14})',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})',
        'currency': r'(VND|USD)'
    }
    for img in images:
        # scale image to gray
        img = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        # Deskewing text
        gray = cv2.bitwise_not(img)
        thresh = cv2.threshold(gray, 0, 255,cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
        coords = np.column_stack(np.where(thresh > 0))
        angle = cv2.minAreaRect(coords)[-1]
        if angle < 1:
            if angle < -45:
                angle = -(90 + angle)
            else:
                angle = -angle - 0.5
            (h, w) = img.shape[:2]
            center = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(center, angle, 1.0)
            img = cv2.warpAffine(img, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

        # show image from pdf
        Image.fromarray(img).show()

        # check image with condition
        lst_condition = ['INTEREST PAYMENT NOTICE', 'PAYMENT NOTICE', 'INTEREST PAYMENT']
        check = all(c not in readImgPytesseractToString(img, 6) for c in lst_condition)
        if check:
            continue

        # amount
        imgAmount = _findCoords(np.array(img), 'amount', bank)
        Image.fromarray(imgAmount).show()
        dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
        dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
        dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace(',','').replace('.',''))

        # interest rate
        imgInterestRate = _findCoords(np.array(img), 'interestRate', bank)
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
        imgDate = _findCoords(np.array(img), 'date', bank)
        # imgDate = imgDate[:-5, :]
        Image.fromarray(imgDate).show()
        dfDate = readImgPytesseractToDataframe(imgDate, 4)
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
        issueDate = dt.datetime.strptime(issueDateText.replace('-',''), '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText.replace('-',''), '%Y/%m/%d')

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(img), 'contractNumber', bank)
        Image.fromarray(imgContractNumber).show()
        dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
        dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
        dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
        if dfContractNumber.empty:
            continue
        dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
        contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']

        # currency
        imgCurrency = _findCoords(np.array(img), 'currency', bank)
        Image.fromarray(imgCurrency).show()
        dfCurrency = readImgPytesseractToDataframe(imgCurrency, 11)
        dfCurrency['check'] = dfCurrency['text'].str.contains(patternDict['currency'])
        dfCurrency = dfCurrency.loc[(dfCurrency['check']) & (dfCurrency['conf'] > 10)]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # paid
        imgPaid = _findCoords(np.array(img), 'paid', bank)
        Image.fromarray(imgPaid).show()
        dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
        dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
        dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
        dfPaid['regex'] = dfPaid['text'].str.extract(patternDict['paid'])
        if dfPaid.empty:
            paid = 0
        else:
            paid = dfPaid.loc[dfPaid.index[0], 'regex']

        # remaining
        remaining = amount - paid

        # interest amount
        interestAmount = amount * termDays * interestRate / 360

        # Append data
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'MEGA')

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runSHINKONG(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)

    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3}\.\d{2})',
        'interestRate': r'(\d{1,2}[.,]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\w+)',
        'currency': r'(VND|USD)'
    }

    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # scale image to gray
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # using opening morphological
        kernel = np.ones((3, 3), np.uint8)
        img = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)

        # check image with condition
        if 'Payment Advice' not in readImgPytesseractToString(img, 6):
            continue

        # amount
        imgAmount = _findCoords(np.array(img), 'amount', bank)
        Image.fromarray(imgAmount).show()
        dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
        dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
        dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace(',', ''))

        # currency
        dfCurrency = dfAmount.loc[dfAmount['text'].str.contains(patternDict['currency'])]
        dfCurrency = dfCurrency.loc[(dfCurrency['check']) & (dfCurrency['conf'] > 10)]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # interest rate
        imgInterestRate = _findCoords(np.array(img), 'interestRate', bank)
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
        imgDate = _findCoords(np.array(img), 'date', bank)
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

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(img), 'contractNumber', bank)
        Image.fromarray(imgContractNumber).show()
        dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 4)
        dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
        dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
        if dfContractNumber.empty:
            continue
        dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
        contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']
        if contractNumber.startswith('LI'):
            contractNumber = contractNumber.replace('LI', 'L1')

        # paid
        paid = 0

        # remaining
        remaining = amount - paid

        # interest amount
        interestAmount = amount * termDays * interestRate / 360

        # Append data
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'SHINKONG')

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runYUANTA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)

    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d{1,2}[.,]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}-\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{4}\d{4}[A-Z]{2}\d{5})',
        'currency': r'(VND|USD)'
    }

    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # grayscale full image
        scaleFullImg = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # check image with condition
        if 'Interest Payment Notice' not in readImgPytesseractToString(img, 6):
            continue

        # detect table in image
        tables = detect_table(img)
        table = tables[0]
        # grayscale table image
        scaleTable = cv2.cvtColor(table, cv2.COLOR_BGR2GRAY)
        Image.fromarray(scaleTable).show()

        # amount
        imgAmount = _findCoords(np.array(scaleTable), 'amount', bank)
        # Image.fromarray(imgAmount).show()
        dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
        dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
        dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace(',', ''))

        # currency
        imgCurrency = _findCoords(np.array(scaleFullImg), 'currency', bank)
        # Image.fromarray(imgCurrency).show()
        dfCurrency = readImgPytesseractToDataframe(imgCurrency, 6)
        dfCurrency['check'] = dfCurrency['text'].str.contains(patternDict['currency'])
        dfCurrency = dfCurrency.loc[(dfCurrency['check']) & (dfCurrency['conf'] > 10)]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # interest rate
        imgInterestRate = _findCoords(np.array(scaleTable), 'interestRate', bank)
        # Image.fromarray(imgInterestRate).show()
        dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 4)
        dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
        dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
        if dfInterestRate.empty:
            continue
        dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
        interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
        interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100

        # ngày hiệu lực, ngày đáo hạn
        imgDate = _findCoords(np.array(scaleTable), 'date', bank)
        # Image.fromarray(imgDate).show()
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

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(scaleFullImg), 'contractNumber', bank)
        # Image.fromarray(imgContractNumber).show()
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
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'YUANTA')

    return balanceTable

# Done
def runSHHK(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3}\.\d{2})',
        'interestRate': r'(\d{1,2}[.,]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[-~]\d{4}/\d{2}/\d{2}|\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{2}\d{7})',
        'currency': r'(VND|USD)',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2}|\d\.\d{2})'
    }

    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # grayscale full image
        scaleFullImg = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # amount
        imgAmount = _findCoords(np.array(scaleFullImg), 'amount', bank)
        Image.fromarray(imgAmount).show()
        dfAmount = readImgPytesseractToDataframe(imgAmount, 4)
        dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
        dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace(',', ''))

        # currency
        imgCurrency = _findCoords(np.array(scaleFullImg), 'currency', bank)
        Image.fromarray(imgCurrency).show()
        dfCurrency = readImgPytesseractToDataframe(imgCurrency, 6)
        dfCurrency['check'] = dfCurrency['text'].str.contains(patternDict['currency'])
        dfCurrency = dfCurrency.loc[(dfCurrency['check']) & (dfCurrency['conf'] > 10)]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # interest rate
        imgInterestRate = _findCoords(np.array(scaleFullImg), 'interestRate', bank)
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
        imgDate = _findCoords(np.array(scaleFullImg), 'date', bank)
        Image.fromarray(imgDate).show()
        dfDate = readImgPytesseractToDataframe(imgDate, 4)
        dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
        dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
        if dfDate.empty:
            # dfDate = pd.DataFrame()
            continue
        elif dfDate['check'].count() == 2:
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
            issueDateText = dfDate.loc[dfDate.index[0], 'regex']
            expireDateText = dfDate.loc[dfDate.index[1], 'regex']
        else:
            dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
            dateText = dfDate.loc[dfDate.index[0], 'regex']
            if '~' in dateText:
                issueDateText, expireDateText = dateText.split('~')
            else:
                issueDateText, expireDateText = dateText.split('-')
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(scaleFullImg), 'contractNumber', bank)
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
        imgPaid = _findCoords(np.array(scaleFullImg), 'paid', bank)
        Image.fromarray(imgPaid).show()
        dfPaid = readImgPytesseractToDataframe(imgPaid, 4)
        dfPaid['check'] = dfPaid['text'].str.contains(patternDict['paid'])
        dfPaid = dfPaid.loc[(dfPaid['check']) & (dfPaid['conf'] > 10)]
        if dfPaid.empty:
            continue
        dfPaid['regex'] = dfPaid['text'].str.extract(patternDict['paid'])
        paid = dfPaid.loc[dfPaid.index[0], 'regex']
        paid = float(paid.replace(',',''))

        # remaining
        remaining = amount - paid

        # interest amount
        interestAmount = amount * termDays * interestRate / 360

        # Append data
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'SHHK')

    return balanceTable

# Done
def runSINOPAC(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d{1,2}[.|,]\d+%)',
        'date': r'(\d{1,2}[.,;:]?[A-Z]{3}[.,;:]?\d{4}to\d{1,2}[.,;:]?[A-Z]{3}[.,;:]?\d{4})',
        'contractNumber': r'(\d{8})',
        'currency': r'(VND|USD)',
        'paid': r'(\d+,[\d,]+\d{3}\.\d{2})'
    }
    if bank != 'SINOPAC':
        bank = 'SINOPAC'
    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # grayscale full image
        scaleFullImg = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        Image.fromarray(scaleFullImg).show()

        # amount
        imgAmount = _findCoords(np.array(scaleFullImg), 'amount', bank)
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
        imgInterestRate = _findCoords(np.array(scaleFullImg), 'interestRate', bank)
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
        imgDate = _findCoords(np.array(scaleFullImg), 'date', bank)
        Image.fromarray(imgDate).show()
        dfDate = readImgPytesseractToDataframe(imgDate, 11)
        # group by dataframe theo block_num
        groupByText = dfDate.groupby(['line_num'])['text'].apply(lambda x: ''.join(list(x)))
        groupByConf = dfDate.groupby(['line_num'])['conf'].mean()
        dfDate = pd.concat([groupByText, groupByConf], axis=1)
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

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(scaleFullImg), 'contractNumber', bank)
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
        imgPaid = _findCoords(scaleFullImg, 'paid', bank)
        Image.fromarray(imgPaid).show()
        dfPaidCurrency = readImgPytesseractToDataframe(imgPaid, 4)

        # currency
        dfCurrency = dfPaidCurrency.loc[
            (dfPaidCurrency['text'].str.contains(patternDict['currency'])) &
            (dfPaidCurrency['conf'] > 10)
        ].reset_index(drop=True)
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # paid
        dfPaid = dfPaidCurrency.loc[
            (dfPaidCurrency['text'].str.contains(patternDict['paid'])) &
            (dfPaidCurrency['conf'] > 10)
        ]
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
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'SINOPAC')

    return balanceTable

# Done
def runCHANGHWA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d{1,2}[.|,]\d+%)',
        'date': r'(\d{1,2}[A-Z]{3}\.\d{4}TO\d{1,2}[A-Z]{3}\.\d{4})',
        'contractNumber': r'(\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3})',
        'currency': r'(VND|USD)',
        'paid': r'(\d+,[\d,]+\d{3})'
    }
    if bank != 'CHANG HWA':
        bank = 'CHANG HWA'
    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # grayscale full image
        scaleFullImg = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        Image.fromarray(scaleFullImg).show()
        # check image with condition
        lst_condition = ['INTEREST PAYMENT NOTICE', 'PAYMENT NOTICE']
        check = all(c not in readImgPytesseractToString(img, 6) for c in lst_condition)
        if check:
            continue

        # amount & currency
        imgAmountCurrency = _findCoords(np.array(scaleFullImg), 'amount', bank)
        Image.fromarray(imgAmountCurrency).show()
        dfAmountCurrency = readImgPytesseractToDataframe(imgAmountCurrency, 11)
        # amount
        dfAmount = dfAmountCurrency.loc[
            (dfAmountCurrency['text'].str.contains(patternDict['amount'])) &
            (dfAmountCurrency['conf'] > 10)
        ]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace(',', ''))

        # currency
        dfCurrency = dfAmountCurrency.loc[
            (dfAmountCurrency['text'].str.contains(patternDict['currency'])) &
            (dfAmountCurrency['conf'] > 10)
        ]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # interest rate
        imgInterestRate = _findCoords(np.array(scaleFullImg), 'interestRate', bank)
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
        imgDate = _findCoords(np.array(scaleFullImg), 'date', bank)
        Image.fromarray(imgDate).show()
        dfDate = readImgPytesseractToDataframe(imgDate, 11)
        # group by dataframe theo block_num
        groupByText = dfDate.groupby(['line_num'])['text'].apply(lambda x: ''.join(list(x)))
        groupByConf = dfDate.groupby(['line_num'])['conf'].mean()
        dfDate = pd.concat([groupByText, groupByConf], axis=1)
        dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
        dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
        if dfDate.empty:
            # dfDate = pd.DataFrame()
            continue
        dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
        dateText = dfDate.loc[dfDate.index[0], 'regex']
        dateText = dateText.replace('.', '')
        issueDateText, expireDateText = dateText.split('TO')
        issueDate = dt.datetime.strptime(issueDateText, '%d%b%Y')
        expireDate = dt.datetime.strptime(expireDateText, '%d%b%Y')

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(scaleFullImg), 'contractNumber', bank)
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
        imgPaid = _findCoords(scaleFullImg, 'paid', bank)
        Image.fromarray(imgPaid).show()
        paidString = readImgPytesseractToString(imgPaid, 6)
        if paidString == '':
            dfPaid = pd.DataFrame()
        else:
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
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'CHANG HWA')

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runKGI(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3}\.\d{2}|\d+,[\d,]+\d{3})',
        'interestRate': r'(\d{1,2}\.\d+)',
        'date': r'(\d{4}/\d{2}/\d{2})',
        'contractNumber': r'([A-Z]{3}\d{7}[A-Z]{2})',
        'currency': r'(VND|USD)'
    }
    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # grayscale full image
        scaleFullImg = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        Image.fromarray(scaleFullImg).show()

        # check image with condition
        if 'NOTICE OF LOAN REPAYMENT' not in readImgPytesseractToString(img, 6):
            continue

        # amount
        imgAmount = _findCoords(np.array(scaleFullImg), 'amount', bank)
        Image.fromarray(imgAmount).show()
        dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
        dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
        dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] >= 0)]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace(',', ''))

        # currency
        imgCurrency = _findCoords(np.array(scaleFullImg), 'currency', bank)
        Image.fromarray(imgCurrency).show()
        dfCurrency = readImgPytesseractToDataframe(imgCurrency, 6)
        dfCurrency['check'] = dfCurrency['text'].str.contains(patternDict['currency'])
        dfCurrency = dfCurrency.loc[(dfCurrency['check']) & (dfCurrency['conf'] > 10)]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # interest rate
        imgInterestRate = _findCoords(np.array(scaleFullImg), 'interestRate', bank)
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
        imgDate = _findCoords(np.array(scaleFullImg), 'date', bank)
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
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(scaleFullImg), 'contractNumber', bank)
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
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'KGI')

    return balanceTable

# chưa thấy file PDF có phần trả khoản vay nên tạm để paid = 0
def runESUN(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+[,|.][\d,|\d.]+\d{3})',
        'interestRate': r'(\d{1,2}[,.]+\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[~-]\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{6})',
        'currency': r'(VND|USD)'
    }
    if bank != 'ESUN':
        bank = 'ESUN'
    for img in images:
        # grayscale full image
        scaleFullImg = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(scaleFullImg).show()

        # check image with condition
        if 'Payment Advice' not in readImgPytesseractToString(img, 6):
            continue

        # amount & currency
        imgAmountCurrency = _findCoords(np.array(scaleFullImg), 'amount', bank)
        Image.fromarray(imgAmountCurrency).show()
        dfAmountCurrency = readImgPytesseractToDataframe(imgAmountCurrency, 4)
        # amount
        dfAmount = dfAmountCurrency.loc[
            (dfAmountCurrency['text'].str.contains(patternDict['amount'])) &
            (dfAmountCurrency['conf'] > 10)
        ]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace('.', '').replace(',', ''))

        # currency
        dfCurrency = dfAmountCurrency.loc[
            (dfAmountCurrency['text'].str.contains(patternDict['currency'])) &
            (dfAmountCurrency['conf'] > 10)
        ]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # interest rate
        imgInterestRate = _findCoords(np.array(scaleFullImg), 'interestRate', bank)
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
        interestRate = float(interestRateText.replace(',', '').replace('%', '')) / 100

        # ngày hiệu lực, ngày đáo hạn
        imgDate = _findCoords(np.array(scaleFullImg), 'date', bank)
        Image.fromarray(imgDate).show()
        dfDate = readImgPytesseractToDataframe(imgDate, 4)
        # group by dataframe theo block_num
        groupByText = dfDate.groupby(['line_num'])['text'].apply(lambda x: ''.join(list(x)))
        groupByConf = dfDate.groupby(['line_num'])['conf'].mean()
        dfDate = pd.concat([groupByText, groupByConf], axis=1)
        dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
        dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
        if dfDate.empty:
            # dfDate = pd.DataFrame()
            continue
        dfDate['regex'] = dfDate['text'].str.extract(patternDict['date'])
        dateText = dfDate.loc[dfDate.index[0], 'regex']
        if '~' in dateText:
            issueDateText, expireDateText = dateText.split('~')
        else:
            issueDateText, expireDateText = dateText.split('-')
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        imgContractNumber = _findCoords(np.array(scaleFullImg), 'contractNumber', bank)
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
        contractNumber = '22OBLN' + contractNumber

        # paid
        paid = 0

        # remaining
        remaining = amount - paid

        # interest amount
        interestAmount = amount * termDays * interestRate / 360

        # Append data
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'ESUN')

    return balanceTable

# chưa ổn lắm, xử lý không dùng opencv
def _runCATHAY(bank: str, month: int):
    # List các trường hợp để nhận biết file hình đúng để xử lý
    now = dt.datetime.now()
    images = convertPDFtoImage(bank, month)
    patternDict = {
        r'\bINT\b|\bCAP\b': 'condition',
        r'\d[A-Z]{7}\d{7}\b|\b[A-Z]{8}\d{7}': 'contractNumber',
        r'\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2}': 'date',
        r'\d+,[\d,|\d.]+\d{3}\.\d{2}': 'amount',
        r'\b\d{1,2}\.\d+%': 'interestRate',
        r'USD': 'currency'
    }
    frames = []

    for img in images:
        try:
            df = readImgPytesseractToDataframe(img, 11)
            # group by dataframe theo block_num
            groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
            groupByConf = df.groupby(['block_num'])['conf'].sum()
            df = pd.concat([groupByText, groupByConf], axis=1)
            # check if LN4030 in file
            check = df['text'].isin(['(LN4030)', 'N4030']).any()
            if not check:
                continue

            # convert dictionary values to list
            patternList = list(patternDict.keys())

            # check xem giá trị nào trong cột text khớp với các pattern trong danh sách pattern
            df['check'] = df['text'].str.contains('(' + '|'.join(patternList) + ')')
            df = df.loc[df['check']]
            df['text'] = df['text'].str.extract('(' + '|'.join(patternList) + ')')
            df['regex'] = df[['text']].replace({'text': patternDict}, regex=True)

            # contractNumber
            contractNumber = df['text'].loc[df['regex'] == 'contractNumber'].item()
            # Ngày hiệu lực, Ngày đáo hạn
            dateText = df['text'].loc[df['regex'] == 'date'].tolist()
            issueDates, expireDates = [], []
            for d in dateText:
                issueDateText, expireDateText = d.split('~')
                if issueDateText == expireDateText:
                    continue
                issueDates.append(issueDateText)
                expireDates.append(expireDateText)
            issueDate = [dt.datetime.strptime(issueDate, '%Y/%m/%d') for issueDate in issueDates]
            expireDate = [dt.datetime.strptime(expireDate, '%Y/%m/%d') for expireDate in expireDates]
            # Term Days
            termDays = [(expireDate - issueDate).days for issueDate, expireDate in zip(issueDate, expireDate)]
            # Term Months
            termMonths = [round(t/30) for t in termDays]
            # Amount
            amount = df['text'].loc[df['regex'].str.contains('amount')].tolist()
            amount = float(min(amount).replace(',', ''))
            # Paid
            paid = []
            if not df['text'].str.contains('CAP').any() and len(termDays) > 1:
                for i in range(len(termDays)):
                    paid.append(0)
            elif df['text'].str.contains('CAP').any() and len(termDays) > 1:
                for i in range(len(termDays)):
                    paid.append(0)
                paid[-1] = amount
            elif df['text'].str.contains('CAP').any() and len(termDays) == 1:
                paid.append(amount)
            else:
                paid.append(0)
            # Remaining
            remaining = [amount - p for p in paid]
            # currency
            currency = df['text'].loc[df['regex'] == 'currency'].tolist()
            if len(currency) > 1:
                currency = currency[0]
            # interest rate
            interestRateText = df['text'].loc[df['regex'].str.contains('interestRate')].tolist()
            if len(interestRateText) != len(set(interestRateText)):
                interestRateText = interestRateText[1:]
            interestRate = [float(i.replace('’', '').replace(',', '.').replace('%', '')) / 100 for i in
                            interestRateText]
            # print(interestRateText, interestRate)
            # Interest amount
            interestAmount = [amount * termDay * iRate / 360 for termDay, iRate in zip(termDays, interestRate)]

            dictionary = {
                'ContractNumber': contractNumber,
                'TermDays': termDays,
                'TermMonths': termMonths,
                'InterestRate': interestRate,
                'IssueDate': issueDate,
                'ExpireDate': expireDate,
                'Amount': amount,
                'Paid': paid,
                'Remaining': remaining,
                'InterestAmount': interestAmount,
                'Currency': currency
            }
            dataFrame = pd.DataFrame(data=dictionary)
            frames.append(dataFrame)
        except (Exception, ValueError):
            dataFrame = pd.DataFrame()
            frames.append(dataFrame)
    balanceTable = pd.concat(frames, ignore_index=True)
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'CATHAY')

    return balanceTable

# viết lại CATHAY theo cách khác (có sử dụng opencv):
def runCATHAY(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,?[\d,]+\d{3}\.\d{2})',
        'interestRate': r'(\d{1,2}\.+\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2})',
        'currency': r'(VND|USD)',
        'contractNumber': r'(\d{1}[A-Z]{7}\d{7})'
    }
    for img in images:
        # grayscale full image
        scaleFullImg = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(scaleFullImg).show()

        conditionList = ['(LN4030)', 'N4030']
        check = any(c in readImgPytesseractToString(scaleFullImg, 6) for c in conditionList)
        if not check:
            continue

        # contract number
        imgContractNumber = _findCoords(np.array(scaleFullImg), 'contractNumber', bank)
        Image.fromarray(imgContractNumber).show()
        dfContractNumber = readImgPytesseractToDataframe(imgContractNumber, 11)
        # group by dfContractNumber theo block_num
        groupByText = dfContractNumber.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
        groupByConf = dfContractNumber.groupby(['block_num'])['conf'].mean()
        dfContractNumber = pd.concat([groupByText, groupByConf], axis=1)
        dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
        dfContractNumber = dfContractNumber.loc[
            (dfContractNumber['check']) & (dfContractNumber['conf'] > 5)
        ]
        if dfContractNumber.empty:
            continue
        dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
        contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']

        # currency
        imgCurrency = _findCoords(np.array(scaleFullImg), 'currency', bank)
        Image.fromarray(imgCurrency).show()
        dfCurrency = readImgPytesseractToDataframe(imgCurrency, 4)
        dfCurrency['check'] = dfCurrency['text'].str.contains(patternDict['currency'])
        dfCurrency = dfCurrency.loc[
            (dfCurrency['check']) & (dfCurrency['conf'] > 10)
        ]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # các trường thông tin khác
        containingPath = os.path.join(os.path.dirname(__file__), 'bank_img', f'{bank}', 'check.png')
        containingImage = cv.imread(containingPath, 0)  # hình trắng đen (array 2 chiều)
        w, h = containingImage.shape[::-1]
        topLeftList = _findTopLeftPointCATHAY(containingImage, scaleFullImg)
        for topLeft in topLeftList:
            top, left = topLeft
            # Ngày hiệu lực, ngày đáo hạn
            topDate = top + w * 2 - 10
            bottomDate = top + w * 8 + 10
            rightDate = left + h + 7
            # Date image
            imgDate = scaleFullImg[left:rightDate, topDate:bottomDate]
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate, 11)
            # group by dataframe theo block_num
            groupByText = dfDate.groupby(['line_num'])['text'].apply(lambda x: ''.join(list(x)))
            groupByConf = dfDate.groupby(['line_num'])['conf'].mean()
            dfDate = pd.concat([groupByText, groupByConf], axis=1)
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
            # Term Days
            termDays = (expireDate - issueDate).days
            # Term Months
            termMonths = round(termDays/30)

            # amount
            topAmount = top + w * 10 - 50
            bottomAmount = top + w * 14 - 35
            rightAmount = left + h + 5
            leftAmount = left - 3
            # Amount image
            imgAmount = scaleFullImg[leftAmount:rightAmount, topAmount:bottomAmount]
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
            groupByText = dfAmount.groupby(['line_num'])['text'].apply(lambda x: ''.join(list(x)))
            groupByConf = dfAmount.groupby(['line_num'])['conf'].mean()
            dfAmount = pd.concat([groupByText, groupByConf], axis=1)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 0)]
            if dfAmount.empty:
                continue
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',', ''))

            # interest rate
            topInterestRate = top + w * 15 + 8
            bottomInterestRate = top + w * 19
            rightInterestRate = int(left + h + 6.5)
            # interest image
            imgInterestRate = scaleFullImg[left:rightInterestRate, topInterestRate:bottomInterestRate]
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate, 11)
            groupByText = dfInterestRate.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
            groupByConf = dfInterestRate.groupby(['block_num'])['conf'].mean()
            dfInterestRate = pd.concat([groupByText, groupByConf], axis=1)
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
            imgPaid = scaleFullImg[leftPaid:rightPaid, top:bottomPaid]
            Image.fromarray(imgPaid).show()
            paidText = readImgPytesseractToString(imgPaid, 6)
            if 'CAP' in paidText:
                paid = amount
            else:
                paid = 0
            # remaining
            remaining = amount - paid
            # Append data
            records.append([contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency])
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'CATHAY')

    return balanceTable

# Mới có file của tháng 8 nên chưa có nhiều data để test, tạm xong
def runBOP(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    patternDict = {
        'amount': r'(\d+,[\d,]+\d{3})',
        'interestRate': r'(\d{1,2}\.+\d+%)',
        'date': r'(\d{4}\d{2}\d{2})',
        'currency': r'(VND|USD)'
    }
    for img in images:
        # grayscale full image
        scaleFullImg = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        Image.fromarray(scaleFullImg).show()

        # check image with condition
        if 'Interest Rate Notice' not in readImgPytesseractToString(img, 6):
            continue

        # amount
        imgAmount = _findCoords(np.array(scaleFullImg), 'amount', bank)
        Image.fromarray(imgAmount).show()
        dfAmount = readImgPytesseractToDataframe(imgAmount, 11)
        dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
        dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] >= 0)]
        if dfAmount.empty:
            continue
        dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
        amountText = dfAmount.loc[dfAmount.index[0], 'regex']
        amount = float(amountText.replace(',', ''))

        # currency
        imgCurrency = _findCoords(np.array(scaleFullImg), 'currency', bank)
        Image.fromarray(imgCurrency).show()
        dfCurrency = readImgPytesseractToDataframe(imgCurrency, 6)
        dfCurrency['check'] = dfCurrency['text'].str.contains(patternDict['currency'])
        dfCurrency = dfCurrency.loc[(dfCurrency['check']) & (dfCurrency['conf'] > 10)]
        if dfCurrency.empty:
            continue
        dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
        currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

        # interest rate
        imgInterestRate = _findCoords(np.array(scaleFullImg), 'interestRate', bank)
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
        interestRate = float(interestRateText.replace(',', '').replace('%', '')) / 100

        # ngày hiệu lực, ngày đáo hạn
        # ngày hiệu lực
        imgIssueDate = _findCoords(np.array(scaleFullImg), 'issueDate', bank)
        Image.fromarray(imgIssueDate).show()
        dfIssueDate = readImgPytesseractToDataframe(imgIssueDate, 4)
        dfIssueDate['text'] = dfIssueDate['text'].astype(str)
        dfIssueDate['check'] = dfIssueDate['text'].str.contains(patternDict['date'])
        dfIssueDate = dfIssueDate.loc[(dfIssueDate['check']) & (dfIssueDate['conf'] > 10)]
        if dfIssueDate.empty:
            # dfDate = pd.DataFrame()
            continue
        dfIssueDate['regex'] = dfIssueDate['text'].str.extract(patternDict['date'])
        issueDateText = dfIssueDate.loc[dfIssueDate.index[0], 'regex']
        issueDate = dt.datetime.strptime(issueDateText, '%Y%m%d')

        # ngày đáo hạn
        imgExpireDate = _findCoords(np.array(scaleFullImg), 'expireDate', bank)
        Image.fromarray(imgExpireDate).show()
        dfExpireDate = readImgPytesseractToDataframe(imgExpireDate, 4)
        dfExpireDate['text'] = dfExpireDate['text'].astype(str)
        dfExpireDate['check'] = dfExpireDate['text'].str.contains(patternDict['date'])
        dfExpireDate = dfExpireDate.loc[(dfExpireDate['check']) & (dfExpireDate['conf'] > 10)]
        if dfExpireDate.empty:
            # dfDate = pd.DataFrame()
            continue
        dfExpireDate['regex'] = dfExpireDate['text'].str.extract(patternDict['date'])
        expireDateText = dfExpireDate.loc[dfExpireDate.index[0], 'regex']
        expireDate = dt.datetime.strptime(expireDateText, '%Y%m%d')

        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = round(termDays/30)

        # contract number
        contractNumber = None

        # paid
        paid = 0

        # remaining
        remaining = amount - paid

        # interest amount
        interestAmount = amount * termDays * interestRate / 360

        # Append data
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid, remaining, interestAmount, currency))
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
            'InterestAmount',
            'Currency'
        ]
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'BOP')

    return balanceTable



