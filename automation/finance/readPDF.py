import os

from automation import *
import cv2

def rotate_pdf(
    bank: str,
    month: int
):
    if bank in ['BOP', 'SINOPAC', 'UBOT']:
        pdf_in = open(join(realpath(dirname(__file__)),'PDF','pdf',f'THÁNG {month}',f'{bank}.pdf'), 'rb')
        reader = PdfFileReader(pdf_in)
        writer = PdfFileWriter()
        for pagenum in range(reader.numPages):
            page = reader.getPage(pagenum)
            page.rotateClockwise(90)
            writer.addPage(page)
        pdf_out = open(join(realpath(dirname(__file__)),'PDF','pdf',f'THÁNG {month}',f'{bank}_rotated.pdf'), 'wb')
        writer.write(pdf_out)
        pdf_out.close()
        pdf_in.close()

def convertPDFtoImage(
    bank: str,
    month: int
):
    # Store Pdf with convert_from_path function
    if bank in ['BOP', 'SINOPAC', 'UBOT']:
        pdf_path = join(realpath(dirname(__file__)),'PDF','pdf',f'THÁNG {month}',f'{bank}_rotated.pdf')
    else:
        pdf_path = join(realpath(dirname(__file__)),'PDF','pdf',f'THÁNG {month}',f'{bank}.pdf')
    images = convert_from_path(
        pdf_path=pdf_path,
        poppler_path=r'C:\Users\namtran\poppler-0.68.0\bin'
    )
    # Save pages as images in the pdf
    for i in range(len(images)):
        # Save pages as images in the pdf
        images[i].save(join(realpath(dirname(__file__)), 'PDF', 'img', f'{bank}', f'{bank}_{i}') + '.jpg', 'JPEG')

def detect_table(bank:str):
    img_file = join(realpath(dirname(__file__)), 'PDF', 'img', f'{bank}.jpg')
    # Load iamge, grayscale, adaptive threshold
    image = cv2.imread(img_file)
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
    i = 0
    for c in cnts:
        if len(c) >= 4:
            x, y, w, h = cv2.boundingRect(c)
            roi = image[y:y + h, x:x + w]
            cv2.imwrite(join(realpath(dirname(__file__)), 'PDF', 'img', f'{bank}_{i}_crop.jpg'), roi)
            cv2.rectangle(image, (x, y), (x + w, y + h), (36, 255, 12), 3)
            i += 1
    cv2.imwrite(join(realpath(dirname(__file__)), 'PDF', 'img', f'{bank}_rectangle.jpg'), image)

def runBOP(bank: str):
    now = dt.datetime.now()
    if bank in ['BOP MÓN VAY']:
        img_name = f'{bank}_0.jpg'
    else:
        img_name = f'{bank}.jpg'
    # read text in image using pytesseract
    string = pytesseract.image_to_string(
        join(realpath(dirname(__file__)), 'PDF', 'img', img_name),
        config='--psm 6'
    )
    # Số tài khoản
    accounts = re.findall(r'\d{13}', string)
    # Ngày hiệu lực, Ngày đáo hạn
    DateText = re.findall(r'\b\d{8}\b|\d{4}/\d{2}/\d{2}', string)
    issueDates = []
    expireDates = []
    for i in range(0, len(DateText), 2):
        issueDates.append(DateText[i])
        expireDates.append(DateText[i + 1])

    issueDates = [dt.datetime.strptime(issueDate, '%Y%m%d') for issueDate in issueDates]
    expireDates = [dt.datetime.strptime(expireDate, '%Y%m%d') for expireDate in expireDates]
    # Term Days
    termDays = [(expireDate - issueDate).days for issueDate, expireDate in zip(issueDates, expireDates)]
    # Term Months
    termMonths = [
        (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month for issueDate, expireDate in zip(issueDates, expireDates)
    ]
    # Interest rate
    iTexts = re.findall(r'\d+\.\d*%', string)
    iRates = [float(iText.replace('%','')) / 100 for iText in iTexts]
    # Balance
    balanceStrings = re.findall(r'\b\d+,[\d,]+\d{3}\b', string)
    balances = [float(balanceString.replace(',','')) for balanceString in balanceStrings]
    # Interest amount
    interestAmountStrings = re.findall(r'\b\d+,[\d]+\.\d{2}\b', string)
    interestAmounts = [float(interestRateString.replace(',','')) for interestRateString in interestAmountStrings]
    # Currency
    currency = re.search(r'\bVND|USD\b', string).group()

    dictionary = {
        'AccountNumber': accounts,
        'TermDays': termDays,
        'TermMonths': termMonths,
        'InterestRate': iRates,
        'IssueDate': issueDates,
        'ExpireDate': expireDates,
        'Balance': balances,
        'InterestAmount': interestAmounts
    }
    balanceTable = pd.DataFrame(data=dictionary)
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'BOP')
    balanceTable.insert(-1,'Currency', currency)

    return balanceTable

def runCATHAY(bank: str):
    now = dt.datetime.now()
    # read text in image using pytesseract
    directory = join(realpath(dirname(__file__)), 'PDF', 'img', f'{bank}')
    records = []
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        if not filename.startswith('CATHAY'):
            continue
        string = pytesseract.image_to_string(
            image=f,
            config='--psm 6'
        )
        if 'Principal/Interest Receipt' in string:
            continue
        else:
            data = string
        # Số tài khoản
        account = re.search(r'\b\d[A-Z]{7}\d{7}\b', data).group()
        # Ngày hiệu lực, Ngày đáo hạn
        issueDateText, expireDateText = re.findall(r'\d{4}/\d{2}/\d{2}', data)[:2]
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'\d+\.\d*%', data).group()
        iRate = float(iText.replace('%',''))/100
        # Balance
        balanceString = re.search(r'\b\d+,[\d,]+\d{3}\.\d{2}\b', data).group()
        balance = float(balanceString.replace(',', ''))
        # Interest amount
        interestAmountString = re.search(r'^\b[0|9]|[1-9]+,\d+\.\d{2}\b', data).group()
        interestAmount = float(interestAmountString.replace(',', ''))
        # Currency
        currency = re.search(r'\bVND|USD\b', data).group()
        # Append
        records.append((account, termDays, termMonths, iRate, issueDate, expireDate, balance,
                        interestAmount, currency))
    balanceTable = pd.DataFrame(
        records,
        columns=['AccountNumber', 'TermDays', 'TermMonths', 'InterestRate', 'IssueDate', 'ExpireDate', 'Balance',
                 'InterestAmount', 'Currency']
    )
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'BOP')

    return balanceTable









