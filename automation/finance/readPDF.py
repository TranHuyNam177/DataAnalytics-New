import re

import pandas as pd

from automation import *
import cv2
import warnings
warnings.filterwarnings("ignore", 'This pattern has match groups')


def rotate_pdf(dir: str, bank: str):
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


def convertPDFtoImage(
        bankName: str,
        month: int
):
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


def detect_table(input_file):
    # Load iamge, grayscale, adaptive threshold
    image = np.array(input_file)
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


def getConfidence(patternList: list, img):
    df = pytesseract.image_to_data(
        image=img,
        config='--psm 6',
        output_type='data.frame'
    )
    lines = df.groupby(['line_num'])['text'].apply(list)
    df['check'] = df['text'].str.contains('(' + '|'.join(patternList) + ')', regex=True)
    df = df.loc[df['check'] == True].reset_index(drop=True)
    if df['conf'].min() < 20:
        # return ra dataframe rỗng thay cho phần gửi mail
        return pd.DataFrame()
    else:
        return df

# chưa rõ mẫu
def runBOP(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        # Số tài khoản
        account = re.search(r'\d{13}', dataText).group()
        # Ngày hiệu lực, Ngày đáo hạn
        issueDateText, expireDateText = re.findall(r'\b\d{8}\b|\d{4}/\d{2}/\d{2}', dataText)
        if '/' in issueDateText or '/' in expireDateText:
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        else:
            issueDate = dt.datetime.strptime(issueDateText, '%Y%m%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y%m%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'\d+\.\d*%', dataText).group()
        iRate = float(iText.replace('%', '')) / 100
        # Balance
        balanceString = re.search(r'\b\d+,[\d,]+\d{3}\b', dataText).group()
        balance = float(balanceString.replace(',', ''))
        # Interest amount
        interestAmountText = re.search(r'\b\d+,[\d]+\.\d{2}\b', dataText).group()
        interestAmount = float(interestAmountText.replace(',', ''))
        # Currency
        currency = re.search(r'\bVND|USD\b', dataText).group()
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


# Done
def runCATHAY(bank: str, month: int):
    # List các trường hợp để nhận biết file hình đúng để xử lý
    now = dt.datetime.now()
    lst_condition = ['(LN4030)', 'N4030']
    frames = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        check = any(c not in dataText for c in lst_condition)
        if check:
            continue
        print(dataText)
        # Số Hợp đồng
        contractNumber = re.search(r'\b\d[A-Z]{7}\d{7}\b|\b[A-Z]{8}\d{7}\b', dataText).group()
        if not contractNumber[0].isdigit() and 'AFOBLNA' in contractNumber:
            if contractNumber[0] == 'Z':
                contractNumber = re.sub(contractNumber[0], '2', contractNumber, 1)
            else:
                contractNumber = re.sub(contractNumber[0], '1', contractNumber, 1)
        # split \n để trả ra list chứa các dòng bắt đầu bằng 'INT '
        dataSplit = [d for d in dataText.split('\n') if 'INT ' in d]
        lst_issueDate = []
        lst_expireDate = []
        lst_termDays = []
        lst_termMonths = []
        lst_iRate = []
        lst_interestAmount = []
        lst_paid = []
        for ele in dataSplit:
            # Ngày hiệu lực, Ngày đáo hạn
            issueDateText, expireDateText = re.findall(r'\d{4}/\d{2}/\d{2}', ele)
            issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
            expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
            lst_issueDate.append(issueDate)
            lst_expireDate.append(expireDate)
            # Term Day
            termDays = (expireDate - issueDate).days
            lst_termDays.append(termDays)
            # Term Month
            termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
            lst_termMonths.append(termMonths)
            # Interest rate
            iText = re.search(r'\b\d{1,2}\.\d+\b|\b\d{1,2}\.\s\d+\b', ele).group()
            if ' ' in iText:
                iText = iText.replace(' ', '')
            iRate = float(iText) / 100
            lst_iRate.append(iRate)
            # Interest amount
            interestAmountText = re.search(r'\b\d{1,2},\d{3}\.\d{2}\b', ele).group()
            interestAmount = float(interestAmountText.replace(',', ''))
            lst_interestAmount.append(interestAmount)
            # Paid
            paid = 0
            lst_paid.append(paid)

        # Amount
        amountText = re.search(r'\b\d+,\s?[\d,]+\d{3}\.\s?\d{2}\b', dataText).group()
        amount = float(amountText.replace(', ', '').replace(',', ''))
        # Remaining, Paid
        if 'CAP' in dataText:
            lst_paid[-1] = amount
        remaining = [amount - p for p in lst_paid]
        # Currency
        currency = re.search(r'\bVND|USD\b', dataText).group()
        dictionary = {
            'ContractNumber': contractNumber,
            'TermDays': lst_termDays,
            'TermMonths': lst_termMonths,
            'InterestRate': lst_iRate,
            'IssueDate': lst_issueDate,
            'ExpireDate': lst_expireDate,
            'Amount': amount,
            'Paid': lst_paid,
            'Remaining': remaining,
            'InterestAmount': lst_interestAmount,
            'Currency': currency
        }
        dataFrame = pd.DataFrame(data=dictionary)
        frames.append(dataFrame)
    balanceTable = pd.concat(frames, ignore_index=True)
    # Date
    if now.hour >= 8:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)  # chạy trong ngày -> xem là số ngày hôm nay
    else:
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'CATHAY')

    return balanceTable

# chưa có mẫu phần trả vay nên chưa set rule được chỗ paid
def runCHANGHWA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        def crop_image(img_file):
            img = Image.fromarray(img_file)
            width, height = img.size
            # Setting the points for cropped image
            bottom = 0.4 * height
            # Cropped image of above dimension
            new_img = img.crop((0, 0, width, bottom))
            return new_img

        # crop image
        img_crop = crop_image(np.array(i))  # convert PIL image to array
        img_crop.show()
        dataText = pytesseract.image_to_string(
            image=img_crop,
            config='--psm 6',
            lang='eng'
        )
        if 'INTEREST PAYMENT NOTICE' not in dataText:
            continue
        dataText = dataText.replace(' ', '')
        print(dataText)
        # Số hợp đồng
        contractNumber = re.search(r'\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3}', dataText).group()
        # Ngày hiệu lực, ngày đáo hạn
        DateText = re.search(r'INTERESTPERIOD:\d{1,2}[A-Z]{3}\.\d{4}TO\d{1,2}[A-Z]{3}\.\d{4}', dataText).group()
        issueDateText, expireDateText = re.findall(r'\d{1,2}[A-Z]{3}\.\d{4}', DateText)
        issueDate = dt.datetime.strptime(issueDateText, '%d%b.%Y')
        expireDate = dt.datetime.strptime(expireDateText, '%d%b.%Y')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'INTERESTRATE:\d{1,2}\.\d+%', dataText).group()
        iRate = float(iText.replace('INTERESTRATE:', '').replace('%', '')) / 100
        # Amount
        amountText = re.search(r'USD\d+,[\d,]+\d{3}', dataText).group()
        amount = float(amountText.replace('USD', '').replace(',', ''))
        # Paid
        paid = 0
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = re.search(
            r'INTERESTOFA/MPERIOD:USD\d{1,2},\d{3}\.\d{2}|INTERESTOFA/MPERIOD:USD\d{1,2},\d{3}', dataText).group()
        interestAmount = float(interestAmountText.replace('INTERESTOFA/MPERIOD:USD', '').replace(',', ''))
        # Currency
        currency = re.search(r'VND|USD', dataText).group()
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'CHANG HWA')

    return balanceTable


# Done
def runESUN(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        if 'Payment Advice' not in dataText:
            continue
        print(dataText)
        # Số hợp đồng
        contractNumber = re.search(r'[A-Z]{3}\d{6}', dataText).group()
        contractNumber = re.sub(r'[A-Z]{3}', '22OBLN', contractNumber)
        # Ngày hiệu lực, ngày đáo hạn
        issueDateText, expireDateText = re.findall(r'\d{4}/\d{2}/\d{2}', dataText)[1:]
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.findall(r'\d{1,2}[.|,]\d+%', dataText)[-1]
        iRate = float(iText.replace('%', '').replace(',', '.')) / 100
        # Currency
        currency = re.search(r'VND|USD', dataText).group()
        # split \n in dataText
        dataSplit = dataText.split('\n')
        amount = ''
        paid = ''
        interestAmount = ''
        remaining = ''
        for d in dataSplit:
            # if 'INTEREST RATE' in d:
            if 'LOAN OUTSTANDING' in d:
                # Amount
                amountText = re.search(r'\d+[,|.][\d(,|.)]+\d{3}\.\d{2}', d).group()
                if amountText.count('.') == 1:
                    amount = float(amountText.replace(',', ''))
                else:
                    amount = float(amountText.replace('.', '', 2))
            elif 'INTEREST DUE' in d:
                # Interest Amount
                interestAmountText = re.search(r'\d{1,2}[,|.]\d{3}\.\d{2}', d).group()
                if interestAmountText.count('.') == 1:
                    interestAmount = float(interestAmountText.replace(',', ''))
                else:
                    interestAmount = float(interestAmountText.replace('.', '', 1))
            elif 'PRINCIPAL REPAYMENT' in d:
                paidText = re.search(r'\d+,[\d,]+\d{3}\.\d{2}|0.00', d).group()
                paid = float(paidText.replace(',', ''))
                remaining = amount - paid
            else:
                continue
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'ESUN')

    return balanceTable


# Done
def runFIRST(bank: str, month: int):
    now = dt.datetime.now()
    lst_condition = ['LOAN REPAYMENT CONFIRMATION', 'RECEIPT']
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        check = any(c in dataText for c in lst_condition)
        if check:
            continue
        # Số hợp đồng
        contractNumber = re.search(r'\b[A-Z]\d[A-Z]\d{7}\b', dataText).group()
        # detect table in image
        tableInImage = detect_table(i)  # return element numpy.ndarray in a list

        def crop_image(img_file):
            img = Image.fromarray(img_file)
            width, height = img.size
            # Setting the points for cropped image
            top = height / 2
            # Cropped image of above dimension
            new_img = img.crop((0, top, width, height))

            return new_img

        # read image table 0
        table_0 = tableInImage[0]
        table_0 = crop_image(table_0)
        dataText = pytesseract.image_to_string(
            image=table_0,
            config='--psm 6'
        )
        print(dataText)
        # Ngày hiệu lực, ngày đáo hạn
        issueDateText, expireDateText = re.findall(r'\d{4}\.\d{1,2}\.\d{1,2}', dataText)
        issueDate = dt.datetime.strptime(issueDateText, '%Y.%m.%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y.%m.%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'\b\d{1,2}\.\d+\b', dataText).group()
        iRate = float(iText) / 100
        # Amount
        amountText = re.search(r'\b\d+,[\d,]+\d{3}\b', dataText).group()
        amount = float(amountText.replace(',', '').replace(' ', ''))
        # read image table 1
        table_1 = tableInImage[1]
        table_1 = crop_image(table_1)
        dataText = pytesseract.image_to_string(
            image=table_1,
            config='preserve_interword_spaces=1 --psm 6 -c tessedit_char_whitelist=0123456789USD,.'
        )
        print(dataText)
        # Paid
        paidText = re.search(r'\d+,[\d,]+\d{3}\b', dataText)
        if paidText is None:
            paid = 0
        else:
            paidText = paidText.group()
            paid = float(paidText.replace(',', '').replace(' ', ''))
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = re.search(r'\d{1,2},\d{3}\.\d{1,2}\b', dataText).group()
        interestAmount = float(interestAmountText.replace(',', ''))
        # Currency
        currency = re.search(r'VND|USD', dataText).group()
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'FIRST')

    return balanceTable


def runFUBON(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        if 'Notice of Interest Payment' not in dataText:
            continue

        dict = {
            'contractNumber': r'\d{14}',
            'date': r'\d{4}/\d{2}/\d{2}',
            'amount': r'\d+,[\d,]+\d{3}\.\d{2}',
            'interestRate': r'\d{1,2}\.\d+',
            'interestAmount': r'\d{1,2},\d{3}\.\d{2}'
        }
        patternList = list(dict.values())
        df = getConfidence(patternList, i)
        date = df['text'].loc[df['text'].str.contains(dict['date'], regex=True)]

# Done
def runKGI(bank: str, month: int):
    """
    1. Chưa có mẫu lúc trả vay nên chưa xử lý được chỗ paid
    2. Thử tìm cách crop hình xem performance có tốt hơn ko
    """
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        # Số hợp đồng
        contractNumbers = re.findall(r'\b[A-Z]{3}\d{7}[A-Z]{2}\b', dataText)
        for contractNumber in contractNumbers:
            for ele in dataText.split('\n'):
                if contractNumber not in ele:
                    continue
                issueDateText, expireDateText = re.findall(r'\d{4}/\d{2}/\d{2}', ele)[2:]
                issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
                expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
                # Term Days
                # Khoảng cách giữa 2 ngày tính tới ngày cuối cùng
                termDays = (expireDate - issueDate).days + 1
                # Term Months
                termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
                # Interest rate
                iText = re.search(r'\b\d{1,2}\.\d+\b', ele).group()
                iRate = round(float(iText) / 100, 7)
                # Amount
                amountText = re.search(r'\b\d+,[\d,]+\d{3}\.\d{2}\b', dataText).group()
                amount = float(amountText.replace(',', '').replace(' ', ''))
                # Paid
                paid = 0
                # Remaining
                remaining = amount - paid
                # Interest Amount
                interestAmountText = re.search(r'\b\d{1,2},\d{3}\.\d{2}\b', ele).group()
                interestAmount = float(interestAmountText.replace(',', ''))
                # Currency
                currency = re.search(r'\bVND|USD\b', dataText).group()
                # Append data
                records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid,
                                remaining, interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'KGI')

    return balanceTable


# Done
def runMEGA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6',
        )
        if 'INTEREST PAYMENT NOTICE' not in dataText:
            continue
        print(dataText)
        # contract number
        contractNumber = re.search(r'\d{14}', dataText)
        # Currency
        currency = re.search(r'VND|USD', dataText).group()
        # Paid
        paidText = re.search(r'\d+,[\d,]+\d{3}\.\d{2}', dataText)
        # Interest rate
        iText = re.search(r'\b\d{1,2}\.\d+', dataText).group()
        iRate = float(iText) / 100
        # Amount
        amountText = re.search(r'\d+,[\d,]+\d{3}', dataText).group()
        amount = float(amountText.replace(',', ''))
        # Paid
        if not paidText:
            paid = 0
        else:
            paid = amount
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = re.search(r'\d{1,2},\d{3}\.\d{2}', dataText).group()
        if ' ' in interestAmountText:
            interestAmount = float(interestAmountText.replace(', ', ''))
        else:
            interestAmount = float(interestAmountText.replace(',', ''))
        # Ngày hiệu lực, ngày đáo hạn
        dataText = dataText.replace(' ', '')
        DateText = re.findall(r'\d{4}/\d{2}/\d{2}', dataText)
        # drop duplicate date in list
        DateText = list(dict.fromkeys(DateText))
        if DateText[0] > DateText[1]:
            issueDateText, expireDateText = DateText[1], DateText[0]
        else:
            issueDateText, expireDateText = DateText[0], DateText[1]
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')

        # check contract numbers
        def crop_image(img_file):
            img = Image.fromarray(img_file)
            width, height = img.size
            # Setting the points for cropped image
            top = height / 2
            bottom = 4 * top / 3
            # Cropped image of above dimension
            new_img = img.crop((0, top, width, bottom))

            return new_img

        if not contractNumber:
            # read image after crop
            img_crop = crop_image(np.array(i))  # convert PIL image to array
            dataText = pytesseract.image_to_string(
                image=img_crop,
                config='--psm 6',
            )
            contractNumber = re.search(r'\d{14}', dataText).group()
        else:
            contractNumber = contractNumber.group()
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'MEGA')

    return balanceTable


# Done
def runSHHK(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        # Số hợp đồng
        contractNumber = re.search(r'\b\d{3}[A-Z]{2}\d{7}\b', dataText).group()
        # Ngày hiệu lực, Ngày đáo hạn
        issueDateText, expireDateText, _ = re.findall(r'\d{4}/\d{2}/\d{2}', dataText)
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'\b\d{1,2}\.\d+\b', dataText).group()
        iRate = round(float(iText) / 100, 7)
        # Amount
        amountText = re.search(r'\b\d+,[\d,]+\d{3}\.\d{2}\b', dataText).group()
        amount = float(amountText.replace(',', ''))
        # Paid
        paidText = ''
        for ele in dataText.split('\n'):
            if 'repay' in ele:
                paidText += ele
        paid = float(paidText.replace('Loan amount repay: USD', '').split(' ')[0].replace(',', ''))
        # Remaining
        remaining = amount - paid
        # Interest amount
        interestAmountText = re.search(r'\b\d{1,2},\d{3}\.\d{2}\b', dataText).group()
        interestAmount = float(interestAmountText.replace(',', ''))
        # Currency
        currency = re.search(r'\bVND|USD\b', dataText).group()
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'SHHK')

    return balanceTable


# rule paid chưa chốt nên không chắc chắn chỗ paid
def runSHINKONG(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6',
        )
        if 'Payment Advice' not in dataText:
            continue
        # contract number
        contractNumber = None

        def crop_image(img_file):
            img = Image.fromarray(img_file)
            width, height = img.size
            # Setting the points for cropped image
            top = height / 4
            bottom = height / 2
            # Cropped image of above dimension
            new_img = img.crop((0, top, width, bottom))
            return new_img

        # read image after crop
        img_crop = crop_image(np.array(i))  # convert PIL image to array
        dataTextImgCrop = pytesseract.image_to_string(
            image=img_crop,
            config='--psm 6',
        )
        dataTextImgCrop = dataTextImgCrop.replace(' ', '')
        print(dataTextImgCrop)
        # Ngày hiệu lực, ngày đáo hạn
        # issueDateText, expireDateText = re.findall(r'\d{4}/\d{2}/\d{2}', dataTextImgCrop)
        dateText = re.search(r'InterestTerm:\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2}', dataTextImgCrop).group()
        issueDateText, expireDateText = re.findall(r'\d{4}/\d{2}/\d{2}', dateText)
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'\d{1,2}\.\d+%', dataTextImgCrop).group()
        iRate = float(iText.replace('%', '')) / 100
        # Amount
        amountText = re.search(r'Principal:USD\d+,[\d,]+\d{3}\.\d{2}', dataTextImgCrop).group()
        amount = float(amountText.split('USD')[-1].replace(',', ''))
        # Paid
        paidText = re.search(r'Totalpayment:USD\d+,[\d,]+\d{3}\.\d{2}', dataTextImgCrop)
        if not paidText:
            paid = 0
        else:
            paid = float(paidText.group().split('USD')[-1].replace(',', ''))
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = re.search(r'Interest:USD\d{1,2},\d{3}\.\d{2}', dataTextImgCrop).group()
        interestAmount = float(interestAmountText.split('USD')[-1].replace(',', ''))
        # Currency
        currency = re.search(r'VND|USD', dataTextImgCrop).group()
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'SHINKONG')

    return balanceTable

# Done
def runSINOPAC(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        # contract number
        contractNumber = re.search(r'\d{8}', dataText).group()
        # Paid
        paidText = re.search(r'\d+,[\d,]+\d{3}\.\d{2}', dataText)
        # Currency
        currency = re.search(r'VND|USD', dataText).group()
        # detect table in image
        tableInImage = detect_table(i)
        tableInImage = Image.fromarray(tableInImage[0])
        dataText = pytesseract.image_to_string(
            image=tableInImage,
            config='--psm 6'
        )
        print(dataText)
        # Ngày hiệu lực, ngày đáo hạn
        DateText = re.findall(r'\d{1,2}[.|,|;]?\s?[A-Z]{3}[.|,|;]?\s?\d{4}', dataText)
        issueDateText, expireDateText = [
            d.replace('. ', '.').replace(', ', '.').replace('; ', '.').replace(' ', '.') for d in DateText
        ]
        issueDate = dt.datetime.strptime(issueDateText, '%d.%b.%Y')
        expireDate = dt.datetime.strptime(expireDateText, '%d.%b.%Y')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'\b\d{1,2}[.|,]\d+\b', dataText).group()
        if ',' in iText:
            iRate = float(iText.replace(',', '.')) / 100
        else:
            iRate = float(iText) / 100
        # Amount
        amountText = re.search(r'\d+,[\d,]+\d{3}', dataText).group()
        amount = float(amountText.replace(',', '').replace(' ', ''))
        # paid
        if not paidText:
            paid = 0
        else:
            paid = amount
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = re.search(r'\d{1,2},\d{3}\.\d{1,2}', dataText).group()
        interestAmount = float(interestAmountText.replace(',', ''))
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'SINOPAC')

    return balanceTable


# còn file tháng 1 chưa xử lý được
def runTAISHIN(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        def crop_image(img_file):
            img = Image.fromarray(img_file)
            width, height = img.size
            # Setting the points for cropped image
            bottom = 0.55 * height
            # Cropped image of above dimension
            new_img = img.crop((0, 0, width, bottom))
            return new_img

        # crop lần 1
        img_crop = crop_image(np.array(i))  # convert PIL image to array
        dataText = pytesseract.image_to_string(
            image=img_crop,
            config='--psm 6',
            lang='eng'
        )
        if 'Interest Notice' not in dataText:
            continue
        print(dataText, "\n=========================\n")
        # Số hợp đồng
        contractNumber = re.search(r'\b[A-Z]{4}\d{2}[A-Z]{2}\d{8}\b', dataText).group()
        # crop lần 2
        img_crop = crop_image(np.array(img_crop))  # convert PIL image to array
        img_crop.show()
        dataText = pytesseract.image_to_string(
            image=img_crop,
            config='--psm 6',
            lang='eng'
        )
        print(dataText)
        # Ngày hiệu lực, ngày đáo hạn
        issueDateText, expireDateText = re.findall(r'\d{4}/\d{2}/\d{2}\b', dataText)
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = re.search(r'\d{1,2}\.\d+%', dataText).group()
        iRate = float(iText.replace('%', '')) / 100
        # Amount
        amountText = re.search(r'USD\d+,\d{3}\b', dataText).group()
        amount = float(amountText.replace('USD', '').replace(',', '')) * 1000
        # Paid
        paid = 0
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = re.search(r'USD\d{1,2},\d{3}\.\d{2}\b', dataText).group()
        interestAmount = float(interestAmountText.split('USD')[-1].replace(',', ''))
        # Currency
        currency = re.search(r'VND|USD', dataText).group()
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'TAISHIN')

    return balanceTable


# Chưa có mẫu lúc trả vay nên chưa xử lý được chỗ paid
def runUBOT(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for i in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=i,
            config='--psm 6'
        )
        if 'Interest Payment Notice' not in dataText:
            continue
        print(dataText)
        # Số hợp đồng
        contractNumber = re.search(r'\b\d{4}[A-Z]{1,2}\d{5,6}\b', dataText).group()
        if 'L0' in contractNumber:
            contractNumber = contractNumber.replace('L0', 'LO')
        issueDateText = ''
        expireDateText = ''
        iText = ''
        interestAmountText = ''
        for ele in dataText.split('\n'):
            ele = ele.replace(' ', '')
            if 'InterestPeriod' in ele:
                # Ngày hiệu lực, Ngày đáo hạn Text
                issueDateText, expireDateText = re.findall(r'\d{4}\.\d{1,2}\.\d{2}', ele)
            elif 'All-inRate' in ele:
                # Interest rate text
                iText = re.search(r'\d{1,2}\.\d+', ele).group()
            elif 'InterestPayableAmount' in ele:
                # Interest amount text
                interestAmountText = re.search(r'\d{1,2},\d{3}\.\d{2}', ele).group()
            else:
                continue
        issueDate = dt.datetime.strptime(issueDateText, '%Y.%m.%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y.%m.%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iRate = round(float(iText) / 100, 7)
        # Amount
        amountText = re.search(r'\d+,\s?[\d,]+\s?\d{3}', dataText).group()
        amount = float(amountText.replace(',', '').replace(' ', ''))
        # Paid
        paid = 0
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmount = float(interestAmountText.replace(',', ''))
        # Currency
        currency = re.search(r'\bVND|USD\b', dataText).group()
        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'UBOT')

    return balanceTable


def runYUANTA(bank: str, month: int):
    # List các trường hợp để nhận biết file hình đúng để xử lý
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    for img in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=img,
            config='--psm 6'
        )
        if 'Interest Payment Notice' not in dataText:
            continue
        print(dataText)
        # contract number
        contractNumber = re.search(r'\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3}-[A-Z]{2}', dataText).group()
        # currency
        currency = re.search(r'VND|USD', dataText).group()

        # detect table in image
        tables = detect_table(img)
        tableInImg = Image.fromarray(tables[0])
        tableInImg.show()

        dictionary = {
            'contractNumber': r'(\d{2}/\d{4}/[A-Z]{2}/[A-Z]{3}-[A-Z]{2})',
            'date': r'(\d{4}/\d{2}/\d{2}-\d{4}/\d{2}/\d{2})',
            'amount': r'(\d{1,3},[,\d]+$)',
            'interestRate': r'(\d{1,2}\.\d+%$)',
            'interestAmount': r'((\d+,?)*\d{3}\.\d{2}$)',
        }
        # convert dictionary values to list
        patternList = list(dictionary.values())
        dfConf = getConfidence(patternList, tableInImg)

        # Ngày hiệu lực, ngày đáo hạn
        dateText = dfConf['text'].loc[dfConf['text'].str.contains(dictionary['date'])].item()
        issueDateText, expireDateText = dateText.split('-')
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = dfConf['text'].loc[dfConf['text'].str.contains(dictionary['interestRate'])].item()
        iRate = float(iText.replace('|','').replace('%','')) / 100
        # Amount
        amountText = dfConf['text'].loc[dfConf['text'].str.contains(dictionary['amount'])].item()
        amount = float(amountText.replace('US$','').replace(',', ''))
        # Paid
        paid = 0
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = dfConf['text'].loc[dfConf['text'].str.contains(dictionary['interestAmount'])].item()
        if ' ' in interestAmountText:
            interestAmount = float(interestAmountText.replace('|US$','').replace(', ', ''))
        else:
            interestAmount = float(interestAmountText.replace('|US$','').replace(',', ''))

        # Append data
        records.append((contractNumber, termDays, termMonths, iRate, issueDate, expireDate, amount, paid, remaining,
                        interestAmount, currency))
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'YUANTA')
    return balanceTable

