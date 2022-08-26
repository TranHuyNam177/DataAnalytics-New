import numpy as np
import pandas as pd
import pytesseract

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


def detect_table(input_file):
    """
    Hàm nhận diện bảng trong 1 image
    """
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


def readImgPytesseractToString(img):
    string = pytesseract.image_to_string(
        image=img,
        config=f'--psm 6'
    )

    return string


def readImgPytesseractToDataframe(img):
    df = pytesseract.image_to_data(
        image=img,
        config=f'--psm 11',
        output_type='data.frame'
    )
    df = df[df['conf'] != -1]

    return df


# chưa có mẫu đã trả vay nên chưa làm được chỗ Paid, remaining
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
        # tableInImg.show()

        dictionary = {
            'date': r'(\d{4}/\d{2}/\d{2}-\d{4}/\d{2}/\d{2})',
            'amount': r'([1-9]{1,3}[,\d]+$)',
            'interestRate': r'(\d{1,2}\.\d+%$)',
            'interestAmount': r'((\d+,\s?)*\d{3}\.\d{2}$)'
        }
        # convert dictionary values to list
        patternList = list(dictionary.values())
        # run def getConfidence --> return to dataframe
        df = readImgPytesseractToDataframe(tableInImg)
        if df.empty:
            # gửi mail
            balanceTable = df
        # Ngày hiệu lực, ngày đáo hạn
        dateText = df['text'].loc[df['text'].str.contains(dictionary['date'])].item()
        issueDateText, expireDateText = dateText.split('-')
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = df['text'].loc[df['text'].str.contains(dictionary['interestRate'])].item()
        iRate = float(iText.replace('|', '').replace('%', '')) / 100
        # Amount
        amountText = df['text'].loc[df['text'].str.contains(dictionary['amount'])].item()
        amount = float(amountText.replace('US$', '').replace(',', ''))
        # Paid
        paid = 0
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = df['text'].loc[df['text'].str.contains(dictionary['interestAmount'])].item()
        if ' ' in interestAmountText:
            interestAmount = float(interestAmountText.replace('|US$', '').replace(', ', ''))
        else:
            interestAmount = float(interestAmountText.replace('|US$', '').replace(',', ''))

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


def runENTIE(bank: str, month: int):
    # List các trường hợp để nhận biết file hình đúng để xử lý
    now = dt.datetime.now()
    condition = ['ADVICE OF', 'REPAYMENT']
    records = []
    images = convertPDFtoImage(bank, month)
    for img in images:
        # read text in image using pytesseract
        dataText = pytesseract.image_to_string(
            image=img,
            config='--psm 6'
        )
        check = any(c in dataText for c in condition)
        if not check:
            continue
        # currency
        currency = re.search(r'VND|USD', dataText).group()
        dictionary = {
            'contractNumber': r'\d{11}-\d{2}',
            'date': r'\d{4}/\d{2}/\d{2} TO \d{4}/\d{2}/\d{2}\b',
            'amount': r'[1-9]{1,3}[,\d]+\.\d{2}$',
            'interestRate': r'\d{1,2}\.\d+%$',
            'interestAmount': r'[[1-9]+,?]*\d{3}\.\d{2}$'
        }
        # convert dictionary values to list
        patternList = list(dictionary.values())
        # run def getConfidence --> return to dataframe
        df = readImgPytesseractToDataframe(img)
        df['regex'] = df['text'].str.extract('(' + '|'.join(patternList) + ')')
        if df.empty:
            # gửi mail
            balanceTable = df

        # contract number
        contractNumber = df['regex'].loc[df['regex'].str.contains(dictionary['contractNumber'])].item() + ' USD1'
        # Ngày hiệu lực, ngày đáo hạn
        dateText = df['regex'].loc[df['regex'].str.contains(dictionary['date'])].item()
        issueDateText, expireDateText = dateText.replace(' ', '').split('TO')
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Interest rate
        iText = df['regex'].loc[df['regex'].str.contains(dictionary['interestRate'])].item()
        iRate = float(iText.replace('%', '')) / 100
        # Amount
        amountText = df['text'].loc[df['text'].str.contains(dictionary['amount'])].iloc[0]
        amount = float(amountText.replace('USD', '').replace(',', ''))
        # Paid
        paid = 0
        # Remaining
        remaining = amount - paid
        # Interest Amount
        interestAmountText = df['regex'].loc[df['regex'].str.contains(dictionary['interestAmount'])].iloc[0]
        if ' ' in interestAmountText:
            interestAmount = float(interestAmountText.replace(', ', ''))
        else:
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
    balanceTable.insert(1, 'Bank', 'ENTIE')
    return balanceTable


######################################################################

# Done
def runCATHAY(bank: str, month: int):
    # List các trường hợp để nhận biết file hình đúng để xử lý
    now = dt.datetime.now()
    images = convertPDFtoImage(bank, month)
    frames = []
    for img in images:
        try:
            df = readImgPytesseractToDataframe(img)
            # group by dataframe theo block_num
            groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
            groupByConf = df.groupby(['block_num'])['conf'].sum()
            df = pd.concat([groupByText, groupByConf], axis=1)
            # check if LN4030 in file
            check = df['text'].isin(['(LN4030)', 'N4030']).any()
            if not check:
                continue

            patternDict = {
                r'\bINT\b|\bCAP\b': 'condition',
                r'\d[A-Z]{7}\d{7}\b|\b[A-Z]{8}\d{7}': 'contractNumber',
                r'\d{4}/\d{2}/\d{2}~\d{4}/\d{2}/\d{2}': 'date',
                r'[1-9]+[0-9]{0,},[\d,]+\d{3}\.\d{2}': 'amount',
                r'\b\d{1,2}\.\d+%': 'interestRate',
                r'USD': 'currency'
            }
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
            termMonths = [
                (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month for issueDate, expireDate
                in zip(issueDate, expireDate)
            ]
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'CATHAY')

    return balanceTable


# Done
def runSHHK(bank: str, month: int):
    now = dt.datetime.now()
    frames = []
    images = convertPDFtoImage(bank, month)

    def crop_image(img_file):
        img = Image.fromarray(img_file)
        width, height = img.size
        # Setting the points for cropped image
        top = height * 0.2
        bottom = 0.4 * height
        # Cropped image of above dimension
        new_img = img.crop((0, top, width, bottom))
        return new_img

    for img in images:
        try:
            img_crop = crop_image(np.array(img))  # convert PIL image to array
            img_crop.show()
            df = readImgPytesseractToDataframe(img_crop)
            # group by dataframe theo block_num
            groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ' '.join(list(x)))
            groupByConf = df.groupby(['block_num'])['conf'].sum()
            dfGroupBy = pd.concat([groupByText, groupByConf], axis=1)

            patternDict = {
                r'[A-Z]{2}\d{7}': 'contractNumber',
                r'\d{4}/\d{2}/\d{2}[~|-]\d{4}/\d{2}/\d{2}|\d{4}/\d{2}/\d{2}': 'date',
                r'[1-9]+[0-9]{0,},[\d,]+\d{3}\.\d{2}': 'amount',
                r'\b\d{1,2}\.\d+%': 'interestRate',
                r'USD': 'currency'
            }
            # convert dictionary values to list
            patternList = list(patternDict.keys())
            # kiểm tra xem giá trị nào khớp với các pattern trong pattern list
            df['check'] = df['text'].str.contains('(' + '|'.join(patternList) + ')')
            df = df.loc[df['check']]
            df['text'] = df['text'].str.extract('(' + '|'.join(patternList) + ')')
            df['regex'] = df[['text']].replace({'text': patternDict}, regex=True)

            # contractNumber
            contractNumber = df['text'].loc[df['regex'] == 'contractNumber'].item()
            contractNumber = contractNumber.replace(contractNumber[0:2], '888LN')
            # currency
            currency = df['text'].loc[df['regex'] == 'currency'].tolist()
            if len(currency) > 1:
                currency = currency[0]
            # Ngày hiệu lực, Ngày đáo hạn
            dateText = df['text'].loc[df['regex'] == 'date'].tolist()
            issueDates, expireDates = [], []
            if not any(c in dateText[0] for c in ['~', '-']):
                issueDateText, expireDateText, _ = dateText
                issueDates.append(issueDateText)
                expireDates.append(expireDateText)
            else:
                if '~' in dateText[0]:
                    issueDateText, expireDateText = dateText[0].split('~')
                else:
                    issueDateText, expireDateText = dateText[0].split('-')
                issueDates.append(issueDateText)
                expireDates.append(expireDateText)
            issueDate = [dt.datetime.strptime(issueDate, '%Y/%m/%d') for issueDate in issueDates]
            expireDate = [dt.datetime.strptime(expireDate, '%Y/%m/%d') for expireDate in expireDates]
            # Term Days
            termDays = [(expireDate - issueDate).days for issueDate, expireDate in zip(issueDate, expireDate)]
            # Term Months
            termMonths = [
                (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month for issueDate, expireDate
                in zip(issueDate, expireDate)
            ]
            # Amount
            amount = df['text'].loc[df['regex'].str.contains('amount')].tolist()
            amount = float(min(amount).replace(',', ''))
            # interest rate
            interestRateText = df['text'].loc[df['regex'].str.contains('interestRate')].tolist()
            interestRate = [float(i.replace(',', '.').replace('%', '')) / 100 for i in interestRateText]
            # Interest amount
            interestAmount = [amount * termDay * iRate / 360 for termDay, iRate in zip(termDays, interestRate)]
            # Paid
            paid = dfGroupBy['text'].loc[dfGroupBy['text'].str.contains('Loan amount repay')].tolist()
            paid = [float(p.replace('Loan amount repay: USD', '').replace(' “', '').replace(',', '')) for p in paid]
            # Remaining
            remaining = [amount - p for p in paid]

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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'SHHK')

    return balanceTable


# Done
def runKGI(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)

    def crop_image(img_file):
        img = Image.fromarray(img_file)
        width, height = img.size
        # Setting the points for cropped image
        top = height * 0.3
        bottom = 0.35 * height
        # Cropped image of above dimension
        new_img = img.crop((0, top, width, bottom))
        return new_img

    for img in images:
        img_crop = crop_image(np.array(img))  # convert PIL image to array
        df = readImgPytesseractToDataframe(img_crop)
        # group by dataframe theo block_num
        groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
        groupByConf = df.groupby(['block_num'])['conf'].sum()
        dfGroupBy = pd.concat([groupByText, groupByConf], axis=1)

        patternDict = {
            r'[A-Z]{3}\d{7}[A-Z]{2}': 'contractNumber',
            r'\d{4}/\d{2}/\d{2}': 'date',
            r'\b\d{1,2}\.\d+': 'interestRate',
            r'USD': 'currency'
        }
        # convert dictionary values to list
        patternList = list(patternDict.keys())

        # kiểm tra xem giá trị nào khớp với các pattern trong pattern list
        df['check'] = df['text'].str.contains('(' + '|'.join(patternList) + ')')
        df = df.loc[df['check']]
        df['text'] = df['text'].str.extract('(' + '|'.join(patternList) + ')')
        df['regex'] = df[['text']].replace({'text': patternDict}, regex=True)

        # contractNumber
        contractNumber = df['text'].loc[df['regex'] == 'contractNumber'].item()
        # currency
        currency = df['text'].loc[df['regex'] == 'currency'].item()
        # Ngày hiệu lực, Ngày đáo hạn
        dateText = df['text'].loc[df['regex'] == 'date'].tolist()[2:]
        issueDateText, expireDateText = dateText
        issueDate = dt.datetime.strptime(issueDateText, '%Y/%m/%d')
        expireDate = dt.datetime.strptime(expireDateText, '%Y/%m/%d')
        # Term Days
        termDays = (expireDate - issueDate).days + 1
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # Amount
        amount = dfGroupBy['text'].loc[
            dfGroupBy['text'].str.contains(r'[1-9]+[0-9]{0,},[\d,]+\d{3}\.\d{2}', regex=True)].item()
        amount = float(amount.replace(',', ''))
        # interest rate
        interestRateText = df['text'].loc[df['regex'].str.contains('interestRate')].item()
        interestRate = float(interestRateText.replace(',', '.')) / 100
        # Interest amount
        interestAmount = amount * termDays * interestRate / 360
        # Paid
        paid = 0
        # Remaining
        remaining = amount - paid
        # Append data
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid,
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
def runSINOPAC(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)

    for img in images:
        df = readImgPytesseractToDataframe(img)
        df = df.loc[df['conf'] > 40]
        # contractNumber
        contractNumber = df['text'].str.extract('(\d{8})').dropna().iloc[0].item()
        # currency
        currency = df['text'].str.extract('(USD)').dropna().iloc[0].item()
        # check condition for "paid"
        paid = df['text'].str.contains(r'([1-9]+[0-9]{0,},[\d,]+\d{3}\.\d{2})', regex=True)
        # detect table in image
        tableInImage = detect_table(img)
        tableInImage = Image.fromarray(tableInImage[0])
        tableInImage.show()
        df = readImgPytesseractToDataframe(tableInImage)
        # group by dataframe theo block_num
        groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ''.join(list(x)))
        groupByConf = df.groupby(['block_num'])['conf'].sum()
        df = pd.concat([groupByText, groupByConf], axis=1)

        patternDict = {
            r'[A-Z]{3}\d{7}[A-Z]{2}': 'contractNumber',
            r'\d{1,2}[.|,|;|:]?\s?[A-Z]{3}[.|,|;|:]?\s?\d{4}': 'date',
            r'[1-9]+[0-9]{0,},[\d,]+\d{3}': 'amount',
            r'\b\d{1,2}[.|,]\d+%': 'interestRate'
        }
        # convert dictionary values to list
        patternList = list(patternDict.keys())

        # kiểm tra xem giá trị nào khớp với các pattern trong pattern list
        df['text'] = df['text'].apply(lambda x: x.replace('..', '.'))
        df['check'] = df['text'].str.contains('(' + '|'.join(patternList) + ')')
        df = df.loc[df['check']]
        df['text'] = df['text'].str.extract('(' + '|'.join(patternList) + ')')
        df['regex'] = df[['text']].replace({'text': patternDict}, regex=True)
        # Ngày hiệu lực, Ngày đáo hạn
        issueDateText, expireDateText = df['text'].loc[df['regex'] == 'date'].tolist()
        issueDateText = issueDateText.replace(' ', '').replace('.', '').replace(',', '').replace(';', '')
        expireDateText = expireDateText.replace(' ', '').replace('.', '').replace(',', '').replace(';', '').replace(':',                                                                                                      '')
        issueDate = dt.datetime.strptime(issueDateText, '%d%b%Y')
        expireDate = dt.datetime.strptime(expireDateText, '%d%b%Y')
        # Term Days
        termDays = (expireDate - issueDate).days
        # Term Months
        termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month
        # amount
        amountText = df['text'].loc[df['regex'].str.contains('amount')].item()
        amount = float(amountText.replace(',', ''))
        # interest rate
        interestRateText = df['text'].loc[df['regex'].str.contains('interestRate')].item()
        interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100
        # Interest amount
        interestAmount = amount * termDays * interestRate / 360
        # Paid
        if paid.any():
            paid = amount
        else:
            paid = 0
        # Remaining
        remaining = amount - paid
        # Append data
        records.append((contractNumber, termDays, termMonths, interestRate, issueDate, expireDate, amount, paid,
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
    balanceTable.insert(1, 'Bank', 'SINOPAC')

    return balanceTable


def _findTopLeftPoint(containingImage, pdfImage):
    containingImage = np.array(containingImage)  # đảm bảo image để được đưa về numpy array
    pdfImage = np.array(pdfImage)  # đảm bảo image để được đưa về numpy array
    matchResult = cv.matchTemplate(containingImage, pdfImage, cv.TM_CCOEFF)
    _, _, _, topLeft = cv.minMaxLoc(matchResult)
    return topLeft[0], topLeft[1]  # cho compatible với openCV


def _findCoords(pdfImage, name, bank):
    if name == 'condition':
        fileName = 'condition.png'
    elif name == 'amount':
        fileName = 'amount.png'
    elif name == 'interestRate':
        fileName = 'interestRate.png'
    elif name == 'date':
        fileName = 'date.png'
    elif name == 'contractNumber':
        fileName = 'contractNumber.png'
    elif name == 'paid':
        fileName = 'paid.png'
    elif name == 'currency':
        fileName = 'currency.png'
    else:
        raise ValueError(
            'colName must be either "amount" or '
            '"iRate" or '
            '"period" or '
            '"contractNumber" or '
            '"paid" or '
            '"currency"'
        )
    containingPath = os.path.join(os.path.dirname(__file__), 'bank_img', f'{bank}', fileName)
    containingImage = cv.imread(containingPath,0)
    w, h = containingImage.shape[::-1]
    top, left = _findTopLeftPoint(pdfImage, containingImage)
    if bank == 'MEGA':
        if name in ['paid', 'contractNumber', 'currency']:
            top = top
            left = left
            right = left + h
            return pdfImage[left:right, top:]
        elif name == 'condition':
            top = top
            bottom = top + w
            left = left
            right = left + h
            return pdfImage[left:right, top:bottom]
        else:
            bottom = top + w
            right = left + h * 2
            top = top
            left = left + h
            return pdfImage[left:right, top:bottom]


def runMEGA(bank: str, month: int):
    now = dt.datetime.now()
    records = []
    images = convertPDFtoImage(bank, month)
    if bank != 'MEGA':
        bank = 'MEGA'

    patternDict = {
        'amount': r'([1-9]+[0-9]{0,},[\d,]+\d{3})',
        'interestRate': r'(\d{1,2}[.|,]\d+%)',
        'date': r'(\d{4}/\d{2}/\d{2}[~|-]+?\d{4}/\d{2}/\d{2})',
        'contractNumber': r'(\d{14})',
        'paid': r'([1-9]+[0-9]{0,},[\d,]+\d{3}\.\d{2})',
        'currency': r'(VND|USD)'
    }

    for img in images:
        # convert PIL to np array
        img = np.array(img)
        # scale image to gray
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # using adaptiveThreshold algo
        img = cv.adaptiveThreshold(img, 255, cv.ADAPTIVE_THRESH_MEAN_C, cv.THRESH_BINARY, 7, 5)

        Image.fromarray(img).show()

        # check image with condition
        if 'INTEREST PAYMENT NOTICE' not in readImgPytesseractToString(img):
            continue
        try:
            # amount
            imgAmount = _findCoords(np.array(img), 'amount', bank)
            Image.fromarray(imgAmount).show()
            dfAmount = readImgPytesseractToDataframe(imgAmount)
            dfAmount['check'] = dfAmount['text'].str.contains(patternDict['amount'])
            dfAmount = dfAmount.loc[(dfAmount['check']) & (dfAmount['conf'] > 10)]
            if dfAmount.empty:
                return pd.DataFrame()
            dfAmount['regex'] = dfAmount['text'].str.extract(patternDict['amount'])
            amountText = dfAmount.loc[dfAmount.index[0], 'regex']
            amount = float(amountText.replace(',','').replace('.',''))

            # interest rate
            imgInterestRate = _findCoords(np.array(img), 'interestRate', bank)
            Image.fromarray(imgInterestRate).show()
            dfInterestRate = readImgPytesseractToDataframe(imgInterestRate)
            dfInterestRate['check'] = dfInterestRate['text'].str.contains(patternDict['interestRate'])
            dfInterestRate = dfInterestRate.loc[(dfInterestRate['check']) & (dfInterestRate['conf'] > 10)]
            if dfInterestRate.empty:
                return pd.DataFrame()
            dfInterestRate['regex'] = dfInterestRate['text'].str.extract(patternDict['interestRate'])
            interestRateText = dfInterestRate.loc[dfInterestRate.index[0], 'regex']
            interestRate = float(interestRateText.replace(',', '.').replace('%', '')) / 100

            # ngày hiệu lực, ngày đáo hạn
            imgDate = _findCoords(np.array(img), 'date', bank)
            # imgDate = imgDate[:-5, :]
            Image.fromarray(imgDate).show()
            dfDate = readImgPytesseractToDataframe(imgDate)
            dfDate['check'] = dfDate['text'].str.contains(patternDict['date'])
            dfDate = dfDate.loc[(dfDate['check']) & (dfDate['conf'] > 10)]
            if dfDate.empty:
                return pd.DataFrame()
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
            termMonths = (expireDate.year - issueDate.year) * 12 + expireDate.month - issueDate.month

            # contract number
            imgContractNumber = _findCoords(np.array(img), 'contractNumber', bank)
            Image.fromarray(imgContractNumber).show()
            dfContractNumber = readImgPytesseractToDataframe(imgContractNumber)
            dfContractNumber['check'] = dfContractNumber['text'].str.contains(patternDict['contractNumber'])
            dfContractNumber = dfContractNumber.loc[(dfContractNumber['check']) & (dfContractNumber['conf'] > 10)]
            if dfContractNumber.empty:
                return pd.DataFrame()
            dfContractNumber['regex'] = dfContractNumber['text'].str.extract(patternDict['contractNumber'])
            contractNumber = dfContractNumber.loc[dfContractNumber.index[0], 'regex']

            # currency
            imgCurrency = _findCoords(np.array(img), 'currency', bank)
            Image.fromarray(imgCurrency).show()
            dfCurrency = readImgPytesseractToDataframe(imgCurrency)
            dfCurrency['check'] = dfCurrency['text'].str.contains(patternDict['currency'])
            dfCurrency = dfCurrency.loc[(dfCurrency['check']) & (dfCurrency['conf'] > 10)]
            if dfCurrency.empty:
                return pd.DataFrame()
            dfCurrency['regex'] = dfCurrency['text'].str.extract(patternDict['currency'])
            currency = dfCurrency.loc[dfCurrency.index[0], 'regex']

            # paid
            imgPaid = _findCoords(np.array(img), 'paid', bank)
            Image.fromarray(imgPaid).show()
            dfPaid = readImgPytesseractToDataframe(imgPaid)
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
        except (Exception,):
            return pd.DataFrame()
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
        d = (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0,
                                                 microsecond=0)  # chạy đầu ngày -> xem là số ngày hôm trước
    balanceTable.insert(0, 'Date', d)
    # Bank
    balanceTable.insert(1, 'Bank', 'MEGA')

    return balanceTable

