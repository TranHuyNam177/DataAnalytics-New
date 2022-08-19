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
        config='--psm 11',
        output_type='data.frame'
    )
    df = df[df['conf'] != -1]
    groupByText = df.groupby(['block_num'])['text'].apply(lambda x: ' '.join(list(x)))
    groupByConf = df.groupby(['block_num'])['conf'].sum()
    df = pd.concat([groupByText,groupByConf], axis=1)
    df['check'] = df['text'].str.contains('(' + '|'.join(patternList) + ')')
    df = df.loc[df['check'] == True]
    condition = df.loc[(df['check'] == True) & (df['conf'] < -50)]
    if not condition.empty:
        df = pd.DataFrame()
    return df.reset_index(drop=True)