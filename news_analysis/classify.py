from request.stock import *
from request import *
from automation.risk_management import get_info, dept_folder


def FilterNewsByKeywords(
    *tickers,
    hours:int,
):

    """
    This function pick out important news (first layer filter)
    :param tickers: tickers to check, default is all tickers in margin list
    :param hours: how many hours to scan news
    """

    start = time.time()

    mlist = internal.mlist('all')
    if tickers == (): # empty tuple
        tickers = mlist

    since = (dt.datetime.now()- dt.timedelta(hours=hours)).strftime('%Y-%m-%d %H:%M:%S')
    data = pd.read_sql(
        f"""
        SELECT
            [Time],
            [Ticker],
            CONCAT([Title],'\n',[Description],'\n',[Body]) [Paragraphs],
            CONCAT('=HYPERLINK("',[URL],'","Link")') [URL]
        FROM [TinChungKhoan]
        WHERE [Time] >= '{since}' AND [Ticker] <> ''
        """,
        connect_DWH_ThiTruong
    )
    data['Paragraphs'] = data['Paragraphs'].str.split('\n')

    def tickerFilter(x,ticker):
        if ticker in x:
            return True
        else:
            return False

    with open(join(dirname(__file__),'Keywords'),encoding='utf8',) as file:
        badWords = file.read().split('\n')

    def badNewsIdentifier(content):
        appearedKeyWords = []
        for word in badWords:
            if word in content.lower(): # đưa hết về chữ thường
                appearedKeyWords.append(word)
        if len(appearedKeyWords) > 0:
            return 'Có Thể Quan Trọng', ', '.join(appearedKeyWords)
        else:
            return '', ''

    tableTickers = []
    tableTimes = []
    tableContents = []
    tableKeyWord = []
    tableClasses = []
    tableURLs = []
    for ticker in tickers:
        tickerMask = data['Ticker'].apply(tickerFilter,ticker=ticker)
        tickerData = data.loc[tickerMask].reset_index(drop=True)
        for row in tickerData.index:
            # Time & URL
            resultTime = tickerData.loc[row,'Time']
            resultURL = tickerData.loc[row,'URL']
            # Select paragraph & classify
            paragraphs = tickerData.loc[row,'Paragraphs']
            pickedParagraphs = itertools.compress(paragraphs,(ticker in p for p in paragraphs))
            resultContent = '\n'.join(pickedParagraphs)
            resultClass, resultKeyWords = badNewsIdentifier(resultContent)
            # Fill in
            tableTickers.append(ticker)
            tableTimes.append(resultTime)
            tableContents.append(resultContent)
            tableKeyWord.append(resultKeyWords)
            tableClasses.append(resultClass)
            tableURLs.append(resultURL)

    resultTable = pd.DataFrame({
        'Ticker': tableTickers,
        'Time': tableTimes,
        'Content': tableContents,
        'PredictedClass': tableClasses,
        'KeyWords': tableKeyWord,
        'URL': tableURLs,
    })
    resultTable = resultTable.sort_values(['PredictedClass','Ticker','Time'],ascending=[False,True,False],ignore_index=True)

    # Write to Excel
    info = get_info('daily',dt.datetime.now())
    period = info['period']
    t0_date = info['end_date'].replace('/','-')
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    if dt.datetime.now().time() < dt.time(hour=12):
        session = 'Buổi Sáng'
    else:
        session = 'Buổi Chiều'

    fdate = dt.datetime.strptime(t0_date,'%Y.%m.%d').strftime('%d.%m.%Y')
    file_name = f'Tin Tức {session} {fdate}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ## Write sheet MO TAi KHOAN
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'bg_color':'#99CCFF',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    time_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'num_format':'yyyy-mm-dd hh:mm:ss',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',20)
    worksheet.set_column('C:C',55)
    worksheet.set_column('D:D',0) # ẩn cột Keywords
    worksheet.set_column('E:E',18)
    worksheet.set_column('F:F',12)

    headers = [
        'Mã Cổ Phiếu',
        'Thời Gian',
        'Nội Dung',
        'Keywords',
        'Loại Tin',
        'Link',
    ]
    worksheet.write_row('A1',headers,header_format)
    worksheet.write_column('A2',resultTable['Ticker'],text_center_format)
    worksheet.write_column('B2',resultTable['Time'],time_format)
    worksheet.write_column('C2',resultTable['Content'],text_left_format)
    worksheet.write_column('D2',resultTable['KeyWords'],text_left_format)
    worksheet.write_column('E2',resultTable['PredictedClass'],text_center_format)
    worksheet.write_column('F2',resultTable['URL'],text_center_format)

    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')




