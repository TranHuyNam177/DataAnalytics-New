from automation.trading_service.giaodichluuky import *
from news_collector import scrape_vietstock_trading_price


def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    # Lấy danh sách hủy niêm yết
    huyniemyet = pd.read_excel(
        r'C:\Users\hiepdang\Shared Folder\Trading Service\Huy Niem Yet\DS mã hủy đăng ký.xls',
        usecols=['SYMBOL'],
        squeeze=True,
    ).to_list()

    # Lấy từ file của Sở
    day = period.split('.')[2]
    month = period.split('.')[1]
    year = period.split('.')[0]
    price_exchange_folder = r'C:\Users\hiepdang\Shared Folder\Trading Service\Price'

    ## move file to correct folders
    unmoved_files = pd.Series([file for file in listdir(price_exchange_folder) if file.endswith('.xlsx')],dtype=object)
    if unmoved_files.shape[0]!=0:
        dates = unmoved_files.str.split('.').str.get(0).str[-8:]
        dates = dates.str[-2:]+'.'+dates.str[4:6]+'.'+dates.str[:4]
        for d,file in zip(dates,unmoved_files):
            if not os.path.isdir(join(price_exchange_folder,d)):
                os.mkdir(join(price_exchange_folder,d))
            shutil.move(join(price_exchange_folder,file),join(price_exchange_folder,d,file))

    def get_file():
        folder = join(price_exchange_folder,f'{day}.{month}.{year}')
        price_hose = pd.read_excel(join(folder,f'SECURITIES_INFO_HSX_{year}{month}{day}.xlsx'))
        price_hnx = pd.read_excel(join(folder,f'SECURITIES_INFO_HNX_{year}{month}{day}.xlsx'))
        price_upcom = pd.read_excel(join(folder,f'SECURITIES_INFO_UPCOM_{year}{month}{day}.xlsx'))
        return price_hose,price_hnx,price_upcom

    while True:
        try:
            price_hose,price_hnx,price_upcom = get_file()
            break
        except FileNotFoundError:
            print('No files received from Stock Exchanges. Still listening...')
            time.sleep(15)

    price_exchange = pd.concat([price_hose,price_hnx,price_upcom])
    price_exchange.rename({'SYMBOL':'ticker','TRADEPLACE':'exchange','AVGPRICE':'price'},axis=1,inplace=True)
    price_exchange = price_exchange.loc[price_exchange['ticker'].str.len()==3]
    price_exchange['exchange'] = price_exchange['exchange'].map({1:'HOSE',2:'HNX',5:'UPCOM'})
    price_exchange = price_exchange.loc[price_exchange['ticker'].str.len()==3]
    price_exchange = price_exchange.loc[~price_exchange['ticker'].isin(huyniemyet)]
    price_exchange.set_index('ticker',inplace=True)

    # Lấy data từ VietStock
    price_vietstock = scrape_vietstock_trading_price.run(period)
    price_vietstock = price_vietstock.loc[price_vietstock.index.str.len()==3]
    price_vietstock = price_vietstock.loc[~price_vietstock.index.isin(huyniemyet)]
    price_vietstock['exchange'].replace('HSX','HOSE',inplace=True)
    price_vietstock['exchange'].replace('UPCoM','UPCOM',inplace=True)
    original_price_vietstock = price_vietstock['price'].copy()

    # bỏ đi trước những thằng trùng
    ticker_idx = price_exchange.index.union(price_vietstock.index)
    price_exchange = price_exchange.reindex(ticker_idx)
    price_vietstock = price_vietstock.reindex(ticker_idx)
    diff = price_vietstock['price']!=price_exchange['price']
    price_exchange = price_exchange.loc[diff.index[diff]]
    price_vietstock = price_vietstock.loc[diff.index[diff]]

    def modified_round(price):  # theo rule lam tron cua DVKH
        if price%100==50:  # 11850 -> 11900, 125550 -> 125600
            result = price+50
        else:
            result = np.round(price,-2)  # 11840 -> 11800, 11860 -> 11900
        return result

    price_vietstock.loc[price_vietstock['exchange'].isin(['HOSE','HNX']),'price'] = np.round(price_vietstock['price'],0)
    price_vietstock.loc[price_vietstock['exchange']=='UPCOM','price'] = price_vietstock['price'].map(modified_round)

    result = price_exchange['price'].compare(price_vietstock['price'],keep_equal=True)
    result.rename({'self':'sogiaodich','other':'vietstock'},axis=1,inplace=True)

    # create report table
    tickers = result.index.to_list()
    report_table = pd.DataFrame(
        index=pd.RangeIndex(1,len(tickers)+1,name='stt'),
        columns=['ticker','exchange','price_sogiaodich','price_vietstock','diff','note']
    )
    report_table['ticker'] = tickers
    report_table['exchange'] = report_table['ticker'].map(price_exchange['exchange'])
    report_table['price_sogiaodich'] = report_table['ticker'].map(price_exchange['price'])
    report_table['price_vietstock'] = report_table['ticker'].map(original_price_vietstock)
    report_table['diff'] = report_table['price_sogiaodich']-report_table['price_vietstock']
    report_table['note'] = ''
    report_table.fillna('',inplace=True)

    # =========================================================================
    # Write to Báo cáo check giá
    file_name = f'Báo cáo check giá {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    # =========================================================================
    # Write to sheet CheckGia
    sheet = workbook.add_worksheet('CheckGia')
    sheet.hide_gridlines(option=2)
    # set column width
    sheet.set_column('A:A',5)
    sheet.set_column('B:F',11)
    sheet.set_column('G:G',82)
    sheet.set_row(1,21)
    sheet.set_row(5,37)
    headline_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':14,
            'align':'center',
            'valign':'vcenter',
        }
    )
    date_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'vcenter',
        }
    )
    header_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'bg_color':'#ddebf7',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_center_text_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_left_text_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_value_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'valign':'vcenter',
        }
    )
    sheet.merge_range('A2:G2','BÁO CÁO DANH SÁCH CÁC MÃ CHỨNG KHOÁN CHÊNH LỆCH GIÁ ĐÓNG CỬA',headline_fmt)
    sheet.merge_range('A3:G3',f'{day}.{month}.{year}',date_fmt)
    sheet.write_row('A6',['STT','Mã CK','Sàn GD','Giá từ Sở','Giá từ Vietstock','Chênh lệch','GHI CHÚ'],header_fmt)
    sheet.write_column('A7',report_table.index,incell_center_text_fmt)
    sheet.write_column('B7',report_table['ticker'],incell_center_text_fmt)
    sheet.write_column('C7',report_table['exchange'],incell_center_text_fmt)
    sheet.write_column('D7',report_table['price_sogiaodich'],incell_value_fmt)
    sheet.write_column('E7',report_table['price_vietstock'],incell_value_fmt)
    sheet.write_column('F7',report_table['diff'],incell_value_fmt)
    sheet.write_column('G7',report_table['note'],incell_left_text_fmt)

    writer.close()

    # =========================================================================
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
