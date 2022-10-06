import pandas as pd
import os
import datetime as dt
import argparse
from os.path import join, expanduser
from automation.trading_service.giaodichluukyRework import scrape_ticker_by_exchange

def run(
    startDate=None,
    endDate=None
):
    dateRun = dt.datetime.now().strftime("%d.%m.%Y")

    result = scrape_ticker_by_exchange.run(False)

    user_name = expanduser("~")  # tên thư mục tương ứng với user đang logging
    folder_name = os.path.join(user_name, "BaoCaoTuDong")
    # create folder
    if not os.path.isdir(join(folder_name, dateRun)):
        os.makedirs(join(folder_name, dateRun))

    file_name = f'Báo cáo danh sách mã CK trên bảng điện VCBS {dateRun}.xlsx'
    writer = pd.ExcelWriter(
        join(folder_name, dateRun, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###########################################################################
    ###########################################################################
    ###########################################################################

    header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
        }
    )
    text_incell_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
        }
    )
    worksheet = workbook.add_worksheet('DS Chứng Khoán')
    worksheet.hide_gridlines(option=2)

    worksheet.set_column('A:A',15)
    worksheet.set_column('B:B',10)

    worksheet.write_row('A1',['Ticker','Exchange'],header_format)
    worksheet.write_column('A2',result.index,text_incell_format)
    worksheet.write_column('B2',result['exchange'],text_incell_format)

    writer.close()


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument("-s", "--startDate", required=True, help="input start date with format: dd/mm/YYYY")
    ap.add_argument("-e", "--endDate", required=True, help="input end date with format: dd/mm/YYYY")
    args = vars(ap.parse_args())
    run(args['startDate'], args['endDate'])

