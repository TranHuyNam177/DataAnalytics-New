from automation.trading_service.giaodichluuky import *
from news_collector import scrape_ticker_by_exchange


def run():
    start = time.time()
    info = get_info('daily',None)
    period = info['period']
    folder_name = 'DanhSachCoPhieuBangDienSSI'

    result = scrape_ticker_by_exchange.run(False)

    file_name = f'Báo cáo danh sách mã CK trên bảng điện SSI {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dirname(dirname(dept_folder)),folder_name,file_name),
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

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
