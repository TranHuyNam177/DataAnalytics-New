import numpy as np
import pandas as pd
import os
from os import listdir
from os.path import join
import time
import unidecode
from datawarehouse import SYNC
from automation.trading_service import get_info
from automation.trading_service.thanhtoanbutru import dept_folder
from datawarehouse.DWH_CoSo import connect_DWH_CoSo
from datawarehouse import BDATE

# DONE
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    t0_date = info['end_date'].replace('/','-')
    t1_date = BDATE(t0_date,-1)
    next_bdate = BDATE(t1_date,1)
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    folder_path = r"C:\Users\hiepdang\Shared Folder\Trading Service\Report\ThanhToanBuTru\FileFromBanks"
    t1_date_str = t1_date.replace('-','.')
    eib_folder = join(folder_path,t1_date_str,'EIB')
    ocb_folder = join(folder_path,t1_date_str,'OCB')
    eib_filename = listdir(eib_folder)[0]  # should have 1 file
    ocb_filename = listdir(ocb_folder)[0]  # should have 1 file

    # đọc file EIB
    eib_sent_data = pd.read_excel(
        join(eib_folder,eib_filename),
        skiprows=5,
        usecols=['FORACID','ACCOUNT NAME','ACCOUNT BALANCE'],
        dtype={'FORACID':object},
    )
    eib_sent_data.rename(
        {
            'FORACID':'bank_account',
            'ACCOUNT NAME':'account_name',
            'ACCOUNT BALANCE':'bank_balance',
        },
        axis=1,
        inplace=True
    )
    eib_sent_data['bank'] = 'EIB'
    eib_sent_data.set_index(['bank','bank_account'],inplace=True)

    # đọc file OCB
    ocb_sent_data = pd.read_excel(
        join(ocb_folder,ocb_filename),
        usecols=['TÀI KHOẢN','TÊN KHÁCH HÀNG','SỐ DƯ'],
        dtype={'TÀI KHOẢN':object},
    )
    ocb_sent_data.rename(
        {
            'TÀI KHOẢN':'bank_account',
            'TÊN KHÁCH HÀNG':'account_name',
            'SỐ DƯ':'bank_balance',
        },
        axis=1,
        inplace=True
    )
    ocb_sent_data['bank'] = 'OCB'
    ocb_sent_data.set_index(['bank','bank_account'],inplace=True)

    # concat
    bank_sent_balance = pd.concat([eib_sent_data,ocb_sent_data])

    # imported data
    SYNC(connect_DWH_CoSo,'spimported_bank_balance',t0_date,t0_date)
    imported_balance = pd.read_sql(
        f"""
        SELECT 
            [imported_bank_balance].[date],
            [imported_bank_balance].[effective_date],
            [imported_bank_balance].[bank_code] [bank],
            [imported_bank_balance].[bank_account],
            [imported_bank_balance].[account_name],
            [imported_bank_balance].[balance] [flex_balance]
        FROM 
            [imported_bank_balance] 
        WHERE 
            [imported_bank_balance].[effective_date] = '{next_bdate}'
        AND 
            [imported_bank_balance].[status] = 1
        """,
        connect_DWH_CoSo,
        index_col=['bank','bank_account']
    )

    table = imported_balance.join(bank_sent_balance['bank_balance'],how='outer')

    # try to fill missing names using bank files
    table.loc[table['account_name'].isna(),'account_name'] = bank_sent_balance['account_name']

    def name_converter(x):
        try:
            result = unidecode.unidecode(x)
        except (Exception,):
            result = ''  # in case of missing names -> fill with empty string
        return result

    table['account_name'] = table['account_name'].map(name_converter)

    table[['date','effective_date']] = table[['date','effective_date']].fillna(method='ffill')
    table[['flex_balance','bank_balance']] = table[['flex_balance','bank_balance']].fillna(0)
    table['diff'] = table['flex_balance'] - table['bank_balance']
    table.reset_index(inplace=True)

    # ----------------- Write file Import -----------------

    eod = f'{t0_date[-2:]}.{t0_date[5:7]}.{t0_date[:4]}'
    file_name = f"Đối chiếu file import TKNHLK buổi sáng {eod}.xlsx"
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ## Write sheet ChenhLech
    header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'num_format':'dd/mm/yyyy',
            'text_wrap':True,
        }
    )
    text_cell_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
        }
    )
    num_cell_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'text_wrap':True,
        }
    )
    diff_cell_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'font_color':'red',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'text_wrap':True,
        }
    )
    worksheet = workbook.add_worksheet('Chênh Lệch')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:B',14)
    worksheet.set_column('C:C',13)
    worksheet.set_column('D:D',21)
    worksheet.set_column('E:E',30)
    worksheet.set_column('F:H',21)
    worksheet.set_row(0,30)
    worksheet.write_row(
        'A1',
        [
            'TXDATE',
            'EFFECTIVE DATE',
            'NGÂN HÀNG',
            'TÀI KHOẢN',
            'TÊN KHÁCH HÀNG',
            'SỐ DƯ ĐÃ IMPORT',
            'SỐ DƯ NGÂN HÀNG GỬI NGÀY HÔM TRƯỚC',
            'LỆCH',
        ],
        header_format
    )
    worksheet.write_column('A2',table['date'],date_format)
    worksheet.write_column('B2',table['effective_date'],date_format)
    worksheet.write_column('C2',table['bank'],text_cell_format)
    worksheet.write_column('D2',table['bank_account'],text_cell_format)
    worksheet.write_column('E2',table['account_name'],text_cell_format)
    worksheet.write_column('F2',table['flex_balance'],num_cell_format)
    worksheet.write_column('G2',table['bank_balance'],num_cell_format)
    worksheet.write_column('H2',table['diff'],diff_cell_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
