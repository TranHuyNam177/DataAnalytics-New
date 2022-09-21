import numpy as np
import pandas as pd
import os
from os.path import join
import time
import datetime as dt
from automation.trading_service import get_info
from automation.trading_service.thanhtoanbutru import dept_folder
from datawarehouse.DWH_CoSo import connect_DWH_CoSo
from datawarehouse import BDATE
from function import iterable_to_sqlstring


# DONE
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('monthly',run_time)
    # lay ngay 26 thang nay den 25 thang sau
    start_date = (BDATE(info['start_date'],-1)[:-2]+'26').replace('/','-')
    end_date = (info['end_date'][:-2]+'25').replace('/','-')
    period = info['period']
    folder_name = info['folder_name']

    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query và xử lý dataframe ---------------------
    ros_group = [
        '022C015559',
        '022C015959',
        '022C017264',
        '022C017940',
        '022C017289',
        '022C017482',
        '022C017535',
        '022C018358',
        '022C019357',
        '022C696969',
        '022C040921',
        '022C040945',
        '022C041906',
        '022C042028',
        '022C016567',
        '022C016608',
        '022C043586',
        '022C043638',
        '022C043773',
        '022C015598',
        '022C045535',
        '022C043763',
    ]
    table = pd.read_sql(
        f"""
        WITH 
        [ros_group] AS (
        SELECT 
            [sub_account].[sub_account], 
            [sub_account].[account_code], 
            [account].[customer_name]
        FROM [sub_account]
        LEFT JOIN [account]
        ON [account].[account_code] = [sub_account].[account_code]
        WHERE [sub_account].[account_code] IN {iterable_to_sqlstring(ros_group)}
        ),
        [uttb] AS (
        SELECT 
            [payment_in_advance].[sub_account], 
            [payment_in_advance].[fee_at_phs] 
        FROM [payment_in_advance]
        WHERE [payment_in_advance].[date] BETWEEN '{start_date}' AND '{end_date}'
        )
        SELECT 
            [ros_group].[account_code],
            [ros_group].[customer_name],
            SUM([uttb].[fee_at_phs]) AS [fee_phs]
        FROM 
            [ros_group]
        LEFT JOIN
            [uttb]
        ON 
            [ros_group].[sub_account] = [uttb].[sub_account]
        GROUP BY 
            [ros_group].[account_code],
            [ros_group].[customer_name]
        ORDER BY [account_code]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    ).fillna(0)

    table['customer_name'] = table['customer_name'].str.title()

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO DOANH THU UTTB NHÓM ROS TẠO RA
    som = dt.datetime.strptime(start_date,"%Y-%m-%d").strftime("%d-%m")
    eom = dt.datetime.strptime(end_date,"%Y-%m-%d").strftime("%d-%m-%Y")
    file_name = f'Doanh thu UTTB nhóm ROS tạo ra từ {som} đến {eom}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # ------------- Viết sheet -------------
    # Format
    sheet_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':18,
            'font_name':'Times New Roman',
        }
    )
    STK_ten_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    STT_header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    fee_phs_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'bg_color':'#FFF2CB'
        }
    )
    stt_col_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    total_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    total_money_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        }
    )
    sod = dt.datetime.strptime(start_date,"%Y-%m-%d").strftime("%d/%m/%Y")
    eod = dt.datetime.strptime(end_date,"%Y-%m-%d").strftime("%d/%m/%Y")
    sheet_title_name = 'DOANH THU UTTB NHÓM ROS TẠO RA'
    from_to_day = f'Từ {sod} đến {eod}'

    worksheet = workbook.add_worksheet('worksheet')
    worksheet.hide_gridlines(option=2)

    sum_start_row = table.shape[0]+6

    # Set Column Width and Row Height
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',15)
    worksheet.set_column('C:C',67)
    worksheet.set_column('D:D',24)
    worksheet.merge_range('A2:D2',sheet_title_name,sheet_title_format)
    worksheet.merge_range('A3:D3',from_to_day,sheet_title_format)

    worksheet.write('A5','STT',STT_header_format)
    worksheet.write('B5','SỐ TKCK',STK_ten_format)
    worksheet.write('C5','TÊN',STK_ten_format)
    worksheet.write('D5','DOANH THU UTTB',fee_phs_format)
    worksheet.write_column('A6',np.arange(table.shape[0])+1,stt_col_format)
    worksheet.write_column('B6',table.index,text_center_format)
    worksheet.write_column('C6',table['customer_name'],text_left_format)
    worksheet.write_column('D6',table['fee_phs'],money_format)
    worksheet.merge_range(f'A{sum_start_row}:C{sum_start_row}','Total',total_format)
    worksheet.write(f'D{sum_start_row}',f'=SUBTOTAL(9,D6:D{sum_start_row-1})',total_money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
