import numpy as np
import pandas as pd
import os
import argparse
from os.path import join, expanduser
import datetime as dt
from request import connect_DWH_CoSo
from automation.trading_service import convert_int

def run(
        startDate=None,
        endDate=None
):
    dateRun = dt.datetime.now().strftime("%d.%m.%Y")
    start_date = dt.datetime.strptime(startDate, "%d/%m/%Y").strftime("%Y/%m/%d")
    end_date = dt.datetime.strptime(endDate, "%d/%m/%Y").strftime("%Y/%m/%d")

    user_name = expanduser("~")  # tên thư mục tương ứng với user đang logging
    folder_name = os.path.join(user_name, "BaoCaoTuDong")
    # create folder
    if not os.path.isdir(join(folder_name, dateRun)):
        os.makedirs(join(folder_name, dateRun))

    depository_fee = pd.read_sql(
        f"""
        SELECT 
            [date], 
            [sub_account], 
            [fee_amount]
        FROM [depository_fee]
        WHERE [date] BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col=['date','sub_account'],
    )
    branch_id = pd.read_sql(
        f"""
        SELECT [date], [sub_account], [branch_id]
        FROM [relationship]
        WHERE [date] BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col=['date','sub_account'],
    ).squeeze()
    depository_fee['branch_id'] = branch_id
    depository_fee = depository_fee.groupby('branch_id')['fee_amount'].sum()
    branch_name_mapper = {
        '0001':'HQ',
        '0101':'Quận 3',
        '0102':'PMH',
        '0104':'Q7',
        '0105':'TB',
        '0116':'P.QLTK1',
        '0111':'InB1',
        '0113':'IB',
        '0201':'Hà nội',
        '0202':'TX',
        '0301':'Hải phòng',
        '0117':'Quận 1',
        '0118':'P.QLTK3',
        '0119':'InB2',
    }
    result = pd.DataFrame(
        columns=[
            'STT',
            'Tên Chi Nhánh',
            'Phí Lưu Ký'
        ],
        index=branch_name_mapper.keys()
    )
    result['STT'] = np.arange(1,result.shape[0]+1)
    result['Tên Chi Nhánh'] = result.index.map(branch_name_mapper)
    result['Phí Lưu Ký'] = depository_fee
    result['Phí Lưu Ký'].fillna(0,inplace=True)
    result.index.name = 'Mã Chi Nhánh'
    result.reset_index(inplace=True)
    result = result[['STT','Tên Chi Nhánh','Mã Chi Nhánh','Phí Lưu Ký']]

    sum_fee = result['Phí Lưu Ký'].sum()

    startDate = startDate.replace('/', '.')
    endDate = endDate.replace('/', '.')

    startMonth = int(startDate.split('.')[1])
    startYear = int(startDate.split('.')[-1])
    endMonth = int(endDate.split('.')[1])
    endYear = int(endDate.split('.')[-1])
    if startMonth == endMonth and startYear == endYear:
        timePeriod = f'{convert_int(startMonth)}.{startYear}'
    else:
        timePeriod = f'{convert_int(startMonth)}.{startYear} - {convert_int(endMonth)}.{endYear}'

    table_title = f'PHÍ LƯU KÝ {timePeriod}'

    # Write to Excel
    file_name = f'Báo cáo phí tạm tính phí lưu ký từ {startDate} đến {endDate}_.xlsx'
    writer = pd.ExcelWriter(
        join(folder_name, dateRun, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(timePeriod)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',21)
    worksheet.set_column('C:D',18)

    title_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
        }
    )
    header_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    stt_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    tenchinhanh_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    machinhanh_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    philuuky_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    tong_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    tongphiluuky_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1,
        }
    )
    worksheet.merge_range('A1:D1',table_title,title_format)
    for col_num,col_name in enumerate(result.columns):
        worksheet.write(2,col_num,col_name,header_format)
    for row in range(result.shape[0]):
        worksheet.write(row+3,0,result.iloc[row,0],stt_format)
        worksheet.write(row+3,1,result.iloc[row,1],tenchinhanh_format)
        worksheet.write(row+3,2,result.iloc[row,2],machinhanh_format)
        worksheet.write(row+3,3,result.iloc[row,3],philuuky_format)
    tong_row = result.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,2,'Tổng',tong_format)
    worksheet.write(tong_row,3,sum_fee,tongphiluuky_format)
    writer.close()


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument("-s", "--startDate", required=True, help="input start date with format: dd/mm/YYYY")
    ap.add_argument("-e", "--endDate", required=True, help="input end date with format: dd/mm/YYYY")
    args = vars(ap.parse_args())
    run(args['startDate'], args['endDate'])