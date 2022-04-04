from automation.trading_service.giaodichluuky import *


def run(
    run_time=None,
):
    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    trading_record = pd.read_sql(
        f"""
        SELECT date, sub_account, exchange, type_of_asset, volume, value
        FROM trading_record
        WHERE date BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col=['date','sub_account']
    )
    branch_id = pd.read_sql(
        f"""
        SELECT date, sub_account, branch_id
        FROM relationship
        WHERE date BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col=['date','sub_account']
    )
    trading_record['branch_id'] = branch_id
    result = trading_record.groupby(['branch_id','exchange','type_of_asset']).sum()
    result.loc[pd.IndexSlice[:,'HNX',:],'service_fee'] = result['value']*0.027/100
    result.loc[pd.IndexSlice[:,'UPCOM',:],'service_fee'] = result['value']*0.018/100
    check_cp = 'Cổ phiếu thường' in result.index.get_level_values(2)
    check_cw = 'Chứng quyền' in result.index.get_level_values(2)
    check_ccq = 'Chứng chỉ quỹ' in result.index.get_level_values(2)
    check_tpcp = 'Trái phiếu chính phủ' in result.index.get_level_values(2)
    check_tpdn = 'Trái phiếu doanh nghiệp' in result.index.get_level_values(2)
    if check_cp:
        try:
            result.loc[pd.IndexSlice[:,'HOSE','Cổ phiếu thường'],'service_fee'] = result['value']*0.027/100
        except (Exception,):
            pass
    if check_cw:
        try:
            result.loc[pd.IndexSlice[:,'HOSE','Chứng quyền'],'service_fee'] = result['value']*0.018/100
        except (Exception,):
            pass
    if check_ccq:
        try:
            result.loc[pd.IndexSlice[:,'HOSE','Chứng chỉ quỹ'],'service_fee'] = result['value']*0.018/100
        except (Exception,):
            pass
    if check_tpcp:
        try:
            result.loc[pd.IndexSlice[:,'HOSE','Trái phiếu chính phủ'],'service_fee'] = result['value']*0.0054/100
        except (Exception,):
            pass
    if check_tpdn:
        try:
            result.loc[pd.IndexSlice[:,'HOSE','Trái phiếu doanh nghiệp'],'service_fee'] = result['value']*0.0054/100
        except (Exception,):
            pass

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
    result.reset_index(inplace=True)
    result.insert(1,'branch_name',result['branch_id'].map(branch_name_mapper))
    result_hose = result.loc[result['exchange'].isin(['HOSE'])].reset_index()
    result_hose['stt'] = result_hose.index+1
    result_hose = result_hose[['stt','exchange','branch_id','branch_name','volume','value','service_fee']]

    table_title_hose = f'BẢNG KÊ GIÁ DỊCH VỤ GIAO DỊCH {period} - HOSE'
    sum_value_hose = result_hose[['volume','value','service_fee']].sum(axis=0)
    result_hose.columns = ['STT','Sàn','Mã Chi Nhánh','Tên Chi Nhánh','Tổng Số Lượng','Thành Tiền','Phí Phải Nộp']

    # write to Bảng kê giá dịch vụ HOSE
    file_name_hose = f'Báo cáo giá dịch vụ giao dịch {period} - HOSE.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name_hose),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:B',10)
    worksheet.set_column('C:C',14)
    worksheet.set_column('D:D',16)
    worksheet.set_column('E:G',20)

    title_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
        }
    )
    header_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    stt_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    san_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    machinhanh_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    tenchinhanh_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    value_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    tong_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    tong_giatri_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1,
        }
    )
    worksheet.merge_range('A1:G1',table_title_hose,title_format_hose)
    worksheet.write_row(2,0,result_hose.columns,header_format_hose)
    worksheet.write_column(3,0,result_hose['STT'],stt_format_hose)
    worksheet.write_column(3,1,result_hose['Sàn'],san_format_hose)
    worksheet.write_column(3,2,result_hose['Mã Chi Nhánh'],machinhanh_format_hose)
    worksheet.write_column(3,3,result_hose['Tên Chi Nhánh'],tenchinhanh_format_hose)
    for col,name in enumerate(['Tổng Số Lượng','Thành Tiền','Phí Phải Nộp']):
        worksheet.write_column(3,col+4,result_hose[name],value_format_hose)
    tong_row = result_hose.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,3,'Tổng',tong_format_hose)
    worksheet.write_row(tong_row,4,sum_value_hose,tong_giatri_format_hose)
    writer.close()

    result_hnx_upcom = result.loc[result['exchange'].isin(['HNX','UPCOM'])].reset_index()
    result_hnx_upcom['stt'] = result_hnx_upcom.index+1
    result_hnx_upcom = result_hnx_upcom[
        ['stt','branch_id','branch_name','exchange','volume','value','service_fee']
    ]
    result_hnx_upcom['exchange'].replace({'UPCOM':'UP'},inplace=True)
    result_hnx_upcom.insert(4,'',result_hnx_upcom['branch_id']+result_hnx_upcom['exchange'])
    result_hnx_upcom.sort_values(['exchange','branch_id'],inplace=True)
    table_title_hnx_upcom = f'BẢNG KÊ GIÁ DỊCH VỤ GIAO DỊCH {period} - HNX & UPCOM'
    sum_value_hnx_upcom = result_hnx_upcom[['volume','value','service_fee']].sum(axis=0)
    result_hnx_upcom.columns = [
        'STT','Mã Chi Nhánh','Tên Chi Nhánh','Sàn','','Tổng Số Lượng','Thành Tiền','Phí Phải Nộp'
    ]
    # write to Bảng kê giá dịch vụ HNX & UPCOM
    file_name_hnx_upcom = f'Báo cáo giá dịch vụ giao dịch {period} - HNX & UPCOM.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name_hnx_upcom),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:B',8)
    worksheet.set_column('C:C',12)
    worksheet.set_column('D:D',11)
    worksheet.set_column('E:E',15)
    worksheet.set_column('F:H',20)

    title_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
        }
    )
    header_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    stt_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    san_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    machinhanh_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    tenchinhanh_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    value_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    tong_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    tong_giatri_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1,
        }
    )
    worksheet.merge_range('A1:H1',table_title_hnx_upcom,title_format_hnx_upcom)
    worksheet.write('A3','STT',header_format_hnx_upcom)
    worksheet.merge_range('B3:C3','Chi Nhánh',header_format_hnx_upcom)
    worksheet.write_row('D3',result_hnx_upcom.columns[3:],header_format_hnx_upcom)
    worksheet.write_column('A4',result_hnx_upcom['STT'],stt_format_hnx_upcom)
    worksheet.write_column('B4',result_hnx_upcom['Mã Chi Nhánh'],machinhanh_format_hnx_upcom)
    worksheet.write_column('C4',result_hnx_upcom['Tên Chi Nhánh'],tenchinhanh_format_hnx_upcom)
    worksheet.write_column('D4',result_hnx_upcom['Sàn'],san_format_hnx_upcom)
    worksheet.write_column('E4',result_hnx_upcom[''],san_format_hnx_upcom)
    for col,name in enumerate(['Tổng Số Lượng','Thành Tiền','Phí Phải Nộp']):
        worksheet.write_column(3,col+5,result_hnx_upcom[name],value_format_hnx_upcom)
    tong_row = result_hnx_upcom.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,4,'Tổng',tong_format_hnx_upcom)
    worksheet.write_row(tong_row,5,sum_value_hnx_upcom,tong_giatri_format_hnx_upcom)
    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
