from automation.trading_service.giaodichluuky import *


def run(
    run_time=None,
):
    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    begin_of_year = f'{start_date[:4]}/01/01'
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    ytd_trading_record = pd.read_sql(
        f"""
        SELECT 
        [trading_record].[sub_account],
        [trading_record].[exchange],
        [account].[account_type],
        [trading_record].[type_of_asset],
        [trading_record].[type_of_order],
        [trading_record].[value]
        FROM [trading_record]
        LEFT JOIN [relationship]
	    ON [relationship].[sub_account] = [trading_record].[sub_account]
        LEFT JOIN [account]
	    ON [account].[account_code] = [relationship].[account_code]
	    where
	    [relationship].[date] = '{end_date}'
        AND
        [trading_record].[date] BETWEEN '{begin_of_year}' AND '{end_date}'
        AND [trading_record].[settlement_period] IN (1,2);
        """,
        connect_DWH_CoSo,
        index_col='sub_account',
    )
    period_trading_record = pd.read_sql(
        f"""
        SELECT 
        [trading_record].[sub_account],
        [trading_record].[exchange],
        [account].[account_type],
        [trading_record].[type_of_asset],
        [trading_record].[type_of_order],
        [trading_record].[value]
        FROM [trading_record]
        LEFT JOIN [relationship]
        ON [relationship].[sub_account] = [trading_record].[sub_account]
        LEFT JOIN [account]
        ON [account].[account_code] = [relationship].[account_code]
        where
        [relationship].[date] = '{end_date}'
        AND
        [trading_record].[date] BETWEEN '{start_date}' AND '{end_date}'
        AND [trading_record].[settlement_period] IN (1,2);
        """,
        connect_DWH_CoSo,
        index_col='sub_account',
    )

    def f(df,ttype):
        df['exchange'].replace('UPCOM','HNX',inplace=True)
        df['exchange'].replace('HOSE','HSX',inplace=True)
        domestic = df['account_type'].str.endswith('trong nước')
        foreign = df['account_type'].str.endswith('nước ngoài')
        tudoanh = df['account_type']=='Tự doanh'
        df.loc[domestic,'account_type'] = 'domestic'
        df.loc[foreign,'account_type'] = 'foreign'
        df.loc[tudoanh,'account_type'] = 'dealing'
        df['type_of_asset'] = df['type_of_asset'].map(
            {
                'Cổ phiếu thường':'stock',
                'Chứng chỉ quỹ':'fund_certificate',
                'Chứng quyền':'cw',
                'Trái phiếu doanh nghiệp':'bond',
                'Trái phiếu chính phủ':'bond',
            }
        )
        result = df.groupby(['type_of_order','exchange','type_of_asset','account_type']).sum()
        if ttype=='ytd':
            result.columns = pd.Index(['value_ytd'],name='type_of_period')
        elif ttype=='period':
            result.columns = pd.Index(['value_period'],name='type_of_period')
        else:
            raise TypeError("ttype only accepts 'ytd' or 'period'")
        return result

    ytd_table = f(ytd_trading_record,'ytd')
    period_table = f(period_trading_record,'period')
    result = pd.concat([period_table,ytd_table],axis=1)
    result = result.unstack(['type_of_order','exchange'])
    result = result.reorder_levels(['type_of_order','type_of_period','exchange'],axis=1)
    result.fillna(0,inplace=True)
    needed_idx = pd.MultiIndex.from_tuples([
        ('stock','domestic'),
        ('stock','foreign'),
        ('fund_certificate','domestic'),
        ('fund_certificate','foreign'),
        ('stock','dealing'),
        ('bond','dealing'),
        ('fund_certificate','dealing'),
    ],names=['type_of_asset','account_type'])
    result = pd.DataFrame(data=0,columns=result.columns,index=needed_idx).add(result,
                                                                              fill_value=0)  # ensure complete index
    result[('T','value_period','HSX')] = result[('B','value_period','HSX')]+result[('S','value_period','HSX')]
    result[('T','value_period','HNX')] = result[('B','value_period','HNX')]+result[('S','value_period','HNX')]
    result[('T','value_ytd','HSX')] = result[('B','value_ytd','HSX')]+result[('S','value_ytd','HSX')]
    result[('T','value_ytd','HNX')] = result[('B','value_ytd','HNX')]+result[('S','value_ytd','HNX')]
    cols = [
        ('B','value_period','HNX'),
        ('B','value_period','HSX'),
        ('B','value_ytd','HNX'),
        ('B','value_ytd','HSX'),
        ('S','value_period','HNX'),
        ('S','value_period','HSX'),
        ('S','value_ytd','HNX'),
        ('S','value_ytd','HSX'),
        ('T','value_period','HNX'),
        ('T','value_period','HSX'),
        ('T','value_ytd','HNX'),
        ('T','value_ytd','HSX'),
    ]
    result = result[cols]  # ensure column order matched with excel output

    # Write to Báo cáo phí chuyển khoản
    file_name = f'Báo cáo tình hình HĐKD (Biểu II.6) {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',16)
    worksheet.set_column('C:N',12)
    worksheet.set_row(7,47)
    worksheet.set_row(8,47)
    worksheet.set_row(9,47)
    worksheet.set_row(10,47)
    title_format = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':13,
            'align':'center',
        }
    )
    dvt_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'italic':True,
            'align':'right',
        }
    )
    supheader_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'border':1,
        }
    )
    header_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'valign':'vcenter',
            'border':1,
        }
    )
    stt_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'valign':'vcenter',
            'border':1,
        }
    )
    index_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
            'valign':'top',
            'border':1
        }
    )
    supindex_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'text_wrap':True,
            'valign':'top',
            'border':1
        }
    )
    value_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'valign':'vcenter',
            'border':1
        }
    )
    worksheet.merge_range('A1:N1',f'Báo Cáo Tình Hình HĐKD - Biểu II.6 - {period}',title_format)
    worksheet.merge_range('A3:A5','STT',header_format)
    worksheet.merge_range('B3:B5','Loại chứng khoán',header_format)
    worksheet.merge_range('C3:F3','Tổng mua',supheader_format)
    worksheet.merge_range('G3:J3','Tổng bán',supheader_format)
    worksheet.merge_range('K3:N3','Tổng mua + bán',supheader_format)
    worksheet.merge_range('C4:D4','Trong kỳ',header_format)
    worksheet.merge_range('E4:F4','Lũy kế từ đầu năm',header_format)
    worksheet.merge_range('G4:H4','Trong kỳ',header_format)
    worksheet.merge_range('I4:J4','Lũy kế từ đầu năm',header_format)
    worksheet.merge_range('K4:L4','Trong kỳ',header_format)
    worksheet.merge_range('M4:N4','Lũy kế từ đầu năm',header_format)
    worksheet.write_row('C5',['HNX','HSX']*6,header_format)
    worksheet.write_row('A6',np.arange(14)+1,header_format)
    worksheet.write_column('A7',np.arange(9)+1,stt_format)
    idx_list = [
        'A. Nhà đầu tư',
        '1. Giao dịch cổ phiếu của nhà đầu tư trong nước',
        '2. Giao dịch cổ phiếu của nhà đầu tư nước ngoài',
        '3. Giao dịch chứng chỉ quỹ của nhà đầu tư trong nước',
        '4. Giao dịch chứng chỉ quỹ của nhà đầu tư nước ngoài',
        'B. Tự doanh',
        '1. Cổ phiếu',
        '2. Trái phiếu',
        '3. Chứng chỉ quỹ',
    ]
    for row,idx in enumerate(idx_list):
        if idx in ['A. Nhà đầu tư','B. Tự doanh']:
            fmt = supindex_format
        else:
            fmt = index_format
        worksheet.write(row+6,1,idx,fmt)
    worksheet.write_row('C7',[None]*12,value_format)
    worksheet.write_row('C12',[None]*12,value_format)
    dvt = 'mil'
    if dvt=='bil':
        div = 1e9;
        dvt_text = 'Đơn vị tính: tỷ đồng'
    elif dvt=='mil':
        div = 1e6;
        dvt_text = 'Đơn vị tính: triệu đồng'
    elif dvt=='k':
        div = 1e3;
        dvt_text = 'Đơn vị tính: nghìn đồng'
    elif dvt=='unit':
        div = 1;
        dvt_text = 'Đơn vị tính: đồng'
    else:
        raise ValueError("dvt must be either 'bil', 'mil', 'k', 'unit'")
    worksheet.merge_range('L2:N2',dvt_text,dvt_format)
    worksheet.write_row('C8',result.loc[('stock','domestic')]/div,value_format)
    worksheet.write_row('C9',result.loc[('stock','foreign')]/div,value_format)
    worksheet.write_row('C10',result.loc[('fund_certificate','domestic')]/div,value_format)
    worksheet.write_row('C11',result.loc[('fund_certificate','foreign')]/div,value_format)
    worksheet.write_row('C13',result.loc[('stock','dealing')]/div,value_format)
    worksheet.write_row('C14',result.loc[('bond','dealing')]/div,value_format)
    worksheet.write_row('C15',result.loc[('fund_certificate','dealing')]/div,value_format)
    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
