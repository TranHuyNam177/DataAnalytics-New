from automation.risk_management import *
from request.stock import ta
from datawarehouse import DELETE, INSERT

# DONE
def run(  # chạy daily sau batch cuối ngày
    run_time=dt.datetime.now(),
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    t0_date = info['end_date'].replace('/','-')
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))
        
    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        WITH 
        [t] AS ( -- Bảng này nhóm các bảng theo sub_account thành account_code
            SELECT 
                [sub_account].[account_code],
                [vmr9004].[ticker],
                SUM(ISNULL([vmr9004].[total_volume],0)) OVER (PARTITION BY [sub_account].[account_code], [vmr9004].[ticker]) [volume]
            FROM [vcf0051]
            LEFT JOIN [vmr9004] ON [vmr9004].[sub_account] = [vcf0051].[sub_account] AND [vmr9004].[date] =  [vcf0051].[date]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [vmr9004].[sub_account]
            WHERE [vcf0051].[date] = '{t0_date}'
                AND [vcf0051].[contract_type] LIKE 'MR%'
                AND LEN([vmr9004].[ticker]) <= 3
                AND [vmr9004].[total_volume] > 0
                AND [sub_account].[account_code] LIKE '022C%'
            -- ORDER BY account_code, ticker
        )
        SELECT
            [t].*,
            CASE
                WHEN ISNULL([outstanding].[p_outstanding],0) + ISNULL([outstanding].[i_outstanding],0) - ISNULL([rmr0062].[cash],0) < 0 THEN 0
                ELSE ISNULL([outstanding].[p_outstanding],0) + ISNULL([outstanding].[i_outstanding],0) - ISNULL([rmr0062].[cash],0) 
            END [outstanding_net_cash],
            ISNULL([vpr0109].[margin_max_price],0) [max_price],
            ISNULL([vpr0109].[margin_ratio],0) [m_ratio],
            ISNULL([vpr0109].[margin_max_price],0) * ISNULL([vpr0109].[margin_ratio],0) / 100 [breakeven_price_of_stock]
        FROM [t]
        LEFT JOIN (
            SELECT 
                [margin_outstanding].[account_code],
                SUM(ISNULL([margin_outstanding].[principal_outstanding],0)) [p_outstanding],
                SUM(ISNULL([margin_outstanding].[interest_outstanding],0)) [i_outstanding]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[type] IN (N'Margin',N'Trả chậm',N'Bảo lãnh') AND [margin_outstanding].[date] = '{t0_date}'
            GROUP BY [margin_outstanding].[account_code]
            ) [outstanding] 
            ON [outstanding].[account_code] = [t].[account_code]
        LEFT JOIN [rmr0062] ON [rmr0062].[account_code] = [t].[account_code] AND [rmr0062].[date] = '{t0_date}' AND [rmr0062].[loan_type] = 1
        LEFT JOIN [vpr0109] ON [vpr0109].[ticker_code] = [t].[ticker] AND [vpr0109].[date] = '{t0_date}'AND [vpr0109].[room_code] LIKE 'TC01%'
        ORDER BY [t].[account_code], [t].[ticker]
        """,
        connect_DWH_CoSo,
    )

    # Lấy giá
    priceSeries = pd.Series(index=table['ticker'].unique(),dtype=np.float64)
    startDate = bdate(t0_date,-5)
    endDate = t0_date
    for ticker in priceSeries.index:
        try:
            priceFull = ta.hist(ticker,startDate,endDate)
            closePrice = priceFull.iloc[-1,priceFull.columns.get_loc('close')] * 1000
            priceSeries.loc[ticker] = closePrice
        except KeyError:
            pass # ko có giá do hủy niêm yết
    table['close_price'] = table['ticker'].map(priceSeries)
    table = table.dropna(subset=['close_price'])# loại các mã đã hủy niêm yết (ko cần tính giá HV)

    # Credit Rating
    rating_path = join(realpath(dirname(dirname(dirname(__file__)))),'credit_rating','result')
    # Adjust result_table_gen
    gen_table = pd.read_csv(join(rating_path,'result_table_gen.csv'),index_col=['ticker']).iloc[:,-1]
    bank_table = pd.read_csv(join(rating_path,'result_table_bank.csv'),index_col=['ticker']).iloc[:,-1]
    ins_table = pd.read_csv(join(rating_path,'result_table_ins.csv'),index_col=['ticker']).iloc[:,-1]
    sec_table = pd.read_csv(join(rating_path,'result_table_sec.csv'),index_col=['ticker']).iloc[:,-1]
    scoringSeries = pd.concat([gen_table,bank_table,ins_table,sec_table]).squeeze().fillna(0)
    scoringSeries.drop(['BM_'],inplace=True)

    def ratingMapper(x):
        if x >= 75:
            group = 'A'
        elif x >= 50:
            group = 'B'
        elif x >= 25:
            group = 'C'
        else:
            group = 'D'
        return group

    ratingSeries = scoringSeries.map(ratingMapper)
    table['rating'] = table['ticker'].map(ratingSeries).fillna('D')

    # Value
    table['value'] = table['volume'] * table['close_price']
    # Asset Value By Stock
    table['asset_value_by_stock'] = table['volume'] * table['close_price']
    # Discounted Stock Value
    table['base_discount_rate'] = table['rating'].map({'A':0,'B':0.05,'C':0.1,'D':0.15}) # dựa theo rating
    table['case1_value'] = table['value'] * (1 - table['base_discount_rate'] - 0)    # case 1: lấy theo rating
    table['case2_value'] = table['value'] * (1 - table['base_discount_rate'] - 0.05) # case 2: trừ thêm 5%
    table['case3_value'] = table['value'] * (1 - table['base_discount_rate'] - 0.1)  # case 3: trừ thêm 10%
    table['case4_value'] = table['value'] * (1 - table['base_discount_rate'] - 0.15) # case 4: trừ thêm 15%

    table[['weight','case1_result','case2_result','case3_result','case4_result']] = np.nan
    resultCols = ['case1_result','case2_result','case3_result','case4_result']
    for account_code in table['account_code'].unique():
        acccountMask = table['account_code']==account_code
        subTable = table.loc[acccountMask].copy()
        # Asset value
        asset_value = subTable['value'].sum() # có thể bằng 0 với tk chỉ nắm các mã đã hủy niêm yết. VD 022C001551 chỉ nắm KSA ngày 17/02
        # Outstanding
        outstanding = subTable['outstanding_net_cash'].mean() # contain exact 1 value
        tickers = subTable['ticker']
        for ticker in tickers:
            other_tickers = [t for t in tickers if t != ticker]
            othertickersMask = subTable['ticker'].isin(other_tickers)
            tickerMask = subTable['ticker']==ticker
            volume = subTable.loc[tickerMask,'volume'].squeeze()
            value = subTable.loc[tickerMask,'value'].squeeze()
            # Weight & write to table
            weight = value / asset_value
            table.loc[acccountMask & (table['ticker']==ticker),'weight'] =  weight
            if weight == 1:
                # 4 Case như nhau:
                subTable.loc[tickerMask,resultCols] = outstanding / volume
            else:
                # Case 1:
                othertickersValue = subTable.loc[othertickersMask,'case1_value'].sum()
                subTable.loc[tickerMask,'case1_result'] = (outstanding - othertickersValue) / volume
                # Case 2:
                othertickersValue = subTable.loc[othertickersMask,'case2_value'].sum()
                subTable.loc[tickerMask,'case2_result'] = (outstanding - othertickersValue) / volume
                # Case 3:
                othertickersValue = subTable.loc[othertickersMask,'case3_value'].sum()
                subTable.loc[tickerMask,'case3_result'] = (outstanding - othertickersValue) / volume
                # Case 4:
                othertickersValue = subTable.loc[othertickersMask,'case4_value'].sum()
                subTable.loc[tickerMask,'case4_result'] = (outstanding - othertickersValue) / volume

        # Write asset value
        table.loc[acccountMask,'asset_value'] = asset_value
        # Write result
        table.loc[acccountMask,resultCols] = subTable[resultCols]

    result = table[[
        'account_code',
        'ticker',
        'volume',
        'outstanding_net_cash',
        'asset_value_by_stock',
        'asset_value',
        'weight',
        'close_price',
        'max_price',
        'm_ratio',
        'breakeven_price_of_stock',
        'case1_result',
        'case2_result',
        'case3_result',
        'case4_result',
    ]]

    # Insert to DWH-CoSo
    DELETE(connect_DWH_CoSo,'breakeven_price_portfolio',f"WHERE [date] = '{t0_date}'")
    table_for_sql = result.copy()
    table_for_sql.insert(0,'date',dt.datetime.strptime(t0_date,'%Y.%m.%d'))
    INSERT(connect_DWH_CoSo,'breakeven_price_portfolio',table_for_sql)

    # filter điều kiện
    result = result.loc[(result['outstanding_net_cash'] > 0) & (result['max_price'] > 0)] # điều kiện lọc của RMD
    result = result.reset_index(drop=True)

    # ----------------- Write file Import -----------------

    eod = f'{t0_date[-2:]}.{t0_date[5:7]}.{t0_date[:4]}'
    file_name = f"Breakeven Price on Portfolio {eod}.xlsx"
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ## Write sheet ChenhLech
    info_company_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vcenter',
            'bold':True,
            'font_size':8,
            'font_name':'Times New Roman',
        }
    )
    line_format = workbook.add_format(
        {
            'bottom':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':18,
            'font_name':'Times New Roman',
        }
    )
    period_format = workbook.add_format(
        {
            'italic':True,
            'align':'right',
            'valign':'top',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    description1_format = workbook.add_format(
        {
            'italic':True,
            'valign':'bottom',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True,
        }
    )
    description2_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FFC000'
        }
    )
    description3_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#00B050'
        }
    )
    info_header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
            'bg_color':'#FFC000'
        }
    )
    ref_header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
            'bg_color':'#FDE9D9'
        }
    )
    result_header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
            'bg_color':'#00B050'
        }
    )
    num_cell_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
        }
    )
    text_cell_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    pct_cell_format = workbook.add_format(
        {
            'italic':True,
            'border':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0.00%'
        }
    )
    red_num_cell_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'bg_color':'#F68E8E',
        }
    )
    sheet_title_name = 'Breakeven Price on Portfolio'
    eod_sub = dt.datetime.strptime(t0_date,"%Y.%m.%d").strftime("%d/%m/%Y")

    description1 = 'The asset value per stock will be discounted based on the rating as follows:\n' \
                   'Base case: + Group A: 0%, + Group B: 5%, + Group C: 10%, + Group D: 15%\n' \
                   'Plus 5%: + Group A: 5%, + Group B: 10%, + Group C: 15%, + Group D: 20%\n'\
                   'Plus 10%: + Group A: 10%, + Group B: 15%, + Group C: 20%, + Group D: 25%\n'\
                   'Plus 15%: + Group A: 15%, + Group B: 20%, + Group C: 25%, + Group D: 30%'

    description2 = 'Breakeven Price'
    description3 = 'Calculated based on P.Outs after deducting the value of other stocks'

    sub_title_name = f'Date: {eod_sub}' + ' '*36
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.65,'y_scale':0.71})

    worksheet.set_column('A:A',7)
    worksheet.set_column('B:B',13)
    worksheet.set_column('C:C',9)
    worksheet.set_column('D:F',18)
    worksheet.set_column('G:G',0) # ẩn cột asset value
    worksheet.set_column('H:H',9)
    worksheet.set_column('I:I',14)
    worksheet.set_column('I:K',9)
    worksheet.set_column('L:L',10)
    worksheet.set_column('M:P',16)

    # description
    worksheet.merge_range('M5:P9',description1,description1_format)
    worksheet.merge_range('M10:P10',description2,description2_format)
    worksheet.merge_range('M11:P11',description3,description3_format)

    # merge row
    worksheet.merge_range('C1:P1',CompanyName,info_company_format)
    worksheet.merge_range('C2:P2',CompanyAddress,info_company_format)
    worksheet.merge_range('C3:P3',CompanyPhoneNumber,info_company_format)
    worksheet.merge_range('A4:P4','',line_format)
    worksheet.merge_range('A9:I9',sheet_title_name,title_format)
    worksheet.merge_range('A10:I10',sub_title_name,period_format)

    # Freeze panes
    worksheet.freeze_panes(12,3)

    infoCols = [
        'No.',
        'Account',
        'Stock',
        f'Volume in portfolio\n(the end of {eod_sub})',
        f'P.Outs net cash\n(the end of {eod_sub})',
        f'Asset value by stock\n(the end of {eod_sub})',
        'Total asset value',
        'Weighted',
    ]
    refCols = [
        'Market Price',
        'Max price',
        'Ratio',
        'Breakeven price of stock',
    ]
    resultCols = [
        'Base case',
        'Plus 5%',
        'Plus 10%',
        'Plus 15%',
    ]
    worksheet.write_row('A12',infoCols,info_header_format)
    worksheet.write_row('I12',refCols,ref_header_format)
    worksheet.write_row('M12',resultCols,result_header_format)
    worksheet.write_column('A13',np.arange(result.shape[0])+1,num_cell_format)
    worksheet.write_column('B13',result['account_code'],text_cell_format)
    worksheet.write_column('C13',result['ticker'],text_cell_format)
    worksheet.write_column('D13',result['volume'],num_cell_format)
    worksheet.write_column('E13',result['outstanding_net_cash'],num_cell_format)
    worksheet.write_column('F13',result['asset_value_by_stock'],num_cell_format)
    worksheet.write_column('G13',result['asset_value'],num_cell_format)
    worksheet.write_column('H13',result['weight'],pct_cell_format)
    worksheet.write_column('I13',result['close_price'],num_cell_format)
    worksheet.write_column('J13',result['max_price'],num_cell_format)
    worksheet.write_column('K13',result['m_ratio'],num_cell_format)
    worksheet.write_column('L13',result['breakeven_price_of_stock'],num_cell_format)
    for row in range(result.shape[0]):
        for df_col, excel_col in zip(['case1_result','case2_result','case3_result','case4_result'],'MNOP'):
            value = result.loc[row,df_col]
            if value > 0:
                fmt = red_num_cell_format
            else:
                fmt = num_cell_format
            worksheet.write(f'{excel_col}{row+13}',value,fmt)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
