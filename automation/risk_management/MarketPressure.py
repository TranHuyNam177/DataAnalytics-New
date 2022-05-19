from automation.risk_management import *
from datawarehouse import BDATE


def run(  # chạy hàng ngày
    run_time = dt.datetime.now()
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    dataDate = info['end_date']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    # list các tài khoản nợ xấu cố định và loại bỏ thêm 1 tài khoản tự doanh
    badDebtAccounts = {
        '022C078252',
        '022C012620',
        '022C012621',
        '022C012622',
        '022C089535',
        '022C089950',
        '022C089957',
        '022C050302',
        '022C006827',
        '022P002222',
    }
    detailTable = pd.read_sql(
        f"""
        WITH
        [BranchTable] AS (
            SELECT DISTINCT
                [relationship].[account_code],
                [branch].[branch_name]
            FROM [relationship]
            LEFT JOIN [branch] ON [relationship].[branch_id] = [branch].[branch_id]
            WHERE [relationship].[date] = '{dataDate}'
        ),
        [LoanTable] AS (
            SELECT
                [BranchTable].[branch_name] [Location],
                [margin_outstanding].[account_code] [Custody],
                SUM([principal_outstanding]) [OriginalLoan],
                SUM([interest_outstanding]) [Interest],
                SUM([principal_outstanding])+SUM([interest_outstanding])+SUM([fee_outstanding]) [TotalLoan]
            FROM [margin_outstanding]
            LEFT JOIN [BranchTable] 
                ON [margin_outstanding].[account_code] = [BranchTable].[account_code]
            WHERE [margin_outstanding].[date] = '{dataDate}'
                AND [margin_outstanding].[type] IN (N'Margin', N'Trả chậm', N'Bảo lãnh')
                AND [margin_outstanding].[account_code] NOT IN {iterable_to_sqlstring(badDebtAccounts)}
            GROUP BY [BranchTable].[branch_name], [margin_outstanding].[account_code]
        ),
        [CashMargin] AS (
            SELECT
                [rmr0062].[account_code] [Custody],
                SUM([rmr0062].[cash]) [CashAndPIA],
                SUM([rmr0062].[margin_value]) [MarginValue]
            FROM [rmr0062]
            WHERE [rmr0062].[date] = '{dataDate}' AND [rmr0062].[loan_type] = 1
            GROUP BY [rmr0062].[account_code]
        ),
        [Asset] AS (
            SELECT
                [sub_account].[account_code] [Custody],
                SUM([rmr0015].[market_value]) [TotalAssetValue]
            FROM [rmr0015]
            LEFT JOIN [sub_account] 
                ON [sub_account].[sub_account] = [rmr0015].[sub_account]
            WHERE [rmr0015].[date] = '{dataDate}'
            GROUP BY [account_code]
        ),
        [MidResult] AS (
            SELECT
                [LoanTable].*,
                ISNULL([CashMargin].[CashAndPIA],0) [CashAndPIA],
                ISNULL([CashMargin].[MarginValue],0) [MarginValue],
                ISNULL([Asset].[TotalAssetValue],0) [TotalAssetValue],
                CASE
                    WHEN ISNULL([CashMargin].[CashAndPIA],0) > [LoanTable].[TotalLoan] THEN 100
                    WHEN ISNULL([CashMargin].[MarginValue],0) = 0 THEN 0
                    ELSE (1 - ([LoanTable].[TotalLoan] - [CashMargin].[CashAndPIA]) / [CashMargin].[MarginValue]) * 100 
                END [MMRMA],
                CASE
                    WHEN ISNULL([CashMargin].[CashAndPIA],0) > [LoanTable].[TotalLoan] THEN 100
                    WHEN ISNULL([Asset].[TotalAssetValue],0) = 0 THEN 0
                    ELSE (1 - ([LoanTable].[TotalLoan] - [CashMargin].[CashAndPIA]) / [Asset].[TotalAssetValue]) * 100 
                END [MMRTA],
                '' [Note]
            FROM [LoanTable]
            LEFT JOIN [CashMargin] ON [LoanTable].[Custody] = [CashMargin].[Custody]
            LEFT JOIN [Asset] ON [LoanTable].[Custody] = [Asset].[Custody]
        )
        SELECT 
            ROW_NUMBER() OVER (ORDER BY [MidResult].[MMRTA],[MidResult].[MMRMA]) [No.],
            [MidResult].*,
            CASE
                WHEN [MidResult].[MMRTA] BETWEEN 80 AND 100 THEN '[80-100]'
                WHEN [MidResult].[MMRTA] BETWEEN 75 AND 80 THEN '[75-80]'
                WHEN [MidResult].[MMRTA] BETWEEN 70 AND 75 THEN '[70-75]'
                WHEN [MidResult].[MMRTA] BETWEEN 65 AND 70 THEN '[65-70]'
                WHEN [MidResult].[MMRTA] BETWEEN 60 AND 65 THEN '[60-65]'
                WHEN [MidResult].[MMRTA] BETWEEN 55 AND 60 THEN '[55-60]'
                WHEN [MidResult].[MMRTA] BETWEEN 50 AND 55 THEN '[50-55]'
                WHEN [MidResult].[MMRTA] BETWEEN 45 AND 50 THEN '[45-50]'
                WHEN [MidResult].[MMRTA] BETWEEN 40 AND 45 THEN '[40-45]'
                WHEN [MidResult].[MMRTA] BETWEEN 35 AND 40 THEN '[35-40]'
                WHEN [MidResult].[MMRTA] BETWEEN 30 AND 35 THEN '[30-35]'
                WHEN [MidResult].[MMRTA] BETWEEN 25 AND 30 THEN '[25-30]'
                WHEN [MidResult].[MMRTA] BETWEEN 20 AND 25 THEN '[20-25]'
                WHEN [MidResult].[MMRTA] BETWEEN 15 AND 20 THEN '[15-20]'
                WHEN [MidResult].[MMRTA] BETWEEN 10 AND 15 THEN '[10-15]'
                ELSE '[00-10]'
            END [Group]
        FROM [MidResult]
        ORDER BY [MidResult].[MMRTA],[MidResult].[MMRMA]
        """,
        connect_DWH_CoSo
    )
    summaryTable = detailTable.groupby('Group')['OriginalLoan'].agg(['count','sum'])
    groupsMapper = {
        '[00-10]':'Market Pressure < 10%',
        '[10-15]':'10%<=Market Pressure < 15%',
        '[15-20]':'15%<=Market Pressure < 20%',
        '[20-25]':'20%<=Market Pressure < 25%',
        '[25-30]':'25%<=Market Pressure < 30%',
        '[30-35]':'30%<=Market Pressure < 35%',
        '[35-40]':'35%<=Market Pressure < 40%',
        '[40-45]':'40%<=Market Pressure < 45%',
        '[45-50]':'45%<=Market Pressure < 50%',
        '[50-55]':'50%<=Market Pressure < 55%',
        '[55-60]':'55%<=Market Pressure < 60%',
        '[60-65]':'60%<=Market Pressure < 65%',
        '[65-70]':'65%<=Market Pressure < 70%',
        '[70-75]':'70%<=Market Pressure < 75%',
        '[75-80]':'75%<=Market Pressure < 80%',
        '[80-100]':'Market Pressure >= 80%'
    }
    summaryTable = summaryTable.reindex(groupsMapper.keys()).fillna(0).reset_index()
    summaryTable['Group'] = summaryTable['Group'].map(groupsMapper)
    summaryTable = summaryTable.rename({'count':'AccountNumber','sum':'Outstanding'},axis=1)
    summaryTable['Outstanding'] /= 1000000
    summaryTable['Proportion'] = summaryTable['Outstanding'] / summaryTable['Outstanding'].sum() * 100

    ###################################################
    ###################################################
    ###################################################

    t0_day = dataDate[-2:]
    t0_month = calendar.month_name[int(dataDate[5:7])]
    t0_year = dataDate[0:4]
    file_name = f'RMD_Market Pressure _end of {t0_day}.{t0_month} {t0_year}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Sheet Summary
    cell_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Calibri'
        }
    )
    title1_red_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Calibri',
            'color': '#FF0000'
        }
    )
    title_2_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    title_2_color_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'color': '#FF0000'
        }
    )
    title_3_format = workbook.add_format(
        {
            'bold': True,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    headers_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap':True,
            'border':1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    text_left_merge_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    text_left_color_format = workbook.add_format(
        {
            'border': 1,
            'bold':True,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'color': '#FF0000'
        }
    )
    num_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    sum_num_format = workbook.add_format(
        {
            'bold':True,
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    money_normal_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    money_small_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '0.00'
        }
    )
    percent_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '0.00'
        }
    )
    headers = [
        'Criteria',
        'No of accounts',
        'Outstanding',
        '% Total Oustanding'
    ]
    title_2 = f'Data is as at end {t0_day}.{t0_month} {t0_year} (it is not inculde 08 accounts that belong to Accumulated Negative Value)'
    title_3 = 'C. Market Pressure (%) is used to indicate the breakeven point of loan with assumption that whole portfolio may drop at same percentage.'

    summary_sheet = workbook.add_worksheet('Summary')
    summary_sheet.set_column('A:A',31)
    summary_sheet.set_column('B:B',15)
    summary_sheet.set_column('C:C',14)
    summary_sheet.set_column('D:D',11)
    summary_sheet.set_column('E:E',19)
    summary_sheet.set_column('F:F',21)
    summary_sheet.set_column('G:G',0)

    summary_sheet.set_row(5,30)
    summary_sheet.merge_range('A1:I1','',cell_format)
    summary_sheet.write_rich_string(
        'A1','SUMMARY RISK REPORT FOR ',title1_red_format,'Market Pressure (%)',cell_format
    )
    summary_sheet.merge_range('A2:F2',title_2,title_2_format)
    summary_sheet.merge_range('A3:F3','', cell_format)
    summary_sheet.write_rich_string(
        'A3', 'Unit for Outstanding: ',title_2_color_format,'million dong',title_2_format
    )
    summary_sheet.write('A4',title_3,title_3_format)
    summary_sheet.write_row('A6',headers,headers_format)
    summary_sheet.merge_range('A7:A8','0 < Market Pressure < 10%',text_left_merge_format)
    summary_sheet.write_rich_string('A9','10%<= Market Pressure',text_left_color_format,' <15%',text_left_format)
    summary_sheet.write_rich_string('A10','15%<= Market Pressure',text_left_color_format,' <20%',text_left_format)
    summary_sheet.write_rich_string('A11','20%<= Market Pressure',text_left_color_format,' <25%',text_left_format)
    summary_sheet.write_rich_string('A12','25%<= Market Pressure',text_left_color_format,' <30%',text_left_format)
    summary_sheet.write_rich_string('A13','30%<= Market Pressure',text_left_color_format,' <35%',text_left_format)
    summary_sheet.write_rich_string('A14','35%<= Market Pressure',text_left_color_format,' <40%',text_left_format)
    summary_sheet.write_rich_string('A15','40%<= Market Pressure',text_left_color_format,' <45%',text_left_format)
    summary_sheet.write_rich_string('A16','45%<= Market Pressure',text_left_color_format,' <50%',text_left_format)
    summary_sheet.write_rich_string('A17','50%<= Market Pressure',text_left_color_format,' <55%',text_left_format)
    summary_sheet.write_rich_string('A18','55%<= Market Pressure',text_left_color_format,' <60%',text_left_format)
    summary_sheet.write_rich_string('A19','60%<= Market Pressure',text_left_color_format,' <65%',text_left_format)
    summary_sheet.write_rich_string('A20','65%<= Market Pressure',text_left_color_format,' <70%',text_left_format)
    summary_sheet.write_rich_string('A21','70%<= Market Pressure',text_left_color_format,' <75%',text_left_format)
    summary_sheet.write_rich_string('A22','75%<= Market Pressure',text_left_color_format,' <80%',text_left_format)
    summary_sheet.write_rich_string('A23','Market Pressure',text_left_color_format,' >=80%', text_left_format)

    summary_sheet.merge_range('B7:B8',detailTable.loc[detailTable.index[0],'AccountNumber'],num_right_format)
    summary_sheet.merge_range('D7:D8',detailTable.loc[detailTable.index[0],'Proportion'],percent_format)
    summary_sheet.write_column('B9',detailTable.loc[detailTable.index[1:],'Proportion'],num_right_format)
    for loc, value in enumerate(detailTable['sum']):
        if value > 100 or value == 0:
            fmt = money_normal_format
        else:
            fmt = money_small_format
        if row == 0:
            summary_sheet.merge_range('C7:C8',value,fmt)
        else:
            summary_sheet.write(loc+7,2,value,fmt)

    summary_sheet.write_column('D9',table_sheet_1['%TotalOutstanding'][1:],percent_format)
    sum_row = table_sheet_1.shape[0]+8
    summary_sheet.write(f'A{sum_row}','Total',headers_format)
    summary_sheet.write(f'B{sum_row}',table_sheet_1['count'].sum(),sum_num_format)
    summary_sheet.write(f'C{sum_row}',table_sheet_1['sum'].sum(),sum_num_format)
    summary_sheet.write(f'D{sum_row}','', sum_num_format)

    ###################################################
    ###################################################
    ###################################################

    # Sheet Detail
    # Format
    sum_num_color_format = workbook.add_format(
        {
            'bold': True,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '#,##0',
            'color': '#FF0000'
        }
    )
    sum_num_format = workbook.add_format(
        {
            'bold': True,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '#,##0',
        }
    )
    header_1_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#FFC000'
        }
    )
    header_2_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#FFF2CC'
        }
    )
    header_3_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'color': '#FF0000',
            'bg_color': '#FFF2CC'
        }
    )
    header_4_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    percent_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '0.00'
        }
    )
    headers_1 = [
        'No.',
        'Location',
        'Custody',
        'Original Loan',
        'interest',
        'Total Loan',
    ]
    headers_2 = [
        'Total Cash & PIA (MR0062 có vay xuất cuối ngày làm việc)',
        'Total Margin value (RMR0062)',
        'Total Asset Value (RMR0015 with market price)'
    ]
    headers_3 = [
        'MMR (base on Marginable Asset)',
        'MMR (base on Total Asset)'
    ]

    worksheet = workbook.add_worksheet('Detail')

    worksheet.set_column('A:A',0)
    worksheet.set_column('B:B',5.5)
    worksheet.set_column('C:C',11.5)
    worksheet.set_column('D:D',17)
    worksheet.set_column('E:G',19)
    worksheet.set_column('H:K',0)
    worksheet.set_column('L:L',14)
    worksheet.set_column('M:M',16)
    worksheet.set_column('N:N',0)
    worksheet.set_column('O:O',9)

    worksheet.set_row(1,52)
    worksheet.write('A2','Bad Loans',header_4_format)
    worksheet.write('B1',table_sheet_2.shape[0],sum_num_color_format)
    worksheet.write('E1',table_sheet_2['original_loan'].sum(),sum_num_color_format)
    worksheet.write('F1',table_sheet_2['original_loan'].sum()/pow(10,6),sum_num_format)
    worksheet.write_row('B2',headers_1,header_1_format)
    worksheet.write_row('H2',headers_2,header_2_format)
    worksheet.write_row('K2',headers_3,header_3_format)
    worksheet.write('M2','Group/deal',header_2_format)
    worksheet.write('N2','bad loan',header_4_format)
    worksheet.write('O2','Note',header_4_format)
    worksheet.write_column('B3',np.arange(table_sheet_2.shape[0])+1,text_center_format)
    worksheet.write_column('C3',table_sheet_2['location'],text_center_format)
    worksheet.write_column('D3',table_sheet_2['custody'],text_left_format)
    worksheet.write_column('E3',table_sheet_2['original_loan'],money_format)
    worksheet.write_column('F3',table_sheet_2['interest'],money_format)
    worksheet.write_column('G3',table_sheet_2['total_loan'],money_format)
    worksheet.write_column('H3',table_sheet_2['total_cash'],money_format)
    worksheet.write_column('I3',table_sheet_2['total_margin_val'],money_format)
    worksheet.write_column('J3',table_sheet_2['total_asset_val'],money_format)
    worksheet.write_column('K3',table_sheet_2['MMR_MarginableAsset'],percent_format)
    worksheet.write_column('L3',table_sheet_2['MMR_TotalAsset'],percent_format)
    worksheet.write_column('M3',[0]*table_sheet_2.shape[0],money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

