from automation.trading_service.thanhtoanbutru import *
from datawarehouse import EXEC, BDATE

# DONE
def generateTempData(
    run_time = dt.datetime.now()
):
    """
    Hàm này được chạy lưu file template sau batch cuối ngày
    """

    info = get_info('daily',run_time)
    end_date = info['end_date']

    table = pd.read_sql(
        f"""
        SELECT
            [sub_account_deposit].[sub_account],
            [sub_account_deposit].[opening_balance],
            [sub_account_deposit].[closing_balance]
        FROM [sub_account_deposit]
        WHERE [sub_account_deposit].[date] ='{end_date}'
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    table.to_pickle(join(dirname(__file__),'sodutientaikhoankhachhang',f'{end_date.replace("/","")}.pickle'))


def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    folder_name = info['folder_name']
    t0_date = info['start_date'].replace('/','-')
    t1_date = BDATE(t0_date,-1)
    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    EXEC(connect_DWH_CoSo,'spsub_account_deposit',FrDate=t0_date,ToDate=t0_date)

    # --------------------- Viết Query ---------------------
    info_table = pd.read_sql(
        f"""
        SELECT
            [relationship].[account_code],
            [relationship].[sub_account],
            [account].[customer_name],
            [broker].[broker_name],
            [branch].[branch_name]
        FROM
            [relationship]
        LEFT JOIN [account] ON [account].[account_code] = [relationship].[account_code]
        LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
        LEFT JOIN [broker] ON [broker].[broker_id] = [relationship].[broker_id]
        WHERE [relationship].[date] = '{t1_date}'
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    t1_table = pd.read_sql(
        f"""
        SELECT
            [sub_account_deposit].[sub_account], 
            [sub_account_deposit].[opening_balance] AS [opening_balance_t1],
            [sub_account_deposit].[closing_balance] AS [closing_balance_t1]
        FROM [sub_account_deposit]
        WHERE [sub_account_deposit].[date] = '{t1_date}'
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    t0_table = pd.read_sql(
        f"""
        SELECT
            [sub_account_deposit].[sub_account],
            [sub_account_deposit].[opening_balance]
        FROM [sub_account_deposit]
        WHERE [sub_account_deposit].[date] = '{t0_date}'
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # read file đã lưu sau batch cuối ngày hôm qua
    saved_folder = join(dirname(__file__),'sodutientaikhoankhachhang')
    t1_table_ref = pd.read_pickle(join(saved_folder,f'{t1_date.replace("-","")}.pickle'))
    # rename columns
    t1_table_ref = t1_table_ref.add_suffix('_t1_ref')
    t0_table = t0_table.add_suffix('_t0')
    # join
    table = pd.concat([info_table,t1_table,t0_table,t1_table_ref],join='outer',axis=1)
    table = table.loc[~table.iloc[:,-5:].isna().all(axis=1)]
    table.fillna(0,inplace=True)
    # check
    check_1 = table['opening_balance_t0']==table['closing_balance_t1']
    check_2 = table['opening_balance_t1']==table['opening_balance_t1_ref']
    check_3 = table['closing_balance_t1']==table['closing_balance_t1_ref']
    check_4 = (table[['opening_balance_t0','opening_balance_t1','closing_balance_t1']]>=0).all(axis=1)
    table['check'] = check_1 & check_2 & check_3 & check_4
    table['check'] = table['check'].replace(True,'Khớp')
    table['check'] = table['check'].replace(False,'Bất thường')
    table.sort_values('check',inplace=True)

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU SỐ DƯ TIỀN TÀI KHOẢN KHÁCH HÀNG
    eod = dt.datetime.strptime(t0_date,"%Y-%m-%d").strftime("%d.%m.%Y")
    file_name = f'Báo cáo Đối chiếu số dư tiền TKKH {eod}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    company_name_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    company_info_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    empty_row_format = workbook.add_format(
        {
            'bottom':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    sheet_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':14,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    sub_title_date_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    headers_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(  # for customer name only
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    sum_money_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    footer_dmy_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    footer_text_format = workbook.add_format(
        {
            'bold':True,
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU SỐ DƯ TIỀN TÀI KHOẢN KHÁCH HÀNG'
    eod_sub = dt.datetime.strptime(t0_date,"%Y-%m-%d").strftime("%d/%m/%Y")
    sub_title_name = f'Ngày {eod_sub}'

    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.66,'y_scale':0.71})
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:C',13)
    worksheet.set_column('D:D',28)
    worksheet.set_column('E:I',15)
    worksheet.set_column('J:J',9)
    worksheet.set_column('K:K',20)
    worksheet.set_column('L:L',28)
    worksheet.merge_range('C1:L1',CompanyName,company_name_format)
    worksheet.merge_range('C2:L2',CompanyAddress,company_info_format)
    worksheet.merge_range('C3:L3',CompanyPhoneNumber,company_info_format)
    worksheet.write_row('A4',['']*14,empty_row_format)
    worksheet.merge_range('A7:L7',sheet_title_name,sheet_title_format)
    worksheet.merge_range('A8:L8',sub_title_name,sub_title_date_format)
    worksheet.merge_range('A10:A11','STT',headers_format)
    worksheet.merge_range('B10:B11','Số tài khoản',headers_format)
    worksheet.merge_range('C10:C11','Số tiểu khoản',headers_format)
    worksheet.merge_range('D10:D11','Tên khách hàng',headers_format)
    worksheet.merge_range('E10:G10','Dữ liệu tại Flex',headers_format)
    worksheet.merge_range('H10:I10','Dữ liệu lưu ngoài Flex',headers_format)
    sub_headers = [
        'Số dư tiền đầu ngày T-1',
        'Số dư tiền cuối ngày T-1',
        'Số dư tiền đầu ngày T0',
        'Số dư tiền đầu ngày T-1',
        'Số dư tiền cuối ngày T-1'
    ]
    worksheet.write_row('E11',sub_headers,headers_format)
    worksheet.merge_range('J10:J11','Bất thường',headers_format)
    worksheet.merge_range('K10:K11','Chi nhánh quản lý',headers_format)
    worksheet.merge_range('L10:L11','Nhân viên quản lý tài khoản',headers_format)

    sum_start_row = table.shape[0]+12
    worksheet.merge_range(f'A{sum_start_row}:D{sum_start_row}','Tổng',headers_format)
    worksheet.merge_range(f'J{sum_start_row}:L{sum_start_row}','',text_center_format)
    footer_start_row = sum_start_row+2
    footer_date = bdate(t0_date,1).split('-')
    worksheet.merge_range(
        f'J{footer_start_row}:L{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format,
    )
    worksheet.merge_range(
        f'J{footer_start_row+1}:L{footer_start_row+1}',
        'Người duyệt',
        footer_text_format,
    )
    worksheet.merge_range(
        f'A{footer_start_row+1}:C{footer_start_row+1}',
        'Người lập',
        footer_text_format,
    )
    worksheet.write_column('A12',np.arange(table.shape[0])+1,text_center_format)
    worksheet.write_column('B12',table['account_code'],text_center_format)
    worksheet.write_column('C12',table.index,text_center_format)
    worksheet.write_column('D12',table['customer_name'].str.title(),text_left_format)
    worksheet.write_column('E12',table['opening_balance_t1'],money_format)
    worksheet.write_column('F12',table['closing_balance_t1'],money_format)
    worksheet.write_column('G12',table['opening_balance_t0'],money_format)
    worksheet.write_column('H12',table['opening_balance_t1_ref'],money_format)
    worksheet.write_column('I12',table['closing_balance_t1_ref'],money_format)
    worksheet.write_column('J12',table['check'],text_center_format)
    worksheet.write_column('K12',table['branch_name'],text_left_format)
    worksheet.write_column('L12',table['broker_name'].str.title(),text_left_format)
    for col in 'EFGHI':
        worksheet.write(f'{col}{sum_start_row}',f'=SUBTOTAL(9,{col}12:{col}{sum_start_row-1})',sum_money_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')} ::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
