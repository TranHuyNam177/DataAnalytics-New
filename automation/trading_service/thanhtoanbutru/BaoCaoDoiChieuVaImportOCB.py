from automation.trading_service.thanhtoanbutru import *

"""
(Change Request ngày 27/12/2021)
Sửa lại Rule : Nếu EIB có sớm trước, sinh ra báo cáo "đối chiếu số dư cuối ngày EIB" để 
bên em đối chiếu trước với EIB, và tạo ra file import EIB. Còn OCB có file trễ hơn, thì 
sinh ra file đối chiếu, file import lúc này: cả dữ liệu EIB và OCB

SOLUTION: duy trì song song 2 module: 
    - BaoCaoDoiChieuVaImportEIB
    - BaoCaoDoiChieuVaImportOCB
"""


# DONE

def retry(func):
    def decorated_func(*args):
        while True:
            try:
                result = func(*args)
                break
            except ValueError:
                # ValueError: max() arg is an empty sequence: xảy ra khi chưa có email
                print('No files received from OCB. Still waiting...')
                time.sleep(30)
        return result

    return decorated_func


@retry
def _get_outlook_files(
    run_time=None,
):
    """
    Hàm này save file NH từ mail và trả về đường dẫn đến các file đó
    Chỉ trả về đường dẫn khi nhận được mail
    """

    info = get_info('daily',run_time)
    t0_date = info['end_date']
    period = info['period']

    outlook = Dispatch('outlook.application').GetNamespace("MAPI")
    for account in outlook.Accounts:
        # print which account is being logged
        print(f"Account {account.DeliveryStore.DisplayName} is being logged")

    inbox = outlook.Folders.Item(1).Folders['Inbox'].Folders['FilesFromOCB']
    messages = inbox.Items
    # Lọc ngày (chỉ xét các emails nhận tại ngày chạy)
    run_date = dt.datetime.strptime(t0_date,'%Y/%m/%d').date()

    def catch(m):
        try:
            result = m.ReceivedTime.date()==run_date
        except (Exception,):
            result = True
        return result

    messages = [mes for mes in messages if catch(mes)]

    # Tạo folder
    if not os.path.isdir(join(dept_folder,'FileFromBanks',period)):
        os.mkdir(join(dept_folder,'FileFromBanks',period))

    def f(x):  # Hàm lọc mail có sender là ngân hàng
        try:
            sender = x.SenderEmailAddress
        except (Exception,):
            try:
                sender = x.Sender.GetExchangeUser().PrimarySmtpAddress
            except (Exception,):
                sender = ''
        return re.search('.*@ocb.*',sender)

    # Tạo folder
    if not os.path.isdir(join(dept_folder,'FileFromBanks',period,'OCB')):
        os.mkdir(join(dept_folder,'FileFromBanks',period,'OCB'))

    # tạo DS emails được nhận vào ngày chạy và có sender là ngân hàng
    # (có thể nhiều hơn 1 email vì ngân hàng có thể gửi nhiều lần trong ngày)
    emails = [mes for mes in messages if f(mes) is not None]
    received_times = [email.ReceivedTime.time() for email in emails]

    # Delete existing files (if any)
    files = os.listdir(join(dept_folder,'FileFromBanks',period,'OCB'))
    for file in files:
        os.remove(join(dept_folder,'FileFromBanks',period,'OCB',file))
    # Get file
    max_time = max(received_times)
    idx = received_times.index(max_time)
    email = emails[idx]  # lấy mail cuối cùng cho đến lúc chạy
    for attachment in email.Attachments:
        check_file_extension = attachment.FileName.endswith('.xlsx') or attachment.FileName.endswith('.xls')
        if check_file_extension:
            attachment.SaveASFile(os.path.join(dept_folder,'FileFromBanks',period,'OCB',attachment.FileName))

    # Return đường dẫn
    file = os.listdir(join(dept_folder,'FileFromBanks',period,'OCB'))[0]  # should only have 1 file
    result = join(dept_folder,'FileFromBanks',period,'OCB',file)

    return result


# run daily
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    t0_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']
    if run_time is None:
        run_time = dt.datetime.now()

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    # đọc mail + download file
    file_path = _get_outlook_files(run_time)

    # đọc file OCB
    sent_data = pd.read_excel(
        file_path,
        usecols=['TÀI KHOẢN','TÊN KHÁCH HÀNG','SỐ DƯ'],
        dtype={'TÀI KHOẢN':object},
    ).drop_duplicates() # ngân hàng có lúc copy paste 2 lần ra 2 dòng trùng nhau
    sent_data.rename(
        {
            'TÀI KHOẢN':'bank_account',
            'TÊN KHÁCH HÀNG':'account_name',
            'SỐ DƯ':'s_balance',
        },
        axis=1,
        inplace=True
    )
    sent_data['bank'] = 'OCB'
    sent_data.set_index(['bank','bank_account'],inplace=True)

    # lấy số trên Flex
    output = pd.read_sql(
        f"""
        WITH 
        [c] AS (
            SELECT
                [cashflow_bank].[sub_account]
                , [cashflow_bank].[bank_account]
                , SUM([cashflow_bank].[outflow_amount]) [increase]
                , SUM([cashflow_bank].[inflow_amount]) [decrease]
            FROM 
                [cashflow_bank]
            WHERE 
                [cashflow_bank].[bank] = 'OCB'
                AND [cashflow_bank].[date] = '{t0_date}'
            GROUP BY
                [cashflow_bank].[sub_account], 
                [cashflow_bank].[bank_account]
        )
        , [in] AS (
            SELECT
                [money_in_out_transfer].[sub_account]
                , [money_in_out_transfer].[bank_account]
                , SUM([money_in_out_transfer].[amount]) [deposit]
            FROM 
                [money_in_out_transfer]
            WHERE 
                [money_in_out_transfer].[transaction_id] = '6692'
                AND [money_in_out_transfer].[date] = '{t0_date}'
                AND [money_in_out_transfer].[bank] = 'OCB'
                AND [money_in_out_transfer].[status] = N'Hoàn tất'
            GROUP BY
                [money_in_out_transfer].[sub_account], 
                [money_in_out_transfer].[bank_account]
        )
        , [out] AS (
            SELECT
                [money_in_out_transfer].[sub_account]
                , [money_in_out_transfer].[bank_account]
                , SUM([money_in_out_transfer].[amount]) [withdraw]
            FROM 
                [money_in_out_transfer]
            WHERE 
                [money_in_out_transfer].[transaction_id] = '6693'
                AND [money_in_out_transfer].[date] = '{t0_date}'
                AND [money_in_out_transfer].[bank] = 'OCB'
                AND [money_in_out_transfer].[status] = N'Hoàn tất'
            GROUP BY
                [money_in_out_transfer].[sub_account], 
                [money_in_out_transfer].[bank_account]
        )
        , [import] AS (
            SELECT 
                [vcf0051].[sub_account]
                , [imported_bank_balance].[bank_account]
                , [imported_bank_balance].[account_name]
                , [imported_bank_balance].[balance] [o_balance]
            FROM 
                [imported_bank_balance]
            LEFT JOIN [vcf0051]
                ON [vcf0051].[date] = [imported_bank_balance].[date]
                AND [vcf0051].[bank_account] = [imported_bank_balance].[bank_account]
            WHERE 
                [imported_bank_balance].[effective_date] = '{t0_date}'
                AND [imported_bank_balance].[bank_code] = 'OCB'
        )
        , [table] AS (
            SELECT
                COALESCE (
                    [c].[sub_account],
                    [in].[sub_account],
                    [out].[sub_account],
                    [import].[sub_account]
                    ) [sub_account]
                , 'OCB' [bank]
                , COALESCE (
                    [import].[bank_account],
                    [c].[bank_account],
                    [in].[bank_account],
                    [out].[bank_account]
                    ) [bank_account]
                , [import].[account_name]
                , ISNULL([import].[o_balance],0) [o_balance]
                , ISNULL([c].[increase],0) [increase]
                , ISNULL([c].[decrease],0) [decrease]
                , ISNULL([in].[deposit],0) [deposit]
                , ISNULL([out].[withdraw],0) [withdraw]
                , ISNULL([import].[o_balance],0) 
                    + ISNULL([c].[increase],0) 
                    - ISNULL([c].[decrease],0) 
                    + ISNULL([in].[deposit],0) 
                    - ISNULL([out].[withdraw],0) 
                    [e_balance]
            FROM
                [import]
            FULL JOIN [c] ON [c].[bank_account] = [import].[bank_account]
            FULL JOIN [in] ON [in].[bank_account] = [import].[bank_account]
            FULL JOIN [out] ON [out].[bank_account] = [import].[bank_account]
        )
        SELECT 
            [sub_account].[account_code]
            , [table].[bank]
            , [table].[bank_account]
            , [table].[account_name]
            , SUM([table].[o_balance]) [o_balance]
            , SUM([table].[increase]) [increase]
            , SUM([table].[decrease]) [decrease]
            , SUM([table].[deposit]) [deposit]
            , SUM([table].[withdraw]) [withdraw]
            , SUM([table].[e_balance]) [e_balance]
        FROM
            [table]
        LEFT JOIN [sub_account]
            ON [table].[sub_account] = [sub_account].[sub_account]
        GROUP BY
            [account_code], [bank], [bank_account], [account_name]
        -- Do bị NULL trên cột KEY của các bảng -> phải làm COALESCE -> phải thêm bước GROUP BY
        """,
        connect_DWH_CoSo,
        index_col=['bank','bank_account']
    )
    # fill sent date
    output = output.join(sent_data['s_balance'],how='outer')
    output.loc[:,'o_balance':] = output.loc[:,'o_balance':].fillna(0)
    output['diff'] = output['e_balance']-output['s_balance']

    # try to fill missing names using bank files
    output.loc[output['account_name'].isna(),'account_name'] = sent_data['account_name']

    # fill remaining missing values with empty string
    def name_converter(x):
        if x!=x or x is None:
            x = ''
        else:
            x = unidecode.unidecode(x)
        return x
    output['account_name'] = output['account_name'].map(name_converter)
    output.reset_index(inplace=True)

    # tách dataframe làm hai cho hai mục đích: đối chiếu và import
    # đối chiếu
    file_check = output.loc[output['diff']!=0].copy()

    # df import
    file_import = output[[
        'bank',
        'bank_account',
        'account_name',
        's_balance',
    ]].copy()
    file_import.columns = ['BANKCODE','ACCOUNT','ACCOUNT NAME','BALANCE']
    destination = join(dirname(__file__),'fileimportsodunganhang',period.replace('.',''))
    if not os.path.isdir(destination):
        os.mkdir(destination)
    file_import.to_pickle(f'{destination}/OCB.pickle')

    # get generated files (if any)
    files = os.listdir(destination)
    frames = []
    for file in files:
        frames.append(pd.read_pickle(f'{destination}/{file}'))
    file_import = pd.concat(frames)

    """
    1.	Nếu import file bank trước bắt đầu chạy batch cuối ngày. (trước 19h)
    -	Ngày TXDATE: N
    -	Ngày EFFECTIVEDATE: N+1 (nếu N là thứ 6 hay ngày nghỉ lễ, tết, 
        thì Ngày EFFECTIVEDATE là ngày làm việc đầu tiên sau nghỉ lễ, tết)
    -	Tên file import: NNmmyyyy
    2.	Nếu import file bank sau batch cuối ngày (do lỗi bank gửi file muộn…) 
        bên TTBT chưa thể import file bank lên trước 19h thì:
    -	Ngày TXDATE: N+1 (ngày làm việc tiếp theo)
    -	Ngày EFFECTIVEDATE: N+1 (ngày làm việc tiếp theo)
    -	Tên file import: (N+1)(N+1)mmyyyy
    """

    if run_time.time()<dt.time(hour=19):
        t = run_time
        txdate = t.strftime('%d/%m/%Y')
        effective_date = (t+dt.timedelta(days=1)).strftime('%d/%m/%Y')
    else:
        t = run_time+dt.timedelta(days=1)
        txdate = t.strftime('%d/%m/%Y')
        effective_date = t.strftime('%d/%m/%Y')

    file_import.insert(0,'EFFECTIVEDATE',[effective_date]*file_import.shape[0])
    file_import.insert(0,'TXDATE',[txdate]*file_import.shape[0])
    file_import.insert(0,'STT',file_import.index+1)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # ----------------- Write file Đối chiếu -----------------

    # đặt tên file không dấu theo yêu cầu của Nhàn TTBT
    file_name = f"Doi chieu so du OCB truoc khi import {txdate.replace('/','.')}.xlsx"
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
    text_subtotal_format = workbook.add_format(
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
    num_subtotal_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'text_wrap':True,
        }
    )
    worksheet = workbook.add_worksheet('Chênh Lệch')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',13)
    worksheet.set_column('B:B',11)
    worksheet.set_column('C:C',18)
    worksheet.set_column('D:D',35)
    worksheet.set_column('E:E',15)
    worksheet.set_column('F:L',13)
    worksheet.write_row(
        'A1',
        [
            'ACCOUNT CODE',
            'NGÂN HÀNG',
            'TÀI KHOẢN',
            'TÊN KHÁCH HÀNG',
            'SỐ DƯ ĐẦU KỲ',
            'TĂNG TIỀN',
            'GIẢM TIỀN',
            'NHẬP TIỀN',
            'RÚT TIỀN',
            'DỰ KIẾN CUỐI NGÀY',
            'SỐ NGÂN HÀNG GỬI',
            'LỆCH'
        ],
        header_format
    )
    worksheet.write_column('A2',file_check['account_code'],text_cell_format)
    worksheet.write_column('B2',file_check['bank'],text_cell_format)
    worksheet.write_column('C2',file_check['bank_account'],text_cell_format)
    worksheet.write_column('D2',file_check['account_name'],text_cell_format)
    worksheet.write_column('E2',file_check['o_balance'],num_cell_format)
    worksheet.write_column('F2',file_check['increase'],num_cell_format)
    worksheet.write_column('G2',file_check['decrease'],num_cell_format)
    worksheet.write_column('H2',file_check['deposit'],num_cell_format)
    worksheet.write_column('I2',file_check['withdraw'],num_cell_format)
    worksheet.write_column('J2',file_check['e_balance'],num_cell_format)
    worksheet.write_column('K2',file_check['s_balance'],num_cell_format)
    worksheet.write_column('L2',file_check['diff'],diff_cell_format)

    # Subtotal row
    subtotal_row = file_check.shape[0]+2
    worksheet.merge_range(f'A{subtotal_row}:D{subtotal_row}','Subtotal',text_subtotal_format)
    for col in 'EFGHIJKL':
        value = f'=SUBTOTAL(9,{col}2:{col}{subtotal_row-1})'
        worksheet.write(f'{col}{subtotal_row}',value,num_subtotal_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # ----------------- Write file Import -----------------

    file_name = f"{txdate.replace('/','')} file import.xlsx"
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ## Write sheet Sheet 1
    header_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_name':'Calibri',
            'font_size':11,
            'text_wrap':True,
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Calibri',
            'font_size':11,
            'text_wrap':True,
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Calibri',
            'font_size':11,
            'text_wrap':True,
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Calibri',
            'font_size':11,
            'num_format':'dd/mm/yyyy',
            'text_wrap':True,
        }
    )
    num_cell_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Calibri',
            'font_size':11,
            'num_format':'_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-',
            'text_wrap':True,
        }
    )
    num_subtotal_format = workbook.add_format(
        {
            'bold':True,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Calibri',
            'font_size':11,
            'num_format':'_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-',
            'text_wrap':True,
        }
    )
    worksheet = workbook.add_worksheet('Sheet 1')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:D',14)
    worksheet.set_column('E:E',18)
    worksheet.set_column('F:F',32)
    worksheet.set_column('G:G',20)
    worksheet.write('G1',f'=SUBTOTAL(9,G3:G{file_import.shape[0]+2})',num_subtotal_format)
    worksheet.write_row(
        'A2',
        ['STT','TXDATE','BANK','BANKCODE','ACCOUNT','ACCOUNT NAME','BALANCE'],
        header_format
    )
    worksheet.write_column('A3',file_import['STT'],text_center_format)
    worksheet.write_column('B3',file_import['TXDATE'],date_format)
    worksheet.write_column('C3',file_import['EFFECTIVEDATE'],date_format)
    worksheet.write_column('D3',file_import['BANKCODE'],text_left_format)
    worksheet.write_column('E3',file_import['ACCOUNT'],text_center_format)
    worksheet.write_column('F3',file_import['ACCOUNT NAME'],text_left_format)
    worksheet.write_column('G3',file_import['BALANCE'],num_cell_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')