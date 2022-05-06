from automation.trading_service.giaodichluuky import *


def run(
    run_time=None,
):

    start = time.time()
    info = get_info('monthly',run_time)
    run_time = info['run_time']
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    month = int(period.split('.')[0])
    year = int(period.split('.')[1])

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))
        
    period_account = pd.read_sql(
        f"""
        SELECT
            [account].[account_type],
            [account].[account_code],
            [account].[customer_name],
            [account].[nationality],
            [account].[address],
            [account].[customer_id_number],
            [account].[date_of_issue],
            [account].[place_of_issue],
            [account].[date_of_open],
            [account].[date_of_close]
        FROM [account]
        WHERE 
            ([account].[date_of_open] BETWEEN '{start_date}' AND '{end_date}')
        OR 
            ([account].[date_of_close] BETWEEN '{start_date}' AND '{end_date}')
        """,
        connect_DWH_CoSo,
        index_col='account_code',
    )
    contract_type = pd.read_sql(
        """
        SELECT
            [customer_information].[sub_account],
            [sub_account].[account_code],
            [customer_information].[contract_code],
            [customer_information].[contract_type]
        FROM 
            [customer_information]
        LEFT JOIN 
            [sub_account] 
        ON 
            [customer_information].[sub_account] = [sub_account].[sub_account]
        """,
        connect_DWH_CoSo,
        index_col='sub_account',
    )
    customer_information_change = pd.read_sql(
        f"""
        SELECT 
            [rcf0005].[account_code],
            [account].[customer_name],
            [rcf0005].[date_of_change],
            [rcf0005].[old_id_number],
            [rcf0005].[new_id_number],
            [rcf0005].[old_date_of_issue],
            [rcf0005].[new_date_of_issue],
            [rcf0005].[old_place_of_issue],
            [rcf0005].[new_place_of_issue],
            [rcf0005].[old_address],
            [rcf0005].[new_address],
            [rcf0005].[old_nationality],
            [rcf0005].[new_nationality]
        FROM 
            [rcf0005]
        LEFT JOIN
            [account]
        ON
            [account].[account_code] = [rcf0005].[account_code]
        WHERE 
            [rcf0005].[date_of_change] BETWEEN '{start_date}' AND '{end_date}'
        ORDER BY
            [rcf0005].[account_code],
            [rcf0005].[date_of_change]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    ).drop_duplicates()  # RCF0005 có duplicated data
    authorization = pd.read_sql(
        f"""
        SELECT * 
        FROM 
            [authorization] 
        WHERE 
            [authorization].[date_of_authorization] BETWEEN '{start_date}' AND '{end_date}'
        AND 
            [authorization].[scope_of_authorization] IS NOT NULL
        AND 
            [authorization].[scope_of_authorization] <> 'I,IV,V'
        """,
        connect_DWH_CoSo,
        index_col='account_code',
    )
    # Highlight cac uy quyen duoc mo moi chi de dang ky uy quyen them (rule ben DVKH)
    highlight_account = pd.read_sql(
        f"""
        SELECT
            [authorization_change].[account_code]
        FROM 
            [authorization_change]
        WHERE 
            [authorization_change].[new_end_date] BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo,
    )
    authorization_change = pd.read_sql(
        f"""
        SELECT *
        FROM 
            [authorization_change]
        WHERE 
            [authorization_change].[date_of_change] BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    authorization['scope_of_authorization'] = 'I,II,IV,V,VII,IX,X'
    authorization.loc[authorization[
                          'authorized_person_name']=='CTY CP CHỨNG KHOÁN PHÚ HƯNG','authorized_person_address'] = CompanyAddress
    mapper = lambda x:'Thường' if x.startswith('Thường') else 'Ký Quỹ'
    contract_type['contract_type'] = contract_type['contract_type'].map(mapper)

    margin_account = contract_type.loc[contract_type['contract_type']=='Ký Quỹ','account_code']
    period_account.loc[period_account.index.isin(margin_account),'remark'] = 'TKKQ'
    period_account['remark'].fillna('',inplace=True)
    period_account.dropna(subset=['account_type'],inplace=True)  # bỏ tài khoản quỹ
    period_account.loc[period_account['account_type'].str.startswith('Cá nhân'),'entity_type'] = 'CN'
    period_account.loc[period_account['account_type'].str.startswith('Tổ chức'),'entity_type'] = 'TC'
    open_mask = (period_account['date_of_open'].dt.month==month)&(period_account['date_of_open'].dt.year==year)
    account_open = period_account.loc[open_mask]
    close_mask = (period_account['date_of_close'].dt.month==month)&(period_account['date_of_close'].dt.year==year)
    account_close = period_account.loc[close_mask]

    ###########################################################################
    ###########################################################################
    ###########################################################################
    ######################## Write to HOSE old file ###########################
    ###########################################################################
    ###########################################################################
    ###########################################################################

    file_name = f'Danh sách KH đóng mở ủy quyền tài khoản PHS (old) HOSE {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ## Write sheet MO TAi KHOAN
    headline_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'bg_color':'#99CCFF',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_motaikhoan = workbook.add_worksheet('Mở TK')
    sheet_motaikhoan.hide_gridlines(option=2)
    # set column width
    sheet_motaikhoan.set_column('A:A',6.5)
    sheet_motaikhoan.set_column('B:B',35.6)
    sheet_motaikhoan.set_column('C:D',16)
    sheet_motaikhoan.set_column('E:E',58)
    sheet_motaikhoan.set_column('F:F',15.9)
    sheet_motaikhoan.set_column('G:G',25)
    sheet_motaikhoan.set_column('H:H',5.7)
    sheet_motaikhoan.set_column('I:I',13)
    sheet_motaikhoan.set_column('J:J',11.6)
    sheet_motaikhoan.set_column('K:K',5)
    sheet_motaikhoan.set_default_row(27)  # set all row height = 27
    sheet_motaikhoan.set_row(0,15)
    sheet_motaikhoan.set_row(1,15)
    sheet_motaikhoan.set_row(2,15)
    sheet_motaikhoan.set_row(3,15)
    sheet_motaikhoan.set_row(4,22)
    sheet_motaikhoan.set_row(5,30)
    sheet_motaikhoan.merge_range('A1:K1',CompanyName,headline_format)
    sheet_motaikhoan.merge_range('A2:K2',CompanyAddress,headline_format)
    sheet_motaikhoan.merge_range('A3:K3',CompanyPhoneNumber,headline_format)
    month = end_date[5:7]
    year = end_date[:4]
    sheet_motaikhoan.merge_range('A4:K4',f'DANH SÁCH KHÁCH HÀNG MỞ TÀI KHOẢN THÁNG {month}.{year}',sup_title_format)
    sheet_motaikhoan.merge_range('A5:K5',f'Kính gửi : SỞ GIAO DỊCH CHỨNG KHOÁN TP.HCM',sup_title_format)
    headers = [
        'STT',
        'HỌ TÊN',
        'MÃ TÀI KHOẢN',
        'CMND/HỘ CHIẾU/GP ĐKKD',
        'ĐỊA CHỈ',
        'NGÀY CẤP',
        'NƠI CẤP',
        'LOẠI HÌNH',
        'NGÀY MỞ',
        'QUỐC TỊCH',
        'GHI CHÚ',
    ]
    sheet_motaikhoan.write_row('A6',headers,header_format)
    stt_column = np.arange(0,account_open.shape[0])+1
    sheet_motaikhoan.write_column('A7',stt_column,text_center_format)
    sheet_motaikhoan.write_column('B7',account_open['customer_name'],text_left_format)
    sheet_motaikhoan.write_column('C7',account_open.index,text_center_format)
    sheet_motaikhoan.write_column('D7',account_open['customer_id_number'],text_center_format)
    sheet_motaikhoan.write_column('E7',account_open['address'],text_left_format)
    sheet_motaikhoan.write_column('F7',account_open['date_of_issue'].map(convertNaTtoSpaceString),date_format)
    sheet_motaikhoan.write_column('G7',account_open['place_of_issue'],text_center_format)
    sheet_motaikhoan.write_column('H7',account_open['entity_type'],text_center_format)
    sheet_motaikhoan.write_column('I7',account_open['date_of_open'].map(convertNaTtoSpaceString),date_format)
    sheet_motaikhoan.write_column('J7',account_open['nationality'],text_center_format)
    sheet_motaikhoan.write_column('K7',account_open['remark'],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################
    # Write sheet DONG TAi KHOAN

    headline_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'bg_color':'#99CCFF',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_dongtaikhoan = workbook.add_worksheet('Đóng TK')
    sheet_dongtaikhoan.hide_gridlines(option=2)

    sheet_dongtaikhoan.set_column('A:A',4)
    sheet_dongtaikhoan.set_column('B:B',30)
    sheet_dongtaikhoan.set_column('C:D',15)
    sheet_dongtaikhoan.set_column('E:E',60)
    sheet_dongtaikhoan.set_column('F:F',13)
    sheet_dongtaikhoan.set_column('G:G',21)
    sheet_dongtaikhoan.set_column('H:H',6)
    sheet_dongtaikhoan.set_column('I:I',13)
    sheet_dongtaikhoan.set_column('J:K',12)
    sheet_dongtaikhoan.set_column('L:L',5)
    sheet_dongtaikhoan.set_default_row(27)  # set all row height = 27
    sheet_dongtaikhoan.set_row(0,15)
    sheet_dongtaikhoan.set_row(1,15)
    sheet_dongtaikhoan.set_row(2,15)
    sheet_dongtaikhoan.set_row(3,15)
    sheet_dongtaikhoan.set_row(4,22)
    sheet_dongtaikhoan.set_row(5,30)
    sheet_dongtaikhoan.merge_range('A1:L1',CompanyName,headline_format)
    sheet_dongtaikhoan.merge_range('A2:L2',CompanyAddress,headline_format)
    sheet_dongtaikhoan.merge_range('A3:L3',CompanyPhoneNumber,headline_format)
    month = end_date[5:7]
    year = end_date[:4]
    sheet_dongtaikhoan.merge_range('A4:L4',f'DANH SÁCH KHÁCH HÀNG ĐÓNG TÀI KHOẢN THÁNG {month}.{year}',sup_title_format)
    sheet_dongtaikhoan.merge_range('A5:L5',f'Kính gửi : SỞ GIAO DỊCH CHỨNG KHOÁN TP.HCM',sup_title_format)
    headers = [
        'STT',
        'HỌ TÊN',
        'MÃ TÀI KHOẢN',
        'CMND/HỘ CHIẾU/GP ĐKKD',
        'ĐỊA CHỈ',
        'NGÀY CẤP',
        'NƠI CẤP',
        'LOẠI HÌNH',
        'NGÀY MỞ',
        'NGÀY ĐÓNG',
        'QUỐC TỊCH',
        'GHI CHÚ',
    ]
    sheet_dongtaikhoan.write_row('A6',headers,header_format)
    stt_column = np.arange(0,account_close.shape[0])+1
    sheet_dongtaikhoan.write_column('A7',stt_column,text_center_format)
    sheet_dongtaikhoan.write_column('B7',account_close['customer_name'],text_left_format)
    sheet_dongtaikhoan.write_column('C7',account_close.index,text_center_format)
    sheet_dongtaikhoan.write_column('D7',account_close['customer_id_number'],text_center_format)
    sheet_dongtaikhoan.write_column('E7',account_close['address'],text_left_format)
    sheet_dongtaikhoan.write_column('F7',account_close['date_of_issue'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('G7',account_close['place_of_issue'],text_center_format)
    sheet_dongtaikhoan.write_column('H7',account_close['entity_type'],text_center_format)
    sheet_dongtaikhoan.write_column('I7',account_close['date_of_open'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('J7',account_close['date_of_close'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('K7',account_close['nationality'],text_center_format)
    sheet_dongtaikhoan.write_column('L7',account_close.shape[0]*[''],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet THAY DOI THONG TIN
    headline_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'bg_color':'#C0C0C0',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_thaydoithongtin = workbook.add_worksheet('Thay đổi thông tin')
    sheet_thaydoithongtin.hide_gridlines(option=2)
    sheet_thaydoithongtin.set_column('A:A',4)
    sheet_thaydoithongtin.set_column('B:B',25)
    sheet_thaydoithongtin.set_column('C:F',13)
    sheet_thaydoithongtin.set_column('G:G',24)
    sheet_thaydoithongtin.set_column('H:I',13)
    sheet_thaydoithongtin.set_column('J:J',24)
    sheet_thaydoithongtin.set_column('K:L',26)
    sheet_thaydoithongtin.set_column('M:P',8)
    sheet_thaydoithongtin.set_row(0,15)
    sheet_thaydoithongtin.set_row(1,15)
    sheet_thaydoithongtin.set_row(2,15)
    sheet_thaydoithongtin.set_row(3,15)
    sheet_thaydoithongtin.set_row(4,22)
    sheet_thaydoithongtin.set_row(5,30)
    sheet_thaydoithongtin.merge_range('A1:P1',CompanyName,headline_format)
    sheet_thaydoithongtin.merge_range('A2:P2',CompanyAddress,headline_format)
    sheet_thaydoithongtin.merge_range('A3:P3',CompanyPhoneNumber,headline_format)
    month = end_date[5:7]
    year = end_date[:4]
    sheet_thaydoithongtin.merge_range('A4:P4',f'DANH SÁCH KHÁCH HÀNG THAY ĐỖI THÔNG TIN THÁNG {month}.{year}',
                                      sup_title_format)
    sheet_thaydoithongtin.merge_range('A5:P5',f'Kính gửi : SỞ GIAO DỊCH CHỨNG KHOÁN TP.HCM',sup_title_format)
    sheet_thaydoithongtin.merge_range('A6:A7','STT',header_format)
    sheet_thaydoithongtin.merge_range('B6:B7','Tên khách hàng',header_format)
    sheet_thaydoithongtin.merge_range('C6:C7','Mã TK cũ',header_format)
    sheet_thaydoithongtin.merge_range('D6:D7','Ngày thay đổi thông tin',header_format)
    sheet_thaydoithongtin.merge_range('E6:J6','Thay đổi thông tin về CMND/ Hộ chiếu/ Giấy ĐKKD',header_format)
    sheet_thaydoithongtin.merge_range('K6:L6','Thay đổi thông tin về địa chỉ',header_format)
    sheet_thaydoithongtin.merge_range('M6:N6','Thay đổi TT về Q.tịch',header_format)
    sheet_thaydoithongtin.merge_range('O6:P6','Thay đổi thông tin về Ghi chú',header_format)
    sub_header = [
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD cũ',
        'Ngày cấp',
        'Nơi cấp',
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD mới',
        'Ngày cấp',
        'Nơi cấp',
        'Địa chỉ cũ',
        'Địa chỉ mới',
        'Quốc tịch cũ',
        'Quốc tịch mới',
        'Ghi chú cũ',
        'Ghi chú mới',
    ]
    sheet_thaydoithongtin.write_row('E7',sub_header,header_format)
    sheet_thaydoithongtin.write_column(
        'A8',
        [f'({i})' for i in np.arange(customer_information_change.shape[0])+1],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'B8',
        customer_information_change['customer_name'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'C8',
        customer_information_change.index,
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'D8',
        customer_information_change['date_of_change'].map(convertNaTtoSpaceString),
        date_format,
    )
    sheet_thaydoithongtin.write_column(
        'E8',
        customer_information_change['old_id_number'].map(convertNaTtoSpaceString),
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'F8',
        customer_information_change['old_date_of_issue'].map(convertNaTtoSpaceString),
        date_format,
    )
    sheet_thaydoithongtin.write_column(
        'G8',
        customer_information_change['old_place_of_issue'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'H8',
        customer_information_change['new_id_number'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'I8',
        customer_information_change['new_date_of_issue'].map(convertNaTtoSpaceString),
        date_format,
    )
    sheet_thaydoithongtin.write_column(
        'J8',
        customer_information_change['new_place_of_issue'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'K8',
        customer_information_change['old_address'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'L8',
        customer_information_change['new_address'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'M8',
        customer_information_change['old_nationality'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'N8',
        customer_information_change['new_nationality'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'O8',
        ['']*customer_information_change.shape[0],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'P8',
        ['']*customer_information_change.shape[0],
        text_left_format,
    )
    ###########################################################################
    ###########################################################################
    ###########################################################################

    ## Write sheet UY QUYEN
    headline_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'bg_color':'#C0C0C0',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_highlight_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'bg_color':'#FFFF00'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_highlight_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'bg_color':'#FFFF00'
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_highlight_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10,
            'bg_color':'#FFFF00'
        }
    )
    sheet_uyquyen = workbook.add_worksheet('Ủy Quyền')
    sheet_uyquyen.hide_gridlines(option=2)

    # set column width
    sheet_uyquyen.set_column('A:A',4)
    sheet_uyquyen.set_column('B:B',21)
    sheet_uyquyen.set_column('C:D',14)
    sheet_uyquyen.set_column('E:E',45)
    sheet_uyquyen.set_column('F:F',11)
    sheet_uyquyen.set_column('G:G',18)
    sheet_uyquyen.set_column('H:H',14)
    sheet_uyquyen.set_column('I:I',28)
    sheet_uyquyen.set_column('J:J',14)
    sheet_uyquyen.set_default_row(39)  # set all row height = 39
    sheet_uyquyen.set_row(0,15)
    sheet_uyquyen.set_row(1,15)
    sheet_uyquyen.set_row(2,15)
    sheet_uyquyen.set_row(3,15)
    sheet_uyquyen.set_row(4,22)
    sheet_uyquyen.set_row(5,30)
    sheet_uyquyen.merge_range('A1:J1',CompanyName,headline_format)
    sheet_uyquyen.merge_range('A2:J2',CompanyAddress,headline_format)
    sheet_uyquyen.merge_range('A3:J3',CompanyPhoneNumber,headline_format)
    month = end_date[5:7]
    year = end_date[:4]
    sheet_uyquyen.merge_range('A4:J4',f'DANH SÁCH KHÁCH HÀNG ỦY QUYỀN THÁNG {month}.{year}',sup_title_format)
    sheet_uyquyen.merge_range('A5:J5',f'Kính gửi : SỞ GIAO DỊCH CHỨNG KHOÁN TP.HCM',sup_title_format)
    headers = [
        'STT',
        'TÊN NGƯỜI ỦY QUYỀN ',
        'MÃ TÀI KHOẢN',
        'CMND/HỘ CHIẾU/GP ĐKKD',
        'ĐỊA CHỈ',
        'NGÀY ỦY QUYỀN',
        'TÊN NGƯỜI ỦY QUYỀN',
        'CMND/HỘ CHIẾU/GP ĐKKD',
        'ĐỊA CHỈ',
        'PV.UQ',
    ]
    authorization['date_of_authorization'] = authorization['date_of_authorization'].map(convertNaTtoSpaceString)
    sheet_uyquyen.write_row('A6',headers,header_format)
    for row in range(authorization.shape[0]):
        ticker = authorization.index[row]
        if ticker in highlight_account.values:
            fmt1 = text_highlight_center_format
            fmt2 = text_highlight_left_format
            fmt3 = date_highlight_format
        else:
            fmt1 = text_center_format
            fmt2 = text_left_format
            fmt3 = date_format
        sheet_uyquyen.write(row+6,0,row+1,fmt1)
        sheet_uyquyen.write(row+6,1,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_name')],
                            fmt2)
        sheet_uyquyen.write(row+6,2,authorization.index[row],fmt1)
        sheet_uyquyen.write(row+6,3,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_id')],fmt1)
        sheet_uyquyen.write(row+6,4,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_address')],
                            fmt2)
        sheet_uyquyen.write(row+6,5,authorization.iloc[row,authorization.columns.get_loc('date_of_authorization')],fmt3)
        sheet_uyquyen.write(row+6,6,authorization.iloc[row,authorization.columns.get_loc('authorized_person_name')],
                            fmt1)
        sheet_uyquyen.write(row+6,7,authorization.iloc[row,authorization.columns.get_loc('authorized_person_id')],fmt1)
        sheet_uyquyen.write(row+6,8,authorization.iloc[row,authorization.columns.get_loc('authorized_person_address')],
                            fmt2)
        sheet_uyquyen.write(row+6,9,authorization.iloc[row,authorization.columns.get_loc('scope_of_authorization')],
                            fmt1)

    ###########################################################################
    ###########################################################################
    ###########################################################################
    # Write to sheet Thay Doi Uy Quyen

    headline_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'bg_color':'#C0C0C0',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    signature_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_thaydoiuyquyen = workbook.add_worksheet('Thay Đổi Ủy Quyền')
    sheet_thaydoiuyquyen.hide_gridlines(option=2)

    sheet_thaydoiuyquyen.set_column('A:A',4)
    sheet_thaydoiuyquyen.set_column('B:B',20)
    sheet_thaydoiuyquyen.set_column('C:D',12)
    sheet_thaydoiuyquyen.set_column('E:E',11)
    sheet_thaydoiuyquyen.set_column('F:F',20)
    sheet_thaydoiuyquyen.set_column('G:H',11)
    sheet_thaydoiuyquyen.set_column('I:J',12)
    sheet_thaydoiuyquyen.set_column('K:L',20)
    sheet_thaydoiuyquyen.set_column('O:P',11)
    sheet_thaydoiuyquyen.set_row(0,15)
    sheet_thaydoiuyquyen.set_row(1,15)
    sheet_thaydoiuyquyen.set_row(2,15)
    sheet_thaydoiuyquyen.set_row(3,15)
    sheet_thaydoiuyquyen.set_row(4,22)
    sheet_thaydoiuyquyen.set_row(5,30)
    sheet_thaydoiuyquyen.set_row(6,33)
    sheet_thaydoiuyquyen.merge_range('A1:P1',CompanyName,headline_format)
    sheet_thaydoiuyquyen.merge_range('A2:P2',CompanyAddress,headline_format)
    sheet_thaydoiuyquyen.merge_range('A3:P3',CompanyPhoneNumber,headline_format)
    month = end_date[5:7]
    year = end_date[:4]
    sheet_thaydoiuyquyen.merge_range('A4:P4',f'DANH SÁCH KHÁCH HÀNG THAY ĐÔI ỦY QUYỀN GIAO DỊCH THÁNG {month}.{year}',
                                     sup_title_format)
    sheet_thaydoiuyquyen.merge_range('A5:P5',f'Kính gửi : SỞ GIAO DỊCH CHỨNG KHOÁN TP.HCM',sup_title_format)
    sheet_thaydoiuyquyen.merge_range('A6:A7','STT',header_format)
    sheet_thaydoiuyquyen.merge_range('B6:B7','Tên khách hàng uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('C6:C7','Mã TK',header_format)
    sheet_thaydoiuyquyen.merge_range('D6:D7','Số CMND/ Hộ chiếu/ Giấy ĐKKD của khách hàng UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('E6:E7','Ngày Uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('F6:F7','Tên người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('G6:G7','Ngày chấm dứt Uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('H6:H7','Ngày thay đổi ND uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('I6:J6','Thay đổi CMND/ Hộ chiếu người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('K6:L6','Thay đổi địa chỉ người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('M6:N6','Thay đổi phạm vi uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('O6:P6','Thay đổi thời hạn ủy quyền',header_format)
    sub_header = [
        'Số CMND/ Hộ chiếu cũ',
        'Số CMND/ Hộ chiếu mới',
        'Địa chỉ cũ',
        'Địa chỉ mới',
        'Phạm vi uỷ quyền cũ',
        'Phạm vi uỷ quyền mới',
        'Thời hạn cũ',
        'Thời hạn mới',
    ]
    sheet_thaydoiuyquyen.write_row('I7',sub_header,header_format)
    sheet_thaydoiuyquyen.write_row('A8',[f'({i})' for i in np.arange(16)+1],text_center_format)
    sheet_thaydoiuyquyen.write_column('A9',np.arange(authorization_change.shape[0])+1,text_center_format)
    sheet_thaydoiuyquyen.write_column('B9',authorization_change['authorizing_person_name'],text_left_format)
    sheet_thaydoiuyquyen.write_column('C9',authorization_change.index,text_center_format)
    sheet_thaydoiuyquyen.write_column('D9',authorization_change['authorizing_person_id'],text_left_format)
    sheet_thaydoiuyquyen.write_column('E9',authorization_change['date_of_authorization'].map(convertNaTtoSpaceString),
                                      date_format)
    sheet_thaydoiuyquyen.write_column('F9',authorization_change['authorized_person_name'],text_center_format)
    sheet_thaydoiuyquyen.write_column('G9',authorization_change['date_of_termination'].map(convertNaTtoSpaceString),
                                      text_center_format)
    sheet_thaydoiuyquyen.write_column('H9',authorization_change['date_of_change'].map(convertNaTtoSpaceString),
                                      date_format)
    sheet_thaydoiuyquyen.write_column('I9',authorization_change['old_authorized_person_id'],text_center_format)
    sheet_thaydoiuyquyen.write_column('J9',authorization_change['new_authorized_person_id'],text_center_format)
    sheet_thaydoiuyquyen.write_column('K9',authorization_change['old_authorized_person_address'],text_center_format)
    sheet_thaydoiuyquyen.write_column('L9',authorization_change['new_authorized_person_address'],text_center_format)
    sheet_thaydoiuyquyen.write_column('M9',authorization_change['old_scope_of_authorization'],text_center_format)
    sheet_thaydoiuyquyen.write_column('N9',authorization_change['new_scope_of_authorization'],text_center_format)
    sheet_thaydoiuyquyen.write_column('O9',authorization_change['old_end_date'].map(convertNaTtoSpaceString),
                                      date_format)
    sheet_thaydoiuyquyen.write_column('P9',authorization_change['new_end_date'].map(convertNaTtoSpaceString),
                                      date_format)

    row_of_signature = 7+authorization_change.shape[0]+2
    run_day = convert_int(run_time.day)
    run_month = convert_int(run_time.month)
    run_year = convert_int(run_time.year)
    sheet_thaydoiuyquyen.merge_range(row_of_signature,10,row_of_signature,15,
                                     f'TP.HCM, ngày {run_day} tháng {run_month} năm {run_year}',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+1,10,row_of_signature+1,15,'ĐIỀN CHỨC DANH VÀO Ô',
                                     signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+2,10,row_of_signature+2,15,'(Ký, ghi rõ họ tên, đóng dấu)',
                                     signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+8,10,row_of_signature+8,15,'ĐIỀN HỌ TÊN VÀO Ô',signature_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
