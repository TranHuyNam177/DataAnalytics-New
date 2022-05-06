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
    start_account = pd.read_sql(
        f"""
        SELECT 
            [account].[account_code], 
            [account].[account_type]
        FROM [account]
        WHERE
            ([account].[date_of_open] <= '{bdate(start_date,-1)}') 
        AND (
            ([account].[date_of_close] IS NULL) 
            OR ([account].[date_of_close] > '{bdate(start_date,-1)}' AND [account].[date_of_close] != '2099-12-31' )
            ) -- mot so tai khoan dong rat lau roi duoc gan ngay dong la ngay nay
        """,
        connect_DWH_CoSo,
        index_col='account_code',
    ).squeeze()
    end_account = pd.read_sql(
        f"""
        SELECT 
            [account].[account_code], 
            [account].[account_type]
        FROM
            [account]
        WHERE
            ([account].[date_of_open] <= '{end_date}') 
        AND (
            ([account].[date_of_close] IS NULL) 
            OR ([account].[date_of_close] > '{end_date}' AND [account].[date_of_close] != '2099-12-31' )
            ) -- mot so tai khoan dong rat lau roi duoc gan ngay dong la ngay nay
        """,
        connect_DWH_CoSo,
        index_col='account_code',
    ).squeeze()
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
    authorization.loc[authorization['authorized_person_name']=='CTY CP CHỨNG KHOÁN PHÚ HƯNG','authorized_person_address'] = CompanyAddress
    mapper = lambda x:'Thường' if x.startswith('Thường') else 'Ký Quỹ'
    contract_type['contract_type'] = contract_type['contract_type'].map(mapper)

    margin_account = contract_type.loc[contract_type['contract_type']=='Ký Quỹ','account_code']
    period_account.loc[period_account.index.isin(margin_account),'remark'] = 'TKKQ'
    period_account['remark'].fillna('',inplace=True)
    period_account.loc[period_account['account_type'].str.startswith('Cá nhân'),'entity_type'] = 'CN'
    period_account.loc[period_account['account_type'].str.startswith('Tổ chức'),'entity_type'] = 'TC'
    open_mask = (period_account['date_of_open'].dt.month==month)&(period_account['date_of_open'].dt.year==year)
    account_open = period_account.loc[open_mask]
    close_mask = (period_account['date_of_close'].dt.month==month)&(period_account['date_of_close'].dt.year==year)
    account_close = period_account.loc[close_mask]

    # Tinh bien dong TK
    open_ind_domestic = (account_open['account_type']=='Cá nhân trong nước').sum()
    open_ins_domestic = (account_open['account_type']=='Tổ chức trong nước').sum()
    open_ind_foreign = (account_open['account_type']=='Cá nhân nước ngoài').sum()
    open_ins_foreign = (account_open['account_type']=='Tổ chức nước ngoài').sum()
    close_ind_domestic = (account_close['account_type']=='Cá nhân trong nước').sum()
    close_ins_domestic = (account_close['account_type']=='Tổ chức trong nước').sum()
    close_ind_foreign = (account_close['account_type']=='Cá nhân nước ngoài').sum()
    close_ins_foreign = (account_close['account_type']=='Tổ chức nước ngoài').sum()
    open_total_domestic = open_ind_domestic+open_ins_domestic
    open_total_foreign = open_ind_foreign+open_ins_foreign
    close_total_domestic = close_ind_domestic+close_ins_domestic
    close_total_foreign = close_ind_foreign+close_ins_foreign
    open_total = open_total_domestic+open_total_foreign
    close_total = close_total_domestic+close_total_foreign

    # Dau ky
    start_account_count = start_account.value_counts()
    opening_ind_domestic = start_account_count.loc['Cá nhân trong nước']
    opening_ins_domestic = start_account_count.loc['Tổ chức trong nước']
    opening_ind_foreign = start_account_count.loc['Cá nhân nước ngoài']
    opening_ins_foreign = start_account_count.loc['Tổ chức nước ngoài']

    opening_total_domestic = opening_ind_domestic+opening_ins_domestic
    opening_total_foreign = opening_ind_foreign+opening_ins_foreign
    opening_total = opening_total_domestic+opening_total_foreign

    # Cuoi ky
    end_account_count = end_account.value_counts()
    closing_ind_domestic = end_account_count.loc['Cá nhân trong nước']
    closing_ins_domestic = end_account_count.loc['Tổ chức trong nước']
    closing_ind_foreign = end_account_count.loc['Cá nhân nước ngoài']
    closing_ins_foreign = end_account_count.loc['Tổ chức nước ngoài']

    closing_totaL_domestic = closing_ind_domestic+closing_ins_domestic
    closing_totaL_foreign = closing_ind_foreign+closing_ins_foreign
    closing_total = closing_totaL_domestic+closing_totaL_foreign

    ###########################################################################
    ###########################################################################
    ###########################################################################
    ########################### Write to HNX file #############################
    ###########################################################################
    ###########################################################################
    ###########################################################################

    file_name = f'Danh sách KH đóng mở ủy quyền tài khoản PHS HNX {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ## Write sheet TONG HOP
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
        }
    )
    sup_note_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    kinhgui_format = workbook.add_format(
        {
            'bold':False,
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    str_bold = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    str_bold_center = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    stt_column_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    header_cell_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'border':1,
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
        }
    )
    normal_cell_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
        }
    )
    header_value = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
        }
    )
    normal_value = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
        }
    )
    sheet_tonghop = workbook.add_worksheet('Tông hợp')
    sheet_tonghop.hide_gridlines(option=2)

    sheet_tonghop.set_column('A:A',8)
    sheet_tonghop.set_column('B:B',26)
    sheet_tonghop.set_column('C:C',22)
    sheet_tonghop.set_column('D:E',20.5)
    sheet_tonghop.set_column('F:F',22)
    sheet_tonghop.set_row(1,34)

    sup_title = r'BÁO CÁO TÌNH HÌNH ĐÓNG/MỞ TÀI KHOẢN VÀ KHÁCH HÀNG ỦY QUYỀN'
    sheet_tonghop.merge_range('A1:F1',sup_title,sup_title_format)
    note_title = '(Kèm theo Quy chế Thành viên giao dịch thị trường niêm yết và ' \
                 'thị trường đăng ký giao dịch tại SGDCKHN ban hành theo Quyết định ' \
                 'số 430/QĐ-SGDHN ngày 03/07/2019 của Tổng Giám Đốc Sở Giao Dịch ' \
                 'Chứng Khoán Hà Nội)'
    sheet_tonghop.merge_range('A2:F2',note_title,sup_note_format)
    sheet_tonghop.merge_range('A4:F4','Kính gửi: Sở Giao dịch Chứng khoán Hà Nội',kinhgui_format)
    sheet_tonghop.merge_range('A6:B6','Tên thành viên:',str_bold)
    sheet_tonghop.merge_range('C6:D6',CompanyName,str_bold)
    sheet_tonghop.merge_range('A7:B7','Mã thành viên:',str_bold)
    sheet_tonghop.write('C7',CompanyCode,str_bold)
    sheet_tonghop.write('E7','Kỳ báo cáo:',str_bold_center)
    sheet_tonghop.write('F7',f'Tháng {end_date[5:7]} Năm {end_date[:4]}',str_bold)
    sheet_tonghop.merge_range('A9:B9','I. Tổng hợp',str_bold)
    sheet_tonghop.merge_range('A11:A12','STT',header_format)
    sheet_tonghop.merge_range('B11:B12','KHÁCH HÀNG',header_format)
    sheet_tonghop.merge_range('C11:F11','SỐ LƯỢNG TÀI KHOẢN',header_format)
    sheet_tonghop.write_row('C12',['Đầu kỳ','Mở trong kỳ','Đóng trong kỳ','Cuối kỳ'],header_format)
    sheet_tonghop.write_column('A13',['1','','','2','','',''],stt_column_format)
    sheet_tonghop.write('B13','TRONG NƯỚC',header_cell_format)
    sheet_tonghop.write('B16','NƯỚC NGOÀI',header_cell_format)
    sheet_tonghop.write_column('B14',['     Cá nhân','     Tổ chức'],normal_cell_format)
    sheet_tonghop.write_column('B17',['     Cá nhân','     Tổ chức'],normal_cell_format)
    sheet_tonghop.write('B19','TỔNG CỘNG',header_cell_format)
    closing_column = np.array(
        [
            closing_totaL_domestic,
            closing_ind_domestic,
            closing_ins_domestic,
            closing_totaL_foreign,
            closing_ind_foreign,
            closing_ins_foreign,
            closing_total,
        ]
    )
    close_column = np.array(
        [
            close_total_domestic,
            close_ind_domestic,
            close_ins_domestic,
            close_total_foreign,
            close_ind_foreign,
            close_ins_foreign,
            close_total,
        ]
    )
    open_column = np.array(
        [
            open_total_domestic,
            open_ind_domestic,
            open_ins_domestic,
            open_total_foreign,
            open_ind_foreign,
            open_ins_foreign,
            open_total,
        ]
    )
    opening_column = np.array(
        [
            opening_total_domestic,
            opening_ind_domestic,
            opening_ins_domestic,
            opening_total_foreign,
            opening_ind_foreign,
            opening_ins_foreign,
            opening_total,
        ]
    )
    value_array = np.array([opening_column,open_column,close_column,closing_column]).transpose()
    for col in range(4):
        for row in range(7):
            if row in [0,3,6]:
                fmt = header_value
            else:
                fmt = normal_value
            sheet_tonghop.write(12+row,2+col,value_array[row,col],fmt)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ## Write sheet MO TAi KHOAN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':14,
            'text_wrap':True,
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
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
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
        }
    )
    text_center_format = workbook.add_format(
        {
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
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
        }
    )
    sheet_motaikhoan = workbook.add_worksheet('Mở TK')
    sheet_motaikhoan.hide_gridlines(option=2)
    # set column width
    sheet_motaikhoan.set_column('A:A',4.6)
    sheet_motaikhoan.set_column('B:B',19.7)
    sheet_motaikhoan.set_column('C:C',10.1)
    sheet_motaikhoan.set_column('D:D',14.1)
    sheet_motaikhoan.set_column('E:E',57.3)
    sheet_motaikhoan.set_column('F:F',11.9)
    sheet_motaikhoan.set_column('G:G',16.6)
    sheet_motaikhoan.set_column('H:H',7.9)
    sheet_motaikhoan.set_column('I:I',11.7)
    sheet_motaikhoan.set_column('J:J',12.4)
    sheet_motaikhoan.set_column('K:K',9.1)

    sheet_motaikhoan.set_row(0,30)
    sheet_motaikhoan.merge_range('A1:K1','II. Danh sách khách hàng mở tài khoản',sup_title_format)
    headers = [
        'STT',
        'Tên khách hàng',
        'Mã TK',
        'Số CMND/ Hộ chiếu/Giấy ĐKKD',
        'Địa chỉ',
        'Ngày cấp',
        'Nơi cấp',
        'Loại hình',
        'Ngày mở',
        'Quốc tịch',
        'Ghi chú',
    ]
    sheet_motaikhoan.write_row('A2',headers,header_format)
    header_num = [f'({i})' for i in np.arange(len(headers))+1]
    sheet_motaikhoan.write_row('A3',header_num,header_format)
    stt_column = [i for i in np.arange(0,account_open.shape[0])+1]
    sheet_motaikhoan.write_column('A4',stt_column,text_center_format)
    sheet_motaikhoan.write_column('B4',account_open['customer_name'],text_left_format)
    sheet_motaikhoan.write_column('C4',account_open.index,text_center_format)
    sheet_motaikhoan.write_column('D4',account_open['customer_id_number'],text_center_format)
    sheet_motaikhoan.write_column('E4',account_open['address'],text_left_format)
    sheet_motaikhoan.write_column('F4',account_open['date_of_issue'].map(convertNaTtoSpaceString),date_format)
    sheet_motaikhoan.write_column('G4',account_open['place_of_issue'].map(convertNaTtoSpaceString),text_left_format)
    sheet_motaikhoan.write_column('H4',account_open['entity_type'],text_center_format)
    sheet_motaikhoan.write_column('I4',account_open['date_of_open'],date_format)
    sheet_motaikhoan.write_column('J4',account_open['nationality'],text_center_format)
    sheet_motaikhoan.write_column('K4',account_open['remark'],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet DONG TAi KHOAN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':14,
            'text_wrap':True,
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
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
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
        }
    )
    text_center_format = workbook.add_format(
        {
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
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
        }
    )
    sheet_dongtaikhoan = workbook.add_worksheet('Đóng TK')
    sheet_dongtaikhoan.hide_gridlines(option=2)
    # set column width
    sheet_dongtaikhoan.set_column('A:A',4.4)
    sheet_dongtaikhoan.set_column('B:B',26.3)
    sheet_dongtaikhoan.set_column('C:C',13.1)
    sheet_dongtaikhoan.set_column('D:D',11.9)
    sheet_dongtaikhoan.set_column('E:E',36.2)
    sheet_dongtaikhoan.set_column('F:F',13.7)
    sheet_dongtaikhoan.set_column('G:G',12.6)
    sheet_dongtaikhoan.set_column('H:H',6.7)
    sheet_dongtaikhoan.set_column('I:J',11.7)
    sheet_dongtaikhoan.set_column('K:K',9.7)
    sheet_dongtaikhoan.set_column('L:L',8.4)
    sheet_dongtaikhoan.set_default_row(27)  # set all row height = 27
    sheet_dongtaikhoan.set_row(0,30)
    sheet_dongtaikhoan.set_row(2,15)
    sheet_dongtaikhoan.merge_range('A1:L1','III. Danh sách khách hàng đóng tài khoản',sup_title_format)
    headers = [
        'STT',
        'Tên khách hàng',
        'Mã tài khoản',
        'Số CMND/ Hộ chiếu/Giấy ĐKKD',
        'Địa chỉ',
        'Ngày cấp',
        'Nơi cấp',
        'Loại hình',
        'Ngày mở TK',
        'Ngày đóng TK',
        'Quốc tịch',
        'Ghi chú',
    ]
    sheet_dongtaikhoan.write_row('A2',headers,header_format)
    header_num = [f'({i})' for i in np.arange(len(headers))+1]
    sheet_dongtaikhoan.write_row('A3',header_num,header_format)
    stt_column = [i for i in np.arange(0,account_close.shape[0])+1]
    sheet_dongtaikhoan.write_column('A4',stt_column,text_left_format)
    sheet_dongtaikhoan.write_column('B4',account_close['customer_name'],text_left_format)
    sheet_dongtaikhoan.write_column('C4',account_close.index,text_center_format)
    sheet_dongtaikhoan.write_column('D4',account_close['customer_id_number'],text_center_format)
    sheet_dongtaikhoan.write_column('E4',account_close['address'],text_left_format)
    sheet_dongtaikhoan.write_column('F4',account_close['date_of_issue'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('G4',account_close['place_of_issue'],text_center_format)
    sheet_dongtaikhoan.write_column('H4',account_close['entity_type'],text_center_format)
    sheet_dongtaikhoan.write_column('I4',account_close['date_of_open'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('J4',account_close['date_of_close'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('K4',account_close['nationality'],text_center_format)
    sheet_dongtaikhoan.write_column('L4',account_close.shape[0]*[''],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet THAY DOI THONG TIN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':14,
            'text_wrap':True,
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
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
            'font_size':10,
            'text_wrap':True,
        }
    )
    sheet_thaydoithongtin = workbook.add_worksheet('Thay đổi thông tin')
    sheet_thaydoithongtin.hide_gridlines(option=2)

    sheet_thaydoithongtin.set_column('A:A',5)
    sheet_thaydoithongtin.set_column('B:B',20.5)
    sheet_thaydoithongtin.set_column('C:D',11.6)
    sheet_thaydoithongtin.set_column('E:E',14.1)
    sheet_thaydoithongtin.set_column('F:F',9.7)
    sheet_thaydoithongtin.set_column('G:G',11.9)
    sheet_thaydoithongtin.set_column('H:H',13.6)
    sheet_thaydoithongtin.set_column('I:I',9.5)
    sheet_thaydoithongtin.set_column('J:J',13.9)
    sheet_thaydoithongtin.set_column('K:L',28)
    sheet_thaydoithongtin.set_column('M:N',8.5)
    sheet_thaydoithongtin.set_column('O:P',5)
    sheet_thaydoithongtin.set_row(0,30)
    sheet_thaydoithongtin.set_row(1,38)
    sheet_thaydoithongtin.set_row(2,39)

    sheet_thaydoithongtin.merge_range('A1:P1','IV. Danh sách khách hàng thay đổi thông tin',sup_title_format)
    sheet_thaydoithongtin.merge_range('A2:A3','STT',header_format)
    sheet_thaydoithongtin.merge_range('B2:B3','Tên khách hàng',header_format)
    sheet_thaydoithongtin.merge_range('C2:C3','Mã TK cũ',header_format)
    sheet_thaydoithongtin.merge_range('D2:D3','Ngày thay đổi thông tin',header_format)
    sheet_thaydoithongtin.merge_range('E2:J2','Thay đổi thông tin về CMND/ Hộ chiếu/ Giấy ĐKKD',header_format)
    sheet_thaydoithongtin.merge_range('K2:L2','Thay đổi thông tin về địa chỉ',header_format)
    sheet_thaydoithongtin.merge_range('M2:N2','Thay đổi TT về Q.tịch',header_format)
    sheet_thaydoithongtin.merge_range('O2:P2','Thay đổi thông tin về Ghi chú',header_format)
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
    sheet_thaydoithongtin.write_row('E3',sub_header,header_format)
    sheet_thaydoithongtin.write_row(
        'A4',
        [f'{i}' for i in np.arange(16)+1],  # cong them 2 cot ghi chu
        header_format,
    )
    sheet_thaydoithongtin.write_column(
        'A5',
        [f'({i})' for i in np.arange(customer_information_change.shape[0])+1],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'B5',
        customer_information_change['customer_name'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'C5',
        customer_information_change.index,
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'D5',
        customer_information_change['date_of_change'].map(convertNaTtoSpaceString),
        date_format,
    )
    sheet_thaydoithongtin.write_column(
        'E5',
        customer_information_change['old_id_number'].map(convertNaTtoSpaceString),
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'F5',
        customer_information_change['old_date_of_issue'].map(convertNaTtoSpaceString),
        date_format,
    )
    sheet_thaydoithongtin.write_column(
        'G5',
        customer_information_change['old_place_of_issue'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'H5',
        customer_information_change['new_id_number'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'I5',
        customer_information_change['new_date_of_issue'].map(convertNaTtoSpaceString),
        date_format,
    )
    sheet_thaydoithongtin.write_column(
        'J5',
        customer_information_change['new_place_of_issue'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'K5',
        customer_information_change['old_address'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'L5',
        customer_information_change['new_address'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'M5',
        customer_information_change['old_nationality'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'N5',
        customer_information_change['new_nationality'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'O5',
        ['']*customer_information_change.shape[0],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'P5',
        ['']*customer_information_change.shape[0],
        text_left_format,
    )

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet UY QUYEN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':14,
            'text_wrap':True,
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
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
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10,
            'text_wrap':True,
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
            'text_wrap':True,
            'bg_color':'#FFFF00'
        }
    )
    sheet_uyquyen = workbook.add_worksheet('Ủy quyền')
    sheet_uyquyen.hide_gridlines(option=2)

    sheet_uyquyen.set_column('A:A',3)
    sheet_uyquyen.set_column('B:B',18.9)
    sheet_uyquyen.set_column('C:D',11.6)
    sheet_uyquyen.set_column('E:E',38.9)
    sheet_uyquyen.set_column('F:F',9.8)
    sheet_uyquyen.set_column('G:G',20.4)
    sheet_uyquyen.set_column('H:H',13.7)
    sheet_uyquyen.set_column('I:I',34)
    sheet_uyquyen.set_column('J:J',13.3)
    sheet_uyquyen.set_column('K:K',5.8)
    sheet_uyquyen.set_row(0,30)
    sheet_uyquyen.set_row(1,51)

    sheet_uyquyen.merge_range('A1:K1','V. Danh sách khách hàng ủy quyền',sup_title_format)
    headers = [
        'STT',
        'Tên khách hàng ủy quyền',
        'Mã TK',
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD người UQ',
        'Địa chỉ  người UQ',
        'Ngày Uỷ quyền',
        'Tên người nhận uỷ quyền',
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD người nhận UQ',
        'Địa chỉ người nhận UQ',
        'Phạm vi uỷ quyền',
        'Ghi chú',
    ]
    authorization['date_of_authorization'] = authorization['date_of_authorization'].map(convertNaTtoSpaceString)
    sheet_uyquyen.write_row('A2',headers,header_format)
    sheet_uyquyen.write_row('A3',[f'({i})' for i in np.arange(len(headers))+1],header_format)
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
        sheet_uyquyen.write(row+3,0,row+1,fmt1)
        sheet_uyquyen.write(row+3,1,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_name')],fmt2)
        sheet_uyquyen.write(row+3,2,authorization.index[row],fmt1)
        sheet_uyquyen.write(row+3,3,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_id')],fmt1)
        sheet_uyquyen.write(row+3,4,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_address')],fmt2)
        sheet_uyquyen.write(row+3,5,authorization.iloc[row,authorization.columns.get_loc('date_of_authorization')],fmt3)
        sheet_uyquyen.write(row+3,6,authorization.iloc[row,authorization.columns.get_loc('authorized_person_name')],fmt1)
        sheet_uyquyen.write(row+3,7,authorization.iloc[row,authorization.columns.get_loc('authorized_person_id')],fmt1)
        sheet_uyquyen.write(row+3,8,authorization.iloc[row,authorization.columns.get_loc('authorized_person_address')],fmt2)
        sheet_uyquyen.write(row+3,9,authorization.iloc[row,authorization.columns.get_loc('scope_of_authorization')],fmt1)
        sheet_uyquyen.write(row+3,10,'',fmt2)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet THAY DOI UY QUYEN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':14,
            'text_wrap':True,
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10,
        }
    )
    signature_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
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
            'font_size':10,
            'text_wrap':True,
        }
    )
    sheet_thaydoiuyquyen = workbook.add_worksheet('Thay đổi ủy quyền')
    sheet_thaydoiuyquyen.hide_gridlines(option=2)

    sheet_thaydoiuyquyen.set_column('A:A',4)
    sheet_thaydoiuyquyen.set_column('B:B',22)
    sheet_thaydoiuyquyen.set_column('C:E',12)
    sheet_thaydoiuyquyen.set_column('F:F',22)
    sheet_thaydoiuyquyen.set_column('G:J',12)
    sheet_thaydoiuyquyen.set_column('K:L',17)
    sheet_thaydoiuyquyen.set_column('M:P',12)
    sheet_thaydoiuyquyen.set_row(0,30)
    sheet_thaydoiuyquyen.set_row(1,30)
    sheet_thaydoiuyquyen.set_row(2,36)

    sheet_thaydoiuyquyen.merge_range(
        'A1:P1',
        'VI. Danh sách khách hàng chấm dứt, thay đổi nội dung ủy quyền',
        sup_title_format
    )
    sheet_thaydoiuyquyen.merge_range('A2:A3','STT',header_format)
    sheet_thaydoiuyquyen.merge_range('B2:B3','Tên khách hàng uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('C2:C3','Mã TK',header_format)
    sheet_thaydoiuyquyen.merge_range('D2:D3','Số CMND/ Hộ chiếu/ Giấy ĐKKD của khách hàng UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('E2:E3','Ngày Uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('F2:F3','Tên người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('G2:G3','Ngày chấm dứt Uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('H2:H3','Ngày thay đổi ND uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('I2:J2','Thay đổi CMND/ Hộ chiếu người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('K2:L2','Thay đổi địa chỉ người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('M2:N2','Thay đổi phạm vi uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('O2:P2','Thay đổi thời hạn ủy quyền',header_format)
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
    sheet_thaydoiuyquyen.write_row('I3',sub_header,header_format)
    sheet_thaydoiuyquyen.write_row('A4',[f'({i})' for i in np.arange(16)+1],header_format)
    sheet_thaydoiuyquyen.write_column('A5',np.arange(authorization_change.shape[0])+1,text_center_format)
    sheet_thaydoiuyquyen.write_column('B5',authorization_change['authorizing_person_name'],text_left_format)
    sheet_thaydoiuyquyen.write_column('C5',authorization_change.index,text_center_format)
    sheet_thaydoiuyquyen.write_column('D5',authorization_change['authorizing_person_id'],text_left_format)
    sheet_thaydoiuyquyen.write_column('E5',authorization_change['date_of_authorization'].map(convertNaTtoSpaceString),date_format)
    sheet_thaydoiuyquyen.write_column('F5',authorization_change['authorized_person_name'],text_center_format)
    sheet_thaydoiuyquyen.write_column('G5',authorization_change['date_of_termination'].map(convertNaTtoSpaceString),date_format)
    sheet_thaydoiuyquyen.write_column('H5',authorization_change['date_of_change'].map(convertNaTtoSpaceString),date_format)
    sheet_thaydoiuyquyen.write_column('I5',authorization_change['old_authorized_person_id'],text_center_format)
    sheet_thaydoiuyquyen.write_column('J5',authorization_change['new_authorized_person_id'],text_center_format)
    sheet_thaydoiuyquyen.write_column('K5',authorization_change['old_authorized_person_address'],text_center_format)
    sheet_thaydoiuyquyen.write_column('L5',authorization_change['new_authorized_person_address'],text_center_format)
    sheet_thaydoiuyquyen.write_column('M5',authorization_change['old_scope_of_authorization'],text_center_format)
    sheet_thaydoiuyquyen.write_column('N5',authorization_change['new_scope_of_authorization'],text_center_format)
    sheet_thaydoiuyquyen.write_column('O5',authorization_change['old_end_date'].map(convertNaTtoSpaceString),date_format)
    sheet_thaydoiuyquyen.write_column('P5',authorization_change['new_end_date'].map(convertNaTtoSpaceString),date_format)

    row_of_signature = 3+authorization_change.shape[0]+2
    sheet_thaydoiuyquyen.set_row(row_of_signature-1,55)
    sheet_thaydoiuyquyen.set_row(row_of_signature+2,80)
    sheet_thaydoiuyquyen.merge_range(row_of_signature,1,row_of_signature,4,'Người lập',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+1,1,row_of_signature+1,4,'(Ký, ghi rõ họ tên)',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+3,1,row_of_signature+3,4,'ĐIỀN HỌ TÊN VÀO Ô',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature,10,row_of_signature,15,'ĐIỀN CHỨC DANH VÀO Ô',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+1,10,row_of_signature+1,15,'(Ký, ghi rõ họ tên)',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+3,10,row_of_signature+3,15,'ĐIỀN HỌ TÊN VÀO Ô',signature_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
