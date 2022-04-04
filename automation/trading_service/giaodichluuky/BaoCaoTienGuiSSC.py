from automation.trading_service.giaodichluuky import *


def run(
    run_time=None,
):
    start = time.time()
    info = get_info('weekly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    # # chờ file tài chính
    # day = period.split('.')[0][-2:]
    # month = period.split('.')[1]
    # year = period.split('.')[2]
    # fin_dept_folder = f'THANG {month}'
    # fin_dept_file = f'Daily Bank Report {day} {month} {year}.xls'
    # fin_dept_path = r"\\192.168.10.120\Data Warehouse\Datawarehouse\Datawarehouse\Bank Report"
    # while not isfile(join(fin_dept_path,fin_dept_folder,fin_dept_file)):
    #     print('No files received from Finance. Still listening...')
    #     time.sleep(15)

    # get data
    machitieu = pd.read_excel(
        join(dirname(dirname(__file__)),'machitieu.xlsx'),
        sheet_name='account_type',
        dtype=str,
        index_col='customer_type',
        squeeze=True
    )
    all_account = pd.read_sql(
        f"""
        SELECT 
            [account].[account_type], 
            COUNT(DISTINCT relationship.account_code) [count]
        FROM 
            [relationship]
        LEFT JOIN 
            [customer_information] 
        ON 
            [customer_information].[sub_account] = [relationship].[sub_account] 
        LEFT JOIN 
            [account]
        ON 
            [account].[account_code] = [relationship].[account_code] 
        WHERE 
            [relationship].[date] BETWEEN N'{start_date}' AND N'{end_date}'
            AND [customer_information].[status] IN ('A','B','N','P') 
            AND [account].[account_type] IN (
                N'Cá nhân trong nước',
                N'Cá nhân nước ngoài',
                N'Tổ chức trong nước',
                N'Tổ chức nước ngoài'
            )
        GROUP BY [account_type]
        ORDER BY [account_type]
        """,
        connect_DWH_CoSo,
    )
    all_account['machitieu'] = all_account['account_type'].map(machitieu)
    all_account.sort_values('machitieu',inplace=True)
    all_account = all_account.append({'account_type':'Tổng','count':all_account['count'].sum()},ignore_index=True)
    all_account.fillna('',inplace=True)
    all_account.index = pd.RangeIndex(start=1,stop=6,name='STT')

    # =========================================================================
    # Write to Báo cáo tiền gửi và số lượng tài khoản
    file_name = f'Báo cáo tiền gửi và số lượng tài khoản {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    # =========================================================================
    # Write to sheet T.Bia
    bia_sheet = workbook.add_worksheet('T.Bia')
    bia_sheet.hide_gridlines(option=2)
    # set column width
    bia_sheet.set_column('A:A',3)
    bia_sheet.set_column('B:B',19)
    bia_sheet.set_column('C:C',34)
    bia_sheet.set_column('D:D',39)
    bia_sheet.set_default_row(15)
    bia_sheet.set_row(6,20)

    bold_headline_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':10,
            'valign':'vcenter',
        }
    )
    normal_headline_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':10,
            'valign':'vcenter',
        }
    )
    title_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':14,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    italic_large_fmt = workbook.add_format(
        {
            'italic':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'vcenter',
        }
    )
    italic_small_fmt = workbook.add_format(
        {
            'italic':True,
            'font_name':'Times New Roman',
            'font_size':9,
            'align':'center',
            'valign':'vcenter',
        }
    )
    ghichu_fmt = workbook.add_format(
        {
            'italic':True,
            'font_name':'Times New Roman',
            'font_size':9,
            'valign':'vcenter',
        }
    )
    underline_fmt = workbook.add_format(
        {
            'underline':1,
            'font_name':'Times New Roman',
            'font_size':9,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bold_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'align':'center',
            'valign':'vcenter',
        }
    )
    header_center_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'bg_color':'#fabf8f',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    header_left_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'bg_color':'#fabf8f',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_left_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_center_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    bia_sheet.write('B2',CompanyName,bold_headline_fmt)
    bia_sheet.write('B3',CompanyAddress,normal_headline_fmt)
    bia_sheet.write('B4',CompanyPhoneNumber,normal_headline_fmt)
    bia_sheet.merge_range('B7:D7','BÁO CÁO TIỀN GỬI NHÀ ĐẦU TƯ',title_fmt)
    bia_sheet.write('D10','Thông tư số 210/2012/TT-BTC Phụ lục 17',italic_large_fmt)
    bia_sheet.write('B11','STT',header_center_fmt)
    bia_sheet.write_row('C11',['Nội dung','Tên sheet'],header_left_fmt)
    bia_sheet.write_column('B12',['1','2'],incell_center_fmt)
    bia_sheet.write_column('C12',['Số lượng tài khoản nhà đầu tư','Tiền gửi giao dịch của nhà đầu tư'],incell_left_fmt)
    bia_sheet.write_url('D12',"internal:'I_06211'!A1",incell_left_fmt,'I_06211')
    bia_sheet.write_url('D13',"internal:'II_06212'!A1",incell_left_fmt,'II_06212')
    bia_sheet.write_row('B14',['']*3,header_center_fmt)
    bia_sheet.write('B16','Ghi chú:',underline_fmt)
    bia_sheet.write_column(
        'C16',[
            'Không đổi tên sheet',
            'Những chỉ tiêu không có số liệu có thể không phải trình bày nhưng không được đánh lại “Mã chỉ tiêu”.',
            'Không được xóa cột trên sheet'
        ],
        ghichu_fmt,
    )
    bia_sheet.merge_range('A23:B23','Người lập biểu',bold_fmt)
    bia_sheet.merge_range('A24:B24','(Ký, họ tên)',italic_small_fmt)
    bia_sheet.merge_range('A29:B29','Điền tên người lập',bold_fmt)
    bia_sheet.write('D22',f'Lập, ngày {end_date[-2:]} tháng {end_date[5:7]} năm {end_date[:4]}',italic_large_fmt)
    bia_sheet.write('D23','Tổng Giám đốc',bold_fmt)
    bia_sheet.write('D24','(Ký, họ tên, đóng dấu)',italic_small_fmt)
    bia_sheet.write('D29','Chen Chia Ken',bold_fmt)

    # =========================================================================
    sheet_I06211 = workbook.add_worksheet('I_06211')
    sheet_I06211.hide_gridlines(option=2)
    # set column width
    sheet_I06211.set_column('A:A',5)
    sheet_I06211.set_column('B:B',28)
    sheet_I06211.set_column('C:D',22)
    header_left_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'bg_color':'#fabf8f',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    header_center_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'bg_color':'#fabf8f',
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_normal_text_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_bold_text_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'border':1,
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_value_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'num_format':'#,##0',
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_bold_value_fmt = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'num_format':'#,##0',
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    incell_code_fmt = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':11,
            'num_format':'0',
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    sheet_I06211.write('A1',all_account.index.name,header_left_fmt)
    sheet_I06211.write('B1','Loại khách hàng',header_left_fmt)
    sheet_I06211.write_row('C1',['Số lượng','Mã chỉ tiêu'],header_center_fmt)
    sheet_I06211.write_column('A2',all_account.index,incell_code_fmt)
    sheet_I06211.write_column('B2',all_account['account_type'],incell_normal_text_fmt)
    sheet_I06211.write_column('C2',all_account['count'],incell_value_fmt)
    sheet_I06211.write_column('D2',all_account['machitieu'],incell_code_fmt)
    # overwrite last row
    sheet_I06211.write('A6',5,incell_bold_value_fmt)
    sheet_I06211.write('B6','Tổng',incell_bold_text_fmt)
    sheet_I06211.write('C6',all_account.loc[all_account['account_type']=='Tổng','count'],incell_bold_value_fmt)
    sheet_I06211.write('D6','',incell_code_fmt)

    # =========================================================================
    # Write to sheet II_06212
    #
    # table = pd.read_excel(
    #     join(fin_dept_path,fin_dept_folder,fin_dept_file),
    #     skiprows=7,
    #     usecols=[2,3,5],
    #     names=['bank_name','account_no','balance']
    # )
    # bank_name_mapper = {
    #     "0021100002115004": "Ngân hàng Phương Đông - PGD Tú Xương",
    #     "140114851002285": "Ngân hàng Xuất Nhập Khẩu VN - CN Sài Gòn",
    #     "160314851020212": "Ngân hàng Xuất Nhập Khẩu VN - CN Hải Phòng",
    #     "1017816-069": "Ngân hàng TNHH Indovina - CN Phú Mỹ Hưng",
    #     "1017816-066": "Ngân hàng TNHH Indovina - CN Phú Mỹ Hưng",
    #     "11910000132943": "Ngân hàng BIDV - CN NKKN",
    #     "26110002677688": "Ngân hàng BIDV - CN Tràng An",
    #     "007.100.1264078": "Ngân hàng Vietcombank - CN TPHCM",
    #     "147001536591": "Ngân hàng Vietin - CN4 TPHCM",
    #     "122000069726": "Ngân hàng Vietin - CN4 TPHCM",
    # }
    # table = table.loc[table['account_no'].isin(bank_name_mapper.keys())]
    # table.fillna(0,inplace=True)
    # table.index = pd.RangeIndex(start=1,stop=table.shape[0]+1)
    # table['bank_name'] = table['account_no'].map(bank_name_mapper)
    # machitieu = pd.read_excel(
    #     join(dirname(dirname(__file__)),'machitieu.xlsx'),
    #     sheet_name='bank_account',
    #     dtype=str,
    #     usecols=['account_no','code'],
    #     index_col='account_no',
    #     squeeze=True
    # )
    # table['machitieu'] = table['account_no'].map(machitieu)
    # table['account_no'].replace('1017816-069','1017816-069 (trong nước)',inplace=True)
    # table['account_no'].replace('1017816-066','1017816-066 (nước ngoài)',inplace=True)
    # sum_balance = table['balance'].sum()
    #
    # sheet_II06211 = workbook.add_worksheet('II_06212')
    # sheet_II06211.hide_gridlines(option=2)
    #
    # sheet_II06211.set_column('A:A',4)
    # sheet_II06211.set_column('B:B',50)
    # sheet_II06211.set_column('C:C',32)
    # sheet_II06211.set_column('D:D',23)
    # sheet_II06211.set_column('E:E',13)
    #
    # headline_fmt = workbook.add_format(
    #     {
    #         'bold': True,
    #         'font_name': 'Times New Roman',
    #         'font_size': 14,
    #         'align': 'center',
    #         'valign': 'vcenter',
    #         'text_wrap': True,
    #     }
    # )
    # header_fmt = workbook.add_format(
    #     {
    #         'bold': True,
    #         'font_name': 'Times New Roman',
    #         'font_size': 11,
    #         'border': 1,
    #         'bg_color': '#fabf8f',
    #         'align':'center',
    #         'valign': 'vcenter',
    #         'text_wrap': True,
    #     }
    # )
    # incell_left_text_fmt = workbook.add_format(
    #     {
    #         'font_name': 'Times New Roman',
    #         'font_size': 11,
    #         'border': 1,
    #         'align': 'left',
    #         'valign': 'vcenter',
    #         'text_wrap': True,
    #     }
    # )
    # incell_center_text_fmt = workbook.add_format(
    #     {
    #         'font_name': 'Times New Roman',
    #         'font_size': 11,
    #         'border': 1,
    #         'align': 'center',
    #         'valign': 'vcenter',
    #         'text_wrap': True,
    #     }
    # )
    # incell_value_fmt = workbook.add_format(
    #     {
    #         'font_name': 'Times New Roman',
    #         'font_size': 11,
    #         'num_format': '#,##0',
    #         'border': 1,
    #         'align': 'right',
    #         'valign': 'vcenter',
    #         'text_wrap': True,
    #     }
    # )
    # incell_bold_value_fmt = workbook.add_format(
    #     {
    #         'bold': True,
    #         'font_name': 'Times New Roman',
    #         'font_size': 11,
    #         'num_format': '#,##0',
    #         'border': 1,
    #         'align': 'right',
    #         'valign': 'vcenter',
    #         'text_wrap': True,
    #     }
    # )
    # incell_bold_text_fmt = workbook.add_format(
    #     {
    #         'bold': True,
    #         'font_name': 'Times New Roman',
    #         'font_size': 11,
    #         'border': 1,
    #         'align': 'center',
    #         'valign': 'vcenter',
    #         'text_wrap': True,
    #     }
    # )
    # sheet_II06211.merge_range('A1:E1','Tiền gửi giao dịch của nhà đầu tư',headline_fmt)
    # sheet_II06211.write_row('A3',['STT','Ngân hàng nhận tiền gửi','Tài khoản tại ngân hàng tiền gửi','Số dư trên tài khoản','Mã chỉ tiêu'],header_fmt)
    # sheet_II06211.write_column('A4',table.index,incell_center_text_fmt)
    # sheet_II06211.write_column('B4',table['bank_name'],incell_left_text_fmt)
    # sheet_II06211.write_column('C4',table['account_no'],incell_left_text_fmt)
    # sheet_II06211.write_column('D4',table['balance'],incell_value_fmt)
    # sheet_II06211.write_column('E4',table['machitieu'],incell_center_text_fmt)
    # sum_row, sum_col = table.shape[0]+3, 0
    # sheet_II06211.merge_range(sum_row,sum_col,sum_row,2,'Tổng',incell_bold_text_fmt)
    # sheet_II06211.write(sum_row,3,sum_balance,incell_bold_value_fmt)
    # sheet_II06211.write(sum_row,4,'',incell_bold_value_fmt)

    # =========================================================================

    writer.close()

    # =========================================================================
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
