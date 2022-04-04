from automation.trading_service.thanhtoanbutru import *


# DONE
def run(  # BC thang
    run_time=None,
):

    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    start_day = start_date[-2:]
    start_month = start_date[5:7]
    start_year = start_date[:4]
    end_day = end_date[-2:]
    end_month = end_date[5:7]
    end_year = end_date[:4]

    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir(join(dept_folder,folder_name,period))

    ###################################################
    ###################################################
    ###################################################

    # query KSNB hàng tháng
    ksnb_6692_6693_query = pd.read_sql(
        f"""
        SELECT
            [money_in_out_transfer].[sub_account], 
            [money_in_out_transfer].[date], 
            [money_in_out_transfer].[time], 
            [money_in_out_transfer].[transaction_id],
            [account].[account_code], 
            [money_in_out_transfer].[bank], 
            [account].[customer_name]
        FROM 
            [money_in_out_transfer]
        LEFT JOIN 
            [sub_account]
        ON 
            [money_in_out_transfer].[sub_account] = [sub_account].[sub_account]
        LEFT JOIN 
            [account]
        ON 
            [sub_account].[account_code] = [account].[account_code]
        WHERE 
            ([money_in_out_transfer].[bank] IN ('EIB','OCB') OR [money_in_out_transfer].[bank] IS NULL) -- cac TK ngat lien ket NH se bi NULL tren he thong
        AND 
            [money_in_out_transfer].[transaction_id] IN ('6692','6693')
        AND 
            [money_in_out_transfer].[date] BETWEEN '{start_date}' AND '{end_date}'
        ORDER BY 
            [money_in_out_transfer].[date], [money_in_out_transfer].[time] 
            ASC
    """,
        connect_DWH_CoSo
    )
    ksnb_1187_query = pd.read_sql(
        f"""
        SELECT 
            [money_in_out_transfer].[sub_account],
            [transactional_record].[date],
            [transactional_record].[time],
            [transactional_record].[transaction_id], 
            [account].[account_code],
            [money_in_out_transfer].[bank],
            [account].[customer_name]
        FROM 
            [transactional_record]
        LEFT JOIN 
            [employee] 
        ON 
            [employee].[employee_id] = [transactional_record].[approver]
        LEFT JOIN 
            [department]
        ON 
            [department].[dept_id] = [employee].[dept_id]
        LEFT JOIN 
            [money_in_out_transfer]
        ON 
            [money_in_out_transfer].[sub_account] = [transactional_record].[sub_account]
        LEFT JOIN 
            [sub_account] 
        ON 
            [sub_account].[sub_account] = [transactional_record].[sub_account]
        LEFT JOIN 
            [account] 
        ON 
            [sub_account].[account_code] = [account].[account_code]
        WHERE 
            [transactional_record].[date] BETWEEN '{start_date}' AND '{end_date}'
        AND 
            [transactional_record].[transaction_id] = '1187'
        AND 
            ([transactional_record].[approver] = '0832' OR [department].[dept_name] = 'Margin settlement')
        ORDER BY 
            [transactional_record].[date]
        """,
        connect_DWH_CoSo
    )
    table = pd.concat([ksnb_6692_6693_query,ksnb_1187_query])

    # Tìm ngân hàng liên kết cuối cùng trước khi ngắt liên kết
    missing_bank_mask = table['bank'].isna()
    table.loc[missing_bank_mask,'bank'] = table.loc[missing_bank_mask,'sub_account'].map(get_bank_name)
    table = table.drop('sub_account',axis=1)

    ###################################################
    ###################################################
    ###################################################

    # Write file excel Bao cao doi chieu file ngan hang
    file_name = f'Báo Cáo File Ghi Âm Gửi KSNB {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )

    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    sheet_title_vn_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':14
        }
    )
    sheet_title_eng_format = workbook.add_format(
        {
            'bold':True,
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':14,
            'color':'#00B050'
        }
    )
    dia_diem_vn_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12,
        }
    )
    dia_diem_eng_format = workbook.add_format(
        {
            'bold':True,
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12,
            'color':'#00B050'
        }
    )
    header_vn_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'bottom':0,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12,
        }
    )
    header_eng_format = workbook.add_format(
        {
            'bold':True,
            'italic':True,
            'border':1,
            'top':0,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12,
            'color':'#00B050'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    call_type_vn_left_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'bottom',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    call_type_eng_left_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'text_wrap':True,
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    text_right_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    customer_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True
        }
    )
    time_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'h:mm:ss'
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'dd/mm/yyyy'
        }
    )
    headers_vn = [
        'Loại cuộc gọi',
        'Stt',
        'Ngày',
        'Giờ',
        'Số tài khoản',
        'Ngân hàng',
        'Tên Khách hàng',
        'Số điện thoại',
        'Ghi chú',
    ]
    headers_eng = [
        'Type of call',
        'No.',
        'Date',
        'Time',
        'Account number',
        'Bank',
        'Name of customer',
        'Phone number',
        'Note',
    ]
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    # content in sheet
    month_mapper = {
        '01':'Jan',
        '02':'Feb',
        '03':'Mar',
        '04':'Apr',
        '05':'May',
        '06':'Jun',
        '07':'Jul',
        '08':'Aug',
        '09':'Sep',
        '10':'Oct',
        '11':'Nov',
        '12':'Dec',
    }
    sheet_title_vn = f'DANH SÁCH CUỘC GỌI ĐẾN KHÁCH HÀNG_THÁNG {period} ({start_day}/{start_month}/{start_year} - {end_day}/{end_month}/{end_year})'
    sheet_title_eng = f'LIST OF PHONE CALLS TO CUSTOMERS__{month_mapper[end_month]} of {end_year} ({start_day}/{start_month}/{start_year} - {end_day}/{end_month}/{end_year})'
    location_vn = 'Địa điểm: Phòng TTBT- HỘI SỞ'
    location_eng = 'Location: TTBT Department - HỘI SỞ'
    content_type_of_call_vn = 'Xác nhận rút nộp ngân hàng liên kết'
    content_type_of_call_eng = 'Confirmation withdrawal-payment of affiliate bank'

    # Set Column Width and Row Height
    worksheet.set_column('A:A',35)
    worksheet.set_column('B:B',6)
    worksheet.set_column('C:C',15)
    worksheet.set_column('D:D',15)
    worksheet.set_column('E:E',19)
    worksheet.set_column('F:F',20)
    worksheet.set_column('G:G',37)
    worksheet.set_column('H:H',45)
    worksheet.set_column('I:I',12)
    worksheet.set_column('J:J',16)
    worksheet.set_column('K:K',28)
    worksheet.set_column('L:L',36)
    worksheet.set_column('M:M',24)

    # Write value in each column, row
    worksheet.merge_range('A1:I1',sheet_title_vn,sheet_title_vn_format)
    worksheet.merge_range('A2:I2',sheet_title_eng,sheet_title_eng_format)
    worksheet.merge_range('A3:I3',location_vn,dia_diem_vn_format)
    worksheet.merge_range('A4:I4',location_eng,dia_diem_eng_format)

    worksheet.write('A8',content_type_of_call_vn,call_type_vn_left_format)
    worksheet.merge_range(f'A9:A{table.shape[0]+7}',content_type_of_call_eng,call_type_eng_left_format)
    worksheet.write_row('A6',headers_vn,header_vn_format)
    worksheet.write_row('A7',headers_eng,header_eng_format)

    worksheet.write_column(
        'B8',
        np.arange(table.shape[0])+1,
        text_center_format
    )
    worksheet.write_column(
        'C8',
        table['date'],
        date_format
    )
    worksheet.write_column(
        'D8',
        table['time'],
        time_format
    )
    worksheet.write_column(
        'E8',
        table['account_code'],
        text_center_format
    )
    worksheet.write_column(
        'F8',
        table['bank'],
        text_center_format
    )
    worksheet.write_column(
        'G8',
        table['customer_name'],
        customer_format
    )
    worksheet.write_column(
        'H8',
        ['(+84 28) 54135479 - Ext: 8168, 8169, 8145']*table.shape[0],
        text_center_format
    )
    worksheet.write_column(
        'I8',
        ['']*table.shape[0],
        text_right_format
    )

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
