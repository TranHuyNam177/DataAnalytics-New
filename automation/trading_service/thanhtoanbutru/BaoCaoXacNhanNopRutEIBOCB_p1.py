"""
Duy trì 2 file:
    1. Xuất tại thời điểm ngày đang chạy (p1)
    2. Xuất cho tháng trước (p2)
"""

from automation.trading_service.thanhtoanbutru import *


# DONE
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    t0_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###########################################################################
    ###########################################################################
    ###########################################################################

    nop_rut_table = pd.read_sql(
        f"""
        SELECT
            [money_in_out_transfer].[date],
            [account].[customer_name],
            [money_in_out_transfer].[bank],
            [money_in_out_transfer].[bank_account],
            [relationship].[account_code],
            [money_in_out_transfer].[transaction_id],
            [money_in_out_transfer].[amount],
            [money_in_out_transfer].[status]
        FROM
            [money_in_out_transfer]
        LEFT JOIN [relationship] 
            ON (
            [money_in_out_transfer].[sub_account] = [relationship].[sub_account] 
            AND [money_in_out_transfer].[date] = [relationship].[date]
            )
        LEFT JOIN [account]
            ON [relationship].[account_code] = [account].[account_code]
        WHERE 
            ([money_in_out_transfer].[bank] IN ('EIB','OCB') OR [money_in_out_transfer].[bank] IS NULL) 
            -- cac TK ngat lien ket NH se bi NULL tren he thong
        AND
            [money_in_out_transfer].[transaction_id] IN ('6692','6693')
        AND 
            [money_in_out_transfer].[date] = '{t0_date}'
        ORDER BY 
        [money_in_out_transfer].[bank]
        """,
        connect_DWH_CoSo
    )

    # Tìm ngân hàng liên kết cuối cùng trước khi ngắt liên kết
    missing_bank_mask = nop_rut_table['bank'].isna()
    nop_rut_table.loc[missing_bank_mask,'bank'] = nop_rut_table.loc[missing_bank_mask,'account_code'].map(get_bank_name)

    # chia theo 2 bank
    nop_rut_eib = nop_rut_table.loc[nop_rut_table['bank']=='EIB'].copy()
    nop_rut_eib['NopTien'] = nop_rut_eib.loc[nop_rut_eib['transaction_id']=='6692','amount']
    nop_rut_eib['NopTien'].fillna('',inplace=True)
    nop_rut_eib['RutTien'] = nop_rut_eib.loc[nop_rut_eib['transaction_id']=='6693','amount']
    nop_rut_eib['RutTien'].fillna('',inplace=True)

    nop_rut_ocb = nop_rut_table.loc[nop_rut_table['bank']=='OCB'].copy()
    nop_rut_ocb['NopTien'] = nop_rut_ocb.loc[nop_rut_ocb['transaction_id']=='6692','amount']
    nop_rut_ocb['NopTien'].fillna('',inplace=True)
    nop_rut_ocb['RutTien'] = nop_rut_ocb.loc[nop_rut_ocb['transaction_id']=='6693','amount']
    nop_rut_ocb['RutTien'].fillna('',inplace=True)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # EXIMBANK
    eod = dt.datetime.strptime(t0_date,"%Y/%m/%d").strftime("%d.%m.%Y")
    file_name = f'Báo cáo Xác nhận nhập rút EIB {eod}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    larger_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vbottom',
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    file_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vbottom',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':14
        }
    )
    sub_title_format = workbook.add_format(
        {
            'valign':'vbottom',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'num_format':'dd/mm/yyyy',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    amount_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':9,
            'num_format':'#,##0'
        }
    )
    xacnhancuanganhang_format = workbook.add_format(
        {
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    xacnhancuaphuhung_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    nguoilap_truongdonvi_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    year = t0_date[:4]
    month = convert_int(t0_date[5:7])
    day = convert_int(t0_date[-2:])
    worksheet = workbook.add_worksheet(f'{day}.{month}.{year}')
    worksheet.hide_gridlines(option=2)

    # Set Column Width and Row Height
    worksheet.set_column('A:A',11)
    worksheet.set_column('B:B',0)
    worksheet.set_column('C:C',26)
    worksheet.set_column('D:D',18)
    worksheet.set_column('E:E',14)
    worksheet.set_column('F:F',0)
    worksheet.set_column('G:G',13)
    worksheet.set_column('H:H',14)
    worksheet.set_column('I:I',12)
    worksheet.set_column('J:J',13)

    worksheet.set_default_row(26)
    worksheet.set_row(0,16)
    worksheet.set_row(1,16)
    worksheet.set_row(2,15)
    worksheet.set_row(3,15)
    worksheet.set_row(4,15)
    worksheet.set_row(5,19)

    headers = [
        'Ngày',
        'Mã GD',
        'HỌ VÀ TÊN',
        'TK NGÂN HÀNG',
        'TK CHỨNG KHOÁN',
        'SỐ TIỀN',
        'NỘP TIỀN',
        'RÚT TIỀN',
        'NH XN',
        'GHI CHÚ',
    ]
    worksheet.write_row('A8',headers,header_format)

    footer_start_row = nop_rut_eib.shape[0]+10
    worksheet.merge_range('A1:J1','CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',larger_title_format)
    worksheet.merge_range('A2:J2','Độc lập – Tự do – Hạnh phúc',larger_title_format)
    worksheet.merge_range('A6:J6','GIẤY XÁC NHẬN NHẬP NỘP RÚT TIỀN',file_title_format)
    worksheet.merge_range('H4:J4',f'TP.HCM, Ngày ... Tháng {month} Năm {year}',sub_title_format)

    worksheet.write_column('A9',nop_rut_eib['date'],date_format)
    worksheet.write_column('B9',['']*nop_rut_eib.shape[0],text_center_format)
    worksheet.write_column('C9',nop_rut_eib['customer_name'],text_left_format)
    worksheet.write_column('D9',nop_rut_eib['bank_account'],text_center_format)
    worksheet.write_column('E9',nop_rut_eib['account_code'],text_center_format)
    worksheet.write_column('F9',['']*nop_rut_eib.shape[0],text_center_format)
    worksheet.write_column('G9',nop_rut_eib['NopTien'],amount_format)
    worksheet.write_column('H9',nop_rut_eib['RutTien'],amount_format)
    worksheet.write_column('I9',['ĐỒNG Ý']*nop_rut_eib.shape[0],text_center_format)
    worksheet.write_column('J9',['']*nop_rut_eib.shape[0],text_center_format)

    worksheet.merge_range(
        f'A{footer_start_row}:D{footer_start_row}',
        '                  XÁC NHẬN CỦA EXIM CHI NHÁNH SÀI GÒN',xacnhancuanganhang_format
    )
    worksheet.merge_range(f'A{footer_start_row+2}:C{footer_start_row+2}','Người lập',nguoilap_truongdonvi_format)
    worksheet.write(f'D{footer_start_row+2}','Người duyệt',nguoilap_truongdonvi_format)
    worksheet.merge_range(f'A{footer_start_row+6}:C{footer_start_row+6}','.....',nguoilap_truongdonvi_format)
    worksheet.write(f'D{footer_start_row+6}','.....',nguoilap_truongdonvi_format)
    worksheet.merge_range(
        f'G{footer_start_row}:J{footer_start_row}','XÁC NHẬN CỦA CHỨNG KHOÁN PHÚ HƯNG',
        xacnhancuaphuhung_format
    )
    worksheet.merge_range(f'G{footer_start_row+2}:H{footer_start_row+2}','Người lập',nguoilap_truongdonvi_format)
    worksheet.merge_range(f'I{footer_start_row+2}:J{footer_start_row+2}','Người duyệt',nguoilap_truongdonvi_format)
    worksheet.merge_range(f'G{footer_start_row+6}:H{footer_start_row+6}','.....',nguoilap_truongdonvi_format)
    worksheet.merge_range(f'I{footer_start_row+6}:J{footer_start_row+6}','.....',nguoilap_truongdonvi_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # OCEANBANK
    file_name = f'Báo cáo Xác nhận nhập rút OCB {eod}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    larger_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vtop',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Cambria',
            'font_size':15
        }
    )
    sub_title_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vbottom',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'num_format':'dd/mm/yyyy',
            'font_name':'Times New Roman',
            'font_size':9
        }
    )
    amount_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':9,
            'num_format':'#,##0'
        }
    )
    xacnhancuanganhang_format = workbook.add_format(
        {
            'valign':'vcenter',
            'font_name':'Cambria',
            'font_size':12
        }
    )
    xacnhancuaphuhung_format = workbook.add_format(
        {
            'valign':'vcenter',
            'font_name':'Cambria',
            'font_size':12
        }
    )
    nguoilap_truongdonvi_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_name':'Cambria',
            'font_size':12
        }
    )
    # Column headers
    headers = [
        'Ngày',
        'Mã GD',
        'HỌ VÀ TÊN',
        'TK NH',
        'TK CK',
        'SỐ TIỀN',
        'NỘP TIỀN',
        'RÚT TIỀN',
        'XN BÊN NH',
        'GHI CHÚ',
    ]
    worksheet = workbook.add_worksheet(f'Tháng {month}')
    worksheet.hide_gridlines(option=2)

    # Set Column Width and Row Height
    worksheet.set_column('A:A',12)
    worksheet.set_column('B:B',0)
    worksheet.set_column('C:C',25)
    worksheet.set_column('D:D',18)
    worksheet.set_column('E:E',13)
    worksheet.set_column('F:F',0)
    worksheet.set_column('G:H',14)
    worksheet.set_column('I:J',12)

    worksheet.set_default_row(23)
    worksheet.set_row(0,15)
    worksheet.set_row(1,16)
    worksheet.set_row(2,15)
    worksheet.set_row(3,15)
    worksheet.set_row(4,24)
    worksheet.set_row(5,20)

    last_row_of_df = nop_rut_ocb.shape[0]
    worksheet.set_row(last_row_of_df+8,16)
    worksheet.set_row(last_row_of_df+9,16)
    worksheet.set_row(last_row_of_df+10,16)
    worksheet.set_row(last_row_of_df+15,16)

    # Insert image OCB
    worksheet.insert_image('A1',join(dirname(__file__),'img','ocb_logo.png'))

    worksheet.write_row('A8',headers,header_format)
    footer_start_row = nop_rut_ocb.shape[0]+10
    worksheet.merge_range('G1:J1','CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',larger_title_format)
    worksheet.merge_range('G2:J2','Độc lập – Tự do – Hạnh phúc',larger_title_format)
    worksheet.merge_range(
        'A5:J6',
        'GIẤY XÁC NHẬN ĐỒNG Ý CHO CÁC KHÁCH HÀNG RÚT TIỀN\nCHUYỂN KHOẢN HOẶC NỘP TIỀN',
        title_format
    )
    worksheet.merge_range('H4:J4',f'TP.HCM, Ngày ... Tháng {month} Năm {year}',sub_title_format)

    worksheet.write_column('A9',nop_rut_ocb['date'],date_format)
    worksheet.write_column('B9',nop_rut_ocb['transaction_id'],text_center_format)
    worksheet.write_column('C9',nop_rut_ocb['customer_name'],text_left_format)
    worksheet.write_column('D9',nop_rut_ocb['bank_account'],text_center_format)
    worksheet.write_column('E9',nop_rut_ocb['account_code'],text_center_format)
    worksheet.write_column('F9',nop_rut_ocb['amount'],text_center_format)
    worksheet.write_column('G9',nop_rut_ocb['NopTien'],amount_format)
    worksheet.write_column('H9',nop_rut_ocb['RutTien'],amount_format)
    worksheet.write_column('I9',['ĐỒNG Ý']*nop_rut_ocb.shape[0],text_center_format)
    worksheet.write_column('J9',['']*nop_rut_ocb.shape[0],text_center_format)

    worksheet.merge_range(
        f'A{footer_start_row}:D{footer_start_row}',
        '     XÁC NHẬN CỦA OCB - PHẠM VĂN HAI ',
        xacnhancuanganhang_format
    )
    worksheet.merge_range(f'C{footer_start_row+1}:D{footer_start_row+1}','NGƯỜI DUYỆT',nguoilap_truongdonvi_format)
    worksheet.merge_range(f'G{footer_start_row+1}:H{footer_start_row+1}','NGƯỜI LẬP',nguoilap_truongdonvi_format)
    worksheet.merge_range(
        f'G{footer_start_row}:J{footer_start_row}',
        '                     XÁC NHẬN CỦA CTY CK PHÚ HƯNG',xacnhancuaphuhung_format
    )
    worksheet.merge_range(f'I{footer_start_row+1}:J{footer_start_row+1}','NGƯỜI DUYỆT',nguoilap_truongdonvi_format)
    worksheet.write(f'A{footer_start_row+1}','NGƯỜI LẬP',nguoilap_truongdonvi_format)
    worksheet.merge_range(f'G{footer_start_row+6}:H{footer_start_row+6}','.....',nguoilap_truongdonvi_format)
    worksheet.merge_range(f'I{footer_start_row+6}:J{footer_start_row+6}','.....',nguoilap_truongdonvi_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
