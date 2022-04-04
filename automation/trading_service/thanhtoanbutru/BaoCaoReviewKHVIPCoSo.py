"""
    1. VIP CN => mỗi quý review 1 lần => lấy data quý đó
    NORMAL, GOLD và SILV => 6 tháng review 1 lần => lấy data 6 tháng
    =>  Báo cáo chạy vào cuối quý 2, quý 4 => xét toàn bộ KH (normal,gold,silv,vip)
        Báo cáo chạy vào cuối quý 1, quý 3 => chỉ xét vip

    2. Nếu từ được lên VIP (approved date) cho tới ngày cuối chu kỳ > 30 ngày => giữ (có review)
    còn KH nào thời gian lên VIP tới ngày cuối chu kỳ <= 30 => bỏ (không review)

    3. Qui tắc xét ngày lấy giá trị:
    - Nếu ngày KH lên VIP vào trước chu kỳ => xét giá trị từ đầu chu kỳ tới cuối chu kỳ
    - Nếu ngày KH lên VIP nằm trong chu kỳ => xét giá trị từ ngày lên VIP tới cuối chu kỳ

    4. Fee for assessment
    công thức: Fee for assessment = 100% * Phí GD + 30% * (Phí UTTB + lãi vay)
    chú ý: tính tổng phí trung bình tháng và tổng lãi trung bình tháng
           trước khi thay vào công thức trên
    - phí GD: ROD0040
    - phí UTTB: RCI0015
    = lãi vay: RLN0019

    5. Cột I Average Net Asset Value -> lấy từ bảng nav

    6. cột H - % Fee for assessment / Criteria Fee

    7. cột Rate lấy từ số % nằm trong 'contract_type' nằm sau cụm từ Margin.PIA ...

    8. Rule cột After review
    KH GOLD:
        if: KH có Fee for Aseessment >= 40tr -> GOLD
        (nếu khác GOLD thì tăng bậc là PROMOTE GOLD, còn nếu là GOLD thì giữ nguyên GOLD)
        else:
            if: KH GOLD có cột H >= 80%, I >= 4 tỷ -> giữ là GOLD
            else: DEMOTE SILV
    KH SILV:
        if: KH có 20tr <= Fee for Aseessment < 40tr -> SILV
        (nếu khác SILV tăng bậc là PROMOTE SILV, còn nếu là SILV thì giữ nguyên SILV)
        else:
            if: KH SILV có cột H >= 80%, I >= 2 tỷ -> giữ là SILV
            else: DEMOTE DP
    KH VIP CN:
        if: fee < 20tr OR H < 80%, I < 4 tỷ -> DEMOTE DP
        else:
            if: 20tr <= fee < 40tr -> PROMOTE SILV
            else: fee >= 40tr -> PROMOTE GOLD
    KH Normal:
        if: 20tr <= fee < 40tr -> PROMOTE SILV
        else: fee >= 40tr -> PROMOTE GOLD

    9. Các cột M, N, O, P, R để trống

"""

from automation.trading_service.thanhtoanbutru import *


# DONE
def review(level,start_date,end_date):

    valid_params = ['normal','silv','gold','vipcn']
    if level not in valid_params:
        raise ValueError(f"'level' must either {valid_params}")

    saved_sql = join(dirname(__file__),'saved_sql',f'{level}.sql')
    with open(saved_sql,'r',encoding='utf-8') as file:
        query = file.read().replace('__input__[0]',f'{start_date}').replace('__input__[1]',f'{end_date}')

    table = pd.read_sql(query,connect_DWH_CoSo)
    table['rate'] = table['rate'].str.split('Margin.PIA').str.get(1).str.split('%').str.get(0).astype(float)/100

    return table


# Có thể chạy lùi ngày, bắt đầu từ 27/12/2021 (ngày bắt đầu lưu VCF0051 trên DWH)
def run(
    run_time=None,
):

    start = time.time()
    info = get_info('quarterly',run_time)
    end_date = info['end_date'].replace('/','-')
    period = info['period']
    folder_name = info['folder_name']

    if run_time is None:
        run_time = dt.datetime.now()

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir(join(dept_folder,folder_name,period))

    ###################################################
    ###################################################
    ###################################################

    year = end_date[:4]

    if end_date[5:7] in ('06','12'):
        tables = []
        if end_date[5:7]=='06':
            s3m = f'{year}-04-01'
            s6m = f'{year}-01-01'
            e6m = f'{year}-06-30'
        else:
            s3m = f'{year}-10-01'
            s6m = f'{year}-07-01'
            e6m = f'{year}-12-31'
        zip_tuple = (
            ['normal','silv','gold','vipcn'],[s6m,s6m,s6m,s3m],[e6m]*4
        )
        for a,b,c in zip(*zip_tuple):
            tables.append(review(a,b,c))
        table = pd.concat(tables)

    elif end_date[5:7] in ('03','09'):
        if end_date[5:7]=='03':
            s3m = f'{year}-01-01'
            e3m = f'{year}-03-31'
        else:
            s3m = f'{year}-07-01'
            e3m = f'{year}-09-30'
        table = review('vipcn',s3m,e3m)

    else:
        raise ValueError('Invalid end_date')

    # --------------------- Viet File ---------------------
    # Write file excel Báo cáo review KH vip
    file_date = f'{end_date[-2:]}.{end_date[5:7]}.{end_date[:4]}'
    file_name = f'Báo cáo review KH VIP {file_date}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    sheet_title_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':20
        }
    )
    sub_title_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True
        }
    )
    sub_title_1_format = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':14,
        }
    )
    sub_title_2_format = workbook.add_format(
        {
            'italic':True,
            'font_name':'Times New Roman',
            'font_size':14,
        }
    )
    no_and_date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12,
        }
    )
    description_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'top',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12,
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
            'font_size':8,
            'bg_color':'#FFC000'
        }
    )
    header_2_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':8,
            'bg_color':'#FFFF00'
        }
    )
    header_3_format = workbook.add_format(
        {
            'border':1,
            'bold': True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':8,
        }
    )
    empty_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
        }
    )
    stt_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':8,
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':8,
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':8,
            'color':'#222B33',
            'bg_color':'#FFCCFF'
        }
    )
    text_left_wrap_text_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':8,
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd/mm/yyyy',
            'font_name':'Times New Roman',
            'font_size':8
        }
    )
    criteria_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':8,
            'num_format':'#,##0',
            'color':'#222B33',
            'bg_color':'#FFCCFF'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':8,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    percent_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':8,
            'num_format':'0.00%'
        }
    )
    footer_format = workbook.add_format(
        {
            'text_wrap':True,
            'align':'left',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11,
        }
    )
    table_header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11,
        }
    )
    table_sub_header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':9,
        }
    )
    table_empty_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'top',
            'font_name':'Times New Roman',
            'font_size':11,
        }
    )
    table_content_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11,
        }
    )
    date_header = dt.datetime.strptime(end_date,"%Y-%m-%d").strftime("%d.%m.%Y")
    headers = ['No.','Account','Name','Branch','Approve day','Criteria Fee',
               'Fee for assessment','% Fee for assessment / Criteria Fee',
               'Average Net Asset Value','Current VIP','After review','Rate',
               f'Group & Deal đợt {date_header}','Opinion of Location Manager',
               'Opinion of Trading Service','Decision of Deputy General Director',
               'Môi giới quản lý','Note']

    #  Viết Sheet VIP PHS
    worksheet = workbook.add_worksheet('REVIEW VIP')
    worksheet.hide_gridlines(option=2)

    # Content of sheet
    description = 'To: Deputy General Director of Phu Hung Securities Corporation' \
                  '\nProposer: Nguyen Thi Tuyet'
    worksheet.merge_range('A2:C3','',empty_format)
    worksheet.insert_image('A2',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.49,'y_scale':0.53})

    # Set Column Width and Row Height
    worksheet.set_column('A:A',4)
    worksheet.set_column('B:B',11)
    worksheet.set_column('C:C',23)
    worksheet.set_column('D:D',12)
    worksheet.set_column('E:I',13)
    worksheet.set_column('J:K',15)
    worksheet.set_column('L:L',9)
    worksheet.set_column('M:P',18)
    worksheet.set_column('Q:Q',23)
    worksheet.set_column('R:R',19)

    worksheet.set_row(1,25.5)
    worksheet.set_row(2,18.75)
    worksheet.set_row(4,36.5)
    worksheet.set_row(5,46.5)

    # merge row
    worksheet.merge_range('D2:M2','SUBMISSION',sheet_title_format)
    report_date = dt.datetime.strptime(end_date,"%Y-%m-%d").strftime("%B, %Y")
    worksheet.merge_range('D3:M3','',empty_format)
    worksheet.write_rich_string(
        'D3',
        sub_title_1_format,'Subject :',
        sub_title_2_format,f' REVIEW VIP (THE END OF {report_date.upper()})',
        sub_title_format,
    )
    worksheet.write('N2','No.:',no_and_date_format)
    worksheet.write('N3','Date:',no_and_date_format)
    worksheet.merge_range('O2:P2',f'     /{year}/TTr-TRS',no_and_date_format)
    worksheet.merge_range('O3:P3',f'{run_time.strftime("%d/%m/%Y")}',no_and_date_format)

    worksheet.merge_range('A5:P5',description,description_format)

    worksheet.write_row('A6',headers,header_format)
    worksheet.write('G6',headers[6],header_2_format)
    worksheet.write('H6',headers[7],header_2_format)
    worksheet.write('I6',headers[8],header_2_format)
    worksheet.write('Q6',headers[-2],header_3_format)
    worksheet.write('R6',headers[-1],header_2_format)

    worksheet.write_column('A7',np.arange(table.shape[0])+1,stt_format)
    worksheet.write_column('B7',table['account_code'],text_center_format)
    worksheet.write_column('C7',table['customer_name'],text_left_wrap_text_format)
    worksheet.write_column('D7',table['branch_name'],text_left_format)
    worksheet.write_column('E7',table['approve_date'],date_format)
    worksheet.write_column('F7',table['criteria_fee'],criteria_format)
    worksheet.write_column('G7',table['fee'],money_format)
    worksheet.write_column('H7',table['pct_fee'],percent_format)
    worksheet.write_column('I7',table['nav'],money_format)
    worksheet.write_column('J7',table['level'],text_center_format)
    worksheet.write_column('K7',table['after_review'],text_left_wrap_text_format)
    worksheet.write_column('L7',table['rate'],percent_format)
    worksheet.write_column('M7',['']*table.shape[0],text_left_format)
    worksheet.write_column('N7',['']*table.shape[0],text_left_format)
    worksheet.write_column('O7',['']*table.shape[0],text_left_format)
    worksheet.write_column('P7',['']*table.shape[0],text_left_format)
    worksheet.write_column('Q7',table['broker_name'],text_left_format)
    worksheet.write_column('R7',['']*table.shape[0],text_left_format)
    start_row = table.shape[0]+9
    footer = f'Effective: Promoted accounts be applied from ......./......./{year} ' \
             f'& another accounts be applied from ........../........../{year}'
    worksheet.merge_range(f'A{start_row}:P{start_row}',footer,footer_format)
    worksheet.merge_range(
        f'B{start_row+2}:P{start_row+2}',
        'TRADING SERVICE DIVISION',
        table_header_format
    )
    worksheet.merge_range(
        f'B{start_row+3}:H{start_row+3}',
        'PROPOSER',
        table_sub_header_format
    )
    worksheet.merge_range(
        f'I{start_row+3}:P{start_row+3}',
        'DIRECTOR OF TRADING SERVICE DIVISION',
        table_sub_header_format
    )
    worksheet.merge_range(f'B{start_row+4}:H{start_row+7}','',table_empty_format)
    worksheet.merge_range(f'I{start_row+4}:P{start_row+7}','',table_empty_format)
    worksheet.merge_range(
        f'B{start_row+9}:H{start_row+9}',
        'Decision of Deputy General Director:',
        table_sub_header_format
    )
    worksheet.merge_range(
        f'I{start_row+9}:P{start_row+9}',
        'DEPUTY GENERAL DIRECTOR',
        table_sub_header_format
    )
    worksheet.merge_range(
        f'B{start_row+10}:H{start_row+14}',
        "o Agree\no Disagree\no Others:................",
        table_content_format
    )
    worksheet.merge_range(f'I{start_row+10}:P{start_row+14}','',table_empty_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
