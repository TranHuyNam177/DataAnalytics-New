"""
BC này không thể chạy lùi ngày
(tuy nhiên có thể viết lại để chạy lùi ngày do đã bắt đầu lưu VCF0051)
"""

from automation.trading_service.thanhtoanbutru import *


# DONE
def run(
    run_time=None
):

    start = time.time()
    info = get_info('daily',run_time)
    t0_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir(join(dept_folder,folder_name,period))

    ###################################################
    ###################################################
    ###################################################

    vip_phs = pd.read_sql(
        f"""
        WITH 
        [contract_code] AS (
            SELECT
                [account_type].[type],
                CASE
                    WHEN [account_type].[type_name] LIKE N'%NOR%' THEN 'NORM'
                    WHEN [account_type].[type_name] LIKE N'%SILV%' THEN 'SILV'
                    WHEN [account_type].[type_name] LIKE N'%GOLD%' THEN 'GOLD'
                    WHEN [account_type].[type_name] LIKE N'%VIPCN%' THEN 'VIP'
                    ELSE N'Tiểu khoản thường'
                END [contract_type],
                [account_type].[type_name]
            FROM
                [account_type]
        )
        , [eod_change] AS (
            SELECT 
                [t].[account_code],
                [t].[date_of_change],
                [t].[contract_type],
                [t].[date_of_approval]
            FROM (
                SELECT 
                    [customer_information_change].[account_code],
                    [customer_information_change].[date_of_change],
                    [customer_information_change].[time_of_change],
                    [customer_information_change].[date_of_approval],
                    MAX(
                        CASE 
                            WHEN [customer_information_change].[change_content] = 'Loai hinh hop dong'
                                AND [contract_code].[contract_type] <> N'Tiểu khoản thường'
                            THEN [customer_information_change].[time_of_change] 
                        END
                    ) OVER (PARTITION BY [account_code], [date_of_change]) [eof_time_of_change],
                    [contract_code].[type],
                    [contract_code].[contract_type]
                FROM
                    [customer_information_change]
                LEFT JOIN [contract_code] ON [customer_information_change].[after] = [contract_code].[type]
                WHERE [customer_information_change].[date_of_change] <= '{t0_date}'
                    AND [customer_information_change].[change_content] = 'Loai hinh hop dong'
                    AND [contract_code].[contract_type] <> N'Tiểu khoản thường'
            ) [t]
            WHERE [t].[time_of_change] = [t].[eof_time_of_change]
        )
        , [full] AS (
            SELECT 
                [t].*
            FROM (
                SELECT 
                    [eod_change].[account_code],
                    [eod_change].[date_of_change],
                    [eod_change].[date_of_approval],
                    [eod_change].[contract_type],
                    MAX(CASE WHEN [eod_change].[contract_type] = 'NORM' THEN [eod_change].[date_of_change] END)
                        OVER (PARTITION BY [eod_change].[account_code]) [norm_end], -- ngày cuối cùng là normal
                    MAX([eod_change].[date_of_change])
                        OVER (PARTITION BY [eod_change].[account_code]) [last_change_date] -- ngày thay đổi cuối cùng
                FROM [eod_change]
            ) [t]
            WHERE ([t].[date_of_change] > [t].[norm_end] OR [t].[norm_end] IS NULL) -- Bỏ qua các thay đổi trước ngày cuối cùng là normal
        )
        , [processing] AS (
            SELECT 
                [full].*,
                MIN([date_of_change]) OVER (PARTITION BY [full].[account_code]) [effective_date], -- Ngày hiệu lực đầu tiên
                MAX(CASE WHEN [full].[date_of_change] = [full].[last_change_date] THEN [full].[contract_type] END)
                    OVER (PARTITION BY [full].[account_code]) [last_type], -- loại hình hợp đồng tại thời điểm gần nhất
                LAG([full].[contract_type]) OVER (PARTITION BY [full].[account_code] ORDER BY [full].[date_of_change]) [previous_type] -- loại hình hợp đồng trước đó
            FROM 
                [full]
        )
        , [final] AS (
        SELECT DISTINCT
            [processing].[account_code],
            [processing].[effective_date],
            MAX(CASE WHEN [processing].[contract_type] = [processing].[last_type] THEN [processing].[date_of_approval] END)
                OVER (PARTITION BY [processing].[account_code]) [approve_date] -- Tính approve date dựa vào ngày thay đổi đầu tiên trong cùng một loại hình
        FROM 
            [processing]
        WHERE ([processing].[contract_type] <> [processing].[previous_type] OR [processing].[previous_type] IS NULL) -- Chỉ lấy ngày thay đổi đầu tiên trong cùng một loại hình
        )
        
        SELECT DISTINCT -- do một số KH thay đổi 2 tiểu khoản cùng lúc. VD: '022C002768' ngày 29/10/2021
            CONCAT(N'Tháng ',FORMAT(MONTH([account].[date_of_birth]),'00')) [birth_month],
            [account].[account_code], 
            [account].[customer_name], 
            [branch].[branch_name], 
            [account].[date_of_birth], 
            CASE 
                WHEN [account].[gender] = '001' THEN 'Male'
                WHEN [account].[gender] = '002' THEN 'Female'
                ELSE ''
            END [gender],
            CASE 
                WHEN [vcf0051].[contract_type] LIKE N'%GOLD%' THEN 'GOLD PHS'
                WHEN [vcf0051].[contract_type] LIKE N'%SILV%' THEN 'SILV PHS'
                WHEN [vcf0051].[contract_type] LIKE N'%VIP%' THEN 'VIP Branch'
            END [contract_type],
            CASE
                WHEN [account].[account_code] = '022C108142' THEN DATETIMEFROMPARTS(2017,11,28,0,0,0,0) -- Dữ liệu xảy ra trước ngày DWH là 2018-01-01
                WHEN [account].[account_code] = '022C103838' THEN DATETIMEFROMPARTS(2017,9,13,0,0,0,0) -- Dữ liệu xảy ra trước ngày DWH là 2018-01-01
                ELSE [final].[effective_date]
            END [effective_date],
            CASE
                WHEN [account].[account_code] = '022C108142' THEN DATETIMEFROMPARTS(2017,11,28,0,0,0,0) -- Dữ liệu xảy ra trước ngày DWH là 2018-01-01
                WHEN [account].[account_code] = '022C103838' THEN DATETIMEFROMPARTS(2017,9,13,0,0,0,0) -- Dữ liệu xảy ra trước ngày DWH là 2018-01-01
                ELSE [final].[approve_date]
            END [approve_date],
            CASE
                WHEN [vcf0051].[contract_type] LIKE N'%GOLD%' OR [vcf0051].[contract_type] LIKE N'%SILV%' THEN 
                    CASE
                        WHEN DATEPART(MONTH,'{t0_date}') < 6 THEN DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}'),6,30,0,0,0,0)
                        WHEN DATEPART(MONTH,'{t0_date}') < 12 THEN DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}'),12,31,0,0,0,0)
                        ELSE DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}')+1,6,30,0,0,0,0) 
                    END
                WHEN [vcf0051].[contract_type] LIKE N'%VIP%' THEN
                    CASE
                        WHEN DATEPART(MONTH,'{t0_date}') < 3 THEN DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}'),3,31,0,0,0,0)
                        WHEN DATEPART(MONTH,'{t0_date}') < 6 THEN DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}'),6,30,0,0,0,0)
                        WHEN DATEPART(MONTH,'{t0_date}') < 9 THEN DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}'),9,30,0,0,0,0)
                        WHEN DATEPART(MONTH,'{t0_date}') < 12 THEN DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}'),12,31,0,0,0,0)
                        ELSE  DATETIMEFROMPARTS(DATEPART(YEAR,'{t0_date}')+1,3,31,0,0,0,0) 
                    END
            END [review_date],
            [broker].[broker_id],
            [broker].[broker_name]
        FROM [vcf0051] 
        LEFT JOIN [relationship]
            ON [vcf0051].[sub_account] = [relationship].[sub_account]
            AND [vcf0051].[date] = [relationship].[date]
        LEFT JOIN [account]
            ON [relationship].[account_code] = [account].[account_code]
        LEFT JOIN [broker]
            ON [relationship].[broker_id] = [broker].[broker_id]
        LEFT JOIN [branch]
            ON [relationship].[branch_id] = [branch].[branch_id]
        LEFT JOIN [final]
            ON [final].[account_code] = [relationship].[account_code]
        WHERE (
            [vcf0051].[contract_type] LIKE N'%GOLD%' 
            OR [vcf0051].[contract_type] LIKE N'%SILV%' 
            OR [vcf0051].[contract_type] LIKE N'%VIP%'
        )
        AND [vcf0051].[status] IN ('A','B')
        AND [vcf0051].[date] = '{t0_date}'
        ORDER BY [birth_month], [date_of_birth]

        """,
        connect_DWH_CoSo,
    )

    branch_groupby_table = vip_phs.groupby(['branch_name','contract_type'])['customer_name'].count().unstack()
    branch_groupby_table.fillna(0,inplace=True)

    # --------------------- Viet File ---------------------

    eod = dt.datetime.strptime(t0_date,"%Y/%m/%d").strftime("%d.%m.%Y")
    file_name = f'Danh sách KH VIP cơ sở {eod}.xlsx'
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
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':18
        }
    )
    sub_title_format = workbook.add_format(
        {
            'bold':True,
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'color':'#FF0000',
            'bg_color':'#FFFFCC'
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
            'font_size':11,
            'bg_color':'#65FF65'
        }
    )
    no_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11,
        }
    )
    header_format_tong_theo_cn = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11,
            'bg_color':'#92D050'
        }
    )
    list_customer_vip_row_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'color':'#FF0000',
            'bg_color':'#FFFF00',
        }
    )
    ngay_hieu_luc_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'color':'#FF0000',
        }
    )
    list_customer_vip_col_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11,
            'bg_color':'#FFCCFF',
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'num_format':'dd/mm/yyyy',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'valign':'vcenter',
            'align':'left',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'valign':'vcenter',
            'align':'center',
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    num_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11,
            'num_format':'#,##0',
        }
    )
    num_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11,
            'num_format':'#,##0'
        }
    )
    headers = [
        'No',
        'Tháng sinh',
        'Account',
        'Name',
        'Location',
        'Birthday',
        'Gender',
        'LIST CUSTOMER VIP',
        'Ngày hiệu lực đầu tiên',
        'Approve date',
        'Review date',
        'Mã MG quản lý tài khoản',
        'Tên MG quản lý tài khoản',
        'Note',
    ]
    headers_tong_theo_cn = [
        'Location',
        'GOLD PHS',
        'SILV PHS',
        'VIP BRANCH',
        'SUM',
        'NOTE'
    ]

    #  Viết Sheet VIP PHS
    vip_phs_sheet = workbook.add_worksheet('VIP PHS')
    vip_phs_sheet.hide_gridlines(option=2)

    # Content of sheet
    sheet_title_name = f'UPDATED LIST OF COMPANY VIP TO {eod}'
    sub_title_name = 'Chỉ tặng quà sinh nhật cho Gold PHS & Silv PHS'

    # Set Column Width and Row Height
    vip_phs_sheet.set_column('A:A',4)
    vip_phs_sheet.set_column('B:C',13)
    vip_phs_sheet.set_column('D:D',40)
    vip_phs_sheet.set_column('E:E',26)
    vip_phs_sheet.set_column('F:F',11)
    vip_phs_sheet.set_column('G:G',10)
    vip_phs_sheet.set_column('H:H',13)
    vip_phs_sheet.set_column('I:I',11)
    vip_phs_sheet.set_column('J:J',11)
    vip_phs_sheet.set_column('K:K',12)
    vip_phs_sheet.set_column('L:L',11)
    vip_phs_sheet.set_column('M:M',40)
    vip_phs_sheet.set_column('N:N',20)
    vip_phs_sheet.set_row(0,23)
    vip_phs_sheet.set_row(1,27)
    vip_phs_sheet.set_row(2,12)
    vip_phs_sheet.set_row(3,50)

    # merge row
    vip_phs_sheet.merge_range('A1:N1',sheet_title_name,sheet_title_format)
    vip_phs_sheet.merge_range('A2:N2',sub_title_name,sub_title_format)

    # write row and write column
    vip_phs_sheet.write_row('A4',headers,header_format)
    vip_phs_sheet.write('A4','No',no_format)
    vip_phs_sheet.write('H4','LIST CUSTOMER VIP',list_customer_vip_row_format)
    vip_phs_sheet.write('I4','Ngày hiệu lực đầu tiên',ngay_hieu_luc_format)
    vip_phs_sheet.write_column('A5',np.arange(vip_phs.shape[0])+1,num_center_format)
    vip_phs_sheet.write_column('B5',vip_phs['birth_month'],text_center_format)
    vip_phs_sheet.write_column('C5',vip_phs['account_code'],num_left_format)
    vip_phs_sheet.write_column('D5',vip_phs['customer_name'],text_left_format)
    vip_phs_sheet.write_column('E5',vip_phs['branch_name'],text_left_format)
    vip_phs_sheet.write_column('F5',vip_phs['date_of_birth'],date_format)
    vip_phs_sheet.write_column('G5',vip_phs['gender'],text_center_format)
    vip_phs_sheet.write_column('H5',vip_phs['contract_type'],list_customer_vip_col_format)
    vip_phs_sheet.write_column('I5',vip_phs['effective_date'],date_format)
    vip_phs_sheet.write_column('J5',vip_phs['approve_date'],date_format)
    vip_phs_sheet.write_column('K5',vip_phs['review_date'],date_format)
    vip_phs_sheet.write_column('L5',vip_phs['broker_id'],text_center_format)
    vip_phs_sheet.write_column('M5',vip_phs['broker_name'],text_left_format)
    vip_phs_sheet.write_column('N5',['']*vip_phs.shape[0],text_left_format)

    ##############################################
    ##############################################
    ##############################################

    # Viết Sheet TONG THEO CN
    sum_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'valign':'vcenter',
            'align':'right',
            'font_size':11,
            'font_name':'Calibri',
            'num_format':'#,##0'
        }
    )
    sum_title_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'valign':'vcenter',
            'align':'left',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    location_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'align':'left',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    num_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'align':'right',
            'font_size':11,
            'font_name':'Calibri',
            'num_format':'#,##0'
        }
    )
    note_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'align':'left',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    branch_groupby_sheet = workbook.add_worksheet('TONG THEO CN')
    branch_groupby_sheet.hide_gridlines(option=2)
    # Set Column Width and Row Height
    branch_groupby_sheet.set_column('A:A',28)
    branch_groupby_sheet.set_column('B:B',14)
    branch_groupby_sheet.set_column('C:C',15)
    branch_groupby_sheet.set_column('D:F',18)

    branch_groupby_sheet.write_row('A1',headers_tong_theo_cn,header_format_tong_theo_cn)
    branch_groupby_sheet.write_column('A2',branch_groupby_table.index,location_format)
    branch_groupby_sheet.write_column('B2',branch_groupby_table['GOLD PHS'],num_format)
    branch_groupby_sheet.write_column('C2',branch_groupby_table['SILV PHS'],num_format)
    branch_groupby_sheet.write_column('D2',branch_groupby_table['VIP Branch'],num_format)
    sum_row = branch_groupby_table['GOLD PHS']+branch_groupby_table['SILV PHS']+branch_groupby_table['VIP Branch']
    branch_groupby_sheet.write_column('E2',sum_row,sum_format)
    note_col = ['']*branch_groupby_table.shape[0]
    branch_groupby_sheet.write_column('F2',note_col,note_format,)
    start_sum_row = branch_groupby_table.shape[0]+2
    branch_groupby_sheet.write(f'A{start_sum_row}','SUM',sum_title_format)
    for col in 'BCDE':
        branch_groupby_sheet.write(f'{col}{start_sum_row}',f'=SUBTOTAL(9,{col}2:{col}{start_sum_row-1})',sum_format)
    branch_groupby_sheet.write(f'F{start_sum_row}','',sum_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
