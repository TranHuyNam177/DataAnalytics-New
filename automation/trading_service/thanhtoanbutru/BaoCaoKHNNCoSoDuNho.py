from automation.trading_service.thanhtoanbutru import *
from datawarehouse import SYNC

# DONE
def run(
    run_time=None
):  # BC này chỉ chạy lùi được từ ngày lưu VRM6631 (chạy vào đầu tháng)

    start = time.time()
    info = get_info('daily',run_time)
    t0_date = info['end_date']
    period = info['period']
    folder_name = 'BaoCaoThang'

    def convert(x):
        month = convert_int(int(x[5:7]))
        if month=='00':
            month = 12
            year = int(x[:4])-1
        else:
            year = int(x[:4])
        return f'{month}.{year}'

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,convert(period))):
        os.mkdir((join(dept_folder,folder_name,convert(period))))

    ###################################################
    ###################################################
    ###################################################

    if run_time is None or run_time.date() == dt.datetime.now().date(): # chạy vào hiện tại
        # Update data:
        SYNC(connect_DWH_CoSo,'spvrm6631',FrDate=t0_date,ToDate=t0_date)

    # Query
    def get_full_list(account_type):
        if account_type=='institutional':
            sql_condition = "[vrm6631].[cash_balance_at_bank] <= 33000 AND [r].[account_type] = N'Tổ chức nước ngoài'"
        elif account_type=='individual':
            sql_condition = "[vrm6631].[cash_balance_at_bank] <= 11000 AND [r].[account_type] = N'Cá nhân nước ngoài'"
        else:
            raise ValueError('Invalid Input')
        return pd.read_sql(
            f"""
            WITH [r] AS (
                SELECT 
                    [sub_account].[sub_account],
                    [sub_account].[account_code],
                    [account].[customer_name],
                    [account].[account_type]
                FROM
                    [sub_account]
                LEFT JOIN
                    [account]
                ON [sub_account].[account_code] = [account].[account_code]
            )
            SELECT 
                MAX([vrm6631].[date]) [date],
                [r].[account_code],
                MAX([r].[customer_name]) [customer_name],
                SUM([vrm6631].[cash_balance_at_bank]) [cash_balance_at_bank],
                '' [email]
            FROM [vrm6631]
            LEFT JOIN [r]
                ON [r].[sub_account] = [vrm6631].[sub_account]
            WHERE {sql_condition}
                AND [vrm6631].[date] = '{t0_date}'
            GROUP BY
                [r].[account_code]
            """,
            connect_DWH_CoSo
        )

    instituional = get_full_list('institutional')
    individual = get_full_list('individual')

    ###################################################
    ###################################################
    ###################################################

    file_name = f'Số dư TK FII {period[-2:]}.{period[5:7]}.{period[:4]} - gửi IT.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,convert(period),file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # ------------- Viết sheet -------------
    headers_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri',
            'bg_color':'#FFFF00',
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri',
            'num_format':'#,##0'
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri',
            'num_format':'dd-mm-yyyy'
        }
    )
    headers = ['Ngày','Số TK','Tên','Số tiền','Email']
    # --------- sheet Cá nhân ---------
    sheet_indvidual = workbook.add_worksheet('Cá nhân')
    sheet_indvidual.hide_gridlines(option=2)
    sheet_indvidual.set_tab_color('#C00000')
    sheet_indvidual.set_column('A:A',17)
    sheet_indvidual.set_column('B:B',13)
    sheet_indvidual.set_column('C:C',30)
    sheet_indvidual.set_column('D:D',38)
    sheet_indvidual.set_column('E:E',14)
    sheet_indvidual.write_row('A1',headers,headers_format)
    sheet_indvidual.write_column('A2',individual['date'],date_format)
    sheet_indvidual.write_column('B2',individual['account_code'],text_center_format)
    sheet_indvidual.write_column('C2',individual['customer_name'],text_left_format)
    sheet_indvidual.write_column('D2',individual['cash_balance_at_bank'],money_format)
    sheet_indvidual.write_column('E2',['']*individual.shape[0],text_left_format)

    # --------- sheet Tổ chức ---------
    sheet_institutional = workbook.add_worksheet('Tổ chức')
    sheet_institutional.hide_gridlines(option=2)
    sheet_institutional.set_tab_color('#FFC000')
    sheet_institutional.set_column('A:A',17)
    sheet_institutional.set_column('B:B',13)
    sheet_institutional.set_column('C:C',30)
    sheet_institutional.set_column('D:D',38)
    sheet_institutional.set_column('E:E',14)
    sheet_institutional.write_row('A1',headers,headers_format)
    sheet_institutional.write_column('A2',instituional['date'],date_format)
    sheet_institutional.write_column('B2',instituional['account_code'],text_center_format)
    sheet_institutional.write_column('C2',instituional['customer_name'],text_left_format)
    sheet_institutional.write_column('D2',instituional['cash_balance_at_bank'],money_format)
    sheet_institutional.write_column('E2',['']*instituional.shape[0],text_left_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
