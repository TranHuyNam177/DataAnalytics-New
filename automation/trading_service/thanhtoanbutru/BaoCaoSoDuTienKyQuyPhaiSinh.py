from automation.trading_service.thanhtoanbutru import *


# DONE
def run(
    run_time=None,
):
    start = time.time()
    info = get_info('monthly',run_time)
    start_date = bdate(info['start_date'],-1)
    end_date = info['end_date'].replace('/','-')
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    def get_cash_balance_at_vsd(t):
        return pd.read_sql(
            f"""    
            WITH [t] AS (
                SELECT 
                    [rdt0121].[date],
                    SUM([rdt0121].[cash_balance_at_vsd]) [balance]
                FROM [rdt0121]
                WHERE [rdt0121].[date] BETWEEN '{bdate(t,-1)}' AND '{t}'
                GROUP BY [rdt0121].[date]
                )
                SELECT [t].[balance] FROM [t]
                WHERE [t].[date] = (SELECT MAX([t].[date]) FROM [t]
                )
            """,
                connect_DWH_PhaiSinh
            ).squeeze()

    o_balance = get_cash_balance_at_vsd(start_date)
    c_balance = get_cash_balance_at_vsd(end_date)

    ###################################################
    ###################################################
    ###################################################

    file_name = f'Số dư tiền ký quỹ phái sinh {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':12,
            'font_name':'Arial',
            'text_wrap':True
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Arial',
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Arial',
        }
    )
    num_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Arial',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',10)
    worksheet.set_column('B:C',17)
    worksheet.merge_range('A1:C1','Tổng số dư ký quỹ',title_format)
    worksheet.write_row('B3',['Đầu tháng','Cuối tháng'],header_format)
    worksheet.write('A4','Tiền mặt',text_left_format)
    worksheet.write('B4',o_balance,num_format)
    worksheet.write('C4',c_balance,num_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
