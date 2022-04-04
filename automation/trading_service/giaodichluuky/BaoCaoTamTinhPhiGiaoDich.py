from automation.trading_service.giaodichluuky import *


def run(
    run_time=None
):
    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    result_hose = pd.read_sql(
        f"""
        SELECT
            ROW_NUMBER() OVER (ORDER BY [branch_id]) [STT],
            [rod0040].[exchange] [Sàn],
            [rod0040].[branch_id] [Mã Chi Nhánh],
            CASE
                WHEN [rod0040].[branch_id] = '0001' THEN N'HQ'
                WHEN [rod0040].[branch_id] = '0101' THEN N'Quận 3'
                WHEN [rod0040].[branch_id] = '0102' THEN N'PMH'
                WHEN [rod0040].[branch_id] = '0104' THEN N'Q7'
                WHEN [rod0040].[branch_id] = '0105' THEN N'TB'
                WHEN [rod0040].[branch_id] = '0116' THEN N'P.QLTK1'
                WHEN [rod0040].[branch_id] = '0111' THEN N'InB1'
                WHEN [rod0040].[branch_id] = '0113' THEN N'IB'
                WHEN [rod0040].[branch_id] = '0201' THEN N'Hà Nội'
                WHEN [rod0040].[branch_id] = '0202' THEN N'TX'
                WHEN [rod0040].[branch_id] = '0301' THEN N'Hải Phòng'
                WHEN [rod0040].[branch_id] = '0117' THEN N'Quận 1'
                WHEN [rod0040].[branch_id] = '0118' THEN N'P.QLTK3'
                WHEN [rod0040].[branch_id] = '0119' THEN N'InB2'
            ELSE [rod0040].[branch_id]
            END [Tên Chi Nhánh],
            [rod0040].[volume] [Tổng Số Lượng],
            [rod0040].[value] [Thành Tiền],
            CASE
                WHEN [rod0040].[type_of_asset] = N'Cổ phiếu thường' THEN [rod0040].[value] * 0.027/100
                WHEN [rod0040].[type_of_asset] IN (N'Chứng quyền', N'Chứng chỉ quỹ') 
                    THEN [rod0040].[value] * 0.018/100
                WHEN [rod0040].[type_of_asset] IN (N'Trái phiếu chính phủ', N'Trái phiếu doanh nghiệp')
                    THEN [rod0040].[value] * 0.0054/100
            END [Phí Phải Nộp]
        FROM (
            SELECT
                [trading_record].[exchange],
                [trading_record].[type_of_asset],
                [branch].[branch_id],
                SUM([trading_record].[volume]) [volume],
                SUM([trading_record].[value]) [value]
            FROM 
                [trading_record]
            LEFT JOIN [relationship]
            ON [relationship].[sub_account] = [trading_record].[sub_account] 
                AND [relationship].[date] = [trading_record].[date]
            LEFT JOIN [branch] 
            ON [branch].[branch_id] = [relationship].[branch_id]
            WHERE [trading_record].[date] BETWEEN '{start_date}' AND '{end_date}'
            GROUP BY [branch].[branch_id], [exchange], [type_of_asset]
        ) [rod0040]
        WHERE [rod0040].[exchange] = 'HOSE'
        ORDER BY [Sàn], [Mã Chi Nhánh]
        """,
        connect_DWH_CoSo,
    )
    result_hnx_upcom = pd.read_sql(
        f"""
        SELECT
            ROW_NUMBER() OVER (ORDER BY [branch_id]) [STT],
            [rod0040].[branch_id] [Mã Chi Nhánh],
            CASE
                WHEN [rod0040].[branch_id] = '0001' THEN N'HQ'
                WHEN [rod0040].[branch_id] = '0101' THEN N'Quận 3'
                WHEN [rod0040].[branch_id] = '0102' THEN N'PMH'
                WHEN [rod0040].[branch_id] = '0104' THEN N'Q7'
                WHEN [rod0040].[branch_id] = '0105' THEN N'TB'
                WHEN [rod0040].[branch_id] = '0116' THEN N'P.QLTK1'
                WHEN [rod0040].[branch_id] = '0111' THEN N'InB1'
                WHEN [rod0040].[branch_id] = '0113' THEN N'IB'
                WHEN [rod0040].[branch_id] = '0201' THEN N'Hà Nội'
                WHEN [rod0040].[branch_id] = '0202' THEN N'TX'
                WHEN [rod0040].[branch_id] = '0301' THEN N'Hải Phòng'
                WHEN [rod0040].[branch_id] = '0117' THEN N'Quận 1'
                WHEN [rod0040].[branch_id] = '0118' THEN N'P.QLTK3'
                WHEN [rod0040].[branch_id] = '0119' THEN N'InB2'
            ELSE [rod0040].[branch_id]
            END [Tên Chi Nhánh],
            CONCAT([rod0040].[branch_id],[rod0040].[exchange]) [ ],
            CASE
                WHEN [rod0040].[exchange] = 'UPCOM' THEN 'UP'
                ELSE [rod0040].[exchange]
            END [Sàn],
            [rod0040].[volume] [Tổng Số Lượng],
            [rod0040].[value] [Thành Tiền],
            CASE
                WHEN [rod0040].[type_of_asset] IN (N'Trái phiếu chính phủ', N'Trái phiếu doanh nghiệp') 
                    THEN [rod0040].[value] * 0.0054/100
                WHEN [rod0040].[exchange] = 'HNX' THEN [rod0040].[value] * 0.027/100
                WHEN [rod0040].[exchange] = 'UPCOM' THEN [rod0040].[value] * 0.018/100
                
            END [Phí Phải Nộp]
        FROM (
            SELECT
                [trading_record].[exchange],
                [trading_record].[type_of_asset],
                [branch].[branch_id],
                SUM([trading_record].[volume]) [volume],
                SUM([trading_record].[value]) [value]
            FROM 
                [trading_record]
            LEFT JOIN 
                [relationship]
            ON 
                [relationship].[sub_account] = [trading_record].[sub_account]
                AND [relationship].[date] = [trading_record].[date]
            LEFT JOIN 
                [branch]
            ON 
                [branch].[branch_id] = [relationship].[branch_id]
            WHERE [trading_record].[date] BETWEEN '{start_date}' AND '{end_date}'
            GROUP BY [branch].[branch_id], [exchange], [type_of_asset]
        ) [rod0040]
        WHERE [rod0040].[exchange] IN ('HNX','UPCOM')
        ORDER BY [Sàn], [Mã Chi Nhánh]
        """,
        connect_DWH_CoSo
    )

    ########################################################################################
    ########################################################################################
    ########################################################################################

    table_title_hose = f'BẢNG KÊ GIÁ DỊCH VỤ GIAO DỊCH {period} - HOSE'

    # write to Bảng kê giá dịch vụ HOSE
    file_name_hose = f'Báo cáo giá dịch vụ giao dịch {period} - HOSE.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name_hose),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:B',10)
    worksheet.set_column('C:C',14)
    worksheet.set_column('D:D',16)
    worksheet.set_column('E:G',20)

    title_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
        }
    )
    header_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    stt_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    san_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    machinhanh_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    tenchinhanh_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    value_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    tong_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    tong_giatri_format_hose = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1,
        }
    )
    worksheet.merge_range('A1:G1',table_title_hose,title_format_hose)
    worksheet.write_row('A3',result_hose.columns,header_format_hose)
    worksheet.write_column('A4',result_hose['STT'],stt_format_hose)
    worksheet.write_column('B4',result_hose['Sàn'],san_format_hose)
    worksheet.write_column('C4',result_hose['Mã Chi Nhánh'],machinhanh_format_hose)
    worksheet.write_column('D4',result_hose['Tên Chi Nhánh'],tenchinhanh_format_hose)
    worksheet.write_column('E4',result_hose['Tổng Số Lượng'],value_format_hose)
    worksheet.write_column('F4',result_hose['Thành Tiền'],value_format_hose)
    worksheet.write_column('G4',result_hose['Phí Phải Nộp'],value_format_hose)
    tong_row = result_hose.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,3,'Tổng',tong_format_hose)
    sum_value_hose = result_hose[['Tổng Số Lượng','Thành Tiền','Phí Phải Nộp']].sum(axis=0)
    worksheet.write_row(tong_row,4,sum_value_hose,tong_giatri_format_hose)

    writer.close()

    ########################################################################################
    ########################################################################################
    ########################################################################################

    table_title_hnx_upcom = f'BẢNG KÊ GIÁ DỊCH VỤ GIAO DỊCH {period} - HNX & UPCOM'

    # write to Bảng kê giá dịch vụ HNX & UPCOM
    file_name_hnx_upcom = f'Báo cáo giá dịch vụ giao dịch {period} - HNX & UPCOM.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name_hnx_upcom),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:B',8)
    worksheet.set_column('C:C',12)
    worksheet.set_column('D:D',11)
    worksheet.set_column('E:E',15)
    worksheet.set_column('F:H',20)

    title_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
        }
    )
    header_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    stt_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    san_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    machinhanh_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    tenchinhanh_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    value_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    tong_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    tong_giatri_format_hnx_upcom = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1,
        }
    )
    worksheet.merge_range('A1:H1',table_title_hnx_upcom,title_format_hnx_upcom)
    worksheet.write_row('A3',result_hnx_upcom.columns,header_format_hnx_upcom)
    worksheet.write_column('A4',result_hnx_upcom['STT'],stt_format_hnx_upcom)
    worksheet.write_column('B4',result_hnx_upcom['Mã Chi Nhánh'],machinhanh_format_hnx_upcom)
    worksheet.write_column('C4',result_hnx_upcom['Tên Chi Nhánh'],tenchinhanh_format_hnx_upcom)
    worksheet.write_column('D4',result_hnx_upcom['Sàn'],san_format_hnx_upcom)
    worksheet.write_column('E4',result_hnx_upcom[' '],san_format_hnx_upcom)
    worksheet.write_column('F4',result_hnx_upcom['Tổng Số Lượng'],value_format_hnx_upcom)
    worksheet.write_column('G4',result_hnx_upcom['Thành Tiền'],value_format_hnx_upcom)
    worksheet.write_column('H4',result_hnx_upcom['Phí Phải Nộp'],value_format_hnx_upcom)
    tong_row = result_hnx_upcom.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,4,'Tổng',tong_format_hnx_upcom)
    sum_value_hnx_upcom = result_hnx_upcom[['Tổng Số Lượng','Thành Tiền','Phí Phải Nộp']].sum(axis=0)
    worksheet.write_row(tong_row,5,sum_value_hnx_upcom,tong_giatri_format_hnx_upcom)
    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')