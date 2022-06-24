from automation.risk_management import *
from datawarehouse import BDATE

def run(  # chạy hàng ngày
    run_time=dt.datetime.now()
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    t0_date = info['end_date']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    file_name = f'Summary High Risk_{t0_date[-2:]}{t0_date[5:7]}{t0_date[0:4]}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    VN30 = [
        'ACB','BID','BVH','CTG','FPT',
        'GAS','GVR','HDB','HPG','KDH',
        'MBB','MSN','MWG','NVL','PDR',
        'PLX','PNJ','POW','SAB','SSI',
        'STB','TCB','TPB','VCB','VHM',
        'VIC','VJC','VNM','VPB','VRE',
    ]
    HNX30 = [
        'BVS','CAP','CEO','DDG','DHT',
        'DP3','DTD','HUT','KLF','L14',
        'LAS','LHC','MBS','NBC','NDN',
        'NRC','NTP','NVB','PVB','PVC',
        'PVS','SHS','SLS','THD','TNG',
        'TVC','VC3','VCS','VMC','DP3',
    ]
    with open(join(dirname(__file__),'sql','HighRisk.sql'),'r') as file:
        SQL = file.read().replace('\n','')
        SQL = SQL.replace(
            '<dataDate>',f"""'{t0_date}'"""
        ).replace(
            '<VN30>',iterable_to_sqlstring(VN30)
        ).replace(
            '<HNX30>',iterable_to_sqlstring(HNX30)
        )
    table = pd.read_sql(SQL,connect_DWH_CoSo)

    # Điều kiện lọc 1:
    filter1a = table['TotalLoan'] > 200
    filter1b = table['SCR'] > 90
    filter1sum = filter1a & filter1b
    # Điều kiện lọc 2:
    filter2a = table['TotalLoan'] > 200
    filter2b = table['SCR'] > 70
    filter2c = table['DL'] > 0.5
    filter2sum = filter2a & filter2b & filter2c
    # Điều kiện lọc 3:
    filter3a = table['TotalLoan'] > 200
    filter3b = ~table['Index'].isin(['VN30','HNX30'])
    filter3c = table['SCR'] > 70
    filter3sum = filter3a & filter3b & filter3c
    # Điều kiện lọc 4:
    filter4a = table['TotalLoan'] > 200
    filter4b = ~table['Index'].isin(['VN30','HNX30'])
    filter4c = table['SCR'] > 20
    filter4d = table['DL'] > 1.5
    filter4sum = filter4a & filter4b & filter4c & filter4d
    # Điều kiện lọc 5:
    filter5a = table['BreakevenPrice'] >= 0
    filter5sum = filter5a
    # Lọc:
    table = table.loc[(filter1sum|filter2sum|filter3sum|filter4sum)&filter5sum]
    # Export danh sách tài khoản - cổ phiếu cho lần chạy sau
    folder = join(dirname(__file__),'temp')
    file_name = f'TempDataOldAccountTicker_{t0_date[0:4]}{t0_date[5:7]}{t0_date[-2:]}.pickle'
    table[['Account','Stock']].to_pickle(join(folder,file_name))
    # Lấy danh sách tài khoản - cổ phiếu gần nhất cho lần chạy hiện tại
    tempFiles = [file for file in listdir(folder) if file.startswith('TempDataOldAccountTicker_')]
    latestFile = sorted(tempFiles)[-2] # lấy file áp cuối vì file cuối là file vừa export
    oldKeys = pd.read_pickle(join(folder,latestFile))
    oldKeys['Key'] = oldKeys['Account'] + oldKeys['Stock']
    # Xác định tài khoản cũ/mới
    table['Type'] = (table['Account']+table['Stock']).isin(oldKeys['Key']) # True: là TK cũ, False: là TK mới
    # Tìm MinRow/MaxRow của từng nhóm (Group): Dùng data allignment
    table = table.sort_values(['Stock','Location','SpecialRoom','Type','SCR'],ascending=[True,True,False,True,False])
    table = table.reset_index(drop=True).reset_index(drop=False) # tạo cột "index" là số tăng dần
    table = table.set_index(['Stock','Location','SpecialRoom'])
    MinMaxLoc = table.groupby(['Stock','Location','SpecialRoom'])['index'].agg(minLoc=np.min,maxLoc=np.max)
    table[['MinLoc','MaxLoc']] = MinMaxLoc
    table = table.reset_index()
    # Không nhóm nếu Special Room = 0 (sửa lại cột MinLoc,MaxLoc)
    table.loc[table['SpecialRoom']==0,'MinLoc'] = table['index']
    table.loc[table['SpecialRoom']==0,'MaxLoc'] = table['index']

    # Write excel
    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    new_case_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#00B050'
        }
    )
    old_case_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FFC000'
        }
    )
    header_normal_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    header_red_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'color':'#FF0000'
        }
    )
    header_bg_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#ED7D31'
        }
    )
    text_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    integer_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'#,##0'
        }
    )
    decimal_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'#,##0.00'
        }
    )
    percent_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0%'
        }
    )
    date_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'[$-409]d-mmm-yy;@'
        }
    )
    headers = [
        'No.',
        'Location',
        'Depository Account',
        'Stock',
        'Quantity',
        'Price',
        'Total Asset Value (Million dong)',
        'Total Loan(Total Outs - Total Cash) (Million dong)',
        'Margin value (Million dong)',
        'SCR Value',
        'DL Value',
        'MR Loan ratio (%)',
        'DP Loan ratio (%)',
        'Max loan price (dong)',
        'General room (setup)',
        'Special room (setup)',
        'Breakeven price (dong)',
        'Total potential outstanding (Billion dong)',
        'Average Volume 3M',
        'Total matched trading volume today',
        '% Break-even price & Max loan price',
        'Reference price',
        '% Reference price & Breakeven price',
        'Industry',
        'Note',
    ]
    HighRiskAccountSheet = workbook.add_worksheet('High Risk Accounts')
    HighRiskAccountSheet.freeze_panes('E6')
    HighRiskAccountSheet.set_column('A:A',3)
    HighRiskAccountSheet.set_column('B:B',9)
    HighRiskAccountSheet.set_column('C:C',10)
    HighRiskAccountSheet.set_column('D:D',6)
    HighRiskAccountSheet.set_column('E:I',10)
    HighRiskAccountSheet.set_column('J:K',6)
    HighRiskAccountSheet.set_column('L:M',5)
    HighRiskAccountSheet.set_column('N:T',10)
    HighRiskAccountSheet.set_column('U:U',7)
    HighRiskAccountSheet.set_column('V:V',10)
    HighRiskAccountSheet.set_column('W:W',7)
    HighRiskAccountSheet.set_column('X:Y',20)

    HighRiskAccountSheet.write('A1','HIGH RISK ACCOUNTS',title_format)
    HighRiskAccountSheet.merge_range('B2:C2',t0_date,date_format)
    HighRiskAccountSheet.merge_range('B3:C3','Old cases',old_case_format)
    HighRiskAccountSheet.merge_range('B4:C4','New cases',new_case_format)
    HighRiskAccountSheet.write_row('A5',headers,header_normal_format)
    HighRiskAccountSheet.write_rich_string(
        'H5',
        header_red_format,'Total Loan ',
        header_normal_format,'(Total Outs - Total Cash) (Million dong)',
        header_normal_format
    )
    HighRiskAccountSheet.write('Q5','Breakeven price (dong)',header_red_format)
    HighRiskAccountSheet.write_row(
        'V5',[
        'Reference price',
        '% Reference price & Breakeven price',
        ],
        header_bg_format
    )
    HighRiskAccountSheet.write_column('A6',np.arange(table.shape[0])+1,integer_format)
    for row in range(table.shape[0]):
        rowType = table.loc[table.index[row],'Type']
        if rowType: # TK cũ
            fmt = old_case_format
        else:
            fmt = new_case_format
        writtenColumns = ['Location','Account','Stock']
        HighRiskAccountSheet.write_row(f'B{row+6}',table.loc[table.index[row],writtenColumns],fmt)
    HighRiskAccountSheet.write_column('E6',table['Quantity'],integer_format)
    HighRiskAccountSheet.write_column('F6',table['Price'],integer_format)
    HighRiskAccountSheet.write_column('G6',table['TotalAsset'],integer_format)
    HighRiskAccountSheet.write_column('H6',table['TotalLoan'],integer_format)
    HighRiskAccountSheet.write_column('I6',table['MarginValue'],integer_format)
    HighRiskAccountSheet.write_column('J6',table['SCR'],integer_format)
    HighRiskAccountSheet.write_column('K6',table['DL'],decimal_format)
    HighRiskAccountSheet.write_column('Q6',table['BreakevenPrice'],integer_format)
    HighRiskAccountSheet.write_column('U6',table['PctBreakevenPriceMaxPrice'],percent_format)
    HighRiskAccountSheet.write_column('V6',table['Price'],integer_format)
    HighRiskAccountSheet.write_column('W6',table['PctBreakevenPriceMarketPrice'],percent_format)
    HighRiskAccountSheet.write_column('X6',['']*table.shape[0],text_format) # Industry

    # Cột MR và DP
    for row in range(table.shape[0]):
        mrRatio = table.loc[table.index[row],'MRRatio']
        dpRatio = table.loc[table.index[row],'DPRatio']
        minLoc = table.loc[table.index[row],'MinLoc']
        maxLoc = table.loc[table.index[row],'MaxLoc']
        if minLoc != maxLoc: # cần phải nhóm
            if minLoc == row: # nếu chưa nhóm
                if mrRatio == dpRatio:
                    HighRiskAccountSheet.merge_range(f'L{minLoc+6}:M{maxLoc+6}',mrRatio,integer_format)
                else:
                    valueString = f'MR:{int(mrRatio)} | DP:{int(dpRatio)}'
                    HighRiskAccountSheet.merge_range(f'L{minLoc+6}:M{maxLoc+6}',valueString,text_format)
            else: # nếu nhóm rồi -> bỏ qua
                continue
        else: # không phải nhóm
            HighRiskAccountSheet.write(f'L{row+6}',mrRatio,integer_format)
            HighRiskAccountSheet.write(f'M{row+6}',dpRatio,integer_format)

    # Cột MaxPrice, GeneralRoom, SpecialRoom, TotalPotentialOutstanding, AvgVolume3M, Volume, Note
    for row in range(table.shape[0]):
        maxPrice = table.loc[table.index[row],'MaxPrice']
        generalRoom = table.loc[table.index[row],'GeneralRoom']
        specialRoom = table.loc[table.index[row],'SpecialRoom']
        totalPotentialOutstanding = table.loc[table.index[row],'TotalPotentialOutstanding']
        avgVolume3M = table.loc[table.index[row],'AvgVolume3M']
        volume = table.loc[table.index[row],'Volume']
        note = ''
        minLoc = table.loc[table.index[row],'MinLoc']
        maxLoc = table.loc[table.index[row],'MaxLoc']
        if minLoc != maxLoc: # cần phải nhóm
            if minLoc == row: # nếu chưa nhóm
                HighRiskAccountSheet.merge_range(f'N{minLoc+6}:N{maxLoc+6}',maxPrice,integer_format)
                HighRiskAccountSheet.merge_range(f'O{minLoc+6}:O{maxLoc+6}',generalRoom,integer_format)
                HighRiskAccountSheet.merge_range(f'P{minLoc+6}:P{maxLoc+6}',specialRoom,integer_format)
                HighRiskAccountSheet.merge_range(f'R{minLoc+6}:R{maxLoc+6}',totalPotentialOutstanding,integer_format)
                HighRiskAccountSheet.merge_range(f'S{minLoc+6}:S{maxLoc+6}',avgVolume3M,integer_format)
                HighRiskAccountSheet.merge_range(f'T{minLoc+6}:T{maxLoc+6}',volume,integer_format)
                HighRiskAccountSheet.merge_range(f'Y{minLoc+6}:Y{maxLoc+6}',note,text_format)
            else: # nếu nhóm rồi -> bỏ qua
                continue
        else: # không phải nhóm
            HighRiskAccountSheet.write(f'N{minLoc+6}',maxPrice,integer_format)
            HighRiskAccountSheet.write(f'O{minLoc+6}',generalRoom,integer_format)
            HighRiskAccountSheet.write(f'P{minLoc+6}',specialRoom,integer_format)
            HighRiskAccountSheet.write(f'R{minLoc+6}',totalPotentialOutstanding,integer_format)
            HighRiskAccountSheet.write(f'S{minLoc+6}',avgVolume3M,integer_format)
            HighRiskAccountSheet.write(f'T{minLoc+6}',volume,integer_format)
            HighRiskAccountSheet.write(f'Y{minLoc+6}',note,text_format)

    ###################################################
    ###################################################
    ###################################################

    with open(join(dirname(__file__),'sql','Liquidity.sql'),'r') as file:
        SQL = file.read().replace('\n','').replace('<dataDate>',f"""'{t0_date}'""")

    table = pd.read_sql(SQL,connect_DWH_CoSo)
    # Tìm MinLoc / MaxLoc (không làm được bằng SQL PARTITION do lỗi data size)
    table = table.reset_index(drop=True).reset_index(drop=False) # tạo cột "index" là số tăng dần
    table = table.set_index('TaiKhoan')
    MinMaxLoc = table.groupby('TaiKhoan')['index'].agg(minLoc=np.min,maxLoc=np.max)
    table[['MinLoc','MaxLoc']] = MinMaxLoc
    table = table.reset_index()

    title_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman'
        }
    )
    header1_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#70AD47'
        }
    )
    header2_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#70AD47'
        }
    )
    header3_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FFC009'
        }
    )
    header4_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FFC000'
        }
    )
    header5_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'color':'#FF0000'
        }
    )
    text_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    text_red_format = workbook.add_format(
        {
            'text_wrap':True,
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'font_color':'#FF0000',
        }
    )
    note_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    date_format = workbook.add_format(
        {
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'dd/mm/yyyy'
        }
    )
    integer_right_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    integer_right_red_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'font_color':'#FF0000'
        }
    )
    integer_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    integer_center_red_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'font_color':'#FF0000',
        }
    )
    decimal_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    subtotal_format = workbook.add_format(
        {
            'bold':True,
            'align':'right',
            'valign':'vbottom',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    headers1 = [
        'No.',
        'Mã Room',
        'Code',
        'Setup date',
    ]
    headers2 = [
        'Approved Date',
        'Stock',
        'Location',
        'Account',
        'Approved Quantity',
        'Set up',
        'MR Approved Ratio (%)',
        'DP Approved Ratio (%)',
        'Maximum loan price',
    ]
    headers3 = [
        'P.Outs',
    ]
    headers4 = [
        'Used quantity today',
        'Average Liquidity Volume of 03 months',
        'Liquidity today',
    ]
    headers5 = [
        'Approval room/3months average liquidity volume',
        'Commitment Fix Type other special',
    ]
    LiquidityDealSheet = workbook.add_worksheet('Liquidity Deal Report')
    LiquidityDealSheet.freeze_panes('H3')
    LiquidityDealSheet.set_column('A:A',5)
    LiquidityDealSheet.set_column('B:B',8,None,{'hidden':True})
    LiquidityDealSheet.set_column('C:C',10,None,{'hidden':True})
    LiquidityDealSheet.set_column('D:D',12,None,{'hidden':True})
    LiquidityDealSheet.set_column('E:E',12)
    LiquidityDealSheet.set_column('F:G',11)
    LiquidityDealSheet.set_column('H:H',35)
    LiquidityDealSheet.set_column('I:I',15)
    LiquidityDealSheet.set_column('J:K',16)
    LiquidityDealSheet.set_column('L:R',15)
    LiquidityDealSheet.set_column('S:S',55)

    # Write title
    LiquidityDealSheet.write('E1',"DAILY  DEAL's LIQUIDITY REPORT",title_format)
    # Write headers
    for colNum,colName in enumerate(headers1+headers2+headers3+headers4+headers5):
        if colName in headers1:
            fmt = header1_format
        elif colName in headers2:
            fmt = header2_format
        elif colName in headers3:
            fmt = header3_format
        elif colName in headers4:
            fmt = header4_format
        else:
            fmt = header5_format
        LiquidityDealSheet.write(1,colNum,colName,fmt)

    # Write các cột ko có rule
    LiquidityDealSheet.write_column('A3',np.arange(table.shape[0])+1,text_format)
    LiquidityDealSheet.write_column('B3',table['MaRoom'],text_format)
    LiquidityDealSheet.write_column('C3',table['Code'],text_format)
    LiquidityDealSheet.write_column('D3',table['SetUpDate'],date_format)
    LiquidityDealSheet.write_column('E3',['']*table.shape[0],text_format)
    LiquidityDealSheet.write_column('G3',table['ChiNhanh'],text_format)
    LiquidityDealSheet.write_column('I3',['']*table.shape[0],text_format)
    LiquidityDealSheet.write_column('J3',table['SetUp'],integer_right_format)
    LiquidityDealSheet.write_column('L3',table['DPRatioChung'],integer_center_format)
    LiquidityDealSheet.write_column('N3',table['Outstanding'],integer_right_format)
    LiquidityDealSheet.write_column('O3',table['UsedRoom'],integer_right_format) # Used Quantity Today
    LiquidityDealSheet.write_column('P3',table['AvgVolume'],integer_right_format)
    LiquidityDealSheet.write_column('Q3',table['LastVolume'],integer_right_format)
    LiquidityDealSheet.write_column('R3',[f'=I{row}/P{row}' for row in np.arange(table.shape[0])+3],decimal_format)
    LiquidityDealSheet.write_column('S3',['']*table.shape[0],note_format)

    # Write các cột có rule: Stock, MR Approved Ratio (%), DP Approved Ratio (%), Maximum loan price
    for row in range(table.shape[0]):
        ticker = table.loc[table.index[row],'Ticker']
        mrRatioRieng = table.loc[table.index[row],'MRRatioRieng']
        mrRatioChung = table.loc[table.index[row],'MRRatioChung']
        maxPriceRieng = table.loc[table.index[row],'MaxPriceRieng']
        maxPriceChung = table.loc[table.index[row],'MaxPriceChung']
        condition1 = mrRatioRieng == mrRatioRieng # (ko bị np.nan)
        condition2 = maxPriceRieng == maxPriceRieng # (ko bị np.nan)
        if condition1:
            LiquidityDealSheet.write(f'K{row+3}',mrRatioRieng,integer_center_red_format)
        else:
            LiquidityDealSheet.write(f'K{row+3}',mrRatioChung,integer_center_format)
        if condition2:
            LiquidityDealSheet.write(f'M{row+3}',maxPriceRieng,integer_right_red_format)
        else:
            LiquidityDealSheet.write(f'M{row+3}',maxPriceChung,integer_right_format)
        if condition1 or condition2:
            LiquidityDealSheet.write(f'F{row+3}',ticker,text_red_format)
        else:
            LiquidityDealSheet.write(f'F{row+3}',ticker,text_format)

    # Nhóm cột tài khoản
    for row in range(table.shape[0]):
        taikhoan = table.loc[table.index[row],'TaiKhoan']
        minLoc = table.loc[table.index[row],'MinLoc']
        maxLoc = table.loc[table.index[row],'MaxLoc']
        if minLoc != maxLoc: # cần phải nhóm
            if minLoc == row: # nếu chưa nhóm
                LiquidityDealSheet.merge_range(f'H{minLoc+3}:H{maxLoc+3}',taikhoan,text_format)
            else: # nếu nhóm rồi
                continue
        else: # không phải nhóm
            LiquidityDealSheet.write(f'H{row+3}',taikhoan,text_format)

    # Tính dòng tổng
    LiquidityDealSheet.write(f'I{table.shape[0]+3}',f'=SUM(I3:I{table.shape[0]+2})',subtotal_format)
    LiquidityDealSheet.write(f'N{table.shape[0]+3}',f'=SUM(N3:N{table.shape[0]+2})',subtotal_format)

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start,1)}s')
