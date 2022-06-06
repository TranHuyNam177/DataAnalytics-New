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

    file_name = f'Summary High Risk_{t0_date[-2:]}{t0_date[5:7]}{t0_date[0:4]} ThuAnh.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Tạm thời hard-key trước, chờ các bảng từ DWh-ThiTruong
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
    # Sheet High Risk
    header_format = workbook.add_format(
        {
            'bold':True,
            'text_wrap':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FFC000',
        }
    )
    red_column = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'bg_color':'#FF0000'
        }
    )
    text_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    decimal_format = workbook.add_format(
        {
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'0.0000_);(0.0000)',
            'border':1,
        }
    )
    integer_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'right',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'num_format':'#,##0_);(#,##0)',
        }
    )
    highRiskSheet = workbook.add_worksheet('HighRisk')
    highRiskSheet.set_tab_color('#FFC000')
    highRiskSheet.set_column('A:A',7)
    highRiskSheet.set_column('B:B',10)
    highRiskSheet.set_column('C:C',11)
    highRiskSheet.set_column('D:D',6)
    highRiskSheet.set_column('F:F',11)
    highRiskSheet.set_column('G:I',8)
    highRiskSheet.set_column('J:W',9)
    highRiskSheet.set_column('X:X',1)
    highRiskSheet.set_column('Y:AB',9)

    headers = [
        'Index',
        'Depository Account',
        'Branch',
        'Stock',
        'Quantity',
        'Market Price',
        'Total Asset Value (Million dong)',
        'Total Loan (Total Outs - Total Cash) (Million dong)',
        'Margin value (Million dong)',
        'SCR Value',
        'DL Value',
        'MR Loan ratio (%)',
        'DP Loan ratio (%)',
        'Max loan price (dong)',
        'General room (approved)',
        'Special room (approved)',
        'Break-even price (dong)',
        'Total potential outstanding (Billion dong)',
        'Average Volume 3M',
        'Total matched trading volume today',
        '% giá hòa vốn&Max loan price',
        'Maket price',
        '% giá hòa vốn& Maket price',
        '',
        'Position market value (Margin value)',
        'Total Asset Value',
        'Total Outstanding',
        'Cash',
    ]
    highRiskSheet.write_row('A1',headers,header_format)
    highRiskSheet.write_column('A2',table['Index'],text_format)
    highRiskSheet.write_column('B2',table['Account'],text_format)
    highRiskSheet.write_column('C2',table['Location'],text_format)
    highRiskSheet.write_column('D2',table['Stock'],text_format)
    highRiskSheet.write_column('E2',table['Quantity'],integer_format)
    highRiskSheet.write_column('F2',table['Price'],integer_format)
    highRiskSheet.write_column('G2',table['TotalAsset'],integer_format)
    highRiskSheet.write_column('H2',table['TotalLoan'],integer_format)
    highRiskSheet.write_column('I2',table['MarginValue'],integer_format)
    highRiskSheet.write_column('J2',table['SCR'],decimal_format)
    highRiskSheet.write_column('K2',table['DL'],decimal_format)
    highRiskSheet.write_column('L2',table['MRRatio'],decimal_format)
    highRiskSheet.write_column('M2',table['DPRatio'],decimal_format)
    highRiskSheet.write_column('N2',table['MaxPrice'],integer_format)
    highRiskSheet.write_column('O2',table['GeneralRoom'],integer_format)
    highRiskSheet.write_column('P2',table['SpecialRoom'],integer_format)
    highRiskSheet.write_column('Q2',table['BreakevenPrice'],integer_format)
    highRiskSheet.write_column('R2',table['TotalPotentialOutstanding'],integer_format)
    highRiskSheet.write_column('S2',table['AvgVolume3M'],integer_format)
    highRiskSheet.write_column('T2',table['Volume'],integer_format)
    highRiskSheet.write_column('U2',table['PctBreakevenPriceMaxPrice'],decimal_format)
    highRiskSheet.write_column('V2',table['ClosePrice'],integer_format)
    highRiskSheet.write_column('W2',table['PctBreakevenPriceMarketPrice'],decimal_format)
    highRiskSheet.write_column('X2',['']*table.shape[0],red_column)
    highRiskSheet.write_column('Y2',table['MarginValue'],integer_format)
    highRiskSheet.write_column('Z2',table['TotalAsset'],integer_format)
    highRiskSheet.write_column('AA2',table['TotalLoan'],integer_format)
    highRiskSheet.write_column('AB2',table['Cash'],integer_format)

    ###################################################
    ###################################################
    ###################################################

    with open(join(dirname(__file__),'sql','SpecialRoom.sql'),'r') as file:
        SQL = file.read().replace('\n','')
        SQL = SQL.replace('<dataDate>',f"""'{t0_date}'""")
    table = pd.read_sql(SQL,connect_DWH_CoSo)
    # Sheet Special Room
    specialRoomSheet = workbook.add_worksheet('SpecialRoom')
    specialRoomSheet.set_tab_color('#FFC000')
    specialRoomSheet.set_column('A:A',15)
    specialRoomSheet.set_column('B:B',9)
    specialRoomSheet.set_column('C:C',10)
    specialRoomSheet.set_column('D:D',11)
    specialRoomSheet.set_column('E:E',7)
    specialRoomSheet.set_column('F:F',14)
    specialRoomSheet.set_column('G:G',22)
    headers = [
        'TK&Stock',
        'Code',
        'Mã Room',
        'Account',
        'Mã CK',
        'Tổng số lượng',
        'Group/Deal'
    ]
    specialRoomSheet.write_row('A1',headers,header_format)
    specialRoomSheet.write_column('A2',table['TKStock'],text_format)
    specialRoomSheet.write_column('B2',table['Code'],text_format)
    specialRoomSheet.write_column('C2',table['MaRoom'],text_format)
    specialRoomSheet.write_column('D2',table['TaiKhoan'],text_format)
    specialRoomSheet.write_column('E2',table['MaCK'],text_format)
    specialRoomSheet.write_column('F2',table['SpecialRoom'],integer_format)
    specialRoomSheet.write_column('G2',table['GroupDeal'],text_format)

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start,1)}s')
