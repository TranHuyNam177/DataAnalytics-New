import pandas as pd
from tqdm import tqdm
from automation.accounting import *


# Client code
def run(
    run_time=dt.datetime.now()
):
    report = Report(run_time)
    for func in tqdm([report.runRDT0121,report.runRDT0141,report.runRDT0127,report.runRDO0002],ncols=70):
        func()


class Report:

    def __init__(self,run_time: dt.datetime):

        info = get_info('daily',run_time)
        period = info['period']
        t0_date = info['end_date'].replace('.','-')
        folder_name = info['folder_name']

        # create folder
        if not os.path.isdir(join(dept_folder,folder_name,period)):
            os.mkdir((join(dept_folder,folder_name,period)))

        self.bravoFolder = join(dirname(dept_folder),'FileFromBravo')
        self.bravoDateString = run_time.strftime('%Y.%m.%d')

        self.file_date = dt.datetime.strptime(t0_date,'%Y-%m-%d').strftime('%d.%m.%Y')
        # date in sub title in Excel
        self.sub_title_date = dt.datetime.strptime(t0_date,'%Y-%m-%d').strftime('%d/%m/%Y')

        self.file_name = f'Đối Chiếu Phái Sinh {self.file_date}.xlsx'
        self.writer = pd.ExcelWriter(
            join(dept_folder,folder_name,period,self.file_name),
            engine='xlsxwriter',
            engine_kwargs={'options':{'nan_inf_to_errors':True}}
        )
        self.workbook = self.writer.book
        self.company_name_format = self.workbook.add_format(
            {
                'bold':True,
                'align':'left',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'text_wrap':True
            }
        )
        self.company_info_format = self.workbook.add_format(
            {
                'align':'left',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'text_wrap':True
            }
        )
        self.empty_row_format = self.workbook.add_format(
            {
                'bottom':1,
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
            }
        )
        self.sheet_title_format = self.workbook.add_format(
            {
                'bold':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':14,
                'font_name':'Times New Roman',
                'text_wrap':True
            }
        )
        self.sub_title_format = self.workbook.add_format(
            {
                'bold':True,
                'italic':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'text_wrap':True
            }
        )
        self.headers_root_format = self.workbook.add_format(
            {
                'border':1,
                'bold':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'text_wrap':True
            }
        )
        self.headers_bravo_format = self.workbook.add_format(
            {
                'border':1,
                'bold':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'text_wrap':True,
                'bg_color':'#DAEEF3',
            }
        )
        self.headers_fds_format = self.workbook.add_format(
            {
                'border':1,
                'bold':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'text_wrap':True,
                'bg_color':'#EBF1DE',
            }
        )
        self.headers_diff_format = self.workbook.add_format(
            {
                'border':1,
                'bold':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'text_wrap':True,
                'bg_color':'#FFFFCC',
            }
        )
        self.text_root_format = self.workbook.add_format(
            {
                'border':1,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
            }
        )
        self.money_bravo_format = self.workbook.add_format(
            {
                'border':1,
                'align':'right',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'num_format':'#,##0_);(#,##0)',
                'bg_color':'#DAEEF3',
            }
        )
        self.money_fds_format = self.workbook.add_format(
            {
                'border':1,
                'align':'right',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'num_format':'#,##0_);(#,##0)',
                'bg_color':'#EBF1DE',
            }
        )
        self.money_diff_format = self.workbook.add_format(
            {
                'border':1,
                'align':'right',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
                'num_format':'#,##0_);(#,##0)',
                'bg_color':'#FFFFCC',
            }
        )

    def __del__(self):
        self.writer.close()

    def runRDT0121(self):
        TaiKhoan3243 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Sổ tổng hợp công nợ 3243_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols='A,B,H',
            names=['SoTaiKhoan','TenKhachHang3243','DuCoCuoi3243']
        )
        TaiKhoan338804 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Sổ tổng hợp công nợ 338804_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols='A,B,H',
            names=['SoTaiKhoan','TenKhachHang338804','DuCoCuoi338804']
        )
        RDT0121 = pd.read_sql(
            f"""
            SELECT
                [relationship].[branch_id] [MaChiNhanh],
                [rdt0121].[account_code] [SoTaiKhoan],
                [account].[customer_name] [TenKhachHangFDS],
                [rdt0121].[cash_balance_at_phs] [TienTaiPHS],
                [rdt0121].[cash_balance_at_vsd] [TienTaiVSD]
            FROM [rdt0121]
            LEFT JOIN [relationship] ON [relationship].[account_code] = [rdt0121].[account_code] AND [relationship].[date] = [rdt0121].[date]
            LEFT JOIN [account] ON [account].[account_code] = [rdt0121].[account_code]
            WHERE [rdt0121].[date] = '{self.bravoDateString}'
            """,
            connect_DWH_PhaiSinh
        )
        table = RDT0121.merge(TaiKhoan3243,how='outer',on='SoTaiKhoan').merge(TaiKhoan338804,how='outer',on='SoTaiKhoan')
        table['TenKhachHang'] = table['TenKhachHang3243'].fillna(table['TenKhachHang338804']).fillna(table['TenKhachHangFDS']).fillna('').str.title()
        table['MaChiNhanh'] = table['MaChiNhanh'].fillna('')
        table = table.fillna(0)
        table['PHSDiff'] = table['TienTaiPHS'] - table['DuCoCuoi3243']
        table['VSDDiff'] = table['TienTaiVSD'] - table['DuCoCuoi338804']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDT0121')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('E13')
        worksheet.set_column('A:A',4)
        worksheet.set_column('B:B',11)
        worksheet.set_column('C:C',15)
        worksheet.set_column('D:D',29)
        worksheet.set_column('E:J',19)

        worksheet.merge_range('A1:J1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:J2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:J3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*10,self.empty_row_format)
        worksheet.merge_range('A6:J6','ĐỐI CHIẾU SỔ DƯ TIỀN NHÀ ĐẦU TƯ',self.sheet_title_format)
        worksheet.merge_range('A7:J7',f'Date: {self.file_date}',self.sub_title_format)
        worksheet.merge_range('A8:J8','Tài khoản: 3243, 338804 & FDS: RDT0121',self.sub_title_format)

        worksheet.merge_range('E10:F10','FDS',self.headers_fds_format)
        worksheet.merge_range('G10:H10','Bravo',self.headers_bravo_format)
        worksheet.merge_range('I10:J10','Chênh lệch',self.headers_diff_format)
        for col, header in zip('ABCD',['STT','Mã chi nhánh','Tài khoản ký quỹ','Tên khách hàng']):
            worksheet.merge_range(f'{col}10:{col}11',header,self.headers_root_format)
        worksheet.write_row('E11',['Số tiền tại công ty\n(RDT0121)','Số tiền ký quỹ tại VSD\n(RDT0121)',],self.headers_fds_format)
        worksheet.write_row('G11',['Số tiền tại công ty\n(Tài khoản: 3243)','Số tiền ký quỹ tại VSD\n(Tài khoản: 338804)',],self.headers_bravo_format)
        worksheet.write_row('I11',['Số tiền tại công ty','Số tiền ký quỹ tại VSD',],self.headers_diff_format)
        worksheet.write_row('A12',np.arange(1,5),self.headers_root_format)
        worksheet.write_row('E12',np.arange(5,7),self.headers_fds_format)
        worksheet.write_row('G12',np.arange(7,9),self.headers_bravo_format)
        worksheet.write_row('I12',np.arange(9,11),self.headers_diff_format)

        worksheet.write_column('A13',np.arange(table.shape[0])+1,self.text_root_format)
        worksheet.write_column('B13',table['MaChiNhanh'],self.text_root_format)
        worksheet.write_column('C13',table['SoTaiKhoan'],self.text_root_format)
        worksheet.write_column('D13',table['TenKhachHang'],self.text_root_format)
        worksheet.write_column('E13',table['TienTaiPHS'],self.money_fds_format)
        worksheet.write_column('F13',table['TienTaiVSD'],self.money_fds_format)
        worksheet.write_column('G13',table['DuCoCuoi3243'],self.money_bravo_format)
        worksheet.write_column('H13',table['DuCoCuoi338804'],self.money_bravo_format)
        worksheet.write_column('I13',table['PHSDiff'],self.money_diff_format)
        worksheet.write_column('J13',table['VSDDiff'],self.money_diff_format)

        worksheet.merge_range(f'A{table.shape[0]+13}:D{table.shape[0]+13}','Tổng cộng:',self.headers_root_format)
        for col in 'EFGHIJ':
            if col in 'EF':
                fmt = self.money_fds_format
            elif col in 'GH':
                fmt = self.money_bravo_format
            else:
                fmt = self.money_diff_format
            sumString = f'=SUM({col}13:{col}{table.shape[0]+12})'
            worksheet.write(f'{col}{table.shape[0]+13}',sumString,fmt)


    def runRDT0141(self):

        TaiKhoan13504 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Sổ tổng hợp công nợ 13504_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols='A,B,E:G',
            names=['SoTaiKhoan','TenKhachHang13504','PhatSinhNo13504','PhatSinhCo13504','DuNoCuoi13504']
        )
        TaiKhoan13505 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Sổ tổng hợp công nợ 13505_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols='A,B,E:G',
            names=['SoTaiKhoan','TenKhachHang13505','PhatSinhNo13505','PhatSinhCo13505','DuNoCuoi13505']
        )
        RDT0141 = pd.read_sql(
            f"""
            SELECT
                [account].[account_code] [SoTaiKhoan],
                [rdt0141].[sub_account] [SoTieuKhoan],
                [account].[customer_name] [TenKhachHangFDS],
                [relationship].[branch_id] [MaChiNhanh],
                [rdt0141].[deferred_payment_amount_opening] [KhoanChamTraDauKyRDT0141],
                [rdt0141].[deferred_payment_fee_opening] [PhiChamTraDauKyRDT0141],
                [rdt0141].[deferred_payment_amount_opening] + [rdt0141].[deferred_payment_fee_opening] [TongTienChamDauKyRDT0141],
                [rdt0141].[deferred_payment_amount_increase] [KhoanChamTraPSTangTrongKyRDT0141],
                [rdt0141].[deferred_payment_fee_increase] [PhiChamTraPSTangTrongKyRDT0141],
                [rdt0141].[deferred_payment_amount_decrease] [KhoanChamTraPSGiamTrongKyRDT0141],
                [rdt0141].[deferred_payment_fee_decrease] [PhiChamTraPSGiamTrongKyRDT0141],
                [rdt0141].[deferred_payment_amount_closing] [KhoanChamTraCuoiKyRDT0141],
                [rdt0141].[deferred_payment_fee_closing] [PhiChamTraCuoiKyRDT0141],
                [rdt0141].[deferred_payment_amount_closing] + [rdt0141].[deferred_payment_fee_closing] [TongTienChamCuoiKyRDT0141]
            FROM [rdt0141]
            LEFT JOIN [relationship] ON [relationship].[sub_account] = [rdt0141].[sub_account] AND [relationship].[date] = [rdt0141].[date]
            LEFT JOIN [account] ON [account].[account_code] = [relationship].[account_code]
            WHERE [rdt0141].[date] = '{self.bravoDateString}'
            """,
            connect_DWH_PhaiSinh
        )
        table = RDT0141.merge(TaiKhoan13504,how='outer',on='SoTaiKhoan').merge(TaiKhoan13505,how='outer',on='SoTaiKhoan')
        table['TenKhachHang'] = table['TenKhachHang13504'].fillna(table['TenKhachHang13505']).fillna(table['TenKhachHangFDS']).fillna('').str.title()
        table['SoTieuKhoan'] = table['SoTieuKhoan'].fillna('')
        table['MaChiNhanh'] = table['MaChiNhanh'].fillna('')
        table = table.fillna(0)

        table['KhoanChamTraCuoiKyDiff'] = table['KhoanChamTraCuoiKyRDT0141'] - table['DuNoCuoi13504']
        table['PhiChamTraCuoiKyDiff'] = table['PhiChamTraCuoiKyRDT0141'] - table['DuNoCuoi13505']
        table['KhoanChamTraPSTangTrongKyDiff'] = table['KhoanChamTraPSTangTrongKyRDT0141'] - table['PhatSinhNo13504']
        table['PhiChamTraPSTangTrongKyDiff'] = table['PhiChamTraPSTangTrongKyRDT0141'] - table['PhatSinhNo13505']
        table['KhoanChamTraPSGiamTrongKyDiff'] = table['KhoanChamTraPSGiamTrongKyRDT0141'] - table['PhatSinhCo13504']
        table['PhiChamTraPSGiamTrongKyDiff'] = table['PhiChamTraPSGiamTrongKyRDT0141'] - table['PhatSinhCo13505']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDT0141')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('E14')
        worksheet.set_column('A:B',11)
        worksheet.set_column('C:C',29)
        worksheet.set_column('D:D',9)
        worksheet.set_column('E:Z',14)

        worksheet.merge_range('A1:Z1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:Z2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:Z3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*26,self.empty_row_format)
        worksheet.merge_range('A6:Z6','ĐỐI CHIẾU SỔ DƯ TIỀN NHÀ ĐẦU TƯ',self.sheet_title_format)
        worksheet.merge_range('A7:Z7',f'Date: {self.file_date}',self.sub_title_format)
        worksheet.merge_range('A8:Z8','Tài khoản: 13504, 13505 & FDS: RDT0141',self.sub_title_format)

        worksheet.merge_range('E10:N10','FDS',self.headers_fds_format)
        worksheet.merge_range('O10:T10','Bravo',self.headers_bravo_format)
        worksheet.merge_range('U10:Z10','Chênh lệch',self.headers_diff_format)

        worksheet.merge_range('E11:G11','Đầu kỳ',self.headers_fds_format)
        worksheet.merge_range('H11:I11','Phát sinh tăng',self.headers_fds_format)
        worksheet.merge_range('J11:K11','Phát sinh giảm',self.headers_fds_format)
        worksheet.merge_range('L11:N11','Cuối kỳ',self.headers_fds_format)
        worksheet.merge_range('O11:P11','Cuối kỳ',self.headers_bravo_format)
        worksheet.merge_range('Q11:R11','Phát sinh tăng',self.headers_bravo_format)
        worksheet.merge_range('S11:T11','Phát sinh giảm',self.headers_bravo_format)
        worksheet.merge_range('U11:V11','Cuối kỳ',self.headers_diff_format)
        worksheet.merge_range('W11:X11','Phát sinh tăng',self.headers_diff_format)
        worksheet.merge_range('Y11:Z11','Phát sinh giảm',self.headers_diff_format)
        for col, header in zip('ABCD',['Tài khoản ký quỹ','Tài khoản giao dịch','Tên khách hàng','Chi nhánh']):
            worksheet.merge_range(f'{col}10:{col}12',header,self.headers_root_format)
        worksheet.write_row('E12',['Khoản chậm trả\n(RDT0141)','Phí chậm trả\n(RDT0141)','Tổng số tiền chậm\n(RDT0141)'],self.headers_fds_format)
        worksheet.write_row('H12',['Khoản chậm trả\n(RDT0141)','Phí chậm trả\n(RDT0141)']*2,self.headers_fds_format)
        worksheet.write_row('L12',['Khoản chậm trả\n(RDT0141)','Phí chậm trả\n(RDT0141)','Tổng số tiền chậm\n(RDT0141)'],self.headers_fds_format)
        worksheet.write_row('O12',['Khoản chậm trả\n(Tài khoản: 13504)','Phí chậm trả\n(Tài khoản: 13505)']*3,self.headers_bravo_format)
        worksheet.write_row('U12',['Khoản chậm trả','Phí chậm trả']*3,self.headers_diff_format)
        worksheet.write_row('A13',np.arange(1,5),self.headers_root_format)
        worksheet.write_row('E13',np.arange(5,15),self.headers_fds_format)
        worksheet.write_row('O13',np.arange(15,22),self.headers_bravo_format)
        worksheet.write_row('U13',np.arange(22,28),self.headers_diff_format)

        worksheet.write_column('A14',table['SoTaiKhoan'],self.text_root_format)
        worksheet.write_column('B14',table['SoTieuKhoan'],self.text_root_format)
        worksheet.write_column('C14',table['TenKhachHang'],self.text_root_format)
        worksheet.write_column('D14',table['MaChiNhanh'],self.text_root_format)
        worksheet.write_column('E14',table['KhoanChamTraDauKyRDT0141'],self.money_fds_format)
        worksheet.write_column('F14',table['PhiChamTraDauKyRDT0141'],self.money_fds_format)
        worksheet.write_column('G14',table['TongTienChamDauKyRDT0141'],self.money_fds_format)
        worksheet.write_column('H14',table['KhoanChamTraPSTangTrongKyRDT0141'],self.money_fds_format)
        worksheet.write_column('I14',table['PhiChamTraPSTangTrongKyRDT0141'],self.money_fds_format)
        worksheet.write_column('J14',table['KhoanChamTraPSGiamTrongKyRDT0141'],self.money_fds_format)
        worksheet.write_column('K14',table['PhiChamTraPSGiamTrongKyRDT0141'],self.money_fds_format)
        worksheet.write_column('L14',table['KhoanChamTraCuoiKyRDT0141'],self.money_fds_format)
        worksheet.write_column('M14',table['PhiChamTraCuoiKyRDT0141'],self.money_fds_format)
        worksheet.write_column('N14',table['TongTienChamCuoiKyRDT0141'],self.money_fds_format)
        worksheet.write_column('O14',table['DuNoCuoi13504'],self.money_bravo_format)
        worksheet.write_column('P14',table['DuNoCuoi13505'],self.money_bravo_format)
        worksheet.write_column('Q14',table['PhatSinhNo13504'],self.money_bravo_format)
        worksheet.write_column('R14',table['PhatSinhNo13505'],self.money_bravo_format)
        worksheet.write_column('S14',table['PhatSinhCo13504'],self.money_bravo_format)
        worksheet.write_column('T14',table['PhatSinhCo13505'],self.money_bravo_format)
        worksheet.write_column('U14',table['KhoanChamTraCuoiKyDiff'],self.money_diff_format)
        worksheet.write_column('V14',table['PhiChamTraCuoiKyDiff'],self.money_diff_format)
        worksheet.write_column('W14',table['KhoanChamTraPSTangTrongKyDiff'],self.money_diff_format)
        worksheet.write_column('X14',table['PhiChamTraPSTangTrongKyDiff'],self.money_diff_format)
        worksheet.write_column('Y14',table['KhoanChamTraPSGiamTrongKyDiff'],self.money_diff_format)
        worksheet.write_column('Z14',table['PhiChamTraPSGiamTrongKyDiff'],self.money_diff_format)
        worksheet.merge_range(f'A{table.shape[0]+14}:D{table.shape[0]+14}','Tổng cộng:',self.headers_root_format)
        for col in 'EFGHIJKLMNOPQRSTUVWXYZ':
            if col in 'EFGHIJKLMN':
                fmt = self.money_fds_format
            elif col in 'OPQRST':
                fmt = self.money_bravo_format
            else:
                fmt = self.money_diff_format
            worksheet.write(f'{col}{table.shape[0]+14}',f'=SUM({col}14:{col}{table.shape[0]+13})',fmt)


    def runRDT0127(self):

        TaiKhoan33353 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Bang ke chung tu 33353.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols='G,K',
            names=['ThueTNCNBravo','SoTaiKhoan']
        )
        RDT0127 = pd.read_sql(
            f"""
            SELECT
                [RDT0127].[SoTaiKhoan],
                SUM([RDT0127].[ThueTNCN]) [ThueTNCNFDS]
            FROM [RDT0127]
            WHERE [RDT0127].[Ngay] = '{self.bravoDateString}'
            GROUP BY [RDT0127].[SoTaiKhoan]
            """,
            connect_DWH_PhaiSinh
        )
        table = pd.merge(TaiKhoan33353,RDT0127,how='outer',on='SoTaiKhoan')
        table = table.fillna(0)
        table['ThueTNCNDiff'] = table['ThueTNCNFDS'] - table['ThueTNCNBravo']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDT0127')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('B11')
        worksheet.set_column('A:A',11)
        worksheet.set_column('B:D',15)

        worksheet.merge_range('A1:H1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:H2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:H3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*4,self.empty_row_format)
        worksheet.merge_range('A6:D6','ĐỐI CHIẾU THUẾ TNCN',self.sheet_title_format)
        worksheet.merge_range('A7:D7',f'Date: {self.file_date}',self.sub_title_format)
        worksheet.merge_range('A8:D8','Tài khoản: 33353 & FDS: RDT0127',self.sub_title_format)

        worksheet.write('A10','Số tài khoản',self.headers_root_format)
        worksheet.write('B10','FDS',self.headers_fds_format)
        worksheet.write('C10','Bravo',self.headers_bravo_format)
        worksheet.write('D10','Chênh lệch',self.headers_diff_format)

        worksheet.write_column('A11',table['SoTaiKhoan'],self.text_root_format)
        worksheet.write_column('B11',table['ThueTNCNFDS'],self.money_fds_format)
        worksheet.write_column('C11',table['ThueTNCNBravo'],self.money_bravo_format)
        worksheet.write_column('D11',table['ThueTNCNDiff'],self.money_diff_format)

        worksheet.write(f'A{table.shape[0]+11}','Tổng',self.headers_root_format)
        worksheet.write(f'B{table.shape[0]+11}',f'=SUM(B11:B{table.shape[0]+10})',self.money_fds_format)
        worksheet.write(f'C{table.shape[0]+11}',f'=SUM(C11:C{table.shape[0]+10})',self.money_bravo_format)
        worksheet.write(f'D{table.shape[0]+11}',f'=SUM(D11:D{table.shape[0]+10})',self.money_diff_format)


    def runRDO0002(self):

        TaiKhoan5115104 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Bang ke chung tu 5115104.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols='G,K',
            names=['PhiGiaoDichBravo','SoTaiKhoan']
        )
        RDO0002 = pd.read_sql(
            f"""
            SELECT
                [sub_account].[account_code] [SoTaiKhoan],
                SUM([rdo0002].[fee]) [PhiGiaoDichFDS]
            FROM [rdo0002]
            LEFT JOIN [sub_account] ON [sub_account].[sub_account] = [rdo0002].[sub_account]
            WHERE [rdo0002].[date] = '{self.bravoDateString}'
            GROUP BY [sub_account].[account_code]
            """,
            connect_DWH_PhaiSinh
        )
        table = pd.merge(TaiKhoan5115104,RDO0002,how='outer',on='SoTaiKhoan')
        table = table.fillna(0)
        table['PhiGiaoDichDiff'] = table['PhiGiaoDichFDS'] - table['PhiGiaoDichBravo']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDO0002')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('B11')
        worksheet.set_column('A:A',11)
        worksheet.set_column('B:D',15)

        worksheet.merge_range('A1:H1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:H2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:H3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*4,self.empty_row_format)
        worksheet.merge_range('A6:D6','ĐỐI CHIẾU PHÍ GIAO DỊCH',self.sheet_title_format)
        worksheet.merge_range('A7:D7',f'Date: {self.file_date}',self.sub_title_format)
        worksheet.merge_range('A8:D8','Tài khoản: 5115104 & FDS: RDO0002',self.sub_title_format)

        worksheet.write('A10','Số tài khoản',self.headers_root_format)
        worksheet.write('B10','FDS',self.headers_fds_format)
        worksheet.write('C10','Bravo',self.headers_bravo_format)
        worksheet.write('D10','Chênh lệch',self.headers_diff_format)

        worksheet.write_column('A11',table['SoTaiKhoan'],self.text_root_format)
        worksheet.write_column('B11',table['PhiGiaoDichFDS'],self.money_fds_format)
        worksheet.write_column('C11',table['PhiGiaoDichBravo'],self.money_bravo_format)
        worksheet.write_column('D11',table['PhiGiaoDichDiff'],self.money_diff_format)

        worksheet.write(f'A{table.shape[0]+11}','Tổng',self.headers_root_format)
        worksheet.write(f'B{table.shape[0]+11}',f'=SUM(B11:B{table.shape[0]+10})',self.money_fds_format)
        worksheet.write(f'C{table.shape[0]+11}',f'=SUM(C11:C{table.shape[0]+10})',self.money_bravo_format)
        worksheet.write(f'D{table.shape[0]+11}',f'=SUM(D11:D{table.shape[0]+10})',self.money_diff_format)
