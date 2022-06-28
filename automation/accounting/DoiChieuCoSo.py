from os.path import dirname, join
import pandas as pd
from tqdm import tqdm
from automation.accounting import *


# Client code
def run(
    run_time=dt.datetime.now()
):
    report = Report(run_time)
    for func in tqdm([report.run324,report.run1231,report.run13226],ncols=70):
        func()

class Report:

    def __init__(self,run_time:dt.datetime):

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
        self.file_name = f'Đối Chiếu Cơ Sở {self.file_date}.xlsx'
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
        self.headers_flex_format = self.workbook.add_format(
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
        self.money_flex_format = self.workbook.add_format(
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

    def run324(self):

        TaiKhoan324 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Sổ tổng hợp công nợ 324_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            names=['SoTaiKhoan','TenKhachHang','DuDauNoBravo','DuDauCoBravo','PhatSinhGiamBravo','PhatSinhTangBravo','DuCuoiNoBravo','DuCuoiCoBravo'],
        )
        TaiKhoan324['DuDauBravo'] = TaiKhoan324['DuDauCoBravo'] - TaiKhoan324['DuDauNoBravo']
        TaiKhoan324['DuCuoiBravo'] = TaiKhoan324['DuCuoiCoBravo'] - TaiKhoan324['DuCuoiNoBravo']
        RCI0001 = pd.read_sql(
            f"""
            SELECT 
                [sub_account].[account_code] [SoTaiKhoan],
                SUM([sub_account_deposit].[opening_balance]) [DuDauFlex],
                SUM([sub_account_deposit].[increase]) [PhatSinhTangFlex],
                SUM([sub_account_deposit].[decrease]) [PhatSinhGiamFlex],
                SUM([sub_account_deposit].[opening_balance]) [DuCuoiFlex]
            FROM [sub_account_deposit]
            LEFT JOIN [sub_account] ON [sub_account_deposit].[sub_account] = [sub_account].[sub_account]
            WHERE [sub_account_deposit].[date] = '{self.bravoDateString}'
            GROUP BY [sub_account].[account_code]
            """,
            connect_DWH_CoSo
        )
        table = pd.merge(TaiKhoan324,RCI0001,how='outer',on='SoTaiKhoan',)
        table['DuDauDiff'] = table['DuDauBravo'] - table['DuDauFlex']
        table['PhatSinhGiamDiff'] = table['PhatSinhGiamBravo'] - table['PhatSinhGiamFlex']
        table['PhatSinhTangDiff'] = table['PhatSinhTangBravo'] - table['PhatSinhTangFlex']
        table['DuCuoiDiff'] = table['DuCuoiBravo'] - table['DuCuoiFlex']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('324')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('C14')
        worksheet.set_column('A:A',10)
        worksheet.set_column('B:B',26)
        worksheet.set_column('C:P',13)

        worksheet.merge_range('A1:P1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:P2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:P3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*16,self.empty_row_format)
        worksheet.merge_range('A6:P6','ĐỐI CHIẾU SỔ TỔNG HỢP CÔNG NỢ',self.sheet_title_format)
        worksheet.merge_range('A7:P7',f'Date: {self.file_date}',self.sub_title_format)
        worksheet.merge_range('A8:P8','Tài khoản: 324 - Phải trả Nhà đầu tư về tiền gửi giao dịch chứng khoán theo phương thức CTCK quản lý',self.sub_title_format)

        worksheet.merge_range('C10:H10','Bravo',self.headers_bravo_format)
        worksheet.merge_range('I10:L10','Flex',self.headers_flex_format)
        worksheet.merge_range('M10:P10','Đối chiếu',self.headers_diff_format)
        worksheet.merge_range('A11:A12','Mã',self.headers_root_format)
        worksheet.merge_range('B11:B12','Tên khách hàng',self.headers_root_format)
        worksheet.merge_range('C11:D11','Dư đầu',self.headers_bravo_format)
        worksheet.merge_range('E11:F11','Phát sinh',self.headers_bravo_format)
        worksheet.merge_range('G11:H11','Dư cuối',self.headers_bravo_format)
        worksheet.write_row('C12',['Nợ','Có']*3,self.headers_bravo_format)
        worksheet.merge_range('I11:I12','Dư đầu',self.headers_flex_format)
        worksheet.merge_range('J11:J12','Giảm',self.headers_flex_format)
        worksheet.merge_range('K11:K12','Tăng',self.headers_flex_format)
        worksheet.merge_range('L11:L12','Dư cuối',self.headers_flex_format)
        worksheet.merge_range('M11:M12','Dư đầu',self.headers_diff_format)
        worksheet.merge_range('N11:N12','Giảm',self.headers_diff_format)
        worksheet.merge_range('O11:O12','Tăng',self.headers_diff_format)
        worksheet.merge_range('P11:P12','Dư cuối',self.headers_diff_format)
        worksheet.write_row('A13',np.arange(1,3),self.headers_root_format)
        worksheet.write_row('C13',np.arange(3,9),self.headers_bravo_format)
        worksheet.write_row('I13',np.arange(9,13),self.headers_flex_format)
        worksheet.write_row('M13',np.arange(13,17),self.headers_diff_format)

        worksheet.write_column('A14',table['SoTaiKhoan'],self.text_root_format)
        worksheet.write_column('B14',table['TenKhachHang'],self.text_root_format)
        worksheet.write_column('C14',table['DuDauNoBravo'],self.money_bravo_format)
        worksheet.write_column('D14',table['DuDauCoBravo'],self.money_bravo_format)
        worksheet.write_column('E14',table['PhatSinhGiamBravo'],self.money_bravo_format)
        worksheet.write_column('F14',table['PhatSinhTangBravo'],self.money_bravo_format)
        worksheet.write_column('G14',table['DuCuoiNoBravo'],self.money_bravo_format)
        worksheet.write_column('H14',table['DuCuoiCoBravo'],self.money_bravo_format)
        worksheet.write_column('I14',table['DuDauFlex'],self.money_flex_format)
        worksheet.write_column('J14',table['PhatSinhGiamFlex'],self.money_flex_format)
        worksheet.write_column('K14',table['PhatSinhTangFlex'],self.money_flex_format)
        worksheet.write_column('L14',table['DuCuoiFlex'],self.money_flex_format)
        worksheet.write_column('M14',table['DuDauDiff'],self.money_diff_format)
        worksheet.write_column('N14',table['PhatSinhGiamDiff'],self.money_diff_format)
        worksheet.write_column('O14',table['PhatSinhTangDiff'],self.money_diff_format)
        worksheet.write_column('P14',table['DuCuoiDiff'],self.money_diff_format)


    def run1231(self):

        TaiKhoan1231 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Sổ tổng hợp công nợ 1231_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            names=['SoTaiKhoan','TenKhachHangBravo','DuDauNoBravo','DuDauCoBravo','PhatSinhNoBravo','PhatSinhCoBravo','DuCuoiNoBravo','DuCuoiCoBravo'],
        )

        RLN0006 = pd.read_sql(
            f"""
            SELECT
                [margin_outstanding].[account_code] [SoTaiKhoan],
                MAX([account].[customer_name]) [TenKhachHangFlex],
                SUM([margin_outstanding].[principal_outstanding]) [DuCuoiNoFlex]
            FROM [margin_outstanding]
            LEFT JOIN [account] ON [account].[account_code] = [margin_outstanding].[account_code]
            WHERE [margin_outstanding].[date] = '{self.bravoDateString}'
                AND  [margin_outstanding].[type] <> N'Ứng trước cổ tức'
            GROUP BY [margin_outstanding].[account_code]
            """,
            connect_DWH_CoSo
        )
        table = pd.merge(TaiKhoan1231,RLN0006,how='outer',on='SoTaiKhoan')
        table['TenKhachHang'] = table['TenKhachHangBravo'].fillna(table['TenKhachHangFlex'])
        table = table.fillna(0)
        table['DuCuoiNoDiff'] = table['DuCuoiNoBravo'] - table['DuCuoiNoFlex']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('1231')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('C13')
        worksheet.set_column('A:A',10)
        worksheet.set_column('B:B',26)
        worksheet.set_column('C:J',13)

        worksheet.merge_range('A1:J1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:J2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:J3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*10,self.empty_row_format)
        worksheet.merge_range('A6:J6','ĐỐI CHIẾU SỔ TỔNG HỢP CÔNG NỢ',self.sheet_title_format)
        worksheet.merge_range('A7:J7',f'Date: {self.file_date}',self.sub_title_format)
        worksheet.merge_range('A8:J8','Tài khoản: 1231 - Cho vay hoạt động Margin',self.sub_title_format)

        worksheet.merge_range('C10:H10','Bravo',self.headers_bravo_format)
        worksheet.write('I10','Flex',self.headers_flex_format)
        worksheet.merge_range('J10:J11','Chênh lệch',self.headers_diff_format)
        worksheet.write_row('A11',['Mã đối tượng','Tên đối tượng'],self.headers_root_format)
        worksheet.write_row('C11',['Dư nợ đầu','Dư có đầu','Ps nợ','Ps có','Dư nợ cuối','Dư có cuối'],self.headers_bravo_format)
        worksheet.write('I11','Dư nợ gốc',self.headers_flex_format)

        worksheet.write_row('A12',np.arange(1,3),self.headers_root_format)
        worksheet.write_row('C12',np.arange(3,9),self.headers_bravo_format)
        worksheet.write_row('I12',np.arange(9,10),self.headers_flex_format)
        worksheet.write_row('J12',np.arange(10,11),self.headers_diff_format)

        worksheet.write_column('A13',table['SoTaiKhoan'],self.text_root_format)
        worksheet.write_column('B13',table['TenKhachHang'],self.text_root_format)
        worksheet.write_column('C13',table['DuDauNoBravo'],self.money_bravo_format)
        worksheet.write_column('D13',table['DuDauCoBravo'],self.money_bravo_format)
        worksheet.write_column('E13',table['PhatSinhNoBravo'],self.money_bravo_format)
        worksheet.write_column('F13',table['PhatSinhCoBravo'],self.money_bravo_format)
        worksheet.write_column('G13',table['DuCuoiNoBravo'],self.money_bravo_format)
        worksheet.write_column('H13',table['DuCuoiCoBravo'],self.money_bravo_format)
        worksheet.write_column('I13',table['DuCuoiNoFlex'],self.money_flex_format)
        worksheet.write_column('J13',table['DuCuoiNoDiff'],self.money_diff_format)

    def run13226(self):

        TaiKhoan13226 = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Sổ tổng hợp công nợ 13226_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            names=['SoTaiKhoan','TenKhachHangBravo','DuDauNoBravo','DuDauCoBravo','PhatSinhNoBravo','PhatSinhCoBravo','DuCuoiNoBravo','DuCuoiCoBravo'],
        )
        RLN0006 = pd.read_sql(
            f"""
            SELECT
                [margin_outstanding].[account_code] [SoTaiKhoan],
                MAX([account].[customer_name]) [TenKhachHangFlex],
                SUM([margin_outstanding].[interest_outstanding]) [DuCuoiNoFlex]
            FROM [margin_outstanding]
            LEFT JOIN [account] ON [account].[account_code] = [margin_outstanding].[account_code]
            WHERE [margin_outstanding].[date] = '{self.bravoDateString}'
                AND  [margin_outstanding].[type] <> N'Ứng trước cổ tức'
            GROUP BY [margin_outstanding].[account_code]
            """,
            connect_DWH_CoSo
        )
        table = pd.merge(TaiKhoan13226,RLN0006,how='outer',on='SoTaiKhoan')
        table['TenKhachHang'] = table['TenKhachHangBravo'].fillna(table['TenKhachHangFlex'])
        table = table.fillna(0)
        table['DuCuoiNoDiff'] = table['DuCuoiNoBravo'] - table['DuCuoiNoFlex']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('13226')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('C13')
        worksheet.set_column('A:A',10)
        worksheet.set_column('B:B',26)
        worksheet.set_column('C:J',13)

        worksheet.merge_range('A1:J1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:J2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:J3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*10,self.empty_row_format)
        worksheet.merge_range('A6:J6','ĐỐI CHIẾU SỔ TỔNG HỢP CÔNG NỢ',self.sheet_title_format)
        worksheet.merge_range('A7:J7',f'Date: {self.file_date}',self.sub_title_format)
        worksheet.merge_range('A8:J8','Tài khoản: 1231 - Cho vay hoạt động Margin',self.sub_title_format)

        worksheet.merge_range('C10:H10','Bravo',self.headers_bravo_format)
        worksheet.write('I10','Flex',self.headers_flex_format)
        worksheet.merge_range('J10:J11','Chênh lệch',self.headers_diff_format)
        worksheet.write_row('A11',['Mã đối tượng','Tên đối tượng'],self.headers_root_format)
        worksheet.write_row('C11',['Dư nợ đầu','Dư có đầu','Ps nợ','Ps có','Dư nợ cuối','Dư có cuối'],self.headers_bravo_format)
        worksheet.write('I11','Dư nợ gốc',self.headers_flex_format)

        worksheet.write_row('A12',np.arange(1,3),self.headers_root_format)
        worksheet.write_row('C12',np.arange(3,9),self.headers_bravo_format)
        worksheet.write_row('I12',np.arange(9,10),self.headers_flex_format)
        worksheet.write_row('J12',np.arange(10,11),self.headers_diff_format)

        worksheet.write_column('A13',table['SoTaiKhoan'],self.text_root_format)
        worksheet.write_column('B13',table['TenKhachHang'],self.text_root_format)
        worksheet.write_column('C13',table['DuDauNoBravo'],self.money_bravo_format)
        worksheet.write_column('D13',table['DuDauCoBravo'],self.money_bravo_format)
        worksheet.write_column('E13',table['PhatSinhNoBravo'],self.money_bravo_format)
        worksheet.write_column('F13',table['PhatSinhCoBravo'],self.money_bravo_format)
        worksheet.write_column('G13',table['DuCuoiNoBravo'],self.money_bravo_format)
        worksheet.write_column('H13',table['DuCuoiCoBravo'],self.money_bravo_format)
        worksheet.write_column('I13',table['DuCuoiNoFlex'],self.money_flex_format)
        worksheet.write_column('J13',table['DuCuoiNoDiff'],self.money_diff_format)
