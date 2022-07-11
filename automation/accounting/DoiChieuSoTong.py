import pandas as pd
from tqdm import tqdm
from automation.accounting import *
import math


# Client code
def run(
    run_time=dt.datetime.now()
):
    report = Report(run_time)
    for func in tqdm([report.runPhaiSinh,report.runCoSo,report.runTuDoanhLoLe],ncols=70):
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

        self.file_name = f'Đối Chiếu Số Tổng {self.file_date}.xlsx'
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
                'align':'left',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Times New Roman',
            }
        )
        self.number_bravo_format = self.workbook.add_format(
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
        self.number_system_format = self.workbook.add_format(
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
        self.number_diff_format = self.workbook.add_format(
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

    def runPhaiSinh(self):

        BangCanDoiSoPhatSinhBravo = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Bảng cân đối số phát sinh {self.file_date}.xls',),
            skiprows=10,
            skipfooter=1,
            usecols='A:B,E:H',
            names=['SoHieuTaiKhoan','TaiKhoanKeToan','TrongKyNo','TrongKyCo','CuoiKyNo','CuoiKyCo'],
            dtype={
                'SoHieuTaiKhoan': object,
                'TaiKhoanKeToan': object,
                'TrongKyNo': np.int64,
                'TrongKyCo': np.int64,
                'CuoiKyNo': np.int64,
                'CuoiKyCo': np.int64,
            }
        )
        accountingItems = (
            '13504','13505','3214','3243','324301','324302',
            '32682','119','1191','33353','13513',
        )
        BangCanDoiSoPhatSinhBravo = BangCanDoiSoPhatSinhBravo.loc[BangCanDoiSoPhatSinhBravo['SoHieuTaiKhoan'].isin(accountingItems)]
        BangCanDoiSoPhatSinhFDS = pd.read_sql(
            f"""
            DECLARE @Date DATETIME
            SET @Date = '2022-01-25';
            
            SELECT 
                '13504' [SoHieuTaiKhoan],
                SUM([deferred_payment_amount_increase]) [TrongKyNo],
                SUM([deferred_payment_amount_decrease]) [TrongKyCo],
                SUM([deferred_payment_amount_closing]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [rdt0141]
            WHERE [date] = @Date
            UNION ALL
            SELECT 
                '13505' [SoHieuTaiKhoan],
                SUM([deferred_payment_fee_increase]) [TrongKyNo],
                SUM([deferred_payment_fee_decrease]) [TrongKyCo],
                SUM(deferred_payment_fee_closing) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [rdt0141]
            WHERE [date] = @Date
            UNION ALL
            SELECT 
                '3214' [SoHieuTaiKhoan],
                SUM([VMLai]) [TrongKyNo],
                SUM([VMLo]) [TrongKyCo],
                NULL [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [RDT0117]
            WHERE [Ngay] = @Date
            UNION ALL
            SELECT
                '3243' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([cash_balance_at_phs]) [CuoiKyCo]
            FROM [rdt0121]
            WHERE [date] = @Date
            UNION ALL
            SELECT
                '324301' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([cash_balance_at_phs]) [CuoiKyCo]
            FROM [rdt0121]
            WHERE [date] = @Date AND [account_code] LIKE '022C%'
            UNION ALL
            SELECT
                '324302' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([cash_balance_at_phs]) [CuoiKyCo]
            FROM [rdt0121]
            WHERE [date] = @Date AND [account_code] LIKE '022F%'
            UNION ALL
            SELECT
                '32682' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([VMLai]) [CuoiKyCo]
            FROM [RDT0117]
            WHERE [Ngay] = @Date
            UNION ALL
            SELECT
                '119' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([cash_balance_at_vsd]) [CuoiKyNo],
                NULL [CuoiKyNo]
            FROM [rdt0121]
            WHERE [date] = @Date
            UNION ALL
            SELECT
                '1191' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([cash_balance_at_vsd]) [CuoiKyNo],
                NULL [CuoiKyNo]
            FROM [rdt0121]
            WHERE [date] = @Date
            UNION ALL
            SELECT
                '33353' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([ThueTNCN]) [CuoiKyCo]
            FROM [RDT0127]
            WHERE [Ngay] = @Date
            UNION ALL
            SELECT
                '13513' [SoHieuTaiKhoan],
                SUM([fee]) [TrongKyNo],
                SUM([fee]) [TrongKyCo],
                NULL [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [rdo0002]
            WHERE [date] = @Date
            """,
            connect_DWH_PhaiSinh
        )
        table = pd.merge(BangCanDoiSoPhatSinhBravo,BangCanDoiSoPhatSinhFDS,how='outer',on='SoHieuTaiKhoan',suffixes=['Bravo','FDS']).fillna('')
        table['TrongKyNoDiff'] = ''
        table['TrongKyCoDiff'] = ''
        table['CuoiKyNoDiff'] = ''
        table['CuoiKyCoDiff'] = ''
        for colPrefix in ('TrongKyNo','TrongKyCo','CuoiKyNo','CuoiKyCo'):
            for item in accountingItems:
                fdsValue = table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}FDS'].squeeze()
                if fdsValue == '':
                    table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}Bravo'] = ''
                else:
                    bravoValue = table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}Bravo'].squeeze()
                    if bravoValue != fdsValue:
                        table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}Diff'] = bravoValue - fdsValue


        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('PhaiSinh')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('C12')
        worksheet.set_column('A:A',8)
        worksheet.set_column('B:B',55)
        worksheet.set_column('C:N',14)

        worksheet.merge_range('A1:N1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:N2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:N3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*14,self.empty_row_format)
        worksheet.merge_range('A6:N6','ĐỐI CHIẾU PHÁI SINH',self.sheet_title_format)
        worksheet.merge_range('A7:N7',f'Date: {self.file_date}',self.sub_title_format)

        worksheet.merge_range('C9:F9','Bravo',self.headers_bravo_format)
        worksheet.merge_range('G9:J9','FDS',self.headers_fds_format)
        worksheet.merge_range('K9:N9','Đối chiếu (Bravo - FDS)',self.headers_diff_format)

        worksheet.merge_range('A9:A11','Số hiệu tài khoản',self.headers_root_format)
        worksheet.merge_range('B9:B11','Tài tài khoản kế toán',self.headers_root_format)
        worksheet.merge_range('C10:D10','Phát sinh trong kỳ',self.headers_bravo_format)
        worksheet.merge_range('E10:F10','Số dư cuối kỳ',self.headers_bravo_format)
        worksheet.write_row('C11',['Nợ','Có']*2,self.headers_bravo_format)
        worksheet.write_row('A12',['A','B'],self.headers_root_format)
        worksheet.write_row('C12',['1','2','3','4'],self.headers_bravo_format)
        worksheet.write('K12','C',self.headers_diff_format)

        worksheet.merge_range('G10:H10','Phát sinh trong kỳ',self.headers_fds_format)
        worksheet.merge_range('I10:J10','Số dư cuối kỳ',self.headers_fds_format)
        worksheet.write_row('G11',['Nợ','Có']*2,self.headers_fds_format)
        worksheet.write_row('G12',['1','2','3','4'],self.headers_fds_format)

        worksheet.merge_range('K10:L10','Phát sinh trong kỳ',self.headers_diff_format)
        worksheet.merge_range('M10:N10','Số dư cuối kỳ',self.headers_diff_format)
        worksheet.write_row('K11',['Nợ','Có']*2,self.headers_diff_format)
        worksheet.write_row('K12',['1','2','3','4'],self.headers_diff_format)

        worksheet.write_column('A13',table['SoHieuTaiKhoan'],self.text_root_format)
        worksheet.write_column('B13',table['TaiKhoanKeToan'],self.text_root_format)
        worksheet.write_column('C13',table['TrongKyNoBravo'],self.number_bravo_format)
        worksheet.write_column('D13',table['TrongKyCoBravo'],self.number_bravo_format)
        worksheet.write_column('E13',table['CuoiKyNoBravo'],self.number_bravo_format)
        worksheet.write_column('F13',table['CuoiKyCoBravo'],self.number_bravo_format)
        worksheet.write_column('G13',table['TrongKyNoFDS'],self.number_system_format)
        worksheet.write_column('H13',table['TrongKyCoFDS'],self.number_system_format)
        worksheet.write_column('I13',table['CuoiKyNoFDS'],self.number_system_format)
        worksheet.write_column('J13',table['CuoiKyCoFDS'],self.number_system_format)
        worksheet.write_column('K13',table['TrongKyNoDiff'],self.number_diff_format)
        worksheet.write_column('L13',table['TrongKyCoDiff'],self.number_diff_format)
        worksheet.write_column('M13',table['CuoiKyNoDiff'],self.number_diff_format)
        worksheet.write_column('N13',table['CuoiKyCoDiff'],self.number_diff_format)


    def runCoSo(self):

        BangCanDoiSoPhatSinhBravo = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Bảng cân đối số phát sinh {self.file_date}.xls',),
            skiprows=10,
            skipfooter=1,
            usecols='A:B,E:H',
            names=['SoHieuTaiKhoan','TaiKhoanKeToan','TrongKyNo','TrongKyCo','CuoiKyNo','CuoiKyCo'],
            dtype={
                'SoHieuTaiKhoan': object,
                'TaiKhoanKeToan': object,
                'TrongKyNo': np.int64,
                'TrongKyCo': np.int64,
                'CuoiKyNo': np.int64,
                'CuoiKyCo': np.int64,
            }
        )
        accountingItems = (
            '12311','12313','12314','12321','1322601',
            '1322602','1322603','13511','13541','13571',
            '3211','3212','3213','3241','3242','32681',
        )
        BangCanDoiSoPhatSinhBravo = BangCanDoiSoPhatSinhBravo.loc[BangCanDoiSoPhatSinhBravo['SoHieuTaiKhoan'].isin(accountingItems)]
        BangCanDoiSoPhatSinhFlex = pd.read_sql(
            f"""
            DECLARE @T0 DATETIME;
            SET @T0 = '2022-01-25';
            
            DECLARE @T1 DATETIME;
            SET @T1 = '2022-01-24';
            
            DECLARE @BuyValueP DECIMAL(20,2);
            SET @BuyValueP = (
                SELECT ISNULL(SUM([value]),0) 
                FROM [trading_record] 
                WHERE [date] IN (@T0,@T1) 
                    AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022P%') 
                    AND [type_of_order] = 'B'
                );
            
            DECLARE @SellValueP DECIMAL(20,2);
            SET @SellValueP = (
                SELECT ISNULL(SUM([value]),0)
                FROM [trading_record] 
                WHERE [date] IN (@T0,@T1) 
                    AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022P%') 
                    AND [type_of_order] = 'S'
                );
            
            DECLARE @BuyValueC DECIMAL(20,2);
            SET @BuyValueC = (
                SELECT ISNULL(SUM([value]),0) 
                FROM [trading_record] 
                WHERE [date] IN (@T0,@T1) 
                    AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022C%') 
                    AND [type_of_order] = 'B'
                );
            
            DECLARE @SellValueC DECIMAL(20,2);
            SET @SellValueC = (
                SELECT ISNULL(SUM([value]),0)
                FROM [trading_record] 
                WHERE [date] IN (@T0,@T1) 
                    AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022C%') 
                    AND [type_of_order] = 'S'
                );
            
            DECLARE @BuyValueF DECIMAL(20,2);
            SET @BuyValueF = (
                SELECT ISNULL(SUM([value]),0) 
                FROM [trading_record] 
                WHERE [date] IN (@T0,@T1) 
                    AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022F%') 
                    AND [type_of_order] = 'B'
                );
            
            DECLARE @SellValueF DECIMAL(20,2);
            SET @SellValueF = (
                SELECT ISNULL(SUM([value]),0)
                FROM [trading_record] 
                WHERE [date] IN (@T0,@T1) 
                    AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022F%') 
                    AND [type_of_order] = 'S'
                );
            
            SELECT
                '12311' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([principal_outstanding]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [margin_outstanding]
            WHERE [date] = @T0 AND [type] = N'Margin'
                AND [account_code] LIKE '022[CF]%'
            UNION ALL
            SELECT 
                '12313' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([principal_outstanding]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [margin_outstanding]
            WHERE [date] = @T0 AND [type] = N'Trả chậm'
                AND [account_code] LIKE '022[CF]%'
            UNION ALL
            SELECT 
                '12314' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([principal_outstanding]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [margin_outstanding]
            WHERE [date] = @T0 AND [type] = N'Bảo lãnh' 
                AND [account_code] LIKE '022[CF]%'
            UNION ALL
            SELECT 
                '12321' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([receivable]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [payment_in_advance]
            WHERE [date] IN (@T0, @T1)
                AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022[CF]%') 
            UNION ALL
            SELECT 
                '1322601' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([interest_outstanding]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [margin_outstanding]
            WHERE [date] = @T0 AND [type] = N'Margin'
                AND [account_code] LIKE '022[CF]%'
            UNION ALL
            SELECT 
                '1322602' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([interest_outstanding]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [margin_outstanding]
            WHERE [date] = @T0 AND [type] = N'Trả chậm'
                AND [account_code] LIKE '022[CF]%'
            UNION ALL
            SELECT 
                '1322603' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([interest_outstanding]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [margin_outstanding]
            WHERE [date] = @T0 AND [type] = N'Bảo lãnh'
                AND [account_code] LIKE '022[CF]%'
            UNION ALL
            SELECT 
                '13511' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([fee]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [trading_record]
            WHERE [date] IN (@T0,@T1) AND [type_of_order] = 'S'
                AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022[CF]%')
            UNION ALL
            SELECT 
                '13571' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                SUM([principal_outstanding]) + SUM([interest_outstanding]) + SUM([fee_outstanding]) [CuoiKyNo],
                NULL [CuoiKyCo]
            FROM [margin_outstanding]
            WHERE [date] = @T0 AND [type] = N'Ứng trước cổ tức' AND [account_code] LIKE '022[CF]%'
            UNION ALL 
            SELECT
                '3211' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                @BuyValueP - @SellValueP [CuoiKyCo]
            UNION ALL 
            SELECT
                '3212' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                @BuyValueC - @SellValueC [CuoiKyCo]
            UNION ALL 
            SELECT
                '3213' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                @BuyValueF - @SellValueF [CuoiKyCo]
            UNION ALL
            SELECT 
                '3241' [SoHieuTaiKhoan],
                SUM([decrease]) [TrongKyNo],
                SUM([increase]) [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([closing_balance]) [CuoiKyCo]
            FROM [sub_account_deposit]
            WHERE [date] = @T0 AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022C%') 
            UNION ALL
            SELECT 
                '3242' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([closing_balance]) [CuoiKyCo]
            FROM [sub_account_deposit]
            WHERE [date] = @T0 AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022F%') 
            UNION ALL
            SELECT 
                '32681' [SoHieuTaiKhoan],
                NULL [TrongKyNo],
                NULL [TrongKyCo],
                NULL [CuoiKyNo],
                SUM([value]) [CuoiKyCo]
            FROM [trading_record]
            WHERE [type_of_order] = 'S' 
                AND [date] IN (@T0,@T1) 
                AND [sub_account] IN (SELECT [sub_account] FROM [sub_account] WHERE [account_code] LIKE '022[CF]%');
            """,
            connect_DWH_CoSo
        )
        table = pd.merge(BangCanDoiSoPhatSinhBravo,BangCanDoiSoPhatSinhFlex,how='outer',on='SoHieuTaiKhoan',suffixes=['Bravo','Flex']).fillna('')
        table['TrongKyNoDiff'] = ''
        table['TrongKyCoDiff'] = ''
        table['CuoiKyNoDiff'] = ''
        table['CuoiKyCoDiff'] = ''
        for colPrefix in ('TrongKyNo','TrongKyCo','CuoiKyNo','CuoiKyCo'):
            for item in accountingItems:
                fdsValue = table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}Flex'].squeeze()
                if fdsValue == '':
                    table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}Bravo'] = ''
                else:
                    bravoValue = table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}Bravo'].squeeze()
                    if bravoValue != fdsValue:
                        table.loc[table['SoHieuTaiKhoan']==item,f'{colPrefix}Diff'] = bravoValue - fdsValue

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('CoSo')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('C12')
        worksheet.set_column('A:A',8)
        worksheet.set_column('B:B',75)
        worksheet.set_column('C:N',14)

        worksheet.merge_range('A1:N1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:N2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:N3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*14,self.empty_row_format)
        worksheet.merge_range('A6:N6','ĐỐI CHIẾU CƠ SỞ',self.sheet_title_format)
        worksheet.merge_range('A7:N7',f'Date: {self.file_date}',self.sub_title_format)

        worksheet.merge_range('C9:F9','Bravo',self.headers_bravo_format)
        worksheet.merge_range('G9:J9','Flex',self.headers_fds_format)
        worksheet.merge_range('K9:N9','Đối chiếu (Bravo - Flex)',self.headers_diff_format)

        worksheet.merge_range('A9:A11','Số hiệu tài khoản',self.headers_root_format)
        worksheet.merge_range('B9:B11','Tài tài khoản kế toán',self.headers_root_format)
        worksheet.merge_range('C10:D10','Phát sinh trong kỳ',self.headers_bravo_format)
        worksheet.merge_range('E10:F10','Số dư cuối kỳ',self.headers_bravo_format)
        worksheet.write_row('C11',['Nợ','Có']*2,self.headers_bravo_format)
        worksheet.write_row('A12',['A','B'],self.headers_root_format)
        worksheet.write_row('C12',['1','2','3','4'],self.headers_bravo_format)
        worksheet.write('K12','C',self.headers_diff_format)

        worksheet.merge_range('G10:H10','Phát sinh trong kỳ',self.headers_fds_format)
        worksheet.merge_range('I10:J10','Số dư cuối kỳ',self.headers_fds_format)
        worksheet.write_row('G11',['Nợ','Có']*2,self.headers_fds_format)
        worksheet.write_row('G12',['1','2','3','4'],self.headers_fds_format)

        worksheet.merge_range('K10:L10','Phát sinh trong kỳ',self.headers_diff_format)
        worksheet.merge_range('M10:N10','Số dư cuối kỳ',self.headers_diff_format)
        worksheet.write_row('K11',['Nợ','Có']*2,self.headers_diff_format)
        worksheet.write_row('K12',['1','2','3','4'],self.headers_diff_format)

        worksheet.write_column('A13',table['SoHieuTaiKhoan'],self.text_root_format)
        worksheet.write_column('B13',table['TaiKhoanKeToan'],self.text_root_format)
        worksheet.write_column('C13',table['TrongKyNoBravo'],self.number_bravo_format)
        worksheet.write_column('D13',table['TrongKyCoBravo'],self.number_bravo_format)
        worksheet.write_column('E13',table['CuoiKyNoBravo'],self.number_bravo_format)
        worksheet.write_column('F13',table['CuoiKyCoBravo'],self.number_bravo_format)
        worksheet.write_column('G13',table['TrongKyNoFlex'],self.number_system_format)
        worksheet.write_column('H13',table['TrongKyCoFlex'],self.number_system_format)
        worksheet.write_column('I13',table['CuoiKyNoFlex'],self.number_system_format)
        worksheet.write_column('J13',table['CuoiKyCoFlex'],self.number_system_format)
        worksheet.write_column('K13',table['TrongKyNoDiff'],self.number_diff_format)
        worksheet.write_column('L13',table['TrongKyCoDiff'],self.number_diff_format)
        worksheet.write_column('M13',table['CuoiKyNoDiff'],self.number_diff_format)
        worksheet.write_column('N13',table['CuoiKyCoDiff'],self.number_diff_format)

    def runTuDoanhLoLe(self):

        BangCanDoiSoPhatSinhBravo = pd.read_excel(
            join(self.bravoFolder,f'{self.bravoDateString}',f'Báo cáo trích lập dự phòng (PP2).xlsx',),
            skiprows=6,
            skipfooter=1,
            usecols='A,B,O',
            names=['MaChungKhoan','Loai','SoLuongChungKhoanBravo'],
        )
        BangCanDoiSoPhatSinhBravo['MaChungKhoan'] = BangCanDoiSoPhatSinhBravo['MaChungKhoan'].str.strip()
        BangCanDoiSoPhatSinhBravo['Loai'] = BangCanDoiSoPhatSinhBravo['Loai'].str.strip()
        BangCanDoiSoPhatSinhBravo.loc[~BangCanDoiSoPhatSinhBravo['Loai'].isin(['CKTD','CKLL']),'Loai'] = np.nan
        BangCanDoiSoPhatSinhBravo['Loai'] = BangCanDoiSoPhatSinhBravo['Loai'].fillna(method='ffill')
        BangCanDoiSoPhatSinhBravo = BangCanDoiSoPhatSinhBravo.dropna(how='any')
        BangCanDoiSoPhatSinhFlex = pd.read_sql(
            """
            SELECT 
                [MaCK] [MaChungKhoan],
                CASE [SoTieuKhoan] 
                    WHEN '0001920028' THEN 'CKLL'
                    WHEN '0001002222' THEN 'CKTD'
                    ELSE ''
                END [Loai],
                SUM([GiaoDich]) [SoLuongChungKhoanFlex] 
            FROM [RSE0008]
            WHERE [Ngay] = '2022-01-25'
                AND [SoTieuKhoan] IN ('0001920028','0001002222')
            GROUP BY [MaCK], [SoTieuKhoan]
            """,
            connect_DWH_CoSo
        )
        table = pd.merge(BangCanDoiSoPhatSinhBravo,BangCanDoiSoPhatSinhFlex,how='outer',on=['MaChungKhoan','Loai']).fillna(0)
        table['SoLuongChungKhoanDiff'] = table['SoLuongChungKhoanBravo'] - table['SoLuongChungKhoanFlex']
        table = table.sort_values('SoLuongChungKhoanDiff',key=lambda series: series.abs(),ascending=False)
        summary = table.groupby('Loai')[['SoLuongChungKhoanBravo','SoLuongChungKhoanFlex','SoLuongChungKhoanDiff']].agg(np.sum)

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('TuDoanhLoLe')
        worksheet.hide_gridlines(option=2)
        worksheet.set_column('A:D',17)

        worksheet.merge_range('A1:D1',CompanyName,self.company_name_format)
        worksheet.merge_range('A2:D2',CompanyAddress,self.company_info_format)
        worksheet.merge_range('A3:D3',CompanyPhoneNumber,self.company_info_format)
        worksheet.write_row('A4',['']*4,self.empty_row_format)
        worksheet.merge_range('A6:D6','ĐỐI CHIẾU SỐ LƯỢNG CHỨNG KHOÁN TỰ DOANH & LÔ LẺ',self.sheet_title_format)
        worksheet.merge_range('A7:D7',f'Date: {self.file_date}',self.sub_title_format)

        worksheet.write('A9','Đối chiếu số tổng:',self.sub_title_format)
        worksheet.write('A10','Loại',self.headers_root_format)
        worksheet.write('B10','Bravo',self.headers_bravo_format)
        worksheet.write('C10','Flex',self.headers_fds_format)
        worksheet.write('D10','Đối chiếu\n(Bravo - Flex)',self.headers_diff_format)

        worksheet.write_column('A11',['Kho lô lẻ','Kho tự doanh','Tổng'],self.text_root_format)
        worksheet.write_column('B11',summary['SoLuongChungKhoanBravo'],self.number_bravo_format)
        worksheet.write_column('C11',summary['SoLuongChungKhoanFlex'],self.number_system_format)
        worksheet.write_column('D11',summary['SoLuongChungKhoanDiff'],self.number_diff_format)
        worksheet.write('B13','=ABS(B11)+ABS(B12)',self.number_bravo_format)
        worksheet.write('C13','=ABS(C11)+ABS(C12)',self.number_system_format)
        worksheet.write('D13','=ABS(D11)+ABS(D12)',self.number_diff_format)

        worksheet.write('A15','Đối chiếu chi tiết:',self.sub_title_format)
        worksheet.write('A16','Mã Chứng Khoán',self.headers_root_format)
        worksheet.write('B16','Bravo',self.headers_bravo_format)
        worksheet.write('C16','Flex',self.headers_fds_format)
        worksheet.write('D16','Đối chiếu\n(Bravo - Flex)',self.headers_diff_format)

        worksheet.write_column('A17',table['MaChungKhoan'],self.text_root_format)
        worksheet.write_column('B17',table['SoLuongChungKhoanBravo'],self.number_bravo_format)
        worksheet.write_column('C17',table['SoLuongChungKhoanFlex'],self.number_system_format)
        worksheet.write_column('D17',table['SoLuongChungKhoanDiff'],self.number_diff_format)

        

