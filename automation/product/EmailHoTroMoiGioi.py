from automation.product import *
from abc import ABC, abstractmethod
from datawarehouse import BDATE

class AbstractTable(ABC):

    def __init__(self,runTime):
        info = get_info('daily',runTime)
        self.period = info['period']
        self.t0_date = info['end_date']
        self.t1_date = BDATE(self.t0_date,-1)
        self.t2_date = BDATE(self.t0_date,-2)
        self.folder_name = 'EmailHoTroMoiGioi'
        self.resultTable = None

        # đọc từ file cố định, phải cập nhật định kỳ
        self.emailTable = pd.read_excel(
            join(dirname(__file__),'file','DanhSachNhanVien.xlsx'),
            dtype={'employeeCode':str,'email':str}
        )
        self.emailTable['email'] = self.emailTable['email'].str.replace('[^a-z@phs\.vn]','',regex=True)

        # create folder
        self.exportPath = join(dept_folder,self.folder_name,self.period)
        if not os.path.isdir(self.exportPath):
            os.mkdir(self.exportPath)

    @abstractmethod
    def createTable(self):
        pass

    def findEmail(self,dataTable):
        """
        dataTable phải có cột 'BrokerID', 'EmailFromFlex'
        """
        table = pd.merge(dataTable,self.emailTable,left_on='BrokerID',right_on='employeeCode',how='left')
        table = table.drop('employeeCode',axis=1).rename({'email':'EmailFromHR'},axis=1)
        # Ưu tiên lấy mail từ HR trước, không có thì lấy mail từ Flex
        table['Email'] = table['EmailFromHR']
        table['Email'] = table['Email'].fillna(table['EmailFromFlex'])
        # Bỏ các dòng không xác định được email
        table = table.loc[table['Email'].notnull()]
        return table


class ThongTinTaiKhoanThayDoiGiaTriGiaoDich(AbstractTable):

    def __init__(self,runTime):
        super().__init__(runTime)

    def createTable(self):
        mainTable = pd.read_sql(
            f"""
            WITH
            [Portfolio] AS (
                SELECT 
                    [SoTieuKhoan],
                    SUM([GiaVon]) [GiaVon],
                    SUM([GiaTriThiTruong]) [GiaTriChungKhoan],
                    SUM([LaiLoDuTinh]) [LaiLoDuTinh]
                FROM [VSE9985]
                WHERE [Date] = '{self.t0_date}'
                GROUP BY [SoTieuKhoan]
            ),
            [AssetT0] AS (
                SELECT
                    [SubAccount],
                    SUM([TotalCash]) [TongTien],
                    SUM([TotalAssetValue]) [TaiSanRong]
                FROM [RMR0063]
                WHERE [Date] ='{self.t0_date}'
                GROUP BY [SubAccount]
            ),
            [AssetT1] AS (
                SELECT
                    [SubAccount],
                    SUM([TotalAssetValue]) [TaiSanRong]
                FROM [RMR0063]
                WHERE [Date] = '{self.t1_date}'
                GROUP BY [SubAccount]
            )
            SELECT
                [branch].[branch_name] [TenChiNhanh],
                [broker].[broker_id] [BrokerID],
                [broker].[email] [EmailFromFlex],
                [broker].[broker_name] [TenMoiGioi],
                [relationship].[account_code] [SoTaiKhoan],
                [AssetT0].[SubAccount] [SoTieuKhoan],
                [account].[customer_name] [TenKhachHang],
                [AssetT0].[TongTien],
                [Portfolio].[GiaTriChungKhoan],
                [AssetT0].[TongTien] / [Portfolio].[GiaTriChungKhoan] [Tien/GiaTriChungKhoan],
                [AssetT1].[TaiSanRong] [TaiSanRongT1],
                [AssetT0].[TaiSanRong] [TaiSanRongT0],
                [AssetT0].[TaiSanRong] / [AssetT1].[TaiSanRong] - 1 [PhanTramThayDoiTaiSan],
                [Portfolio].[LaiLoDuTinh],
                [Portfolio].[GiaVon],
                [Portfolio].[LaiLoDuTinh] / [Portfolio].[GiaVon] [PhanTramLaiLoDuTinh]
            FROM [AssetT0]
            LEFT JOIN [Portfolio] ON [Portfolio].[SoTieuKhoan] = [AssetT0].[SubAccount]
            LEFT JOIN [AssetT1] ON [AssetT1].[SubAccount] = [AssetT0].[SubAccount]
            LEFT JOIN [relationship] ON [relationship].[sub_account] = [AssetT0].[SubAccount] AND [relationship].[date] = '{self.t0_date}'
            LEFT JOIN [account] ON [account].[account_code] = [relationship].[account_code]
            LEFT JOIN [broker] ON [broker].[broker_id] = [relationship].[broker_id]
            LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
            WHERE [Portfolio].[GiaTriChungKhoan] <> 0 
                AND [AssetT1].[TaiSanRong] <> 0 AND [Portfolio].[GiaVon] <> 0
                AND [broker].[broker_id] LIKE 'A%'
            """,
            connect_DWH_CoSo
        )
        mainTable['TenKhachHang'] = mainTable['TenKhachHang'].str.title()
        # Lấy dữ liệu giá
        market = pd.read_sql(
            f"""
            SELECT
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker],
                [DWH-ThiTruong].[dbo].[DanhSachMa].[Exchange],
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ref] * 1000 [RefPrice],
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Close] * 1000 [ClosePrice]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay] 
            LEFT JOIN [DWH-ThiTruong].[dbo].[DanhSachMa] 
                ON [DWH-ThiTruong].[dbo].[DanhSachMa].[Ticker] = [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker]
            WHERE [DWH-ThiTruong].[dbo].[DanhSachMa].[Date] = '{self.t0_date}'
                AND [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date] IN ('{self.t0_date}','{self.t1_date}','{self.t2_date}')
                AND [DWH-ThiTruong].[dbo].[DanhSachMa].[Exchange] IN ('HOSE','HNX','UPCOM')
            """,
            connect_DWH_CoSo
        )
        # Lấy danh sách cổ phiếu KH đang nắm giữ
        holdings = pd.read_sql(
            f"""
            WITH
            [RawData] AS (
                SELECT DISTINCT
                    [Date],
                    [SoTieuKhoan],
                    CASE CHARINDEX('_',[ChungKhoan])
                        WHEN 0 THEN [ChungKhoan]
                        ELSE SUBSTRING([ChungKhoan],1,CHARINDEX('_',[ChungKhoan])-1)
                    END [MaChungKhoan]
                FROM [VSE9985]
                WHERE [Date] IN ('{self.t0_date}','{self.t1_date}','{self.t2_date}')
            )
            SELECT [SoTieuKhoan], [MaChungKhoan], [NamGiuDu3Ngay] FROM (
                SELECT
                    [Date],
                    [SoTieuKhoan],
                    [MaChungKhoan],
                    CASE WHEN [Date] = '{self.t0_date}' THEN 1 ELSE 0 END [DangNamGiu],
                    CASE WHEN COUNT([Date]) OVER(PARTITION BY [SoTieuKhoan], [MaChungKhoan]) = 3 THEN 1 ELSE 0 END [NamGiuDu3Ngay]
                FROM [RawData]
            ) [x] WHERE [x].[DangNamGiu] = 1
            """,
            connect_DWH_CoSo,
        )
        holdings['NamGiuDu3Ngay'] = holdings['NamGiuDu3Ngay'].astype(bool)
        # Tìm danh sách các mã giảm sàn 3 phiên liên tiếp
        tickers = holdings['MaChungKhoan'].unique()
        market = market.loc[market['Ticker'].isin(tickers)]
        # VSE9985 chỉ bắt đầu lưu từ 19/05/2022, chạy hàm run() trước ngày 19/5/2022 bảng holdings sẽ rỗng -> lỗi tại đây
        market['FloorPrice'] = market.apply(
            lambda x: fc_price(x['RefPrice'],'floor',x['Exchange']),
            axis=1,
        )
        market['FloorCheck'] = market['FloorPrice'] == market['ClosePrice']
        floorCount = market.groupby('Ticker')['FloorCheck'].sum()
        floorTickers = floorCount.loc[floorCount==3].index.to_list()
        # Đánh dấu các mã giảm sàn 3 phiên liên tiếp
        holdings['GiamSan3PhienLienTiep'] = holdings['MaChungKhoan'].isin(floorTickers)
        # Tạo cột MaChungKhoanDanhMuc
        portfolio = holdings.groupby(['SoTieuKhoan'])['MaChungKhoan'].agg(lambda x: ','.join(x))
        # Tạo cột MaChungKhoanGiamSan3Phien
        floor3DaysHoldings = holdings.loc[(holdings['NamGiuDu3Ngay']==True)&(holdings['GiamSan3PhienLienTiep']==True)]
        floor3DaysHoldings = floor3DaysHoldings.groupby('SoTieuKhoan')['MaChungKhoan'].agg(lambda x: ','.join(x))
        # Ghép MaChungKhoanGiamSan3Phien vào MaChungKhoanDanhMuc
        subTable = pd.merge(portfolio,floor3DaysHoldings,how='left',on='SoTieuKhoan',suffixes=('DanhMuc','GiamSan3Phien')).reset_index()
        subTable['MaChungKhoanDanhMuc'] = subTable['MaChungKhoanDanhMuc'].fillna('')
        subTable['MaChungKhoanGiamSan3Phien'] = subTable['MaChungKhoanGiamSan3Phien'].fillna('')
        # Tính PhanTramMaGiamSan3Phien
        countTicker = lambda x: len([e for e in x.split(',') if e != ''])
        SoMaGiamSan3Phien = subTable['MaChungKhoanGiamSan3Phien'].astype(str).map(countTicker)
        SoMaTrongDanhMuc = subTable['MaChungKhoanDanhMuc'].astype(str).map(countTicker)
        subTable['PhanTramMaGiamSan3Phien'] = SoMaGiamSan3Phien / SoMaTrongDanhMuc

        table = pd.merge(mainTable,subTable,how='left',on='SoTieuKhoan')
        condition1 = table['Tien/GiaTriChungKhoan'] >= 1
        condition2 = table['PhanTramThayDoiTaiSan'] <= -0.25
        condition3 = table['PhanTramLaiLoDuTinh'] <= -0.5
        condition4 = table['PhanTramMaGiamSan3Phien'] >= 0.5
        table = table.loc[condition1 | condition2 | condition3 | condition4]
        self.resultTable = self.findEmail(table)

        return self.resultTable


class ThongTinTieuKhoanDong(AbstractTable):

    def __init__(self,runTime):
        super().__init__(runTime)

    def createTable(self):
        table = pd.read_sql(
            f"""
            SELECT 
                [branch].[branch_name] [TenChiNhanh],
                [broker].[broker_id] [BrokerID],
                [broker].[email] [EmailFromFlex],
                [broker].[broker_name] [TenMoiGioi],
                [relationship].[account_code] [SoTaiKhoan],
                [account].[customer_name] [TenKhachHang],
                [RCF0002].[SoTieuKhoan] [SoTieuKhoanDong],
                CASE
                    WHEN [vcf0051].[contract_type] LIKE N'MR%'
                        THEN 'Margin'
                    ELSE N'Thường'
                END [LoaiHinh]
            FROM [RCF0002]
            LEFT JOIN [relationship]
                ON [relationship].[sub_account] = [RCF0002].[SoTieuKhoan]
                AND [relationship].[date] = [RCF0002].[Ngay]
            LEFT JOIN [broker]
                ON [relationship].[broker_id] = [broker].[broker_id]
            LEFT JOIN [branch]
                ON [branch].[branch_id] = [relationship].[branch_id]
            LEFT JOIN [account] 
                ON [account].[account_code] = [relationship].[account_code]
            LEFT JOIN [vcf0051]
                ON [vcf0051].[sub_account] = [RCF0002].[SoTieuKhoan]
                AND [vcf0051].[date] = [RCF0002].[Ngay]
            WHERE [RCF0002].[Ngay] = '{self.t0_date}'
            -- WHERE [RCF0002].[Ngay] >= '2022-01-01'
            """,
            connect_DWH_CoSo
        )
        self.resultTable = self.findEmail(table)

        return self.resultTable


class ThongTinTaiKhoanChuyenChungKhoan(AbstractTable):

    def __init__(self,runTime):
        super().__init__(runTime)

    def createTable(self):
        table = pd.read_sql(
            f"""
            WITH
            [RawData] AS (
                SELECT
                    [010015].[trading_date] [Ngay],
                    CASE
                        WHEN [RSA0004].[sub_account] IS NULL
                            THEN [010015].[sub_account_or_id]
                        ELSE [RSA0004].[sub_account]
                    END [SoTieuKhoan],
                    CASE 
                        WHEN CHARINDEX('_',[010015].[ticker]) = 0
                            THEN [010015].[ticker]
                        ELSE
                            SUBSTRING([010015].[ticker],1,CHARINDEX('_',[010015].[ticker])-1)
                    END [MaChungKhoan],
                    [010015].[value_or_volume] [SoLuongChuyen]
                FROM [transaction_in_system] [010015]
                INNER JOIN [transactional_record] [RSA0004]
                    ON [RSA0004].[document_number] = [010015].[document_number]
                    AND [RSA0004].[effective_date] = [010015].[trading_date]
                    AND [RSA0004].[transaction_id] = [010015].[transaction_id]
                WHERE [010015].[transaction_id] IN ('2255','2257')
                    AND [010015].[trading_date] = '{self.t0_date}'
                    -- AND [010015].[trading_date] >= '2022-01-01'
            )
            SELECT
                [branch].[branch_name] [TenChiNhanh],
                [broker].[broker_id] [BrokerID],
                [broker].[email] [EmailFromFlex],
                [broker].[broker_name] [TenMoiGioi],
                [relationship].[account_code] [SoTaiKhoanChuyen],
                [account].[customer_name] [TenKhachHang],
                [RawData].[MaChungKhoan],
                [RawData].[SoLuongChuyen]
            FROM [RawData]
            LEFT JOIN [relationship] 
                ON [relationship].[date] = [RawData].[Ngay]
                AND [relationship].[sub_account] = [RawData].[SoTieuKhoan]
            LEFT JOIN [account] ON [account].[account_code] = [relationship].[account_code]
            LEFT JOIN [broker] ON [broker].[broker_id] = [relationship].[broker_id]
            LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
            """,
            connect_DWH_CoSo,
        )
        self.resultTable = self.findEmail(table)

        return self.resultTable

class Container:

    def __init__(self,runTime):
        self.tableObject1 = ThongTinTaiKhoanThayDoiGiaTriGiaoDich(runTime)
        self.tableObject2 = ThongTinTieuKhoanDong(runTime)
        self.tableObject3 = ThongTinTaiKhoanChuyenChungKhoan(runTime)

        self.table1 = self.tableObject1.createTable()
        self.table2 = self.tableObject2.createTable()
        self.table3 = self.tableObject3.createTable()

        self.dateString = runTime.strftime('%d/%m/%Y')
        self.exportPath = self.tableObject1.exportPath
        self.runTime = runTime
        self.filePath = None
        self.email = None
        self.brokerName = None
        self.branchName = None
        self.workbookPassword = None

    def _createSubTables(self,email):
        subTable1 = self.table1.loc[self.table1['Email']==email]
        subTable2 = self.table2.loc[self.table2['Email']==email]
        subTable3 = self.table3.loc[self.table3['Email']==email]
        return subTable1,subTable2,subTable3

    @staticmethod
    def _generatePassword():
        return ''.join(np.random.choice(10,4).astype(str))

    @staticmethod
    def _findBrokerName(*subTables): # tìm tên môi giới trong subTables
        for table in subTables:
            if not table.empty:
                brokerName = table.loc[table.index[0],'TenMoiGioi'].title()
                return brokerName

    @staticmethod
    def _findBranchName(*subTables): # tìm chi nhánh của môi giới trong subTables
        for table in subTables:
            if not table.empty:
                branchName = table.loc[table.index[0],'TenChiNhanh']
                return branchName

    def findEmailList(self):
        emailList = set()
        for table in (self.table1,self.table2,self.table3):
            emailList.update(table['Email'])
        return emailList

    def toExcel(self,email,protectSheet:bool):
        folder_name = self.tableObject1.folder_name
        period = self.tableObject1.period
        subTable1,subTable2,subTable3 = self._createSubTables(email)
        self.email = email
        self.brokerName = self._findBrokerName(subTable1,subTable2,subTable3)
        self.branchName = self._findBranchName(subTable1,subTable2,subTable3)

        self.filePath = join(dept_folder,folder_name,period,f'{email}.xlsx')
        writer = pd.ExcelWriter(
            self.filePath,
            engine='xlsxwriter',
            engine_kwargs={'options':{'nan_inf_to_errors':True}}
        )
        workbook = writer.book
        worksheet = workbook.add_worksheet('Sheet1')
        worksheet.hide_gridlines(option=2)
        worksheet.insert_image('A1',join(dirname(__file__),'img','phs_logo.png'),{'x_scale':0.65,'y_scale':0.71})

        company_name_format = workbook.add_format(
            {
                'bold':True,
                'align':'left',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
                'text_wrap':True
            }
        )
        company_info_format = workbook.add_format(
            {
                'align':'left',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
                'text_wrap':True
            }
        )
        empty_row_format = workbook.add_format(
            {
                'bottom':1,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
            }
        )
        title_format = workbook.add_format(
            {
                'bold':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':14,
                'font_name':'Arial',
                'text_wrap':True
            }
        )
        date_string_format = workbook.add_format(
            {
                'italic':True,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
                'text_wrap':True
            }
        )
        broker_name_format = workbook.add_format(
            {
                'italic':True,
                'align':'right',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
            }
        )
        table_name_format = workbook.add_format(
            {
                'bold':True,
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
            }
        )
        table_header_format = workbook.add_format(
            {
                'text_wrap':True,
                'bold':True,
                'border':1,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
            }
        )
        table_header_ratio_format = workbook.add_format(
            {
                'text_wrap':True,
                'bold':True,
                'border':1,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
                'bg_color':'#C4D79B',
            }
        )
        text_incell_center_format = workbook.add_format(
            {
                'border':1,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
            }
        )
        text_incell_left_format = workbook.add_format(
            {
                'border':1,
                'align':'left',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
            }
        )
        integer_incell_format = workbook.add_format(
            {
                'border':1,
                'align':'center',
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
                'text_wrap':True,
                'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            }
        )
        percent_incell_format = workbook.add_format(
            {
                'border':1,
                'valign':'vcenter',
                'font_size':10,
                'font_name':'Arial',
                'text_wrap':True,
                'num_format':'0.00%',
            }
        )

        worksheet.set_column('A:A',6)
        worksheet.set_column('B:B',13)
        worksheet.set_column('C:C',25)
        worksheet.set_column('D:D',13)
        worksheet.set_column('E:P',15)
        worksheet.set_column('Q:Q',1)  # margin
        worksheet.set_column('R:XFD',0)

        sheet_title_name = 'BIẾN ĐỘNG GIAO DỊCH CỦA KHÁCH HÀNG'
        file_date = self.tableObject1.t0_date
        eod_sub = dt.datetime.strptime(file_date,"%Y.%m.%d").strftime("%d.%m.%Y")
        subtitle_name = f'Ngày {eod_sub}'

        worksheet.merge_range('C1:P1',CompanyName,company_name_format)
        worksheet.merge_range('C2:P2',CompanyAddress,company_info_format)
        worksheet.merge_range('C3:P3',CompanyPhoneNumber,company_info_format)
        worksheet.merge_range('A7:P7',sheet_title_name,title_format)
        worksheet.merge_range('A8:P8',subtitle_name,date_string_format)
        worksheet.write_row('A4',['']*16,empty_row_format)
        worksheet.write('P10',f'NVQLTK: {self.brokerName} ({self.branchName})',broker_name_format)

        # Số dòng cố định của heading
        headingRows = 8
        startRow = headingRows  # bắt đầu tại dòng cuối của heading

        # Write table 1
        if not subTable1.empty:
            startRow = startRow + 3
            table1Name = '1. Thông tin tài khoản thay đổi giá trị giao dịch:'
            worksheet.write(f'A{startRow-1}',table1Name,table_name_format)
            headers = [
                'STT',
                'Số TK',
                'Họ tên khách hàng',
                'Tiểu khoản',
                'Tổng Tiền',
                'Giá trị chứng khoán',
                '% Tiền/CK',
                'Tài sản ròng\n(T-1)',
                'Tài sản ròng',
                '% Thay đổi TS',
                'Tổng giá vốn danh mục',
                'Lãi lỗ dự tính',
                '% Lãi lỗ dự tính',
                'Chứng khoán nắm giữ',
                'Chứng khoán giảm sàn (≥ 3 phiên)',
                '% mã CK giảm sàn (≥ 3 phiên)',
            ]
            subHeaders = [
                '','(1)','(2)','(3)',
                '(4)','(5)','(6) = (4) / (5)',
                '(7)','(8)','(9) = (8) / (7) - 1',
                '(10)','(11)','(12) = (11) / (10)',
                '(13)','(14)','(15) = (14) / (13)',
            ]
            for colName, header, subHeader in zip('ABCDEFGHIJKLMNOP',headers,subHeaders):
                if header.startswith('%'):
                    fmt = table_header_ratio_format
                else:
                    fmt = table_header_format
                worksheet.write(f'{colName}{startRow}',header,fmt)
                worksheet.write(f'{colName}{startRow+1}',subHeader,fmt)

            worksheet.write_column(f'A{startRow+2}',np.arange(subTable1.shape[0])+1,text_incell_center_format)
            worksheet.write_column(f'B{startRow+2}',subTable1['SoTaiKhoan'],text_incell_center_format)
            worksheet.write_column(f'C{startRow+2}',subTable1['TenKhachHang'],text_incell_center_format)
            worksheet.write_column(f'D{startRow+2}',subTable1['SoTieuKhoan'],text_incell_center_format)
            # Tạo khung
            for colName in 'EFGHIJKLMNOP':
                worksheet.write_column(f'{colName}{startRow+2}',['']*subTable1.shape[0],integer_incell_format)
            # Viết dữ liệu
            for rowNum in range(subTable1.shape[0]):
                TongTien = subTable1.loc[subTable1.index[rowNum],'TongTien']
                GiaTriChungKhoan = subTable1.loc[subTable1.index[rowNum],'GiaTriChungKhoan']
                TienGiaTriChungKhoan = subTable1.loc[subTable1.index[rowNum],'Tien/GiaTriChungKhoan']
                TaiSanRongT0 = subTable1.loc[subTable1.index[rowNum],'TaiSanRongT0']
                TaiSanRongT1 = subTable1.loc[subTable1.index[rowNum],'TaiSanRongT1']
                PhanTramThayDoiTaiSan = subTable1.loc[subTable1.index[rowNum],'PhanTramThayDoiTaiSan']
                GiaVon = subTable1.loc[subTable1.index[rowNum],'GiaVon']
                LaiLoDuTinh = subTable1.loc[subTable1.index[rowNum],'LaiLoDuTinh']
                PhanTramLaiLoDuTinh = subTable1.loc[subTable1.index[rowNum],'PhanTramLaiLoDuTinh']
                MaChungKhoanDanhMuc = subTable1.loc[subTable1.index[rowNum],'MaChungKhoanDanhMuc']
                MaChungKhoanGiamSan3Phien = subTable1.loc[subTable1.index[rowNum],'MaChungKhoanGiamSan3Phien']
                PhanTramMaGiamSan3Phien = subTable1.loc[subTable1.index[rowNum],'PhanTramMaGiamSan3Phien']
                if TienGiaTriChungKhoan >= 1:
                    worksheet.write(f'E{startRow+2+rowNum}',TongTien,integer_incell_format)
                    worksheet.write(f'F{startRow+2+rowNum}',GiaTriChungKhoan,integer_incell_format)
                    worksheet.write(f'G{startRow+2+rowNum}',TienGiaTriChungKhoan,percent_incell_format)
                if PhanTramThayDoiTaiSan <= -0.25:
                    worksheet.write(f'H{startRow+2+rowNum}',TaiSanRongT1,integer_incell_format)
                    worksheet.write(f'I{startRow+2+rowNum}',TaiSanRongT0,integer_incell_format)
                    worksheet.write(f'J{startRow+2+rowNum}',PhanTramThayDoiTaiSan,percent_incell_format)
                if PhanTramLaiLoDuTinh <= -0.5:
                    worksheet.write(f'K{startRow+2+rowNum}',GiaVon,integer_incell_format)
                    worksheet.write(f'L{startRow+2+rowNum}',LaiLoDuTinh,integer_incell_format)
                    worksheet.write(f'M{startRow+2+rowNum}',PhanTramLaiLoDuTinh,percent_incell_format)
                if PhanTramMaGiamSan3Phien >= 0.5:
                    worksheet.write(f'N{startRow+2+rowNum}',MaChungKhoanDanhMuc,text_incell_left_format)
                    worksheet.write(f'O{startRow+2+rowNum}',MaChungKhoanGiamSan3Phien,text_incell_left_format)
                    worksheet.write(f'P{startRow+2+rowNum}',PhanTramMaGiamSan3Phien,percent_incell_format)

        # Write table 2
        startRow = startRow + subTable1.shape[0] + 1
        if not subTable2.empty:
            startRow = startRow + 3
            table1Name = '2. Thông tin tiểu khoản đóng:'
            worksheet.write(f'A{startRow-1}',table1Name,table_name_format)
            headers = [
                'STT',
                'Số TK lưu ký',
                'Họ tên khách hàng',
                'Số tiểu khoản đóng',
                'Loại hình',
            ]
            worksheet.write_row(f'A{startRow}',headers,table_header_format)
            worksheet.write_column(f'A{startRow+1}',np.arange(subTable2.shape[0])+1,text_incell_center_format)
            worksheet.write_column(f'B{startRow+1}',subTable2['SoTaiKhoan'],text_incell_center_format)
            worksheet.write_column(f'C{startRow+1}',subTable2['TenKhachHang'],text_incell_center_format)
            worksheet.write_column(f'D{startRow+1}',subTable2['SoTieuKhoanDong'],text_incell_center_format)
            worksheet.write_column(f'E{startRow+1}',subTable2['LoaiHinh'],text_incell_center_format)

        # Write table 3
        startRow = startRow + subTable2.shape[0]
        if not subTable3.empty:
            startRow = startRow + 3
            table1Name = '3. Thông tin tài khoản chuyển chứng khoán:'
            worksheet.write(f'A{startRow-1}',table1Name,table_name_format)
            headers = [
                'STT',
                'Số TK chuyển',
                'Họ tên khách hàng',
                'Mã CK',
                'Số lượng chuyển',
            ]
            worksheet.write_row(f'A{startRow}',headers,table_header_format)
            worksheet.write_column(f'A{startRow+1}',np.arange(subTable3.shape[0])+1,text_incell_center_format)
            worksheet.write_column(f'B{startRow+1}',subTable3['SoTaiKhoanChuyen'],text_incell_center_format)
            worksheet.write_column(f'C{startRow+1}',subTable3['TenKhachHang'],text_incell_center_format)
            worksheet.write_column(f'D{startRow+1}',subTable3['MaChungKhoan'],text_incell_center_format)
            worksheet.write_column(f'E{startRow+1}',subTable3['SoLuongChuyen'],integer_incell_format)

        # Write footer
        startRow = startRow + subTable3.shape[0] + 3
        worksheet.merge_range(f'A{startRow-1}:P{startRow-1}','--End--',empty_row_format)

        if protectSheet:
            worksheet.protect(
                password='admin@PHS12345',
                options={
                    'objects':True,
                    'scenarios':True,
                    'select_locked_cells':True,
                    'select_unlocked_cells':True,
                    'sort':True,
                    'autofilter':True,
                }
            )
        writer.close()
        return self

    def encrypt(self):
        self.workbookPassword = self._generatePassword()
        excel = Dispatch('Excel.Application')
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(self.filePath)
        workbook.SaveAs(self.filePath,51,self.workbookPassword) # 51 means .xlsx
        workbook.Close()
        return self

    def sendMail(self):

        outlook = Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'hiepdang@phs.vn' # = self.email
        mail.Subject = f"Biến động giao dịch của khách hàng, ngày {self.dateString}_NVQLTK: {self.email}"
        body = f"""
            <html>
                <head></head>
                <body>
                    <p style="font-family:Arial; font-size:90%">
                        Kính gửi Anh/Chị,
                    </p>
                    <p></p>
                    <p style="font-family:Arial; font-size:90%">
                        Để hỗ trợ Khối Môi Giới chăm sóc khách hàng, 
                        PHS gửi đến Anh/Chị biến động giao dịch của khách hàng ngày {self.dateString}. 
                        Anh/Chị vui lòng sử dụng mật khẩu <b style="font-size:120%">{self.workbookPassword}</b> để xem file đính kèm.
                    </p>
                    <p></p>
                    <p style="font-family:Arial; font-size:90%">
                        Mọi thắc mắc Anh/Chị vui lòng liên hệ với ...
                    </p>
                    <p></p>
                    <p style="font-family:Arial; font-size:90%">
                        Trân trọng.
                    </p>
                    <p></p>
                    <p style="font-family:Times New Roman; font-size:90%"><i>
                        -- Generated by Reporting System
                    </i></p>
                </body>
            </html>
            """
        mail.HTMLBody = body
        mail.Attachments.Add(self.filePath)
        mail.Send()


# client code
def run(
    runTime:dt.datetime=dt.datetime.now(),
    send_mail:bool=False,
):

    container = Container(runTime)
    emailList = container.findEmailList()
    for email in list(emailList):
        if send_mail:
            container.toExcel(email,protectSheet=True).encrypt().sendMail()
        else:
            container.toExcel(email,protectSheet=False)

    return container.tableObject1, container.tableObject2, container.tableObject3

