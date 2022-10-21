from bs4 import BeautifulSoup
import requests
import re
from abc import ABC
import datetime as dt
import calendar
import decimal
from dateutil.relativedelta import relativedelta

"""
* Tỷ lệ thực hiện của tất cả các quyền nếu trên web là num:num thì luôn trả về num/num để nhập vào flex
1. Quyền tham dự đại hội cổ đông
    - Tỷ lệ thực hiện: 
        + không bao giờ có trường hợp nhiều hơn 3 (Ngày, Tháng, Qúy)
        + khi có trường hợp 2 (Ngày, Tháng, Qúy) thì lấy cái xa nhất
2. Quyền mua
    - Tỷ lệ thực hiện:
        Rule trong file Excel
        + Dòng 11: Tỷ lệ cổ phiếu sở hữu/quyền - mặc định 1/1
        + Dòng 12: Tỷ lệ quyền/cổ phiếu được mua - có thay đổi dựa vào Tỷ lệ thực hiện trên trang web (phần: "... quyền
        được mua ... cổ phiếu mới")
    - Ngày bắt đầu ĐK quyền mua, Ngày ĐK quyền mua cuối cùng
        + Lấy từ "+ Thời gian đăng ký đặt mua và nộp tiền mua cổ phiếu: hoặc
        từ "Thời gian đăng ký và đặt mua cổ phiếu:" dưới phần LỊCH TRÌNH THỰC HIỆN QUYỀN MUA - lấy theo cái nào cũng '
        được, 2 cái là 1

"""

class Base(ABC):
    def __init__(self,url):
        self.url = url
        self.__headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
        }
        self.HTML = self._toHTML()
        self.__contentNews = self.getContentNews()
        self.__ngayDKCC = dt.datetime.strptime(self.getNgayDKCC(), '%d/%m/%Y')

    # concrete method, vì có implementation
    def _toHTML(self):
        """
        This function return html
        """
        with requests.Session() as session:
            retry = requests.packages.urllib3.util.retry.Retry(connect=5, backoff_factor=1)
            adapter = requests.adapters.HTTPAdapter(max_retries=retry)
            session.mount('https://', adapter)
            html = session.get(self.url, headers=self.__headers, timeout=10).text
            soup = BeautifulSoup(html,'html5lib')
            return soup.find(class_='container-small')

    def getTitleNews(self):
        return self.HTML.find(class_="title-category").get_text()

    def getTimeNews(self):
        return self.HTML.find(class_="time-newstcph").get_text()

    def getContentNews(self):
        contentNews = self.HTML.find(class_='col-md-12').get_text()
        # return re.sub(r'\s{2,}', ' ', contentNews).replace('\n', ' ')
        return re.sub(r'\s{3,}', '\n', contentNews).replace('\xa0',' ')

    def getMaCK(self):
        pattern = r'(Mã chứng khoán\s*:\s*[\r\n])(.*)([\r\n])'
        return re.search(pattern, self.__contentNews).group(2).strip()

    def getMaISIN(self):
        pattern = r'(Mã ISIN\s*:\s*[\r\n])(.*)([\r\n])'
        return re.search(pattern, self.__contentNews).group(2).strip()

    def getNgayDKCC(self):
        pattern = r'(Ngày đăng ký cuối cùng\s*:\s*[\r\n])(.*)([\r\n])'
        return re.search(pattern, self.__contentNews).group(2).strip()

    def getLyDoMucDich(self):
        pattern = r'(Lý do mục đích\s*:\s*[\r\n])(.*)([\r\n])'
        return re.search(pattern, self.__contentNews).group(2).rstrip('.')

    def __findNgayThucHienDuKien(self):
        pattern = re.compile(r'((Thời gian|Ngày) thực thanh toán\s*:?\s*)(.*)([\r\n])')
        if re.search(pattern, self.__contentNews):
            stringRegex = re.search(pattern, self.__contentNews).group(3).strip().upper()
        else:
            pattern = re.compile(r'((Thời gian|Ngày) (thực hiện|thanh toán)\s*:?\s*)(.*)([\r\n])')
            stringRegex = re.search(pattern, self.__contentNews).group(4).strip().upper()

        # xử lý case có format %m/%Y
        def processCaseMonth(listString):
            dateString = listString[-1]
            dateString = dateString.replace('THÁNG ', '')
            month = int(re.split(r'(/|NĂM)', dateString)[0])
            year = int(re.split(r'(/|NĂM)', dateString)[-1])
            day = calendar.monthrange(year, month)[1]
            return dt.datetime(year,month,day).strftime('%d/%m/%Y')

        # xử lý case có format Quý/Năm
        def processCaseQuarter(listTuple):
            quarterString = listTuple[-1][0]
            quarterString = quarterString.replace('QUÝ ', '')
            quarter = re.split(r'(/|NĂM)', quarterString)[0]
            year = int(re.split(r'(/|NĂM)', quarterString)[-1])
            if quarter == 'I':
                return dt.datetime(year, 3, 31).strftime('%d/%m/%Y')
            elif quarter == 'II':
                return dt.datetime(year, 6, 30).strftime('%d/%m/%Y')
            elif quarter == 'III':
                return dt.datetime(year, 9, 30).strftime('%d/%m/%Y')
            else:
                return dt.datetime(year, 12, 31).strftime('%d/%m/%Y')

        # xử lý case có format %d/%m/%Y
        def processCaseDate(listString):
            return max([dt.datetime.strptime(d, "%d/%m/%Y") for d in listString]).strftime("%d/%m/%Y")

        caseMapper = {
            r'(QUÝ\s*(IV|III|II|I)/\d{4})': processCaseQuarter,
            r'(THÁNG\s*\d{1,2}[/\d{4}|NĂM\s*\d{4}]*)': processCaseMonth,
            r'\b\d{1,2}/\d{1,2}/\d{4}\b': processCaseDate,
        }
        for pattern in caseMapper.keys():
            results = re.findall(pattern, stringRegex)
            if results:
                print(results)
                return caseMapper[pattern](results)  # chỗ này catch trường hợp có nhiều date, month, quarter trong 1 câu

        return self.__ngayDKCC.replace(month=self.__ngayDKCC.month+1).strftime('%d/%m/%Y')  # case thông báo sau

    def getNgayThucHienDuKien(self):
        return self.__findNgayThucHienDuKien()

    def __findTyLeThucHien(self):
        """
        Nếu là Class QuyenMua thi giá trị trả về sẽ dùng để truyền vào
        trường thông tin Tỷ lệ quyền/cổ phiếu được mua trong flex
        """
        if re.search(r'(Tỷ lệ thực hiện\s*:?\s*)(.*)([\r\n])', self.__contentNews):
            stringRegex = re.search(r'(Tỷ lệ thực hiện\s*:?\s*)(.*)([\r\n])', self.__contentNews).group(2).strip()
            # stringRegex = re.search(pattern, self.__subContentNews).group(2).strip()
            print("StringRegex TyLeThucHien From Other Class:", stringRegex)
            pattern = r'(\d+[\d.]*[,]?[\d+]*%|\d+[\d.]*[,]?[\d+]*:\d+[\d.]*[,]?[\d+]*)'
            if re.search(pattern, stringRegex):
                tyLeThucHien = re.search(pattern, stringRegex).group()
                tyLeThucHien = re.sub('[.%]', '', tyLeThucHien).replace(',', '.').replace(':','/')
            else:
                if ',' in stringRegex:
                    stringRegex = stringRegex.split(',')[-1]
                pattern = r'\b\d+[\d.]*[,]?[\d+]*\b'
                listNum = re.findall(pattern, stringRegex)
                num1 = listNum[0].replace(',', '.')
                num2 = listNum[1].replace(',', '.')
                tyLeThucHien = f'{num1}/{num2}'
        elif re.search(r'(Tỷ lệ chuyển đổi\s*:?\s*)(.*)([\r\n])', self.__contentNews):
            stringRegex = re.search(r'(Tỷ lệ chuyển đổi\s*:?\s*)(.*)([\r\n])', self.__contentNews).group(2).strip()
            print("StringRegex TyLeThucHien From Class ChuyenDoiTraiPhieuThanhCP:", stringRegex)
            pattern = r'\d+[\d.]*[,]?[\d+]*:\d+[\d.]*[,]?[\d+]*'
            if re.search(pattern, stringRegex):
                tyLeString = re.search(pattern, stringRegex).group()
                tyLeString = tyLeString.replace('.','').replace(',', '.')
                num1 = float(tyLeString.split(':')[0]) * 100
                num2 = float(tyLeString.split(':')[1]) * 100
            else:
                pattern = r'\b\d+[\d.]*[,]?[\d+]*\b'
                listNum = re.findall(pattern, stringRegex)
                num1 = float(listNum[0].replace(',', '.')) * 100
                num2 = float(listNum[1].replace(',', '.')) * 100
            tyLeThucHien = f'{num1}/{num2}'
        else:
            pattern = r'(Tỷ lệ thanh toán\s*:?\s*)(.*)([\r\n])'
            stringRegex = re.search(pattern, self.getContentNews()).group(2).strip()
            print("StringRegex TyLeThucHien class TraLaiTraiPhieu:", stringRegex)
            if re.search('(tiền gốc|tiền lãi)', stringRegex):
                soTien = re.search(r'(\d+[\d.]*[,]?[\d+]*)\s*đồng tiền lãi', stringRegex).group(1)
                soTien = soTien.replace('.', '').replace(',', '.')
            else:
                soTienStrings = re.findall(r'\b\d+[\d.]*[,]?[\d+]*\b', stringRegex.replace('.', '').replace(',', '.'))
                soTien = max(soTienStrings)
            print(soTien)
            tyLeThucHien = float(soTien) / 100000 * 100
        return tyLeThucHien

    def getTyLeThucHien(self):
        """
            "tỷ lệ thực hiện, tỷ lệ chia tiền,
            (%) Tỉ lệ chia cổ tức %, (%) Tỉ lệ thưởng %,
            (%) lãi suất/kỳ"
            xài cùng 1 hàm này
        """
        return self.__findTyLeThucHien()

    def __findExamplePara(self):
        """
        hàm dùng để lấy đoạn ví dụ trong bài báo
        """
        contentNews = self.__contentNews.replace('\n', ' ')
        # contentNews = self.__subContentNews.replace('\n', ' ')
        pattern = r'(Phương án làm tròn[,]?\s*)(.*)(\s*Địa điểm thực hiện)'
        if re.search(pattern, contentNews):
            exampleString = re.search(pattern, contentNews).group(2).strip()
        else:
            pattern = r'(Toàn bộ phần lẻ\s*)(.*)(\s*Thời gian đăng ký chuyển đổi)'
            exampleString = re.search(pattern, contentNews).group(2).strip()
        return exampleString

    def findGiaQuiDoiCPLe(self):
        stringRegex = self.__findExamplePara()
        patternMuaLai = r'(mua lại.*giá\s*)(.*)(\s*đồng/cổ phiếu)'
        if re.search(patternMuaLai, stringRegex):
            giaQuiDoiCPLe = re.search(patternMuaLai, stringRegex).group(2).replace('.', '')
        else:
            giaQuiDoiCPLe = 0
        return giaQuiDoiCPLe

    def findSoChuSoThapPhanCKLe(self):
        stringRegex = self.__findExamplePara()
        pattern = r'(\s*=\s*)(\d+[\d.]*,\d+)'
        if re.search(pattern, stringRegex):
            result = re.search(pattern, stringRegex).group(2).replace(' ', '').replace(',', '.')
        else:
            pattern = r'(\d+[\d.]*,\d+)(\s*cổ phiếu)(.*=)'
            result = re.search(pattern, stringRegex).group(1).replace(',', '.')
        chuSoThapPhan = abs(decimal.Decimal(result).as_tuple().exponent)
        return chuSoThapPhan

class ThamDuDHCD(Base):
    def __init__(self, url):
        super().__init__(url)
        self.__ngayDKCC = dt.datetime.strptime(self.getNgayDKCC(), '%d/%m/%Y')
        self.__contentNews = self.getContentNews()

    def getTyLeThucHien(self):
        return '1/1'  # theo rule là mặc định với quyền Tham dự đại hội cổ đông

class ChiaCoTucBangTien(Base):
    def __init__(self, url):
        super().__init__(url)
        self.__contentNews = self.getContentNews()

class ChiaCoTucBangCoPhieu(Base):
    def __init__(self, url):
        super().__init__(url)
        self.__ngayDKCC = dt.datetime.strptime(self.getNgayDKCC(), '%d/%m/%Y')

    def getNgayThucHienDuKien(self):
        return (self.__ngayDKCC + relativedelta(months=3)).strftime('%d/%m/%Y')

    def getGiaQuiDoiCPLe(self):
        return self.findGiaQuiDoiCPLe()

    def getSoChuSoThapPhanCKLe(self):
        return self.findSoChuSoThapPhanCKLe()

class CoPhieuThuong(Base):
    def __init__(self, url):
        super().__init__(url)
        self.__ngayDKCC = dt.datetime.strptime(self.getNgayDKCC(), '%d/%m/%Y')

    def getNgayThucHienDuKien(self):
        return (self.__ngayDKCC + relativedelta(months=3)).strftime('%d/%m/%Y')

    def getGiaQuiDoiCPLe(self):
        return self.findGiaQuiDoiCPLe()

    def getSoChuSoThapPhanCKLe(self):
        return self.findSoChuSoThapPhanCKLe()

class QuyenMua(Base):
    def __init__(self, url):
        super().__init__(url)
        self.__ngayDKCC = dt.datetime.strptime(self.getNgayDKCC(), '%d/%m/%Y')
        self.__contentNews = self.getContentNews().replace('\n', ' ')

    def getNgayThucHienDuKien(self):
        return (self.__ngayDKCC + relativedelta(months=3)).strftime('%d/%m/%Y')

    def getGiaPhatHanh(self):
        """
        Trong flex là Giá mua
        """
        pattern = r'(Giá phát hành)(\d+[\d.]*[\d+]*)(\s*đồng[/]?)'
        return re.search(pattern, self.__contentNews).group(2).strip().replace('.','')

    def getTyLeCPSoHuu_Quyen(self):
        return '1/1'

    def getTyLeQuyen_CPDuocMua(self):
        return self.getTyLeThucHien()

    def __findTransferDate(self):
        if 'Điều chỉnh thời gian thực hiện quyền mua' in self.getTitleNews():
            pattern = r'(Thông tin điều chỉnh)(.*)(Thời gian đăng ký và đặt mua cổ phiếu)'
        else:
            pattern = r'(Thời gian chuyển nhượng quyền mua cổ phiếu)(.*)(Thời gian đăng ký và đặt mua cổ phiếu)'
        stringRegex = re.search(pattern, self.__contentNews).group(2)
        print("StringRegex TransferDate:", stringRegex)
        dateStrings = re.findall(r'\d{1,2}/\d{1,2}/\d{4}', stringRegex)
        dates = sorted([dt.datetime.strptime(dateString, '%d/%m/%Y') for dateString in dateStrings])
        transferStartDate, transferEndDate = dates[-2:]
        return transferStartDate, transferEndDate

    def getTransferStartDate(self):
        return self.__findTransferDate()[0].strftime('%d/%m/%Y')

    def getTransferEndDate(self):
        return self.__findTransferDate()[1].strftime('%d/%m/%Y')

    def __findRegisterBuyDate(self):
        if 'Điều chỉnh thời gian thực hiện quyền mua' in self.getTitleNews():
            pattern = r'(Thông tin điều chỉnh)(.*)(Thời hạn thành viên lưu ký)'
        else:
            pattern = r'(Thời gian đăng ký và đặt mua cổ phiếu)(.*)(Thời hạn TVLK)'
        stringRegex = re.search(pattern, self.__contentNews).group(2)
        print("StringRegex RegisterBuyDate", stringRegex)
        dateStrings = re.findall(r'\d{1,2}/\d{1,2}/\d{4}', stringRegex)
        dates = sorted([dt.datetime.strptime(dateString, '%d/%m/%Y') for dateString in dateStrings])
        registerBuyRightStartDate, registerBuyRightEndDate = dates[-2:]
        return registerBuyRightStartDate, registerBuyRightEndDate

    def getRegisterBuyStartDate(self):
        return self.__findRegisterBuyDate()[0].strftime('%d/%m/%Y')

    def getRegisterBuyEndDate(self):
        return self.__findRegisterBuyDate()[1].strftime('%d/%m/%Y')

class TraLaiTraiPhieu(Base):
    def __init__(self, url):
        super().__init__(url)

class TraGocVaLaiTraiPhieu(Base):
    def __init__(self, url):
        super().__init__(url)

class ChuyenDoiTraiPhieuThanhCP(Base):
    def __init__(self, url):
        super().__init__(url)

    def getGiaQuiDoiCPLe(self):
        return self.findGiaQuiDoiCPLe()

    def getSoChuSoThapPhanCKLe(self):
        return self.findSoChuSoThapPhanCKLe()

    def __findNgayThucHienDuKien(self):
        pattern = r'(Thời gian đăng ký chuyển đổi)(.*)(Thời gian tạm ngừng)'
        stringRegex = re.search(pattern, self.getContentNews().replace('\n', ' ')).group(2)
        print("StringRegex NgayThucHienDuKien Class ChuyenDoiTraiPhieuThanhCP:", stringRegex)
        dateStrings = re.findall(r'\d{1,2}/\d{1,2}/\d{4}', stringRegex)
        ngayThucHienDuKien = max([dt.datetime.strptime(dateString, '%d/%m/%Y') for dateString in dateStrings])
        return ngayThucHienDuKien

    def getNgayThucHienDuKien(self):
        return self.__findNgayThucHienDuKien().strftime('%d/%m/%Y')

# client code
class F220001:
    def __init__(self,url):
        # phan loai quyen
        title = Base(url).getTitleNews()
        if '....' in title:
            rightObject = ThamDuDHCD(url)

    def insertRight(self):
        # insert Title
        # insert ISIN
        pass