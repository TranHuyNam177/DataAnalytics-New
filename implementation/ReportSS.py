from implementation import TaskMonitor
from request.stock import *

@TaskMonitor
def SS_BaoCaoCheckGia():
    from automation.trading_service.giaodichluuky import BaoCaoCheckGia
    BaoCaoCheckGia.run()

@TaskMonitor
def SS_BaoCaoDanhSachChungKhoanVCBS():
    from automation.trading_service.giaodichluuky import BaoCaoDanhSachChungKhoanVCBS
    BaoCaoDanhSachChungKhoanVCBS.run()

@TaskMonitor
def SS_BaoCaoDongMoUyQuyenTK_HNX():
    from automation.trading_service.giaodichluuky import BaoCaoDongMoUyQuyenTK_HNX
    BaoCaoDongMoUyQuyenTK_HNX.run()

@TaskMonitor
def SS_BaoCaoDongMoUyQuyenTK_HOSEnew():
    from automation.trading_service.giaodichluuky import BaoCaoDongMoUyQuyenTK_HOSEnew
    BaoCaoDongMoUyQuyenTK_HOSEnew.run()

@TaskMonitor
def SS_BaoCaoDongMoUyQuyenTK_HOSEold():
    from automation.trading_service.giaodichluuky import BaoCaoDongMoUyQuyenTK_HOSEold
    BaoCaoDongMoUyQuyenTK_HOSEold.run()

@TaskMonitor
def SS_BaoCaoHoatDongLuuKyNDTNN():
    from automation.trading_service.giaodichluuky import _BaoCaoHoatDongLuuKyNDTNN
    BaoCaoHoatDongLuuKyNDTNN.run()

@TaskMonitor
def SS_BaoCaoTamTinhPhiChuyenKhoan():
    from automation.trading_service.giaodichluuky import BaoCaoTamTinhPhiChuyenKhoan
    BaoCaoTamTinhPhiChuyenKhoan.run()

@TaskMonitor
def SS_BaoCaoTamTinhPhiGiaoDich():
    from automation.trading_service.giaodichluuky import BaoCaoTamTinhPhiGiaoDich
    BaoCaoTamTinhPhiGiaoDich.run()

@TaskMonitor
def SS_BaoCaoTamTinhPhiLuuKy():
    from automation.trading_service.giaodichluuky import BaoCaoTamTinhPhiLuuKy
    BaoCaoTamTinhPhiLuuKy.run()

@TaskMonitor
def SS_BaoCaoTienGuiSSC():
    from automation.trading_service.giaodichluuky import BaoCaoTienGuiSSC
    BaoCaoTienGuiSSC.run()

@TaskMonitor
def SS_BaoCaoTinhHinhHDKD():
    from automation.trading_service.giaodichluuky import BaoCaoTinhHinhHDKD
    BaoCaoTinhHinhHDKD.run()

@TaskMonitor
def SS_BaoCaoDoiChieuLaiTienGuiPhatSinh_p1():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuLaiTienGuiPhatSinh_p1
    BaoCaoDoiChieuLaiTienGuiPhatSinh_p1.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_BaoCaoDoiChieuLaiTienGuiPhatSinh_p2():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuLaiTienGuiPhatSinh_p2
    BaoCaoDoiChieuLaiTienGuiPhatSinh_p2.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_SaveTempData1():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuSoDuTaiKhoanKhachHang
    BaoCaoDoiChieuSoDuTaiKhoanKhachHang.generateTempData()

@TaskMonitor
def SS_BaoCaoDoiChieuSoDuTaiKhoanKhachHang():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuSoDuTaiKhoanKhachHang
    BaoCaoDoiChieuSoDuTaiKhoanKhachHang.run()

@TaskMonitor
def SS_BaoCaoDoiChieuTTBTTienMuaBanChungKhoan():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuTTBTTienMuaBanChungKhoan
    BaoCaoDoiChieuTTBTTienMuaBanChungKhoan.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_BaoCaoDoiChieuUTTB():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuUTTB
    BaoCaoDoiChieuUTTB.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_BaoCaoDoiChieuVaImportEIB():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuVaImportEIB
    BaoCaoDoiChieuVaImportEIB.run()

@TaskMonitor
def SS_BaoCaoDoiChieuVaImportOCB():
    from automation.trading_service.thanhtoanbutru import BaoCaoDoiChieuVaImportOCB
    BaoCaoDoiChieuVaImportOCB.run()

@TaskMonitor
def SS_BaoCaoDuLieuGuiKiemToan():
    from automation.trading_service.thanhtoanbutru import BaoCaoDuLieuGuiKiemToan
    BaoCaoDuLieuGuiKiemToan.run()

@TaskMonitor
def SS_BaoCaoFileGhiAmGuiKSNB():
    from automation.trading_service.thanhtoanbutru import BaoCaoFileGhiAmGuiKSNB
    BaoCaoFileGhiAmGuiKSNB.run()

@TaskMonitor
def SS_BaoCaoFileImportTKNganHangBuoiSang():
    from automation.trading_service.thanhtoanbutru import BaoCaoFileImportTKNganHangBuoiSang
    BaoCaoFileImportTKNganHangBuoiSang.run()

@TaskMonitor
def SS_BaoCaoKHNNCoSoDuNho():
    from automation.trading_service.thanhtoanbutru import BaoCaoKHNNCoSoDuNho
    BaoCaoKHNNCoSoDuNho.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_BaoCaoPhatSinhGiaoDichTien():
    from automation.trading_service.thanhtoanbutru import BaoCaoPhatSinhGiaoDichTien
    BaoCaoPhatSinhGiaoDichTien.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_BaoCaoPSTheoDoiTKKHCanChuyenTienTuVSDVeTKGD():
    from automation.trading_service.thanhtoanbutru import BaoCaoPSTheoDoiTKKHCanChuyenTienTuVSDVeTKGD
    BaoCaoPSTheoDoiTKKHCanChuyenTienTuVSDVeTKGD.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_BaoCaoReviewKHVIPCoSo():
    from automation.trading_service.thanhtoanbutru import BaoCaoReviewKHVIPCoSo
    BaoCaoReviewKHVIPCoSo.run()

@TaskMonitor
def SS_BaoCaoSoDuTienKyQuyPhaiSinh():
    from automation.trading_service.thanhtoanbutru import BaoCaoSoDuTienKyQuyPhaiSinh
    BaoCaoSoDuTienKyQuyPhaiSinh.run()

@TaskMonitor
def SS_BaoCaoSoLieuThongKeDVTC():
    from automation.trading_service.thanhtoanbutru import BaoCaoSoLieuThongKeDVTC
    BaoCaoSoLieuThongKeDVTC.run()

@TaskMonitor
def SS_BaoCaoUngTienROS():
    from automation.trading_service.thanhtoanbutru import BaoCaoUngTienROS
    BaoCaoUngTienROS.run()

@TaskMonitor
def SS_BaoCaoXacNhanNopRutEIBOCB_p1():
    from automation.trading_service.thanhtoanbutru import BaoCaoXacNhanNopRutEIBOCB_p1
    BaoCaoXacNhanNopRutEIBOCB_p1.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_BaoCaoXacNhanNopRutEIBOCB_p2():
    from automation.trading_service.thanhtoanbutru import BaoCaoXacNhanNopRutEIBOCB_p2
    BaoCaoXacNhanNopRutEIBOCB_p2.run()

@TaskMonitor
def SS_DanhSachKHVIPCoSo():
    from automation.trading_service.thanhtoanbutru import DanhSachKHVIPCoSo
    DanhSachKHVIPCoSo.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def SS_DanhSachKHVIPPhaiSinh():
    from automation.trading_service.thanhtoanbutru import DanhSachKHVIPPhaiSinh
    DanhSachKHVIPPhaiSinh.run(dt.datetime.now()-dt.timedelta(days=1))


"""
Cần viết tool email báo có chênh lệch trong các BC đối chiếu 
(as decorators that wrap functions from automation package)
"""
