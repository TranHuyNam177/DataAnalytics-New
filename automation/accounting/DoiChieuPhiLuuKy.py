import re
import pandas as pd
from tqdm import tqdm
from automation.accounting import *


# Client code
def run(
    run_time=dt.datetime.now()
):

    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    t0_date = info['end_date'].replace('.','-')
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    NoPhiLuuKyFlex = pd.read_sql(
        """
        SELECT [SoTieuKhoan], [PhiLuuKy]
        FROM [NoPhiLuuKy]
        WHERE [Ngay] = '2022-07-12'
        """,
        connect_DWH_CoSo,
    )
    NoPhiLuuKyBravo = pd.read_excel(
        """
        
        """,
        connect_DWH_CoSo,
    )
    mappingSeries = pd.read_sql(
        f"""
        SELECT [sub_account], [account_code] FROM [sub_account]
        """,
        connect_DWH_CoSo,
        index_col='sub_account',
    ).squeeze()
    IB['SoTaiKhoan'] = IB['NoiDungIB'].apply(_findAccount,args=(mappingSeries,))

    # Group, ignore content
    RCI1025agg = RCI1025.groupby(['SoTaiKhoan','NganHang','TaiKhoanNganHang'],as_index=False)['TienNop'].agg(np.sum)
    IBagg = IB.groupby(['SoTaiKhoan','NganHang','TaiKhoanNganHang'],as_index=False)['TienNop'].agg(np.sum)

    # Tạo bảng so sánh
    compareTable = pd.merge(RCI1025agg,IBagg,how='outer',on=['SoTaiKhoan','NganHang','TaiKhoanNganHang'],suffixes=('Flex','IB'))
    compareTable['TienNopFlex'] = compareTable['TienNopFlex'].fillna(0)
    compareTable['TienNopIB'] = compareTable['TienNopIB'].fillna(0)
    compareTable['TienNopDiff'] = compareTable['TienNopFlex'] - compareTable['TienNopIB']
    compareTable = compareTable.sort_values('TienNopDiff',key=lambda x: x.abs(),ascending=False)

    ############################################################
    ############################################################
    ############################################################

    file_date = dt.datetime.strptime(t0_date,'%Y-%m-%d').strftime('%d.%m.%Y')
    file_name = f'Đối Chiếu Nhập Tiền {file_date}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    company_name_format = workbook.add_format(
        {
            'bold':True,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    company_info_format = workbook.add_format(
        {
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    empty_row_format = workbook.add_format(
        {
            'bottom':1,
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    sheet_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':14,
            'font_name':'Times New Roman',
            'text_wrap':True
        }
    )
    sub_title_format = workbook.add_format(
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
    headers_root_format = workbook.add_format(
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
    headers_ib_format = workbook.add_format(
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
    headers_flex_format = workbook.add_format(
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
    headers_diff_format = workbook.add_format(
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
    text_root_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    text_root_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_size':10,
            'font_name':'Times New Roman',
        }
    )
    number_ib_format = workbook.add_format(
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
    number_flex_format = workbook.add_format(
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
    number_diff_format = workbook.add_format(
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

    ############################################################
    ############################################################
    ############################################################

    worksheet = workbook.add_worksheet('SoSanh')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',17)
    worksheet.set_column('B:B',9)
    worksheet.set_column('C:E',17)
    worksheet.set_column('F:F',19)

    worksheet.merge_range('A1:F1',CompanyName,company_name_format)
    worksheet.merge_range('A2:F2',CompanyAddress,company_info_format)
    worksheet.merge_range('A3:F3',CompanyPhoneNumber,company_info_format)
    worksheet.write_row('A4',['']*6,empty_row_format)
    worksheet.merge_range('A6:F6','Đối chiếu RCI1025 & Internet Banking',sheet_title_format)
    worksheet.merge_range('A7:F7',f'Date: {file_date}',sub_title_format)

    worksheet.write_row('A9',['Số tài khoản','Ngân hàng','Tài khoản ngân hàng'],headers_root_format)
    worksheet.write('D9','Tiền nộp Flex',headers_flex_format)
    worksheet.write('E9','Tiền nộp IB',headers_ib_format)
    worksheet.write('F9','Chênh lệch (Flex - IB)',headers_diff_format)

    worksheet.write_column('A10',compareTable['SoTaiKhoan'],text_root_center_format)
    worksheet.write_column('B10',compareTable['NganHang'],text_root_center_format)
    worksheet.write_column('C10',compareTable['TaiKhoanNganHang'],text_root_center_format)
    worksheet.write_column('D10',compareTable['TienNopFlex'],number_flex_format)
    worksheet.write_column('E10',compareTable['TienNopIB'],number_ib_format)
    worksheet.write_column('F10',compareTable['TienNopDiff'],number_diff_format)

    worksheet.merge_range(f'A{compareTable.shape[0]+10}:C{compareTable.shape[0]+10}','Tổng',headers_root_format)
    worksheet.write(f'D{compareTable.shape[0]+10}',f'=SUM(D10:D{compareTable.shape[0]+9})',number_flex_format)
    worksheet.write(f'E{compareTable.shape[0]+10}',f'=SUM(E10:E{compareTable.shape[0]+9})',number_ib_format)
    worksheet.write(f'F{compareTable.shape[0]+10}',f'=SUM(F10:F{compareTable.shape[0]+9})',number_diff_format)

    ############################################################
    ############################################################
    ############################################################
    
    worksheet = workbook.add_worksheet('InternetBanking')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',15)
    worksheet.set_column('B:B',9)
    worksheet.set_column('C:C',17)
    worksheet.set_column('D:D',105)
    worksheet.set_column('E:E',17)

    worksheet.merge_range('A1:E1',CompanyName,company_name_format)
    worksheet.merge_range('A2:E2',CompanyAddress,company_info_format)
    worksheet.merge_range('A3:E3',CompanyPhoneNumber,company_info_format)
    worksheet.write_row('A4',['']*5,empty_row_format)
    worksheet.merge_range('A6:E6','Internet Banking',sheet_title_format)
    worksheet.merge_range('A7:E7',f'Date: {file_date}',sub_title_format)

    worksheet.write_row('A9',['Số tài khoản','Ngân hàng','Tài khoản ngân hàng','Nội dung','Tiền nộp'],headers_root_format)
    worksheet.write_column('A10',IB['SoTaiKhoan'],text_root_center_format)
    worksheet.write_column('B10',IB['NganHang'],text_root_center_format)
    worksheet.write_column('C10',IB['TaiKhoanNganHang'],text_root_center_format)
    worksheet.write_column('D10',IB['NoiDungIB'],text_root_left_format)
    worksheet.write_column('E10',IB['TienNop'],number_ib_format)

    worksheet.merge_range(f'A{IB.shape[0]+10}:D{IB.shape[0]+10}','Tổng',headers_root_format)
    worksheet.write(f'E{IB.shape[0]+10}',f'=SUM(E10:E{IB.shape[0]+9})',number_ib_format)

    ############################################################
    ############################################################
    ############################################################
    
    worksheet = workbook.add_worksheet('RCI1025')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',15)
    worksheet.set_column('B:B',9)
    worksheet.set_column('C:C',17)
    worksheet.set_column('D:D',50)
    worksheet.set_column('E:E',17)

    worksheet.merge_range('A1:E1',CompanyName,company_name_format)
    worksheet.merge_range('A2:E2',CompanyAddress,company_info_format)
    worksheet.merge_range('A3:E3',CompanyPhoneNumber,company_info_format)
    worksheet.write_row('A4',['']*5,empty_row_format)
    worksheet.merge_range('A6:E6','RCI1025 (Flex) - Mã giao dịch: 1141',sheet_title_format)
    worksheet.merge_range('A7:E7',f'Date: {file_date}',sub_title_format)

    worksheet.write_row('A9',['Số tài khoản','Ngân hàng','Tài khoản ngân hàng','Nội dung','Tiền nộp'],headers_root_format)
    worksheet.write_column('A10',RCI1025['SoTaiKhoan'],text_root_center_format)
    worksheet.write_column('B10',RCI1025['NganHang'],text_root_center_format)
    worksheet.write_column('C10',RCI1025['TaiKhoanNganHang'],text_root_center_format)
    worksheet.write_column('D10',RCI1025['NoiDungFlex'],text_root_left_format)
    worksheet.write_column('E10',RCI1025['TienNop'],number_flex_format)

    worksheet.merge_range(f'A{RCI1025.shape[0]+10}:D{RCI1025.shape[0]+10}','Tổng',headers_root_format)
    worksheet.write(f'E{RCI1025.shape[0]+10}',f'=SUM(E10:E{RCI1025.shape[0]+9})',number_flex_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')


