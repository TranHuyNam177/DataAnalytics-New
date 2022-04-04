from implementation import TaskMonitor
from request.stock import *
from news_collector import newsts


@TaskMonitor
def NewsSS(num_hours:int=48):
    """
    This method runs all functions in module news_collector.newsts (try till success)
    and exports all resulted DataFrames to a single excel file in the specified folder
    for daily usage of TS. This function is called in a higher-level module and
    automatically run on a daily basis
  
    :param num_hours: number of hours in the past that's in our concern
  
    :return: None
    """

    now = dt.date.today()
    date_string = now.strftime('%Y%m%d')

    path = r'C:\Users\hiepdang\Shared Folder\Trading Service\News Update'
    file_name = f'{date_string}_news.xlsx'
    file_path = fr'{path}\{file_name}'

    while True:
        try:
            vsd_TCPH = newsts.vsd.tinTCPH(num_hours)
            break
        except newsts.vsd.ignored_exceptions:
            vsd_TCPH = pd.DataFrame()
            continue
        except newsts.NoNewsFound:
            vsd_TCPH = pd.DataFrame()
            break

    while True:
        try:
            vsd_TVLK = newsts.vsd.tinTVLK(num_hours)
            break
        except newsts.vsd.ignored_exceptions:
            vsd_TVLK = pd.DataFrame()
            continue
        except newsts.NoNewsFound:
            vsd_TVLK = pd.DataFrame()
            break

    while True:
        try:
            hnx_tintuso = newsts.hnx.tintuso(num_hours)
            break
        except newsts.hnx.ignored_exceptions:
            hnx_tintuso = pd.DataFrame()
            continue
        except newsts.NoNewsFound:
            hnx_tintuso = pd.DataFrame()
            break

    while True:
        try:
            hose_tinTCNY = newsts.hose.tinTCNY(num_hours)
            break
        except newsts.hose.ignored_exceptions:
            hose_tinTCNY = pd.DataFrame()
            continue
        except newsts.NoNewsFound:
            hose_tinTCNY = pd.DataFrame()
            break

    while True:
        try:
            hose_tinCW = newsts.hose.tinCW(num_hours)
            break
        except newsts.hose.ignored_exceptions:
            hose_tinCW = pd.DataFrame()
            continue
        except newsts.NoNewsFound:
            hose_tinCW = pd.DataFrame()
            break

    writer = pd.ExcelWriter(file_path,engine='xlsxwriter')
    workbook = writer.book
    vsd_tinTCPH_sheet = workbook.add_worksheet('vsd_tinTCPH')
    vsd_tinTVLK_sheet = workbook.add_worksheet('vsd_tinTVLK')
    hnx_tintuso_sheet = workbook.add_worksheet('hnx_tintuso')
    hose_tinTCNY_sheet = workbook.add_worksheet('hose_tinTCNY')
    hose_tinCW_sheet = workbook.add_worksheet('hose_tinCW')

    header_fmt = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'bold':True,
            'border':1,
            'text_wrap':True,
        }
    )
    regular_fmt = workbook.add_format(
        {
            'valign':'vcenter',
            'border':1,
            'text_wrap':True,
        }
    )
    time_fmt = workbook.add_format(
        {
            'num_format':'dd/mm/yyyy hh:mm:ss',
            'valign':'vcenter',
            'border':1,
            'text_wrap':True,
        }
    )
    vsd_tinTCPH_sheet.hide_gridlines(option=2)
    vsd_tinTVLK_sheet.hide_gridlines(option=2)
    hnx_tintuso_sheet.hide_gridlines(option=2)
    hose_tinTCNY_sheet.hide_gridlines(option=2)
    hose_tinCW_sheet.hide_gridlines(option=2)

    # --------- Write vsd_tinTCPH
    vsd_tinTCPH_sheet.set_column('A:A',18)
    vsd_tinTCPH_sheet.set_column('B:B',12)
    vsd_tinTCPH_sheet.set_column('C:C',25)
    vsd_tinTCPH_sheet.set_column('D:D',90)
    vsd_tinTCPH_sheet.set_column('E:E',8)
    if not vsd_TCPH.empty:
        vsd_tinTCPH_sheet.write_row('A1',vsd_TCPH.columns,header_fmt)
        vsd_tinTCPH_sheet.write_column('A2',vsd_TCPH['Thời gian'],time_fmt)
        vsd_tinTCPH_sheet.write_column('B2',vsd_TCPH['Mã cổ phiếu'],regular_fmt)
        vsd_tinTCPH_sheet.write_column('C2',vsd_TCPH['Lý do'],regular_fmt)
        vsd_tinTCPH_sheet.write_column('D2',vsd_TCPH['Nội dung'],regular_fmt)
        vsd_tinTCPH_sheet.write_column('E2',vsd_TCPH['Link'],regular_fmt)

    # --------- Write vsd_tinTVLK
    vsd_tinTVLK_sheet.set_column('A:A',18)
    vsd_tinTVLK_sheet.set_column('B:B',25)
    vsd_tinTVLK_sheet.set_column('C:C',30)
    vsd_tinTVLK_sheet.set_column('D:D',90)
    vsd_tinTVLK_sheet.set_column('E:E',8)
    if not vsd_TVLK.empty:
        vsd_tinTVLK_sheet.write_row('A1',vsd_TVLK.columns,header_fmt)
        vsd_tinTVLK_sheet.write_column('A2',vsd_TVLK['Thời gian'],time_fmt)
        vsd_tinTVLK_sheet.write_column('B2',vsd_TVLK['Mã cổ phiếu / chứng quyền'],regular_fmt)
        vsd_tinTVLK_sheet.write_column('C2',vsd_TVLK['Lý do, mục đích'],regular_fmt)
        vsd_tinTVLK_sheet.write_column('D2',vsd_TVLK['Nội dung'],regular_fmt)
        vsd_tinTVLK_sheet.write_column('E2',vsd_TVLK['Link'],regular_fmt)

    # --------- Write hnx_tintuso
    hnx_tintuso_sheet.set_column('A:A',18)
    hnx_tintuso_sheet.set_column('B:B',8)
    hnx_tintuso_sheet.set_column('C:C',50)
    hnx_tintuso_sheet.set_column('D:D',80)
    hnx_tintuso_sheet.set_column('E:E',8)
    if not hnx_tintuso.empty:
        hnx_tintuso_sheet.write_row('A1',hnx_tintuso.columns,header_fmt)
        hnx_tintuso_sheet.write_column('A2',hnx_tintuso['Thời gian'],time_fmt)
        hnx_tintuso_sheet.write_column('B2',hnx_tintuso['Mã CK'],regular_fmt)
        hnx_tintuso_sheet.write_column('C2',hnx_tintuso['Tiêu đề'],regular_fmt)
        hnx_tintuso_sheet.write_column('D2',hnx_tintuso['Nội dung'],regular_fmt)
        hnx_tintuso_sheet.write_column('E2',hnx_tintuso['File đính kèm'],regular_fmt)

    # --------- Write hose_tintonghop
    hose_tinTCNY_sheet.set_column('A:A',18)
    hose_tinTCNY_sheet.set_column('B:B',12)
    hose_tinTCNY_sheet.set_column('C:C',40)
    hose_tinTCNY_sheet.set_column('D:D',80)
    hose_tinTCNY_sheet.set_column('E:E',13)
    if not hose_tinTCNY.empty:
        hose_tinTCNY_sheet.write_row('A1',hose_tinTCNY.columns,header_fmt)
        hose_tinTCNY_sheet.write_column('A2',hose_tinTCNY['Thời gian'],time_fmt)
        hose_tinTCNY_sheet.write_column('B2',hose_tinTCNY['Mã cổ phiếu'],regular_fmt)
        hose_tinTCNY_sheet.write_column('C2',hose_tinTCNY['Tiêu đề'],regular_fmt)
        hose_tinTCNY_sheet.write_column('D2',hose_tinTCNY['Nội dung'],regular_fmt)
        hose_tinTCNY_sheet.write_column('E2',hose_tinTCNY['File đính kèm'],regular_fmt)

    # --------- Write hose_tinCW
    hose_tinCW_sheet.set_column('A:A',18)
    hose_tinCW_sheet.set_column('B:B',45)
    hose_tinCW_sheet.set_column('C:C',90)
    hose_tinCW_sheet.set_column('D:D',15)
    hose_tinCW_sheet.set_column('E:E',15)
    if not hose_tinCW.empty:
        hose_tinCW_sheet.write_row('A1',hose_tinCW.columns,header_fmt)
        hose_tinCW_sheet.write_column('A2',hose_tinCW['Thời gian'],time_fmt)
        hose_tinCW_sheet.write_column('B2',hose_tinCW['Tiêu đề'],regular_fmt)
        hose_tinCW_sheet.write_column('C2',hose_tinCW['Nội dung'],regular_fmt)
        hose_tinCW_sheet.write_column('D2',hose_tinCW['File đính kèm'],regular_fmt)
        hose_tinCW_sheet.write_column('E2',hose_tinCW['Link'],regular_fmt)

    writer.close()