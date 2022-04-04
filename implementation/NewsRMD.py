from implementation import TaskMonitor
from request.stock import *
from news_collector import newsrmd


@TaskMonitor
def NewsRMD():
    """
    This method runs all functions in module news_collector.newsrmd (try till success)
    and exports all resulted DataFrames to a single excel file in the specified folder
    for daily usage of RMD. This function is called in a higher-level module and
    automatically run on a daily basis
  
    :return: None
    """
    now = dt.datetime.now()
    time_string = now.strftime('%Y%m%d_@%H%M')
    if dt.time(hour=0,minute=0,second=0)<=now.time()<=dt.time(hour=11,minute=59,second=59):
        previous_bdate = bdate(now.strftime('%Y-%m-%d'),-1)
        time_point = dt.datetime.strptime(previous_bdate,'%Y-%m-%d')
        time_point = time_point.replace(hour=18,minute=0,second=0,microsecond=0)
    if dt.time(hour=12,minute=0,second=0)<=now.time()<=dt.time(hour=23,minute=59,second=59):
        time_point = now.replace(hour=10,minute=0,second=0,microsecond=0)

    path = r'\\192.168.10.101\phs-storge-2018' \
           r'\RiskManagementDept\RMD_Data' \
           r'\Luu tru van ban\RMC Meeting 2018' \
           r'\00. Meeting minutes\Data\News Update'
    file_name = f'{time_string}.xlsx'
    file_path = fr'{path}\{file_name}'

    while True:
        try:
            vsd_TCPH = newsrmd.vsd.tinTCPH()
            break
        except newsrmd.ignored_exceptions:
            vsd_TCPH = pd.DataFrame()
            continue
        except newsrmd.NoNewsFound:
            vsd_TCPH = pd.DataFrame()
            break

    while True:
        try:
            vsd_TVBT = newsrmd.vsd.tinTVBT()
            break
        except newsrmd.ignored_exceptions:
            vsd_TVBT = pd.DataFrame()
            continue
        except newsrmd.NoNewsFound:
            vsd_TVBT = pd.DataFrame()
            break

    while True:
        try:
            hnx_TCPH = newsrmd.hnx.tinTCPH()
            break
        except newsrmd.ignored_exceptions:
            hnx_TCPH = pd.DataFrame()
            continue
        except newsrmd.NoNewsFound:
            hnx_TCPH = pd.DataFrame()
            break

    while True:
        try:
            hnx_tintuso = newsrmd.hnx.tintuso()
            break
        except newsrmd.ignored_exceptions:
            hnx_tintuso = pd.DataFrame()
            continue
        except newsrmd.NoNewsFound:
            hnx_tintuso = pd.DataFrame()
            break

    while True:
        try:
            hose_tintonghop = newsrmd.hose.tintonghop()
            break
        except newsrmd.ignored_exceptions:
            hose_tintonghop = pd.DataFrame()
            continue
        except newsrmd.NoNewsFound:
            hose_tintonghop = pd.DataFrame()
            break

    writer = pd.ExcelWriter(file_path,engine='xlsxwriter')
    workbook = writer.book
    vsd_TCPH_sheet = workbook.add_worksheet('vsd_TCPH')
    vsd_TVBT_sheet = workbook.add_worksheet('vsd_TVBT')
    hnx_TCPH_sheet = workbook.add_worksheet('hnx_TCPH')
    hnx_tintuso_sheet = workbook.add_worksheet('hnx_tintuso')
    hose_tintonghop_sheet = workbook.add_worksheet('hose_tintonghop')

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
    highlight_regular_fmt = workbook.add_format(
        {
            'valign':'vcenter',
            'border':1,
            'bg_color':'#FFF024',
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
    highlight_time_fmt = workbook.add_format(
        {
            'num_format':'dd/mm/yyyy hh:mm:ss',
            'valign':'vcenter',
            'border':1,
            'bg_color':'#FFF024',
            'text_wrap':True,
        }
    )
    vsd_TCPH_sheet.hide_gridlines(option=2)
    vsd_TVBT_sheet.hide_gridlines(option=2)
    hnx_TCPH_sheet.hide_gridlines(option=2)
    hnx_tintuso_sheet.hide_gridlines(option=2)
    hose_tintonghop_sheet.hide_gridlines(option=2)

    check_dt = lambda dt_time:True if dt_time>time_point else False

    vsd_TCPH_sheet.set_column('A:A',18)
    vsd_TCPH_sheet.set_column('B:D',7)
    vsd_TCPH_sheet.set_column('E:E',50)
    vsd_TCPH_sheet.set_column('F:F',7)
    if not vsd_TCPH.empty:
        mask_vsd_TCPH = vsd_TCPH['Thời gian'].map(check_dt)
        vsd_TCPH_sheet.write_row('A1',vsd_TCPH.columns,header_fmt)
        for col in range(vsd_TCPH.shape[1]):
            for row in range(vsd_TCPH.shape[0]):
                if col==0 and mask_vsd_TCPH.loc[row]==True:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],highlight_time_fmt)
                elif col!=0 and mask_vsd_TCPH.loc[row]==True:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],highlight_regular_fmt)
                elif col==0 and mask_vsd_TCPH.loc[row]==False:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],time_fmt)
                else:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],regular_fmt)

    vsd_TVBT_sheet.set_column('A:A',18)
    vsd_TVBT_sheet.set_column('B:B',90)
    vsd_TVBT_sheet.set_column('C:C',7)
    vsd_TVBT_sheet.write_row('A1',vsd_TVBT.columns,header_fmt)
    if not vsd_TVBT.empty:
        mask_vsd_TVBT = vsd_TVBT['Thời gian'].map(check_dt)
        for col in range(vsd_TVBT.shape[1]):
            for row in range(vsd_TVBT.shape[0]):
                if col==0 and mask_vsd_TVBT.loc[row]==True:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],highlight_time_fmt)
                elif col!=0 and mask_vsd_TVBT.loc[row]==True:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],highlight_regular_fmt)
                elif col==0 and mask_vsd_TVBT.loc[row]==False:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],time_fmt)
                else:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],regular_fmt)

    hnx_TCPH_sheet.set_column('A:A',18)
    hnx_TCPH_sheet.set_column('B:D',7)
    hnx_TCPH_sheet.set_column('E:E',30)
    hnx_TCPH_sheet.set_column('F:F',120)
    hnx_TCPH_sheet.set_column('G:G',7)
    hnx_TCPH_sheet.write_row('A1',hnx_TCPH.columns,header_fmt)
    if not hnx_TCPH.empty:
        mask_hnx_TCPH = hnx_TCPH['Thời gian'].map(check_dt)
        for col in range(hnx_TCPH.shape[1]):
            for row in range(hnx_TCPH.shape[0]):
                if col==0 and mask_hnx_TCPH.loc[row]==True:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],highlight_time_fmt)
                elif col!=0 and mask_hnx_TCPH.loc[row]==True:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],highlight_regular_fmt)
                elif col==0 and mask_hnx_TCPH.loc[row]==False:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],time_fmt)
                else:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],regular_fmt)

    hnx_tintuso_sheet.set_column('A:A',18)
    hnx_tintuso_sheet.set_column('B:D',7)
    hnx_tintuso_sheet.set_column('E:E',70)
    hnx_tintuso_sheet.set_column('F:F',7)
    hnx_tintuso_sheet.write_row('A1',hnx_tintuso,header_fmt)
    if not hnx_tintuso.empty:
        mask_hnx_tintuso = hnx_tintuso['Thời gian'].map(check_dt)
        for col in range(hnx_tintuso.shape[1]):
            for row in range(hnx_tintuso.shape[0]):
                if col==0 and mask_hnx_tintuso.loc[row]==True:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],highlight_time_fmt)
                elif col!=0 and mask_hnx_tintuso.loc[row]==True:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],highlight_regular_fmt)
                elif col==0 and mask_hnx_tintuso.loc[row]==False:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],time_fmt)
                else:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],regular_fmt)

    hose_tintonghop_sheet.set_column('A:A',18)
    hose_tintonghop_sheet.set_column('B:D',7)
    hose_tintonghop_sheet.set_column('E:E',70)
    hose_tintonghop_sheet.set_column('F:F',80)
    hose_tintonghop_sheet.write_row('A1',hose_tintonghop,header_fmt)
    if not hose_tintonghop.empty:
        mask_hose_tintonghop = hose_tintonghop['Thời gian'].map(check_dt)
        for col in range(hose_tintonghop.shape[1]):
            for row in range(hose_tintonghop.shape[0]):
                if col==0 and mask_hose_tintonghop.iloc[row]==True:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],highlight_time_fmt)
                elif col!=0 and mask_hose_tintonghop.iloc[row]==True:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],highlight_regular_fmt)
                elif col==0 and mask_hose_tintonghop.iloc[row]==False:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],time_fmt)
                else:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],regular_fmt)

    writer.close()
