from automation import *

dept_folder = r'C:\Users\hiepdang\Shared Folder\Product'

def get_info(
    periodicity:str,
    run_time=dt.datetime.now(),
):

    run_year = run_time.year
    run_month = run_time.month

    """
    Mon >> run_weekday = 2
    Tue >> run_weekday = 3
    Wed >> run_weekday = 4
    Thu >> run_weekday = 5
    Fri >> run_weekday = 6
    Sat >> run_weekday = 7
    Sun >> run_weekday = 8
    """
    run_weekday = run_time.weekday() + 2

    # Calculate time for quarterly report
    if run_month in [1,2,3]:
        soq = dt.datetime(run_year-1,10,1)
        eoq = dt.datetime(run_year-1,12,31)
    elif run_month in [4,5,6]:
        soq = dt.datetime(run_year,1,1)
        eoq = dt.datetime(run_year,3,31)
    elif run_month in [7,8,9]:
        soq = dt.datetime(run_year,4,1)
        eoq = dt.datetime(run_year,6,30)
    else:
        soq = dt.datetime(run_year,7,1)
        eoq = dt.datetime(run_year,9,30)

    # Calculate time for monthly report
    if run_month==1:
        mreport_year = run_year - 1
        mreport_month = 12
    else:
        mreport_year = run_year
        mreport_month = run_month-1
    som = dt.datetime(mreport_year,mreport_month,1)
    eom = dt.datetime(run_year,run_month,1)-dt.timedelta(days=1)

    # Calculate time for weekly report
    if run_weekday in [2,3,4,5]:
        sow = run_time - dt.timedelta(days=run_weekday+5)
        eow = run_time - dt.timedelta(days=run_weekday+1)
    else:  # run_weekday in [6,7,8]
        sow = run_time - dt.timedelta(days=run_weekday-2)
        eow = run_time - dt.timedelta(days=run_weekday-6)

    # select name of the folder
    folder_mapper = {
        'daily':'BaoCaoNgay',
        'weekly':'BaoCaoTuan',
        'monthly':'BaoCaoThang',
        'quarterly':'BaoCaoQuy',
    }
    folder_name = folder_mapper[periodicity]

    # choose dates and period
    if periodicity.lower() == 'daily':
        start_date = run_time.strftime('%Y.%m.%d')
        end_date = start_date
        period = run_time.strftime('%Y.%m.%d')
    elif periodicity.lower() == 'weekly':
        start_date = sow.strftime('%Y.%m.%d')
        end_date = eow.strftime('%Y.%m.%d')
        start_str = convert_int(sow.day)
        end_str = f'{convert_int(eow.day)}.{convert_int(eow.month)}.{eow.year}'
        period = f'{start_str}-{end_str}'
    elif periodicity.lower() == 'monthly':
        start_date = som.strftime('%Y.%m.%d')
        end_date = eom.strftime('%Y.%m.%d')
        period = f'{convert_int(eom.month)}.{eom.year}'
    elif periodicity.lower() == 'quarterly':
        start_date = soq.strftime('%Y.%m.%d')
        end_date = eoq.strftime('%Y.%m.%d')
        quarter_mapper = {
            3:'Q1',
            6:'Q2',
            9:'Q3',
            12:'Q4',
        }
        quarter = quarter_mapper[eoq.month]
        period = f'{quarter}.{eoq.year}'
    else:
        raise ValueError('Invalid periodicity')

    result_as_dict = {
        'run_time':run_time,
        'start_date':start_date,
        'end_date':end_date,
        'period':period,
        'folder_name':folder_name,
    }

    return result_as_dict
