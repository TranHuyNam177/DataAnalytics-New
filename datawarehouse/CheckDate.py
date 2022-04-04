from request import *

def run(db):

    endDate = dt.datetime.now().strftime('%Y-%m-%d')

    if db == 'DWH-CoSo':
        conn = connect_DWH_CoSo
        sqlTables = TableNames_DWH_CoSo['TABLE_NAME']
        startDate = '2018-01-01'
    elif db == 'DWH-PhaiSinh':
        conn = connect_DWH_PhaiSinh
        sqlTables = TableNames_DWH_PhaiSinh['TABLE_NAME']
        startDate = '2020-12-24'
    else:
        raise ValueError("db must be either 'DWH-CoSo' or 'DWH-PhaiSinh'")

    output = pd.DataFrame(columns=sqlTables,index=pd.date_range(startDate,endDate))
    for table in sqlTables:
        try:
            sqlStatement = f'SELECT TOP 0 * FROM [{table}]'
            cols= pd.read_sql(sqlStatement,conn).columns.tolist()
            date_col = [col for col in cols if 'date' in col][0]
            sqlStatement = f'SELECT DISTINCT {date_col} FROM [{table}] ORDER BY {date_col}'
            d = pd.read_sql(sqlStatement,conn,index_col=date_col)
            d['check'] = 'OK'
            output[table] = d['check']
        except (IndexError,):
            pass

    output = output.fillna('')

    return output