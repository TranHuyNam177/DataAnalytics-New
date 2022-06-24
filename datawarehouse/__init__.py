from request import *


def BATCHINSERT(
    conn,
    table:str,
    df:pd.DataFrame,
):
    """
    This function INSERT the whole batch of data from a pd.DataFrame to a particular [db].[table].
    Must make sure the data type of pd.DataFrame align with [db].[table], must fill missing values by None, can't insert
    tables with empty column
    More performant, less flexible than SEQUENTIALINSERT

    :param conn: connection object of the Database
    :param table: name of the table in the Database
    :param df: inserted pd.DataFrame

    :return: None
    """

    sqlStatement = f"INSERT INTO [{table}] ([{'],['.join(df.columns)}]) VALUES ({','.join(['?']*df.shape[1])})"
    cursor = conn.cursor()
    cursor.executemany(sqlStatement,df.values.tolist())
    cursor.commit()
    cursor.close()

def SEQUENTIALINSERT(
    conn,
    table:str,
    df:pd.DataFrame,
):
    """
    This function INSERT row by row a pd.DataFrame to a particular [db].[table].
    Must make sure the data type of pd.DataFrame align with [db].[table]. Able to treat missing values with NULL automatically.
    Less performant, more flexible than BATCHINSERT

    :param conn: connection object of the Database
    :param table: name of the table in the Database
    :param df: inserted pd.DataFrame

    :return: None
    """

    cursor = conn.cursor()
    for row in df.index:
        truncatedRow = df.loc[row].dropna()
        if not truncatedRow.shape[0]: # dòng rỗng
            continue
        sqlStatement = f"INSERT INTO [{table}] ([{'],['.join(truncatedRow.index)}]) VALUES ({','.join(['?']*truncatedRow.shape[0])})"
        cursor.execute(sqlStatement,truncatedRow.to_list())
        cursor.commit()
    cursor.close()

def DELETE(
    conn,
    table:str,
    where:str,
):
    """
    This function DELETE entire rows from a [db].[table] given a paticular WHERE clause.
    If WHERE = '', it completely clears all data from the [db].[table]

    :param conn: connection object of the Database
    :param table: name of the table in the Database
    :param where: WHERE clause in DELETE statement
    """
    if where == '':
        where = ''
    else:
        if not where.startswith('WHERE'):
            where = 'WHERE ' + where
    sqlStatement = f"DELETE FROM [{table}] {where}"
    cursor = conn.cursor()
    cursor.execute(sqlStatement)
    cursor.commit()
    cursor.close()


def DROPDUPLICATES(
    conn,
    table:str,
    *columns:str,
):
    """
    This function DELETE duplicates values from [db].[table] given a list of columns
    on which we check for duplicates

    (Hàm này có downside là không biết nó DELETE dòng nào trong các dòng trùng)

    :param conn: connection object of the Database
    :param table: name of the table in the Database
    :param columns: columns to check for duplicates
    """

    columnList = '[' + '],['.join(columns) + ']'
    sqlStatement = f"""
        WITH [tempTable] AS (
            SELECT *, ROW_NUMBER() OVER (PARTITION BY {columnList} ORDER BY {columnList}) [rowNum]
            FROM [{table}]
        )
        DELETE FROM [tempTable]
        WHERE [rowNum] > 1
    """
    cursor = conn.cursor()
    cursor.execute(sqlStatement)
    cursor.commit()
    cursor.close()


def EXEC(
    conn,
    sp:str,
    **params,
):

    """
    This function EXEC the specified stored procedure in SQL

    :param conn: connection object of the Database
    :param sp: name of the stored procedure in the Database
    :param params: parameters passed to the stored procedure

    Example: EXEC(connect_DWH_CoSo, 'spvrm6631', FrDate='2022-03-01', ToDate='2022-03-01')
    """

    cursor = conn.cursor()
    sqlStatement = f'SET NOCOUNT ON; EXEC {sp}'
    for k,v in params.items():
        sqlStatement += f" @{k} = '{v}',"
    sqlStatement = sqlStatement.rstrip(',')

    cursor.execute(sqlStatement)
    print(sqlStatement)
    cursor.commit()
    cursor.close()


def SYNC(
    conn,
    StoredProcedure:str,
    FrDate:str=None,
    ToDate:str=None,
    SoLanLui:str=None,
):

    """
    This function runs the underlying stored procedure through Web Service to sync the associated table

    :param conn: connection object of the Database
    :param StoredProcedure: name of the stored procedure
    :param FrDate: FrDate string
    :param ToDate: ToDate string
    :param SoLanLui: Số ngày lùi
    """

    dbName = conn.getinfo(pyodbc.SQL_DATABASE_NAME)
    if dbName == 'DWH-CoSo':
        action = 'RunStoreCoSo'
    elif dbName == 'DWH-PhaiSinh':
        action = 'RunStorePhaiSinh'
    else:
        raise ValueError('Invalid Connection Object')

    URL = 'http://192.168.6.103/RunStoreDWH/RunStore.asmx'
    hostName = '192.168.6.103'
    key = 'abcdef123456'
    body = f"""<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
            <soap:Body>
                <{action} xmlns="http://tempuri.org/">
                <StoreName>{StoredProcedure}</StoreName>
                <FrDate>{FrDate}</FrDate>
                <ToDate>{ToDate}</ToDate>
                <SoLanLui>{str(SoLanLui)}</SoLanLui>
                <Key>{key}</Key>
            </{action}>
            </soap:Body>
        </soap:Envelope>"""
    headers = {
        'Content-Type': 'text/xml; charset=utf-8',
        'Host': hostName,
        'SOAPAction': f'http://tempuri.org/{action}',
        'Content-Length': str(len(body))
    }
    r = requests.post(URL,data=body,headers=headers)
    if r.status_code != 200:
        raise RuntimeError('Không thể chạy Stored Procedure qua Web Service')
    if 'thành công' not in r.text:
        raise RuntimeError(r.text)


def CONFIG(
    conn,
    StoreName:str,
    **kwargs,
):

    """
    This function changes the config table that controls the running rules of stored procedures

    :param conn: connection object of the Database
    :param StoreName: name of the stored procedure
    :param kwargs: name of the attribute to change
    """

    dbName = conn.getinfo(pyodbc.SQL_DATABASE_NAME)
    cursor = conn.cursor()
    sqlStatement = f'USE [{dbName}]; UPDATE [StoreConfig] SET'
    for k,v in kwargs.items():
        sqlStatement += f" [{k}] = '{v}',"
    sqlStatement = sqlStatement.rstrip(',')
    sqlStatement += f" WHERE [StoreName] = '{StoreName}'"
    cursor.execute(sqlStatement)
    print(sqlStatement)
    cursor.commit()
    cursor.close()


def BDATE(
    date:str,
    bdays:int=0
) -> str:

    """
    This function return the business date before/after a certain business days
    since a given date

    :param date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd'
    :param bdays: allow positive integer (after) or negative integer (before)

    :return: string of type 'yyyy-mm-dd'
    """

    dateType = pd.read_sql(
        f"""
        SELECT [t].[Work]
        FROM [Date] [t]
        WHERE [t].[Date] = '{date}'
        """,
        connect_DWH_CoSo,
    ).squeeze()

    if dateType: # Ngày làm việc
        top = abs(bdays) + 1
    else: # Ngày nghỉ
        top = abs(bdays)

    if bdays < 0:
        return pd.read_sql(
            f"""
            WITH [FullTable] AS (
                SELECT TOP {top} [t].[Date]
                FROM [Date] [t]
                WHERE [t].[Work] = 1 AND [t].[Date] <= '{date}'
                ORDER BY [t].[Date] DESC
            )
            SELECT MIN([FullTable].[Date])
            FROM [FullTable]
            """,
            connect_DWH_CoSo,
        ).squeeze().strftime('%Y-%m-%d')
    elif bdays > 0:
        return pd.read_sql(
            f"""
            WITH [FullTable] AS (
                SELECT TOP {top} [t].[Date]
                FROM [Date] [t]
                WHERE [t].[Work] = 1 AND [t].[Date] >= '{date}'
                ORDER BY [t].[Date] ASC
            )
            SELECT MAX([FullTable].[Date])
            FROM [FullTable]
            """,
            connect_DWH_CoSo,
        ).squeeze().strftime('%Y-%m-%d')
    else:
        return f'{date[:4]}-{date[5:7]}-{date[-2:]}'


def LASTBATCH(conn,batchType:int):
    dbName = conn.getinfo(pyodbc.SQL_DATABASE_NAME)
    if dbName == 'DWH-CoSo':
        return pd.read_sql(
            f"""
            SELECT MAX([batch_end]) FROM [batch] 
            WHERE [batch_type] = {batchType}
            """,
            conn,
        ).squeeze().to_pydatetime()
    elif dbName == 'DWH-PhaiSinh':
        return pd.read_sql(
            f"""
            SELECT MAX([batch_end]) FROM [batch] 
            """,
            conn,
        ).squeeze().to_pydatetime()
    else:
        raise ValueError("Invalid Database")


def CHECKBATCH(conn,batchType:int) -> bool:

    """
    This function SYNC spbatch to see if batch job finishes.
    Return True if batch finishes, False if not finishes

    :param conn: connection object of the Database
    :param batchType: 1 for mid-day batch, 2 for end-day batch
    """

    todayString = dt.datetime.now().strftime('%Y-%m-%d')
    dbName = conn.getinfo(pyodbc.SQL_DATABASE_NAME)
    if dbName == 'DWH-CoSo':
        SYNC(conn,'spbatch',FrDate=todayString,ToDate=todayString)
        batchTable = pd.read_sql(
            f"""
            SELECT * FROM [batch] WHERE [batch].[date] = '{todayString}' AND [batch].[batch_type] = {batchType}
            """,
            conn,
        )
    elif dbName == 'DWH-PhaiSinh':
        SYNC(conn,'spbatch',FrDate=todayString,ToDate=todayString)
        batchTable = pd.read_sql(
            f"""
            SELECT * FROM [batch] WHERE [batch].[date] = '{todayString}'
            """,
            conn,
        )
    else:
        raise ValueError(f'Invalid Database')

    if batchTable.shape[0]:
        print(f"Batch done at: {batchTable.loc[batchTable.index[-1],'batch_end']}")
        return True
    else:
        return False


def AFTERBATCH(conn,batchType): # decorator

    """
    Pause the function till batch done
    """

    def wrapper(func): # nhận hàm ban đầu

        def decoratedFunction(*args,**kwargs):
            while True:
                if CHECKBATCH(conn,batchType):
                    break
                time.sleep(30)
            func(*args,**kwargs)

        return decoratedFunction # trả ra hàm decorated rồi

    return wrapper


def LASTSYNC(conn,tableName=None):

    """
    Last sync time of a table. If tableName = None, return last run time of RunCoSo, RunPhaiSinh
    """

    dbName = conn.getinfo(pyodbc.SQL_DATABASE_NAME)
    if dbName == 'DWH-CoSo':
        if tableName is None:
            return pd.read_sql(
                f"""
                SELECT MAX([EXEC_DATE]) FROM [ExecTaskLog] 
                WHERE [STATUS] = 'END' AND [DESCRIPTION] = 'RunCoSo'
                """,
                conn,
            ).squeeze().to_pydatetime()
        else:
            if tableName not in TableNames_DWH_CoSo.squeeze().to_list():
                raise ValueError("Invalid Table Name")
            return pd.read_sql(
                f"""
                SELECT MAX([EXEC_DATE]) FROM [ExecTaskLog] 
                WHERE [STATUS] = 'OK' AND [DESCRIPTION] = '{tableName}'
                """,
                conn,
            ).squeeze().to_pydatetime()
    elif dbName == 'DWH-PhaiSinh':
        if tableName is None:
            return pd.read_sql(
                f"""
                SELECT MAX(EXEC_DATE) FROM [ExecTaskLog] 
                WHERE [STATUS] = 'END' AND [DESCRIPTION] = 'RunPhaiSinh'
                """,
                conn,
            ).squeeze().to_pydatetime()
        else:
            if tableName not in TableNames_DWH_PhaiSinh.squeeze().to_list():
                raise ValueError("Invalid Table Name")
            return pd.read_sql(
                f"""
                SELECT MAX([EXEC_DATE]) FROM [ExecTaskLog] 
                WHERE [STATUS] = 'OK' AND [DESCRIPTION] = '{tableName}'
                """,
                conn,
            ).squeeze().to_pydatetime()
    else:
        raise ValueError("Invalid Database")


def CHECKSYNC(
    conn,
    batchType:int,
    tableName:str=None,
):

    """
    Check if the sync process has been done following the batch job. If tableName = None, return last run time of RunCoSo, RunPhaiSinh
    """

    lastBatchTime = LASTBATCH(conn,batchType)
    lastSyncTime = LASTSYNC(conn,tableName)

    return lastSyncTime.time() > lastBatchTime.time() and lastSyncTime.date() == lastBatchTime.date()


def NOTIFYSYNCSTATUSTODAY(db):

    check_time = dt.datetime.now()

    ignored_Tables = [
        'v_sqlrun',
        'ExcludeDays',
        'ExecTaskLog',
        'bank_account_list',
        'de083',
        'breakeven_price_portfolio',
        'storerun',
        'VW_GETSEACCOUNTROOM_DB',
        'V_GETSECMARGINASSET',
        'V_GETSECMARGINRELEASE_MST1',
        'V_GETSECMARGINRELEASE_MST2',
        'V_GETSECMARGINRELEASE',
        'VW_MR0004',
        'vw_getsecmargindetail_detail',
        'vw_getsecmargindetail',
        'vw_mr0001_all',
        'BankTransactionHistory',
        'BankCurrentBalance',
        'BankDepositBalance',
    ]

    if db == 'DWH-CoSo':
        conn = connect_DWH_CoSo
        db_Tables = TableNames_DWH_CoSo.squeeze()
        prefix = '[DWH-CoSo]'
        description = 'RunCoSo'
        since = dt.datetime.now()-dt.timedelta(minutes=30)

    elif db == 'DWH-PhaiSinh':
        conn = connect_DWH_PhaiSinh
        db_Tables = TableNames_DWH_PhaiSinh.squeeze()
        prefix = '[DWH-PhaiSinh]'
        description = 'RunPhaiSinh'
        since = dt.datetime.now()-dt.timedelta(minutes=30)

    else:
        raise ValueError('The module currently checks DWH-CoSo or DWH-PhaiSinh only')

    db_Tables = db_Tables.loc[~db_Tables.isin(ignored_Tables)]

    # check xem n phút vừa rồi có chạy chưa (n được quy định ở trên)
    run_check = pd.read_sql(
        f"""
        SELECT [ExecTaskLog].[EXEC_DATE] [TIME]
        FROM [ExecTaskLog]
        WHERE [ExecTaskLog].[EXEC_DATE] >= '{since.strftime("%Y-%m-%d %H:%M:%S")}'
            AND [ExecTaskLog].[STATUS] = 'START'
        """,
        conn,
    ).squeeze(axis=0)

    if run_check.empty: # nếu chưa chạy
        NotRunTables = db_Tables
    else: # nếu chạy rồi
        # check xem nếu chạy rồi thì có đủ bảng chưa
        run_Tables = pd.read_sql(
            f"""
            SELECT [ExecTaskLog].[DESCRIPTION]
            FROM [ExecTaskLog]
            WHERE [ExecTaskLog].[STATUS] = 'OK'
            AND [ExecTaskLog].[EXEC_DATE] >= (
                SELECT MAX([ExecTaskLog].[EXEC_DATE]) [Time]
                FROM [ExecTaskLog]
                WHERE [ExecTaskLog].[STATUS] = 'START'
                AND [ExecTaskLog].[DESCRIPTION] = '{description}'
            )
            """,
            conn,
        ).squeeze()
        NotRunTables = db_Tables.loc[~db_Tables.isin(run_Tables)]

    missing_Tables = NotRunTables.to_frame()
    missing_Tables.columns = ['Missing Tables']
    missing_Tables = prefix + '.[' + missing_Tables + ']'
    html_table = missing_Tables.to_html(index=False)
    html_table = html_table.replace('<tr>','<tr align="center">')  # center columns
    html_table = html_table.replace('border="1"','border="1" style="border-collapse:collapse"')  # make thinner borders


    if missing_Tables.empty:
        content = f"""
            <p style="font-family:Times New Roman; font-size:100%"><i>
                ĐỦ SỐ LƯỢNG BẢNG
            </i></p>
        """
    else:
        # HTML table for email
        content = html_table

    # Send mail
    body = f"""
    <html>
        <head></head>
        <body>
            {content}
            <p style="font-family:Times New Roman; font-size:90%"><i>
                -- Generated by Reporting System
            </i></p>
        </body>
    </html>
    """

    outlook = Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mapi = outlook.GetNamespace("MAPI")

    for account in mapi.Accounts:
        print(f"Account {account.DeliveryStore.DisplayName} is being logged")

    mail.To = 'hiepdang@phs.vn; tupham@phs.vn'
    mail.Subject = f"{prefix} Missing Tables {check_time.strftime('%Y-%m-%d %H:%M:%S')}"
    mail.HTMLBody = body
    mail.Send()


def NOTIFYSYNCSTATUSBACKDATE(db):

    """
    Được chạy vào:
        - Mon-Fri: 12:15, 18:00, 21:00
        - Sat-Sun: 01:00
    """

    check_time = dt.datetime.now()

    if db == 'DWH-CoSo':
        conn = connect_DWH_CoSo
        prefix = '[DWH-CoSo]'
        description = 'RunCoSoLui'
        since = dt.datetime.now()-dt.timedelta(minutes=60)

    elif db == 'DWH-PhaiSinh':
        conn = connect_DWH_PhaiSinh
        prefix = '[DWH-PhaiSinh]'
        description = 'RunPhaiSinhLui'
        since = dt.datetime.now()-dt.timedelta(minutes=60)

    else:
        raise ValueError('The module currently checks DWH-CoSo or DWH-PhaiSinh only')

    hour = dt.datetime.now().hour
    if 22 <= hour <= 24 or 0 <= hour <= 5:
        days = 4
    else:
        days = 1

    # check xem n phút vừa rồi có chạy chưa (n được quy định ở trên)
    checkTable = pd.read_sql(
        f"""
        SELECT [ExecTaskLog].[STATUS]
        FROM [ExecTaskLog]
        WHERE [ExecTaskLog].[EXEC_DATE] >= '{since.strftime("%Y-%m-%d %H:%M:%S")}'
            AND [ExecTaskLog].[DESCRIPTION] = '{description}'
        """,
        conn,
    ).squeeze().value_counts().reindex(['START','END']).fillna(0)

    if checkTable['START'] < days or checkTable['END'] < days:
        content = f"""
            <p style="font-family:Times New Roman; font-size:100%"><i>
                SYNC BACKDATE THẤT BẠI
            </i></p>
        """
    else:
        content = f"""
            <p style="font-family:Times New Roman; font-size:100%"><i>
                SYNC BACKDATE THÀNH CÔNG
            </i></p>
        """

    body = f"""
    <html>
        <head></head>
        <body>
            {content}
            <p style="font-family:Times New Roman; font-size:90%"><i>
                -- Generated by Reporting System
            </i></p>
        </body>
    </html>
    """

    outlook = Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mapi = outlook.GetNamespace("MAPI")

    for account in mapi.Accounts:
        print(f"Account {account.DeliveryStore.DisplayName} is being logged")

    mail.To = 'hiepdang@phs.vn; tupham@phs.vn'
    mail.Subject = f"{prefix} Check Update BackDate {check_time.strftime('%Y-%m-%d %H:%M:%S')}"
    mail.HTMLBody = body
    mail.Send()
