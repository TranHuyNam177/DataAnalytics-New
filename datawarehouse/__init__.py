import time

from request import *


def INSERT(
    conn,
    table:str,
    df:pd.DataFrame,
):
    """
    This function INSERT a pd.DataFrame to a particular [db].[table].
    Must make sure the order / data type of pd.DataFrame align with [db].[table]

    :param conn: connection object of the Database
    :param table: name of the table in the Database
    :param df: inserted pd.DataFrame

    :return: None
    """

    sqlStatement = f"INSERT INTO [{table}] VALUES ({','.join(['?']*df.shape[1])})"
    cursor = conn.cursor()
    cursor.executemany(sqlStatement,df.values.tolist())
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


def DROP_DUPLICATES(
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

    sqlStatement = f'SET NOCOUNT ON; EXEC {sp}'
    for k,v in params.items():
        sqlStatement += f" @{k} = '{v}',"

    sqlStatement = sqlStatement.rstrip(',')
    print(sqlStatement)

    cursor = conn.cursor()
    cursor.execute(sqlStatement)
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


def CHECKBATCH(conn,batchType:int) -> bool:

    """
    This function EXEC spbatch to see if batch job finishes.
    Return True if batch finishes, False if not finishes

    :param conn: connection object of the Database
    :param batchType: 1 for mid-day batch, 2 for end-day batch
    """

    todayString = dt.datetime.now().strftime('%Y-%m-%d')
    dbName = conn.getinfo(pyodbc.SQL_DATABASE_NAME)
    if dbName == 'DWH-CoSo':
        EXEC(conn,'spbatch',FrDate=todayString,ToDate=todayString)
        batchTable = pd.read_sql(
            f"""
            SELECT * FROM [batch] WHERE [batch].[date] = '{todayString}' AND [batch].[batch_type] = {batchType}
            """,
            conn,
        )
    elif dbName == 'DWH-PhaiSinh':
        EXEC(conn,'spbatch',FrDate=todayString,ToDate=todayString)
        batchTable = pd.read_sql(
            f"""
            SELECT * FROM [batch] WHERE [batch].[date] = '{todayString}'
            """,
            conn,
        )
    else:
        raise ValueError(f'Invalid database: {dbName}')

    if batchTable.shape[0]:
        print(f"Batch done at: {batchTable.loc[batchTable.index[-1],'batch_end']}")
        return True
    else:
        return False

def AfterBatch(conn,batchType): # decorator

    """
    Pause the function till batch done
    """

    def wrapper(func):

        while True:
            checkResult = CHECKBATCH(conn,batchType)
            if checkResult:
                break
            time.sleep(15)

        func(*args,**kwargs)

    return wrapper


