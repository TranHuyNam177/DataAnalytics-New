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


