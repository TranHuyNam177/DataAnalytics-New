from datawarehouse import connect_DWH_CoSo, EXEC
from request import *

def UPDATE():

    """
    This function EXEC the stored procedures to update all tables
    """

    EXEC(connect_DWH_CoSo,'spRunCoSo')


def UPDATEBACKDATE(
    days:int
):

    """
    This function EXEC spRunCoSoBack to update all table back date

    :param days: number of back days
    """

    EXEC(connect_DWH_CoSo,'spRunCoSoBack',Ngaylui=days)


def CHECKBATCH(batchType:int) -> bool:

    """
    This function EXEC spbatch to see if batch job finishes.
    Return True if batch finishes, False if not finishes

    :param batchType: 1 for mid-day batch, 2 for end-day batch
    """

    todayString = dt.datetime.now().strftime('%Y-%m-%d')
    EXEC(connect_DWH_CoSo,'spbatch',FrDate=todayString,ToDate=todayString)
    batchTable = pd.read_sql(
        f"""
        SELECT * FROM [batch] WHERE [batch].[date] = '{todayString}' AND [batch].[batch_type] = {batchType}
        """,
        connect_DWH_CoSo,
    )

    if batchTable.shape[0] == 0:
        return False
    else:
        return True

