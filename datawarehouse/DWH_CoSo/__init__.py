from request import connect_DWH_CoSo
from datawarehouse import SYNC

def SYNCTODAY():

    """
    This function sync all tables from OLTP with today's data
    """

    SYNC(connect_DWH_CoSo,StoredProcedure='spRunCoSo')


def SYNCBACKDATE(
    days:int
):

    """
    This function sync all tables from OLTP back date

    :param days: number of back days
    """

    SYNC(connect_DWH_CoSo,StoredProcedure='spRunCoSoBack',SoLanLui=days)

