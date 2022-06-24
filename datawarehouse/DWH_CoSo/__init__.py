from datawarehouse import connect_DWH_CoSo, SYNC, AFTERBATCH
from request import *

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

