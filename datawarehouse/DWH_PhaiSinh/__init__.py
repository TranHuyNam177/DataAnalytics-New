from request import connect_DWH_PhaiSinh
from datawarehouse import SYNC

def SYNCTODAY():

    """
    This function sync all tables from OLTP with today's data
    """

    SYNC(connect_DWH_PhaiSinh,StoredProcedure='spRunPhaiSinh')


def SYNCBACKDATE(
    days:int
):

    """
    This function sync all tables from OLTP back date

    :param days: number of back days
    """

    SYNC(connect_DWH_PhaiSinh,StoredProcedure='spRunPhaiSinhBack',SoLanLui=days)

