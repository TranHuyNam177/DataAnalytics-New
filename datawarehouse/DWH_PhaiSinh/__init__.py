from datawarehouse import connect_DWH_PhaiSinh, SYNC, AFTERBATCH
from request import *

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

