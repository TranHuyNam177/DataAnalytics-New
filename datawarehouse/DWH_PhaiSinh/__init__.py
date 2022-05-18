from datawarehouse import connect_DWH_PhaiSinh, EXEC
from request import *

def UPDATE():

    """
    This function EXEC the stored procedures to update all tables
    """

    EXEC(connect_DWH_PhaiSinh,'spRunPhaiSinh')


def UPDATEBACKDATE(
    days: int
):
    """
    This function EXEC spRunPhaiSinhBack to update all table back date

    :param days: number of back days
    """

    EXEC(connect_DWH_PhaiSinh,'spRunPhaiSinhBack',Ngaylui=days)
