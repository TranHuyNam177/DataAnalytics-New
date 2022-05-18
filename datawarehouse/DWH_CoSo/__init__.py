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

