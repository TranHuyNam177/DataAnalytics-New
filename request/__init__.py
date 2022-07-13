from function import *

with open(r'C:\Users\namtran\Desktop\Passwords\DataBase\DataBase.txt') as file:
    user, password, _ = file.readlines()
    user = user.replace('\n','')
    password = password.replace('\n','')

# Risk Database Information
driver_RMD = '{SQL Server}'
server_RMD = 'SRV-RPT'
db_RMD = 'RiskDb'

connect_RMD = pyodbc.connect(
    f'Driver={driver_RMD};'
    f'Server={server_RMD};'
    f'Database={db_RMD};'
    f'uid={user};'
    f'pwd={password}'
)
TableNames_RMD = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect_RMD,
)

# # DWH-Base Database Information
# driver_DWH_Base = '{SQL Server}'
# server_DWH_Base = 'SRV-RPT'
# db_DWH_Base = 'DWH-Base'
# connect_DWH_Base = pyodbc.connect(
#     f'Driver={driver_DWH_Base};'
#     f'Server={server_DWH_Base};'
#     f'Database={db_DWH_Base};'
#     f'uid={user};'
#     f'pwd={password}'
# )
# TableNames_DWH_Base = pd.read_sql(
#     'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
#     connect_DWH_Base
# )


# DWH-ThiTruong Database Information
driver_DWH_ThiTruong = '{SQL Server}'
server_DWH_ThiTruong = 'SRV-RPT'
db_DWH_ThiTruong = 'DWH-ThiTruong'
connect_DWH_ThiTruong = pyodbc.connect(
    f'Driver={driver_DWH_ThiTruong};'
    f'Server={server_DWH_ThiTruong};'
    f'Database={db_DWH_ThiTruong};'
    f'uid={user};'
    f'pwd={password}'
)
TableNames_DWH_ThiTruong = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect_DWH_ThiTruong
)

# DWH-CoSo Database Information
driver_DWH_CoSo = '{SQL Server}'
server_DWH_CoSo = 'SRV-RPT'
db_DWH_CoSo = 'DWH-CoSo'
connect_DWH_CoSo = pyodbc.connect(
    f'Driver={driver_DWH_CoSo};'
    f'Server={server_DWH_CoSo};'
    f'Database={db_DWH_CoSo};'
    f'uid={user};'
    f'pwd={password}'
)
TableNames_DWH_CoSo = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect_DWH_CoSo
)

# DWH-PhaiSinh Database Information
driver_DWH_PhaiSinh = '{SQL Server}'
server_DWH_PhaiSinh = 'SRV-RPT'
db_DWH_PhaiSinh = 'DWH-PhaiSinh'
connect_DWH_PhaiSinh = pyodbc.connect(
    f'Driver={driver_DWH_PhaiSinh};'
    f'Server={server_DWH_PhaiSinh};'
    f'Database={db_DWH_PhaiSinh};'
    f'uid={user};'
    f'pwd={password}'
)
TableNames_DWH_PhaiSinh = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect_DWH_PhaiSinh
)
