from automation.trading_service import *

driver = '{SQL Server}'
server = 'SRV-RPT'
database = 'RiskDb'
user_id = 'hiep'
user_password = '5B7Cv6huj2FcGEM4'

connect = pyodbc.connect(
    f'Driver={driver};'
    f'Server={server};'
    f'Database={database};'
    f'uid={user_id};'
    f'pwd={user_password}'
)

TableNames = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',connect
)
