import sys
import os.path
import pandas as pd

file_path = os.getcwd()
ticker = sys.argv[1]

try:
    files = [file for file in os.listdir(os.path.join(file_path,'backtest_result')) if
             file.startswith(f'C{ticker}') and file.endswith('.xlsx')]
    lastest_period = max([int(file.split('.')[0][4:]) for file in files])
    file = f'C{ticker}{lastest_period}.xlsx'
    hedge_table = pd.read_excel(
        os.path.join(file_path,'backtest_result',file),
        index_col=0
    )
    json_string = hedge_table.to_json()
except (Exception,):
    json_string = '{"ERROR":"True}'

print(json_string)
