import numpy as np
import pandas as pd
from request.stock import *

###############################################################################
###############################################################################
###############################################################################

model_2pct = pd.read_csv(
    r"\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Gia Hoa Von\25.11.2021_0.02.csv",
    usecols=['Ticker','0% at Risk'],
    index_col='Ticker',
)
model_2pct.columns = ['low_price_2pct']
model_5pct = pd.read_csv(
    r"\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Gia Hoa Von\25.11.2021_0.05.csv",
    usecols=['Ticker','0% at Risk'],
    index_col='Ticker',
)
model_5pct.columns = ['low_price_5pct']
model_adjustment_table = pd.concat([model_2pct,model_5pct],axis=1)
model_adjustment_table.dropna(how='all',inplace=True)
model_adjustment_table[
    [
        'market_price_2411',
    ]
] = np.nan
for ticker in model_adjustment_table.index:
    model_adjustment_table.loc[ticker,'market_price_2411'] = ta.hist(ticker,'2021-11-24').loc[0,'close']*1000
