from request.customer import *

info_table = core.info().copy()
market_n_accounts = pd.read_csv(r'C:\Users\hiepdang\Downloads\SL_TKGD.csv')
market_n_accounts.columns = ['DATE','MARKET_DOMESTIC_COUNT','MARKET_FOREIGN_COUNT']
market_n_accounts['DATE'] = market_n_accounts['DATE'].map(lambda x:date(int(x[:4]),int(x[-2:]),15))
market_n_accounts['MARKET_DOMESTIC_COUNT'] = market_n_accounts['MARKET_DOMESTIC_COUNT'].str.replace(',','').map(int)
market_n_accounts['MARKET_FOREIGN_COUNT'] = market_n_accounts['MARKET_FOREIGN_COUNT'].str.replace(',','').map(int)
market_n_accounts.set_index('DATE',inplace=True)
market_n_accounts.sort_index(inplace=True)
market_n_accounts['MARKET_DOMESTIC_PCT_CHANGE'] = market_n_accounts['MARKET_DOMESTIC_COUNT'].pct_change(periods=1)
market_n_accounts['MARKET_FOREIGN_PCT_CHANGE'] = market_n_accounts['MARKET_FOREIGN_COUNT'].pct_change(periods=1)

phs_n_accounts = info_table[['OPENING_DATE']].loc[info_table['CUSTOMER_TYPE']=='Retail'].copy()
phs_n_accounts.reset_index(inplace=True)
phs_n_accounts.insert(0,'TYPE',
                      phs_n_accounts['TRADING_CODE'].map(lambda x:'Foreign' if x.startswith('022F') else 'Domestic'))
phs_n_accounts['OPENING_DATE'] = phs_n_accounts['OPENING_DATE'].map(lambda x:date(x.year,x.month,15))
phs_n_accounts = phs_n_accounts.groupby(['TYPE','OPENING_DATE'],as_index=False).count()
phs_domestic_n_accounts = phs_n_accounts.loc[phs_n_accounts['TYPE']=='Domestic'].drop('TYPE',axis=1).set_index(
    'OPENING_DATE').sort_index().rename({'TRADING_CODE':'PHS_DOMESTIC_COUNT'},axis=1)
phs_foreign_n_accounts = phs_n_accounts.loc[phs_n_accounts['TYPE']=='Foreign'].drop('TYPE',axis=1).set_index(
    'OPENING_DATE').sort_index().rename({'TRADING_CODE':'PHS_FOREIGN_COUNT'},axis=1)

phs_domestic_n_accounts['PHS_DOMESTIC_COUNT'] = phs_domestic_n_accounts['PHS_DOMESTIC_COUNT'].cumsum()
phs_foreign_n_accounts['PHS_FOREIGN_COUNT'] = phs_foreign_n_accounts['PHS_FOREIGN_COUNT'].cumsum()
phs_domestic_n_accounts['PHS_DOMESTIC_PCT_CHANGE'] = phs_domestic_n_accounts['PHS_DOMESTIC_COUNT'].pct_change(periods=1)
phs_foreign_n_accounts['PHS_FOREIGN_PCT_CHANGE'] = phs_foreign_n_accounts['PHS_FOREIGN_COUNT'].pct_change(periods=1)

result_n_accounts = pd.concat([market_n_accounts,phs_domestic_n_accounts,phs_foreign_n_accounts],axis=1)
result_n_accounts.sort_index(inplace=True)
result_n_accounts.dropna(inplace=True)
result_n_accounts.to_excel(r'C:\Users\hiepdang\Desktop\account_result.xlsx')

###############################################################################
nav_table = core.nav(subtype=['institutional']).copy()
nav_table.reset_index(inplace=True)
nav_table['BRANCH'] = nav_table['TRADING_CODE'].map(info_table['BRANCH'])

start = date(2010,1,1)
end = date.today()
bdates = pd.bdate_range(start,end,freq='BM')
bdates = [d.strftime('%Y-%m-%d') for d in bdates]
bdates = pd.Series([bdate(d,-4) for d in bdates])
bdates = bdates.map(lambda x:dt.datetime.strptime(x,'%Y-%m-%d'))

nav_table = nav_table.loc[nav_table['TRADING_DATE'].isin(bdates)]
nav_result = nav_table.groupby(['BRANCH','TRADING_DATE']).sum()
nav_result.reset_index(inplace=True)
nav_result.to_excel(r'C:\Users\hiepdang\Desktop\nav_result.xlsx')

###############################################################################

market_value = pd.read_csv(r'C:\Users\hiepdang\Downloads\trd_value_market.csv')
market_value.set_index(['TICKER','DATE'],inplace=True)
market_value.sort_index(inplace=True)
market_value = market_value.groupby('DATE').sum()
market_value.drop(['VOL','CLOSE'],axis=1,inplace=True)
market_value.index = pd.DatetimeIndex(market_value.index.map(lambda x:dt.date(int(x[:4]),int(x[5:7]),int(x[-2:]))))
market_value = market_value.groupby(pd.Grouper(freq='M')).sum()
market_value.columns = ['MARKET']
market_value *= 1e9

value_table = core.value().copy()
phs_value = value_table.groupby('TRADING_DATE').sum().groupby(pd.Grouper(freq='M')).sum()
phs_value.rename({'TRADING_VALUE':'PHS'},axis=1,inplace=True)

result_value = pd.concat([market_value,phs_value],axis=1).dropna()
result_value = result_value.astype(np.int64)
result_value.sort_index(inplace=True)
result_value.insert(1,'PCT_CHANGE',result_value['MARKET'].pct_change(periods=1))
result_value.insert(3,'PCT_CHANGE_',result_value['PHS'].pct_change(periods=1))
result_value.to_excel(r'C:\Users\hiepdang\Desktop\result_value.xlsx')
