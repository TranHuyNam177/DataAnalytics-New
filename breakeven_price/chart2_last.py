import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.animation as ani
import scipy as sc
from scipy import stats
from datetime import datetime,timedelta
import requests
import json
from matplotlib.dates import DateFormatter

fs = 'CLG'
days = 66
alpha = 0.001
simulation = 100000
seed = 10
np.random.seed(seed)


def reformat_large_tick_values(tick_val,pos):
    """
    Turns large tick values (in the billions, millions and thousands) such as 4500 into 4.5K
    and also appropriately turns 4000 into 4K (no zero after the decimal)
    """
    if tick_val>=1000:
        val = round(tick_val/1000,1)
        if tick_val%1000>0:
            new_tick_format = '{:,}K'.format(val)
        else:
            new_tick_format = '{:,}K'.format(int(val))
    else:
        new_tick_format = int(tick_val)
    new_tick_format = str(new_tick_format)
    return new_tick_format


# customize display for numpy and pandas
np.set_printoptions(linewidth=np.inf,precision=0,suppress=True,threshold=int(10e10))
pd.set_option("display.max_rows",None,"display.max_column",None,
              'display.width',None,'display.max_colwidth',20)

# manipulate date
pd.options.mode.chained_assignment = None
fromdate = (datetime.now()-timedelta(days=1000)).strftime("%Y-%m-%d")
todate = datetime.now().strftime("%Y-%m-%d")
# extract data from API
r = requests.post('https://api.phs.vn/market/utilities.svc/GetShareIntraday',
                  data=json.dumps({'symbol':fs,'fromdate':fromdate,'todate':todate}),
                  headers={'content-type':'application/json'})
df = pd.DataFrame(json.loads(r.json()['d']))

# cleaning data
df['trading_date'] = pd.to_datetime(df['trading_date'])
df['close_price'].loc[df['close_price']==0] = df['prior_price'].loc[df['close_price']==0]
df['prior_price'].loc[df['prior_price']==0] = df['close_price'].loc[df['prior_price']==0]
df['change_percent'] = df['close_price']/df['prior_price']-1
df['log_r'] = np.log(1+df['change_percent'])
for i in range(1,len(df.index)):
    df['change_log_r'] = df.loc[:,'log_r'].iloc[i]-df.loc[:,'log_r'].iloc[i-1]
df['change_log_r'].iloc[0] = 0
df['close_price'] = df['close_price']*1000
df['change'] = df['change']*1000
df.drop(columns=['change','prior_price'])

# Fundamental Statistics:
mean_logr = np.average(df['log_r'])
std_logr = np.std(df['log_r'])
kur_logr = sc.stats.kurtosis(df['log_r'])
skew_logr = sc.stats.skew(df['log_r'])

# D'Agostino-Pearson test for log return model
stat2_logr,p2_logr = sc.stats.normaltest(df['log_r'],nan_policy='omit')
print(f'p2_logr of {fs} is',p2_logr)

if p2_logr<=alpha:

    # Testing whether normal skewness
    stat3_logr,p3_logr = sc.stats.skewtest(df['log_r'],nan_policy='omit')
    print(f'p3_logr of {fs} is',p3_logr)

    # Testing whether normal kurtosis.
    stat4_logr,p4_logr = sc.stats.kurtosistest(df['log_r'],nan_policy='omit')
    print(f'p4_logr of {fs} is',p4_logr)

    if p3_logr<=alpha and p4_logr<=alpha:
        log_r = np.random.normal(mean_logr,std_logr,size=(simulation,days))
    if p3_logr<=alpha<p4_logr:
        log_r = sc.stats.t.rvs(df=df['log_r'].count()-1,loc=mean_logr,
                               scale=std_logr,size=(simulation,days))
    if p4_logr<=alpha<p3_logr:
        log_r = sc.stats.skewnorm.rvs(a=skew_logr,loc=mean_logr,
                                      scale=std_logr,size=(simulation,days))
    if p3_logr>alpha and p4_logr>alpha:
        log_r = sc.stats.nct.rvs(df=df['log_r'].count()-1,nc=skew_logr,loc=mean_logr,
                                 scale=std_logr,size=(simulation,days))

    # Convert log_r back to simulated price
    price_t = df['close_price'].loc[df['trading_date']==df['trading_date'].max()].iloc[0]
    simulated_price = np.zeros(shape=(simulation,days),dtype=np.int64)
    for i in range(simulation):
        simulated_price[i,0] = np.exp(log_r[i,0])*price_t
        for j in range(1,days):
            simulated_price[i,j] = np.exp(log_r[i,j])*simulated_price[i,j-1]
    df_historical = df[['trading_date','close_price']].iloc[int(max(-254,-df['trading_date']
                                                                    .count())):]

    # Post-processing and graphing
    pro_days = list()
    for j in range(days*2):
        if pd.to_datetime(df['trading_date'].max()+pd.Timedelta(days=j+1)).weekday()<5:
            pro_days.append(df['trading_date'].max()+pd.Timedelta(days=j+1))
    pro_days = pro_days[:days]

    simulation_no = [i for i in range(1,simulation+1)]
    df_simulated_price = pd.DataFrame(data=simulated_price,
                                      columns=pro_days,index=simulation_no).transpose()
    price_last = df_simulated_price.iloc[-1,:]

else:

    # Fundamental Statistics
    mean_change_logr = np.average(df['change_log_r'])
    std_change_logr = np.std(df['change_log_r'])
    kur_change_logr = sc.stats.kurtosis(df['change_log_r'])
    skew_change_logr = sc.stats.skew(df['change_log_r'])

    # D'Agostino-Pearson test for change in log return model
    stat2_change_logr,p2_change_logr = sc.stats.normaltest(df['change_log_r'],
                                                           nan_policy='omit')
    print(f'p2_change_logr of {fs} is',p2_change_logr)

    if p2_change_logr<=alpha:

        # Testing whether normal skewness
        stat3_change_logr,p3_change_logr = sc.stats.skewtest(df['change_log_r'],
                                                             nan_policy='omit')
        print(f'p3_change_logr of {fs} is',p3_change_logr)

        # Testing whether normal kurtosis
        stat4_change_logr,p4_change_logr = sc.stats.kurtosistest(df['change_log_r'],
                                                                 nan_policy='omit')
        print(f'p4_change_logr of {fs} is',p4_change_logr)

        if p3_change_logr<=alpha and p4_change_logr<=alpha:
            change_logr = np.random.normal(loc=mean_change_logr,
                                           cale=std_change_logr,size=(simulation,days))
        if p3_change_logr<=alpha<p4_change_logr:
            change_logr = sc.stats.t.rvs(df=df['change_log_r'].count()-1,
                                         loc=mean_change_logr,
                                         scale=std_change_logr,size=(simulation,days))
        if p4_change_logr<=alpha<p3_change_logr:
            change_logr = sc.stats.skewnorm.rvs(a=skew_change_logr,
                                                loc=mean_change_logr,
                                                scale=std_change_logr,size=(simulation,days))
        if p3_change_logr>alpha and p4_change_logr>alpha:
            change_logr = sc.stats.nct.rvs(df=df['change_log_r'].count()-1,
                                           nc=skew_change_logr,
                                           loc=mean_change_logr,
                                           scale=std_change_logr,size=(simulation,days))

        # Convert change_logr back to simulated price
        price_t = df['close_price'].loc[df['trading_date']==df['trading_date'].max()].iloc[0]
        price_t1 = df['close_price'].loc[df['trading_date']==df['trading_date'].max()
                                         -pd.Timedelta('1 day')].iloc[0]
        simulated_price = np.zeros(shape=(simulation,days),dtype=np.int64)
        for i in range(simulation):
            simulated_price[i,0] = np.exp(change_logr[i,0])*price_t**2/price_t1
            simulated_price[i,1] = np.exp(change_logr[i,1])* \
                                   simulated_price[i,0]**2/price_t
            for j in range(days):
                simulated_price[i,j] = np.exp(change_logr[i,j])* \
                                       simulated_price[i,j-1]**2/simulated_price[i,j-2]
        df_historical = df[['trading_date','close_price']].iloc[int(max(-254,
                                                                        df['trading_date']
                                                                        .count())):]
        # Post-processing and graphing
        pro_days = list()
        for j in range(days*2):
            if pd.to_datetime(df['trading_date'].max()+
                              pd.Timedelta(days=j+1)).weekday()<5:
                pro_days.append(df['trading_date'].max()+pd.Timedelta(days=j+1))
        pro_days = pro_days[:days]

        simulation_no = [i for i in range(1,simulation+1)]
        df_simulated_price = pd.DataFrame(data=simulated_price,
                                          columns=pro_days,index=simulation_no).transpose()
        price_last = df_simulated_price.iloc[-1,:]

    else:
        raise ValueError(f'{fs} cannot be simulated with given significance level')

alpha = 0.5
fig3,ax3 = plt.subplots(1,1,figsize=(7,5.2))

if i==0:
    ubound = pd.DataFrame(df_simulated_price.iloc[:,1])
    dbound = pd.DataFrame(df_simulated_price.iloc[:,1])
else:
    dbound = pd.DataFrame(df_simulated_price.quantile(q=0.00,axis=1,
                                                      interpolation='linear'))
    ubound = pd.DataFrame(df_simulated_price.quantile(q=0.70,axis=1,
                                                      interpolation='linear'))
ax3.set_ylim(np.round(np.min([df_historical['close_price'].min(),
                              np.quantile(df_simulated_price.values,0.00)])*0.95,-2),
             np.round(np.max([df_historical['close_price'].max(),
                              np.quantile(df_simulated_price.values,0.95)])*1.05,-2))

doublediff1 = np.diff(np.sign(np.diff(df_historical['close_price'].values)))
peak_locations = np.where(doublediff1==-2)[0]+1+df_historical.index[0]
doublediff2 = np.diff(np.sign(np.diff(-1*df_historical['close_price'].values)))
trough_locations = np.where(doublediff2==-2)[0]+1+df_historical.index[0]
ax3.scatter(df_historical['trading_date'][peak_locations],
            df_historical['close_price'][peak_locations],
            marker=matplotlib.markers.CARETUPBASE,color='tab:green',
            s=16,label='Peak')
ax3.scatter(df_historical['trading_date'][trough_locations],
            df_historical['close_price'][trough_locations],
            marker=matplotlib.markers.CARETDOWNBASE,color='tab:red',
            s=16,label='Trough')

ax3.plot(df_historical['trading_date'],df_historical['close_price'],
         color='midnightblue',alpha=0.9,lw=2,label='{}'.format(fs))
ax3.set_ylabel('Stock Price',fontsize=12)
ax3.xaxis.set_major_formatter(DateFormatter('%d/%m/%Y'))
ax3.yaxis.set_major_formatter(matplotlib.fs.FuncFormatter(reformat_large_tick_values))
plt.xticks(rotation=15)
plt.yticks(fontsize=11)
ax3.grid(axis='both',alpha=0.3)

breakeven_price = dbound.min().iloc[0]
connect_date = pd.date_range(df['trading_date'].max(),pro_days[0])[[0,-1]]
connect = pd.DataFrame({'price_u':[price_t,ubound.iloc[0,0]],
                        'price_d':[price_t,dbound.iloc[0,0]]},index=connect_date)
ax3.plot(connect,color='tab:red',lw=2,alpha=alpha)
ax3.plot(ubound,color='tab:red',lw=2,alpha=alpha)
ax3.plot(dbound,color='tab:red',lw=2,alpha=alpha,label='Best Case/ Worst Case')
ax3.legend(loc='upper right',fontsize=9)

# ax3.plot(df_simulated_price.iloc[:, i],
#         color='midnightblue', alpha=alpha, lw=0.5, linestyle='-')
ax3.fill_between(pro_days,ubound.iloc[:,0],dbound.iloc[:,0],color='green',
                 alpha=0.15)
ax3.axhline(breakeven_price,ls='--',lw=0.5,color='red')
# ax3.text(0.005, 1.06, "After " + "{}".format(i + 1) + " Simulations:",
#         fontsize=14,
#         color='tab:red', transform=ax3.transAxes, fontweight=700)
ax3.text(df['trading_date'].iloc[int(max(-200,-df['trading_date'].count()))],
         breakeven_price*1.15,"Breakeven Price: "+
         str(f"{round(breakeven_price):,d}"),fontsize=10,fontstyle='italic')
ax3.text(0.69,1.01,"Worst case: "+
         '{:,}'.format(round((breakeven_price/price_t-1)*100,2))+'%',
         fontsize=10,transform=ax3.transAxes,fontstyle='italic')
ax3.text(0.69,1.045,"Working days: "+
         '{}'.format(days),fontsize=10,transform=ax3.transAxes,fontstyle='italic')
ax3.text(0.69,1.08,"Breakeven Price: "+
         '{:,}'.format(int(breakeven_price)),fontsize=10,
         transform=ax3.transAxes,fontstyle='italic')
