from request.stock import *

destination_dir = join(dirname(realpath(__file__)),'result_chart')
backtest_dir = join(dirname(realpath(__file__)),'backtest_chart')

q_upper = 0.95
q_lower = 0

# CREDIT RATING
rating_path = join(realpath(dirname(dirname(__file__))),'credit_rating','result')
# adjust result_table_gen
gen_table = pd.read_csv(join(rating_path,'result_table_gen.csv'),index_col=['ticker'])
bank_table = pd.read_csv(join(rating_path,'result_table_bank.csv'),index_col=['ticker'])
ins_table = pd.read_csv(join(rating_path,'result_table_ins.csv'),index_col=['ticker'])
sec_table = pd.read_csv(join(rating_path,'result_table_sec.csv'),index_col=['ticker'])
rating_table = pd.concat([gen_table,bank_table,ins_table,sec_table])
rating_table.drop(['BM_'],inplace=True)
rating_table.fillna('Not enough data',inplace=True)


def run(
    ticker: str,
    hdays=252,
    pdays=66,
    simulation=100000,
    savefigure: bool = True,
    seed: int = 1,
    run_date: str = 'today',
):
    start_time = time.time()
    if run_date=='today':
        now = dt.datetime.now()
        model_date = now.strftime("%Y-%m-%d")
        if now.weekday() in holidays.WEEKEND or now in holidays.VN():
            model_date = bdate(model_date,-1)
    else:
        model_date = run_date
        year = int(run_date[:4])
        month = int(run_date[5:7])
        day = int(run_date[-2:])
        run_datetime = dt.datetime(year,month,day)
        if run_datetime.weekday() in holidays.WEEKEND or run_datetime in holidays.VN():
            model_date = bdate(run_date,-1)

    df = ta.hist(ticker,bdate(model_date,-hdays),model_date)
    df = df.loc[df['close']!=0]

    # cleaning data
    df['trading_date'] = pd.to_datetime(df['trading_date'])
    df.loc[df['close']==0,'close'] = df.loc[df['close']==0,'ref']
    df.loc[df['ref']==0,'ref'] = df.loc[df['ref']==0,'close']
    df['change_percent'] = df['close']/df['ref']-1
    df['logr'] = np.log(1+df['change_percent'])
    df['change_logr'] = 0
    for i in range(1,len(df.index)):
        df.loc[df.index[i],'change_logr'] = df.loc[df.index[i],'logr']-df.loc[df.index[i-1],'logr']
    df.loc[df.index[0],'change_logr'] = df.loc[df.index[1],'change_logr']
    df['close'] = df['close']*1000
    df.drop(columns=['ref','high','low','change','change_percent','total_volume','total_value'],inplace=True)

    # D'Agostino-Pearson test for log return model (fail the test -> abnormal price movement, manipulated price)
    stat_logr,p_logr = sc.stats.normaltest(df['logr'],nan_policy='omit')
    print(f'p_logr of {ticker} is',np.round(p_logr,3))

    if p_logr<=0.05:  # abnormal price movement

        # Testing whether normal skewness
        stat_skew,p_skew = sc.stats.skewtest(df['logr'],nan_policy='omit')
        print(f'p_skew of {ticker} is',np.round(p_skew,3))

        # Testing whether normal kurtosis.
        stat_kur,p_kur = sc.stats.kurtosistest(df['logr'],nan_policy='omit')
        print(f'p_kur of {ticker} is',np.round(p_kur,3))

        if p_skew<=0.05 and p_kur<=0.05:  # recheck on individual statistics
            loc = np.nanmean(df['logr'])
            scale = np.nanstd(df['logr'])
            np.random.seed(seed)
            logr = np.random.normal(loc,scale,size=(simulation,pdays))
        elif p_skew<=0.05<p_kur:  # abnormal movement case 1
            mean = np.nanmean(df['logr'])
            std = np.nanstd(df['logr'])
            deg_free = df['logr'].count()-1
            loc = mean
            scale = (std**2*(deg_free-2)/deg_free)**0.5
            logr = sc.stats.t.rvs(deg_free,loc,scale,size=(simulation,pdays),random_state=seed)
        elif p_skew>0.05>=p_kur:  # abnormal movement case 2
            mean = np.nanmean(df['logr'])
            std = np.nanstd(df['logr'])
            skew = sc.stats.skew(df['logr'],nan_policy='omit')
            theta = (np.pi/2*abs(skew)**(2/3)/(abs(skew)**(2/3)+((4-np.pi)/2)**(2/3)))**0.5*np.sign(skew)
            a = theta/(1-theta**2)**0.5
            scale = (std**2/(1-2*theta**2/np.pi))**0.5
            loc = mean-scale*theta*(2/np.pi)**0.5
            logr = sc.stats.skewnorm.rvs(a,loc,scale,size=(simulation,pdays),random_state=seed)
        else:  # abnormal movement case 3
            mean = np.nanmean(df['logr'])
            std = np.nanstd(df['logr'])
            deg_free = df['logr'].count()-1
            loc = mean
            scale = std
            nc = mean*(1-3/(4*deg_free-1))
            logr = sc.stats.nct.rvs(deg_free,nc,loc,scale,size=(simulation,pdays),random_state=seed)

    else:  # normal price movement
        loc = np.nanmean(df['logr'])
        scale = np.nanstd(df['logr'])
        np.random.seed(seed)
        logr = np.random.normal(loc,scale,size=(simulation,pdays))

    # Convert logr back to simulated price
    price_t = df.loc[df['trading_date']==df['trading_date'].max(),'close'].squeeze()
    simulated_price = np.zeros(shape=(simulation,pdays),dtype=np.int64)
    for i in range(simulation):
        simulated_price[i,0] = np.exp(logr[i,0])*price_t
        for j in range(1,pdays):
            simulated_price[i,j] = np.exp(logr[i,j])*simulated_price[i,j-1]
    df_historical = df[['trading_date','close']].iloc[int(max(-254,-df['trading_date'].count())):]

    # Post-processing and graphing
    pro_days = list()
    for j in range(pdays*2):
        if pd.to_datetime(df['trading_date'].max()+pd.Timedelta(days=j+1)).weekday()<5:
            pro_days.append(df['trading_date'].max()+pd.Timedelta(days=j+1))
    pro_days = pro_days[:pdays]

    simulation_no = [i for i in range(1,simulation+1)]
    df_simulated_price = pd.DataFrame(data=simulated_price,columns=pro_days,index=simulation_no).transpose()
    last_price = df_simulated_price.iloc[-1,:]
    ubound = pd.DataFrame(df_simulated_price.quantile(q=q_upper,axis=1,interpolation='linear'))
    dbound = pd.DataFrame(df_simulated_price.quantile(q=q_lower,axis=1,interpolation='linear'))
    lv0_price = dbound.min().iloc[0]
    lv1_price = last_price.loc[last_price.lt(price_t)].quantile(0.01)
    lv2_price = last_price.loc[last_price.lt(price_t)].quantile(0.03)
    lv3_price = last_price.loc[last_price.lt(price_t)].quantile(0.05)

    convert_to_valid_price = lambda x:int(adjprice(x).replace(',',''))
    lv0_price = convert_to_valid_price(lv0_price)
    lv1_price = convert_to_valid_price(lv1_price)
    lv2_price = convert_to_valid_price(lv2_price)
    lv3_price = convert_to_valid_price(lv3_price)

    try:
        rating = rating_table.loc[ticker,rating_table.columns[-1]]
    except KeyError:
        rating = 'Not in industry classification'

    if rating in ['Not enough data','Not in industry classification']:
        group = 'Not enough data'
        breakeven_price = lv0_price  # not enough fundamental data for credit rating means convertively making no markup
    elif rating>=75:
        group = 'A'
        breakeven_price = lv3_price
    elif rating>=50:
        group = 'B'
        breakeven_price = lv2_price
    elif rating>=25:
        group = 'C'
        breakeven_price = lv1_price
    else:
        group = 'D'
        breakeven_price = lv0_price

    connect_date = pd.date_range(df['trading_date'].max(),pro_days[0])[[0,-1]]
    connect = pd.DataFrame({'price_u':[price_t,ubound.iloc[0,0]],
                            'price_d':[price_t,dbound.iloc[0,0]]},
                           index=connect_date)

    fig,ax = plt.subplots(1,2,figsize=(12,5.5),sharey=True,gridspec_kw={'width_ratios':[2.5,1]})
    plt.subplots_adjust(left=0.05,right=0.95,bottom=0.05,top=0.95,wspace=0.01)

    ymin = np.round(np.min([df_historical['close'].min(),np.quantile(df_simulated_price.values,0.00)])*0.95,-2)
    ymax = np.round(np.max([df_historical['close'].max(),np.quantile(df_simulated_price.values,0.95)])*1.05,-2)
    ax[0].set_ylim(ymin,ymax)

    doublediff1 = np.diff(np.sign(np.diff(df_historical['close'].values)))
    peak_locations = np.where(doublediff1==-2)[0]+1+df_historical.index[0]
    doublediff2 = np.diff(np.sign(np.diff(-1*df_historical['close'].values)))
    trough_locations = np.where(doublediff2==-2)[0]+1+df_historical.index[0]
    ax[0].scatter(df_historical['trading_date'][peak_locations],
                  df_historical['close'][peak_locations],
                  marker='^',color='tab:green',
                  s=10,label='Peak')
    ax[0].scatter(df_historical['trading_date'][trough_locations],
                  df_historical['close'][trough_locations],
                  marker='v',color='tab:red',
                  s=10,label='Trough')
    ax[0].plot(df_historical['trading_date'],df_historical['close'],color='midnightblue',
               alpha=0.9,lw=1.5,label=f'{ticker}')
    ax[0].set_ylabel('Stock Price',fontsize=12)
    ax[0].xaxis.set_major_formatter(DateFormatter('%m/%Y'))
    ax[0].yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(priceKformat))
    plt.xticks(rotation=15)
    plt.yticks(fontsize=11)
    ax[0].grid(axis='both',alpha=0.3)
    ax[0].plot(connect,color='tab:red',lw=2,alpha=0.5)
    ax[0].plot(ubound,color='tab:red',lw=2,alpha=0.5)
    ax[0].plot(dbound,color='tab:red',lw=2,alpha=0.5,label='Best Case / Worst Case')
    ax[0].legend(loc='upper left',bbox_to_anchor=(-0.008,1.07),fontsize=9,ncol=4,columnspacing=1.2)

    ax[0].fill_between(pro_days,ubound.iloc[:,0],dbound.iloc[:,0],color='green',alpha=0.15)
    ax[0].axhline(breakeven_price,ls='--',lw=0.6,color='red')
    ax[0].annotate('Breakeven Price: '+adjprice(breakeven_price),
                   xy=(0.65,breakeven_price),xycoords=transforms
                   .blended_transform_factory(ax[0].transAxes,ax[0].transData),
                   xytext=(0,2),textcoords='offset pixels',
                   ha='left',va='bottom',fontsize=8,color='red')

    if run_date!='today':
        today_price = ta.last(ticker)*1000
        ax[0].axhline(today_price,ls='--',lw=0.6,color='red')
        ax[0].annotate(f'Last Close Price: '+adjprice(today_price),
                       xy=(0.65,today_price),xycoords=transforms
                       .blended_transform_factory(ax[0].transAxes,ax[0].transData),
                       xytext=(0,2),textcoords='offset pixels',
                       ha='left',va='bottom',fontsize=8,color='red')

    ax[0].yaxis.set_major_formatter(FuncFormatter(priceKformat))
    ax[0].annotate('Discount Rate: '
                   +f'{round((breakeven_price/price_t-1)*100,2):,} % \n'
                   +f'Trading Days: {pdays} \n'
                   +'Breakeven Price: '+adjprice(breakeven_price),
                   xy=(0.615,1.01),xycoords=ax[0].transAxes,
                   ha='left',va='bottom',fontsize=9,linespacing=1.3)

    sns.histplot(y=last_price,ax=ax[1],bins=200,legend=False,color='tab:blue',stat='density')
    ax[1].set_xlabel('Density')
    ax[1].set_ylabel('Stock Price')
    ax[1].axhline(lv0_price,ls='--',linewidth=0.5,color='tab:red')
    ax[1].axhline(lv1_price,ls='--',linewidth=0.5,color='tab:red')
    ax[1].axhline(lv2_price,ls='--',linewidth=0.5,color='tab:red')
    ax[1].axhline(lv3_price,ls='--',linewidth=0.5,color='tab:red')

    ax[1].annotate(f'Lowest Price: '+adjprice(lv0_price),
                   xy=(0.5,lv0_price),
                   xycoords=transforms.blended_transform_factory
                   (ax[1].transAxes,ax[1].transData),
                   xytext=(0,2),textcoords='offset pixels',
                   ha='left',va='bottom',fontsize=8,color='tab:red')
    ax[1].annotate('1% Worst Case: '+adjprice(lv1_price),
                   xy=(0.5,lv1_price),
                   xycoords=transforms.blended_transform_factory
                   (ax[1].transAxes,ax[1].transData),
                   xytext=(0,2),textcoords='offset pixels',
                   ha='left',va='bottom',fontsize=8,color='tab:red')
    ax[1].annotate('3% Worst Case: '+adjprice(lv2_price),
                   xy=(0.5,lv2_price),
                   xycoords=transforms.blended_transform_factory
                   (ax[1].transAxes,ax[1].transData),
                   xytext=(0,2),textcoords='offset pixels',
                   ha='left',va='bottom',fontsize=8,color='tab:red')
    ax[1].annotate('5% Worst Case: '+adjprice(lv3_price),
                   xy=(0.5,lv3_price),
                   xycoords=transforms.blended_transform_factory
                   (ax[1].transAxes,ax[1].transData),
                   xytext=(0,2),textcoords='offset pixels',
                   ha='left',va='bottom',fontsize=8,color='tab:red')

    f = lambda x:int(adjprice(x).replace(',',''))
    ax[1].axhspan(ymin=f(lv0_price),ymax=f(lv1_price),xmin=0,xmax=1,facecolor='tab:red',alpha=0.2)
    ax[1].axhspan(ymin=f(lv1_price),ymax=f(lv2_price),xmin=0,xmax=1,facecolor='tab:red',alpha=0.1)
    ax[1].axhspan(ymin=f(lv2_price),ymax=f(lv3_price),xmin=0,xmax=1,facecolor='tab:red',alpha=0.05)

    ax[1].annotate('The distribution plot describes probabilities of possible prices \n '
                   'in the next 66 trading days\n'
                   f'- For Group A, we propose to take higher risk, at {adjprice(lv3_price)} dong\n'
                   f'- For Group B, we propose to take moderate risk, at {adjprice(lv2_price)} dong\n'
                   f'- For Group C, we propose to take minor risk, at {adjprice(lv1_price)} dong\n'
                   f'- For Group D, we propose to take no risk, at {adjprice(lv0_price)} dong',
                   xy=(-0.3,1.01),xycoords=ax[1].transAxes,ha='left',va='bottom',fontsize=9,linespacing=1.3)

    ax[1].tick_params(axis='x',bottom=False,labelbottom=False)
    ax[1].yaxis.set_major_formatter(FuncFormatter(priceKformat))

    ax[1].grid(axis='both',alpha=0.1)

    if savefigure is True and run_date=='today':
        plt.savefig(join(destination_dir,f'{ticker}.png'),bbox_inches='tight')
    if savefigure is True and run_date!='today':
        plt.savefig(join(backtest_dir,f'{ticker}_{run_date}.png'),bbox_inches='tight')

    print(f'Breakeven price of {ticker} is '+str(breakeven_price))
    print("The execution time is: %s seconds"%(int(time.time()-start_time)))

    return lv0_price,lv1_price,lv2_price,lv3_price,breakeven_price,group
