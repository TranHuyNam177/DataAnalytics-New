from request.stock import *

destination_dir = join(dirname(realpath(__file__)),'result')

# Pre-classification by industry
standard = 'bics'
level = 'bics_l1'

# Parameters:
centroids = 3
min_tickers = 3
nstd_bound = 2

start_time = time.time()
pd.options.mode.chained_assignment = None

agg_data = fa.all('gen')  #################### expensive
periods = fa.periods[-13:]  # trailing 3 years
converted_idx = (agg_data.index.get_level_values(0).map(str)+'q'+agg_data.index.get_level_values(1).map(str))
agg_data = agg_data.loc[converted_idx.isin(periods)]

quantities = [
    'revenue',
    'cogs',
    'gross_profit',
    'interest',
    'pbt',
    'net_income',
    'cur_asset',
    'cash',
    'ar',
    'lt_receivable',
    'inv',
    'ppe',
    'asset',
    'liability',
    'cur_liability',
    'lt_debt',
    'equity',
    'retn_earnings',
    'period_retn_earnings',
    'lt_invst',
    'depre_tgbl',
    'depre_gw',
    's_exp',
    'ga_exp',
    'oprtng_income',
    'cfo',
    'st_debt'
]

tickers = fa.tickers('gen')

years = list()
quarters = list()
for period in periods:
    years.append(int(period[:4]))
    quarters.append(int(period[-1]))

period_tuple = list(zip(years,quarters))
inds = [x for x in itertools.product(period_tuple,tickers)]

for i in range(len(inds)):
    inds[i] = inds[i][0]+tuple([inds[i][1]])

index = pd.MultiIndex.from_tuples(inds,names=['year','quarter','ticker'])
col = pd.Index(quantities,name='quantity')

df = pd.DataFrame(columns=col,index=index)
df['revenue'] = agg_data[('is','3.')]
df['cogs'] = -agg_data['is','4.']
df['gross_profit'] = agg_data[('is','5.')]
df['interest'] = -agg_data[('is','8.')]
df['pbt'] = agg_data[('is','17.')]
df['net_income'] = agg_data[('is','21.')]
df['cur_asset'] = agg_data[('bs','1.')]
df['cash'] = agg_data[('bs','2.')]
df['ar'] = agg_data[('bs','9.')]
df['lt_receivable'] = agg_data[('bs','28.')]
df['inv'] = agg_data[('bs','18.')]
df['ppe'] = agg_data[('bs','36.')]
df['asset'] = agg_data[('bs','66.')]
df['liability'] = agg_data[('bs','67.')]
df['cur_liability'] = agg_data[('bs','68.')]
df['lt_debt'] = agg_data[('bs','91.')]
df['equity'] = agg_data[('bs','98.')]
df['retn_earnings'] = agg_data[('bs','113.')]
df['period_retn_earnings'] = agg_data[('bs','115.')]
df['lt_invst'] = agg_data[('bs','53.')]
df['depre_tgbl'] = agg_data[('cfi','3.')]  # tai sao mot so thang <0?
df['depre_gw'] = agg_data[('cfi','4.')]
df['s_exp'] = -agg_data[('is','10.')]
df['ga_exp'] = -agg_data[('is','11.')]
df['oprtng_income'] = agg_data[('is','12.')]
df['cfo'] = agg_data[('cfi','1.')]
df['st_debt'] = agg_data[('bs','78.')]

del agg_data  # for memory savings

# replace 0 values with 1000 VND to avoid 0 denominator
df = df.loc[~(df==0).all(axis=1)]  # remove no-data companies first
df = df.replace(to_replace=0,value=1e3)
# remove negative revenue companies
df = df.loc[df['revenue']>0]

df['cash&ppe'] = df['cash']+df['ppe']
df['equity_'] = df['equity']
df['revenue_'] = df['revenue']
df['cur_ratio'] = df['cur_asset']/df['cur_liability']
df['quick_ratio'] = (df['cur_asset']-df['inv'])/df['cur_liability']
df['cash_ratio'] = df['cash']/df['cur_liability']
df['(-)inv/cur_asset'] = -df['inv']/df['cur_asset']
df['wc_turnover'] = df['revenue']/(df['cur_asset']-df['cur_liability'])
df['inv_turnover'] = df['cogs']/df['inv']
df['ar_turnover'] = df['revenue']/df['ar']
df['ppe_turnover'] = df['revenue']/df['ppe']
df['(-)lib/asset'] = -df['liability']/df['asset']
df['(-)lt_debt/equity'] = -df['lt_debt']/df['equity']
df['gross_margin'] = df['gross_profit']/df['revenue']
df['net_margin'] = df['net_income']/df['revenue']
df['roe'] = df['net_income']/df['equity']
df['roa'] = df['net_income']/df['asset']
df['ebit/int'] = (df['pbt']+df['interest'])/df['interest']
df['wc/asset'] = (df['cur_asset']-df['cur_liability'])/df['asset']
df['retn_earning/asset'] = df['retn_earnings']/df['asset']
df['ebit/asset'] = (df['pbt']+df['interest'])/df['asset']
df['rev/asset'] = df['revenue']/df['asset']
df['dsri'] = -df['ar']/df['revenue']
df['aqi'] = (df['cur_asset']+df['ppe']+df['lt_invst'])/df['asset']
df['depi'] = (df['depre_tgbl']+df['depre_gw']) \
             /(df['ppe']+df['depre_tgbl']+df['depre_gw'])
df['sgai'] = -(df['s_exp']+df['ga_exp'])/df['revenue']
df['lvgi'] = -(df['cur_liability']+df['lt_debt']+df['st_debt'])/df['asset']
df['tata'] = -(df['oprtng_income']-df['cfo'])/df['asset']
df['cfo/debt'] = df['cfo']/(df['lt_debt']+df['st_debt'])
df['ebitda/rev'] = (df['pbt']+df['interest']+df['depre_tgbl']+df['depre_gw']) \
                   /df['revenue']
df['ebit/rev'] = (df['pbt']+df['interest'])/df['revenue']
df['pbt/rev'] = df['pbt']/df['revenue']
df['roce'] = (df['pbt']+df['interest'])/(df['cur_asset']-df['cur_liability'])
df['roic'] = df['period_retn_earnings'] \
             /(df['lt_debt']+df['st_debt']+df['equity'])
df['tr_turnover'] = df['revenue']/(df['ar']+df['lt_receivable'])
df['asset_turnover'] = df['revenue']/df['asset']
df['equity_turnover'] = df['revenue']/df['equity']
df['(-)lt_debt/asset'] = -df['lt_debt']/df['asset']
df['(-)debt/equity'] = -(df['st_debt']+df['lt_debt'])/df['equity']
df['(-)debt/asset'] = -(df['st_debt']+df['lt_debt'])/df['asset']
df['(-)cur_lib/equity'] = -df['cur_liability']/df['equity']
df['(-)cur_lib/asset'] = -df['cur_liability']/df['asset']
df['(-)liability/equity'] = -df['liability']/df['equity']
df['(-)liability/asset'] = -df['liability']/df['asset']

df = df.drop(columns=quantities)
df.sort_index(axis=1,inplace=True)
df.replace([np.inf,-np.inf],np.nan,inplace=True)

quantities_new = df.columns.to_list()

df.dropna(inplace=True,how='all')

ind_names = fa.industries(standard,int(level[-1]))
kmeans_index = pd.Index(ind_names,name='industry')

benmrks = pd.DataFrame(index=kmeans_index,columns=periods)
kmeans = pd.DataFrame(index=kmeans_index,columns=periods)
labels = pd.DataFrame(index=kmeans_index,columns=periods)
centers = pd.DataFrame(index=kmeans_index,columns=periods)
kmeans_tickers = pd.DataFrame(index=kmeans_index,columns=periods)
kmeans_coord = pd.DataFrame(index=kmeans_index,columns=periods)

ticker_list = dict()
industry_classifation = fa.classification('bics')[level].squeeze()
industry_list = list(industry_classifation.unique())

for industry in industry_list:
    ticker_list[industry] = industry_classifation.loc[industry_classifation==industry].index.to_list()
    for year,quarter in zip(years,quarters):
        # cross section
        tickers = ticker_list[industry]
        try:
            df_xs = df.loc[(year,quarter,tickers)]
        except KeyError:
            continue
        df_xs.dropna(axis=0,how='any',inplace=True)
        if df_xs.shape[0]<min_tickers:
            kmeans.loc[industry] = None
            labels.loc[industry] = None
            centers.loc[industry] = None
        else:
            for quantity in quantities_new:
                # remove outliers (Interquartile Range Method)
                ## (have to ensure symmetry)
                df_xs_median = df_xs[quantity].median()
                df_xs_75q = df_xs[quantity].quantile(q=0.75)
                df_xs_25q = df_xs[quantity].quantile(q=0.25)
                # deal with rare cases in which most tickers = 0, few != 0
                if df_xs_75q==df_xs_25q==0:
                    pass
                else:
                    cut_off = (df_xs_75q-df_xs_25q)*1.5
                    df_xs[quantity].clip(df_xs_25q-cut_off,df_xs_75q+cut_off,inplace=True)

                # standardize to mean=0
                df_xs_mean = df_xs[quantity].mean()
                df_xs[quantity] = df_xs[quantity]-df_xs_mean

                # standardize to range (-1,1)
                df_xs_min = df_xs[quantity].min()
                df_xs_max = df_xs[quantity].max()
                if df_xs_max==df_xs_min:
                    df_xs.drop(columns=quantity,inplace=True)
                else:
                    df_xs[quantity] = -1+(df_xs[quantity]-df_xs_min)/(df_xs_max-df_xs_min)*2

            # PCA algorithm
            X = df_xs.values
            benmrk_vector = np.array([-1 for n in range(df_xs.shape[1])])
            benmrk_vector = benmrk_vector.reshape((1,benmrk_vector.shape[0]))
            X = np.append(benmrk_vector,X,axis=0)
            PCs = PCA(n_components=0.9).fit_transform(X)

            idx = df_xs.index.to_list()
            idx.insert(0,(year,quarter,'BM_'))
            idx = pd.MultiIndex.from_tuples(idx)
            cols = ['PC_'+str(i+1) for i in range(PCs.shape[1])]
            df_xs = pd.DataFrame(PCs,index=idx,columns=cols)

            benmrks.loc[industry,f'{year}q{quarter}'] = PCs[0]

            # Kmeans algorithm
            kmeans.loc[industry,f'{year}q{quarter}'] = KMeans(
                n_clusters=centroids,
                init='k-means++',
                n_init=10,
                max_iter=1000,
                tol=1e-6,
                random_state=1
            ).fit(df_xs.dropna(axis=0,how='any'))

            kmeans_tickers.loc[industry,f'{year}q{quarter}'] = df_xs.index.get_level_values(2).to_list()
            kmeans_coord.loc[industry,f'{year}q{quarter}'] = df_xs.values
            labels.loc[industry,f'{year}q{quarter}'] = kmeans.loc[industry,f'{year}q{quarter}'].labels_.tolist()

            centers.loc[industry,f'{year}q{quarter}'] = kmeans.loc[
                industry,f'{year}q{quarter}'].cluster_centers_.tolist()
            print(f'Passed:: {industry} -- Year: {year}, Quarter: {quarter}')

del df_xs  # for memory saving

radius_centers = pd.DataFrame(index=kmeans_index,columns=periods)
for row in range(centers.shape[0]):
    for col in range(centers.shape[1]):
        if centers.iloc[row,col] is None:
            radius_centers.iloc[row,col] = None
        else:
            distance = np.zeros(centroids)
            for center in range(centroids):
                # origin at (-1,-1,-1,...) whose dimension varies by PCA
                distance[center] = ((np.array(centers.iloc[row,col][center])-np.array(
                    benmrks.iloc[row,col]))**2).sum()**(1/2)
            radius_centers.iloc[row,col] = distance

center_scores = pd.DataFrame(index=kmeans_index,columns=periods)
for row in range(centers.shape[0]):
    for col in range(centers.shape[1]):
        if radius_centers.iloc[row,col] is None:
            center_scores.iloc[row,col] = None
        else:
            center_scores.iloc[row,col] = rankdata(radius_centers.iloc[row,col])
            for n in range(1,centroids+1):
                center_scores.iloc[row,col] = \
                    np.where(center_scores.iloc[row,col]==n,100/(centroids+1)*n,center_scores.iloc[row,col])

radius_tickers = pd.DataFrame(index=kmeans_index,columns=periods)
for row in range(labels.shape[0]):
    for col in range(labels.shape[1]):
        if labels.iloc[row,col] is None:
            radius_tickers.iloc[row,col] = None
        else:
            distance = np.zeros(len(labels.iloc[row,col]))
            for ticker in range(len(labels.iloc[row,col])):
                distance[ticker] \
                    = (((np.array(kmeans_coord.iloc[row,col][ticker]))-np.array(benmrks.iloc[row,col]))**2).sum()**(1/2)
            radius_tickers.iloc[row,col] = distance

ticker_scores = pd.DataFrame(index=kmeans_index,columns=periods)
for row in range(radius_tickers.shape[0]):
    for col in range(radius_tickers.shape[1]):
        if radius_tickers.iloc[row,col] is None:
            ticker_scores.iloc[row,col] = None
        else:
            min_ = min(radius_centers.iloc[row,col])
            max_ = max(radius_centers.iloc[row,col])
            range_ = max_-min_
            f = interp1d(np.sort(np.append(
                radius_centers.iloc[row,col],
                [min_-range_/(centroids-1),max_+range_/(centroids-1)]
            )),np.sort(np.append(center_scores.iloc[row,col],[0,100])),
                kind='linear',bounds_error=False,fill_value=(0,100))
            ticker_scores.iloc[row,col] = f(radius_tickers.iloc[row,col])
            for n in range(len(ticker_scores.iloc[row,col])):
                ticker_scores.iloc[row,col][n] = int(ticker_scores.iloc[row,col][n])

ind_names = list()
ind_tickers = list()
ind_periods = list()
for industry in fa.industries('bics',int(level[-1])):
    for period in periods:
        try:
            if isinstance(kmeans_tickers.loc[industry,period],str):
                ind_names.append(industry)
                ind_tickers.append(kmeans_tickers.loc[industry,period])
                ind_periods.append(period)
            else:
                for ticker in kmeans_tickers.loc[industry,period]:
                    ind_names.append(industry)
                    ind_tickers.append(ticker)
                    ind_periods.append(period)
        except TypeError:
            continue

result_index = pd.MultiIndex.from_arrays(
    [ind_names,ind_tickers,ind_periods],names=['industry','ticker','period']
)

result_table = pd.DataFrame(index=result_index,columns=['credit_score'])
for industry in fa.industries('bics',int(level[-1])):
    for period in periods:
        try:
            for n in range(len(kmeans_tickers.loc[industry,period])):
                result_table.loc[(
                    industry,kmeans_tickers.loc[industry,period][n],period)] \
                    = ticker_scores.loc[industry,period][n]
        except TypeError:
            continue

result_table = result_table.unstack(level=2)
result_table.columns = result_table.columns.droplevel(0)
result_table.index = result_table.index.droplevel(0)
result_table.drop('BM_',axis=0,inplace=True)

component_filename = 'component_talbe_gen'
df.to_csv(join(destination_dir,component_filename+'.csv'))

result_filename = 'result_table_gen'
result_table.to_csv(join(destination_dir,result_filename+'.csv'))


def graph_tickers(tickers: list):
    table = pd.DataFrame(
        index=['credit_score'],
        columns=periods
    )
    for ticker in tickers:
        available_tickers = result_table.index
        if ticker not in available_tickers:
            continue
        for period in periods:
            table.loc['credit_score',period] = result_table.loc[ticker,period]
        fig,ax = plt.subplots(nrows=1,ncols=1,figsize=(8,6))
        ax.set_title(ticker,
                     fontsize=15,fontweight='bold',color='darkslategrey',
                     fontfamily='Times New Roman')

        xloc = np.arange(table.shape[1])  # label locations
        rects = ax.bar(xloc,table.iloc[0],width=0.8,
                       color='tab:blue',
                       label='Credit Score',
                       edgecolor='black')
        for rect in rects:
            height = rect.get_height()
            ax.annotate('{:.0f}'.format(height),
                        xy=(rect.get_x()+rect.get_width()/2,height),
                        xytext=(0,3),  # 3 points vertical offset
                        textcoords="offset points",
                        ha='center',va='bottom',fontsize=11)

        ax.set_xticks(np.arange(len(xloc)))
        ax.set_xticklabels(table.columns.tolist(),rotation=45,x=xloc,
                           fontfamily='Times New Roman',fontsize=11)

        ax.set_yticks(np.array([0,25,50,75,100]))
        ax.tick_params(axis='y',labelcolor='black',labelsize=11)

        Acolor = 'green'
        Bcolor = 'olivedrab'
        Ccolor = 'darkorange'
        Dcolor = 'firebrick'

        ax.axhline(100,ls='--',linewidth=0.5,color=Acolor)
        ax.axhline(75,ls='--',linewidth=0.5,color=Bcolor)
        ax.axhline(50,ls='--',linewidth=0.5,color=Ccolor)
        ax.axhline(25,ls='--',linewidth=0.5,color=Dcolor)
        ax.fill_between([-0.4,xloc[-1]+0.4],100,75,
                        color=Acolor,alpha=0.2)
        ax.fill_between([-0.4,xloc[-1]+0.4],75,50,
                        color=Bcolor,alpha=0.25)
        ax.fill_between([-0.4,xloc[-1]+0.4],50,25,
                        color=Ccolor,alpha=0.2)
        ax.fill_between([-0.4,xloc[-1]+0.4],25,0,
                        color=Dcolor,alpha=0.2)

        plt.xlim(-0.6,xloc[-1]+0.6)

        ax.set_ylim(top=110)
        midpoints = np.array([87.5,62.5,37.5,12.5])/110
        labels = ['Group A','Group B','Group C','Group D']
        colors = [Acolor,Bcolor,Ccolor,Dcolor]
        for loc in zip(midpoints,labels,colors):
            ax.annotate(loc[1],
                        xy=(-0.1,loc[0]),
                        xycoords='axes fraction',
                        textcoords="offset points",
                        xytext=(0,-5),
                        ha='center',va='bottom',
                        color=loc[2],fontweight='bold',
                        fontsize='large')

        ax.legend(loc='best',framealpha=0.5)
        ax.margins(tight=True)
        plt.subplots_adjust(left=0.15,bottom=0.1,right=0.95,top=0.9)
        plt.savefig(join(destination_dir,'Result',f'{ticker}_result.png'))


def compare_industry(tickers: list):
    components = {
        'cash&ppe':'Cash Plus PPE',
        'equity_':'Equity',
        'revenue_':'Revenue',
        'cur_ratio':'Current Ratio',
        'quick_ratio':'Quick Ratio',
        'cash_ratio':'Cash Ratio',
        '(-)inv/cur_asset':'Inventory on Current Asset',
        'wc_turnover':'Working Capital Turnover',
        'inv_turnover':'Inventory Turnover',
        'ar_turnover':'Account Receivable Turnover',
        'ppe_turnover':'PPE Turnover',
        '(-)lib/asset':'Debt Ratio',
        '(-)lt_debt/equity':'Long Term Debt to Equity',
        'gross_margin':'Gross Margin',
        'net_margin':'Net Profit Margin',
        'roe':'Return on Equity',
        'roa':'Return on Assets',
        'ebit/int':'Interest Coverage Ratio',
        'wc/asset':'Working Capital Over Total Assets Ratio',
        'retn_earning/asset':'Retained Earnings Over Total Assets',
        'ebit/asset':'Basic Earning Power',
        'rev/asset':'Asset Turnover Ratio',
        'dsri':'Days Sales in Receivables Index',
        'aqi':'Asset Quality Index',
        'depi':'Depreciation Index',
        'sgai':'Sales General and Administrative Expenses Index',
        'lvgi':'Leverage Index',
        'tata':'Total Accruals to Total Assets',
        'cfo/debt':'Cash Flow to Debt Ratio',
        'ebitda/rev':'EBITDA to Sales Ratio',
        'ebit/rev':'EBIT to Sales Ratio',
        'pbt/rev':'Pretax Profit Margin',
        'roce':'Return on Caption Employed',
        'roic':'Return on Invested Capital',
        'tr_turnover':'Sales on Accounts Receivable',
        'asset_turnover':'Asset Turnover',
        'equity_turnover':'Equity Turnover',
        '(-)lt_debt/asset':'Long-Term Debt to Total Assets Ratio',
        '(-)debt/equity':'Debt to Equity Ratio',
        '(-)debt/asset':'Debt to Assets Ratio',
        '(-)cur_lib/equity':'Current Liabilities to Equity Ratio',
        '(-)cur_lib/asset':'Current Liabilities to Assets Ratio',
        '(-)liability/equity':'Total Liabilities to Equity Ratio',
        '(-)liability/asset':'Total Liabilities to Assets Ratio',
    }

    full_list = fa.classification('bics')[level]
    for ticker in tickers:
        if ticker not in result_table.index:
            continue
        # to avoid cases of missing data right at the first period, result in mis-shaped
        industry = full_list.loc[ticker]
        peers = full_list.loc[full_list==industry].index.tolist()

        table = df.loc[df.index.get_level_values(2).isin(peers)]
        table.dropna(axis=0,how='all',inplace=True)

        median = table.groupby(axis=0,level=[0,1]).median()
        quantities = pd.DataFrame(np.zeros_like(median),
                                  index=median.index,
                                  columns=median.columns)

        ref_table = table.xs(ticker,axis=0,level=2)
        quantities = pd.concat([quantities,ref_table],axis=0)
        quantities = quantities.groupby(level=[0,1],axis=0).sum()

        comparison = pd.concat([quantities,median],axis=1,join='outer',keys=[ticker,'median'])

        periods = [f'{q[0]}q{q[1]}' for q in comparison.index]
        variables = comparison.columns.get_level_values(1).drop_duplicates().to_numpy()

        chartsperrow = 4
        rowsrequired = int(np.ceil(len(variables)/chartsperrow))
        variables = np.append(variables,[None]*(chartsperrow*rowsrequired-len(variables)))
        variables = np.reshape(variables,(rowsrequired,chartsperrow))
        fig,ax = plt.subplots(rowsrequired,chartsperrow,
                              figsize=(chartsperrow*6,rowsrequired*4.5),
                              tight_layout=True)
        for row in range(rowsrequired):
            for col in range(chartsperrow):
                w = 0.35
                l = np.arange(len(periods))  # the label locations
                if variables[row,col] is None:
                    ax[row,col].axis('off')
                else:
                    ax[row,col].bar(
                        l-w/2,
                        quantities.iloc[:,row*chartsperrow+col],
                        width=w,
                        label=ticker,
                        color='tab:orange',
                        edgecolor='black'
                    )
                    ax[row,col].bar(
                        l+w/2,
                        median.iloc[:,row*chartsperrow+col],
                        width=w,
                        label='Industry\'s Average',
                        color='tab:blue',
                        edgecolor='black'
                    )
                    ax[row,col].axhline(y=0,color='black',alpha=0.5,lw=1)

                    plt.setp(ax[row,col].xaxis.get_majorticklabels(),rotation=45)
                    ax[row,col].set_xticks(l)
                    ax[row,col].set_xticklabels(periods,fontsize=13)
                    ax[row,col].set_yticks([])
                    ax[row,col].set_autoscaley_on(True)
                    ax[row,col].set_title(components[variables[row,col]],fontsize=18)

        fig.subplots_adjust(top=0.9)
        fig.suptitle(f'{ticker}\n'
                     f'Comparison With The Industry\'s Average \n'
                     f'All {len(quantities_new)} Variables In Use \n',
                     fontweight='bold',color='darkslategrey',
                     fontfamily='Times New Roman',fontsize=30,y=0.99)
        handles,labels = ax[0,0].get_legend_handles_labels()
        fig.legend(handles,labels,loc='upper left',
                   bbox_to_anchor=(0.01,0.98),ncol=2,fontsize=18,
                   markerscale=0.7)
        plt.savefig(join(destination_dir,'Compare with Industry',f'{ticker}_compare_industry.png'))
        plt.close('all')


graph_tickers(fa.tickers('gen'))
compare_industry(fa.tickers('gen'))

execution_time = time.time()-start_time
print(f"The execution time is: {int(execution_time)}s seconds")
