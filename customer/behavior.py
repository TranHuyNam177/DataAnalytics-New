from request.customer import *

info_table = core.info()

age = info_table['AGE']
age = age[~age.isna()]

gender = info_table['GENDER']
gender = gender[~(gender=='Unknown')]
male_count = gender[gender=='Male'].count()
female_count = gender[gender=='Female'].count()

opening_date = info_table['OPENING_DATE']


def sum_margin(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    # Note: average over time, sum over trading codes, sum over ages in each age group
    table = aggregation.on_margin_by('all','mean',fromdate,todate,subtype,locations=[location])
    table = table.loc[pd.IndexSlice[location,:,:],:].droplevel(0)
    table.reset_index(inplace=True)

    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(genders=[x['GENDER']],
                                                 ages=[x['AGE']],
                                                 locations=[location],
                                                 subtype=subtype)),
                      axis=1)

    table['AGE'] = table['AGE'].map(age_group_function)
    table = table.groupby(['GENDER','AGE'],as_index=False).sum()

    return table


def sum_value(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    # Note: sum over time, sum over trading codes, sum over ages in each age group
    table = aggregation.on_value_by('all','mean',fromdate,todate,subtype,locations=[location])
    table = table.loc[pd.IndexSlice[location,:,:],:].droplevel(0)
    table.reset_index(inplace=True)

    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(genders=[x['GENDER']],
                                                 ages=[x['AGE']],
                                                 locations=[location],
                                                 subtype=subtype)),
                      axis=1)

    table['AGE'] = table['AGE'].map(age_group_function)
    table = table.groupby(['GENDER','AGE'],as_index=False).sum()

    return table


def sum_fee(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    # Note: sum over time, sum over trading codes, sum over ages in each age group
    table = aggregation.on_fee_by('all','mean',fromdate,todate,subtype,locations=[location])
    table = table.loc[pd.IndexSlice[location,:,:],:].droplevel(0)
    table.reset_index(inplace=True)

    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(genders=[x['GENDER']],
                                                 ages=[x['AGE']],
                                                 locations=[location],
                                                 subtype=subtype)),
                      axis=1)

    table['AGE'] = table['AGE'].map(age_group_function)
    table = table.groupby(['GENDER','AGE'],as_index=False).sum()

    return table


def sum_interest(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    # Note: sum over time, sum over trading codes, sum over ages in each age group
    table = aggregation.on_interest_by('all','mean',fromdate,todate,subtype,locations=[location])
    table = table.loc[pd.IndexSlice[location,:,:],:].droplevel(0)
    table.reset_index(inplace=True)

    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(genders=[x['GENDER']],
                                                 ages=[x['AGE']],
                                                 locations=[location],
                                                 subtype=subtype)),
                      axis=1)

    table['AGE'] = table['AGE'].map(age_group_function)
    table = table.groupby(['GENDER','AGE'],as_index=False).sum()

    return table


def sum_nav(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    # Note: sum over time, sum over trading codes, sum over ages in each age group
    table = aggregation.on_interest_by('all','mean',fromdate,todate,subtype,locations=[location])
    table = table.loc[pd.IndexSlice[location,:,:],:].droplevel(0)
    table.reset_index(inplace=True)

    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(genders=[x['GENDER']],
                                                 ages=[x['AGE']],
                                                 locations=[location],
                                                 subtype=subtype)),
                      axis=1)

    table['AGE'] = table['AGE'].map(age_group_function)
    table = table.groupby(['GENDER','AGE'],as_index=False).sum()

    return table


################################## GRAPHING ###################################


def info():
    fig,ax = plt.subplots(2,2,figsize=(15,12))
    plt.subplots_adjust(left=0.1,right=0.95,
                        bottom=0.1,top=0.92,
                        wspace=0.2,hspace=0.18)

    # chart 1
    sns.histplot(x=age,ax=ax[0,0],discrete=True)
    ax[0,0].set_xlabel('Age',labelpad=8,fontsize=12)
    ax[0,0].set_ylabel('Number of Customers',fontsize=12)
    ax[0,0].tick_params(axis='both',labelsize=11)

    ax1 = ax[0,0].twinx()
    sns.kdeplot(x=age,ax=ax1)
    ax1.set_ylabel('')
    ax1.tick_params(axis='y',right=False,labelright=False)

    # Chart 2
    rects = ax[0,1].bar(x=['Male','Female'],
                        height=[male_count,
                                female_count],
                        width=0.7,edgecolor='black',
                        linewidth='1')

    male_percent = male_count/(male_count+female_count)
    female_percent = 1-male_percent
    label_list = [f'{value:.1%}' for value in [male_percent,female_percent]]

    ax[0,1].bar_label(container=rects,
                      labels=label_list,
                      padding=3,fontsize=12)
    ax[0,1].set_ylim([0,(male_count+female_count)*.8])
    ax[0,1].set_ylabel('')

    ax[0,1].tick_params(axis='y',left=False,labelleft=False)
    ax[0,1].tick_params(axis='x',labelsize=12)

    # Chart 3
    table_ = info_table[~(info_table['GENDER']=='Unknown')]
    sns.histplot(data=table_,x='AGE',hue='GENDER',hue_order=['Female','Male'],
                 ax=ax[1,0],bins=50,color='white',alpha=0,linewidth=0)
    ax[1,0].set_xlabel('Age',labelpad=8,fontsize=12)
    ax[1,0].set_ylabel('Number of Customers',fontsize=12)
    ax[1,0].tick_params(axis='both',labelsize=11)

    ax3 = ax[1,0].twinx()
    sns.kdeplot(data=table_,x='AGE',hue='GENDER',hue_order=['Female','Male'],
                ax=ax3,fill=True,palette='Set1',linewidth=.5)
    ax3.set_xlabel('')
    ax3.set_ylabel('')
    ax3.tick_params(axis='both',
                    right=False,labelright=False,
                    bottom=False,labelbottom=False)

    # Chart 4
    sns.histplot(data=opening_date.map(lambda x:x.year),ax=ax[1,1],
                 bins=5,color='tab:blue',discrete=True)
    ax[1,1].set_ylabel('Number of New Accounts',fontsize=12)
    ax[1,1].set_xlabel('Year',labelpad=8,fontsize=12)
    ax[1,1].tick_params(axis='both',labelsize=11)

    plt.suptitle('Descriptive Statistics of Customers Pool',
                 fontsize=20,fontweight='book')

    plt.savefig(join(dirname(__file__),f'info.png'))


###############################################################################

def age_group_function(age):
    if 16<=age<=25:
        group_name = '18-25'
    elif 26<=age<=30:
        group_name = '26-30'
    elif 31<=age<=35:
        group_name = '31-35'
    elif 36<=age<=40:
        group_name = '36-40'
    elif 41<=age<=45:
        group_name = '41-45'
    elif 46<=age<=50:
        group_name = '46-50'
    elif 51<=age<=55:
        group_name = '51-55'
    elif 56<=age<=60:
        group_name = '56-60'
    elif 61<=age<=65:
        group_name = '61-65'
    elif 66<=age<=70:
        group_name = '66-70'
    elif 71<=age<=75:
        group_name = '71-75'
    else:
        group_name = '75+'
    return group_name


def graph_total_margin_at_location_by_age_gender(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    """
    This graph displays total average margin outstandings over a period of time by subtypes in any one specific location
    """

    table = sum_margin(location,fromdate,todate,subtype)

    fig,ax = plt.subplots(1,1,figsize=(16,12))
    sns.barplot(x='AGE',y='MARGIN_OUTSTANDINGS',hue='GENDER',
                data=table,ax=ax,palette=['tab:red','tab:blue'])

    fig.subplots_adjust(top=0.87,bottom=0.08,left=0.08,right=0.92)
    ax.tick_params(axis='both',which='major',labelsize=13)
    ax.set_ylabel('TOTAL MARGIN OUTSTANDINGS',labelpad=10,fontsize=15)
    ax.set_xlabel('AGE',labelpad=10,fontsize=15)
    ax.legend(prop={'size':14})

    if todate is None:
        todate = core.time
        todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    else:
        todate_text = f'{todate[-2:]}/{todate[5:7]}/{todate[:4]}'

    if fromdate is None:
        fromdate_text = '01/01/2018'
    else:
        fromdate_text = f'{fromdate[-2:]}/{fromdate[5:7]}/{fromdate[:4]}'

    ax.grid(True,lw=0.6,alpha=0.5)
    ax.set_axisbelow(True)

    rects = ax.patches
    # Make labels.
    labels = table['COUNT'].astype('int64').tolist()

    for rect,label in zip(rects,labels):
        height = rect.get_height()
        ax.annotate(f'{label}\nppl',
                    (rect.get_x()+rect.get_width()/2,height),
                    xytext=(0,5),textcoords='offset pixels',ha='center',fontsize=12)

    ax.yaxis.offsetText.set_visible(False)
    ax.annotate('Billion VND',xy=(0.01,1.015),xycoords='axes fraction',fontsize=12)

    locs,labels = yticks()
    yticks(locs,map(lambda x:'%.0f'%x,locs/1e9))

    fig.suptitle(f"Which Customer Groups Consume Most Margin\n"
                 f"in {location}\n"
                 f"From {fromdate_text} to {todate_text}",
                 y=0.95,fontsize=20,fontfamily='Arial Unicode MS')

    plt.savefig(join(dirname(__file__),f'TotalMargin_{location}.png'))

    return table


def graph_total_trading_value_at_location_by_age_gender(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    """
    This graph displays total trading value over a period of time by subtypes in any one specific location
    """

    table = sum_value(location,fromdate,todate,subtype)

    fig,ax = plt.subplots(1,1,figsize=(16,12))
    sns.barplot(x='AGE',y='TRADING_VALUE',hue='GENDER',
                data=table,ax=ax,palette=['tab:red','tab:blue'])

    fig.subplots_adjust(top=0.87,bottom=0.08,left=0.08,right=0.92)
    ax.tick_params(axis='both',which='major',labelsize=13)
    ax.set_ylabel('TOTAL TRADING VALUE',labelpad=10,fontsize=15)
    ax.set_xlabel('AGE',labelpad=10,fontsize=15)
    ax.legend(prop={'size':14})

    if todate is None:
        todate = core.time
        todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    else:
        todate_text = f'{todate[-2:]}/{todate[5:7]}/{todate[:4]}'

    if fromdate is None:
        fromdate_text = '01/01/2018'
    else:
        fromdate_text = f'{fromdate[-2:]}/{fromdate[5:7]}/{fromdate[:4]}'

    ax.grid(True,lw=0.6,alpha=0.5)
    ax.set_axisbelow(True)

    rects = ax.patches
    # Make labels.
    labels = table['COUNT'].astype('int64').tolist()

    for rect,label in zip(rects,labels):
        height = rect.get_height()
        ax.annotate(f'{label}\nppl',
                    (rect.get_x()+rect.get_width()/2,height),
                    xytext=(0,5),textcoords='offset pixels',ha='center',fontsize=12)

    ax.yaxis.offsetText.set_visible(False)
    ax.annotate('Billion VND',xy=(0.01,1.015),xycoords='axes fraction',fontsize=12)

    locs,labels = yticks()
    yticks(locs,map(lambda x:'%.0f'%x,locs/1e9))

    fig.suptitle(f"Which Customer Groups Trade Most\n"
                 f"in {location}\n"
                 f"From {fromdate_text} to {todate_text}",
                 y=0.95,fontsize=20,fontfamily='Arial Unicode MS')

    plt.savefig(join(dirname(__file__),f'TotalTradingValue{location}.png'))

    return table


def graph_total_trading_fee_at_location_by_age_gender(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    """
    This graph displays total trading fee over a period of time by subtypes in any one specific location
    """

    # Note: sum over time, sum over trading codes, sum over ages in each age group
    table = sum_fee(location,fromdate,todate,subtype)

    fig,ax = plt.subplots(1,1,figsize=(16,12))
    sns.barplot(x='AGE',y='TRADING_FEES',hue='GENDER',
                data=table,ax=ax,palette=['tab:red','tab:blue'])

    fig.subplots_adjust(top=0.87,bottom=0.08,left=0.08,right=0.92)
    ax.tick_params(axis='both',which='major',labelsize=13)
    ax.set_ylabel('TOTAL TRADING FEE',labelpad=10,fontsize=15)
    ax.set_xlabel('AGE',labelpad=10,fontsize=15)
    ax.legend(prop={'size':14})

    if todate is None:
        todate = core.time
        todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    else:
        todate_text = f'{todate[-2:]}/{todate[5:7]}/{todate[:4]}'

    if fromdate is None:
        fromdate_text = '01/01/2018'
    else:
        fromdate_text = f'{fromdate[-2:]}/{fromdate[5:7]}/{fromdate[:4]}'

    ax.grid(True,lw=0.6,alpha=0.5)
    ax.set_axisbelow(True)

    rects = ax.patches
    # Make labels.
    labels = table['COUNT'].astype('int64').tolist()

    for rect,label in zip(rects,labels):
        height = rect.get_height()
        ax.annotate(f'{label}\nppl',
                    (rect.get_x()+rect.get_width()/2,height),
                    xytext=(0,5),textcoords='offset pixels',ha='center',fontsize=12)

    ax.yaxis.offsetText.set_visible(False)
    ax.annotate('Billion VND',xy=(0.01,1.015),xycoords='axes fraction',fontsize=12)

    locs,labels = yticks()
    yticks(locs,map(lambda x:'%.0f'%x,locs/1e9))

    fig.suptitle(f"Which Customer Groups Incur Most Trading Fee\n"
                 f"in {location}\n"
                 f"From {fromdate_text} to {todate_text}",
                 y=0.95,fontsize=20,fontfamily='Arial Unicode MS')

    plt.savefig(join(dirname(__file__),f'TotalTradingFee{location}.png'))

    return table


def graph_total_interest_expense_at_location_by_age_gender(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    # Note: sum over time, sum over trading codes, sum over ages in each age group
    table = sum_interest(location,fromdate,todate,subtype)

    fig,ax = plt.subplots(1,1,figsize=(16,12))
    sns.barplot(x='AGE',y='INTEREST_EXPENSE',hue='GENDER',
                data=table,ax=ax,palette=['tab:red','tab:blue'])

    fig.subplots_adjust(top=0.87,bottom=0.08,left=0.08,right=0.92)
    ax.tick_params(axis='both',which='major',labelsize=13)
    ax.set_ylabel('TOTAL INTEREST EXPENSE',labelpad=10,fontsize=15)
    ax.set_xlabel('AGE',labelpad=10,fontsize=15)
    ax.legend(prop={'size':14})

    if todate is None:
        todate = core.time
        todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    else:
        todate_text = f'{todate[-2:]}/{todate[5:7]}/{todate[:4]}'

    if fromdate is None:
        fromdate_text = '01/01/2018'
    else:
        fromdate_text = f'{fromdate[-2:]}/{fromdate[5:7]}/{fromdate[:4]}'

    ax.grid(True,lw=0.6,alpha=0.5)
    ax.set_axisbelow(True)

    rects = ax.patches
    # Make labels.
    labels = table['COUNT'].astype('int64').tolist()

    for rect,label in zip(rects,labels):
        height = rect.get_height()
        ax.annotate(f'{label}\nppl',
                    (rect.get_x()+rect.get_width()/2,height),
                    xytext=(0,5),textcoords='offset pixels',ha='center',fontsize=12)

    ax.yaxis.offsetText.set_visible(False)
    ax.annotate('Billion VND',xy=(0.01,1.015),xycoords='axes fraction',fontsize=12)

    locs,labels = yticks()
    yticks(locs,map(lambda x:'%.0f'%x,locs/1e9))

    fig.suptitle(f"Which Customer Groups Incur Most Interest Expense\n"
                 f"in {location}\n"
                 f"From {fromdate_text} to {todate_text}",
                 y=0.95,fontsize=20,fontfamily='Arial Unicode MS')

    plt.savefig(join(dirname(__file__),f'TotalInterestExpense_{location}.png'))

    return table


def graph_total_nav_at_location_by_age_gender(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
):
    # Note: average over time, sum over trading codes, sum over ages in each age group
    table = sum_nav(location,fromdate,todate,subtype)

    fig,ax = plt.subplots(1,1,figsize=(16,12))
    sns.barplot(x='AGE',y='NAV',hue='GENDER',
                data=table,ax=ax,palette=['tab:red','tab:blue'])

    fig.subplots_adjust(top=0.87,bottom=0.08,left=0.08,right=0.92)
    ax.tick_params(axis='both',which='major',labelsize=13)
    ax.set_ylabel('TOTAL NAV',labelpad=10,fontsize=15)
    ax.set_xlabel('AGE',labelpad=10,fontsize=15)
    ax.legend(prop={'size':14})

    if todate is None:
        todate = core.time
        todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    else:
        todate_text = f'{todate[-2:]}/{todate[5:7]}/{todate[:4]}'

    if fromdate is None:
        fromdate_text = '01/01/2018'
    else:
        fromdate_text = f'{fromdate[-2:]}/{fromdate[5:7]}/{fromdate[:4]}'

    ax.grid(True,lw=0.6,alpha=0.5)
    ax.set_axisbelow(True)

    rects = ax.patches
    # Make labels.
    labels = table['COUNT'].astype('int64').tolist()

    for rect,label in zip(rects,labels):
        height = rect.get_height()
        ax.annotate(f'{label}\nppl',
                    (rect.get_x()+rect.get_width()/2,height),
                    xytext=(0,5),textcoords='offset pixels',ha='center',fontsize=12)

    ax.yaxis.offsetText.set_visible(False)
    ax.annotate('Billion VND',xy=(0.01,1.015),xycoords='axes fraction',fontsize=12)

    locs,labels = yticks()
    yticks(locs,map(lambda x:'%.0f'%x,locs/1e9))

    fig.suptitle(f"Which Customer Groups Have Highest NAV\n"
                 f"in {location}\n"
                 f"From {fromdate_text} to {todate_text}",
                 y=0.95,fontsize=20,fontfamily='Arial Unicode MS')

    plt.savefig(join(dirname(__file__),f'TotalNAV_{location}.png'))

    return table


###############################################################################

def graph_turnover_by_location(
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
    least: int = 500
):
    if todate is None:
        todate = core.time
    else:
        todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

    if fromdate is None:
        fromdate = dt.date(2018,1,1)
    else:
        fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

    location_series = core.info('all',subtype)['LOCATION'].dropna()
    location_series = location_series.loc[~(location_series=='Unknown')]

    idx = pd.IndexSlice

    nav_table = core.nav('all','all',subtype)
    nav_table = nav_table.loc[idx[fromdate:todate,:],:]
    nav_table.insert(0,'LOCATION',nav_table.index.get_level_values(1).map(location_series))
    sum_nav_over_trading_codes = nav_table.groupby(['TRADING_DATE','LOCATION']).sum()
    result_avg_nav = sum_nav_over_trading_codes.groupby(['LOCATION']).agg('mean')
    result_avg_nav = result_avg_nav.astype('int64')

    margin_table = core.margin('all','all',subtype)
    margin_table = margin_table.loc[idx[fromdate:todate,:],:]
    margin_table.insert(0,'LOCATION',margin_table.index.get_level_values(1).map(location_series))
    sum_margin_over_trading_codes = margin_table.groupby(['TRADING_DATE','LOCATION']).sum()
    result_avg_margin = sum_margin_over_trading_codes.groupby(['LOCATION']).agg('mean')
    result_avg_margin = result_avg_margin.astype('int64')

    value_table = core.value('all','all',subtype)
    value_table = value_table.loc[idx[fromdate:todate,:],:]
    value_table.insert(0,'LOCATION',value_table.index.get_level_values(1).map(location_series))
    sum_value_over_trading_codes = value_table.groupby(['TRADING_DATE','LOCATION']).sum()
    result_sum_value = sum_value_over_trading_codes.groupby(['LOCATION']).agg('sum')
    result_sum_value = result_sum_value.astype('int64')

    table = pd.concat([result_avg_nav,result_avg_margin,result_sum_value],axis=1)
    table.reset_index(inplace=True)
    table['TRADING_TURNOVER'] = table['TRADING_VALUE'].div(table['NAV'].add(table['MARGIN_OUTSTANDINGS'],fill_value=0),
                                                           fill_value=0)

    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(locations=[x['LOCATION']],
                                                 subtype=subtype)),
                      axis=1)
    table = table.loc[table['COUNT'].ge(least)]
    table.sort_values(['TRADING_TURNOVER'],inplace=True,ascending=False)

    fig,ax = plt.subplots(1,1,figsize=(16,12))
    sns.barplot(x='LOCATION',y='TRADING_TURNOVER',data=table,ax=ax,palette=['tab:blue'])

    fig.subplots_adjust(top=0.92,bottom=0.11,left=0.08,right=0.92)
    ax.tick_params(axis='both',which='major',labelsize=13)
    ax.set_ylabel('TRADING TURNOVER',labelpad=10,fontsize=15)
    ax.set_xlabel('LOCATION',labelpad=10,fontsize=15)
    ax.legend(prop={'size':14})

    ax.grid(True,lw=0.6,alpha=0.5)
    ax.set_axisbelow(True)

    rects = ax.patches
    # Make labels.
    labels = table['COUNT'].astype('int64').tolist()

    for rect,label in zip(rects,labels):
        height = rect.get_height()
        ax.annotate(f'{label}\nppl',(rect.get_x()+rect.get_width()/2,height),
                    xytext=(0,5),textcoords='offset pixels',ha='center',fontsize=12)

    todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    fromdate_text = f'{fromdate.day}/{fromdate.month}/{fromdate.year}'

    fig.suptitle(f"Where Customers Trade Most Often\n"
                 f"From {fromdate_text} to {todate_text}",
                 y=0.98,fontsize=20,fontfamily='Arial Unicode MS')

    plt.xticks(rotation=45)
    plt.savefig(join(dirname(__file__),'Turnover_by_Location.png'))

    return table


def graph_margin_usage_by_location(
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
    least: int = 500
):
    if todate is None:
        todate = core.time
    else:
        todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

    if fromdate is None:
        fromdate = dt.date(2000,1,1)
    else:
        fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

    location_series = core.info('all',subtype)['LOCATION'].dropna()
    location_series = location_series.loc[~(location_series=='Unknown')]

    idx = pd.IndexSlice

    nav_table = core.nav('all','all',subtype)
    nav_table = nav_table.loc[idx[fromdate:todate,:],:]
    nav_table.insert(0,'LOCATION',nav_table.index.get_level_values(1).map(location_series))
    sum_nav_over_trading_codes = nav_table.groupby(['TRADING_DATE','LOCATION']).sum()
    result_avg_nav = sum_nav_over_trading_codes.groupby(['LOCATION']).agg('mean')
    result_avg_nav = result_avg_nav.astype('int64')

    margin_table = core.margin('all','all',subtype)
    margin_table = margin_table.loc[idx[fromdate:todate,:],:]
    margin_table.insert(0,'LOCATION',margin_table.index.get_level_values(1).map(location_series))
    sum_margin_over_trading_codes = margin_table.groupby(['TRADING_DATE','LOCATION']).sum()
    result_avg_margin = sum_margin_over_trading_codes.groupby(['LOCATION']).agg('mean')
    result_avg_margin = result_avg_margin.astype('int64')

    table = pd.concat([result_avg_nav,result_avg_margin],axis=1)
    table.reset_index(inplace=True)
    table['MARGIN_USAGE_RATIO'] \
        = table['MARGIN_OUTSTANDINGS'].div(table['NAV'].add(table['MARGIN_OUTSTANDINGS'],fill_value=0),fill_value=0)

    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(locations=[x['LOCATION']],
                                                 subtype=subtype)),
                      axis=1)
    table = table.loc[table['COUNT'].ge(least)]
    table.sort_values(['MARGIN_USAGE_RATIO'],inplace=True,ascending=False)

    fig,ax = plt.subplots(1,1,figsize=(16,12))
    sns.barplot(x='LOCATION',y='MARGIN_USAGE_RATIO',data=table,ax=ax,palette=['tab:blue'])

    fig.subplots_adjust(top=0.92,bottom=0.11,left=0.08,right=0.92)
    ax.tick_params(axis='both',which='major',labelsize=13)
    ax.set_ylabel('MARGIN USAGE RATIO',labelpad=10,fontsize=15)
    ax.set_xlabel('LOCATION',labelpad=10,fontsize=15)
    ax.legend(prop={'size':14})

    ax.grid(True,lw=0.6,alpha=0.5)
    ax.set_axisbelow(True)

    rects = ax.patches
    # Make labels.
    labels = table['COUNT'].astype('int64').tolist()

    for rect,label in zip(rects,labels):
        height = rect.get_height()
        ax.annotate(f'{label}\nppl',(rect.get_x()+rect.get_width()/2,height),
                    xytext=(0,5),textcoords='offset pixels',ha='center',fontsize=12)

    todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    fromdate_text = f'{fromdate.day}/{fromdate.month}/{fromdate.year}'

    fig.suptitle(f"Where Customers Use Much Margin\n"
                 f"From {fromdate_text} to {todate_text}",
                 y=0.98,fontsize=20,fontfamily='Arial Unicode MS')

    plt.xticks(rotation=45)
    plt.savefig(join(dirname(__file__),'MarginUsage_by_Location.png'))

    return table


###############################################################################

def graph_turnover_at_location_by_age_gender(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
    least: int = 10,
    ci: int = 95,
    order: int = 2
):
    table = aggregation.trading_turnover_by('all',subtype,fromdate,todate)
    table = table.loc[pd.IndexSlice[location,:,:],:].droplevel(0)
    table.reset_index(inplace=True)
    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(locations=[location],
                                                 genders=[x['GENDER']],
                                                 ages=[x['AGE']],
                                                 subtype=subtype)),
                      axis=1)

    # only groups with at least 10 customers are considered, others are ignored
    table = table.loc[table['COUNT'].ge(least)]

    g = sns.lmplot(x='AGE',y=f'TRADING_TURNOVER',hue='GENDER',ci=ci,order=order,
                   data=table,height=10,aspect=1.1,
                   palette='Set1',truncate=False,legend=False)

    g.fig.subplots_adjust(top=0.88,bottom=0.08,left=0.08,right=0.92)

    g.set_xticklabels(fontsize=12)
    g.set_yticklabels(fontsize=12)
    g.set_ylabels(f'TRADING TURNOVER',labelpad=10,fontsize=13)
    g.set_xlabels('AGE',labelpad=10,fontsize=13)
    g.set(xlim=(20,None))
    g.set(ylim=(0,None))

    g.ax.legend(title='GENDER',loc='upper right',
                title_fontsize=13,fontsize=13,
                fancybox=True,shadow=True)

    if todate is None:
        todate = core.time
        todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    else:
        todate_text = f'{todate[-2:]}/{todate[5:7]}/{todate[:4]}'

    if fromdate is None:
        fromdate_text = '01/01/2018'
    else:
        fromdate_text = f'{fromdate[-2:]}/{fromdate[5:7]}/{fromdate[:4]}'

    g.fig.suptitle(f'Which Customer Groups Trade Most Frequently\n'
                   f'in {location}\n'
                   f'From {fromdate_text} to {todate_text}',
                   y=0.98,fontsize=18,fontfamily='Arial Unicode MS')

    plt.savefig(join(dirname(__file__),f'Turnover_{location}.png'))

    return table


def graph_margin_usage_at_location_by_age_gender(
    location: str = 'Ho Chi Minh City',
    fromdate: str = None,
    todate: str = None,
    subtype: list = 'all',
    least: int = 10,
    ci: int = 95,
    order: int = 2
):
    table = aggregation.margin_usage_ratio_by('all',subtype,fromdate,todate)
    table = table.loc[pd.IndexSlice[location,:,:],:].droplevel(0)
    table.reset_index(inplace=True)
    table['COUNT'] \
        = table.apply(lambda x:len(core.ofgroups(locations=[location],
                                                 genders=[x['GENDER']],
                                                 ages=[x['AGE']],
                                                 subtype=subtype)),
                      axis=1)

    # only groups with at least 10 customers are considered, others are ignored
    table = table.loc[table['COUNT'].ge(least)]

    g = sns.lmplot(x='AGE',y=f'MARGIN_USAGE_RATIO',hue='GENDER',ci=ci,order=order,
                   data=table,height=10,aspect=1.1,
                   palette='Set1',truncate=False,legend=False)

    g.fig.subplots_adjust(top=0.88,bottom=0.08,left=0.08,right=0.92)

    g.set_xticklabels(fontsize=12)
    g.set_yticklabels(fontsize=12)
    g.set_ylabels(f'MARGIN USAGE RATIO',labelpad=10,fontsize=13)
    g.set_xlabels('AGE',labelpad=10,fontsize=13)
    g.ax.legend(prop={'size':14})
    g.set(xlim=(20,None))
    g.set(ylim=(0,None))

    g.ax.legend(title='GENDER',loc='upper right',
                title_fontsize=13,fontsize=13,
                fancybox=True,shadow=True)

    if todate is None:
        todate = core.time
        todate_text = f'{todate.day}/{todate.month}/{todate.year}'
    else:
        todate_text = f'{todate[-2:]}/{todate[5:7]}/{todate[:4]}'

    if fromdate is None:
        fromdate_text = '01/01/2018'
    else:
        fromdate_text = f'{fromdate[-2:]}/{fromdate[5:7]}/{fromdate[:4]}'

    g.fig.suptitle(f'Which Customer Groups Like To Use Margin\n'
                   f'in {location}\n'
                   f'From {fromdate_text} to {todate_text}',
                   y=0.98,fontsize=18,fontfamily='Arial Unicode MS')

    plt.savefig(join(dirname(__file__),f'MarginUsage{location}.png'))

    return table


class Predict:

    def __init__(self,
                 location: str,
                 account_type: str,
                 n_jobs: int = 10000,
                 applied_irate: float = 0.135,
                 applied_frate: float = 0.0015,
                 model_startdate: str = '2020-01-01',
                 model_enddate: str = None
                 ):

        self.location = location
        self.account_type = account_type
        self.n_jobs = n_jobs
        self.applied_irate = applied_irate
        self.applied_frate = applied_frate
        self.model_startdate = model_startdate
        self.model_enddate = model_enddate

        self.subtype = [account_type,'individual','retail','active','trading']

        self.lower_age = 18
        self.upper_age = 65

        # MARGIN_USAGE_RATIO prediction by fitting y = a0 + a1.AGE + a2.AGE**2
        # TRADING_TURNOVER prediction by fitting y = a0 + a1.AGE + a2.AGE**2
        self.margin_usage_table = aggregation.margin_usage_ratio_by(
            'all',
            model_startdate=self.model_startdate,
            model_enddate=self.model_enddate,
            subtype=self.subtype,
            locations=[self.location],
        )

        self.trading_turnover_table = aggregation.trading_turnover_by(
            'all',
            model_startdate=self.model_startdate,
            model_enddate=self.model_enddate,
            subtype=self.subtype,
            locations=[self.location],
        )

        self.table = pd.concat([self.margin_usage_table,self.trading_turnover_table],axis=1)
        self.table.fillna(0,inplace=True)
        self.table.reset_index(inplace=True)
        self.table.drop(['LOCATION'],axis=1,inplace=True)
        self.table.insert(1,'AGE_SQR',self.table['AGE']**2)

        # work on male
        self.male_sub_result = pd.DataFrame(
            index=pd.MultiIndex.from_product([self.table['AGE'],range(n_jobs)],names=['AGE','STATE']),
            columns=[
                'fit_object',
                'PRED_MARGIN_USAGE_RATIO',
                'PRED_TRADING_TURNOVER'
            ])

        self.male_sub_result.reset_index(level=0,inplace=True)
        self.male_sub_result.insert(1,'AGE_SQR',self.male_sub_result['AGE']**2)

        male_table = self.table.loc[self.table['GENDER']=='Male']
        for state in range(n_jobs):
            train_set,test_set = train_test_split(
                male_table,
                test_size=1/4,
                random_state=state,
                shuffle=True
            )
            # train set
            X_train = train_set[['AGE','AGE_SQR']]
            y_train = train_set[['MARGIN_USAGE_RATIO','TRADING_TURNOVER']]
            # create Linear Regression object
            reg = LinearRegression()
            fit = reg.fit(X_train,y_train)
            self.male_sub_result.loc[state,'fit_object'] = fit

        r = self.male_sub_result.apply(lambda x:x['fit_object'].predict(
            [[x['AGE'],x['AGE_SQR']]])[0],axis=1)

        self.male_sub_result['PRED_MARGIN_USAGE_RATIO'] = r.str.get(0)
        self.male_sub_result['PRED_TRADING_TURNOVER'] = r.str.get(1)
        self.male_sub_result.drop(['AGE_SQR','fit_object'],axis=1,inplace=True)
        self.male_sub_result.reset_index(inplace=True)
        self.male_sub_result.drop('STATE',axis=1,inplace=True)
        self.male_sub_result.insert(0,'GENDER','Male')

        # work on female
        self.female_sub_result = pd.DataFrame(
            index=pd.MultiIndex.from_product([self.table['AGE'],range(n_jobs)],names=['AGE','STATE']),
            columns=[
                'fit_object',
                'PRED_MARGIN_USAGE_RATIO',
                'PRED_TRADING_TURNOVER'
            ])

        self.female_sub_result.reset_index(level=0,inplace=True)
        self.female_sub_result.insert(1,'AGE_SQR',self.female_sub_result['AGE']**2)

        female_sub_result = self.table.loc[self.table['GENDER']=='Female']
        for state in range(n_jobs):
            train_set,test_set = train_test_split(
                female_sub_result,
                test_size=1/4,
                random_state=state,
                shuffle=True
            )
            # train set
            X_train = train_set[['AGE','AGE_SQR']]
            y_train = train_set[['MARGIN_USAGE_RATIO','TRADING_TURNOVER']]
            # create Linear Regression object
            reg = LinearRegression()
            fit = reg.fit(X_train,y_train)
            self.female_sub_result.loc[state,'fit_object'] = fit

        r = self.female_sub_result.apply(lambda x:x['fit_object'].predict([[x['AGE'],x['AGE_SQR']]])[0],axis=1)

        self.female_sub_result['PRED_MARGIN_USAGE_RATIO'] = r.str.get(0)
        self.female_sub_result['PRED_TRADING_TURNOVER'] = r.str.get(1)
        self.female_sub_result.drop(['AGE_SQR','fit_object'],axis=1,inplace=True)
        self.female_sub_result.reset_index(inplace=True)
        self.female_sub_result.drop('STATE',axis=1,inplace=True)
        self.female_sub_result.insert(0,'GENDER','Female')

        # combine
        self.result_table = pd.concat([self.male_sub_result,self.female_sub_result],axis=0)
        self.result_table.set_index(['GENDER','AGE'],inplace=True)

        # remove abnormal accounts who have extremely great NAV
        nav_hist = core.nav('all')
        outliers = set(nav_hist.loc[nav_hist['NAV']>2e9].index.get_level_values(1))
        all_customer = set(core.AllCustomer)
        self.removed_outliers = list(all_customer-outliers)

        # count business days
        if self.model_enddate is None:
            self.model_enddate = core.time.strftime('%Y-%m-%d')

        start_date = self.model_startdate
        self.n_bdays = 0
        while start_date!=self.model_enddate:
            start_date = bdate(start_date,1)
            self.n_bdays += 1

        location = self.location
        subtype = self.subtype
        removed_outliers = self.removed_outliers

        def count_customers(df):
            customers = core.ofgroups(locations=[location],
                                      genders=[df['GENDER']],
                                      ages=[df['AGE']],
                                      subtype=subtype)
            customers = list(set(customers)&set(removed_outliers))
            return len(customers)

        nav_table = aggregation.on_nav_by(
            self.removed_outliers,
            'mean',
            model_startdate=self.model_startdate,
            model_enddate=self.model_enddate,
            subtype=self.subtype,
            locations=[self.location],
        )
        nav_table = nav_table.droplevel(0,axis=0)
        nav_table.reset_index(inplace=True)

        nav_table['N_CUSTOMERS'] = nav_table.apply(count_customers,axis=1)
        # remove groups with less than 5 customers
        # nav_table = nav_table.loc[nav_table['N_CUSTOMERS'].ge(5)]
        nav_table['AVG_DAILY_NAV'] = nav_table['NAV']/nav_table['N_CUSTOMERS']
        nav_table.insert(2,'AGE_SQR',nav_table['AGE']**2)

        self.male_nav_table = nav_table.loc[nav_table['GENDER']=='Male'].copy()
        fit = LinearRegression().fit(self.male_nav_table[['AGE','AGE_SQR']],self.male_nav_table['AVG_DAILY_NAV'])
        self.male_nav_predict = pd.DataFrame({'AGE':range(self.lower_age,self.upper_age,1)})
        self.male_nav_predict['AGE_SQR'] = self.male_nav_predict['AGE']**2
        self.male_nav_predict.insert(0,'GENDER','Male')
        self.male_nav_predict['PRED_NAV'] = fit.predict(self.male_nav_predict[['AGE','AGE_SQR']])
        self.male_nav_predict['PRED_NAV'] = self.male_nav_predict['PRED_NAV'].map(lambda x:max(x,0))
        self.male_nav_predict['PRED_NAV'] = self.male_nav_predict['PRED_NAV'].astype(np.int64)

        self.female_nav_table = nav_table.loc[nav_table['GENDER']=='Female'].copy()
        fit = LinearRegression().fit(self.female_nav_table[['AGE','AGE_SQR']],self.female_nav_table['AVG_DAILY_NAV'])
        self.female_nav_predict = pd.DataFrame({'AGE':range(self.lower_age,self.upper_age,1)})
        self.female_nav_predict['AGE_SQR'] = self.female_nav_predict['AGE']**2
        self.female_nav_predict.insert(0,'GENDER','Female')
        self.female_nav_predict['PRED_NAV'] = fit.predict(self.female_nav_predict[['AGE','AGE_SQR']])
        self.female_nav_predict['PRED_NAV'] = self.female_nav_predict['PRED_NAV'].map(lambda x:max(x,0))
        self.female_nav_predict['PRED_NAV'] = self.female_nav_predict['PRED_NAV'].astype(np.int64)

        self.nav_result_table = pd.concat([self.male_nav_predict,self.female_nav_predict],axis=0)
        self.nav_result_table.drop(['AGE_SQR'],axis=1,inplace=True)
        self.nav_result_table.set_index(['GENDER','AGE'],inplace=True)

    def monthly_income(self,
                       gender: str,
                       age: int,
                       full_output: bool = False,
                       ):

        lower_quantile = 0.001
        upper_quantile = 0.999

        result_table = self.result_table.loc[(gender,age),:].copy()
        result_table.reset_index(inplace=True)

        if gender=='Male':
            avg_daily_nav = self.male_nav_predict
        elif gender=='Female':
            avg_daily_nav = self.female_nav_predict
        else:
            raise TypeError('gender must be either "Male" or "Female"')

        avg_daily_nav = avg_daily_nav.loc[avg_daily_nav['AGE']==age,'PRED_NAV'].squeeze()

        result_table['PRED_INTEREST_INCOME'] \
            = avg_daily_nav*result_table['PRED_MARGIN_USAGE_RATIO'] \
              /(1-result_table['PRED_MARGIN_USAGE_RATIO'])*self.applied_irate/360*30

        result_table['PRED_FEE_INCOME'] \
            = result_table['PRED_TRADING_TURNOVER'] \
              *avg_daily_nav/(1-result_table['PRED_MARGIN_USAGE_RATIO'])*self.applied_frate/self.n_bdays*22

        result_table.fillna(0,inplace=True)

        pred_interest_mean = int(result_table['PRED_INTEREST_INCOME'].mean())
        pred_interest_lower_percentile = int(result_table['PRED_INTEREST_INCOME'].quantile(lower_quantile))
        pred_interest_upper_percentile = int(result_table['PRED_INTEREST_INCOME'].quantile(upper_quantile))

        pred_fee_mean = int(result_table['PRED_FEE_INCOME'].mean())
        pred_fee_lower_percentile = int(result_table['PRED_FEE_INCOME'].quantile(lower_quantile))
        pred_fee_upper_percentile = int(result_table['PRED_FEE_INCOME'].quantile(upper_quantile))

        pred_income_mean = pred_interest_mean+pred_fee_mean
        pred_income_lower_percentile = pred_interest_lower_percentile+pred_fee_lower_percentile
        pred_income_upper_percentile = pred_interest_upper_percentile+pred_fee_upper_percentile

        result = pd.DataFrame(
            {
                'interest':[pred_interest_lower_percentile,
                            pred_interest_mean,
                            pred_interest_upper_percentile],
                'fee':[pred_fee_lower_percentile,
                       pred_fee_mean,
                       pred_fee_upper_percentile],
                'total':[pred_income_lower_percentile,
                         pred_income_mean,
                         pred_income_upper_percentile]
            },
            index=[f'{lower_quantile*100}_percentile',
                   'mean',
                   f'{upper_quantile*100}_percentile'])

        # GRAPHING:
        if self.account_type=='margin':
            graph_table = result_table[['PRED_INTEREST_INCOME','PRED_FEE_INCOME']]/1000
            g = sns.jointplot(data=graph_table,
                              x='PRED_FEE_INCOME',
                              y='PRED_INTEREST_INCOME',
                              kind='hex',gridsize=90,height=8,
                              space=0,
                              xlim=(pred_fee_lower_percentile/1000,
                                    pred_fee_upper_percentile/1000),
                              ylim=(pred_interest_lower_percentile/1000,
                                    pred_interest_upper_percentile/1000))
            g.plot_joint(sns.kdeplot,color='tab:blue',alpha=0.5,levels=[0.05,0.1,0.25,0.5,0.75])
            g.set_axis_labels(xlabel='Predicted Monthly Fee Income (VND\'000)',
                              ylabel='Predicted Monthly Interest Income (VND\'000)',
                              fontsize=10)
            g.fig.suptitle(f'Monthly Income Prediction \n'
                           f'Location: {self.location}, Age: {age}, Gender: {gender}, '
                           f'Account: {self.account_type.capitalize()}',
                           fontsize=7,fontfamily='Arial',color='black',
                           ha='left',position=(0,1))
            g.fig.tight_layout()

            g.savefig(join(dirname(__file__),
                           f'PredictIncome_{self.account_type.capitalize()}_{self.location}_{gender}_{age}.png'))

        elif self.account_type=='normal':
            graph_table = result_table[['PRED_FEE_INCOME']]/1000
            fig,ax = plt.subplots(figsize=(7,6))
            sns.histplot(data=graph_table,x='PRED_FEE_INCOME',stat='density',
                         color='tab:blue',alpha=1,edgecolor='black',ax=ax)
            sns.kdeplot(data=graph_table,x='PRED_FEE_INCOME',color='black',lw=2,ax=ax)
            ax.set_xlabel('Predicted Monthly Fee Income (VND\'000)',fontsize=10.5)
            ax.set_ylabel('')
            plt.tick_params(
                axis='y',
                which='both',
                left=False,
                labelleft=False)
            ax.set_title(f'Monthly Income Prediction - Age: {age}',
                         fontsize=15,fontfamily='Arial',color='black',pad=12)
            ax.annotate(f'Location: {self.location}, Age: {age}, Gender: {gender}, '
                        f'Account: {self.account_type.capitalize()}',
                        xy=(0.005,0.98),xycoords='axes fraction',fontsize=8)
            fig.subplots_adjust(left=0.01,right=0.99,
                                bottom=0.1,top=0.9)
            plt.savefig(join(dirname(__file__),
                             f'PredictIncome_{self.account_type.capitalize()}_{self.location}_{gender}_{age}.png'))
        else:
            raise TypeError('account_type must be either "margin" or "normal"')

        if full_output is False:
            return result
        else:
            return result_table

    def summary(self):

        result_table = self.result_table.query(f'AGE >= {self.lower_age} and AGE <= {self.upper_age}').copy()

        nav_result_table = self.nav_result_table
        result_table['PRED_NAV'] = nav_result_table.squeeze(axis=1)
        result_table.reset_index(inplace=True)

        result_table['PRED_INTEREST_INCOME'] \
            = result_table['PRED_NAV']*result_table['PRED_MARGIN_USAGE_RATIO'] \
              /(1-result_table['PRED_MARGIN_USAGE_RATIO'])*self.applied_irate/360*30
        result_table['PRED_INTEREST_INCOME'] \
            = result_table['PRED_INTEREST_INCOME'].map(lambda x:max(x,0))

        result_table['PRED_FEE_INCOME'] \
            = result_table['PRED_TRADING_TURNOVER'] \
              *result_table['PRED_NAV']/(1-result_table['PRED_MARGIN_USAGE_RATIO'])*self.applied_frate/self.n_bdays*22
        result_table['PRED_FEE_INCOME'] \
            = result_table['PRED_FEE_INCOME'].map(lambda x:max(x,0))

        result_table.fillna(0,inplace=True)
        result_table[['PRED_INTEREST_INCOME','PRED_FEE_INCOME']] /= 1e3
        result_table = result_table.convert_dtypes()

        # GRAPHING
        fig_fee,ax_fee = plt.subplots(figsize=(22,9))
        sns.boxplot(data=result_table,x='AGE',y='PRED_FEE_INCOME',hue='GENDER',
                    fliersize=3,ax=ax_fee)
        sns.despine(offset=8)
        ax_fee.set_title(f'Predicted Monthly Fee Income\n'
                         f'in {self.location}',fontsize=13)
        ax_fee.set_axisbelow(True)
        ax_fee.grid(True,color='gray',alpha=0.2)
        ax_fee.yaxis.offsetText.set_visible(False)
        ax_fee.annotate('VND\'000',xy=(-0.035,1.03),xycoords='axes fraction',fontsize=9)
        fig_fee.tight_layout()
        fig_fee.savefig(join(dirname(__file__),f'SummaryFee_{self.account_type.capitalize()}_{self.location}.png'))

        fig_interest,ax_interest = plt.subplots(figsize=(22,9))
        sns.boxplot(data=result_table,x='AGE',y='PRED_INTEREST_INCOME',hue='GENDER',
                    fliersize=3,ax=ax_interest)
        sns.despine(offset=8)
        ax_interest.set_title(f'Predicted Monthly Interest Income\n'
                              f'in {self.location}',fontsize=13)
        ax_interest.set_axisbelow(True)
        ax_interest.grid(True,color='gray',alpha=0.2)
        ax_interest.yaxis.offsetText.set_visible(False)
        ax_interest.annotate('VND\'000',xy=(-0.035,1.03),xycoords='axes fraction',fontsize=9)
        fig_interest.tight_layout()
        fig_interest.savefig(
            join(dirname(__file__),f'SummaryInterest{self.account_type.capitalize()}_{self.location}.png'))

        return result_table
