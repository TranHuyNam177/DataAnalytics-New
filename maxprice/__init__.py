from request.stock import *
from breakeven_price.monte_carlo import *


def lowest_coefficients_fitting(
    tickers: list = 'all',
    model_period=fa.periods[-1],
    standard: str = 'bics',
    level: int = 1,
    savefigure: bool = True,
    savetable: bool = True
) -> pd.DataFrame:
    """
    This model return the lowest possible stock price given the observed
    relationship between next period's lowest price and current period's
    credit score
  
    :param tickers: ticker list
    :param model_period: the period of calculation, usually is fa.period[-1],
     else of back-testing only
    :param standard: industry classification for credit rating's model
    :param level: industry level for credit rating's model
    :param savefigure: whether to save the figure
    :param savetable: whether to save the table
    """

    # pre-prcessing
    input_folder = join(dirname(dirname(realpath(__file__))),
                        'credit_rating','result')
    if tickers=='all':
        tickers = fa.tickers()

    maxprice_dict = dict()
    for ticker in tickers:
        print(f'Evaluating {ticker} :::')
        segment = fa.segment(ticker)
        file_name = f'result_table_{segment}.csv'
        rating_file = join(input_folder,file_name)

        destination_dir = join(dirname(dirname(realpath(__file__))),
                               'maxprice','result')

        # MaxPrice calculation
        rating_result = pd.read_csv(rating_file,index_col='ticker')
        if segment=='gen':
            rating_result = rating_result.loc[rating_result['level']==f'{standard}_l{level}'].T
            rating_result.drop(index=['standard','level','industry'],inplace=True)
        else:
            rating_result = rating_result.T

        # (some ticker were excluded by not having data for CreditRating)
        try:
            score = rating_result[[ticker]].copy();
            score.columns = ['score']
        except KeyError:
            print(f'{ticker} is not evaluated by credit rating')
            continue
        try:
            highlow = ta.prhighlow(ticker,fquarters=1)
        except ValueError:
            print(f'{ticker} has no historical data in tested period')
            continue
        plow = highlow['low'][['low_price']]*1000
        plow = plow.shift(periods=-1,axis=0).loc[score.index]
        # shift 1 period because we want 1-period forward-looking
        phigh = highlow['high'][['high_price']]*1000
        phigh = phigh.shift(periods=-1,axis=0).loc[score.index]
        # shift 1 period because we want 1-period forward-looking
        table = pd.concat([pd.DataFrame(score),plow],axis=1)
        table.dropna(inplace=True)
        table = table.loc[table['low_price']!=0]
        table = table.loc[:model_period]
        if table.empty:
            print(f'{ticker} does not have enough data')
            continue
        table = table.astype(np.int64)
        last_score = table['score'].iloc[-1]
        table = table.head(-1)  # because we only use observed/happened data
        table['multiple'] = table['low_price']/table['score']
        table['multiple'] = table['multiple'].replace(np.inf,np.nan)
        min_multiple = table['multiple'].mean()

        def fprice(score):
            return min_multiple*score

        maxprice = fprice(last_score)
        if pd.isna(maxprice):
            print(f'{ticker} credit score is NaN')
            continue

        eopdate = seopdate(model_period)[1]
        try:
            refprice = ta.hist(ticker,eopdate,eopdate)['close'].iloc[0]*1000
            discount = maxprice/refprice-1
        except ValueError:
            print(f'{ticker} has no historical data in {eopdate}, hence it has no ref price')
            continue

        def graph_maxprice1():

            ylow = plow.dropna()
            yhigh = phigh.dropna()
            x = rating_result.loc[:model_period,ticker].dropna().to_frame()
            x = x.loc[x.index.intersection(ylow.index).intersection(yhigh.index)]

            fig,ax = plt.subplots(1,1,figsize=(8,6))

            for i in range(x.shape[0]):

                if x.index[i]==model_period:
                    c = 'tab:red'
                    a = 1
                    fw = 'bold'
                else:
                    c = 'black'
                    a = 0.8
                    fw = 'normal'

                ax.scatter(x.iloc[i,0],ylow.iloc[i,0],
                           label='Next Quarter\'s Lowest Price',
                           color='firebrick',edgecolors='firebrick',
                           marker='^',alpha=a)
                ax.scatter(x.iloc[i,0],yhigh.iloc[i,0],
                           label='Next Quarter\'s Highest Price',
                           color='forestgreen',edgecolors='forestgreen',
                           marker='v',alpha=a)
                ax.plot([x.iloc[i,0],x.iloc[i,0]],
                        [ylow.iloc[i,0],yhigh.iloc[i,0]],
                        color='black',alpha=a)
                ax.annotate(x.index[i],xy=(x.iloc[i,0],ylow.iloc[i,0]),
                            xycoords=ax.transData,
                            xytext=(3,5),
                            textcoords='offset pixels',
                            ha='left',fontsize=7,
                            color=c,alpha=a,fontweight=fw)
                ax.set_xlabel('Credit Score',labelpad=10)
                ax.set_ylabel('Stock Price',rotation=90,labelpad=5)
                ax.xaxis.grid(True,alpha=0.05)
                ax.yaxis.grid(True,alpha=0.2)

            ax.xaxis.set_major_locator(MaxNLocator(integer=True,
                                                   steps=[1,2,5]))
            ax.yaxis.set_major_formatter(FuncFormatter(priceKformat))
            ax.axhline(maxprice,ls='--',linewidth=0.5,color='tab:red')
            ax.annotate(f'Max Price = {adjprice(maxprice)}',
                        xy=(0.7,maxprice),
                        xycoords=transforms.blended_transform_factory(
                            ax.transAxes,ax.transData),
                        xytext=(0,3),
                        textcoords='offset pixels',
                        ha='left',fontsize=7,
                        color='tab:red',fontweight='bold')

            ax.set_title(f'{ticker}\n Historical Relation between\n'
                         f'Credit Score and Stock Price',
                         fontfamily='Times New Roman',fontsize=13,
                         fontweight='bold',color='black')

            handles,labels = ax.get_legend_handles_labels()
            handles = handles[1:3];
            labels = labels[1:3]
            ax.legend(handles,labels,loc='best',
                      ncol=1,fontsize=8,numpoints=1,
                      framealpha=1)

            if savefigure is True:
                plt.savefig(join(destination_dir,
                                 'Chart1',
                                 f'{ticker}_chart1.png'),
                            bbox_inches='tight')

        def graph_maxprice2():

            nonlocal rating_result
            try:
                peers = fa.peers(ticker,standard,level+1)
            except KeyError:
                print(f'{ticker} is not in Classification File from Bloomberg hence Chart 2 can\'t be drawn')
                return

            peers = list(set(peers)&set(rating_result.columns))
            # (some ticker were excluded by not having data for CreditRating)
            rating_result = rating_result[peers]
            rating_result = rating_result.loc[:model_period]

            delta_scores = rating_result.diff(periods=1,axis=0)
            delta_scores = delta_scores.loc[[model_period]].T

            fig,ax = plt.subplots(1,1,figsize=(8,6))

            table = delta_scores.copy()
            table.dropna(inplace=True)
            table.columns = ['score']
            table[['low_return','high_return']] = np.nan

            folder = 'price';
            file = 'prhighlow.csv'
            highlow = pd.read_csv(join(dirname(dirname(realpath(__file__))),
                                       'database',folder,file),
                                  index_col=[0])
            highlow = highlow[model_period]
            for ticker_ in table.index:
                if isinstance(highlow.loc[ticker_],str):
                    pass
                else:
                    continue
                string = highlow.loc[ticker_].replace('(','').replace(')','')
                string = string.split(',')
                low_r = float(string[3])
                high_r = float(string[4])
                if pd.isna(low_r) or pd.isna(high_r):
                    continue
                table.loc[ticker_,'low_return'] = low_r
                table.loc[ticker_,'high_return'] = high_r

                xloc = table.loc[ticker_,'score']
                ylow = table.loc[ticker_,'low_return']
                yhigh = table.loc[ticker_,'high_return']

                if ticker_==ticker:
                    c = 'tab:red'
                    a = 1
                    fw = 'bold'
                else:
                    c = 'black'
                    a = 0.2
                    fw = 'normal'

                last_period = fa.periods[-1]
                if model_period==last_period:
                    label_low \
                        = f'Lowest Return since eop {last_period}'
                    label_high \
                        = f'Highest Return since eop {last_period}'
                else:
                    label_low = f'Lowest Return in ' \
                                f'{period_cal(model_period,0,1)}'
                    label_high = f'Lowest Return in ' \
                                 f'{period_cal(model_period,0,1)}'

                ax.scatter(xloc,ylow,
                           label=label_low,
                           color='firebrick',edgecolors='firebrick',
                           marker='^',alpha=a)
                ax.scatter(xloc,yhigh,
                           label=label_high,
                           color='forestgreen',edgecolors='forestgreen',
                           marker='v',alpha=a)
                ax.plot([xloc,xloc],[ylow,yhigh],
                        color='black',alpha=a)
                ax.annotate(ticker_,xy=(xloc,ylow),
                            xycoords=ax.transData,
                            xytext=(0,-20),
                            textcoords='offset pixels',
                            ha='center',fontsize=7,
                            color=c,alpha=a,fontweight=fw)

            ax.xaxis.set_major_locator(MaxNLocator(min_n_ticks=10,
                                                   integer=True,
                                                   steps=[1,2,5]))
            ax.yaxis.set_major_formatter(mpl.ticker.PercentFormatter(1.0))

            ax.axhline(0,ls='--',linewidth=0.5,color='black')
            ax.axvline(0,ls='--',linewidth=0.5,color='black')

            ax.axhline(discount,ls='--',
                       linewidth=0.5,color='tab:red')
            ax.annotate(f'Discount Rate = {discount:.2%}',
                        xy=(0.7,discount),
                        xycoords=transforms.blended_transform_factory(
                            ax.transAxes,ax.transData),
                        xytext=(0,3),
                        textcoords='offset pixels',
                        ha='left',fontsize=7,
                        color='tab:red',fontweight='bold')

            subtext_cur = '{'+f'{model_period}'+'}'
            subtext_pre = '{'+f'{period_cal(model_period,quarters=-1)}'+'}'
            ax.set_xlabel(fr'$CreditScore_{subtext_cur}  -  $'
                          +fr'$CreditScore_{subtext_pre}$',labelpad=10)
            ax.set_ylabel('Return',labelpad=2,rotation=90)
            ax.xaxis.grid(True,alpha=0.05)
            ax.yaxis.grid(True,alpha=0.2)

            ax.set_title(f'{ticker}\n Cross-sectional Relation between\n'
                         r'$\Delta$Credit Score and Return',
                         fontfamily='Times New Roman',fontsize=13,
                         fontweight='bold',color='black')

            handles,labels = ax.get_legend_handles_labels()
            handles = handles[1:3];
            labels = labels[1:3]
            ax.legend(handles,labels,loc='best',
                      ncol=1,fontsize=8,numpoints=1,
                      framealpha=1)

            if savefigure is True:
                plt.savefig(join(destination_dir,
                                 'Chart2',
                                 f'{ticker}_chart2.png'),
                            bbox_inches='tight')

            return

        def graph_comment():

            score_cur = rating_result.loc[model_period,ticker]
            if pd.isna(score_cur):
                print(f'{ticker} is not evaluated by credit rating')
                return

            score_pre \
                = rating_result.loc[period_cal(model_period,0,-1),ticker]

            if score_cur>75:
                group = 'A'
            elif score_cur>50:
                group = 'B'
            elif score_cur>25:
                group = 'C'
            else:
                group = 'D'

            def ftext(score_cur,score_pre):
                if score_cur>score_pre:
                    t = f'upgraded {int(abs(score_cur-score_pre))} points.'
                elif score_cur<score_pre:
                    t = f'downgraded {int(abs(score_cur-score_pre))} points.'
                else:
                    t = 'remained unchanged.'
                return t

            fig,ax = plt.subplots(1,1,figsize=(7.8,2),
                                  edgecolor='black')
            ax.annotate(f'Comment from Hiep:',
                        xy=(0,1),xycoords=ax.transAxes,
                        xytext=(8,-8),textcoords='offset points',
                        ha='left',va='top',
                        fontsize=11,fontstyle='oblique',
                        fontfamily='Times New Roman')
            ax.annotate(f'{ticker} has Credit Score of '
                        f'{int(score_cur)} (Group {group}) in the last period, '
                        +ftext(score_cur,score_pre)+'\n'
                                                    f'Max Price is conservatively set at '
                                                    f'{int(maxprice):,}VND, which takes into account '
                                                    f'all of the company\'s'
                                                    f'\nfinancial aspects.\n',
                        xy=(0,1),xycoords=ax.transAxes,
                        xytext=(8,-34),textcoords='offset points',
                        ha='left',va='top',
                        fontsize=11,fontstyle='italic',
                        fontfamily='Times New Roman')
            ax.annotate(f'This Max Price is equivalent to a discount of'
                        f' {discount:.2%} '
                        f'compared to {ticker}\'s open price in '
                        f'{period_cal(model_period,0,1)}.',
                        xy=(0,1),xycoords=ax.transAxes,
                        xytext=(8,-73),textcoords='offset points',
                        ha='left',va='top',
                        fontsize=11,fontstyle='italic',
                        fontfamily='Times New Roman')
            ax.annotate(f'It should be noticed that this price'
                        f' does not reflect {ticker}\'s market risk.',
                        xy=(0,1),xycoords=ax.transAxes,
                        xytext=(8,-89),textcoords='offset points',
                        ha='left',va='top',
                        fontsize=11,fontstyle='italic',
                        fontfamily='Times New Roman')

            ax.tick_params(bottom=False,
                           left=False,
                           labelbottom=False,
                           labelleft=False)

            if savefigure is True:
                plt.savefig(join(destination_dir,'Comment',
                                 f'{ticker}_comment.png'),
                            bbox_inches='tight')

        if savefigure is True:
            graph_maxprice1()
            graph_maxprice2()
            graph_comment()

        maxprice_dict[ticker] \
            = int(adjprice(maxprice).replace(',',''))

        print(f'Finished {ticker} :::')
        print('--------------------')

    result_table \
        = pd.DataFrame(maxprice_dict,index=['price']).T

    if savetable is True:
        result_table.to_csv(
            join(realpath(dirname(__file__)),'result','result.csv'))

    return result_table
