import pandas as pd

from request.stock import *
from request import *
import scrape_cw_stock
import scrape_missing_info

rf = 0.045


class NoBacktestData(Exception):
    pass


class InvalidDate(Exception):
    pass


class NoCWAvailable(Exception):
    pass


class issuance:

    def __init__(
        self,
        ticker: str,
        K: int,
        T: str,
        t: str,
        k: float,
        n: int,
        r: float,
        isbacktest: bool = False
    ):

        self.ticker = ticker
        self.K = K
        self.T = T
        self.t = t
        self.k = k
        self.n = n
        self.isbacktest = isbacktest

        self.cw_value = None
        self.hedge_table = None

        if t=='now':
            t = date.today().strftime('%Y-%m-%d')

        if dt.datetime.strptime(t,'%Y-%m-%d')>dt.datetime.strptime(T,'%Y-%m-%d'):
            raise InvalidDate('t must be before T')

        if self.isbacktest is False:
            price = ta.hist(ticker)[['trading_date','close']].set_index('trading_date')*1000
        else:
            price = pd.read_csv(join(dirname(realpath(__file__)),'backtest_data','StockPrice.csv'))
            price = price.loc[price['ticker']==self.ticker,['trading_date','close']].set_index('trading_date')
        price = price.astype(np.int64)
        price.sort_index(inplace=True)
        price['return'] = np.log(price['close']/price['close'].shift(1))
        price = price.loc[:t]
        price = price.tail(90)

        self.sigma = price['return'].std()

        year_days = 0
        start = f'{t[:4]}-01-01'
        end = f'{t[:4]}-12-31'
        while dt.datetime.strptime(start,'%Y-%m-%d')<dt.datetime.strptime(end,'%Y-%m-%d'):
            year_days += 1
            start = bdate(start,1)

        # Compute daily risk_management-free rate
        self.r = np.log(1+r)/year_days

        remaining_days = 0
        start = t
        while dt.datetime.strptime(start,'%Y-%m-%d')<dt.datetime.strptime(T,'%Y-%m-%d'):
            remaining_days += 1
            start = bdate(start,1)

        self.delta_t = remaining_days
        self.S = price.iloc[-1,price.columns.get_loc('close')]

        self.d1 = (np.log(self.S/self.K)+(self.r+0.5*self.sigma**2)*self.delta_t)/(self.sigma*np.sqrt(self.delta_t))
        self.d2 = self.d1-self.sigma*np.sqrt(self.delta_t)
        self.delta = stats.norm.cdf(self.d1)
        print(t,'ticker',ticker,'S',self.S,'K',self.K,'r',self.r,'sigma',self.sigma,'delta_t',self.delta_t,'d1',self.d1)

    def valuation(
        self,
    ) -> int:

        self.cw_value = (self.S*stats.norm.cdf(self.d1)-self.K*np.exp(-self.r*self.delta_t)*stats.norm.cdf(self.d2))/self.k
        self.cw_value = int(self.cw_value)

        return self.cw_value

    def hedge(
        self,
        K_: int,
        T_: str,
        k_: float,
        risk_by_underlying='all'
    ) -> pd.DataFrame:

        # Compute delta of third-party CW
        market_delta = issuance(self.ticker,K_,T_,self.t,k_,self.n,self.r,self.isbacktest).delta

        # Create result table
        self.hedge_table = pd.DataFrame(
            columns=[
                'n_stock',
                'n_hedging_cw',
                'pct_hedged_by_stock'
            ]
        )

        if risk_by_underlying!='all':
            self.hedge_table['pct_hedged_by_stock'] = [risk_by_underlying]
        else:
            self.hedge_table['pct_hedged_by_stock'] = np.round(np.arange(0,1.05,0.05),2)
        self.hedge_table['n_stock'] = self.n*self.delta*self.hedge_table['pct_hedged_by_stock']
        self.hedge_table['n_hedging_cw'] = (self.n*self.delta-self.hedge_table['n_stock'])/market_delta

        self.hedge_table['n_hedging_cw'] *= k_
        self.hedge_table['n_stock'] /= self.k

        self.hedge_table.set_index('pct_hedged_by_stock',inplace=True)
        # self.hedge_table= np.round(self.hedge_table,0)

        return self.hedge_table


class preprocessing:

    def __init__(
        self,
        ticker):

        self.info_table = pd.read_csv(join(dirname(realpath(__file__)),'backtest_data','InfoTable.csv'))
        ticker_filter = self.info_table['cw'].map(lambda x:x[1:4]).isin([ticker])
        self.info_table = self.info_table.loc[ticker_filter]
        self.info_table.drop(['term','issue_date'],axis=1,inplace=True)
        self.info_table.set_index('cw',inplace=True)
        self.info_table['t0'] = self.info_table['t0'].map(lambda x:f'{x[-4:]}-{x[3:5]}-{x[:2]}')
        self.info_table['T'] = self.info_table['T'].map(lambda x:f'{x[-4:]}-{x[3:5]}-{x[:2]}')
        self.info_table[['P0','K']] = self.info_table[['P0','K']].applymap(lambda x:float(x.replace(',','')))

        if not ticker_filter.any():
            raise NoBacktestData(f'{ticker} has no backtest data')

        self.test_data = pd.read_csv(
            join(dirname(realpath(__file__)),'backtest_data','PriceTable.csv'),
            usecols=['<Ticker>','<Close>','<Date>']
        )
        self.test_data.rename(
            {
                '<Ticker>':'CW',
                '<Date>':'Date',
                '<Close>':'Price'
            },
            axis=1,inplace=True
        )
        ticker_filter = self.test_data['CW'].map(lambda x:x[1:4]).isin([ticker])
        self.test_data = self.test_data.loc[ticker_filter]
        self.test_data.sort_values(['CW','Date'],inplace=True)
        self.test_data.set_index('CW',inplace=True)
        self.test_data['Date'] = self.test_data['Date'].map(lambda x:f'{x}')
        self.test_data['Date'] = self.test_data['Date'].map(lambda x:dt.date(int(x[:4]),int(x[4:6]),int(x[-2:])))
        self.test_dates = self.test_data['Date'].drop_duplicates()
        self.test_cases = self.test_data.index.unique().to_list()

        self.price_table = pd.DataFrame(columns=self.test_cases,index=self.test_dates)
        for test_case in self.test_cases:
            self.price_table[test_case] = self.test_data.loc[test_case].set_index('Date').squeeze()
        self.price_table *= 1000

        self.dividend_date = pd.read_excel(join(dirname(realpath(__file__)),'backtest_data','ExDividendDate.xlsx'),
                                           index_col='Ticker',usecols=['Ticker','ExDividendDate'],squeeze=True)
        self.dividend_date = self.dividend_date.map(lambda x:dt.date(int(x[-4:]),int(x[3:5]),int(x[0:2])))
        self.dividend_date = self.dividend_date.loc[
            self.dividend_date.index.isin([ticker])].values  # allow empty Series


class backtest(preprocessing):

    def __init__(
        self,
        case: str,
        n: int,
        r: float,
        risk_by_underlying: float = 0.8
    ):

        self.ticker = case[1:4]
        super().__init__(self.ticker)
        self.case = case
        self.K = self.info_table.loc[case,'K']
        self.T = self.info_table.loc[case,'T']
        self.P0 = self.info_table.loc[case,'P0']
        self.k = self.info_table.loc[case,'k']
        self.n = n
        self.r = r
        self.risk_by_underlying = risk_by_underlying
        self.backtest_table = None
        self.days = None
        self.MIRR = None
        self.status = None

    def select_cw(
        self,
        t: str
    ):

        t = dt.date(int(t[:4]),int(t[5:7]),int(t[-2:]))
        market_cws = [cw for cw in self.test_cases if cw!=self.case]
        time_table = self.test_data.loc[self.test_data['Date']>=t,'Date']
        list_cws = time_table.loc[time_table==t].index
        cws = set(list_cws)&set(market_cws)
        time_table = time_table.loc[cws]
        if len(time_table)==0:
            selected_cw = 'No CW Available'
            days_left = np.nan
        else:
            count_days = time_table.index.value_counts()-1
            selected_cw = count_days.idxmax()
            days_left = count_days.loc[selected_cw]

        return selected_cw,days_left

    def run(
        self,
    ) -> pd.DataFrame:

        case_price = self.price_table[self.case].dropna()

        self.backtest_table = pd.DataFrame(
            columns=['date','n_issued_cw','selected_cw','days_left','n_stock',
                     'n_hedging_cw','delta_n_stock','delta_n_hedging_cw',
                     'p_issued_cw','p_stock','p_hedging_cw','dividend',
                     'issued_cw_delta','market_cw_delta','pct_position','cash_flow'])

        self.backtest_table['date'] = case_price.index.sort_values()
        self.backtest_table['n_issued_cw'] = self.n

        for row in self.backtest_table.index:
            if row>=1:
                previous_days_left = self.backtest_table.loc[row-1,'days_left']
                if previous_days_left>1:
                    selected_cw = self.backtest_table.loc[row-1,'selected_cw']
                    self.backtest_table.loc[row,'selected_cw'] = selected_cw
                    self.backtest_table.loc[row,'days_left'] = previous_days_left-1
                    continue
            t = self.backtest_table.loc[row,'date'].strftime('%Y-%m-%d')
            result = pd.Series(self.select_cw(t),index=['selected_cw','days_left'])
            if result.loc['days_left']==0:  # TH: CW con lai duy nhat tren thi truong la CW co days_lef = 0 -> ko chon
                self.backtest_table.loc[row,['selected_cw','days_left']] = ('No CW Available',0)
            else:
                self.backtest_table.loc[row,['selected_cw','days_left']] = result

        find_K_ = lambda selected_cw:self.info_table.loc[selected_cw,'K']
        find_T_ = lambda selected_cw:self.info_table.loc[selected_cw,'T']
        find_k_ = lambda selected_cw:self.info_table.loc[selected_cw,'k']

        rows = self.backtest_table['date'].index
        for row in range(0,len(rows)):
            t = self.backtest_table.loc[row,'date'].strftime('%Y-%m-%d')
            object1 = issuance(self.ticker,self.K,self.T,t,self.k,self.n,self.r,True)
            self.backtest_table.iloc[row,self.backtest_table.columns.get_loc('issued_cw_delta')] = object1.delta
            if self.backtest_table.loc[row,'selected_cw']!='No CW Available':
                K_ = find_K_(self.backtest_table.loc[row,'selected_cw'])
                T_ = find_T_(self.backtest_table.loc[row,'selected_cw'])
                k_ = find_k_(self.backtest_table.loc[row,'selected_cw'])
                result_set = object1.hedge(K_,T_,k_,self.risk_by_underlying).squeeze(axis=0)
                object2 = issuance(self.ticker,K_,T_,t,k_,self.n,self.r,True)
                self.backtest_table.iloc[row,self.backtest_table.columns.get_loc('market_cw_delta')] = object2.delta
            else:
                # select any K,T,k -> arbitrarily select self.K,self.T,self.k
                K_ = self.K
                T_ = self.T
                k_ = self.k
                result_set = object1.hedge(K_,T_,k_,1.0).squeeze(
                    axis=0)  # No CW Available -> 100% into underlying stock
                self.backtest_table.iloc[row,self.backtest_table.columns.get_loc('market_cw_delta')] = 0
            self.backtest_table.loc[row,['n_stock','n_hedging_cw']] = result_set

            # pct_position column
            theoritical_position = self.n*self.backtest_table.loc[row,'issued_cw_delta']
            quantity1 = self.backtest_table.loc[row,'n_stock']*self.k
            quantity2 = self.backtest_table.loc[row,'n_hedging_cw']/k_
            quantity3 = self.backtest_table.loc[row,'market_cw_delta']
            actual_position = quantity1+quantity2*quantity3
            self.backtest_table.loc[row,'pct_position'] = actual_position/theoritical_position

        self.backtest_table['pct_position'] = self.backtest_table['pct_position'].astype(np.float64)
        self.backtest_table.dropna(subset=['pct_position'],inplace=True)
        self.backtest_table.reset_index(drop=True,inplace=True)
        columns = self.backtest_table.columns

        # delta_n_ticker column
        self.backtest_table['delta_n_stock'] = self.backtest_table['n_stock'].diff()
        self.backtest_table.iloc[0,columns.get_loc('delta_n_stock')] \
            = self.backtest_table.iloc[0,columns.get_loc('n_stock')]

        # delta_n_cw_ticker column
        self.backtest_table.iloc[0,columns.get_loc('delta_n_hedging_cw')] \
            = self.backtest_table.iloc[0,columns.get_loc('n_hedging_cw')]
        for row in range(1,self.backtest_table.shape[0]):
            current_cw = self.backtest_table.iloc[row,columns.get_loc('selected_cw')]
            last_cw = self.backtest_table.iloc[row-1,columns.get_loc('selected_cw')]
            if current_cw==last_cw:
                current_n_cw = self.backtest_table.iloc[row,columns.get_loc('n_hedging_cw')]
                last_n_cw = self.backtest_table.iloc[row-1,columns.get_loc('n_hedging_cw')]
            else:
                current_n_cw = self.backtest_table.iloc[row,columns.get_loc('n_hedging_cw')]
                last_n_cw = 0
            self.backtest_table.iloc[row,columns.get_loc('delta_n_hedging_cw')] = current_n_cw-last_n_cw

        # p_stock column
        price_stock = pd.read_csv(join(dirname(realpath(__file__)),'backtest_data','StockPrice.csv'))
        price_stock = price_stock.loc[price_stock['ticker']==self.ticker,['trading_date','close']].set_index(
            'trading_date').squeeze()
        price_stock.index = price_stock.index.map(lambda x:dt.date(int(x[:4]),int(x[5:7]),int(x[-2:])))
        self.backtest_table['p_stock'] = self.backtest_table['date'].map(price_stock)
        self.backtest_table['p_stock'].fillna(method='ffill',inplace=True)  # in case of missing data

        # dividend column
        self.backtest_table.loc[0,'dividend'] = 0  # chac chan bang 0
        for row in range(1,self.backtest_table.shape[0]):
            cur_price = self.backtest_table.iloc[row,columns.get_loc('p_stock')]
            pre_price = self.backtest_table.iloc[row-1,columns.get_loc('p_stock')]
            trade_date = self.backtest_table.iloc[row,columns.get_loc('date')]
            if trade_date in self.dividend_date and cur_price<pre_price:
                self.backtest_table.loc[row,'dividend'] = pre_price-cur_price
            else:
                self.backtest_table.loc[row,'dividend'] = 0

        # p_issued_cw column
        price_issued_cw = self.price_table[self.case]
        self.backtest_table['p_issued_cw'] = self.backtest_table['date'].map(price_issued_cw)
        self.backtest_table['p_issued_cw'].fillna(method='ffill',inplace=True)  # in case of missing data

        # p_hedging_cw column
        selected_cws = self.backtest_table['selected_cw'].drop_duplicates()
        for cw in selected_cws:
            if cw=='No CW Available':
                price_mapper = pd.Series(0,index=self.backtest_table['date'])
            else:
                price_mapper = self.price_table[cw]
            d_range = self.backtest_table['selected_cw']==cw
            self.backtest_table.loc[d_range,'p_hedging_cw'] = self.backtest_table['date'].map(price_mapper)
            self.backtest_table.loc[d_range,'p_hedging_cw'] = self.backtest_table.loc[d_range,'p_hedging_cw'].fillna(
                method='ffill')  # in case of missing data
            # Note: fillna(inplace=True) does not work with list indexing (pandas's limitation)

        # cash_flow column
        for row in self.backtest_table.index:
            from_dividend = self.backtest_table.loc[row,'dividend']
            if row==0:
                from_initial_issuance = self.n*self.P0
                from_stock = - self.backtest_table.loc[row,'n_stock']*self.backtest_table.loc[row,'p_stock']
                from_cw = - self.backtest_table.loc[row,'n_hedging_cw']*self.backtest_table.loc[row,'p_hedging_cw']
                self.backtest_table.loc[row,'cash_flow'] = from_initial_issuance+from_stock+from_cw+from_dividend
            else:
                previous_cw = self.backtest_table.loc[row-1,'selected_cw']
                current_cw = self.backtest_table.loc[row,'selected_cw']
                if current_cw==previous_cw:
                    from_due_cw = 0
                else:
                    from_due_cw = self.backtest_table.loc[row-1,'n_hedging_cw']*self.backtest_table.loc[
                        row-1,'p_hedging_cw']
                if row!=self.backtest_table.index[-1]:
                    from_remaining_stock = 0
                    from_remaining_cw = 0
                    from_issued_cw = 0
                else:
                    from_remaining_stock = self.backtest_table.loc[row,'n_stock']*self.backtest_table.loc[row,'p_stock']
                    from_remaining_cw = self.backtest_table.loc[row,'n_hedging_cw']*self.backtest_table.loc[
                        row,'p_hedging_cw']
                    from_issued_cw = - self.n*self.backtest_table.loc[row,'p_issued_cw']
                from_delta_stock = - self.backtest_table.loc[row,'delta_n_stock']*self.backtest_table.loc[row,'p_stock']
                from_delta_cw = - self.backtest_table.loc[row,'delta_n_hedging_cw']*self.backtest_table.loc[
                    row,'p_hedging_cw']
                self.backtest_table.loc[row,'cash_flow'] = from_delta_stock+from_delta_cw+from_remaining_stock \
                                                           +from_remaining_cw+from_issued_cw+from_due_cw+from_dividend

        # monthly MIRR calculation with reinvest rate, financing rate = 0
        cf = self.backtest_table['cash_flow']
        self.MIRR = computeMIRR(cf)

        # determine maturity status
        col = self.backtest_table.columns.get_loc('p_issued_cw')
        if self.backtest_table.iloc[-1,col]>0:
            self.status = 'ITM'
        else:
            self.status = 'OTM'

        # GRAPHING
        fig,ax = plt.subplots(figsize=(15,12))
        fig.tight_layout()
        fig.subplots_adjust(top=0.95,bottom=0.08,left=0.1,right=0.99)
        backtest_table_for_graph = self.backtest_table[['date','pct_position','cash_flow']].copy()
        backtest_table_for_graph['pct_position'] *= 100
        sns.lineplot(data=backtest_table_for_graph,x='date',y='pct_position',ax=ax,lw=2,color='blue')
        ax.axhline(100,ls='--',linewidth=1,color='black')
        ax.axhline(99,ls='--',linewidth=1,color='red')
        ax.axhline(101,ls='--',linewidth=1,color='red')
        ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(decimals=1))
        ax.set_ylim(ymin=98.5,ymax=101.5)
        ax.yaxis.set_ticks(np.arange(98.5,102,0.5))
        ax.set_ylabel('Risk Position Balance',fontsize=18,labelpad=8)
        ax.tick_params(axis='y',which='major',labelsize=14)
        ax.yaxis.offsetText.set_visible(False)
        ax.grid(axis='y',alpha=0.2)
        ax.set_xlabel('Date',fontsize=16,labelpad=14)
        ax.xaxis.set_major_formatter(DateFormatter('%d-%m-%Y'))
        ax.tick_params(axis='x',which='major',labelsize=14)
        fmt_month = mpl.dates.MonthLocator()
        ax.xaxis.set_minor_locator(fmt_month)
        ax.grid(axis='y',alpha=0.2)

        ax.annotate(f'{self.case}',
                    xy=(0.01,1.02),xycoords='axes fraction',
                    ha='left',va='bottom',fontsize=18,fontweight='bold',color='blue')
        ax.annotate(f'Maturity Status: {self.status}  |',
                    xy=(0.77,1.02),xycoords='axes fraction',
                    ha='right',va='bottom',fontsize=18,color='black')
        if self.MIRR<0:
            color = 'red'
        else:
            color = 'green'
        ax.annotate(f'Monthly MIRR = {np.round(self.MIRR*100,2)}%',
                    xy=(1,1.02),xycoords='axes fraction',
                    ha='right',va='bottom',fontsize=18,color=color)

        fig.savefig(join(dirname(realpath(__file__)),'backtest_result',f'{self.case}.png'))

        self.backtest_table.rename({
            'date':'Date',
            'n_issued_cw':'Number of Issued CWs',
            'selected_cw':'Selected Hedging CWs',
            'days_left':'Days to Maturity of Hedging CWs',
            'n_stock':'Number of Stocks',
            'n_hedging_cw':'Number of Hedging CWs',
            'delta_n_stock':'Increase/Decrease of Stocks',
            'delta_n_hedging_cw':'Increase/Decrease of Hedging CWs',
            'p_issued_cw':'Price of Issued CW',
            'p_stock':'Price of Stock',
            'p_hedging_cw':'Price of Hedging CWs',
            'dividend':'Dividend Received',
            'issued_cw_delta':'Risk of Issued CWs',
            'market_cw_delta':'Risk of Hedging CWs',
            'pct_position':'Risk Position Balance',
            'cash_flow':'Cash Flow'
        },axis=1,inplace=True)

        self.backtest_table.to_excel(join(dirname(realpath(__file__)),'backtest_result',f'{self.case}.xlsx'))

        return self.backtest_table


class multiple_backtest:

    @staticmethod
    def run(
        selected_fold,
        folds=10
    ):
        data_date = pd.read_csv(join(dirname(realpath(__file__)),'backtest_data','PriceTable.csv'),
                                usecols=['<Date>']).squeeze()
        last_date_str = str(data_date.max())
        converter = lambda x:dt.date(int(x[:4]),int(x[4:6]),int(x[6:8]))
        last_date = converter(last_date_str)

        cw_list = pd.read_csv(join(dirname(realpath(__file__)),'backtest_data','InfoTable.csv'),usecols=['cw','T'])
        cw_list = cw_list.loc[cw_list['T'].map(lambda x:dt.date(int(x[-4:]),int(x[3:5]),int(x[:2]))).lt(last_date)]
        cw_list = cw_list.drop('T',axis=1).squeeze()
        cw_list = cw_list.sort_values().reset_index(drop=True)

        size = int(len(cw_list)/folds)+1
        selected_cw = cw_list.iloc[size*(selected_fold-1):size*selected_fold]  # last fold might have size-1 CWs

        for cw in selected_cw:
            print(f'{cw} ::: Back Testing...')
            try:
                backtest(cw,1000000,rf,0.95).run()
                print(f'{cw} ::: Finished')
            except (Exception,):
                try:
                    print(f'{case} cannot run with risk_by_underlying = 0.95, run with 1 instead')
                    backtest(cw,1000000,rf,1).run()
                    print(f'{cw} ::: Finished')
                except (Exception,):
                    print(f'{cw} ::: Backtest Failed')
                    continue
            print('===============')

    @staticmethod
    def rerun(
        risk_by_underlying=0.8
    ):
        folder = join(dirname(realpath(__file__)),'backtest_result')
        file_names = [name for name in listdir(folder) if name.endswith('.xlsx') and name.startswith('C')]
        for file in file_names:
            balance = pd.read_excel(join(folder,file),usecols=['Risk Position Balance']).squeeze()
            balance = (balance<0.8)|(balance>1.2)
            if balance.sum()>1:
                cw = file.split('.')[0]
                print(f'{cw} ::: Back Testing...')
                try:
                    backtest(cw,1000000,rf,risk_by_underlying).run()
                    print(f'{cw} ::: Finished')
                except (Exception,):
                    print(f'{cw} ::: Backtest Failed')
                    continue

    @staticmethod
    def report(
        bmk_balance=0.2,
        bmk_mirr=-0.01
    ):

        folder = join(dirname(realpath(__file__)),'backtest_result')
        file_names = [name for name in listdir(folder) if name.endswith('.xlsx') and name.startswith('C')]
        cw_list = [name.split('.')[0] for name in file_names]

        report_table = pd.DataFrame(index=cw_list,columns=['Final Result','Avg. Position Balance','MIRR'])

        def f(cw):
            result_table = pd.read_excel(join(folder,cw+'.xlsx'),
                                         usecols=['Risk Position Balance','Cash Flow'])
            check_balance = ((result_table['Risk Position Balance']-1).abs()<bmk_balance).all()
            cf = result_table['Cash Flow']
            monthly_mirr = computeMIRR(cf)
            check_mirr = monthly_mirr>=bmk_mirr
            check_both = check_balance and check_mirr
            if check_both:
                final_result = 'Passed'
            else:
                final_result = 'Failed'
            avg_balance = result_table['Risk Position Balance'].median()
            mirr = monthly_mirr
            return final_result,avg_balance,mirr

        for cw in cw_list:
            try:
                report_table.loc[cw] = f(cw)
            except (Exception,):
                continue

        report_table.to_excel(join(dirname(folder),'FinalReport.xlsx'))

        return report_table


class forwardtest:

    def __init__(self):
        self.all_stock = scrape_cw_stock.run(True)

    def filter_stock(
        self,
        c: int,
        score_ratio: str,
    ) \
            -> pd.DataFrame:

        """
        :param c: Total capital for CW issuance
        :param score_ratio: weights of liquidity score and competition score.
        Ex: '50:50', '40:60'
        """
        filter_date = dt.datetime.now().strftime('%Y-%m-%d')
        last_working_date = bdate(filter_date,-1)
        stock_data = pd.read_sql(
            f"""
            EXEC [dbo].[spStockIntraday] 
            @FROM_DATE = N'{last_working_date}'
            @TO_DATE = N'{last_working_date}',
            """,
            connect_RMD,
            index_col='SYMBOL',
        )
        cw_data_hose = pd.read_sql(
            f"""
            EXEC [dbo].[spCoveredWarrantIntraday]
            @FROM_DATE = N'{last_working_date}',
            @TO_DATE = N'{last_working_date}'
            """,
            connect_RMD,
            index_col='SYMBOL',
        )
        cw_data_vstock = scrape_missing_info.run()
        cw_data = pd.concat([cw_data_hose,cw_data_vstock],axis=1)
        cw_data = cw_data[[
            'UNDERLYING_SYMBOL',
            'ISSUER_NAME',
            'LISTED_SHARE',
            'TRADING_DATE',
            'ISSUANCE_DATE',
            'MATURITY_DATE',
            'EXERCISE_RATIO',
            'ISSUANCE_PRICE',
            'EXERCISE_PRICE',
        ]]
        cw_data.dropna(how='any',inplace=True)
        cw_data['TRADING_DATE'] = cw_data['TRADING_DATE'].str.split().str.get(0)
        cw_data[['ISSUANCE_DATE','MATURITY_DATE']] = cw_data[['ISSUANCE_DATE','MATURITY_DATE']].applymap(
            lambda x:f"{x.split('/')[2]}-{x.split('/')[1]}-{x.split('/')[0]}"
        )
        cw_data['EXERCISE_RATIO'] = cw_data['EXERCISE_RATIO'].str.split(':').str.get(0).astype(np.float64)
        cw_data['EXERCISE_PRICE'] *= 0.1
        cw_data['CW_VALUATION'] = np.nan
        for row in range(cw_data.shape[0]):
            stock = cw_data.iloc[row,cw_data.columns.get_loc('UNDERLYING_SYMBOL')]
            K = cw_data.iloc[row,cw_data.columns.get_loc('EXERCISE_PRICE')]
            T = cw_data.iloc[row,cw_data.columns.get_loc('MATURITY_DATE')]
            t = cw_data.iloc[row,cw_data.columns.get_loc('ISSUANCE_DATE')]
            k = cw_data.iloc[row,cw_data.columns.get_loc('EXERCISE_RATIO')]
            n = 1
            r = rf
            val = issuance(stock,K,T,t,k,n,r).valuation()
            cw_data.iloc[row,cw_data.columns.get_loc('CW_VALUATION')] = val

        cw_data['DISCOUNT_RATE'] = cw_data['ISSUANCE_PRICE']/cw_data['CW_VALUATION']-1
        discount_rate = cw_data.groupby('UNDERLYING_SYMBOL')['DISCOUNT_RATE'].mean()

        result = pd.DataFrame(index=self.all_stock)
        result['n_stock'] = c/(result.index.map(ta.last)*1000)
        result['foreign_room'] = stock_data['ROOM']
        # Foreign Room Condition
        result.loc[result['n_stock']<result['foreign_room'],'room_condition'] = 'Passed'
        result.loc[result['n_stock']>=result['foreign_room'],'room_condition'] = 'Failed'
        # 3M Average Volume
        f = lambda x:ta.hist(x).tail(66)['total_volume'].mean()
        result['3m_avg_volume'] = result.index.map(f)
        # Liquidity Percentage
        result.loc[result['room_condition']=='Passed','liquidity_pct'] = result['n_stock']/result['3m_avg_volume']
        # Discount Rate
        result.loc[result['room_condition']=='Passed','discount_rate'] = discount_rate
        result['discount_rate'].fillna('No Issuer',inplace=True)
        # Liquidity Score
        result['liquidity_score'] = result.loc[result['liquidity_pct']!='Failed','liquidity_pct'].astype(
            np.float64).rank(ascending=False)
        # Competition Score
        discount_rate = discount_rate.reindex(result.index)
        result['competition_score'] = discount_rate.loc[result['room_condition']=='Passed'].rank()
        result['competition_score'].fillna(result['competition_score'].max()+1,inplace=True)
        # Score Ratio
        result.loc[result['room_condition']=='Passed','score_ratio'] = score_ratio
        # Final Score
        liquidity_weight = int(score_ratio.split(':')[0])/100
        competition_weight = int(score_ratio.split(':')[1])/100
        liquidity_component = result.loc[result['room_condition']=='Passed','liquidity_score']*liquidity_weight
        competition_component = result.loc[result['room_condition']=='Passed','competition_score']*competition_weight
        result['final_score'] = liquidity_component+competition_component

        result.sort_values('final_score',ascending=False,inplace=True)
        result.fillna('Failed',inplace=True)

        now = dt.datetime.now()
        result.to_excel(f'stock_cw_ranking_{np.round(c/1e9,2)}B_{now.year}{now.month}{now.day}.xlsx')

        return result

    def run(self):

        ranking = pd.read_excel(
            r"C:\Users\hiepdang\PycharmProjects\DataAnalytics\covered_warrant\stock_cw_ranking_1.0B_20211011.xlsx",
            index_col=0
        )
        exercise_price_table = pd.DataFrame(index=ranking.index,columns=np.round(np.arange(0,0.21,0.05),2))
        exercise_price_table[0] = exercise_price_table.index.map(
            lambda x:ta.hist(x,'2021-10-11','2021-10-11')['close'].squeeze()*1000)
        for p_level in exercise_price_table.columns[1:]:
            exercise_price_table[p_level] = exercise_price_table[0]*(1+p_level)
        mirr_table = pd.DataFrame(index=exercise_price_table.index,columns=exercise_price_table.columns)

        def f(ticker,K):
            price = ta.hist(ticker)[['trading_date','close']].set_index('trading_date')*1000
            test_table = pd.DataFrame(
                index=price.loc['2021-10-11':'2021-11-30'].index,
                columns=[
                    'n_issued_cw',
                    'n_stock',
                    'delta_n_stock',
                    'p_issued_cw',
                    'p_stock',
                    'issued_cw_delta',
                    'pct_position',
                    'cash_flow',
                ]
            )
            n = 1000000
            k = 2  # k = 2
            test_table['n_issued_cw'] = n
            for t in test_table.index:
                issuance_obj = issuance(ticker,K,'2021-11-30',t,k,n,rf,False)
                # Column: n_stock
                n_stock = issuance_obj.hedge(1000000,'2021-11-30',1,1)[
                    'n_stock'].squeeze()  # first 3 arguments are dummy when risk_by_underlying = 1
                n_stock = np.round(n_stock,-2)
                test_table.loc[t,'n_stock'] = n_stock
                # Column: p_issued_cw
                val = issuance_obj.valuation()
                test_table.loc[t,'p_issued_cw'] = val
                # Column: issued_cw_delta
                delta = issuance_obj.delta
                test_table.loc[t,'issued_cw_delta'] = delta

            # Column: delta_n_stock
            test_table['delta_n_stock'] = test_table['n_stock'].diff(1)
            test_table.iloc[0,test_table.columns.get_loc('delta_n_stock')] = test_table.iloc[
                0,test_table.columns.get_loc('n_stock')]
            # Column: p_stock
            test_table['p_stock'] = price['close']
            # Column: pct_position
            theoritical_position = n*test_table['issued_cw_delta']
            actual_position = k*test_table['n_stock']
            test_table['pct_position'] = actual_position/theoritical_position
            # Column: cash_flow
            test_table['cash_flow'] = - test_table['delta_n_stock']*test_table['p_stock']
            ## at t = 0:
            issuance_amount = n*test_table.iloc[0,test_table.columns.get_loc('p_issued_cw')]
            test_table.iloc[0,test_table.columns.get_loc('cash_flow')] = issuance_amount+test_table.iloc[
                0,test_table.columns.get_loc('cash_flow')]
            ## at t = T
            cashout_amount = test_table.iloc[-1,test_table.columns.get_loc('n_stock')]*test_table.iloc[
                -1,test_table.columns.get_loc('p_stock')]
            payment_amount = - n*test_table.iloc[-1,test_table.columns.get_loc('p_issued_cw')]
            test_table.iloc[-1,test_table.columns.get_loc('cash_flow')] = payment_amount+cashout_amount+test_table.iloc[
                -1,test_table.columns.get_loc('cash_flow')]

            cf = test_table['cash_flow']
            MIRR = computeMIRR(cf)

            test_table.to_excel(join(dirname(realpath(__file__)),'forwardtest_result',f'{ticker}_{np.round(K,0)}.xlsx'))

            return MIRR

        for ticker in mirr_table.index[17:]:
            for p_level in mirr_table.columns:
                exercise_price = exercise_price_table.loc[ticker,p_level]
                try:
                    mirr_table.loc[ticker,p_level] = f(ticker,exercise_price)
                except (Exception,):
                    continue


def computeMIRR(CF_Series):
    """
    compute monthly MIRR with reinvest rate, financing rate = 0
    """

    CF_Series.sort_index(inplace=True)
    inflow = CF_Series.loc[CF_Series>=0];
    outflow = CF_Series.loc[CF_Series<0]
    FV_sum = inflow.sum();
    PV_sum = outflow.sum()
    days = len(CF_Series)
    MIRR = (FV_sum/-PV_sum)**(30/days)-1

    return MIRR
