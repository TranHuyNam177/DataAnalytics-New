from function import *


###############################################################################


class core(object):
    __driver = '{SQL Server}'
    __server = 'SRV-RPT'
    __database = 'RiskDb'
    __id = 'hiep'
    __password = '5B7Cv6huj2FcGEM4'

    __connect = pyodbc.connect(f'Driver={__driver};'
                               f'Server={__server};'
                               f'Database={__database};'
                               f'uid={__id};'
                               f'pwd={__password}')

    __deal_path = r'\\192.168.10.101\phs-storge-2018' \
                  r'\RiskManagementDept\RMD_Data' \
                  r'\Luu tru van ban\Daily Report' \
                  r'\04. High Risk & Liquidity\Group.Deal'

    __deal_files = listdir(__deal_path)

    # can't be replaced by list comprehension bc of weird behavior in Python 3
    __deal_table = pd.DataFrame()
    for __file in __deal_files:
        if not __file.startswith('~$'):
            __frame = pd.read_excel(f'{__deal_path}/{__file}')
            __deal_table = pd.concat([__deal_table,__frame])

    DealCustomer = __deal_table['Account'].str.strip().drop_duplicates().tolist()

    TableNames = pd.read_sql('SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',__connect)

    __dateconverter = lambda x:x.date() if isinstance(x,dt.datetime) else None
    Info = pd.read_sql('SELECT * FROM ONE_TIME_DATA',__connect)
    Info['DATE_OF_BIRTH'] = Info['DATE_OF_BIRTH'].map(__dateconverter)
    Info['OPENING_DATE'] = Info['OPENING_DATE'].map(__dateconverter)
    Info['FIRST_TRADING_DATE'] = Info['FIRST_TRADING_DATE'].map(__dateconverter)

    # All = Retail + Institutional = Deal + Individual = Margin + Normal
    AllCustomer = Info['TRADING_CODE'].tolist()
    RetailCustomer = Info['TRADING_CODE'].loc[Info['CUSTOMER_TYPE']=='I'].tolist()
    InstitutionalCustomer = Info['TRADING_CODE'].loc[Info['CUSTOMER_TYPE']=='B'].tolist()
    IndividualCustomer = Info['TRADING_CODE'].loc[~(Info['TRADING_CODE'].isin(DealCustomer))].tolist()

    # Customer who has margin sub-account
    MarginCustomer = Info['TRADING_CODE'].loc[Info['IS_MARGIN']].tolist()
    NormalCustomer = pd.Series(AllCustomer).loc[~pd.Series(AllCustomer).isin(MarginCustomer)].tolist()

    NAV = pd.read_sql('SELECT * FROM NET_ASSET_VALUE',
                      con=__connect,
                      index_col=['ID']) \
        .astype({'NAV':np.int64})
    NAV['TRADING_DATE'] = NAV['TRADING_DATE'].map(__dateconverter)

    time = np.nanmax(NAV['TRADING_DATE'])

    Buy = pd.read_sql('SELECT * FROM STOCKS_BUY',
                      con=__connect,
                      index_col=['ID'])
    Buy['TRADING_DATE'] = Buy['TRADING_DATE'].map(__dateconverter)

    Sell = pd.read_sql('SELECT * FROM STOCKS_SELL',
                       con=__connect,
                       index_col=['ID'])
    Sell['TRADING_DATE'] = Sell['TRADING_DATE'].map(__dateconverter)

    MarginOutstandings = pd.read_sql('SELECT * FROM MARGIN_OUTSTANDINGS',
                                     con=__connect,
                                     index_col=['ID']) \
        .astype({'TOTAL_AMT':np.int64})
    MarginOutstandings.rename(columns={'TOTAL_AMT':'MARGIN_OUTSTANDINGS'},
                              inplace=True)
    MarginOutstandings['TRADING_DATE'] \
        = MarginOutstandings['TRADING_DATE'].map(__dateconverter)

    InterestExpense = pd.read_sql('SELECT * FROM INTEREST_EXPENSE',
                                  con=__connect,
                                  index_col=['ID']) \
        .astype({'TOTAL_AMT':np.int64})
    InterestExpense['TRADING_DATE'] = InterestExpense['TRADING_DATE'].map(__dateconverter)

    InterestExpense.rename(columns={'TOTAL_AMT':'INTEREST_EXPENSE'},
                           inplace=True)

    ActiveCustomers = NAV['TRADING_CODE'].drop_duplicates().tolist()

    TradingCustomers = pd.concat([Buy['TRADING_CODE'],
                                  Sell['TRADING_CODE']]).drop_duplicates().tolist()

    def __init__(self):
        pass

    def oftypes(self,*subtype) -> list:

        """
        This method returns a list of trading codes that belong to all of given types

        :param subtype: accept one or many string, including: None (same as 'all'),
        'all', 'individual', 'deal', 'retail', 'institutional', 'margin' ,'normal'

        :return: list of trading_codes of specified types
        """

        subset = set(self.AllCustomer)

        if 'all' in subtype:
            pass
        if 'active' in subtype:
            subset = subset&set(self.ActiveCustomers)
        if 'trading' in subtype:
            subset = subset&set(self.TradingCustomers)
        if 'individual' in subtype:
            subset = subset&set(self.IndividualCustomer)
        if 'deal' in subtype:
            subset = subset&set(self.DealCustomer)
        if 'retail' in subtype:
            subset = subset&set(self.RetailCustomer)
        if 'institutional' in subtype:
            subset = subset&set(self.InstitutionalCustomer)
        if 'margin' in subtype:
            subset = subset&set(self.MarginCustomer)
        if 'normal' in subtype:
            subset = subset&set(self.NormalCustomer)

        return list(subset)

    def info(
        self,
        trading_codes: list = 'all',
        subtype: list = 'all'
    ) \
            -> pd.DataFrame:

        """
        This method returns customer personal data

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        the resulted list of trading codes is of those who belong to all these types

        :return: personal information
        """

        table = self.Info.set_index(['TRADING_CODE'])

        if trading_codes=='all' and subtype=='all':
            trading_codes = self.AllCustomer
        elif trading_codes!='all' and subtype=='all':
            pass
        elif trading_codes=='all' and subtype!='all':
            trading_codes = self.oftypes(*subtype)
        else:
            set1 = set(trading_codes)
            set2 = set(self.oftypes(*subtype))
            trading_codes = list(set1&set2)

        table = table.loc[trading_codes]

        if table.empty is True:
            return pd.DataFrame()

        # gender mapping
        table['GENDER'] = table['GENDER'].map({'001':'Male',
                                               '002':'Female',
                                               '000':'Unknown'})

        # age mapping
        today = date.today()

        def f(time):
            if isinstance(time,date):
                age = int((today-time).days/365)
                if age>100 or age<16:
                    age = None
            else:
                age = None
            return age

        age_column = table['DATE_OF_BIRTH'].map(f)
        table.insert(loc=1,column='AGE',value=age_column)

        # address mapping
        table.insert(4,'LOCATION',table['ADDRESS'].map(process_address))

        # branch mapping
        mapper = {
            'namsaigon1':'Dist.07',
            'quan3':'Dist.03',
            'hoiso':'Head Office',
            'thanhxuan':'Thanh Xuan',
            'chinhanhq7':'Dist.07',
            'internetbroker':'Internet Broker',
            'phumyhung':'PMH',
            'caugiay':'Thanh Xuan',
            'institutionalbusiness01':'INB.01',
            'hanoi':'Ha Noi',
            'crossselling':'Internet Broker',
            'tanbinh':'Tan Binh',
            'phongquanlytaikhoan-01':'AMD.01',
            'chinhanhquan1':'Dist.01',
            'haiphong':'Hai Phong',
            'institutionalbusiness02':'INB.02',
            'phongquanlytaikhoan-03':'AMD.03'
        }

        table['BRANCH'] \
            = table['BRANCH_NAME_OPEN_ACCOUNT'].map(lambda x:unidecode.unidecode(x).replace(' ','').lower())
        table['BRANCH'] = table['BRANCH'].map(mapper)
        table.drop('BRANCH_NAME_OPEN_ACCOUNT',axis=1,inplace=True)

        # account type mapping
        account_type_series = pd.Series('Margin',index=self.MarginCustomer) \
            .append(pd.Series('Normal',index=self.NormalCustomer))
        table.insert(loc=3,column='ACCOUNT_TYPE',value=account_type_series)

        # customer type mapping
        mapper = {'I':'Retail',
                  'B':'Institutional'}
        table['CUSTOMER_TYPE'] = table['CUSTOMER_TYPE'].map(mapper)

        table = table.reindex(['DATE_OF_BIRTH','AGE','GENDER','ACCOUNT_TYPE',
                               'CUSTOMER_TYPE','ADDRESS','LOCATION','OPENING_DATE',
                               'FIRST_TRADING_DATE','BRANCH'],axis=1)

        return table

    def ofgroups(
        self,
        locations: list = 'all',
        genders: list = 'all',
        ages: list = 'all',
        subtype: list = 'all'
    ) \
            -> list:

        """
        This method returns a list of trading codes that belong to all of given location, gender, age, and subtype

        :param locations: list of locations or 'all'
        :param genders: list of genders or 'all'
        :param ages: list of ages or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']

        :return: list of trading_codes of specified types
        """

        table = self.info('all',subtype)

        if locations=='all':
            index_of_loc = table.index
        else:
            index_of_loc = table.loc[table['LOCATION'].isin(locations)].index

        if genders=='all':
            index_of_gender = table.index
        else:
            index_of_gender = table.loc[table['GENDER'].isin(genders)].index

        if ages=='all':
            index_of_age = table.index
        else:
            index_of_age = table.loc[table['AGE'].isin(ages)].index

        result_index = index_of_loc.intersection(index_of_gender).intersection(index_of_age)
        result = result_index.tolist()

        return result

    def nav(
        self,
        trading_codes: list = 'all',
        trading_date: str = 'all',
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns historical NAV of given customers

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param trading_date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd' or 'now' or 'all'
        :param subtype: allow any combination of 'all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal'
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: NAV by customer through time
        """

        section = self.NAV.copy()

        if trading_codes=='all':
            set1 = set(self.AllCustomer)
        else:
            set1 = set(trading_codes)

        if 'locations' not in demography:
            demography['locations'] = 'all'
        if 'genders' not in demography:
            demography['genders'] = 'all'
        if 'ages' not in demography:
            demography['ages'] = 'all'

        set2 = set(self.ofgroups(locations=demography['locations'],
                                 genders=demography['genders'],
                                 ages=demography['ages'],
                                 subtype=subtype))

        trading_codes = list(set1&set2)

        section = section.loc[section['TRADING_CODE'].isin(trading_codes)]
        section.set_index(['TRADING_DATE','TRADING_CODE'],drop=True,inplace=True)

        if trading_date=='now':
            section = section.loc[pd.IndexSlice[self.time,:]]
        elif trading_date=='all':
            pass
        else:
            year = trading_date[:4]
            month = trading_date[5:7]
            day = trading_date[-2:]
            section = section.loc[[dt.date(int(year),int(month),int(day))]]

        hist_nav = section.sort_index()

        return hist_nav

    def buy(
        self,
        trading_codes: list = 'all',
        trading_date: str = 'all',
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns buying history of given customers

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param trading_date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd' or 'now' or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: tickers, buying price, buying volume
        """

        table = self.Buy.copy()

        if trading_codes=='all':
            set1 = set(self.AllCustomer)
        else:
            set1 = set(trading_codes)

        if 'locations' not in demography:
            demography['locations'] = 'all'
        if 'genders' not in demography:
            demography['genders'] = 'all'
        if 'ages' not in demography:
            demography['ages'] = 'all'

        set2 = set(self.ofgroups(locations=demography['locations'],
                                 genders=demography['genders'],
                                 ages=demography['ages'],
                                 subtype=subtype))

        trading_codes = list(set1&set2)

        table = table.loc[table['TRADING_CODE'].isin(trading_codes)]
        table.set_index(['TRADING_DATE','TRADING_CODE'],drop=True,inplace=True)

        if trading_date=='now':
            table = table.loc[pd.IndexSlice[self.time,:]]
        elif trading_date=='all':
            pass
        else:
            year = trading_date[:4]
            month = trading_date[5:7]
            day = trading_date[-2:]
            table = table.loc[[pd.IndexSlice[dt.date(int(year),int(month),int(day)),:]]]

        table = table.sort_index()

        transform_function = lambda x:x.str.split(',')
        table = table.transform(transform_function,axis=1)

        def f(x):
            x = pd.Series(x)
            x = x.map(int)
            return x

        result = pd.DataFrame()
        for date_code in table.index:
            ticker_list = table.loc[date_code,'SYMBOLS']
            price_list = table.loc[date_code,'PRICES']
            volume_list = table.loc[date_code,'VOLUMES']
            fee_list = table.loc[date_code,'TRADING_FEES']
            for transaction in zip(ticker_list,
                                   f(price_list),
                                   f(volume_list),
                                   f(fee_list)):
                idx = pd.MultiIndex.from_tuples([date_code],
                                                names=['TRADING_DATE',
                                                       'TRADING_CODE'])
                frame = pd.DataFrame([transaction],
                                     index=idx,
                                     columns=['SYMBOLS','PRICES',
                                              'VOLUMES','TRADING_FEES'])
                result = pd.concat([result,frame])

        if result.empty is False:
            result = result.convert_dtypes()

        return result

    def sell(
        self,
        trading_codes: list = 'all',
        trading_date: str = 'all',
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns selling history of given customers

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param trading_date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd' or 'now' or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: tickers, selling price, selling volume
        """

        table = self.Sell.copy()

        if trading_codes=='all':
            set1 = set(self.AllCustomer)
        else:
            set1 = set(trading_codes)

        if 'locations' not in demography:
            demography['locations'] = 'all'
        if 'genders' not in demography:
            demography['genders'] = 'all'
        if 'ages' not in demography:
            demography['ages'] = 'all'

        set2 = set(self.ofgroups(locations=demography['locations'],
                                 genders=demography['genders'],
                                 ages=demography['ages'],
                                 subtype=subtype))

        trading_codes = list(set1&set2)

        table = table.loc[table['TRADING_CODE'].isin(trading_codes)]

        if table.empty is True:
            return pd.DataFrame(columns=['SYMBOLS','PRICES',
                                         'VOLUMES','TRADING_FEES'])

        table.set_index(['TRADING_DATE','TRADING_CODE'],drop=True,inplace=True)

        if trading_date=='now':
            table = table.loc[pd.IndexSlice[self.time,:]]
        elif trading_date=='all':
            pass
        else:
            year = trading_date[:4]
            month = trading_date[5:7]
            day = trading_date[-2:]
            table = table.loc[[pd.IndexSlice[dt.date(int(year),int(month),int(day)),:]]]

        table = table.sort_index()

        transform_function = lambda x:x.str.split(',')
        table = table.transform(transform_function,axis=1)

        def f(x):
            x = pd.Series(x)
            x = x.map(int)
            return x

        result = pd.DataFrame()
        for date_code in table.index:
            ticker_list = table.loc[date_code,'SYMBOLS']
            price_list = table.loc[date_code,'PRICES']
            volume_list = table.loc[date_code,'VOLUMES']
            fee_list = table.loc[date_code,'TRADING_FEES']
            for transaction in zip(ticker_list,
                                   f(price_list),
                                   f(volume_list),
                                   f(fee_list)):
                idx = pd.MultiIndex.from_tuples([date_code],
                                                names=['TRADING_DATE',
                                                       'TRADING_CODE'])
                frame = pd.DataFrame([transaction],
                                     index=idx,
                                     columns=['SYMBOLS','PRICES',
                                              'VOLUMES','TRADING_FEES'])
                result = pd.concat([result,frame])

        if result.empty is False:
            result = result.convert_dtypes()

        return result

    def margin(
        self,
        trading_codes: list = 'all',
        trading_date: str = 'all',
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns historical NAV of given customers

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param trading_date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd' or 'now' or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: margin outstanding by customer through time
        """

        section = self.MarginOutstandings.copy()

        if trading_codes=='all':
            set1 = set(self.AllCustomer)
        else:
            set1 = set(trading_codes)

        if 'locations' not in demography:
            demography['locations'] = 'all'
        if 'genders' not in demography:
            demography['genders'] = 'all'
        if 'ages' not in demography:
            demography['ages'] = 'all'

        set2 = set(self.ofgroups(locations=demography['locations'],
                                 genders=demography['genders'],
                                 ages=demography['ages'],
                                 subtype=subtype))

        trading_codes = list(set1&set2)

        section = section.loc[section['TRADING_CODE'].isin(trading_codes)]
        section.set_index(['TRADING_DATE','TRADING_CODE'],drop=True,inplace=True)

        if trading_date=='now':
            section = section.loc[pd.IndexSlice[self.time,:]]
        elif trading_date=='all':
            pass
        else:
            year = trading_date[:4]
            month = trading_date[5:7]
            day = trading_date[-2:]
            section = section.loc[[dt.date(int(year),int(month),int(day))]]

        hist_margin = section.sort_index()

        return hist_margin

    def interest(
        self,
        trading_codes: list = 'all',
        trading_date: str = 'all',
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns incurred interest expense of given customers through time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param trading_date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd' or 'now' or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: daily interest expense incurred
        """

        section = self.InterestExpense.copy()

        if trading_codes=='all':
            set1 = set(self.AllCustomer)
        else:
            set1 = set(trading_codes)

        if 'locations' not in demography:
            demography['locations'] = 'all'
        if 'genders' not in demography:
            demography['genders'] = 'all'
        if 'ages' not in demography:
            demography['ages'] = 'all'

        set2 = set(self.ofgroups(locations=demography['locations'],
                                 genders=demography['genders'],
                                 ages=demography['ages'],
                                 subtype=subtype))

        trading_codes = list(set1&set2)

        section = section.loc[section['TRADING_CODE'].isin(trading_codes)]
        section.set_index(['TRADING_DATE','TRADING_CODE'],drop=True,inplace=True)

        if trading_date=='now':
            section = section.loc[pd.IndexSlice[self.time,:]]
        elif trading_date=='all':
            pass
        else:
            year = trading_date[:4]
            month = trading_date[5:7]
            day = trading_date[-2:]
            section = section.loc[[dt.date(int(year),int(month),int(day))]]

        hist_interest = section.sort_index()
        return hist_interest

    def fee(
        self,
        trading_codes: list = 'all',
        trading_date: str = 'all',
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns total trading fee of given customers through time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param trading_date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd' or 'now' or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: trading fee
        """

        if trading_codes=='all':
            set1 = set(self.AllCustomer)
        else:
            set1 = set(trading_codes)

        if 'locations' not in demography:
            demography['locations'] = 'all'
        if 'genders' not in demography:
            demography['genders'] = 'all'
        if 'ages' not in demography:
            demography['ages'] = 'all'

        set2 = set(self.ofgroups(locations=demography['locations'],
                                 genders=demography['genders'],
                                 ages=demography['ages'],
                                 subtype=subtype))

        trading_codes = list(set1&set2)

        buy_fee = self.Buy.set_index(['TRADING_DATE','TRADING_CODE'],drop=True)
        sell_fee = self.Sell.set_index(['TRADING_DATE','TRADING_CODE'],drop=True)

        buy_fee = buy_fee.loc[buy_fee.index.get_level_values(1).isin(trading_codes)]
        sell_fee = sell_fee.loc[sell_fee.index.get_level_values(1).isin(trading_codes)]

        if trading_date=='now':
            buy_fee = buy_fee.loc[pd.IndexSlice[self.time,:]]
            sell_fee = sell_fee.loc[pd.IndexSlice[self.time,:]]
        elif trading_date=='all':
            pass
        else:
            year = trading_date[:4]
            month = trading_date[5:7]
            day = trading_date[-2:]
            buy_fee = buy_fee.loc[[dt.date(int(year),int(month),int(day))]]
            sell_fee = sell_fee.loc[[dt.date(int(year),int(month),int(day))]]

        f = lambda anylist:np.sum([int(elem) for elem in anylist],dtype='int64')
        buy_fee = buy_fee['TRADING_FEES'].str.split(',').map(f)
        sell_fee = sell_fee['TRADING_FEES'].str.split(',').map(f)

        fee_table = buy_fee.add(sell_fee,fill_value=0)
        fee_table = fee_table.groupby(level=[0,1]).sum()
        fee_table = pd.DataFrame(fee_table).astype('int64')

        return fee_table

    def value(
        self,
        trading_codes: list = 'all',
        trading_date: str = 'all',
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns total trading value (either buy or sell) of given customers through time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param trading_date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd' or 'now' or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: trading value
        """

        if trading_codes=='all':
            set1 = set(self.AllCustomer)
        else:
            set1 = set(trading_codes)

        if 'locations' not in demography:
            demography['locations'] = 'all'
        if 'genders' not in demography:
            demography['genders'] = 'all'
        if 'ages' not in demography:
            demography['ages'] = 'all'

        set2 = set(self.ofgroups(locations=demography['locations'],
                                 genders=demography['genders'],
                                 ages=demography['ages'],
                                 subtype=subtype))

        trading_codes = list(set1&set2)

        buy_value = self.Buy.set_index(['TRADING_DATE','TRADING_CODE'],drop=True)
        sell_value = self.Sell.set_index(['TRADING_DATE','TRADING_CODE'],drop=True)

        buy_value = buy_value.loc[buy_value.index.get_level_values(1).isin(trading_codes)]
        sell_value = sell_value.loc[sell_value.index.get_level_values(1).isin(trading_codes)]

        if trading_date=='now':
            buy_value = buy_value.loc[pd.IndexSlice[self.time,:]]
            sell_value = sell_value.loc[pd.IndexSlice[self.time,:]]
        elif trading_date=='all':
            pass
        else:
            year = trading_date[:4]
            month = trading_date[5:7]
            day = trading_date[-2:]
            buy_value = buy_value.loc[[dt.date(int(year),int(month),int(day))]]
            sell_value = sell_value.loc[[dt.date(int(year),int(month),int(day))]]

        def convert(x):
            int_list = [int(elem) for elem in x]
            result = np.array(int_list)
            return result

        def f(df):
            df['PRICES'] = df['PRICES'].str.split(',')
            df['PRICES'] = df['PRICES'].map(convert)
            df['VOLUMES'] = df['VOLUMES'].str.split(',')
            df['VOLUMES'] = df['VOLUMES'].map(convert)
            df['TRADING_VALUE'] = df['VOLUMES']*df['PRICES']
            df['TRADING_VALUE'] = df['TRADING_VALUE'].apply(np.sum,dtype='int64')
            df.drop(['PRICES','VOLUMES'],axis=1,inplace=True)
            df = df.groupby(level=[0,1]).sum()
            return df

        df = pd.concat([buy_value,sell_value])
        if df.empty:
            value_table = pd.DataFrame(columns=['TRADING_VALUE'])
        else:
            value_table = df[['VOLUMES','PRICES']].copy()
            value_table = f(value_table)

        return value_table


###############################################################################
############################# ANALYSIS FUNCTIONS ##############################
###############################################################################


class rolling(core):

    def __int__(self):
        pass

    def avg_margin(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        trailing_days: int = 30,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns average margin of given customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param trailing_days: number of trailing days
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column DataFrame with index of trading_codes
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        n = trailing_days

        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        margin_table.reset_index(level=1,inplace=True)
        trailing_avg_margin = margin_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
        trailing_avg_margin = trailing_avg_margin.swaplevel()
        trailing_avg_margin = trailing_avg_margin.sort_index(0,['TRADING_DATE','TRADING_CODE'])

        idx = pd.IndexSlice
        result = trailing_avg_margin.loc[idx[fromdate:todate,:],:]
        result = result.groupby(level='TRADING_CODE').mean()

        result.rename({'MARGIN_OUTSTANDINGS':f'{n}D_AVG_MARGIN_OUTSTANDINGS'},
                      axis=1,inplace=True)

        result = result.astype('int64')

        return result

    def total_fee(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        trailing_days: int = 30,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns total trading fee of given customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param trailing_days: number of trailing days
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column DataFrame with index of trading_codes
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        n = trailing_days

        fee_table = self.fee(trading_codes,'all',subtype,**demography)

        if fee_table.empty is True:
            return pd.DataFrame()

        fee_table.reset_index(level=1,inplace=True)
        trailing_total_fee = fee_table.groupby('TRADING_CODE').rolling(window=f'{n}D').sum()
        trailing_total_fee = trailing_total_fee.swaplevel()
        trailing_total_fee = trailing_total_fee.sort_index(0,['TRADING_DATE','TRADING_CODE'])

        idx = pd.IndexSlice
        result = trailing_total_fee.loc[idx[fromdate:todate,:],:]
        result = result.groupby(level='TRADING_CODE').mean()

        result.rename({'TRADING_FEES':f'{n}D_AVG_TRADING_FEES'},
                      axis=1,inplace=True)

        result = result.astype('int64')

        return result

    def avg_nav(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        trailing_days: int = 30,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns average margin of given customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param trailing_days: number of trailing days
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column DataFrame with index of trading_codes
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        n = trailing_days

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        nav_table.reset_index(level=1,inplace=True)
        trailing_avg_nav = nav_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
        trailing_avg_nav = trailing_avg_nav.swaplevel()

        # handle missing data from IT
        trailing_avg_nav = trailing_avg_nav.loc[~trailing_avg_nav['NAV'].isna()]
        trailing_avg_nav.fillna(0,inplace=True)

        trailing_avg_nav = trailing_avg_nav.sort_index(0,['TRADING_DATE','TRADING_CODE'])

        idx = pd.IndexSlice
        result = trailing_avg_nav.loc[idx[fromdate:todate,:],:]
        result = result.groupby(level='TRADING_CODE').mean()

        result.rename({'NAV':f'{n}D_AVG_NAV'},axis=1,inplace=True)
        result = result.astype('int64')

        return result

    def trading_turnover(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        trailing_days: int = 30,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns a time series of trading turnover of given customers through specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param trailing_days: number of trailing days
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column DataFrame with index of trading_codes
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        n = trailing_days

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        value_table = self.value(trading_codes,'all',subtype,**demography)

        if not nav_table.empty:
            nav_table.reset_index(level=1,inplace=True)
            trailing_avg_nav = nav_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
            trailing_avg_nav = trailing_avg_nav.swaplevel()
        else:
            trailing_avg_nav = pd.DataFrame()

        if not margin_table.empty:
            margin_table.reset_index(level=1,inplace=True)
            trailing_avg_margin = margin_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
            trailing_avg_margin = trailing_avg_margin.swaplevel()
        else:
            trailing_avg_margin = pd.DataFrame()

        if not value_table.empty:
            value_table.reset_index(level=1,inplace=True)
            trailing_sum_value = value_table.groupby('TRADING_CODE').rolling(window=f'{n}D').sum()
            trailing_sum_value = trailing_sum_value.swaplevel()
        else:
            trailing_sum_value = pd.DataFrame()

        table = pd.concat([trailing_avg_nav,
                           trailing_avg_margin,
                           trailing_sum_value],
                          axis=1)

        # handle missing data from IT
        table.dropna(how='any',inplace=True)

        if table.empty:
            return pd.DataFrame()

        table.sort_index(level=[0,1],inplace=True)

        table.rename({'TRADING_VALUE':f'{n}D_TOTAL_TRADING_VALUE',
                      'NAV':f'{n}D_AVG_NAV',
                      'MARGIN_OUTSTANDINGS':f'{n}D_AVG_MARGIN_OUTSTANDINGS'},
                     axis=1,inplace=True)

        table[f'{n}D_TRADING_TURNOVER'] = table[f'{n}D_TOTAL_TRADING_VALUE'].div(
            table[f'{n}D_AVG_NAV'].add(table[f'{n}D_AVG_MARGIN_OUTSTANDINGS'],fill_value=0),fill_value=0)

        idx = pd.IndexSlice
        result = table.loc[idx[fromdate:todate,:],:]

        result = result.astype({f'{n}D_TOTAL_TRADING_VALUE':'int64',
                                f'{n}D_AVG_MARGIN_OUTSTANDINGS':'int64',
                                f'{n}D_AVG_NAV':'int64'})
        result.index = result.index.set_levels(result.index.levels[0].date,level=0)

        return result

    def margin_usage_ratio(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        trailing_days: int = 30,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns a time series of margin usage ratio of given customers through specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param trailing_days: number of trailing days
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column DataFrame with index of trading_codes
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        n = trailing_days

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        nav_table.reset_index(level=1,inplace=True)
        trailing_avg_nav = nav_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
        trailing_avg_nav = trailing_avg_nav.swaplevel()

        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        margin_table.reset_index(level=1,inplace=True)
        trailing_avg_margin = margin_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
        trailing_avg_margin = trailing_avg_margin.swaplevel()

        table = pd.concat([trailing_avg_nav,
                           trailing_avg_margin],
                          axis=1)

        # handle missing data from IT
        table = table.loc[~table['NAV'].isna()]
        if table.empty: return pd.DataFrame()
        table.fillna(0,inplace=True)

        table.sort_index(level=[0,1],inplace=True)

        table.rename({'MARGIN_OUTSTANDINGS':f'{n}D_AVG_MARGIN_OUTSTANDINGS',
                      'NAV':f'{n}D_AVG_NAV'},
                     axis=1,inplace=True)
        table[f'{n}D_TRAILING_MARGIN_USAGE_RATIO'] \
            = table[f'{n}D_AVG_MARGIN_OUTSTANDINGS'].div(
            table[f'{n}D_AVG_MARGIN_OUTSTANDINGS'].add(table[f'{n}D_AVG_NAV'],fill_value=0),fill_value=0)

        idx = pd.IndexSlice
        result = table.loc[idx[fromdate:todate,:],:]

        result = result.astype({f'{n}D_AVG_MARGIN_OUTSTANDINGS':'int64',
                                f'{n}D_AVG_NAV':'int64'})
        result.index = result.index.set_levels(result.index.levels[0].date,level=0)

        return result

    def trading_turnover_by(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        trailing_days: int = 30,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns a time series of trading turnover of given customers through specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param trailing_days: number of trailing days
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column DataFrame with index of trading_codes
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        n = trailing_days

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        value_table = self.value(trading_codes,'all',subtype,**demography)

        age_series = self.info(trading_codes,subtype)['AGE'].dropna()
        gender_series = self.info(trading_codes,subtype)['GENDER'].dropna()
        gender_series = gender_series.loc[~(gender_series=='Unknown')]

        if not nav_table.empty:
            nav_table.reset_index(level=1,inplace=True)
            trailing_avg_nav = nav_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
            trailing_avg_nav = trailing_avg_nav.swaplevel()
        else:
            trailing_avg_nav = pd.DataFrame()

        if not margin_table.empty:
            margin_table.reset_index(level=1,inplace=True)
            trailing_avg_margin = margin_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
            trailing_avg_margin = trailing_avg_margin.swaplevel()
        else:
            trailing_avg_margin = pd.DataFrame()

        if not value_table.empty:
            value_table.reset_index(level=1,inplace=True)
            trailing_sum_value = value_table.groupby('TRADING_CODE').rolling(window=f'{n}D').sum()
            trailing_sum_value = trailing_sum_value.swaplevel()
        else:
            trailing_sum_value = pd.DataFrame()

        table = pd.concat([trailing_avg_nav,
                           trailing_avg_margin,
                           trailing_sum_value],
                          axis=1)
        table.fillna(0,inplace=True)

        if table.empty:
            return pd.DataFrame()

        table.sort_index(level=[0,1],inplace=True)

        table['AGE'] = table.index.get_level_values(1).map(age_series)
        table['GENDER'] = table.index.get_level_values(1).map(gender_series)
        table = table.reindex(columns=[
            'GENDER','AGE','NAV','MARGIN_OUTSTANDINGS','TRADING_VALUE'
        ])

        # Step 1: compute sum of values by date, gender, age
        table_raw_sum = table.groupby(['TRADING_DATE','GENDER','AGE'])[
            ['NAV','MARGIN_OUTSTANDINGS','TRADING_VALUE']].sum()
        table_raw_sum.reset_index(level=['GENDER','AGE'],inplace=True)
        # Step 2a: compute rolling average of NAV by gender, age
        nav_trailing_mean = table_raw_sum.groupby(['GENDER','AGE'])['NAV'].rolling(window='30D').mean()
        # Step 2b: compute rolling average of Margin Outstandings by gender, age
        margin_trailing_mean = table_raw_sum.groupby(['GENDER','AGE'])['MARGIN_OUTSTANDINGS'].rolling(
            window='30D').mean()
        # Step 2c: compute rolling sum of trading value by gender, age. Needed at least 20 observations to calculate the sum
        value_trailing_sum = table_raw_sum.groupby(['GENDER','AGE'])['TRADING_VALUE'].rolling(window='30D').sum()

        result = pd.concat([nav_trailing_mean,margin_trailing_mean,value_trailing_sum],axis=1)
        result.reset_index(level=['GENDER','AGE'],inplace=True)

        result.rename({'NAV':f'{n}D_AVG_NAV',
                       'MARGIN_OUTSTANDINGS':f'{n}D_AVG_MARGIN_OUTSTANDINGS',
                       'TRADING_VALUE':f'{n}D_TOTAL_TRADING_VALUE'},
                      axis=1,inplace=True)
        result[f'{n}D_AVG_TRADING_TURNOVER'] \
            = result[f'{n}D_TOTAL_TRADING_VALUE'].div(
            result[f'{n}D_AVG_NAV'].add(result[f'{n}D_AVG_MARGIN_OUTSTANDINGS'],fill_value=0),fill_value=0)

        result.sort_index(inplace=True)
        result = result.loc[fromdate:todate,:]
        result.sort_values(['GENDER','AGE'],inplace=True)

        result = result.astype({f'{n}D_TOTAL_TRADING_VALUE':'int64',
                                f'{n}D_AVG_MARGIN_OUTSTANDINGS':'int64',
                                f'{n}D_AVG_NAV':'int64'})
        result.index = result.index.date

        return result

    def margin_usage_ratio_by(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        trailing_days: int = 30,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns a time series of margin usage ratio of given customers through specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param trailing_days: number of trailing days
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column DataFrame with index of trading_codes
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        n = trailing_days

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        margin_table = self.margin(trading_codes,'all',subtype,**demography)

        age_series = self.info(trading_codes,subtype)['AGE'].dropna()
        gender_series = self.info(trading_codes,subtype)['GENDER'].dropna()
        gender_series = gender_series.loc[~(gender_series=='Unknown')]

        if not nav_table.empty:
            nav_table.reset_index(level=1,inplace=True)
            trailing_avg_nav = nav_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
            trailing_avg_nav = trailing_avg_nav.swaplevel()
        else:
            trailing_avg_nav = pd.DataFrame()

        if not margin_table.empty:
            margin_table.reset_index(level=1,inplace=True)
            trailing_avg_margin = margin_table.groupby('TRADING_CODE').rolling(window=f'{n}D').mean()
            trailing_avg_margin = trailing_avg_margin.swaplevel()
        else:
            trailing_avg_margin = pd.DataFrame()

        table = pd.concat([trailing_avg_nav,
                           trailing_avg_margin],
                          axis=1)
        table.fillna(0,inplace=True)

        if table.empty:
            return pd.DataFrame()

        table.sort_index(level=[0,1],inplace=True)

        table['AGE'] = table.index.get_level_values(1).map(age_series)
        table['GENDER'] = table.index.get_level_values(1).map(gender_series)
        table = table.reindex(columns=[
            'GENDER','AGE','NAV','MARGIN_OUTSTANDINGS'
        ])

        # Step 1: compute sum of values by date, gender, age
        table_raw_sum = table.groupby(['TRADING_DATE','GENDER','AGE'])[['NAV','MARGIN_OUTSTANDINGS']].sum()
        table_raw_sum.reset_index(level=['GENDER','AGE'],inplace=True)
        # Step 2a: compute rolling average of NAV by gender, age
        nav_trailing_mean = table_raw_sum.groupby(['GENDER','AGE'])['NAV'].rolling(window='30D').mean()
        # Step 2b: compute rolling average of Margin Outstandings by gender, age
        margin_trailing_mean = table_raw_sum.groupby(['GENDER','AGE'])['MARGIN_OUTSTANDINGS'].rolling(
            window='30D').mean()

        result = pd.concat([nav_trailing_mean,margin_trailing_mean],axis=1)
        result.reset_index(level=['GENDER','AGE'],inplace=True)

        result.rename({'NAV':f'{n}D_AVG_NAV',
                       'MARGIN_OUTSTANDINGS':f'{n}D_AVG_MARGIN_OUTSTANDINGS'},
                      axis=1,inplace=True)

        result[f'{n}D_AVG_TRAILING_MARGIN_USAGE_RATIO'] \
            = result[f'{n}D_AVG_MARGIN_OUTSTANDINGS'].div(
            result[f'{n}D_AVG_NAV'].add(result[f'{n}D_AVG_MARGIN_OUTSTANDINGS'],fill_value=0),fill_value=0)

        result.sort_index(inplace=True)
        result = result.loc[fromdate:todate,:]
        result.sort_values(['GENDER','AGE'],inplace=True)

        result = result.astype({f'{n}D_AVG_MARGIN_OUTSTANDINGS':'int64',
                                f'{n}D_AVG_NAV':'int64'})
        result.index = result.index.date

        return result


class aggregation(core):

    def __int__(self):
        pass

    def on_margin(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography
    ) \
            -> int:

        """
        This method returns aggregated value of total margin of a group of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: margin outstandings
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        idx = pd.IndexSlice

        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        margin_table = margin_table.loc[idx[fromdate:todate,:],:]

        sum_margin_over_trading_codes = margin_table.groupby('TRADING_DATE').sum()
        result = sum_margin_over_trading_codes.agg(func).agg(func)

        if np.isnan(result):
            result = np.nan
        else:
            result = int(result)

        return result

    def on_fee(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> int:

        """
        This method returns aggregated trading fee of a group of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: trading fee
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        fee_table = self.fee(trading_codes,'all',subtype,**demography)

        if fee_table.empty is True:
            return pd.DataFrame()

        idx = pd.IndexSlice
        fee_table = fee_table[idx[fromdate:todate,:],:]

        sum_fee_over_trading_codes = fee_table.groupby('TRADING_DATE').sum()
        result = sum_fee_over_trading_codes.agg(func).agg(func)

        if np.isnan(result):
            result = np.nan
        else:
            result = int(result)

        return result

    def on_interest(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> int:

        """
        This method returns aggregated interest expense of a group of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: interest expense
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        interest_table = self.interest(trading_codes,'all',subtype,**demography)

        if interest_table.empty is True:
            return pd.DataFrame()

        idx = pd.IndexSlice
        interest_table = interest_table[idx[fromdate:todate,:],:]

        sum_interest_over_trading_codes = interest_table.groupby('TRADING_DATE').sum()
        result = sum_interest_over_trading_codes.agg(func).agg(func)

        if np.isnan(result):
            result = np.nan
        else:
            result = int(result)

        return result

    def on_nav(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> int:

        """
        This method returns aggregated NAV of a froup of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: NAV
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        nav_table = self.nav(trading_codes,'all',subtype,**demography)

        idx = pd.IndexSlice
        nav_table = nav_table.loc[idx[fromdate:todate,:],:]

        sum_nav_over_trading_codes = nav_table.groupby('TRADING_DATE').sum()
        result = sum_nav_over_trading_codes.agg(func).agg(func)

        if np.isnan(result):
            result = np.nan
        else:
            result = int(result)

        return result

    def on_value(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> int:

        """
        This method returns aggregated trading value of a froup of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: trading value
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        value_table = self.value(trading_codes,'all',subtype,**demography)

        idx = pd.IndexSlice
        value_table = value_table.loc[idx[fromdate:todate,:],:]

        sum_value_over_trading_codes = value_table.groupby('TRADING_DATE').sum()
        result = sum_value_over_trading_codes.agg(func).agg(func)

        if np.isnan(result):
            result = np.nan
        else:
            result = int(result)

        return result

    def on_trading_turnover(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography
    ) \
            -> float:

        """
        This method returns trading turnover of a group of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: trading turnover
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        idx = pd.IndexSlice

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        nav_table = nav_table.loc[idx[fromdate:todate,:],:]
        sum_nav_over_trading_codes = nav_table.groupby('TRADING_DATE').sum()
        result_avg_nav = sum_nav_over_trading_codes.mean().mean()

        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        margin_table = margin_table.loc[idx[fromdate:todate,:],:]
        sum_margin_over_trading_codes = margin_table.groupby('TRADING_DATE').sum()
        result_avg_margin = sum_margin_over_trading_codes.mean().mean()

        value_table = self.value(trading_codes,'all',subtype,**demography)
        value_table = value_table.loc[idx[fromdate:todate,:],:]
        sum_value_over_trading_codes = value_table.groupby('TRADING_DATE').sum()
        result_sum_value = sum_value_over_trading_codes.sum().sum()

        result = result_sum_value/(result_avg_nav+result_avg_margin)

        return result

    def on_margin_usage_ratio(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography
    ) \
            -> float:

        """
        This method returns margin usage ratio of a group of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: margin usage ratio
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        idx = pd.IndexSlice

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        nav_table = nav_table.loc[idx[fromdate:todate,:],:]
        sum_nav_over_trading_codes = nav_table.groupby('TRADING_DATE').sum()
        result_avg_nav = sum_nav_over_trading_codes.mean(axis=0).mean()

        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        margin_table = margin_table.loc[idx[fromdate:todate,:],:]
        sum_margin_over_trading_codes = margin_table.groupby('TRADING_DATE').sum()
        result_avg_margin = sum_margin_over_trading_codes.mean().mean()

        result = result_avg_margin/(result_avg_nav+result_avg_margin)

        return result

    def on_margin_by(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> pd.DataFrame:

        """
        This method returns aggregated value of margin of a group of customers over specified period of time, segmented by gender and age

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column pd.DataFrame with MultiIndex of ('LOCATION','GENDER','AGE')
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        age_series = self.info(trading_codes,subtype)['AGE'].dropna()
        gender_series = self.info(trading_codes,subtype)['GENDER'].dropna()
        gender_series = gender_series.loc[~(gender_series=='Unknown')]
        location_series = self.info(trading_codes,subtype)['LOCATION'].dropna()
        location_series = location_series.loc[~(location_series=='Unknown')]

        idx = pd.IndexSlice

        margin_table = self.margin(trading_codes,'all',subtype,**demography)
        margin_table = margin_table.loc[idx[fromdate:todate,:],:]
        margin_table.insert(0,'LOCATION',margin_table.index.get_level_values(1).map(location_series))
        margin_table.insert(0,'AGE',margin_table.index.get_level_values(1).map(age_series))
        margin_table.insert(0,'GENDER',margin_table.index.get_level_values(1).map(gender_series))
        sum_margin_over_trading_codes = margin_table.groupby(['TRADING_DATE','LOCATION','GENDER','AGE']).sum()
        result_agg_margin = sum_margin_over_trading_codes.groupby(['LOCATION','GENDER','AGE']).agg(func)
        result_agg_margin = result_agg_margin.astype('int64')

        return result_agg_margin

    def on_fee_by(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> pd.DataFrame:

        """
        This method returns aggregated trading fee of a group of customers over specified period of time, segmented by gender and age

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column pd.DataFrame with MultiIndex of ('LOCATION','GENDER','AGE')
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        age_series = self.info(trading_codes,subtype)['AGE'].dropna()
        gender_series = self.info(trading_codes,subtype)['GENDER'].dropna()
        gender_series = gender_series.loc[~(gender_series=='Unknown')]
        location_series = self.info(trading_codes,subtype)['LOCATION'].dropna()
        location_series = location_series.loc[~(location_series=='Unknown')]

        idx = pd.IndexSlice

        fee_table = self.fee(trading_codes,'all',subtype,**demography)
        fee_table = fee_table.loc[idx[fromdate:todate,:],:]
        fee_table.insert(0,'LOCATION',fee_table.index.get_level_values(1).map(location_series))
        fee_table.insert(0,'AGE',fee_table.index.get_level_values(1).map(age_series))
        fee_table.insert(0,'GENDER',fee_table.index.get_level_values(1).map(gender_series))
        sum_fee_over_trading_codes = fee_table.groupby(['TRADING_DATE','LOCATION','GENDER','AGE']).sum()
        result_agg_fee = sum_fee_over_trading_codes.groupby(['LOCATION','GENDER','AGE']).agg(func)
        result_agg_fee = result_agg_fee.astype('int64')

        return result_agg_fee

    def on_interest_by(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> pd.DataFrame:

        """
        This method returns aggregated interest expense of a group of customers over specified period of time, segmented by gender and age

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column pd.DataFrame with MultiIndex of ('LOCATION','GENDER','AGE')
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        age_series = self.info(trading_codes,subtype)['AGE'].dropna()
        gender_series = self.info(trading_codes,subtype)['GENDER'].dropna()
        gender_series = gender_series.loc[~(gender_series=='Unknown')]
        location_series = self.info(trading_codes,subtype)['LOCATION'].dropna()
        location_series = location_series.loc[~(location_series=='Unknown')]

        idx = pd.IndexSlice

        interest_table = self.interest(trading_codes,'all',subtype,**demography)
        interest_table = interest_table.loc[idx[fromdate:todate,:],:]
        interest_table.insert(0,'LOCATION',interest_table.index.get_level_values(1).map(location_series))
        interest_table.insert(0,'AGE',interest_table.index.get_level_values(1).map(age_series))
        interest_table.insert(0,'GENDER',interest_table.index.get_level_values(1).map(gender_series))
        sum_interest_over_trading_codes = interest_table.groupby(['TRADING_DATE','LOCATION','GENDER','AGE']).sum()
        result_agg_interest = sum_interest_over_trading_codes.groupby(['LOCATION','GENDER','AGE']).agg(func)
        result_agg_interest = result_agg_interest.astype('int64')

        return result_agg_interest

    def on_nav_by(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> pd.DataFrame:

        """
        This method returns aggregated NAV of a group of customers over specified period of time, segmented by gender and age

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column pd.DataFrame with MultiIndex of ('LOCATION','GENDER','AGE')
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        age_series = self.info(trading_codes,subtype)['AGE'].dropna()
        gender_series = self.info(trading_codes,subtype)['GENDER'].dropna()
        gender_series = gender_series.loc[~(gender_series=='Unknown')]
        location_series = self.info(trading_codes,subtype)['LOCATION'].dropna()
        location_series = location_series.loc[~(location_series=='Unknown')]

        idx = pd.IndexSlice

        nav_table = self.nav(trading_codes,'all',subtype,**demography)
        nav_table = nav_table.loc[idx[fromdate:todate,:],:]
        nav_table.insert(0,'LOCATION',nav_table.index.get_level_values(1).map(location_series))
        nav_table.insert(0,'AGE',nav_table.index.get_level_values(1).map(age_series))
        nav_table.insert(0,'GENDER',nav_table.index.get_level_values(1).map(gender_series))
        sum_nav_over_trading_codes = nav_table.groupby(['TRADING_DATE','LOCATION','GENDER','AGE']).sum()
        result_agg_nav = sum_nav_over_trading_codes.groupby(['LOCATION','GENDER','AGE']).agg(func)
        result_agg_nav = result_agg_nav.astype('int64')

        return result_agg_nav

    def on_value_by(
        self,
        trading_codes: list = 'all',
        func: Callable[[pd.DataFrame],int] = 'mean',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography,
    ) \
            -> pd.DataFrame:

        """
        This method returns aggregated trading value of a group of customers over specified period of time, segmented by gender and age

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param func: any aggregation function or aggregation function name, example np.sum, np.mean, 'sum', 'mean'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column pd.DataFrame with MultiIndex of ('LOCATION','GENDER','AGE')
        """

        if todate is None:
            todate = self.time
        else:
            todate = dt.date(int(todate[:4]),int(todate[5:7]),int(todate[-2:]))

        if fromdate is None:
            fromdate = dt.date(2000,1,1)
        else:
            fromdate = dt.date(int(fromdate[:4]),int(fromdate[5:7]),int(fromdate[-2:]))

        age_series = self.info(trading_codes,subtype)['AGE'].dropna()
        gender_series = self.info(trading_codes,subtype)['GENDER'].dropna()
        gender_series = gender_series.loc[~(gender_series=='Unknown')]
        location_series = self.info(trading_codes,subtype)['LOCATION'].dropna()
        location_series = location_series.loc[~(location_series=='Unknown')]

        idx = pd.IndexSlice

        value_table = self.value(trading_codes,'all',subtype,**demography)
        value_table = value_table.loc[idx[fromdate:todate,:],:]
        value_table.insert(0,'LOCATION',value_table.index.get_level_values(1).map(location_series))
        value_table.insert(0,'AGE',value_table.index.get_level_values(1).map(age_series))
        value_table.insert(0,'GENDER',value_table.index.get_level_values(1).map(gender_series))
        sum_value_over_trading_codes = value_table.groupby(['TRADING_DATE','LOCATION','GENDER','AGE']).sum()
        result_agg_value = sum_value_over_trading_codes.groupby(['LOCATION','GENDER','AGE']).agg(func)
        result_agg_value = result_agg_value.astype('int64')

        return result_agg_value

    def trading_turnover_by(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns trading turnover of a group of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param demography: take value of parameters: locations, genders, ages of function ofgroups, ignored parameters imply 'all'

        :return: 1-column pd.DataFrame with MultiIndex of ('LOCATION','GENDER','AGE')
        """

        result_avg_nav = self.on_nav_by(trading_codes,'mean',fromdate,todate,subtype,**demography).squeeze(axis=1)
        result_avg_margin = self.on_margin_by(trading_codes,'mean',fromdate,todate,subtype,**demography).squeeze(axis=1)
        result_sum_value = self.on_value_by(trading_codes,'sum',fromdate,todate,subtype,**demography).squeeze(axis=1)

        if result_avg_margin.empty:
            result_avg_margin = result_avg_nav.copy()
            result_avg_margin.iloc[:] = 0

        result = result_sum_value.div(result_avg_nav.add(result_avg_margin,fill_value=0),fill_value=0).to_frame(
            name='TRADING_TURNOVER')

        return result

    def margin_usage_ratio_by(
        self,
        trading_codes: list = 'all',
        fromdate: str = None,
        todate: str = None,
        subtype: list = 'all',
        **demography
    ) \
            -> pd.DataFrame:

        """
        This method returns margin usage ratio of a group of customers over specified period of time

        :param trading_codes: list of trading codes, example ['022C000015','022C000308'] or 'all'
        :param subtype: list of customer type ['all', 'individual', 'deal', 'retail', 'institutional', 'margin', 'normal']
        :param fromdate: example '2020-01-01', '2020/01/01', None means '2000/01/01'
        :param todate: example '2021-01-01', '2021/01/01', None means core.time

        :return: 1-column pd.DataFrame with MultiIndex of ('LOCATION','GENDER','AGE')
        """

        result_avg_nav = self.on_nav_by(trading_codes,'mean',fromdate,todate,subtype,**demography).squeeze(axis=1)
        result_avg_margin = self.on_margin_by(trading_codes,'mean',fromdate,todate,subtype,**demography).squeeze(axis=1)

        if result_avg_margin.empty:
            result_avg_margin = result_avg_nav.copy()
            result_avg_margin.iloc[:] = 0

        result = result_avg_margin.div(result_avg_nav.add(result_avg_margin,fill_value=0),fill_value=0).to_frame(
            name='MARGIN_USAGE_RATIO')

        return result


# Because we do not really want to create objects:
core = core()
rolling = rolling()
aggregation = aggregation()

###############################################################################
