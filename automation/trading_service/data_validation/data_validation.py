from automation.trading_service.data_validation import *


###############################################################################


def branch():
    DWH_branch = pd.read_sql(
        'SELECT * FROM branch',connect,index_col='branch_id'
    )
    DWH_branch.sort_index(inplace=True)
    FLEX_branch = pd.read_excel(
        join(realpath(dirname(__file__)),'flex_data','010002.xls'),
        dtype={'Mã chi nhánh':str},
        usecols=['Mã chi nhánh','Tên chi nhánh/đại lý','Trạng thái']
    )
    FLEX_branch.columns = ['branch_id','branch_name','status']
    FLEX_branch.set_index('branch_id',inplace=True)
    FLEX_branch.sort_index(inplace=True)

    check = DWH_branch.equals(FLEX_branch)
    if check:
        return True
    else:
        print("Wrong at:")
        wrong_at = DWH_branch.compare(FLEX_branch)
        return wrong_at


def broker():
    print("Table 'broker' has no validating data from FLEX")


def account():
    print("Table 'account' has no validating data from FLEX")


def sub_account():
    DWH_sub_account = pd.read_sql(
        'SELECT * FROM sub_account',
        connect
    )
    DWH_sub_account.set_index('sub_account',inplace=True)
    DWH_sub_account.sort_index(inplace=True)
    FLEX_excel = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','VCF0051.xls'),
    )
    FLEX_sub_account = pd.DataFrame()
    sheets = FLEX_excel.sheet_names
    for sheet in sheets:
        frame = FLEX_excel.parse(
            sheet=sheet,
            usecols='A:B',
            names=['account_code','sub_account'],
            dtype={'sub_account':str},
        )
        FLEX_sub_account = pd.concat(
            [FLEX_sub_account,frame],
            axis=0
        )
    FLEX_sub_account.drop_duplicates(ignore_index=True,inplace=True)
    FLEX_sub_account.set_index('sub_account',inplace=True)
    FLEX_sub_account.sort_index(inplace=True)
    idx = DWH_sub_account.index.intersection(FLEX_sub_account.index)
    DWH_sub_account = DWH_sub_account.loc[idx]
    FLEX_sub_account = FLEX_sub_account.loc[idx]
    if DWH_sub_account.equals(FLEX_sub_account):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_sub_account.compare(FLEX_sub_account)
        return wrong_at


def trading_record():
    FLEX_excel = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','ROD0040.xls')
    )
    FLEX_sheets = FLEX_excel.sheet_names
    FLEX_trading_record = pd.DataFrame()
    for sheet in FLEX_sheets:
        frame = FLEX_excel.parse(
            sheet,
            usecols='B:D,F:T',
            header=None,
        )
        FLEX_trading_record = pd.concat(
            [FLEX_trading_record,frame],
            ignore_index=True,
            join='outer',
            axis=0
        )
    FLEX_trading_record.dropna(how='all',axis=0,inplace=True)
    FLEX_trading_record.dropna(how='all',axis=1,inplace=True)
    FLEX_trading_record.drop(labels=[0,1],axis=0,inplace=True)
    FLEX_trading_record.columns = [
        'date',
        'type_of_order',
        'stock',
        'channel',
        'user_of_order_placement',
        'account_code',
        'sub_account',
        'buy_volume',
        'buy_price',
        'buy_value',
        'sell_volume',
        'sell_price',
        'sell_value',
        'fee_rate',
        'fee',
        'tax_of_selling',
        'tax_of_share_dividend',
        'total_tax'
    ]
    idx = FLEX_trading_record['type_of_order'].isin(['Mua','Bán'])
    FLEX_trading_record = FLEX_trading_record.loc[idx]
    FLEX_trading_record['type_of_order'] = FLEX_trading_record['type_of_order'].map(
        {'Bán':'S','Mua':'B'}
    )
    FLEX_volume = FLEX_trading_record[['buy_volume','sell_volume']].sum(axis=1)
    FLEX_price = FLEX_trading_record[['buy_price','sell_price']].sum(axis=1)
    FLEX_value = FLEX_trading_record[['buy_value','sell_value']].sum(axis=1)
    cols = FLEX_trading_record.columns
    FLEX_trading_record.insert(cols.get_loc('sub_account')+1,'volume',FLEX_volume)
    FLEX_trading_record.insert(cols.get_loc('sub_account')+2,'price',FLEX_price)
    FLEX_trading_record.insert(cols.get_loc('sub_account')+3,'value',FLEX_value)
    FLEX_trading_record.drop([
        'buy_volume',
        'buy_price',
        'buy_value',
        'sell_volume',
        'sell_price',
        'sell_value',
    ],axis=1,inplace=True)
    f = lambda x:x if isinstance(x,str) else f'0{str(int(x))}'
    FLEX_trading_record['sub_account'] = FLEX_trading_record['sub_account'].map(f)
    FLEX_trading_record = FLEX_trading_record.set_index('sub_account').sort_index()
    start_date = FLEX_trading_record['date'].min().strftime('%Y/%m/%d')
    end_date = FLEX_trading_record['date'].max().strftime('%Y/%m/%d')
    DWH_trading_record = pd.read_sql(
        f"SELECT * FROM trading_record "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}';",
        connect
    )
    DWH_trading_record = DWH_trading_record.set_index('sub_account').sort_index()
    print(DWH_trading_record.columns)
    print(FLEX_trading_record.columns)
    FLEX_trading_record = FLEX_trading_record.astype(
        DWH_trading_record.dtypes.to_dict())  # table trading_record da duoc them cot vao nen khong giong ROD0040 tren flex
    # difference of number of records due to file extraction timming
    DWH_records = DWH_trading_record.shape[0]
    FLEX_records = FLEX_trading_record.shape[0]
    if np.abs(DWH_records/FLEX_records-1)<0.01:
        idx = DWH_trading_record.index.intersection(FLEX_trading_record.index)
        DWH_trading_record = DWH_trading_record.loc[idx]
        FLEX_trading_record = FLEX_trading_record.loc[idx]
    else:
        print('Record mismatching:')
        print(f'DWH_trading_record has {DWH_records} records '
              f'whereas FLEX_trading_record has {FLEX_records} records')
        return None
    if DWH_trading_record.equals(FLEX_trading_record):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_trading_record.compare(FLEX_trading_record)
        return wrong_at


def deposit_withdraw_stock():
    FLEX_excel = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','RSE2009.xls'),
    )
    FLEX_deposit_withdraw_stock = pd.DataFrame()
    sheets = FLEX_excel.sheet_names
    for sheet in sheets:
        frame = FLEX_excel.parse(
            sheet,
            header=None,
            usecols='B,D,F,G,I:K',
            skiprows=[0,1],
            skipfooter=1,
        )
        FLEX_deposit_withdraw_stock = pd.concat(
            [FLEX_deposit_withdraw_stock,frame],
            join='outer',
            ignore_index=True,
            axis=0,
        )
    FLEX_deposit_withdraw_stock.columns = [
        'date',
        'account_code',
        'ticker',
        'type',
        'volume',
        'remark',
        'creator'
    ]
    FLEX_deposit_withdraw_stock = FLEX_deposit_withdraw_stock.astype(
        {'type':str,'volume':float}
    )
    FLEX_deposit_withdraw_stock = FLEX_deposit_withdraw_stock.set_index(
        ['date','account_code']
    ).sort_index()
    start_date = FLEX_deposit_withdraw_stock.index.get_level_values(0).min().strftime('%Y/%m/%d')
    end_date = FLEX_deposit_withdraw_stock.index.get_level_values(0).max().strftime('%Y/%m/%d')
    DWH_deposit_withdraw_stock = pd.read_sql(
        f"SELECT * FROM deposit_withdraw_stock "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}';",
        connect
    )
    DWH_deposit_withdraw_stock = DWH_deposit_withdraw_stock.set_index(
        ['date','account_code']
    ).sort_index()
    # difference of number of records due to file extraction timming
    DWH_records = DWH_deposit_withdraw_stock.shape[0]
    FLEX_records = FLEX_deposit_withdraw_stock.shape[0]
    if np.abs(DWH_records/FLEX_records-1)<0.01:
        idx = DWH_deposit_withdraw_stock.index.intersection(
            FLEX_deposit_withdraw_stock.index
        )
        DWH_deposit_withdraw_stock = DWH_deposit_withdraw_stock.loc[idx]
        FLEX_deposit_withdraw_stock = FLEX_deposit_withdraw_stock.loc[idx]
    else:
        print('Record mismatching:')
        print(f'DWH_deposit_withdraw_stock has {DWH_records} records '
              f'whereas FLEX_deposit_withdraw_stock has {FLEX_records} records')
        return None
    if DWH_deposit_withdraw_stock.equals(FLEX_deposit_withdraw_stock):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_deposit_withdraw_stock.compare(FLEX_deposit_withdraw_stock)
        return wrong_at


def sub_account_deposit():
    FLEX_excel = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','RCI0001.xls'),
    )
    FLEX_sub_account_deposit = pd.DataFrame()
    sheets = FLEX_excel.sheet_names
    for sheet in sheets:
        frame = FLEX_excel.parse(
            sheet,
            header=None,
            usecols='C,E:H',
            names=[
                'sub_account',
                'opening_balance',
                'increase',
                'decrease',
                'closing_balance'
            ]
        )
        FLEX_sub_account_deposit = pd.concat(
            [FLEX_sub_account_deposit,frame],
            axis=0
        )
    date_cell = FLEX_sub_account_deposit.iloc[0,1]
    FLEX_sub_account_deposit.drop(labels=[0,1],inplace=True)
    FLEX_sub_account_deposit.set_index('sub_account',inplace=True)
    FLEX_sub_account_deposit.dropna(how='any',thresh=3,inplace=True)
    FLEX_sub_account_deposit = FLEX_sub_account_deposit.astype(np.float64)
    start_date = date_cell.split()[2]
    start_date = f'{start_date[-4:]}/{start_date[3:5]}/{start_date[:2]}'
    end_date = date_cell.split()[-1]
    end_date = f'{end_date[-4:]}/{end_date[3:5]}/{end_date[:2]}'
    DWH_sub_account_deposit = pd.read_sql(
        "SELECT * FROM sub_account_deposit "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}'",
        connect
    )
    open_idx = DWH_sub_account_deposit.groupby(['sub_account'],as_index=False)['date'].min()
    open_idx = pd.MultiIndex.from_frame(open_idx)
    open_balance = DWH_sub_account_deposit.set_index(['sub_account','date']).loc[open_idx,'opening_balance']
    open_balance = open_balance.reset_index(level=1,drop=True).to_frame()
    end_idx = DWH_sub_account_deposit.groupby(['sub_account'],as_index=False)['date'].max()
    close_idx = pd.MultiIndex.from_frame(end_idx)
    close_balance = DWH_sub_account_deposit.set_index(['sub_account','date']).loc[close_idx,'closing_balance']
    close_balance = close_balance.reset_index(level=1,drop=True).to_frame()
    increase = DWH_sub_account_deposit.groupby(['sub_account'])['increase'].sum()
    decrease = DWH_sub_account_deposit.groupby(['sub_account'])['decrease'].sum()
    DWH_aggregate = pd.concat(
        [open_balance,increase,decrease,close_balance],
        axis=1
    )
    DWH_records = DWH_aggregate.shape[0]
    FLEX_records = FLEX_sub_account_deposit.shape[0]
    if np.abs(DWH_records/FLEX_records-1)<0.01:
        idx = DWH_aggregate.index.intersection(FLEX_sub_account_deposit.index)
        DWH_aggregate = DWH_aggregate.loc[idx]
        FLEX_sub_account_deposit = FLEX_sub_account_deposit.loc[idx]
    else:
        print('Record mismatching:')
        print(f'DWH_new_sub_account has {DWH_records} records '
              f'whereas FLEX_new_sub_account has {FLEX_records} records')
        return None
    FLEX_sub_account_deposit = FLEX_sub_account_deposit.loc[idx]
    DWH_aggregate = DWH_aggregate.loc[idx]
    if DWH_aggregate.equals(FLEX_sub_account_deposit):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_aggregate.compare(FLEX_sub_account_deposit)
        return wrong_at


def new_sub_account():
    FLEX_excel = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','RCF0001.xls')
    )
    FLEX_new_sub_account = pd.DataFrame()
    sheets = FLEX_excel.sheet_names
    for sheet in sheets:
        frame = FLEX_excel.parse(
            sheet,
            usecols='C,E,H,I,K,L',
            header=None,
            skipfooter=2,
            names=[
                'sub_account',
                'address',
                'contract_type',
                'open_date',
                'account_code',
                'sub_account_type'
            ],
            dtype={'sub_account':object},
        )
        FLEX_new_sub_account = pd.concat(
            [FLEX_new_sub_account,frame]
        )
    date_cell = FLEX_new_sub_account.iloc[0,1]
    FLEX_new_sub_account.drop('address',axis=1,inplace=True)
    FLEX_new_sub_account = FLEX_new_sub_account.drop([0,1]).sort_index()
    start_date = date_cell.split()[2]
    start_date = f'{start_date[-4:]}/{start_date[3:5]}/{start_date[:2]}'
    end_date = date_cell.split()[-1]
    end_date = f'{end_date[-4:]}/{end_date[3:5]}/{end_date[:2]}'
    FLEX_new_sub_account.set_index(['sub_account','sub_account_type'],inplace=True)
    FLEX_new_sub_account['open_date'] = FLEX_new_sub_account['open_date'].map(
        lambda x:dt.datetime.strptime(x,'%d/%m/%Y')
    )
    DWH_new_sub_account = pd.read_sql(
        "SELECT * FROM new_sub_account "
        f"WHERE open_date BETWEEN '{start_date}' AND '{end_date}'",
        connect
    )
    DWH_new_sub_account.set_index(['sub_account','sub_account_type'],inplace=True)
    FLEX_new_sub_account = FLEX_new_sub_account[DWH_new_sub_account.columns]
    DWH_records = DWH_new_sub_account.shape[0]
    FLEX_records = FLEX_new_sub_account.shape[0]
    if np.abs(DWH_records/FLEX_records-1)<0.01:
        idx = DWH_new_sub_account.index.intersection(FLEX_new_sub_account.index)
        DWH_new_sub_account = DWH_new_sub_account.loc[idx]
        FLEX_new_sub_account = FLEX_new_sub_account.loc[idx]
    else:
        print('Record mismatching:')
        print(f'DWH_new_sub_account has {DWH_records} records '
              f'whereas FLEX_new_sub_account has {FLEX_records} records')
        return None
    if DWH_new_sub_account.equals(FLEX_new_sub_account):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_new_sub_account.compare(FLEX_new_sub_account)
        return wrong_at


def cashflow_bidv():
    FLEX_excel_RRM0068 = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','RRM0068.xls')
    )
    FLEX_RRM0068 = pd.DataFrame()
    sheets = FLEX_excel_RRM0068.sheet_names
    for sheet in sheets:
        frame = FLEX_excel_RRM0068.parse(
            sheet,
            header=None,
            usecols='C,E:G',
            names=['sub_account','bank_account','outflow_amount','remark'],
            skipfooter=2
        )
        FLEX_RRM0068 = pd.concat(
            [FLEX_RRM0068,frame]
        )
    date_cell = FLEX_RRM0068.iloc[0,3]
    FLEX_RRM0068.drop([0,1],inplace=True)
    FLEX_RRM0068.reset_index(drop=True,inplace=True)
    start_text_date_RRM0068 = date_cell.split()[3]
    start_date_RRM0068 = dt.datetime.strptime(start_text_date_RRM0068,'%d/%m/%Y')
    end_text_date_RRM0068 = date_cell.split()[-1]
    end_date_RRM0068 = dt.datetime.strptime(end_text_date_RRM0068,'%d/%m/%Y')

    FLEX_excel_RRM0069 = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','RRM0068.xls')
    )
    FLEX_RRM0069 = pd.DataFrame()
    sheets = FLEX_excel_RRM0069.sheet_names
    for sheet in sheets:
        frame = FLEX_excel_RRM0069.parse(
            sheet,
            header=None,
            usecols='C,E:G',
            names=['sub_account','bank_account','inflow_amount','remark'],
            skipfooter=2
        )
        FLEX_RRM0069 = pd.concat(
            [FLEX_RRM0069,frame]
        )
    date_cell = FLEX_RRM0069.iloc[0,3]
    FLEX_RRM0069.drop([0,1],inplace=True)
    FLEX_RRM0069.reset_index(drop=True,inplace=True)
    FLEX_cashflow_bidv = pd.concat(
        [FLEX_RRM0068,FLEX_RRM0069],
        ignore_index=True
    )
    FLEX_cashflow_bidv.fillna(0,inplace=True)
    FLEX_cashflow_bidv = FLEX_cashflow_bidv.astype(
        dtype={'inflow_amount':np.float64,'outflow_amount':np.float64}
    )
    start_text_date_RRM0069 = date_cell.split()[3]
    start_date_RRM0069 = dt.datetime.strptime(start_text_date_RRM0069,'%d/%m/%Y')
    end_text_date_RRM0069 = date_cell.split()[-1]
    end_date_RRM0069 = dt.datetime.strptime(end_text_date_RRM0069,'%d/%m/%Y')

    start_date = max([start_date_RRM0068,start_date_RRM0069]).strftime('%Y/%m/%d')
    end_date = min([end_date_RRM0068,end_date_RRM0069]).strftime('%Y/%m/%d')

    DWH_cashflow_bidv = pd.read_sql(
        "SELECT * FROM cashflow_bidv "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}'",
        connect
    )
    DWH_cashflow_bidv.drop('date',axis=1,inplace=True)
    DWH_cashflow_bidv['remark'] = DWH_cashflow_bidv['remark'].str.rstrip()
    FLEX_cashflow_bidv = FLEX_cashflow_bidv[DWH_cashflow_bidv.columns]

    DWH_cashflow_bidv = DWH_cashflow_bidv.sort_values(
        ['sub_account','inflow_amount','outflow_amount']
    ).reset_index(drop=True)
    FLEX_cashflow_bidv = FLEX_cashflow_bidv.sort_values(
        ['sub_account','inflow_amount','outflow_amount']
    ).reset_index(drop=True)
    DWH_records = DWH_cashflow_bidv.shape[0]
    FLEX_records = FLEX_cashflow_bidv.shape[0]
    if np.abs(DWH_records/FLEX_records-1)<0.01:
        idx = DWH_cashflow_bidv.index.intersection(FLEX_cashflow_bidv.index)
        DWH_cashflow_bidv = DWH_cashflow_bidv.loc[idx]
        FLEX_cashflow_bidv = FLEX_cashflow_bidv.loc[idx]
    else:
        print('Record mismatching:')
        print(f'DWH_new_sub_account has {DWH_records} records '
              f'whereas FLEX_new_sub_account has {FLEX_records} records')
        return None
    if DWH_cashflow_bidv.equals(FLEX_cashflow_bidv):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_cashflow_bidv.compare(FLEX_cashflow_bidv)
        return wrong_at


def depository_fee():
    FLEX_excel = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','RSE9992.xls'),
    )
    sheets = FLEX_excel.sheet_names
    FLEX_depository_fee = pd.DataFrame()
    for sheet in sheets:
        frame = FLEX_excel.parse(
            sheet,
            headers=None,
            usecols='B,C,E',
            names=['account_code','sub_account','fee_amount'],
            skipfooter=3,
        )
        FLEX_depository_fee = pd.concat(
            [FLEX_depository_fee,frame]
        )
    date_cell = FLEX_depository_fee.iloc[0,0]
    FLEX_depository_fee.drop('account_code',axis=1,inplace=True)
    FLEX_depository_fee = FLEX_depository_fee.drop([0,1]).set_index('sub_account').sort_index()
    start_date = date_cell.split()[2]
    start_date = f'{start_date[-4:]}/{start_date[3:5]}/{start_date[:2]}'
    end_date = date_cell.split()[-1]
    end_date = f'{end_date[-4:]}/{end_date[3:5]}/{end_date[:2]}'
    DWH_depository_fee = pd.read_sql(
        "SELECT * FROM depository_fee "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}'",
        connect
    )
    DWH_depository_fee = DWH_depository_fee.drop('date',axis=1).set_index('sub_account').sort_index()
    DWH_depository_fee = DWH_depository_fee.groupby('sub_account').sum()
    DWH_records = DWH_depository_fee.shape[0]
    FLEX_records = FLEX_depository_fee.shape[0]
    if np.abs(DWH_records/FLEX_records-1)<0.01:
        idx = DWH_depository_fee.index.intersection(FLEX_depository_fee.index)
        DWH_depository_fee = DWH_depository_fee.loc[idx]
        FLEX_depository_fee = FLEX_depository_fee.loc[idx]
    else:
        print('Record mismatching:')
        print(f'DWH_depository_fee has {DWH_records} records '
              f'whereas FLEX_depository_fee has {FLEX_records} records')
        return None
    if DWH_depository_fee.equals(FLEX_depository_fee):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_depository_fee.compare(FLEX_depository_fee)
        return wrong_at


def money_in_out_transfer():
    FLEX_excel = pd.ExcelFile(
        join(realpath(dirname(__file__)),'flex_data','RCI1025.xls')
    )
    FLEX_money_in_out_transfer = pd.DataFrame()
    sheets = FLEX_excel.sheet_names
    for sheet in sheets:
        frame = FLEX_excel.parse(
            sheet,
            header=None,
            usecols='B,C,E:G,I:K,O',
            names=[
                'date',
                'time',
                'sub_account',
                'bank_account',
                'bank',
                'transaction_id',
                'amount',
                'remark',
                'status'
            ],
            dtype={
                'amount':float,
                'sub_account':object,
                'bank_account':object,
                'transaction_id':object,

            },
            skipfooter=3
        )
        FLEX_money_in_out_transfer = pd.concat(
            [FLEX_money_in_out_transfer,frame]
        )
    date_cell = FLEX_money_in_out_transfer.iloc[1,0]
    FLEX_money_in_out_transfer.drop([0,1,2],inplace=True)
    start_date = date_cell.split()[2]
    start_date = f'{start_date[-4:]}/{start_date[3:5]}/{start_date[:2]}'
    end_date = date_cell.split()[-1]
    end_date = f'{end_date[-4:]}/{end_date[3:5]}/{end_date[:2]}'
    DWH_money_in_out_transfer = pd.read_sql(
        "SELECT * FROM money_in_out_transfer "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}'",
        connect
    )
    if DWH_money_in_out_transfer.equals(FLEX_money_in_out_transfer):
        return True
    else:
        print('Wrong at:')
        wrong_at = DWH_money_in_out_transfer.compare(FLEX_money_in_out_transfer)
        return wrong_at
