from request import *
from request.stock import *

table = pd.read_sql(
    """
    WITH
    [taikhoan] AS (
        SELECT DISTINCT
            [holding].[account_code]
        FROM
            [holding]
        WHERE
            [holding].[date] = '2022-01-26'
            AND [holding].[ticker] IN ('FLC','HPG','DIG')
    )
    SELECT
        [holding].[account_code],
        [holding].[ticker],
        [holding].[volume]
    FROM
        [holding]
    RIGHT JOIN
        [taikhoan]
    ON [taikhoan].[account_code] = [holding].[account_code]
    WHERE [holding].[date] = '2022-01-26'
    """,
    connect_DWH_CoSo
)
tickers = table['ticker'].unique()
cash = table.loc[table['ticker']=='Tien',['account_code','volume']].set_index('account_code').squeeze()

price_mapper = dict()
for ticker in tickers:
    try:
        price_mapper[ticker] = ta.hist(ticker)['close'].iloc[-1]
    except (Exception,):
        price_mapper[ticker] = np.nan

table['price'] = table['ticker'].map(price_mapper)*1000
table = table.loc[~table['price'].isna()]
table['value'] = table['volume'] * table['price']

sumSeries = table.groupby('account_code')['value'].sum()
table['total'] = table['account_code'].map(sumSeries)
table['weight'] = table['value'] / table['total']
table['cash'] = table['account_code'].map(cash)
table['cash'] = table['cash'].fillna(0)

rating_gen = pd.read_csv(r"C:\Users\hiepdang\PycharmProjects\DataAnalytics\credit_rating\result\result_table_gen.csv",index_col='ticker',usecols=['ticker','2021q2'])
rating_bank = pd.read_csv(r"C:\Users\hiepdang\PycharmProjects\DataAnalytics\credit_rating\result\result_table_bank.csv",index_col='ticker',usecols=['ticker','2021q2'])
rating_sec = pd.read_csv(r"C:\Users\hiepdang\PycharmProjects\DataAnalytics\credit_rating\result\result_table_sec.csv",index_col='ticker',usecols=['ticker','2021q2'])
rating_ins = pd.read_csv(r"C:\Users\hiepdang\PycharmProjects\DataAnalytics\credit_rating\result\result_table_ins.csv",index_col='ticker',usecols=['ticker','2021q2'])

rating = pd.concat([rating_gen,rating_bank,rating_ins,rating_sec]).squeeze()

def rating_mapper(x):
    if x < 25 or x != x:
        result = 'D'
    elif x < 50:
        result = 'C'
    elif x < 75:
        result = 'B'
    else:
        result = 'A'
    return result

rating = rating.map(rating_mapper)
rating = rating.drop('BM_')

table['rating'] = table['ticker'].map(rating).fillna('D')
table['discount_rate'] = table['rating'].map({'D':0.15,'C':0.1,'B':0.05,'A':0})
table['discounted_value'] = table['value'] * table['discount_rate']

table.to_excel('table.xlsx')

