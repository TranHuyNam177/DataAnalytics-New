from datawarehouse import *
from automation.finance import getBankData

def check(bank):

    fnData = pd.read_excel(r"C:\Users\hiepdang\Downloads\Daily Bank Report.xlsx",skiprows=3)
    def mapper(x):
        if 'OCB' in x:
            result = 'OCB'
        elif 'EXIMBANK' in x:
            result = 'EIB'
        elif 'VIETINBANK' in x:
            result = 'VTB'
        elif 'INDOVINABANK' in x or 'INDOVINA BANK' in x:
            result = 'IVB'
        elif 'VIETCOMBANK' in x:
            result = 'VCB'
        elif 'BIDV' in x:
            result  = 'BIDV'
        else:
            result = None
        return result

    fnData.columns = ['Date','Bank','AccountNumber','Balance']
    fnData['Bank'] = fnData['Bank'].map(mapper)
    fnData['AccountNumber'] = fnData['AccountNumber'].str.replace('.','')
    fnData = fnData.dropna(how='any')

    dbData = pd.read_sql(
        """
        SELECT [Date], [Bank], [AccountNumber], [Balance] FROM [BankCurrentBalance]
        """,
        connect_DWH_CoSo,
    )

    compareTable = pd.merge(fnData,dbData,how='outer',on=['Date','Bank','AccountNumber'],suffixes=['_fn','_db'])
    compareTable['diff'] = compareTable['Balance_fn'] - compareTable['Balance_db']
    compareTable = compareTable.loc[compareTable['Bank']==bank]

    return compareTable
