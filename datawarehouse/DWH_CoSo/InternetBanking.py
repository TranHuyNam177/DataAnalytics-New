from datawarehouse import *
from automation.finance import *
from automation.finance import BankCurrentBalance
from automation.finance import BankDepositBalance
from automation.finance import BankTransactionHistory

class BankFailedException(Exception):
    pass

def GetDataMonitor(func):

    def wrapper(*args,**kwargs):
        bank = args[0].bank
        outlook = Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'hiepdang@phs.vn'
        try:
            bankObject = func(*args,**kwargs)
            mail.Subject = f"{func.__name__}.{bank} Done"
            body = f"""
                <html>
                    <head></head>
                    <body>
                        <p style="font-family:Times New Roman; font-size:90%"><i>
                            -- Generated by Reporting System
                        </i></p>
                    </body>
                </html>
                """
            mail.HTMLBody = body
            mail.Send()
        except (Exception,):
            mail.Subject = f"{func.__name__}.{bank} Failed"
            traceback_message = traceback.format_exc()
            body = f"""
                <html>
                    <head></head>
                    <body>
                        <p style="font-family:Consolas; font-size:90%">
                        {traceback_message}
                        </p>
                        <p style="font-family:Times New Roman; font-size:90%"><i>
                            -- Generated by Reporting System
                        </i></p>
                    </body>
                </html>
                """
            mail.HTMLBody = body
            mail.Send()
            raise BankFailedException(f"{bank} Failed")
        return bankObject
    return wrapper


@GetDataMonitor
def runBankCurrentBalance(bankObject,fromDate,toDate):
    """
    :param bankObject: Bank name
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    if bankObject.bank == 'BIDV':
        balanceTable = BankCurrentBalance.runBIDV(bankObject,fromDate,toDate)
    elif bankObject.bank == 'EIB':
        balanceTable = BankCurrentBalance.runEIB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'IVB':
        balanceTable = BankCurrentBalance.runIVB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'VTB':
        balanceTable = BankCurrentBalance.runVTB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'VCB':
        balanceTable = BankCurrentBalance.runVCB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'OCB':
        balanceTable = BankCurrentBalance.runOCB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'TCB':
        balanceTable = BankCurrentBalance.runTCB(bankObject,fromDate,toDate)
    else:
        print(f'Invalid bank name: {bankObject.bank}')
        return bankObject

    fromDateString = fromDate.strftime('%Y-%m-%d')
    toDateString = toDate.strftime('%Y-%m-%d')
    DELETE(connect_DWH_CoSo,'BankCurrentBalance',f"""WHERE [Date] BETWEEN '{fromDateString}' AND '{toDateString}' AND [Bank] = '{bankObject.bank}'""")
    BATCHINSERT(connect_DWH_CoSo,'BankCurrentBalance',balanceTable)

    return bankObject # destroy the object to close opening Chrome driver (call __del__ magic method)

@GetDataMonitor
def runBankDepositBalance(bankObject):

    """
    :param bankObject: Bank Object (đã login)
    """

    if bankObject.bank == 'BIDV':
        balanceTable = BankDepositBalance.runBIDV(bankObject)
    elif bankObject.bank == 'IVB':
        balanceTable = BankDepositBalance.runIVB(bankObject)
    elif bankObject.bank == 'VTB':
        balanceTable = BankDepositBalance.runVTB(bankObject)
    elif bankObject.bank == 'VCB':
        balanceTable = BankDepositBalance.runVCB(bankObject)
    elif bankObject.bank == 'OCB':
        balanceTable = BankDepositBalance.runOCB(bankObject)
    else:
        print(f'Invalid bank name: {bankObject.bank}')
        return bankObject

    dateString = balanceTable['Date'].max().strftime('%Y-%m-%d')
    DELETE(connect_DWH_CoSo,'BankDepositBalance',f"""WHERE [Date] = '{dateString}' AND [Bank] = '{bankObject.bank}'""")
    BATCHINSERT(connect_DWH_CoSo,'BankDepositBalance',balanceTable)

    return bankObject

@GetDataMonitor
def runBankTransactionHistory(bankObject,fromDate,toDate):
    """
    :param bankObject: Bank Object (đã login)
    :param fromDate: Ngày bắt đầu lấy dữ liệu
    :param toDate: Ngày kết thúc lấy dữ liệu
    """

    if bankObject.bank == 'BIDV':
        transactionTable = BankTransactionHistory.runBIDV(bankObject,fromDate,toDate)
    elif bankObject.bank == 'IVB':
        transactionTable = BankTransactionHistory.runIVB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'VCB':
        transactionTable = BankTransactionHistory.runVCB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'VTB':
        transactionTable = BankTransactionHistory.runVTB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'EIB':
        transactionTable = BankTransactionHistory.runEIB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'OCB':
        transactionTable = BankTransactionHistory.runOCB(bankObject,fromDate,toDate)
    elif bankObject.bank == 'TCB':
        transactionTable = BankTransactionHistory.runTCB(bankObject,fromDate,toDate)
    else:
        print(f'Invalid bank name: {bankObject.bank}')
        return bankObject

    if transactionTable.empty:
        print('No data to insert')
    else:
        fromTimeString = fromDate.strftime('%Y-%m-%d 00:00:00')
        toTimeString = toDate.strftime('%Y-%m-%d 23:59:59')
        DELETE(connect_DWH_CoSo,'BankTransactionHistory',f"""WHERE [Time] BETWEEN '{fromTimeString}' AND '{toTimeString}' AND [Bank] = '{bankObject.bank}'""")
        BATCHINSERT(connect_DWH_CoSo,'BankTransactionHistory',transactionTable)

    return bankObject

