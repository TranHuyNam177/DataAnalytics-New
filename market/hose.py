from dependency import *
from market import __convertInteger__

def SECURITY(d):

    filePath = fr"Y:\BACKUP{__convertInteger__(d)}\SECURITY.DAT"
    sequenceLength = 294 # tested
    dataMap = {
        'StockNo':'H',
        'StockSymbol':'8s',
        'StockType':'c',
        'Ceiling':'L',
        'Floor':'L',
        'BigLotValue':'d',
        'SecurityName':'25s',
        'SectorNo':'c',
        'Designated':'c',
        'Suspension':'c',
        'Delist':'c',
        'HaltResumeFlag':'c',
        'Split':'c',
        'Benefit':'c',
        'Meeting':'c',
        'Notice':'c',
        'ClientIDRequest':'c',
        'CouponRate':'H',
        'IssueDate':'6s',
        'MatureDate':'6s',
        'AvrPrice':'L',
        'ParValue':'H',
        'SDCFlag':'s',
        'PriorClosePrice':'L',
        'PriorCloseDate':'6s',
        'ProjectOpen':'L',
        'OpenPrice':'L',
        'Last':'L',
        'LastVol':'L',
        'LastVal':'d',
        'Highest':'L',
        'Lowest':'L',
        'Totalshare':'d',
        'TotalValue':'d',
        'AccumulateDeal':'H',
        'BigDeal':'H',
        'BigVol':'L',
        'BigVal':'d',
        'OddDeal':'H',
        'OddVol':'L',
        'OddVal':'d',
        'Best1Bid':'L',
        'Best1BidVolume':'L',
        'Best2Bid':'L',
        'Best2BidVolume':'L',
        'Best3Bid':'L',
        'Best3BidVolume':'L',
        'Best1Offer':'L',
        'Best1OfferVolume':'L',
        'Best2Offer':'L',
        'Best2OfferVolume':'L',
        'Best3Offer':'L',
        'Best3OfferVolume':'L',
    }
    __pattern = '<' + ''.join(dataMap.values())

    with open(filePath,'rb') as file:
        data = file.read()
        sequenceNumber = len(data) // sequenceLength
        records = []
        for sequence in range(sequenceNumber):
            loc = sequence * sequenceLength
            records.append(struct.unpack_from(__pattern,data,loc))

    dataTable = pd.DataFrame(records,columns=dataMap.keys())
    for col in dataTable.columns:
        if dataTable[col].dtype == object: # Xử lý các cột string
            dataTable[col] = dataTable[col].str.decode('utf8').str.rstrip().str.lstrip()
            dataTable[col] = dataTable[col].replace({'00+':None},regex=True).replace({'':None})
    dataTable = dataTable.replace({'\x00+':None},regex=True)

    for col in dataTable.columns:
        if 'date' in col.lower():  # Xử lý các cột ngày
            dataTable[col] = pd.to_datetime(dataTable[col],format='%d%m%y').replace({'':None,pd.NaT:None})

    return dataTable


def CS_VN30():

    filePath = r"X:\CS_VN30.DAT"
    sequenceLength = 10 # tested

    dataMap = {
        'StockNo':'H',
        'StockSymbol':'8s',
    }
    __pattern = '<' + ''.join(dataMap.values())

    with open(filePath,'rb') as file:
        data = file.read()
        sequenceNumber = len(data) // sequenceLength
        records = []
        for sequence in range(sequenceNumber):
            loc = sequence * sequenceLength
            records.append(struct.unpack_from(__pattern,data,loc))

    dataTable = pd.DataFrame(records,columns=dataMap.keys())
    for col in dataTable.columns:
        if dataTable[col].dtype == object: # Xử lý các cột string
            dataTable[col] = dataTable[col].str.decode('utf8').str.rstrip().str.lstrip()
    runDate = dt.datetime.now().replace(hour=0,minute=0,second=0,microsecond=0)
    dataTable.insert(0,'Date',runDate)

    return dataTable


