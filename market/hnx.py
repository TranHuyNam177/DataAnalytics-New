import pickle
from dependency import *
from request import connect_DWH_ThiTruong
from datawarehouse import SEQUENTIALINSERT, DROPDUPLICATES

class HNXMessages:

    def __init__(
        self,
        startPosition:int=0,
    ):

        self.lastPosition = startPosition
        with open(join(dirname(__file__),'mapper','tableNameMapper.pickle'),'rb') as file:
            self.tableNameMapper = pickle.load(file)
        with open(join(dirname(__file__),'mapper','metaData.pickle'),'rb') as file:
            self.metaData = pickle.load(file)
        self.nameFrameIterable = None

    def __repr__(self):
        return '<HNXMessages>'

    def _readFile(self):
        self.filePath = r"C:\Users\hiepdang\Downloads\2022-05-30\ClientLog\Log_05302022.log"
        # self.filePath = join(r'Z:',listdir(r'Z:')[0]) # Z folder should have one file
        with open(self.filePath,'rb') as file:
            file.seek(self.lastPosition)
            self.lastMessages = file.readlines()
            # Strip first message due to incomplete content
            del self.lastMessages[0]
            self.lastMessages = iter(self.lastMessages)
            self.lastPosition = file.tell() # EOF position
        
    def _toIterable(self):
        for message in self.lastMessages:
            messageInList = [content.decode('utf-8') for content in message.split(b'\x01')]
            messageInList.remove('\r\n')
            messageDict = {element.split('=')[0]: self._cleanData(element.split('=')[0])(element.split('=')[1]) for element in messageInList}
            print(messageDict)
            columnsCodes = set(messageDict.keys())
            for tableName, allColumns in self.tableNameMapper.items():
                if columnsCodes.issubset(allColumns) and columnsCodes:
                    messageFrame = pd.DataFrame(messageDict,index=[0])
                    messageFrame = messageFrame.rename(columns=lambda x: self.metaData[x][0])
                    print(tableName)
                    print(messageFrame)
                    print(messageFrame.columns)
                    print(messageFrame.info())
                    yield tableName
                    yield messageFrame
                    break
            else:
                raise ValueError("Can't find table name")

    def _writeDB(self):
        for tableName,frame in zip(*[iter(self.nameFrameIterable)]*2):
            # SEQUENTIALINSERT(connect_DWH_ThiTruong,tableName,frame)
            for colName in frame.columns:
                value = frame.loc[frame.index[0],colName]
                SEQUENTIALINSERT(connect_DWH_ThiTruong,tableName,pd.DataFrame({colName:[value]}))
            print('------------------------')

    def realTime(self):
        runTime = dt.datetime.now().time()
        if runTime.hour < 12:
            session = 'SA'
        else:
            session = 'CH'

        while True:
            if dt.datetime.now().time() > dt.time(hour=12) and session == 'SA':
                break
            elif dt.datetime.now().time() > dt.time(hour=17) and session == 'CH':
                break
            self._readFile()
            self.nameFrameIterable = self._toIterable()
            self._writeDB()

    def eod(self):
        self._readFile()
        self.nameFrameIterable = self._toIterable()
        self._writeDB()

    @staticmethod
    def _cleanData(colCode):
        varcharTags = {
            '8','35','49','2','21','27','18','15','55','425','336','340','167','106','107','232','426','421','422','341','33','56','60','800','425',
            '340','326','327','167',
        }
        timeTags_HHMMSSmm = {'4'}
        timeTags_HHMMSS = {'399'}
        dateTags_YYYYMMDD = {'19','388','541','28'}
        dateTags_DDMMYYYY = {'802','803'}
        datetimeTags = {'52','57','225'}

        if colCode in varcharTags:
            converter = lambda x: x
        # For date/time values, insert as string and let SQL Server converts to date/time itself
        elif colCode in timeTags_HHMMSSmm:
            converter = lambda x: dt.datetime.strptime(x,'%H:%M:%S:%f').strftime('%H:%M:%S')
        elif colCode in timeTags_HHMMSS:
            converter = lambda x: dt.datetime.strptime(x,'%H:%M:%S').strftime('%H:%M:%S')
        elif colCode in dateTags_YYYYMMDD:
            converter = lambda x: '9999-12-31' if int(x[:4])==1 else dt.datetime.strptime(x,'%Y%m%d').strftime('%Y-%m-%d')
        elif colCode in dateTags_DDMMYYYY:
            converter = lambda x: '9999-12-31' if int(x[-4:])==1 else dt.datetime.strptime(x,'%d/%m/%Y').strftime('%Y-%m-%d')
        elif colCode in datetimeTags:
            converter = lambda x: '9999-12-31 23:59:59' if int(x[:4])==1 else dt.datetime.strptime(x,'%Y%m%d-%H:%M:%S').strftime('%Y-%m-%d %H:%M:%S')
        else:
            converter = lambda x: float(x)

        return converter

