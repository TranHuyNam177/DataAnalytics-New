import pandas as pd

class DocGGSheet:
    def __init__(self):
        self.__ggSheetID = '1fivaSKYlcyMt9g1a1CPEm4HFEWsJ3Oni-GWqzBVjQ24'
        self.__originalURL = f'https://docs.google.com/spreadsheets/d/{self.__ggSheetID}/edit#gid=1617498253'
        self.__finalURL = self.__originalURL.replace('/edit#gid=', '/export?format=csv&gid=')

    def readSheet(self):
        table = pd.read_csv(
            self.__finalURL,
            usecols=[0, 1, 2, 3],
            names=['Date', 'SoTK', 'KiemTraLienHe', 'PhanHoiCuaKH'],
            skiprows=1,
            dtype={
                'Date': object,
                'SoTK': object,
                'KiemTraLienHe': object,
                'PhanHoiCuaKH': object
            }
        )
        table['Date'] = pd.to_datetime(table['Date'])
        return table
