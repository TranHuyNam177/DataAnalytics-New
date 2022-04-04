from request.stock import *

room = pd.read_excel(r'C:\Users\hiepdang\Desktop\BaocaoNgay_RMD\FileMau.xlsx',sheet_name='230007',usecols='A:E',
                     index_col=0)
matched = pd.read_excel(r'C:\Users\hiepdang\Desktop\BaocaoNgay_RMD\FileMau.xlsx',sheet_name='Matched',usecols='A:D',
                        index_col=0)
intranet = pd.read_excel(r'C:\Users\hiepdang\Desktop\BaocaoNgay_RMD\FileMau.xlsx',sheet_name='intranet',usecols='A:H',
                         index_col=0)

mlist = pd.read_excel(r'C:\Users\hiepdang\Desktop\BaocaoNgay_RMD\Margin list.xlsx',index_col=0)

tonghop = pd.DataFrame(index=mlist.index)

#############################
# Used general room
# Cach 1:
for cp in tonghop.index:
    tonghop.loc[cp,'Used general room'] = room.loc[cp,'Room hệ thống đã sử dụng']

del tonghop['Used general room']  # xóa kết quả cách 1
# Cach 2:
tonghop.insert(0,'Used general room',room['Room hệ thống đã sử dụng'])

# Note: Ở đây em dùng phương thức insert thay thế cho vòng lặp for, mọi người tham khảo cú pháp phương thức insert ở đây nhé:
# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.insert.html? (Ctrl + Click)

#############################
# Used special room
# Cach 1:
for cp in tonghop.index:
    tonghop.loc[cp,'Used special room'] = room.loc[cp,'Room đặc biệt đã sử dụng']

del tonghop['Used special room']  # xóa kết quả cách 1
# Cach 2:
tonghop.insert(0,'Used special room',room['Room hệ thống đã sử dụng'])

#############################
# Total used room (previous session)
# Cach 1:
for cp in tonghop.index:
    tonghop.loc[cp,'Total used room (previous session)'] \
        = tonghop.loc[cp,'Used general room']+tonghop.loc[cp,'Used special room']

del tonghop['Total used room (previous session)']  # xóa kết quả cách 1
# Cach 2:
tonghop['Total used room (previous session)'] \
    = tonghop['Used general room']+tonghop['Used special room']

#############################
