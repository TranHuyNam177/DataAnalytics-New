from dependency import *


###############################################################################


def trytoint(obj):
    """
    This function tries to convert an arbitrary object to integer
  
    :param obj: arbitrary object
  
    :return: object of int type if possible
    """

    try:
        obj = int(obj)
    except ValueError:
        pass
    return obj


def priceKformat(tick_val,pos) -> str:
    """
    Turns large tick values (in the billions, millions and thousands)
    such as 4500 into 4.5K and also appropriately turns 4000 into 4K
    (no zero after the decimal)
  
    :param tick_val: price value
    :param pos: ignored
    :type tick_val: float
  
    :return: formated price value
    """

    if tick_val>=1000:
        val = round(tick_val/1000,1)
        if tick_val%1000>0:
            new_tick_format = '{:,}K'.format(val)
        else:
            new_tick_format = '{:,}K'.format(int(val))
    else:
        new_tick_format = int(tick_val)
    new_tick_format = str(new_tick_format)
    return new_tick_format


def adjprice(price: Union[float,int]) -> str:
    """
    This function returns adjusted price for minimum price steps,
    used for display only
  
    :param price: stock price
    :type price: float or int
  
    :return: str
    """

    if price<10000:
        price_string = f'{int(round(price,-1)):,d}'
    elif 10000<=price<50000:
        price_string = f'{50*int(round(price/50)):,d}'
    else:
        price_string = f'{int(round(price,-2)):,d}'

    return price_string


def bdate(date: str,bdays: int = 0) -> str:
    """
    This function return the business date before/after a certain business days
    since a given date
  
    :param date: allow string like 'yyyy-mm-dd', 'yyyy/mm/dd'
    :param bdays: allow positive integer (after) or negative integer (before)
  
    :return: string of type 'yyyy-mm-dd'
    """

    year = int(date[:4])
    month = int(date[5:7])
    day = int(date[-2:])

    step = int(np.sign(bdays))

    given_date = dt.datetime(year,month,day)
    result_date = given_date

    d = 0
    while abs(d) < abs(bdays):
        d += step
        result_date += dt.timedelta(days=step)
        while result_date.weekday() in holidays.WEEKEND or result_date in holidays.VN():
            result_date += dt.timedelta(days=step)

    result_date = result_date.strftime('%Y-%m-%d')

    return result_date


def btime(time: str,bhours: int = 0) -> str:
    """
    This function returns the business datetime before/after a certain business hours
    since a given datetime
  
    :param time: allow string like 'yyyy-mm-dd hh:mm:ss', 'yyyy/mm/dd hh:mm:ss'
    :param bhours: allow positive integer (after) or negative integer (before)
  
    :return: string of type 'yyyy-mm-dd hh:mm:ss'
    """

    year = int(time[:4])
    month = int(time[5:7])
    day = int(time[8:10])

    hour = int(time[11:13])
    minute = int(time[14:16])
    second = int(time[17:19])

    step = int(np.sign(bhours))

    given_time = dt.datetime(year,month,day,hour,minute,second)
    result_time = given_time

    d = 0
    while abs(d)<abs(bhours):
        d += step
        result_time += dt.timedelta(hours=step)
        while result_time.weekday() in holidays.WEEKEND or result_time in holidays.VN():
            result_time += dt.timedelta(days=step)

    result_time = result_time.strftime('%Y-%m-%d %H:%M:%S')

    return result_time


def seopdate(period: str) -> tuple:
    """
    This function returns (start date, end date) of a period
  
    :param period: target period: '2020q4', '2020q3', etc.
    :type period: str
  
    :return: tuple
    """

    year = int(period[:4])
    quarter = int(period[-1])

    # start of the period
    sop_date = dt.datetime(year=year,month=3*quarter-2,day=1)
    while sop_date.weekday() in holidays.WEEKEND \
            or sop_date in holidays.VN():
        sop_date += dt.timedelta(days=1)
    sop_date = sop_date.strftime('%Y-%m-%d')

    # end of the period
    fq = lambda quarter:1 if quarter==4 else 3*quarter+1
    fy = lambda year:year+1 if quarter==4 else year
    eop_date = dt.datetime(year=fy(year),month=fq(quarter),day=1)+dt.timedelta(days=-1)
    while eop_date.weekday() in holidays.WEEKEND or eop_date in holidays.VN():
        eop_date -= dt.timedelta(days=1)
    eop_date = eop_date.strftime('%Y-%m-%d')

    return sop_date,eop_date


def period_cal(period: str,years: int = 0,quarters: int = 0) -> str:
    """
    This function return resutled period from addition/substraction operations
  
    :param period: original period
    :param years: number of years added (+) or substracted (-) from the original period
    :param quarters: number of quarters added (+) or substracted (-) from the original period
    :type period: str
    :type years: int
    :type quarters: int
  
    :return: tuple
    """

    year = int(period[:4])
    quarter = int(period[-1])

    if (quarter+quarters)%4==0:
        year_new = year+years+(quarter+quarters)//4-1
        quarter_new = 4
    else:
        year_new = year+years+(quarter+quarters)//4
        quarter_new = (quarter+quarters)%4

    period_new = f'{year_new}q{quarter_new}'

    return period_new


def fc_price(
    ref_price:Union[float,int],
    price_type:str,
    exchange:str
) -> int:
    """
    This function returns ceil/floor price given a reference price
  
    :param ref_price: stock price
    :param price_type: allow either 'ceil' (for ceil price) or 'floor' (for floor price)
    :param exchange: allow values in ['HOSE', 'HNX', 'UPCOM']. Do not allow 'all'
  
    :return: int
    """

    if exchange=='HOSE':
        rreturn = 0.07
    elif exchange=='HNX':
        rreturn = 0.1
    else:
        rreturn = 0.15

    if price_type=='ceil':
        pass
    elif price_type=='floor':
        rreturn *= -1
    else:
        raise ValueError('invalid price type')

    bmk = ref_price*(1+rreturn)

    price = ref_price
    result_price = [price]
    condition = True
    while condition:
        if exchange=='HOSE':
            if price<10000:
                step = 10
            elif 10000<=price<50000:
                step = 50
            else:
                step = 100
        elif exchange=='HNX':
            step = 100
        elif exchange=='UPCOM':
            step = 100
        else:
            raise ValueError('invalid exchange')

        if price_type=='ceil':
            pass
        elif price_type=='floor':
            step *= -1
        else:
            raise ValueError('invalid price type')

        price = int(adjprice(result_price[-1]+step).replace(',',''))
        result_price += [price]
        condition = (result_price[-1]<=bmk and price_type=='ceil') \
                    or (result_price[-1]>=bmk and price_type=='floor')

    return result_price[-2]


def process_address(raw_address):
    def handle_missing_value(x):
        if pd.isnull(x) or x in ['.','']:
            x = 'Unknown'
        return x

    def highest_level(x):
        if ',' in x:
            result = x.split(',')[-1]
        elif '-' in x:
            result = x.split('-')[-1]
        elif '.' in x:
            result = x.split('.')[-1]
        elif '_' in x:
            result = x.split('_')[-1]
        else:
            result = x[-20:]
        return result

    def lowercase(x):
        if x!='Unknown':
            x = x.lower()
        return x

    def remove_accent(x):
        return unidecode.unidecode(x)

    def remove_separator(x):
        for sep in [' ','.','-','/','(',')','_',',',':','#',"'"]:
            if sep in x:
                x = x.replace(sep,'')
        return x

    dictionary = dict()
    dictionary['Ho Chi Minh City'] = [
        'hcm','hochiminh','saigon','quan1','quan2','quan3','quan4','quan5',
        'quan6','quan7','quan8','quan9','quan10','quan11','quan12','q1','q2',
        'q3','q4','q5','q6','q7','q8','q9','q10','q11','q12','govap','tanbinh',
        'tanphu','binhthanh','phunhuan','thuduc','binhtan','binhchanh','cuchi',
        'hocmon','nhabe','cangio','qgv','qtb','qtp','qbt','qpn','qtd','qbc',
        'qcc','qhm','qnb','qcg','phumyhung','cauonglanh','haithuonglanong',
        'hoocmon','ptanphong','ptantaoa','district1','district2','district3',
        'district4','district5','district6','district7','district8','district9',
        'district10','district11','district12'
    ]

    dictionary['Hai Phong'] = [
        'haiphong','hongbang','lechan','ngoquyen','kienan','haian','doson',
        'anlao','kienthuy','thuynguyen','anduong','tienlang','vinhbao','cathai',
        'bachlongvi','duongkinh'
    ]

    dictionary['Ha Noi'] = [
        'hanoi','hn','badinh','hoankiem','haibatrung','dongda','tayho',
        'caugiay','thanhxuan','hoangmai','longbien','bactuliem','thanhtri',
        'gialam','donganh','socson','hadong','thixasontay','bavi','phuctho',
        'thachthat','quocoai','chuongmy','danphuong','hoaiduc','thanhoai',
        'myduc','unghoa','thuongtin','phuxuyen','melinh','namtuliem'
    ]

    dictionary['Ha Giang'] = [
        'hagiang','dongvan','meovac','yenminh','quanba','vixuyen','bacme',
        'hoangsuphi','xinman','bacquang'
    ]

    dictionary['Cao Bang'] = [
        'caobang','thanhphocaobang','baolac','thongnong','haquang','tralinh',
        'trungkhanh','nguyenbinh','hoaan','quanguyen','thachan','halang',
        'baolam','phuchoa'
    ]

    dictionary['Lai Chau'] = [
        'laichau','thanhpholaichau','tamduong','phongtho','sinho','muongte',
        'thanuyen','tanuyen','namnhun'
    ]

    dictionary['Lao Cai'] = [
        'laocai','baothang','baoyen','batxat','bacha','thanhpholaocai',
        'muongkhuong','sapa','simacai','vanban'
    ]

    dictionary['Tuyen Quang'] = [
        'tuyenquang','thanhphotuyenquang','lambinh','nahang','chiemhoa',
        'hamyen','yenson','sonduong'
    ]

    dictionary['Lang Son'] = [
        'langson','thanhpholangson','trangdinh','binhgia','vanlang','bacson',
        'vanquan','caoloc','locbinh','chilang','dinhlap','huulung'
    ]

    dictionary['Bac Kan'] = [
        'backan','thanhphobackan','chodon','bachthong','nari','nganson',
        'babe','chomoi','pacnam'
    ]

    dictionary['Yen Bai'] = [
        'yenbai','thanhphoyenbai','thixanghialo','vanyen','yenbinh',
        'mucangchai','vanchan','tranyen','tramtau','lucyen'
    ]

    dictionary['Son La'] = [
        'sonla','thanhphosonla','quynhnhai','muongla','thuanchau','bacyen',
        'maison','yenchau','songma','mocchau','sopcop','vanho'
    ]

    dictionary['Thanh Hoa'] = [
        'thanhhoa','thanhphothanhhoa','thixabimson','thixasamson','quanhoa',
        'quanson','muonglat','bathuoc','thuongxuan','nhuxuan','nhuthanh',
        'langchanh','ngoclac','thachthanh','camthuy','thoxuan','vinhloc',
        'thieuhoa','trieuson','nongcong','dongson','hatrung','hoanghoa',
        'ngason','hauloc','quangxuong','tinhgia','yendinh'
    ]

    dictionary['Vinh Phuc'] = [
        'vinhphuc','thanhphovinhyen','tamduong','lapthach','vinhtuong',
        'yenlac','binhxuyen','songlo','thixaphucyen','tamdao'
    ]

    dictionary['Quang Ninh'] = [
        'quangninh','halong','mongcai','campha','uongbi','binhlieu',
        'damha','haiha','tienyen','bache','thixadongtrieu','thixaquangyen',
        'hoanhbo','vandon','coto'
    ]

    dictionary['Bac Giang'] = [
        'bacgiang','yenthe','lucngan','sondong','lucnam','tanyen',
        'hiephoa','langgiang','vietyen','yendung'
    ]

    dictionary['Bac Ninh'] = [
        'bacninh','yenphong','quevo','tiendu','thixatuson','thuanthanh',
        'giabinh','luongtai'
    ]

    dictionary['Hai Duong'] = [
        'haiduong','thanhphohaiduong','thixachilinh','namsach','kinhmon',
        'gialoc','tuky','thanhmien','ninhgiang','camgiang','thanhha',
        'kimthanh','binhgiang'
    ]

    dictionary['Hung Yen'] = [
        'hungyen','thanhphohungyen','kimdong','anthi','khoaichau','yenmy',
        'tienlu','phucu','myhao','vanlam','vangiang'
    ]

    dictionary['Ha Nam'] = [
        'hanam','duytien','kimbang','lynhan','thanhliem','binhluc'
    ]

    dictionary['Nam Dinh'] = [
        'namdinh','myloc','xuantruong','giaothuy','yyen','vuban','namtruc',
        'trucninh','nghiahung','haihau'
    ]

    dictionary['Thai Binh'] = [
        'thaibinh','quynhphu','hungha','donghung','vuthu','kienxuong',
        'tienhai','thaithuy'
    ]

    dictionary['Ninh Binh'] = [
        'ninhbinh','tamdiep','nhoquan','giavien','hoalu','yenmo',
        'kimson','yenkhanh'
    ]

    dictionary['Nghe An'] = [
        'nghean','tpvinh','thanhphovinh','thixacualo','quychau','quyhop',
        'nghiadan','quynhluu','kyson','tuongduong','concuong','tanky',
        'yenthanh','dienchau','anhson','doluong','thanhchuong','nghiloc',
        'namdan','hungnguyen','quephong','thixathaihoa','thixahoangmai'
    ]

    dictionary['Ha Tinh'] = [
        'hatinh','thanhphohatinh','thixahonglinh','huongson','ductho',
        'nghixuan','canloc','huongkhe','thachha','camxuyen','kyanh',
        'vuquang','locha','thixakyanh'
    ]

    dictionary['Quang Binh'] = [
        'quangbinh','botrach','donghoi','tuyenhoa','minhhoa','quangtrach',
        'quangninh','lethuy','thixabadon'
    ]

    dictionary['Quang Tri'] = [
        'quangtri','thanhphodongha','thixaquangtri','vinhlinh','giolinh',
        'camlo','trieuphong','hailang','huonghoa','dakrong','daoconco'
    ]

    dictionary['Quang Nam'] = [
        'quangnam','hoian','dailoc','dienban','duyxuyen','thanhphotamky',
        'thixadienban','queson','hiepduc','thangbinh','nuithanh','tienphuoc',
        'bactramy','donggiang','namgiang','phuocson','namtramy','taygiang',
        'phuninh','nongson'
    ]

    dictionary['Da Nang'] = [
        'danang','haichau','camle','lienchieu','nguhanhson','thanhkhe',
        'sontra','hoavang','hoangsa'
    ]

    dictionary['Quang Ngai'] = [
        'quangngai','binhson','sontinh','thanhphoquangngai','tunghia',
        'nghiahanh','moduc','ducpho','bato','minhlong','sonha','sontay',
        'trabong','taytra','lyson'
    ]

    dictionary['Kon Tum'] = [
        'kontum','dakglei','ngochoi','dakto','sathay','konplong',
        'dakha','konray','tumorong','iahdrai'
    ]

    dictionary['Binh Dinh'] = [
        'binhdinh','quynhon','quinhon','hoaian','hoainhon','phumy',
        'phucat','vinhthanh','tayson','vancanh','thixaannhon','tuyphuoc'
    ]

    dictionary['Gia Lai'] = [
        'gialai','pleiku','chupah','mangyang','kbang','thixaankhe','kongchro',
        'ducco','chuprong','chuse','thixaayunpa','krongpa','iagrai','dakdoa',
        'iapa','dakpo','phuthien','chupuh'
    ]

    dictionary['Phu Yen'] = [
        'phuyen','tuyhoa','dongxuan','thixasongcau','tuyan','sonhoa',
        'songhinh','donghoa','phuhoa','tayhoa'
    ]

    dictionary['Dak Lak'] = [
        'daklak','buonmethuot','buonmethuoc','eahleo','krongbuk','krongnang',
        'easup','cumgar','krongpac','eakar','mdrak','krongana','krongbong',
        'lak','buondon','cukuin','thixabuonho'
    ]

    dictionary['Khanh Hoa'] = [
        'khanhhoa','nhatrang','camranh','dienkhanh','vanninh','thixaninhhoa',
        'khanhvinh','khanhson','daotruongsa','camlam'
    ]

    dictionary['Vung Tau'] = [
        'vungtau','brvt','baria','tpvt','xuyenmoc','longdien','condao',
        'tanthanh','chauduc','datdo'
    ]

    dictionary['Lam Dong'] = [
        'lamdong','dalat','baoloc','ductrong','dilinh','donduong','lacduong',
        'dahuoai','dateh','cattien','lamha','baolam','damrong'
    ]

    dictionary['Binh Duong'] = [
        'binhduong','dian','thudaumot','thudau1','thixabencat','thixatanuyen',
        'thixathuanan','thixadian','phugiao','dautieng','bactanuyen','baubang',
    ]

    dictionary['Ninh Thuan'] = [
        'ninhthuan','thapcham','phanrang','ninhson','ninhhai',
        'ninhphuoc','bacai','thuanbac','thuannam'
    ]

    dictionary['Tay Ninh'] = [
        'tayninh','tanbien','tanchau','duongminhchau','chauthanh',
        'hoathanh','bencau','godau','trangbang'
    ]

    dictionary['Binh Thuan'] = [
        'binhthuan','phanthiet','tuyphong','bacbinh','hamthuanbac',
        'hamthuannam','hamtan','duclinh','tanhlinh','daophuquy','thixalagi'
    ]

    dictionary['Dong Thap'] = [
        'dongthap','chauthanh','laivung','lapvo','thanhphosadec','thanhphocaolanh',
        'caolanh','thapmuoi','tamnong','thanhbinh','thixahongngu','hongngu','tanhong'
    ]

    dictionary['Long An'] = [
        'longan','thanhphotanan','vinhhung','mochoa','tanthanh','duchue',
        'duchoa','benluc','thuthua','chauthanh','tantru','canduoc','cangiuoc',
        'thixakientuong'
    ]

    dictionary['An Giang'] = [
        'angiang','longxuyen','chaudoc','anphu','thixatanchau','tinhbien',
        'triton','chauphu','chomoi','chauthanh','thoaison'
    ]

    dictionary['Binh Phuoc'] = [
        'binhphuoc','anlocbinhlong','dongxoai','dongphu','chonthanh',
        'thixabinhlong','locninh','budop','thixaphuoclong','budang','honquan',
        'bugiamap','phurieng'
    ]

    dictionary['Tien Giang'] = [
        'tiengiang','mytho','gocong','caibe','cailay','chauthanh','chogao',
        'gocongdong','tanphuoc','tanphudong'
    ]

    dictionary['Kien Giang'] = [
        'kiengiang','kienggiang','kienluong','kiengluong','hatien','thanhphorachgia',
        'hondat','tanhiep','chauthanh','giongrieng','goquao','anbien','anminh',
        'vinhthuan','phuquoc','kienhai','uminhthuong','giangthanh'
    ]

    dictionary['Can Tho'] = [
        'cantho','cairang','ninhkieu','tpct','binhthuy','omon','phongdien',
        'codo','vinhthanh','thotnot','thoilai'
    ]

    dictionary['Ben Tre'] = [
        'bentre','thanhphobentre','chauthanh','cholach','mocaybac','giongtrom',
        'binhdai','batri','thanhphu','mocaynam'
    ]

    dictionary['Vinh Long'] = [
        'vinhlong','longho','mangthit','thixabinhminh','tambinh','traon',
        'vungliem'
    ]

    dictionary['Tra Vinh'] = [
        'travinh','tracu','canglong','cauke','tieucan','chauthanh','caungang',
        'duyenhai','thixaduyenhai'
    ]

    dictionary['Soc Trang'] = [
        'soctrang','kesach','mytu','myxuyen','longphu','thixavinhchau','culaodung',
        'thixanganam','chauthanh','trande'
    ]

    dictionary['Bac Lieu'] = [
        'baclieu',' vinhloiminhhai','vinhloi','hongdan','thixagiarai','phuoclong',
        'donghai'
    ]

    dictionary['Ca Mau'] = [
        'camau','thoibinh','uminh','tranvanthoi','cainuoc','damdoi','ngochien',
        'namcan','phutan'
    ]

    dictionary['Dien Bien'] = [
        'tinhdienbien','tpdienbienphu','thanhphodienbienphu','thixamuonglay',
        'tuangiao','muongcha','tuachua','dienbiendong','muongnhe',
        'muongang','nampo'
    ]

    dictionary['Dak Nong'] = [
        'daknong','dacnong','thixagianghia','dakrlap','dakmil','cujut',
        'daksong','krongno','dakglong','tuyduc'
    ]

    dictionary['Hau Giang'] = [
        'haugiang','thanhphovithanh','vithuy','longmy','phunghiep',
        'chauthanh','chauthanha','thixangabay','thixalongmy'
    ]

    dictionary['Dong Nai'] = [
        'dongnai','longkhanh','dnai','trangbom','bienhoa','nhontrach',
        'vinhcuu','dinhquan','thongnhat','xuanloc','longthanh',
        'cammy'
    ]

    dictionary['Hue'] = [
        'hue','phongdien','quangdien','thixahuongtra','phuvang',
        'thixahuongthuy','phuloc','namdong','aluoi'
    ]

    dictionary['Hoa Binh'] = [
        'hoabinh','dabac','maichau','tanlac','lacson','kyson','luongson',
        'kimboi','lacthuy','yenthuy','caophong'
    ]

    dictionary['Phu Tho'] = [
        'phutho','thanhphoviettri','thixaphutho','doanhung','thanhba',
        'hahoa','camkhe','yenlap','thanhson','phuninh','lamthao','tamnong',
        'thanhthuy','tanson'
    ]

    dictionary['Thai Binh'] = [
        'thaibinh','quynhphu','hungha','donghung','vuthu','kienxuong',
        'tienhai','thaithuy'
    ]

    dictionary['Thai Nguyen'] = [
        'thainguyen','thanhphothainguyen','thanhphosongcong','dinhhoa',
        'phuluong','vonhai','daitu','donghy','phubinh','thixaphoyen'
    ]

    dictionary['Taiwan'] = [
        'taiwan','taipei','taoyuan','kaohsiung','keelung','roc'
    ]

    dictionary['Canada'] = [
        'canada'
    ]

    dictionary['USA'] = [
        'usa','cavecreek','sandiego'
    ]

    dictionary['France'] = [
        'france'
    ]

    dictionary['Philippines'] = [
        'philippines'
    ]

    dictionary['China'] = [
        'trungquoc','china','beijing','shanghai'
    ]

    dictionary['Malaysia'] = [
        'malaysia'
    ]

    dictionary['Singapore'] = [
        'singapore'
    ]

    dictionary['Korea'] = [
        'korea'
    ]

    dictionary['Japan'] = [
        'japan','shinoharahigasi'
    ]

    dictionary['Hong Kong'] = [
        'hongkong'
    ]

    dictionary['Thailand'] = [
        'thailand'
    ]

    dictionary['Australia'] = [
        'australia','mountelizavic'
    ]

    dictionary['South Africa'] = [
        'southafrica'
    ]

    dictionary['New Zealand'] = [
        'newzealand'
    ]

    dictionary['Germany'] = [
        'germany'
    ]

    dictionary['Austria'] = [
        'austria'
    ]

    dictionary['British Virgin Islands'] = [
        'britishvirgin'
    ]

    dictionary['Belgium'] = [
        'belgium'
    ]

    dictionary['Cayman Islands'] = [
        'cayamanislands'
    ]

    def classify_province(x):
        name = 'Unknown'
        for province in dictionary:
            for local in dictionary[province]:
                if local in x:
                    name = province
        return name

    result = handle_missing_value(raw_address)
    result = highest_level(result)
    result = lowercase(result)
    result = remove_accent(result)
    result = remove_separator(result)
    result = classify_province(result)

    return result


def convertNaTtoSpaceString(d):
    """
    Due to Pandas's limitation that make it unable to work well with NaT,
    this function serves to replace NaT with ' ' (space string)
  
    :param d: datetime object
    """

    if pd.isnull(d):
        d = ' '
    return d


def iterable_to_sqlstring(itr):
    """
    This function convert an Python's iterable to a SQL string
    which then can be inserted to SQL code
  
    :param itr: any 1-dimension Python's interable (list, tuple, set, pd.Series)
    """

    sqlstr = ','.join(itr)
    sqlstr = sqlstr.replace(",","','")
    sqlstr = "('"+sqlstr+"')"

    return sqlstr


def month_mapper(x):
    """
    This function converts month string.
  
    Examples
    --------
  
    >>> month_mapper('01')
    'Jan'
  
    >>> month_mapper('10')
    'Oct'
  
    >>> month_mapper('Jun')
    '06'
  
    >>> month_mapper('Dec')
    '12'
  
    >>> month_mapper(5)
    'May'
  
    >>> month_mapper(12)
    'Dec'
  
    :param x: input
  
    """

    mapper = {
        'Jan':'01',
        'Feb':'02',
        'Mar':'03',
        'Apr':'04',
        'May':'05',
        'Jun':'06',
        'Jul':'07',
        'Aug':'08',
        'Sep':'09',
        'Oct':'10',
        'Nov':'11',
        'Dec':'12',
    }

    if isinstance(x,str):
        if len(x)==3:
            pass
        elif len(x)==2:
            mapper = {v:k for k,v in mapper.items()}
        else:
            raise ValueError('Invalid Input')
        result = mapper[x]

    elif isinstance(x,int):
        if x>=10:
            x = str(x)
        else:
            x = '0'+str(x)
        result = month_mapper(x)

    else:
        raise ValueError('Invalid Input')

    return result


def first(iterable,condition):
    """
    This function return the first element matches a condition in an iterable.
    If none of the element matchs the condition, return None

    :param iterable: iterable to look up
    :param condition: a function that return a boolean
    """

    for value in iterable:
        if condition(value):
            return value