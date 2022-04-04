from implementation import TaskMonitor
from request.stock import *


@TaskMonitor
def CreditRating(
    standard:str='bics',
    level:int=1
):
    destination_path = r'\\192.168.10.28\images\creditrating'
    chart_path = os.path.join(destination_path,'charts')
    table_path = os.path.join(destination_path,'tables')

    # POST CREDIT RATING
    rating_path = join(realpath(dirname(dirname(__file__))),'credit_rating','result')
    rating_files = [file for file in listdir(rating_path) if file.startswith('result') and 'gen' not in file]
    for file in rating_files:
        shutil.copy(join(rating_path,file),join(table_path,'rating'))
    # adjust result_table_gen
    gen_table = pd.read_csv(join(rating_path,'result_table_gen.csv'),index_col=['ticker'])
    gen_table = gen_table.loc[gen_table['level']==f'{standard}_l{level}']
    gen_table.drop(['standard','level','industry'],axis=1,inplace=True)
    gen_table.to_csv(join(table_path,'rating','result_table_gen.csv'))
    # merge to ready-to-use result table
    bank_table = pd.read_csv(join(rating_path,'result_table_bank.csv'),index_col=['ticker'])
    ins_table = pd.read_csv(join(rating_path,'result_table_ins.csv'),index_col=['ticker'])
    sec_table = pd.read_csv(join(rating_path,'result_table_sec.csv'),index_col=['ticker'])
    result_table = pd.concat([gen_table,bank_table,ins_table,sec_table])
    result_table.drop(['BM_'],inplace=True)
    result_table.to_csv(join(table_path,'rating','result_summary.csv'))

    component_files = [file for file in listdir(join(rating_path,'Compare with Industry'))]
    for file in component_files:
        shutil.copy(join(rating_path,'Compare with Industry',file),join(chart_path,'component'))

    result_files = [file for file in listdir(join(rating_path,'Result'))]
    for file in result_files:
        shutil.copy(join(rating_path,'Result',file),join(chart_path,'rating'))

    # CREATE IMAGE
    rating_tickers = {name[:3] for name in result_files}
    component_tickers = {name[:3] for name in component_files}

    # intersect
    tickers = rating_tickers&component_tickers
    k_rating = 1.4
    for ticker in tickers:
        rating_image = Image.open(join(chart_path,'rating',f'{ticker}_result.png'))
        rating_image = rating_image.resize(
            (int(rating_image.width*k_rating),int(rating_image.height*k_rating))
        )
        component_image = Image.open(join(chart_path,'component',f'{ticker}_compare_industry.png'))
        component_image = component_image.resize(
            (int(rating_image.width),int(rating_image.width*component_image.height/component_image.width))
        )
        result_image = Image.new(
            'RGB',
            (rating_image.width,rating_image.height+component_image.height),
            (255,255,255)
        )
        result_image.paste(rating_image)
        result_image.paste(component_image,(0,rating_image.height))
        result_image.save(join(chart_path,'result',f'{ticker}_result.png'))
