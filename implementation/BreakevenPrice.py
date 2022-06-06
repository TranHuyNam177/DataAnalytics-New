from request.stock import *
from breakeven_price import monte_carlo_test
from implementation import TaskMonitor


@TaskMonitor
def BreakevenPrice_5pct(  # run on Wed and Fri weekly
    tickers: list = None,
    exchanges: list = None,
) -> pd.DataFrame:

    start_time = time.time()

    destination_dir_github = join(dirname(dirname(realpath(__file__))),'breakeven_price','result_table')
    destination_dir_network = r'\\192.168.10.28\images\breakeven'
    destination_dir_rmd = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept' \
                          r'\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Gia Hoa Von'

    chart_foler = join(realpath(dirname(dirname(__file__))),'breakeven_price','result_chart')

    if exchanges=='all' or tickers=='all':
        tickers = fa.tickers(exchange='all')
    elif exchanges is not None and exchanges!='all':
        tickers = []
        for exchange in exchanges:
            tickers += fa.tickers(exchange=exchange)
    elif tickers is not None and tickers!='all':
        pass

    network_table_path = join(destination_dir_network,'tables','result.csv')
    table = pd.DataFrame(columns=['ticker','Breakeven Price'])
    table.set_index(keys=['ticker'],inplace=True)
    table.to_csv(network_table_path)

    now = dt.datetime.now()
    github_file_name = f'{now.day}.{now.month}.{now.year}.csv'
    github_table_path = join(destination_dir_github,github_file_name)
    table = pd.DataFrame(
        columns=['Ticker','0% at Risk','1% at Risk','3% at Risk','5% at Risk','Group','Breakeven Price']
    )
    table.set_index(keys=['Ticker'],inplace=True)
    table.to_csv(github_table_path)

    rmd_file_name = f'{now.day}.{now.month}.{now.year}.csv'
    rmd_table_path = join(destination_dir_rmd,rmd_file_name)
    table = pd.DataFrame(
        columns=['Ticker','Group','Breakeven Price','0% at Risk']
    )
    table.set_index(keys=['Ticker'],inplace=True)
    table.to_csv(rmd_table_path)

    for ticker in tickers:
        try:
            lv0_price,lv1_price,lv2_price,lv3_price,breakeven_price,group = monte_carlo_test.run(ticker=ticker,alpha=0.05)
            with open(github_table_path,mode='a',newline='') as github_file:
                github_writer = csv.writer(github_file,delimiter=',')
                github_writer.writerow([ticker,lv0_price,lv1_price,lv2_price,lv3_price,group,breakeven_price])
            with open(network_table_path,mode='a',newline='') as network_file:
                network_writer = csv.writer(network_file,delimiter=',')
                network_writer.writerow([ticker,breakeven_price])
            with open(rmd_table_path,mode='a',newline='') as rmd_file:
                rmd_writer = csv.writer(rmd_file,delimiter=',')
                rmd_writer.writerow([ticker,group,breakeven_price,lv0_price])
        except (ValueError,KeyError,IndexError):
            print(f'{ticker} cannot be simulated')
            print('-------')
        try:
            shutil.copy(join(chart_foler,f'{ticker}.png'),join(realpath(destination_dir_network),'charts'))
        except FileNotFoundError:
            print(f'{ticker} cannot be graphed')

        print('===========================')

    print('Finished!')
    print("Total execution time is: %s seconds"%(time.time()-start_time))


@TaskMonitor
def BreakevenPrice_2pct(  # weekly run as requested by RMD
    tickers: list = None,
    exchanges: list = None,
) \
        -> pd.DataFrame:
    start_time = time.time()

    destination_dir_github = join(dirname(dirname(realpath(__file__))),'breakeven_price','result_table')
    destination_dir_rmd = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept' \
                          r'\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Gia Hoa Von'

    if exchanges=='all' or tickers=='all':
        tickers = fa.tickers(exchange='all')
    elif exchanges is not None and exchanges!='all':
        tickers = []
        for exchange in exchanges:
            tickers += fa.tickers(exchange=exchange)
    elif tickers is not None and tickers!='all':
        pass

    now = dt.datetime.now()
    github_file_name = f'{now.day}.{now.month}.{now.year}_0.02.csv'
    github_table_path = join(destination_dir_github,github_file_name)
    table = pd.DataFrame(
        columns=['Ticker','0% at Risk','1% at Risk','3% at Risk','5% at Risk','Group','Breakeven Price'])
    table.set_index(keys=['Ticker'],inplace=True)
    table.to_csv(github_table_path)

    rmd_file_name = f'{now.day}.{now.month}.{now.year}_0.02.csv'
    rmd_table_path = join(destination_dir_rmd,rmd_file_name)
    table = pd.DataFrame(
        columns=['Ticker','Group','Breakeven Price','0% at Risk']
    )
    table.set_index(keys=['Ticker'],inplace=True)
    table.to_csv(rmd_table_path)

    for ticker in tickers:
        try:
            lv0_price,lv1_price,lv2_price,lv3_price,breakeven_price,group = monte_carlo_test.run(ticker=ticker,alpha=0.02)
            with open(github_table_path,mode='a',newline='') as github_file:
                github_writer = csv.writer(github_file,delimiter=',')
                github_writer.writerow([ticker,lv0_price,lv1_price,lv2_price,lv3_price,group,breakeven_price])
            with open(rmd_table_path,mode='a',newline='') as rmd_file:
                rmd_writer = csv.writer(rmd_file,delimiter=',')
                rmd_writer.writerow([ticker,group,breakeven_price,lv0_price])
        except (ValueError,KeyError,IndexError):
            print(f'{ticker} cannot be simulated')
            print('-------')

        print('===========================')

    print('Finished!')
    print("Total execution time is: %s seconds"%(time.time()-start_time))
