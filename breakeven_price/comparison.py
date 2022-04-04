import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

pd.set_option("display.max_rows",None,"display.max_column",None,'display.width',None,'display.max_colwidth',20)
df_result = pd.read_excel(
    r'C:\Users\Admin\Desktop\PhuHung\BreakevenPrice\MonteCarloSimulation\ChoVay-Resualt2\ResultSummary.xlsm',
    sheet_name='Sheet1',usecols=['Ticker','Current Breakeven Price','Model\'s Breakeven Price'])
df_result['Current Breakeven Price'] = df_result['Current Breakeven Price'].astype(int)
df_result['Model\'s Breakeven Price'] = df_result['Model\'s Breakeven Price'].astype(int)
df_result['Diff'] = df_result['Model\'s Breakeven Price']-df_result['Current Breakeven Price']
df_result['% Diff'] = pd.Series('{0:.2f}%'.format((df_result.iloc[i,2]/df_result.iloc[i,1]-1)*100)
                                for i in range(df_result['Diff'].count()))
number_of_tickers = df_result['Ticker'].count()
for i in range(300-number_of_tickers):
    df_result = df_result.append(pd.DataFrame({'Ticker':'###','Current Breakeven Price':0,'Model\'s Breakeven Price':0,
                                               'Diff':0,'% Diff':0},index=[number_of_tickers+i+1]))

print(df_result)


def reformat_large_tick_values(tick_val,pos):
    """
    Turns large tick values (in the billions, millions and thousands) such as 4500 into 4.5K and also appropriately turns 4000 into 4K (no zero after the decimal)
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


def autolabel(rects):
    """Attach a text label above each bar in *rects*, displaying its height."""
    for rect in rects:
        height = rect.get_height()
        ax[i,j].annotate('{}'.format(height),
                         xy=(rect.get_x()+rect.get_width()/2,height),
                         xytext=(0,3),  # 3 points vertical offset
                         textcoords="offset points",
                         ha='center',va='bottom',rotation=90,fontsize=5)


grid_kws = {"height_ratios":(0.2,0.2,0.2,0.2),"hspace":0.3,"wspace":0.25}
fig,ax = plt.subplots(4,5,gridspec_kw=grid_kws)
fig.suptitle('Comparison between Current Price vs. Model\'s Price',fontsize=14,weight='bold')
fig.subplots_adjust(left=0.05,right=0.95,bottom=0.05,top=0.9,wspace=0.2,hspace=0.4)

labels = list()
for i in range(4):
    for j in range(5):
        labels.append(df_result['Ticker'].iloc[75*i+15*j:75*i+15*(j+1)].to_list())
print(labels)

for i in range(4):
    for j in range(5):
        width = 0.35
        l = np.arange(15)  # the label locations
        rects1 = ax[i,j].bar(l-width/2,df_result['Current Breakeven Price'].
                             iloc[75*i+15*j:75*i+15*(j+1)],width=width,label='Current Breakeven Price')
        rects2 = ax[i,j].bar(l+width/2,df_result['Model\'s Breakeven Price'].
                             iloc[75*i+15*j:75*i+15*(j+1)],width=width,label='Model\'s Breakeven Price')
        ax[i,j].yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(reformat_large_tick_values))
        plt.setp(ax[i,j].xaxis.get_majorticklabels(),rotation=90)
        ax[i,j].grid(linestyle='--',linewidth=0.5,alpha=0.7)
        ax[i,j].set_xticks(l)
        ax[i,j].set_xticklabels(labels[5*i+j])
        ax[i,j].set_autoscaley_on(True)
        ax[i,j].autoscale(enable=True,axis="y",tight=False)

        # autolabel(rects1)
        # autolabel(rects2)

handles,labels = ax[0,0].get_legend_handles_labels()
fig.legend(handles,labels,loc='upper left',bbox_to_anchor=(0.05,0.97))

plt.show()
