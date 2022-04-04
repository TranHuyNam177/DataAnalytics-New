import numpy as np
import matplotlib.pyplot as plt
import matplotlib.animation as ani
import scipy as sc
from scipy import stats

price0 = 10
mean = 0.001
std = 0.05
days = 22*12
bins = 100
seed = 9
alpha = 0.1

fig,(ax1,ax2) = plt.subplots(nrows=2,ncols=1,figsize=(5,8))
plt.subplots_adjust(left=0.15,right=0.95,bottom=0.07,top=0.9,hspace=0.25)
fig.suptitle("We treat trading data as a pool\nNot discrete sequence of stops",
             fontsize='xx-large',fontweight=500,
             fontfamily='Times New Roman')

date = np.array(list(range(days)))
price_ = np.array([price0])
np.random.seed(seed)
return_ = sc.stats.norm.rvs(loc=mean,scale=std,size=days-1)

for i in range(len(return_)-1):
    price_ = np.append(price_,price_[-1]*(1+return_[i]))


def buildanichart1(i=int):
    ax1.clear();
    ax2.clear()
    ax1.set_xlabel('Days')
    ax1.set_ylabel('Stock Price')
    ax2.set_xlabel('Return')
    ax2.set_ylabel('Probability Density')
    ax1.plot(date[:i],price_[:i],lw=1,color='midnightblue')
    ax1.fill_between(date[:i],price0,price_[:i],where=price_[:i]>=price0,
                     color='green',alpha=alpha)
    ax1.fill_between(date[:i],price0,price_[:i],where=price_[:i]<price0,
                     color='red',alpha=alpha)
    np.random.seed(seed)
    return__ = np.sort(sc.stats.norm.rvs(loc=mean,scale=std,size=i))
    np.random.seed(seed)
    ax2.plot(return__[:i],sc.stats.norm.pdf(return__[:i],loc=mean,scale=std),
             lw=1,color='midnightblue')
    ax2.fill_between(return__[:i],0,sc.stats.norm.pdf(return__[:i],loc=mean,scale=std),
                     color='green',alpha=alpha)


animator1 = ani.FuncAnimation(fig,buildanichart1,interval=10,repeat=True,save_count=days)
