# -*- coding: utf-8 -*-
"""
Created on Wed Nov 27 17:47:10 2019

@author: cmoctar
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Nov 22 16:46:14 2019

@author: cmoctar
"""
#Import Dependencies
import matplotlib.pyplot as plt


def plot(df, name):
    ''' This function plot bar graph,takes a data frame and the  name of the dataframe  '''
    tickers = df.columns
    percent = []
    for v in tickers:
        percent.append((len(df) - df[v].isna().sum())/len(df)*100)
    
    # Orient widths. Add labels, tick mark
    y_axis = range(len(tickers))
    
    #colors = np.array(['b']*len(df))
    fig, ax = plt.subplots(figsize=(10, 7))
    clrs = ['r' if (x < 50) else 'g' for x in percent]
    plt.barh(y_axis, percent, align = "center", color = clrs)
    plt.yticks(y_axis, tickers,  rotation="horizontal")
    plt.title(f"{name}", fontsize=14)
    ax.set_xlabel("% of recived data", fontsize=14)
    plt.grid()

    fig.savefig(f'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\Data_Stats\graphs\{name}.png')
    return  print(f'plotting -----------{name}----------')


def stats(df, tm):
    '''This function genrate some stats based on received data from the broker'''
    h = [c for c in df.columns]
    for c in tm.columns:
        if c in h:
            row_n = len(df)- df[c].isna().sum()
            tm.loc[:,c][0] = row_n
        else:
            tm.loc[:,c][0] = 0
           
    return tm