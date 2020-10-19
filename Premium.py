# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 11:41:47 2020

@author: CheikhMoctar
"""

import pandas as pd
import numpy as np
import utils
from functools import reduce

file = "../premium/Ross - Premium Allocation.xlsx"

sov = pd.read_excel('final_Overlay_07-13.xlsx')
cleaned_dfs = {}
def headerFunc(d):
    empty_cols = [col for col in d.columns if d[col].isnull().all()]
    empty_rows = d.index[d.isnull().all(1)]
    d.drop(empty_cols , axis=1,inplace=True)
    d.drop(d.index[empty_rows], inplace=True)
    d[~d.isin(['Unnamed:'])].dropna(how='all')
    if 'Unnamed:' in list(d.columns) :
         d.columns = d.iloc[0]
         d[~d.isin(['Unnamed:'])].dropna(how='all')
         cleaned_dfs.append(d)
         return d
    elif len(list(d.columns))>0:
        pass
    else:
        return d
       
sheets = {'BRE Hotels Allocation':12, 'Mfg Homes Allocation':1, 'GLP Allocation':17, 'Office Retail Allocation':6, 'Link Allocation':15, 
          'Space Center Allocation':0, 'LivCor Allocation':1}
#df = pd.read_excel(file, None);



# Data Cleaning and Address normalization functions
def replace_last(source_string, replace_what, replace_with):
        head, _sep, tail = source_string.rpartition(replace_what)
        return head + replace_with + tail
    
def replace_second(source_string, dict_):
    s = source_string.split(' ')
    if len(s) >1 :
        if s[1].replace('.', '').lower() in dict_.keys():
            s[1] = dict_[s[1].replace('.', '')]
            return ' '.join(s)
        else:    
            return source_string
    else:
        return source_string

#def similar(a, b):
    #return SequenceMatcher(None, a, b).ratio()
    

""" Retained columns from each file
====================================================================================================================="""

Us_States = utils.us_state_abbrev
st_dict = utils.st_dict
dir_dict = utils.dir_dict


sov['State'] = sov['State'].apply(lambda x: Us_States[x] if x in Us_States.keys()  else Us_States[x.capitalize()])
Add_cols = ['Address', 'City', 'State', 'Zip']
""" Adding geolocation into the MDS
====================================================================================================================="""
def lower_No_spaces(df, col):
      df[col] = [(' '.join(str(d).split())).lower().strip() for d in df[col]]
      df[col] = [d.replace(' ', '') for d in df[col]]
      return df[col].astype(str)

def Address_Normal(df):
    #k1 = df.columns[4]
    
    #df['State'] = df['State'].apply(lambda x: Us_States[x])
    df['Address'] = [(' '.join(str(d).split())).lower().strip() for d in df['Address']]
    df['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df['Address']]
    df['st_add'] = [replace_second(d,dir_dict) for d in df['st_add']]
    df['Add'] = [d.replace(' ', '') for d in df['st_add']] + lower_No_spaces(df, 'City') + lower_No_spaces(df, 'State')
   
    
    return df

def merging(df1, df2):
    df = pd.merge(Address_Normal(df1), Address_Normal(df2), on = 'Add', how='outer', suffixes=('_x', '_y'), indicator=True)
    return df

xl = pd.ExcelFile(file)
for sh in xl.sheet_names:
    if sh in sheets.keys():
        var = sh.split(' ')[0]
        df = xl.parse(sh, skiprows=sheets[sh])
        empty_rows = df.index[df.isnull().all(1)]
        df.drop(df.index[empty_rows], inplace=True)
        cleaned_dfs[var] = df


overlay = {k: merging(sov, df) for k, df in cleaned_dfs.items()}
found = {k:len(df[df['_merge'] == 'both']) for k, df in overlay.items()}
left_only = {k:len(df[df['_merge'] == 'left_only']) for k, df in overlay.items()}
right_only = {k:len(df[df['_merge'] == 'right_only']) for k, df in overlay.items()}



dfs = [Address_Normal(df) for df in cleaned_dfs.values()]
#dfs.insert(0, Address_Normal(sov))
#final = reduce(lambda df1, df2: df1.merge(df2, on="Add", how='outer', indicator=True), dfs)
#dfs = [Address_Normal(df).set_index("Add", drop=True) for df in cleaned_dfs.values()]
#final = pd.concat(dfs, axis=1, keys=range(len(dfs)), join='outer', copy=False)

merg1 = pd.merge(Address_Normal(sov), Address_Normal(dfs[0]), on = 'Add', how='outer', suffixes=('', '_z') )
merg1['Track'] = np.nan
merg1['Track'] = np.where(pd.notnull(merg1['Potential Total']), "BRE", merg1['Track'])

merg2 = pd.merge(Address_Normal(merg1), Address_Normal(dfs[1]), on = 'Add', how='outer', suffixes=('', '_z'))
merg2['Track'] = np.where(pd.notnull(merg2['Est Property Total incl Taxes']), 'Mfg', merg2['Track'])
#final.reset_index(drop=False, inplace=True)

merg3 = pd.merge(Address_Normal(merg2), Address_Normal(dfs[2]), on = 'Add', how='outer', suffixes=('', '_z'))
merg3['Track'] = np.where(pd.notnull(merg3['Premium + Fee']), 'Glp', merg3['Track'])

merg4 = pd.merge(Address_Normal(merg3), Address_Normal(dfs[3]), on = 'Add', how='outer', suffixes=('', '_z'))
merg4['Track'] = np.where(pd.notnull(merg4['Property Premium Allocation']), 'Office', merg4['Track'])


merg5 = pd.merge(Address_Normal(merg4), Address_Normal(dfs[4]), on = 'Add', how='outer', suffixes=('', '_z'))
merg5['Track'] = np.where(pd.notnull(merg5['Annual Premium']), 'Link', merg5['Track'])

merg6 = pd.merge(Address_Normal(merg5), Address_Normal(dfs[5]), on = 'Add', how='outer', suffixes=('', '_z'))
merg6['Track'] = np.where(pd.notnull(merg6['Total Property & CA EQ Premium']), 'Space', merg6['Track'])

merg7 = pd.merge(Address_Normal(merg6), Address_Normal(dfs[6]), on = 'Add', how='outer', suffixes=('', '_z'))
merg7['Track'] = np.where(pd.notnull(merg7['Total']), 'LivCor', merg7['Track'])

results = {"Total":len(merg7), "BRE_Matches": len(merg7.query("Track == 'BRE'")),"BRE_NoMatches":168 , "SOV_NoMtches":len(merg7.query("Track == 'nan'")), "SOV_Matches":(4357-1268),
           "Mfg_Matches": len(merg7.query("Track == 'Mfg'")), "Mfg_NoMatches":36, "Glp_Matches": len(merg7.query("Track == 'Glp'")), "Office_Matches": len(merg7.query("Track == 'Office'")), "Link_Matches": len(merg7.query("Track == 'Link'")), "Space_Matches": len(merg7.query("Track == 'Space'")), "LivCode_Matches": len(merg7.query("Track == 'LivCor'"))}

final_results = pd.DataFrame.from_dict(results, orient='index', dtype=None).to_excel("merging_stats.xlsx")