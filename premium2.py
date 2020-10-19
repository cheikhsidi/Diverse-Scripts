# -*- coding: utf-8 -*-
"""
Created on Wed Jul 15 14:46:49 2020

@author: CheikhMoctar
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 11:41:47 2020

@author: CheikhMoctar
"""

import pandas as pd
import numpy as np
import utils_data
from functools import reduce
from difflib import get_close_matches, SequenceMatcher
from fuzzywuzzy import fuzz
from fuzzywuzzy import process, utils

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
       
sheets = {'GLP Allocation':17, 'Office Retail Allocation':6, 'Link Allocation':15, 
          'Space Center Allocation':0}
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

Us_States = utils_data.us_state_abbrev
st_dict = utils_data.st_dict
dir_dict = utils_data.dir_dict


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
sov = sov.drop_duplicates(subset=['Address', 'City', 'State'], keep= 'first')
def merging(df1, df2):
    df1['New_Fund'] = np.nan
    df = pd.merge(Address_Normal(df1), df2, on = 'Add', how='left', suffixes=('', '_y'), validate='m:1')
    df['New_Fund'] = np.where(pd.notnull(df['Fund']), df['Fund'], df['New_Fund'])
    #df = df[list(df1.columns)]
    return df

xl = pd.ExcelFile(file)
for sh in xl.sheet_names:
    if sh in sheets.keys():
        var = sh.split(' ')[0]
        df = xl.parse(sh, skiprows=sheets[sh])
        empty_rows = df.index[df.isnull().all(1)]
        df.drop(df.index[empty_rows], inplace=True)
        cleaned_dfs[var] = df

sov = Address_Normal(sov).drop_duplicates(subset=['Add'], keep= 'first')
overlay = {k: merging(df, sov) for k, df in cleaned_dfs.items()}
found = {k:len(df[df['New_Fund'].isna()]) for k, df in overlay.items()}



#left_only = {k:len(df[df['New_Fund'].isna()]) for k, df in overlay.items()}
#right_only = {k:len(df[df['New_Fund'].isna()]) for k, df in overlay.items()}

"""=================================Possibel Matches ======================================================"""

def getMatches(text, Addresses, count, precision):
    text = str(text)
    if get_close_matches(text, Addresses):
        return '// '.join(get_close_matches(text, Addresses, count, precision))
    else:
        return ''
   
def possibleMatches(count, precision, df, df2):
    df['address'] = df['Address'] +' '+ df['City']
    #df2.index +=2
    #df2.reset_index(inplace=True)
    df2['address'] =  df2['Address'] +' '+ df2['City']
    Addresses = df2['address'].to_list()
    Addresses = [str(d) for d in Addresses]
    df['possibleMatches'] = df['address'].apply(lambda text: getMatches(text, Addresses, count, precision))
    #df.to_excel('test.xlsx', index=False, engine='xlsxwriter')
    return df

sov = Address_Normal(sov)
sov.index +=2
sov.reset_index(inplace=True)
sov['address'] = sov['Address'] + ' ' + sov['City'] + ' ' +  sov['index'].astype(str)
Addresses = sov['address'].to_list()
#Addresses = [str(d) for d in Addresses]






# sov = Address_Normal(sov)
# org_list = sov['Add'].to_list()
# # processed_orgs = {org: utils.default_process(org) for org in org_list}
# df = overlay['GLP']
# Add = df['Add']
# for (i, (query, processed_query)) in enumerate(Add.items()):
#     match = process.extract(processed_query, org_list, processor=None, limit=2,scorer=fuzz.token_sort_ratio)
#     print (match)
#     if match:
#         df.loc[i, 'fuzzy_match'] = str(match)
#         df.loc[i, 'indx'] = sov
#     else:
#         df.loc[i, 'fuzzy_match'] = ''
#         #df.loc[i, 'fuzzy_match_score'] = match[1]

    # if match:
    #     df.loc[i, 'fuzzy_match'] = str(match)
    #     df.loc[i, 'indx'] = sov
    # else:
    #     df.loc[i, 'fuzzy_match'] = ''
    #     #df.loc[i, 'fuzzy_match_score'] = match[1]




writer = pd.ExcelWriter('Output3.xlsx')

for k, df in overlay.items():  
    df.to_excel(writer, k, index=False)
writer.save()


writer = pd.ExcelWriter('Output2.xlsx')

for k, df in overlay.items():  
    df = df[df['Fund'].isna()]
    df['address'] = df['Address'] +' '+ df['City']
    df['possibleMatches'] = df['address'].apply(lambda text: getMatches(text, Addresses, 3, 0.1))
    df.to_excel(writer, k, index=False)
writer.save()


# dfs = [Address_Normal(df) for df in cleaned_dfs.values()]
# dfs.insert(0, Address_Normal(sov))
# final = reduce(lambda df1, df2: df1.merge(df2, on="Add", how='outer', indicator=True), dfs)
# dfs = [Address_Normal(df).set_index("Add", drop=True) for df in cleaned_dfs.values()]
# final = pd.concat(dfs, axis=1, keys=range(len(dfs)), join='outer', copy=False)

