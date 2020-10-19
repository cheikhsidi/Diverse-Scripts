# -*- coding: utf-8 -*-
"""
Created on Sat Jul 18 13:53:32 2020

@author: CheikhMoctar
"""

import pandas as pd
import utils_data
import numpy as np
from difflib import get_close_matches


mds_org = pd.read_excel("../Origami_locations/Origami MDS 07-17-20.xlsx", sheet_name='MDS')
mds = pd.read_excel("LivCor Org 7.8.20.xlsx")
sov = pd.read_excel("07_16_Overlay.xlsx")
possibles = pd.read_excel('Possible_Matches (1).xlsx')

def Diff(li1, li2): 
    return (list(set(li1) - set(li2))) 

Us_States = utils_data.us_state_abbrev
st_dict = utils_data.st_dict
dir_dict = utils_data.dir_dict
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


def lower_No_spaces(df, col):
      df[col] = [(' '.join(str(d).split())).lower().strip() for d in df[col]]
      df[col] = [d.replace(' ', '') for d in df[col]]
      return df[col].astype(str)

def Address_Normal(df):
    df['Address'] = [(' '.join(str(d).split())).lower().strip() for d in df['Address']]
    df['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df['Address']]
    df['st_add'] = [replace_second(d,dir_dict) for d in df['st_add']]
    df['Add'] = [d.replace(' ', '') for d in df['st_add']] + lower_No_spaces(df, 'City') + lower_No_spaces(df, 'State')
    return df

sov['State'] = sov['State'].apply(lambda x: Us_States[x] if x in Us_States.keys()  else Us_States[x.capitalize()])
cols = ['Address', 'City', 'State', 'Zip']

mds1 = Address_Normal(mds)
sov = Address_Normal(sov)
print(len(sov))
sov = sov.drop_duplicates(subset='Add', keep='first')
sov_pr = sov.drop_duplicates(subset='Property', keep='first')

"""finding dups"""
# dups = pd.merge(mds1, mds, on=['Location','Address'], how='outer', indicator=True)\
#        .query("_merge != 'both'")\
#        .drop('_merge', axis=1)\
#        .reset_index(drop=True)
poss = pd.merge(possibles, mds_org[['MDS ID', 'Name']],  how='left', left_on= 'possibleMatches_name', right_on='Name', validate = 'm:1', suffixes=('', '_poss')) 
print(mds['Add'])
# print(len(mds))
"""Merge on Address------------------------------------------------------------------------------------------------------------"""
new_df = pd.merge( mds, sov[['Add', 'Business_ID']],  how='left', on='Add' , validate = 'm:1', suffixes=('', '_mds'))\
.merge(sov_pr[['Property', 'Business_ID']],  how='left', left_on='Location', right_on='Property', validate = 'm:1', suffixes=('', '_mds2'))\
.merge(mds_org[['MDS ID', 'Name']],  how='left', left_on='Location', right_on='Name', validate = 'm:1', suffixes=('', '_org'))\
.merge(poss[['MDS ID', 'Location', 'Address']], how='left', on=['Location', 'Address'], validate = 'm:1', suffixes=('', '_2'))



new_df['MDS ID'] = np.where(pd.notnull(new_df['MDS ID_2']), new_df['MDS ID_2'], new_df['MDS ID'])
new_df['MDS ID'] = np.where((pd.notnull(new_df['Business_ID_mds2']) & new_df['MDS ID'].isna()), new_df['Business_ID_mds2'], new_df['MDS ID'])
new_df['MDS ID'] = np.where((pd.notnull(new_df['Business_ID']) & new_df['MDS ID'].isna()), new_df['Business_ID'], new_df['MDS ID'])


new_df.to_excel('LivCor_finalOverlay.xlsx', index=False)
# test = new_df[(new_df['Location_mds'].notna()) | (new_df['Location_mds2'].notna())]
# print(len(new_df[~new_df['BU #/Location Code'].isna()])) 
# # print(len(new_df[~new_df['MDS Property ID'].isna()]))

# """Merge on property------------------------------------------------------------------------------------------------------------"""
# new_df2 = pd.merge( sov,  mds,  how='left', left_on='Property', right_on='Location', validate = 'm:1', suffixes=('', '_mds2'))
# print(len(new_df2[new_df2['BU #/Location Code'].notna()]))


# add = test[~test['Location_mds'].isna()]['Location_mds'].to_list()
# md = test[~test['Location_mds2'].isna()]['Location_mds2'].to_list()
                
          
# NoMatches = mds.query(f'{add} not in Location and {md} not in Location')
# Matches = mds.query(f'Location == {add} or Location == {md}')

# """Possible Matches------------------------------------------------------------------------"""
# # sov['address'] = sov['Address'] +' '+ sov['City']
# # NoMatches['address'] = NoMatches['Address']+ ' ' + NoMatches['City']
# # Addresses = sov['address'].to_list()
# # props = sov['Property'].to_list()
# # props = [str(d) for d in props]
# def getMatches(text, Addresses, count, precision):
#     text = str(text)
#     if get_close_matches(text, Addresses):
#         return '// '.join(get_close_matches(text, Addresses, count, precision))
#     else:
#         return '// '.join(get_close_matches(text, Addresses, count, 0.3))

# # NoMatches['possibleMatches_add']  = NoMatches['address'].apply(lambda text: getMatches(text, Addresses, 2, 0.6))
# # NoMatches['possibleMatches_prop']  = NoMatches['Location'].apply(lambda text: getMatches(text, props, 2, 0.6))
# # NoMatches.to_excel('Possible_Matches.xlsx', index=False)

# """ Mergin on Location from MDS--------------------------------------------------------"""

# df2_mds = pd.merge( NoMatches,  mds_org[['MDS ID', 'Name']],  how='left', left_on='Location', right_on='Name', validate = 'm:1', suffixes=('', '_org'))
# print(len(df2_mds[df2_mds['MDS ID'].isna()]))
# loc = df2_mds[~df2_mds['MDS ID'].isna()]['Location'].to_list()
# NoMatches2 = mds.query(f'{add} not in Location and {md} not in Location and {loc} not in Location')


# """Possible Matches------------------------------------------------------------------------"""
# # mds_org['address'] = mds_org['Address'] +' '+ mds_org['City']
# # NoMatches['address'] = NoMatches['Address']+ ' ' + NoMatches['City']
# Addresses = mds_org['Name'].to_list()
# # props = mds_org['Name'].to_list()
# # props = [str(d) for d in props]

# NoMatches2['possibleMatches_name']  = NoMatches2['Location'].apply(lambda text: getMatches(text, Addresses, 2, 0.6))
# # NoMatches2['possibleMatches_prop']  = NoMatches2['Location'].apply(lambda text: getMatches(text, props, 2, 0.6))
# NoMatches2.to_excel('Possible_Matches.xlsx', index=False)
