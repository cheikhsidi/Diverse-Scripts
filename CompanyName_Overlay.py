
import pandas as pd
import utils_data
import numpy as np
#from difflib import get_close_matches


Origami = pd.read_excel("../../../Downloads/Origami Locations 10-5.xlsx", skiprows=2)
BRE = pd.read_excel("../../../Downloads/BRE Hotels Resort Property SOV -10-1-20.xlsx")

# mds = pd.read_excel("LivCor Org 7.8.20.xlsx")
sov = pd.read_excel("07_18_Overlay.xlsx").sort_values('Address')
# possibles = pd.read_excel('Possible_Matches (1).xlsx')

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
    df['Add'] = [d.replace(' ', '') for d in df['st_add']]
    return df

# sov['State'] = sov['State'].apply(lambda x: Us_States[x] if x in Us_States.keys()  else Us_States[x.capitalize()])
# cols = ['Address', 'City', 'State', 'Zip']

# mds1 = Address_Normal(mds)
# sov = Address_Normal(sov)
# print(len(sov))
# sov = sov.drop_duplicates(subset='Add', keep='first')
"""finding dups"""
# dups = pd.merge(mds1, mds, on=['Location','Address'], how='outer', indicator=True)\
#        .query("_merge != 'both'")\
#        .drop('_merge', axis=1)\
#        .reset_index(drop=True)

merge = pd.merge( mds, Company[['CompanyID', 'AddressLine1', 'City', 'State']],  how='left', left_on='Workday Company ID', right_on='CompanyID', validate = 'm:1', suffixes=('', '_mds')).sort_values('AddressLine1')
print(len(merge[merge['CompanyID'].isna()]))

merge = merge.drop_duplicates(subset='AddressLine1', keep=False)
merge.rename(columns={'AddressLine1': 'Address'}, inplace=True)
#df = merge.groupby(['Address', 'City', 'State']).size().reset_index(name='count').query("count > 1")
#len(df)
merge2 = pd.merge( Address_Normal(sov), Address_Normal(merge)[['MDS ID','Address', 'Add']],  how='left', on='Add', validate = 'm:1', suffixes=('', '_mds'))
# merge3 = merge2[['Address', 'Address_mds']].sort_values('Address')
merge2['Business_ID'] = np.where(pd.notnull(merge2['MDS ID']), merge2['MDS ID'], merge2['Business_ID'])
merge4= merge2[~merge2['MDS ID'].isna()].sort_values('Address')

merge2.to_excel('07_19_Overlay.xlsx', index=True)
# poss = pd.merge(possibles, mds_org[['MDS ID', 'Name']],  how='left', left_on= 'possibleMatches_name', right_on='Name', validate = 'm:1', suffixes=('', '_poss')) 
# print(mds['Add'])
# # print(len(mds))
# """Merge on Address------------------------------------------------------------------------------------------------------------"""
# new_df = pd.merge( mds, sov[['Add', 'Business_ID']],  how='left', on='Add' , validate = 'm:1', suffixes=('', '_mds'))\
# .merge(sov_pr[['Property', 'Business_ID']],  how='left', left_on='Location', right_on='Property', validate = 'm:1', suffixes=('', '_mds2'))\
# .merge(mds_org[['MDS ID', 'Name']],  how='left', left_on='Location', right_on='Name', validate = 'm:1', suffixes=('', '_org'))\
# .merge(poss[['MDS ID', 'Location', 'Address']], how='left', on=['Location', 'Address'], validate = 'm:1', suffixes=('', '_2'))



#merge2['MDS ID'] = np.where((pd.notnull(merge2['Business_ID_mds2']) & merge2['MDS ID'].isna()), merge2['Business_ID_mds2'], merge2['MDS ID'])
# new_df['MDS ID'] = np.where((pd.notnull(new_df['Business_ID']) & new_df['MDS ID'].isna()), new_df['Business_ID'], new_df['MDS ID'])
