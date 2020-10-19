import pandas as pd
import numpy as np
import re
from difflib import get_close_matches
from xlwt import Workbook
import xlwt
# (.*?) ([^ ]+?) ?((?<= )APT)? ?((?<= )\d*)?$
def parser(st):
    result = re.search(r"[a-zA-z]+", st)
    print(result.group(1))
    


ESIS = pd.read_excel("../../RevantageBlackstone/PossibleMatches_20201014/ESIS Locations.xlsx")
SOV = pd.read_excel("../../../Downloads/Hotels SOV Overlay.xlsx")


# Us_States = utils_data.us_state_abbrev
# st_dict = utils_data.st_dict
# dir_dict = utils_data.dir_dict
# # Data Cleaning and Address normalization functions
# def replace_last(source_string, replace_what, replace_with):
#         head, _sep, tail = source_string.rpartition(replace_what)
#         return head + replace_with + tail
    
# def replace_second(source_string, dict_):
#     s = source_string.split(' ')
#     if len(s) >1 :
#         if s[1].replace('.', '').lower() in dict_.keys():
#             s[1] = dict_[s[1].replace('.', '')]
#             return ' '.join(s)
#         else:    
#             return source_string
#     else:
#         return source_string


# def lower_No_spaces(df, col):
#       df[col] = [(' '.join(str(d).split())).lower().strip() for d in df[col]]
#       df[col] = [d.replace(' ', '') for d in df[col]]
#       return df[col].astype(str)

# def Address_Normal(df):
#     df['Address'] = [(' '.join(str(d).split())).lower().strip() for d in df['Address']]
#     df['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df['Address']]
#     df['st_add'] = [replace_second(d,dir_dict) for d in df['st_add']]
#     df['Add'] = [d.replace(' ', '') for d in df['st_add']]
#     return df

def getMatches(text, Addresses, count, precision):
    text = str(text)
    if get_close_matches(text, Addresses):
        return '// '.join(get_close_matches(text, Addresses, count, precision))
    else:
        return '// '.join(get_close_matches(text, Addresses, count, 0.3))

   
def possibleMatches(count, precision, df, df2):
    # df['address'] = df['Street Address'] +' '+ df['City']
    df2.index +=2
    df2.reset_index(inplace=True)
    df2['address'] =  df2['index'].astype(str) + '  '+ df2['Location Name']
    Addresses = df2['address'].to_list()
    Addresses = [str(d) for d in Addresses]
    df['possibleMatches'] = df['LocationName'].apply(lambda text: getMatches(text, Addresses, count, precision))
    df.to_excel('test.xlsx', index=False, engine='xlsxwriter')
    return df

possibleMatches(4, 0.7, ESIS, SOV)



