
import pandas as pd
import numpy as np



df1 = pd.read_excel("../SOVs/Blackstone Values Revised 5.31.20_Willis Tower and Cosmo Removed_Sec Mod Incl.xlsx")
df2 = pd.read_excel("../SOVs/Final SOV Uploaded to Origami 12-10 - Includes MDS.xlsx")

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

""" Adding geolocation into the MDS
====================================================================================================================="""

#df1 = pd.read_excel(SOV_REV)
# df2 = pd.read_excel(SOV_AIR)
k1 = df1.columns[0]
print(k1)
k2 = df2.columns[5]
print(k2)
st_dict = {'st':'street', 'pkwy':'parkway', 'ave':'avenue', 'dr': 'drive', 'rd':'road', 'blvd':'boulevard', 'ln':'lane', 'hwy':'highway'}
dir_dict = {'e':'east', 'w':'west', 's':'south', 'n': 'north', 'ne':'north east', 'nw':'north west', 'se':'south east', 'sw':'south west'}

df1.dropna(subset=[k1], inplace =True)
df2.dropna(subset=[k2], inplace =True) 

df1 = df1.query(f'{k1}!="nan"')  
df2 = df2.query(f'{k2}!="nan"')  


df1[k1] = [(' '.join(str(d).split())).lower().strip() for d in df1[k1]]
df2[k2] = [(' '.join(str(d).split())).lower().strip() for d in df2[k2]]

# df1[k1] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df1[k1]]
# df1['st_add'] = [replace_second(d,dir_dict) for d in df1['st_add']]
# df2['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df2[k2]]
# df2['st_add'] = [replace_second(d,dir_dict) for d in df2['st_add']]
# df1['add'] = [d.replace(' ', '') for d in df1['st_add']]
# df2['add'] = [d.replace(' ', '') for d in df2['st_add']]

# print(df1.columns)
# print(df2.columns)
columns = list(df1.columns).extend(['Blackstone',
'Business_ID',
'Portfolio_Company_ID',
'Portfolio_Company',
'Fund',
'Investment',
'Occupancy_Asset_Class',
'Legal_Entity',
'Main_Insurance_Portfolio_y',
'Normalized_Construction_Type_y',
'Normalized_Occupancy_y',
'Latitude_y',
'Longitude_y']
)
# filtring out 7000 and 8000, and Information Only
df2 = df2[df2['Main_Insurance_Portfolio'] != 'Information Only']
df2 = df2[( df2['Business_ID'].str.contains("_") == False) & (df2['Business_ID'].str[3] != '8') & (df2['Business_ID'].str[3] != '7' )]
# df2 = df2[df2['Business_ID'].str[3] != 8 ]

# df2 = df2.drop_duplicates(subset=['Address_Line_1','Address_Line_2', 'City']) 
# df_left_Q3 = pd.merge(df1_Q3, df2_Q3, left_on=k1_Q3, right_on=k2_Q3, how = 'left', validate = 'm:1')
new_df = pd.merge(df1, df2,  how='left', left_on=['Address_Line_1','Address_Line_2', 'City'], right_on=['Address_Line_1','Address_Line_2', 'City'], validate = 'm:1')
# df = new_df[[columns]]
print(len(new_df))
print(len(df1), len(df2), len(new_df[new_df['Business_ID'].isna()]))
new_df.to_excel('New_Overlay_SOVs.xlsx')

