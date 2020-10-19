
import pandas as pd
import numpy as np



df1 = pd.read_excel("../Q_Data_Revanatge/Q1'20 US Master Property List_vF_Hardcoded.xlsx", skiprows=9)
df2 = pd.read_excel("final_Overlay_07-13.xlsx")

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
# df2_500 = df2[(df2['Business_ID'].str[3] == '5') & (df2['Business_ID'].str[4] == '0')]
# df2 = df2[( df2['Business_ID'].str.contains("_") == False) & (df2['Business_ID'].str[3] != '8') & (df2['Business_ID'].str[3] != '7' )]
# df2 = df2[df2['Business_ID'].str[3] != 8 ]

df1 = df1.drop_duplicates(subset=('Property Name')) 
# df_left_Q3 = pd.merge(df1_Q3, df2_Q3, left_on=k1_Q3, right_on=k2_Q3, how = 'left', validate = 'm:1')
new_df = pd.merge( df2,  df1,  how='left', left_on='Property', right_on='Property Name', validate = 'm:1', suffixes=('', '_Q1'))
new_df['Business_ID'] = np.where((pd.notnull(new_df['Property MDM ID']) & (new_df['Business_ID'].str[3] == '5') & (new_df['Business_ID'].str[4] == '0')), new_df['Property MDM ID'], new_df['Business_ID'])
# df = new_df[[columns]]


"""SOV Overaly---------------------------------------------------------------------------------------------------------"""
# final = pd.merge( df2,  new_df,  how='left', left_on='Business_ID', right_on='Business_ID', validate = 'm:1')
# final['Portfolio_Company_x'] = np.where(pd.notnull(final['Portfolio Company']), final['Portfolio Company'], final['Portfolio_Company_x'])
print(len(new_df))
print(len(df1), len(df2), len(new_df[~new_df['Portfolio_Company'].isna()]))

print(len(new_df[~new_df['Property MDM ID'].isna()]))
new_df.to_excel('Overlay_07_16.xlsx')

# df['Type'] = np.where(df['Type'] == 4, "partial"+df['Program']+"_"+df['Breadth'],df['Type'])


