import pandas as pd
import numpy as np
import os

file1 = "../Origami_locations/Origami Locations - 10-30-19.xlsx"
file2 = "../Origami_locations/Origami Locations - 7-9-20.xlsx"
column = 'BusinessID'
def df_diff(file1, file2, column, skiprows=0):
    """Find rows which are different between two DataFrames."""
    df1 = pd.read_excel(file1, skiprows=skiprows)
    df2 = pd.read_excel(file2, skiprows=skiprows)
    # df2[column] = df2.column.astype(int)
    df1['Match?'] = np.where(df1[column].isin(df2[column].to_list()), 'True', 'False')
    df2['Match?'] = np.where(df2[column].isin(df1[column].to_list()), 'True', 'False')
    
    print(len(df1[df1['Match?'] == 'False']))
    print(len(df2[df2['Match?'] == 'False']))
    print(df1[column].dtypes)
    print(df2[column].dtypes)
    
        
    Q3 = df1[df1['Match?'] == 'False']
    Q1 = df2[df2['Match?'] == 'False']

    Q1.to_excel("file1_New.xlsx")
    Q3.to_excel("file2_New.xlsx") 
    # return diff_df



df1 = pd.read_excel("../SOVs/Blackstone Values Revised 5.31.20_Willis Tower and Cosmo Removed_Sec Mod Incl.xlsx")
df2 = pd.read_excel("../SOVs/Final SOV Uploaded to Origami 12-10 - Includes MDS.xlsx")


new_df = pd.merge(df1, df2,  how='inner', left_on=['Address_Line_1','Address_Line_2', 'City'], 
                    right_on = ['Address_Line_1','Address_Line_2', 'City'])

# new_df.to_excel('overlay_sov.xlsx')

print(len(df1), len(df2), len(new_df[~new_df['Business_ID'].isna()]))
# df1 = pd.read_excel("../Q_Data_Revanatge/Q3'19 US Master Property List_harcoded_to Rvtg.xlsx", skiprows=9)
# df2 = pd.read_excel("../Q_Data_Revanatge/Q1'20 US Master Property List_vF_Hardcoded.xlsx", skiprows=9)
             
# df1['Match?'] = np.where(df1['iLevel Property Names'].isin(df2['iLevel Property Names'].to_list()), 'True', 'False')
# df2['Match?'] = np.where(df2['iLevel Property Names'].isin(df1['iLevel Property Names'].to_list()), 'True', 'False')           

# # print(len(dataframe_difference(df1, df2)['iLevel Property Names'].to_list()))
# # print(len(Q1), len(Q3))




# print(len(df1[df1['Match?'] == 'False']))

# print(len(df1[df2['Match?'] == 'False']))


# Q3 = df1[df1['Match?'] == 'False']

# Q1 = df2[df2['Match?'] == 'False']

# Q1.to_excel("Q1_New.xlsx")

# Q3.to_excel("Q3_New.xlsx")













# df2['Newly Added?'] = df2['Newly Added?'].astype(str)
# dataframe_difference(df1, df2)

# print(df1.dtypes)
# print(df2.dtypes)

# print (df1.head())
# print (df2.head())