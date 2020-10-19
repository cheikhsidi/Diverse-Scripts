# Importing dependencies
import pandas as pd
import numpy as np


# Loading all files into a dataframe
old_sov = pd.read_excel("../SOVs/Final SOV Uploaded to Origami 12-10 - Includes MDS.xlsx")
sov = pd.read_excel("Overlay_SOVs.xlsx")
Q1 = pd.read_excel("../Q_Data_Revanatge/Q1'20 US Master Property List_vF_Hardcoded.xlsx", skiprows=9)
mds = pd.read_excel("../Origami_locations/Added Column - Origami Locations - 7-9-20.xlsx")
possible = pd.read_excel("Possibel Matches (1).xlsx")

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

retainedCols = ['Address_Line_1','Address_Line_2', 'City', 'possibleMatches SOV 12-10', 'possibleMatches Q1 2020', 'possibleMatches Origami Locations 7-9-20' ]

sov_cols = ['index', 'Business_ID','Location',	'Property_ID',	'Property',	'Portfolio_Company_ID',	'Portfolio_Company',	
            'Fund',	'Investment',	'Occupancy_Asset_Class',	'Legal_Entity',	'Main_Insurance_Portfolio',	
            'Normalized_Construction_Type',	'Normalized_Occupancy',	'Latitude',	'Longitude']

mds_cols = ['index', 'BusinessID',	'Name', 'Legal Entity',	'Investment','Fund', 'Portfolio Company']
Q1_cols = ['index','Property MDM ID',	'Fund',	'Investment Deal', 'Portfolio Company',	'Property Name','Sector']


# drop dups
possible = possible.drop_duplicates(subset=['Address_Line_1','Address_Line_2', 'City'])

# Overlaying sov and manual ross mapping with the possible matches  
merg1 = pd.merge(sov,possible[retainedCols], on=['Address_Line_1','Address_Line_2', 'City'], how='left', validate = 'm:1')
merg1['sov_indexs'] = merg1['possibleMatches SOV 12-10'].apply(lambda x : int(x.split(' ')[0]) if (x != '' and not isinstance(x, float) and not isinstance(x, int)) else x)
merg1['Q1_indexs'] = merg1['possibleMatches Q1 2020'].apply(lambda x : int(x.split(' ')[0]) if (x != '' and not isinstance(x, float) and not isinstance(x, int)) else x)
merg1['mds_indexs'] = merg1['possibleMatches Origami Locations 7-9-20'].apply(lambda x : int(x.split(' ')[0]) if (x != '' and not isinstance(x, float) and not isinstance(x, int)) else x)

# Creating an index to merge the data on
old_sov.index +=2
old_sov.reset_index(inplace=True)
# Merging the entire overlay with the old SOV to pull all manually mapped columns
merge_sov = pd.merge(merg1,old_sov[sov_cols], how='left', left_on='sov_indexs', right_on='index',  validate = 'm:1', suffixes=('','_sov'))


sov_cols.remove('index')
# updating all the columns from the new pulled columns  
for col in sov_cols:
    try:
        merge_sov[f'{col}'] = np.where(pd.notnull(merge_sov[f'{col}_sov']), merge_sov[f'{col}_sov'], merge_sov[f'{col}'])
    except Exception as e: print(e)

# Adding an index to Q1 and performing the merge and the update
Q1.index +=11
Q1.reset_index(inplace=True) 
meregeQ1 = pd.merge(merge_sov,Q1[Q1_cols], how='left', left_on='Q1_indexs', right_on='index',  validate = 'm:1', suffixes=('','_Q1'))
meregeQ1['Business_ID'] = np.where(pd.notnull(meregeQ1['Property MDM ID']), meregeQ1['Property MDM ID'], meregeQ1['Business_ID'])
meregeQ1['Fund'] = np.where(pd.notnull(meregeQ1['Fund_Q1']), meregeQ1['Fund_Q1'], meregeQ1['Fund'])
meregeQ1['Investment'] = np.where(pd.notnull(meregeQ1['Investment Deal']), meregeQ1['Investment Deal'], meregeQ1['Investment'])
meregeQ1['Property'] = np.where(pd.notnull(meregeQ1['Property Name']), meregeQ1['Property Name'], meregeQ1['Property'])
meregeQ1['Occupancy_Asset_Class'] = np.where(pd.notnull(meregeQ1['Sector']), meregeQ1['Sector'], meregeQ1['Occupancy_Asset_Class'])
meregeQ1['Portfolio_Company'] = np.where(pd.notnull(meregeQ1['Portfolio Company']), meregeQ1['Portfolio Company'], meregeQ1['Portfolio_Company'])

# Adding an index to MDS and performing the merge and the update
mds.index +=2
mds.reset_index(inplace=True) 

final = pd.merge(meregeQ1,mds[mds_cols], how='left', left_on='mds_indexs', right_on='index',  validate = 'm:1', suffixes=('', '_mds'))
final['Business_ID'] = np.where(pd.notnull(final['BusinessID']), final['BusinessID'], final['Business_ID'])
final['Location'] = np.where(pd.notnull(final['Name']), final['Name'], final['Location'])
final['Investment'] = np.where(pd.notnull(final['Investment_mds']), final['Investment_mds'], final['Investment'])
final['Legal_Entity'] = np.where(pd.notnull(final['Legal Entity']), final['Legal Entity'], final['Legal_Entity'])
final['Portfolio_Company'] = np.where(pd.notnull(final['Portfolio Company_mds']), final['Portfolio Company_mds'], final['Portfolio_Company'])
final['Fund'] = np.where(pd.notnull(final['Fund_mds']), final['Fund_mds'], final['Fund'])

print(len(final))

# Writing the final results to an excel file
#final[sov.columns.to_list()].to_excel('Test_07-14.xlsx', index=False)






