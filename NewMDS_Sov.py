
import pandas as pd
import numpy as np
from difflib import get_close_matches



mds = pd.read_excel("../Origami_locations/Origami MDS 07-17-20.xlsx", sheet_name='BX MDM w MDS')
sov = pd.read_excel("07_16_Overlay.xlsx")


"""SOV Overaly---------------------------------------------------------------------------------------------------------"""
print(len(mds))
mds = mds[~mds.duplicated(subset=['BX Property ID'], keep=False)]
print(len(mds))
new_df = pd.merge( sov,  mds[['BX Property ID', 'MDS Property ID']],  how='left', left_on='Business_ID', right_on='BX Property ID', validate = 'm:1', suffixes=('', '_mds'))
new_df['Business_ID'] = np.where(pd.notnull(new_df['MDS Property ID']), new_df['MDS Property ID'], new_df['Business_ID'])
# df = new_df[[columns]]
new = new_df[~new_df['BX Property ID'].isna()]
print(len(new_df[~new_df['BX Property ID'].isna()]))
print(len(new_df[~new_df['MDS Property ID'].isna()]))



print(len(new_df))
#new_df.to_excel('Overlay_07_18.xlsx')

"""SOV MDS ID---------------------------------------------------------------------------------------------------------"""
st = ['Alberta', 'British Columbia', 'Manitoba', 'Ontario', 'Quebec', 'Saskatchewan' ]

def getMatches(text, Addresses, count, precision):
    text = str(text)
    if get_close_matches(text, Addresses):
        return '// '.join(get_close_matches(text, Addresses, count, precision))
    else:
        return ''


sov_ca = sov[sov['State'].isin(st)]
mds2 = pd.read_excel("../Origami_locations/Origami MDS 07-17-20.xlsx", sheet_name='MDS')

# properties = mds2['Name'].to_list()
# sov_ca['possibleMatches'] = sov_ca['Property'].apply(lambda text: getMatches(text, properties, 2, 0.6))


# merge_df = pd.merge( sov,  mds2[['Name', 'MDS ID']],  how='left', left_on='Property', right_on='Name', validate = 'm:1', suffixes=('', '_mds'))
# print(len(sov_ca[~sov_ca['possibleMatches'].isna()]))
#sov_ca.to_excel('possibleMatchesMDS.xlsx')

possible = pd.read_excel('possibleMatchesMDS.xlsx')
print(len(possible[possible.duplicated(subset=['Address','Property'])]))
possible = possible.drop_duplicates(subset=['Address','Property'])
matches = pd.merge( possible,  mds2[['Name', 'MDS ID']],  how='left', left_on='possibleMatches', right_on='Name', validate = 'm:1', suffixes=('', '_mds'))
matches_df = pd.merge( sov,  matches[['Address','Property', 'MDS ID']],  how='left', left_on=['Address', 'Property'], right_on=['Address','Property'], validate = 'm:1', suffixes=('', '_mds'))
matches_df['Business_ID'] = np.where((pd.notnull(matches_df['MDS ID']) & matches_df['State'].isin(st)), matches_df['MDS ID'], new_df['Business_ID'])
matches_df.to_excel("07_18_Overlay.xlsx", index=False)
check = matches_df[matches_df['State'].isin(st)][['Business_ID', 'MDS ID']]









