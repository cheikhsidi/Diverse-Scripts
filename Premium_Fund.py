# Importing dependencies
import pandas as pd
import numpy as np

# BRE = pd.read_excel("../../../Downloads/BRE Possible Matches.xlsx")

BRE = pd.read_excel("../../../Downloads/BRE Hotels Resort Property SOV -10-1-20.xlsx")

Origami = pd.read_excel("../../../Downloads/Origami Locations 10-5.xlsx", skiprows=2)
ps = pd.read_excel("../../../Downloads/BRE Possible Matches.xlsx")

# sheets = ['GLP', 'Office', 'Link', 'Space']

# drop dups
# possible = possible.drop_duplicates(subset=['Address_Line_1','Address_Line_2', 'City'])

# Creating an index to merge the data on
Origami.index +=4
Origami.reset_index(inplace=True)


    
# ps['ps_index'] = ps['possibleMatches'].apply(lambda x : int(x.split(' ')[-1]) if (x != '' and not isinstance(x, float) and not isinstance(x, int)) else x)
merg1 = pd.merge(ps,Origami, left_on='Matches_ID', right_on='index', how='left', validate = 'm:1', suffixes=('x', 'y'))
# merg1['Location Number'] = np.where(pd.notnull(merg1['Location Number']), merg1['Location Number'], np.nan)
# merg1['Portfolio Company'] = np.where(pd.notnull(merg1['Portfolio Company']), merg1['Portfolio Company'], np.nan)
# merg1['Fund'] = np.where(pd.notnull(merg1['Fund']), merg1['Fund'], np.nan)
# merg1['Name'] = np.where(pd.notnull(merg1['Name']), merg1['Name'], np.nan)
# merg1['Investment'] = np.where(pd.notnull(merg1['Investment']), merg1['Investment'], np.nan)
# overlays[sh] = (merg1)

print(len(merg1))
print(len(ps))
merg1.to_excel('BRE_OVerlay_20201009.xlsx')






