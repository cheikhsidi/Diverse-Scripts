import pandas as pd

# updated = df1.merge(df2, how='left', on=['Code', 'Name'], suffixes=('', '_new'))
# updated['Value'] = np.where(pd.notnull(updated['Value_new']), updated['Value_new'], updated['Value'])
# updated.drop('Value_new', axis=1, inplace=True)

overlay = pd.read_excel("Overlay_SOVs.xlsx")
df = overlay[overlay['Business_ID'].isna()]
df_sov = pd.read_excel("../SOVs/Final SOV Uploaded to Origami 12-10 - Includes MDS.xlsx")
df_mds = pd.read_excel("../Origami_locations/Origami Locations - 7-9-20.xlsx")
df_Q1 = pd.read_excel("../Q_Data_Revanatge/Q1'20 US Master Property List_vF_Hardcoded.xlsx", skiprows=9)

k1_sov = df_sov.columns[0]
k2_sov = df_sov.columns[2]
k1_add = df_sov.columns[1]

k2_mds = df_mds.columns[0]

k2_Q1 = df_Q1.columns[5]

k_sov = df.columns[4]

for i, d in df[k_sov].items():     
    id_ = d.split(' ')[0]
    df.at[i,"Property ID"] = df_mds.loc[df_mds['index']== id_, "PropertyID"].iloc[0]
    df.at[i,"Location"] = df_mds.loc[df_mds['index']== id_, "Location"].iloc[0]
    df.at[i,"Property"] = df_mds.loc[df_mds['index']== id_, "Property"].iloc[0]
    df.at[i,"Fund"] = df_mds.loc[df_mds['index']== id_, "Fund"].iloc[0]
    df.at[i,"Investment"] = df_mds.loc[df_mds['index']== id_, "Investment"].iloc[0]
    df.at[i,"Legal Entity"] = df_mds.loc[df_mds['index']== id_, "LegalEntity"].iloc[0]
    df.at[i,"Portfolio Company ID"] = df_mds.loc[df_mds['index']== id_, "PortfolioCompanyID"].iloc[0]
    df.at[i,"Portfolio Company"] = df_mds.loc[df_mds['index']== id_, "PortfolioCompany"].iloc[0]
    df.at[i,"Occupancy/Asset Class"] = df_mds.loc[df_mds['index']== id_, "AssetClass"].iloc[0]
    df.loc[i,"Revantage Manage"] = df_mds.loc[df_mds['index']== id_, "Revantage Manage"].iloc[0]
    df.loc[i, 'mds/Q1'] = 'MDS_r'
            
