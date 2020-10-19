# -*- coding: utf-8 -*-
"""
Created on Tue Dec  3 15:29:36 2019
This code performs the overlay from the additional ambigeous matches made by the insurance team.
@author: Cheikh Sidi El Moctar
"""
#importing dependencies
import pandas as pd
import table_mapping


def match(df):
    '''This function takes a dataframe (final SOV file) as argument where its going to pull the additional matches into the final file, and the new assigned 
    Business IDs for missing addresses from MDS and Q3.'''
    
    # setting up the paths to un-matched spreadsheet and MDS and Q3
    SOV_unmatched = r"S:\Revantage\IT\FRP\Cleaned_Data\Un-mapped_addresses.xlsx"
    MDS_REV = r"S:\Revantage\IT\FRP\Cleaned_Data\Final\Revantage\Clean_MDS_v3.xlsx"
    Q3_REV = r"S:\Revantage\IT\FRP\Cleaned_Data\Final\Revantage\Q3'19 US Master Property List_harcoded_to Rvtg.xlsx"
    
    
    df_sov = pd.read_excel(SOV_unmatched)
    df_mds = pd.read_excel(MDS_REV)
    df_Q3 = pd.read_excel(Q3_REV)
    
    k1_sov = df_sov.columns[0]
    k2_sov = df_sov.columns[2]
    k1_add = df_sov.columns[1]
    
    k2_mds = df_mds.columns[0]
    
    k2_Q3 = df_Q3.columns[5]
    
    k_sov = df.columns[4]
    
    print("Starting the un-matched overlay.....................")
    Q3_id = table_mapping.Q3_id
    print(Q3_id)
    n = 0
    a = 0
    m = 700000
    k = 0
    for i, d in df[k_sov].items():    
        if d.strip() in df_sov[k1_add].to_list(): 
            k +=1
            bs = df_sov.loc[df_sov[k1_add] == d, k1_sov].iloc[0]
            Q3_r = df_sov.loc[df_sov[k1_add] == d, k2_sov].iloc[0]
            bs_ids = table_mapping.bs_ids
            if bs in df_mds[k2_mds].to_list() and (pd.isna(df.loc[i, 'Business ID'])):
                B1 = df_mds.loc[df_mds['Business ID']== bs, "Business ID"].iloc[0]
                
                if B1 in bs_ids:
                    s = bs_ids.count(B1)
                    df.loc[i,"Business ID"] = f'CA-{B1}_{s}'
                    bs_ids.append(B1)
                else:
                    df.loc[i,"Business ID"] = f'CA-{B1}'
                    bs_ids.append(B1)
                df.at[i,"Property ID"] = df_mds.loc[df_mds['Business ID']== bs, "PropertyID"].iloc[0]
                df.at[i,"Location"] = df_mds.loc[df_mds['Business ID']== bs, "Location"].iloc[0]
                df.at[i,"Property"] = df_mds.loc[df_mds['Business ID']== bs, "Property"].iloc[0]
                df.at[i,"Fund"] = df_mds.loc[df_mds['Business ID']== bs, "Fund"].iloc[0]
                df.at[i,"Investment"] = df_mds.loc[df_mds['Business ID']== bs, "Investment"].iloc[0]
                df.at[i,"Legal Entity"] = df_mds.loc[df_mds['Business ID']== bs, "LegalEntity"].iloc[0]
                df.at[i,"Portfolio Company ID"] = df_mds.loc[df_mds['Business ID']== bs, "PortfolioCompanyID"].iloc[0]
                df.at[i,"Portfolio Company"] = df_mds.loc[df_mds['Business ID']== bs, "PortfolioCompany"].iloc[0]
                df.at[i,"Occupancy/Asset Class"] = df_mds.loc[df_mds['Business ID']== bs, "AssetClass"].iloc[0]
                df.loc[i,"Revantage Manage"] = df_mds.loc[df_mds['Business ID']== bs, "Revantage Manage"].iloc[0]
                df.loc[i, 'mds/Q3'] = 'MDS_r'
                n +=1
    
            elif Q3_r in df_Q3[k2_Q3].to_list() and (pd.isna(df.loc[i, 'Business ID'])):
                    
                #df.loc[i,"Business ID"] = df_Q3.loc[df_Q3[k2_Q3]== Q3_r, "Business_ID"].iloc[0]
                
                df.loc[i,"Business ID"] = f'CA-{Q3_id}'
                df.at[i,"Fund"] = df_Q3.loc[df_Q3[k2_Q3]== Q3_r, "Fund"].iloc[0]
                df.at[i, "Investment"] = df_Q3.loc[df_Q3[k2_Q3]== Q3_r, "Investment Deal"].iloc[0]
                df.at[i, "Occupancy/Asset Class"] = df_Q3.loc[df_Q3[k2_Q3]== Q3_r, "Sector"].iloc[0]
                df.at[i, "Property"] = df_Q3.loc[df_Q3[k2_Q3]== Q3_r, "Property Name"].iloc[0]
                df.loc[i, "Revantage Manage"] = df_Q3.loc[df_Q3[k2_Q3]== Q3_r, "Revantage Manage"].iloc[0]
                Q3_id += 1
                df.loc[i, 'mds/Q3'] = 'Q3_r'
                a +=1
            elif (pd.isna(df.loc[i, 'Business ID'])):
                df.loc[i,"Business ID"] = f'CA-{m}'
                m +=1
                df.loc[i, 'mds/Q3'] = 'Non_Match'
        elif (pd.isna(df.loc[i, 'Business ID'])):
          df.loc[i,"Business ID"] = f'CA-{m}'
          m +=1
          df.loc[i, 'mds/Q3'] = 'Non Matches'
        if (pd.isna(df.loc[i, 'Location'])):
          df.loc[i,"Location"] = f'No_Information'
    #Non_match = [c for c in df_sov[k1_add].to_list() if not c in df_mds[k2_mds].to_list()]
    #un_match_df = df_sov[~df_sov[k1_add].isin(Non_match)]
    # = df_sov[df_sov[k1_add].isin(Non_match)]
    #un_match_df.to_excel('un-match_test.xlsx')            
    print(f"MDS : {n}, Q3 : {a}, Total : {k}")
    print()        
    return df
        

    
    
    
    
  