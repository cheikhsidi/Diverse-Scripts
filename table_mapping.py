# Importing dependencies
import pandas as pd
import numpy as np
from additiona_match import match
#from difflib import SequenceMatcher
import re


# Reading the excel files
SOV_AIR = r"S:\Revantage\IT\FRP\Cleaned_Data\AIR\FRPDQ_TotalSOV_ForAIR_GeolocatingSummary.xlsx"
#SOV_REV = r"S:\Revantage\IT\FRP\Cleaned_Data\Final\New_Final_SOV.xlsx"

MDS_AIR = r"S:\Revantage\IT\FRP\Cleaned_Data\AIR\FRPDQ_MDS_ForAIR_GeolocatingSummary.xlsx"
MDS_REV = r"S:\Revantage\IT\FRP\Cleaned_Data\Final\Revantage\Clean_MDS_v3.xlsx"

Q3_AIR = r"S:\Revantage\IT\FRP\Cleaned_Data\AIR\FRPDQ_Q3_ForAIR_GeolocatingSummary.xlsx"
Q3_REV = r"S:\Revantage\IT\FRP\Cleaned_Data\Final\Revantage\Q3'19 US Master Property List_harcoded_to Rvtg.xlsx"


Q3_id = 500000
bs_ids = []

def overlay(df1):
    
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
    df2 = pd.read_excel(SOV_AIR)
    k1 = df1.columns[4]
    #print(k1)
    k2 = df2.columns[1]
    st_dict = {'st':'street', 'pkwy':'parkway', 'ave':'avenue', 'dr': 'drive', 'rd':'road', 'blvd':'boulevard', 'ln':'lane', 'hwy':'highway'}
    dir_dict = {'e':'east', 'w':'west', 's':'south', 'n': 'north', 'ne':'north east', 'nw':'north west', 'se':'south east', 'sw':'south west'}
    
    
    
    df1[k1] = [(' '.join(str(d).split())).lower().strip() for d in df1[k1]]
    df2[k2] = [(' '.join(str(d).split())).lower().strip() for d in df2[k2]]
    df1['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df1[k1]]
    df1['st_add'] = [replace_second(d,dir_dict) for d in df1['st_add']]
    df2['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df2[k2]]
    df2['st_add'] = [replace_second(d,dir_dict) for d in df2['st_add']]
    df1['add'] = [d.replace(' ', '') for d in df1['st_add']]
    df2['add'] = [d.replace(' ', '') for d in df2['st_add']]
    
    #c1 = df1[k1].nunique()
    #c2 = df2[k2].nunique()
    df2 = df2.drop_duplicates(subset='add') 
    df_left_sov = pd.merge(df1, df2, on ='add', how = 'left', suffixes =('', '_x'), validate = 'm:1')
     
    
    """ Adding geolocation into the MDS
    ====================================================================================================================="""
    
    df1_mds = pd.read_excel(MDS_REV)
    df2_mds = pd.read_excel(MDS_AIR)
    
    k1_mds = df1_mds.columns[6]
    #print(k1_mds)
    k2_mds = df2_mds.columns[4]
    df1_mds[k1_mds] = [re.sub('\s+',' ',str(d).lower().strip()) for d in df1_mds[k1_mds]]
    df2_mds[k2_mds] = [re.sub('\s+',' ',str(d).lower().strip()) for d in df2_mds[k2_mds]]
    
   
    df2_mds = df2_mds.drop_duplicates(subset=k2_mds)
    df_left_mds = pd.merge(df1_mds, df2_mds, left_on=k1_mds, right_on=k2_mds, how = 'left', validate = 'm:1')
    #count_mds = df_left_mds[k2_mds].isna().sum()
    
    
    """ Adding geolocation into the Q3
    ====================================================================================================================="""
    
    df1_Q3 = pd.read_excel(Q3_REV)
    df2_Q3 = pd.read_excel(Q3_AIR)
    
    k1_Q3 = df1_Q3.columns[11]
    #print(k1_Q3)
    k2_Q3 = df2_Q3.columns[4]
    df1_Q3[k1_Q3] = [re.sub('\s+',' ',str(d).lower().strip()) for d in df1_Q3[k1_Q3]]
    df2_Q3[k2_Q3] = [re.sub('\s+',' ',str(d).lower().strip()) for d in df2_Q3[k2_Q3]]
    
    
    df2_Q3 = df2_Q3.drop_duplicates(subset=k2_Q3) 
    df_left_Q3 = pd.merge(df1_Q3, df2_Q3, left_on=k1_Q3, right_on=k2_Q3, how = 'left', validate = 'm:1')
    #count_Q3 = df_left_Q3[k2_Q3].isna().sum()
        
    
    """ Joining the SOV, and MDS
    ====================================================================================================================="""
    
    # Creating stardard Address in SOV
    df_left_sov['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df_left_sov[k1]]
    df_left_sov['st_add'] = [replace_second(d,dir_dict) for d in df_left_sov['st_add']]
    df_left_sov["add"] = [str(d).replace('-', '').replace('.', '').replace(' ', '').strip().replace('"','').replace("'", "").lower() for d in df_left_sov['st_add']]
    # Creating Standard Address MDS
    df_left_mds['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1].replace('.','') in st_dict.keys() else d for d in df_left_mds[k1_mds]]
    df_left_mds['st_add'] = [replace_second(d,dir_dict) for d in df_left_mds['st_add']]
    df_left_mds["add"] = [str(d).replace('-', '').replace('.', '').replace(' ', '').strip().replace('"','').replace("'", "").lower() for d in df_left_mds['st_add']]
    #df_left_mds = df_left_mds.drop_duplicates(subset='add')
    #df_left_mds = df_left_mds.drop_duplicates(subset=['Latitude', 'Longitude'])
    #Creating Standard Address Q3
    df_left_Q3['st_add'] = [replace_last(d, d.split(' ')[-1], st_dict[d.split(' ')[-1].replace('.','')]) if d.split(' ')[-1] in st_dict.keys() else d for d in df_left_Q3[k1_Q3]]
    df_left_Q3['st_add'] = [replace_second(d,dir_dict) for d in df_left_Q3['st_add']]
    df_left_Q3["add"] = [str(d).replace('-', '').replace('.', '').replace(' ', '').strip().replace('"','').replace("'", "").lower() for d in df_left_Q3['st_add']]
    
    #df_left_Q3 = df_left_Q3.drop_duplicates(subset='add')
    
    df_left_sov['mds/Q3'] = np.nan
    df_left_sov['Revantage Manage'] = np.nan
    
    
    print ("Starting the Overlay----------------------------------------------------------")

    #print("Overlaying the SOV and MDS-------------------sov -  mds ------------------------------")
    df_left_sov['Lat'] = [cor for cor in df_left_sov['Latitude']]
    df_left_sov['Long'] = [cor for cor in df_left_sov['Longitude']]
    
    df_left_mds['Lat'] = [cor for cor in df_left_mds['Latitude']]
    df_left_mds['Long'] = [cor for cor in df_left_mds['Longitude']]
    
    df_left_Q3['Lat'] = [cor for cor in df_left_Q3['Latitude']]
    df_left_Q3['Long'] = [cor for cor in df_left_Q3['Longitude']]
         
         
    df_left_sov['lat-long'] = list(zip(df_left_sov.Lat, df_left_sov.Long))
    df_left_mds['lat-long'] = list(zip(df_left_mds.Lat, df_left_mds.Long))
    df_left_Q3['lat-long'] = list(zip(df_left_Q3.Lat, df_left_Q3.Long))
    
    n = 0
    m = 0
    b = 0
    a = 0
    
    global bs_ids
    for i, d in df_left_sov['add'].items():
       
        if (df_left_sov.loc[i, 'add'] in df_left_mds['add'].to_list()) and (pd.isna(df_left_sov.loc[i, 'Business ID'])) and (df_left_sov.loc[i, 'File_Source'] != 'Great Wolf Resorts - SOV' or df_left_sov.loc[i, 'File_Source'] != 'Colony Industrial - SOV' ):
            Business_id = df_left_mds.loc[df_left_mds['add']== str(d), "Business ID"].iloc[0]
            
            if Business_id in bs_ids:   
                s = bs_ids.count(Business_id)
                bs_ids.append(Business_id)
                
                df_left_sov.loc[i,"Business ID"] = f'CA-{Business_id}_{s}'    
            else:
               df_left_sov.loc[i,"Business ID"] = f'CA-{Business_id}'
               bs_ids.append(Business_id)
            df_left_sov.at[i,"Property ID"] = df_left_mds.loc[df_left_mds['add']== str(d), "PropertyID"].iloc[0]
            df_left_sov.at[i,"Location"] = df_left_mds.loc[df_left_mds['add']== str(d), "Location"].iloc[0]
            df_left_sov.at[i,"Property"] = df_left_mds.loc[df_left_mds['add']== str(d), "Property"].iloc[0]
            df_left_sov.at[i,"Fund"] = df_left_mds.loc[df_left_mds['add']== str(d), "Fund"].iloc[0]
            df_left_sov.at[i,"Investment"] = df_left_mds.loc[df_left_mds['add']== str(d), "Investment"].iloc[0]
            df_left_sov.at[i,"Legal Entity"] = df_left_mds.loc[df_left_mds['add']== str(d), "LegalEntity"].iloc[0]
            df_left_sov.at[i,"Portfolio Company ID"] = df_left_mds.loc[df_left_mds['add']== str(d), "PortfolioCompanyID"].iloc[0]
            df_left_sov.at[i,"Portfolio Company"] = df_left_mds.loc[df_left_mds['add']== str(d), "PortfolioCompany"].iloc[0]
            df_left_sov.at[i,"Occupancy/Asset Class"] = df_left_mds.loc[df_left_mds['add']== str(d), "AssetClass"].iloc[0]
            df_left_sov.loc[i,"Revantage Manage"] = df_left_mds.loc[df_left_mds['add']== str(d), "Revantage Manage"].iloc[0]
            df_left_sov.loc[i, 'mds/Q3'] = 'MDS'
            n +=1
        elif  (df_left_sov.loc[i, 'lat-long'] in df_left_mds['lat-long'].to_list()) and (pd.isna(df_left_sov.loc[i, 'Business ID']) and (df_left_sov.loc[i, 'lat-long'] !=(0,0))) and (df_left_sov.loc[i, 'File_Source'] != 'Great Wolf Resorts - SOV' or df_left_sov.loc[i, 'File_Source'] != 'Colony Industrial - SOV' ): 
            B2 = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "Business ID"].iloc[0]
            
            if B2 in bs_ids:
                s = bs_ids.count(B2)
                bs_ids.append(B2)
                df_left_sov.loc[i,"Business ID"] = f'CA-{B2}_{s}'
                
            else:
                df_left_sov.loc[i,"Business ID"] = f'CA-{B2}'
                bs_ids.append(B2)
            df_left_sov.at[i,"Property ID"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "PropertyID"].iloc[0]
            df_left_sov.at[i,"Location"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "Location"].iloc[0]
            df_left_sov.at[i,"Property"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "Property"].iloc[0]
            df_left_sov.at[i,"Legal Entity"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "LegalEntity"].iloc[0]
            df_left_sov.at[i,"Portfolio Company ID"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "PortfolioCompanyID"].iloc[0]
            df_left_sov.at[i,"Portfolio Company"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "PortfolioCompany"].iloc[0]
            df_left_sov.at[i,"Fund"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "Fund"].iloc[0]
            df_left_sov.at[i,"Investment"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "Investment"].iloc[0]
            df_left_sov.at[i,"Occupancy/Asset Class"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "AssetClass"].iloc[0]
            df_left_sov.loc[i,"Revantage Manage"] = df_left_mds.loc[df_left_mds['lat-long']== df_left_sov.loc[i, 'lat-long'], "Revantage Manage"].iloc[0]
            df_left_sov.loc[i, 'mds/Q3'] = 'MDS'
            m +=1
    
    
   
    #for i, d in df_left_sov['add'].items():
        
        if (df_left_sov.loc[i, 'add'] in df_left_Q3['add'].to_list()) and (pd.isna(df_left_sov.loc[i, 'Business ID'])):
            global Q3_id
            df_left_sov.loc[i,"Business ID"] = f'CA-{Q3_id}'
            #df_left_sov.loc[i,"Business ID"] = df_left_Q3.loc[df_left_Q3['add']== str(d), "Business_ID"].iloc[0]
            df_left_sov.at[i,"Fund"] = df_left_Q3.loc[df_left_Q3['add']== str(d), "Fund"].iloc[0]
            df_left_sov.at[i, "Investment"] = df_left_Q3.loc[df_left_Q3['add']== str(d), "Investment Deal"].iloc[0]
            df_left_sov.at[i, "Occupancy/Asset Class"] = df_left_Q3.loc[df_left_Q3['add']== str(d), "Sector"].iloc[0]
            df_left_sov.at[i, "Property"] = df_left_Q3.loc[df_left_Q3['add']== str(d), "Property Name"].iloc[0]
            df_left_sov.loc[i, "Revantage Manage"] = df_left_Q3.loc[df_left_Q3['add']== str(d), "Revantage Manage"].iloc[0]
            Q3_id += 1
            df_left_sov.loc[i, 'mds/Q3'] = 'Q3'
            a +=1
        elif (df_left_sov.loc[i, 'lat-long'] in df_left_Q3['lat-long'].to_list()) and (pd.isna(df_left_sov.loc[i, 'Business ID']) and (df_left_sov.loc[i, 'lat-long'] !=(0,0))):
            
            
            df_left_sov.at[i,"Fund"] = df_left_Q3.loc[df_left_Q3['lat-long']== df_left_sov.loc[i, 'lat-long'], "Fund"].iloc[0]
            df_left_sov.at[i, "Investment"] = df_left_Q3.loc[df_left_Q3['lat-long']== df_left_sov.loc[i, 'lat-long'], "Investment Deal"].iloc[0]
            #df_left_sov.loc[i,"Business ID"] = df_left_Q3.loc[df_left_Q3['lat-long']== df_left_sov.loc[i, 'lat-long'], "Business_ID"].iloc[0]
            df_left_sov.loc[i,"Business ID"] = f'CA-{Q3_id}'
            
            df_left_sov.at[i, "Occupancy/Asset Class"] = df_left_Q3.loc[df_left_Q3['lat-long']== df_left_sov.loc[i, 'lat-long'], "Sector"].iloc[0]
            df_left_sov.at[i, "Property"] = df_left_Q3.loc[df_left_Q3['lat-long']== df_left_sov.loc[i, 'lat-long'], "Property Name"].iloc[0]
            df_left_sov.loc[i, "Revantage Manage"] = df_left_Q3.loc[df_left_Q3['lat-long']== df_left_sov.loc[i, 'lat-long'], "Revantage Manage"].iloc[0]
            df_left_sov.loc[i, 'mds/Q3'] = 'Q3'
            Q3_id += 1
            b +=1
    

    print(f"{n} overlayed from MDS based on the address")
    print(f"{m} overlayed from MDS based on the lat - long")


    
    #df_sts['overlayed_Q3'] = df_left_sov.groupby('File_Source').count().reset_index()['mds/Q3'] -  df_sts['overlayed_mds']
    print(f"{a} overlayed from Q3 based on the address")
    print(f"{b} overlayed from Q3 based on the lat-long")
        
    #finding un-matched from mds
    #missing_mds = df_left_mds[~df_left_mds['add'].isin(df_left_sov['add'].to_list()) | ~df_left_mds['add'].isin(df_left_sov['add'].to_list())]
    #missing_mds.to_excel('S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\Data_Stats\missing_mds.xlsx')
    
    #finding un-matched from Q3
    #missing_Q3 = df_left_Q3[~df_left_Q3['add'].isin(df_left_sov['add'].to_list()) | ~df_left_Q3['add'].isin(df_left_sov['add'].to_list())]
   #missing_Q3.to_excel('S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\Data_Stats\missing_Q3.xlsx')
    
    """ selecting final columns
    ====================================================================================================================="""
   
    columns =["Business ID", "Location", "Property ID", "Property","Address Line 1", "Address Line 2",
    "City", "State", "Postal Code", "Portfolio Company ID", "Portfolio Company", "Fund",
    "Investment", "Occupancy/Asset Class", "Legal Entity", "Broker Firm", "Broker Contact",
    "Construction Type", "Year Built", "Year Upgraded", "# of Stories", "Sq. Ft.",  "% Sprinklered",
    "Bldg Value", "Contents Value", "BI/EE", "TIV", "Flood Zone", "Shape of Roof",
    "Roof Covering", "Year Roof Replaced", "Basement",  "Frame-Foundation Connection",
    "Roof Strapped", "Additional Structures", "Resistance - Windows", "EM Year Upgrade",
    "Soft Story", "Cladding","Cripple Walls", "Frame Bolted", "EM Purlin Anchoring",
    "URM Retrofit", "Mechanical & Electrical Equipment Bracing","Pounding", "Contingent Coverage",
    "Main/Insurance Portfolio", "Invitation Homes - Region", "Normalized Occupancy",
    "Normalized Construction Type","Latitude", "Longitude",'mds/Q3', 'File_Source', 'Revantage Manage']
     #'File_Source' 
    
   
    
    
    df_mds_dups = pd.read_excel(r'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\MDS_dup (to be added to the final SOV).xlsx')
    df_left_sov = df_left_sov[columns]
    df_left_sov = match(df_left_sov)
    df_mds_dups['Main/Insurance Portfolio'] = f'Information Only'
    
    #Formatting the Buisness ID for Additional duplicates for Information Only
    for i, d in df_mds_dups['Business ID'].items():
        
        if d in bs_ids:
            s = bs_ids.count(d)
            df_mds_dups.loc[i,"Business ID"] = f'CA-{d}_{s}'
            bs_ids.append(d)
        else:
            df_mds_dups.loc[i,"Business ID"] = f'CA-{d}'
            bs_ids.append(d)
    df_left_sov =  df_left_sov.append(df_mds_dups, sort=False)
    print ("Done with the Overlay......................")
    print()
    
  
    return df_left_sov
      
#x = overlay()
#x.to_excel('overlay_v3.xlsx')

