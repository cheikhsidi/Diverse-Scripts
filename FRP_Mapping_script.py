# -*- coding: utf-8 -*-

"""
Created on Thu Oct 31 19:20:08 2019
Project : FRP
@author: Cheikh Sidi El Moctar
Contact : (862)-571-1002
Description :
    This Script is designed to map broker files into one Origami Template
"""

#Importing Dependencies
import pandas as pd
import os
import collections
import sys
import table_mapping
from analytics import stats
import warnings

warnings.simplefilter(action='ignore', category=Warning)

def Mapping(file,  *args):
    
    if file.upper() == 'SOV':
        #SOv paths
        working_dir = r"S:/Revantage/IT/FRP/EndToEnd/SOV/Source"
        template = pd.read_excel(r'S:\Revantage\IT\FRP\Cleaned_Data\Template.xlsx')
        map_sheet = pd.read_excel(r'S:\Revantage\IT\FRP\Cleaned_Data\SOV_MapSheet.xlsx', index_col=0)
        output_path = r'S:/Revantage/IT/FRP/EndToEnd/SOV/ThereCanBeOnlyOne/Final_SOV.xlsx'
        #output_path = r'S:\Revantage\IT\FRP\Cleaned_Data/Final_SOV_test.xlsx'
        log_path = r'S:/Revantage/IT/FRP/EndToEnd/SOV/ThereCanBeOnlyOne.xlsx'
        
    elif file.upper() == 'LOSS':
        
        # Loss paths
        working_dir = r"S:/Revantage/IT/FRP/EndToEnd/LOSS/Source"
        template = pd.read_excel(r'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\Template_loss.xlsx')
        map_sheet = pd.read_excel(r'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\LOSS_MapSheet.xlsx', index_col=0)
        output_path = r'S:/Revantage/IT/FRP/EndToEnd/LOSS/ThereCanBeOnlyOne/Final_LOSS.xlsx'
        #output_path = r'S:\Revantage\IT\FRP\Cleaned_Data/Final_loss_test.xlsx'
        log_path = r'S:/Revantage/IT/FRP/EndToEnd/LOSS/ThereCanBeOnlyOne.xlsx'
    
    else:
        sys.exit("Oh Oh thats not right......Please double check your spelling it should be (SOV or LOSS ")
        
    
    #Creating A log DataFrame toi track all changes
    log_df = pd.DataFrame(columns=['Broker', 'N_Rows'])
    
    # generating latters, indexes
    def num_to_col_letters(num):
        ''' Genrate a dictionary of alphabets and their corresponding numbers'''
        letters = ''
        while num:
            mod = (num-1) % 26
            letters += chr(mod + 65)
            num = (num-1) // 26
        return ''.join(reversed(letters))
    
    # Generating indexes from Alphabets using the previous function
    map_c = {}
    for i in range (1, 100):
        map_c[num_to_col_letters(i)] = i
    
    # creating a list of broker files    
    br_files = []
    for file in os.listdir(working_dir):
        if file.lower().endswith(".xlsx"):
            br_files.append(file)
    
    # Intersection of two lists
    def intersection(lst1, lst2): 
        lst = [value for value in lst1 if value in lst2] 
        return lst 
    
    # Creating a function to rename broker files header to match the template    
    def rename(df, _dict):
        ''' Rename the columns of broker's data to match the target Origami template'''
        h = [c for c in df.columns]
        keys_a = list(map_c.keys())
        keys_b = list(_dict.values())
        shared_keys = list(intersection(keys_b, keys_a))
        duplicates = [item for item, count in collections.Counter(shared_keys).items() if count > 1]
        unq = []
        for k in shared_keys :
            if k not in unq:
                unq.append(k)
        for key in unq:
            i = map_c[key]
            old = h[i-1]
            h[i-1] = [k for k, v in _dict.items() if v == key][0]
            new = h[i-1]
            #retained_clms.append(new)
            log.writelines(f"Renaming {old} to ----->   {new}\n")
            
        df.columns = h
        if len(duplicates) > 0:
            for d in duplicates:
                for n in range(1, shared_keys.count(d)):
                    ix = map_c[d]-1
                    #print(d, ix)
                    od = h[ix]
                    new_clm = [k for k, v in _dict.items() if v == d][n]
                    #print(od, new_clm)
                    df[new_clm] = df[od]
                    h.append(new_clm)
                    #retained_clms.append(new_clm)
                    log.writelines(f"{od} is duplicated in ->   {new_clm}\n")
            
        df.columns = h
        return df
    
    # adding broker map dictionaries for debugging
    broker_map = []
    retain_lst = []
    drop_lst = []
    total_clms = []
    total_rows = [] 
                         
    # droping Irrelevant Columns from SOVs
    def dropping_clmns(df, _dict):
        """ Thsi Function takes a dataframe, and dictionary map arguments, 
        map the datfram to the dictionary, and drop all columns that are not included in the map"""
        N_clms = []
        dr_clms =[]
        total_clms.append(len(df.columns))
        total_rows.append(len(df))
        for c in df.columns:
            if c in [d for d in _dict.keys()]:
                log.writelines(f"{c} is Retained\n")
                N_clms.append(c)
                
            else:
                dr_clms.append(c)
                log.writelines(f"{c} will be dropped\n")
        R = len(N_clms)
        D = len(dr_clms)
        retain_lst.append(R)
        drop_lst.append(D)
        log.writelines(f"Total of {R} Columns Retained \n")
        log.writelines(f"Total of {D} Columns Dropped \n")
        log.writelines(f"Retained Columns are {N_clms}\n")
        log.writelines(f"Dropped Columns are {dr_clms}\n\n\n")
        return df[N_clms]
    
    
    # Aggregate Function
    aggs_clms_ = []
    def aggregate(ag):
        """ This Funcrion perform the aggregation specified in the map sheet"""
        agg_split = str(ag).split('+')
        if len(agg_split) > 1:
            N_aggs = len(agg_split)
            aggs.append(N_aggs)
            ag = str(ag).replace(' ', '')
            
            indx = [(map_c[i]-1) for i in agg_split]
            clm_agg = [headers[c] for c in indx]
            rows = df.iloc[:,indx]
            rows = rows.fillna(0)
            for i in range (len(rows)) :
                r = [row for row in rows.iloc[i]]
                sum = 0
                for j in range(len(clm_agg)):
                    v = r[j]
                    if isinstance(v, (int, float)):
                        sum += v
                    else:
                        try :
                            sum += int(v)
                        except:
                            continue
                df.loc[i, agg_clm] = sum
            log.writelines(f"columns {clm_agg} are summed into {agg_clm}\n" )
        else :
            df.loc[:,agg_clm] = ag
            log.writelines(f"{ag} is populated in column: {agg_clm}\n")
        return df[agg_clm]
    
    
    # Mapping the broker files to the origami template
    cleaned_dfs =[]
    brokers =[]
    portfolio = []
    dfs = []
    if 'SOV' in str(file).upper():
        x = 'SOV'
    else :
        x = 'LOSS'
    # looping through the broker files and performming the mapping
    with open(f'{x}_logs.txt', 'w+') as log:
        for xl in br_files:
            name =os.path.splitext(os.path.basename(f"{working_dir}/{xl}"))[0]
            df = pd.read_excel(f"{working_dir}/{xl}")
            headers = df.columns
            for n in map_sheet.index:
                if n == name:
                    
                    log.writelines(f"\t-Broker file : {name}\n")
                    print(f"Working on : {name}.............")
                    brokers.append(name)
                    log.writelines(f"*************************************************************\n")
                    tm_ = map_sheet.loc[[n]]
                    tm = tm_.dropna(axis=1, how='all')
                    tm_dict = tm.to_dict('records')[0]
                    broker_map.append(tm_dict)
                    rename(df,tm_dict)
                    log.writelines(f"-------------------------------------------\n")
                    #print(df.columns)
                    
                    # Aggregation and constant columns
                    keys_a = list(map_c.keys())
                    keys_b = list(tm_dict.values())
                    aggreg = [item for item in keys_b if item not in keys_a]
                    dup = [item for item, count in collections.Counter(aggreg).items() if count > 1]
                    uniq = []
                    for k in aggreg :
                        if k not in uniq:
                            uniq.append(k)
                                       
                    temp_agg_clms = []
                    #clmns = []
                    #clmns.append('-'.join(temp_agg_clms))
                    aggs =[]
                    #NO_aggs = []
                                      
                    for ag in uniq:
                        agg_clm = [key for key, value in tm_dict.items() if value == ag][0]
                        temp_agg_clms.append(agg_clm)
                        aggregate(ag)
                        
                    if len(dup)> 0:
                        for d in dup:
                            for n in range(1, aggreg.count(d)):
                                agg_clm = [k for k, v in tm_dict.items() if v == d][n]
                                df[agg_clm] = aggregate(d)
                                log.writelines(f"{agg_clm} aggregated values are duplicates\n")
                    log.writelines(f"-----------------------------------------------------------------\n")
                    cleaned_df = dropping_clmns(df, tm_dict)
                    
                    portfolio.append(cleaned_df['Main/Insurance Portfolio'][0])
                    try:
                        for i in range(0, len(cleaned_df['% Sprinklered'])):
                            v = cleaned_df.loc[i,'% Sprinklered']
                            if isinstance(v, (int, float)) and (v < 100): 
                                cleaned_df.loc[i, '% Sprinklered'] = (v)*100
                            else:
                                cleaned_df.loc[i, '% Sprinklered'] = str(v).strip('%')
                    except:
                        
                        print(f"--------{name}-----------")
                    
                    if str(name) == 'Alliant - Invitation Homes SOV':
                        Regions =  cleaned_df.groupby("Invitation Homes - Region").count()
                        p_id = 600000
                        cnt = 0
                        c = 0
                        Num_clas = []
                        for count in Regions["BI/EE"]:
                            sum_c = cnt
                            cnt = sum_c+count
                            Num_clas.append(p_id +sum_c)
                        Regions["PropertyID"] = [N for N in Num_clas]
                        RegionDict = pd.Series(Regions["PropertyID"],index=Regions.index).to_dict()
                        for key in RegionDict.keys():
                            for i, v in cleaned_df["Invitation Homes - Region"].items():
                                if key.lower().strip() == v.lower().strip():
                                    data = RegionDict[v]+c
                                    base = RegionDict[v]
                                    cleaned_df.at[i,"Business ID"] = f'CA-{data}'
                                    cleaned_df.at[i, "Property ID"] = f'CA-{base}'
                                    cleaned_df.at[i, "Property"] = f'IH - {key} - Property'
                                    cleaned_df.at[i, "Location"] = f'{key}'
                                    c+=1
                            c = 0
                    aggs_clms_.append(aggs)       
                    #Running Some Statistics on data completness
                    stat_df = stats(cleaned_df, tm_)
                    dfs.append(stat_df)
                    #plot(cleaned_df, n)
                    
                    # adding some stats to the log
                    No_map_columns = []
                    No_map_columns.append(len(cleaned_df.columns))
                    cleaned_df['File_Source'] = name                   
                    #print(No_map_columns)
                    cleaned_dfs.append(cleaned_df)
                    print(f"================ {name} ===============================")
                    log.writelines(f"--------------------------------------------------\n")
                    
    #stats_ = pd.concat(dfs, sort=False)
    #sov_stats = r'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\Data_Stats\Stats_sov.xlsx'
    #loss_stats = r'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\Data_Stats\Stats_loss.xlsx'
    #if 'SOV' in working_dir:
        #stats_.to_excel(sov_stats)
    #else :
        #stats_.to_excel(loss_stats)
    
                    
    Final_Temp = template.append(cleaned_dfs, sort=False)
    
    #print(f'Writing the final Data Frame to excel')
       
    
    #Writing to logs to dataframe            
    log_df['No RetainedColumns'] = retain_lst
    log_df['No DroppedColumns'] = drop_lst
    log_df['Portfolio'] = portfolio
    log_df['Broker'] = brokers
    log_df['total received data columns'] = total_clms
    log_df['total received data rows']  = total_rows
    #log_df['Number of mapped columns'] = [len(df.columns) for df in cleaned_dfs]
    #log_df['columns Aggregated']  = ['-'.join(d) if len(d)>0 else 'Non' for d in aggs_clms_ ]
    #log_df['Template_aggrgated column']  = clmns
    if 'SOV' in working_dir:
        Final = table_mapping.overlay(Final_Temp)
        #log_df['overlayed'] = Final.groupby('File_Source').count().reset_index()['mds/Q3']
        #log_df['Un-overlayed'] = log_df['total received data rows'] - log_df['overlayed']
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
        "Normalized Construction Type","Latitude", "Longitude", 'File_Source', 'Revantage Manage']        
        
        Final = Final[columns]
    
    else :
        Final = Final_Temp
        
    log_df.to_excel(f'{log_path}')
    
    
    Final.to_excel(f'{output_path}')
    print (f" ---------- Done Successfuly! ------------------------ ")
    return Final
    #return Final_Temp.to_excel(f'{output_path}')

Mapping(*sys.argv[1:])
print("-------------------DOne Successfully-------------------------")
#test = Mapping('SOV')
#Mapping(working_dir,r'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\LOSS_MapSheet.xlsx',  r'S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\Template_loss.xlsx' ,r"S:\Revantage\IT\FRP\Cleaned_Data\Working_dir\testtttttt.xlsx" )
