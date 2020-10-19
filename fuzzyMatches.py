
import pandas as pd 
from fuzzywuzzy import fuzz 
from fuzzywuzzy import process


def checker(wrong_options,correct_options):
    names_array=[]
    ratio_array=[]    
    for wrong_option in wrong_options:
        if wrong_option in correct_options:
           names_array.append(wrong_option)
           ratio_array.append('100')
        else:   
            x=process.extract(wrong_option,correct_options,scorer=fuzz.token_set_ratio)
            names_array.append(x)
            ratio_array.append(x[-1])
    return names_array,ratio_array


str2Match = df_To_beMatched['To be Matched Column'].fillna('######').tolist()
strOptions =df_Original_List['Original Name Column'].fillna('######').tolist()


ame_match,ratio_match=checker(str2Match,strOptions)
df1 = pd.DataFrame()
df1['old_names']=pd.Series(str2Match)
df1['correct_names']=pd.Series(name_match)
df1['correct_ratio']=pd.Series(ratio_match)
df1.to_excel('matched_names.xlsx', engine='xlsxwriter')