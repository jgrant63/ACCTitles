import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import datetime as dt
import openpyxl

acc_titles = pd.read_excel('/Users/jakegrant/PycharmProjects/ACCTitles/ACC_Table_of_Titles.xlsx',sheet_name='Clean Data',engine='openpyxl')
gen_members = pd.read_excel('/Users/jakegrant/PycharmProjects/ACCTitles/ACC_Table_of_Titles.xlsx',sheet_name='General Membership',engine='openpyxl')
sport_members = pd.read_excel('/Users/jakegrant/PycharmProjects/ACCTitles/ACC_Table_of_Titles.xlsx',sheet_name='Sport-Specific Membership',engine='openpyxl')
sport_list = pd.read_excel('/Users/jakegrant/PycharmProjects/ACCTitles/ACC_Table_of_Titles.xlsx',sheet_name='Sports Lookup',engine='openpyxl')
template = pd.read_excel('/Users/jakegrant/PycharmProjects/ACCTitles/ACC_Table_of_Titles.xlsx',sheet_name='Template',engine='openpyxl')

acc_titles = acc_titles.drop(['Class Year','Season'],axis=1).rename(columns={"Academic Year": "Season"})
sport_list = sport_list.drop(['Season'],axis=1)
sport_list = sport_list.squeeze()
#sport_list = sport_list['Sport'].to_list()
#for i in sport_list:

sport_members = sport_members.drop(['This Year:','2021-22','NOTES','LAST','ACC?'],axis=1).rename(columns={"SCHOOL": "School","SPORT":"Sport","FIRST":"First","ACTIVE":"Active","Inclusive End":"End"})
sport_members = sport_members[sport_members['End'].notna()]
sport_members['Seasons Played'] = pd.to_numeric(sport_members['End'].str[:4]) - pd.to_numeric(sport_members['First'].str[:4]) + 1

school_titles_total = acc_titles.groupby(['School']).count().drop(['Shared','Sport'],axis=1)
school_titles_sport = acc_titles.groupby(['School','Sport']).count().drop(['Shared'],axis=1)

shared_titles = acc_titles[acc_titles['Shared'] == True].groupby(['School','Sport']).size()
solo_titles = acc_titles[acc_titles['Shared'] == False].groupby(['School','Sport']).size()
school_titles_sport = pd.concat([school_titles_sport, solo_titles, shared_titles], axis=1)

sport_years = sport_members.groupby(['School']).sum()
active_sports =  sport_years['Active']
sport_years = sport_years['Seasons Played']
historic_sports = sport_members.groupby(['School']).count()
historic_sports = historic_sports['Sport']

sport_members = sport_members.set_index(['School','Sport'])
sport_members = pd.concat([sport_members,school_titles_sport], axis=1)
sport_members = sport_members.rename(columns={'Season':'Titles',0:'Unanimous',1:'Shared'})
sport_members['Win Pct'] = sport_members['Titles'] / sport_members['Seasons Played']

gen_members = gen_members.set_index('SCHOOL')
gen_members = pd.concat([gen_members, school_titles_total, active_sports, historic_sports, sport_years], axis=1)
gen_members = gen_members.rename(columns={'FIRST':'First','Season':'Titles','LAST':'End','Sport':'Historic'})
gen_members['Win Pct'] = gen_members['Titles'] / gen_members['Seasons Played']
gen_members['Yearly Expected'] = gen_members['Win Pct']*gen_members['Active']

school_title_array = school_titles_sport['Season']
title_grid = school_title_array.unstack()

# Sport Specific Stuff
for idx1,val in sport_list.items():
    temp = sport_members.swaplevel()
    temp = temp.loc[val]

    start_end = pd.concat([temp['First'],temp['End']],axis=1)

    for idx2,data in start_end.iterrows():
        print(idx2)









out_xls_name = 'ACC_Titles_Output.xlsx'

with pd.ExcelWriter(out_xls_name) as writer:
    gen_members.to_excel(writer, sheet_name='School Summary')
    sport_members.to_excel(writer, sheet_name='Sport Specific')
    title_grid.to_excel(writer, sheet_name='Table of Titles')

# M_XC_members = sport_members[sport_members['Sport'] == 'M_XC']
# W_XC_members = sport_members[sport_members['Sport'] == 'W_XC']
# W_FH_members = sport_members[sport_members['Sport'] == 'W_FH']
# M_FB_members = sport_members[sport_members['Sport'] == 'M_FB']
# M_SOC_members = sport_members[sport_members['Sport'] == 'M_SOC']
# W_SOC_members = sport_members[sport_members['Sport'] == 'W_SOC']
# W_VB_members = sport_members[sport_members['Sport'] == 'W_VB']
# M_BB_members = sport_members[sport_members['Sport'] == 'M_BB']
# W_BB_members = sport_members[sport_members['Sport'] == 'W_BB']
# M_FEN_members = sport_members[sport_members['Sport'] == 'M_FEN']
# W_FEN_members = sport_members[sport_members['Sport'] == 'W_FEN']
# M_SD_members = sport_members[sport_members['Sport'] == 'M_SD']
# W_SD_members = sport_members[sport_members['Sport'] == 'W_SD']
# M_ITF_members = sport_members[sport_members['Sport'] == 'M_ITF']
# W_ITF_members = sport_members[sport_members['Sport'] == 'W_ITF']
# M_WR_members = sport_members[sport_members['Sport'] == 'M_WR']
# M_BASE_members = sport_members[sport_members['Sport'] == 'M_BASE']
# M_GOLF_members = sport_members[sport_members['Sport'] == 'M_GOLF']
# W_GOLF_members = sport_members[sport_members['Sport'] == 'W_GOLF']
# M_LAX_members = sport_members[sport_members['Sport'] == 'M_LAX']
# W_LAX_members = sport_members[sport_members['Sport'] == 'W_LAX']
# W_ROW_members = sport_members[sport_members['Sport'] == 'W_ROW']
# W_SOFT_members = sport_members[sport_members['Sport'] == 'W_SOFT']
# M_TEN_members = sport_members[sport_members['Sport'] == 'M_TEN']
# W_TEN_members = sport_members[sport_members['Sport'] == 'W_TEN']
# M_OTF_members = sport_members[sport_members['Sport'] == 'M_OTF']
# W_OTF_members = sport_members[sport_members['Sport'] == 'W_OTF']
# W_GYM_members = sport_members[sport_members['Sport'] == 'W_GYM']

# M_XC_active = template
# W_XC_active = template
# W_FH_active = template
# M_FB_active = template
# M_SOC_active = template
# W_SOC_active = template
# W_VB_active = template
# M_BB_active = template
# W_BB_active = template
# M_FEN_active = template
# W_FEN_active = template
# M_SD_active = template
# W_SD_active = template
# M_ITF_active = template
# W_ITF_active = template
# M_WR_active = template
# M_BASE_active = template
# M_GOLF_active = template
# W_GOLF_active = template
# M_LAX_active = template
# W_LAX_active = template
# W_ROW_active = template
# W_SOFT_active = template
# M_TEN_active = template
# W_TEN_active = template
# M_OTF_active = template
# W_OTF_active = template
# W_GYM_active = template
