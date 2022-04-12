import pandas as pd
import datetime as dt
import openpyxl
import matplotlib.pyplot as plt

indiv_sport = 'BASE'
short_list = True
med_list = False
long_list = False

if short_list == True:
    sport_list = ['BASE','MBB','WBB','FB','SOFT','MTEN','WTEN','VB']
    hatch_list = ['/','|','-','+','x','o','.','*']
elif med_list == True:
    sport_list = ['BASE','MBB','WBB','FB','SOFT','MTEN','WTEN','VB','MSD','WSD','GOLF']
elif long_list == True:
    sport_list = ['BASE','MBB','WBB','FB','SOFT','MTEN','WTEN','VB','MSD','WSD','GOLF','MITF','WITF','MOTF','WOTF','MXC','WXC']
else:
    print('Choose Sports List Settings')

years = ['2016-2017','2017-2018','2018-2019','2019-2020','2020-2021','2021-2022']
schools = ['Boston College','Clemson','Duke','Florida State','Louisville','Miami','NC State','North Carolina','Notre Dame','Pittsburgh','Syracuse','Virginia','Virginia Tech']

all_sports = pd.read_excel('/Users/jakegrant/PycharmProjects/ACCTitles/Tech_ACC_Results_2016-Present_030822.xlsx',sheet_name='Data',engine='openpyxl')
all_sports = all_sports.drop(['Month','Year'],axis=1)

# All Sports, All Years, By Opponent, All Games
all_all_opp_all = all_sports.groupby(['Opponent']).size().reset_index(name='Count')
# All Sports, All Years, By Opponent, By Result
all_all_opp_res = all_sports.groupby(['Opponent','Result']).size().reset_index(name="Count")

all_all_opp_all['Wins'] = all_all_opp_res[(all_all_opp_res == "W").any(axis=1)].reset_index()['Count']
all_all_opp_all['Losses'] = all_all_opp_res[(all_all_opp_res == "L").any(axis=1)].reset_index()['Count']

# Total Row
total = pd.DataFrame({'Count': [all_all_opp_all['Count'].sum()], 'Wins': [all_all_opp_all['Wins'].sum()], 'Losses': [all_all_opp_all['Losses'].sum()], 'Opponent':['TOTAL']})
all_all_opp_all = all_all_opp_all.append(total, ignore_index = True)

all_all_opp_all['Percent'] = all_all_opp_all['Wins']/all_all_opp_all['Count']
all_all_opp_all = all_all_opp_all.set_index('Opponent')
all_all_opp_all = pd.concat([all_all_opp_all], axis=1, keys=['Total'])

detail_matrix_opp_tot = all_all_opp_all
detail_matrix_opp_tot.index.name = None

in_file = all_all_opp_all
for select_sport in sport_list:

    # FIND OUT HOW TO ADD TOTAL ROWS

    sport_specs = all_sports[(all_sports == select_sport).any(axis=1)]
    spec_all_opp_all = sport_specs.groupby(['Opponent']).size().reset_index(name='Count')
    spec_all_opp_all = spec_all_opp_all.set_index('Opponent')
    spec_all_opp_res = sport_specs.groupby(['Opponent', 'Result']).size().reset_index(name='Count')
    spec_all_opp_res = spec_all_opp_res.set_index('Opponent')

    spec_all_opp_all['Wins'] = spec_all_opp_res[(spec_all_opp_res == 'W').any(axis=1)]['Count']
    spec_all_opp_all['Losses'] = spec_all_opp_res[(spec_all_opp_res == 'L').any(axis=1)]['Count']

    total = pd.DataFrame({'Count': [spec_all_opp_all['Count'].sum()], 'Wins': [spec_all_opp_all['Wins'].sum()],
                          'Losses': [spec_all_opp_all['Losses'].sum()], 'Opponent': ['TOTAL']}).set_index('Opponent')
    spec_all_opp_all = spec_all_opp_all.append(total)


    spec_all_opp_all['Percent'] = spec_all_opp_all['Wins']/spec_all_opp_all['Count']
    spec_all_opp_all = pd.concat([spec_all_opp_all], axis=1, keys=[select_sport])

    in_file = pd.concat([in_file, spec_all_opp_all], axis=1)

detail_matrix_opp = in_file
detail_matrix_opp.index.name = None

# All Sports, All Years, By School Year, All Games
all_all_year_all = all_sports.groupby(['School Year']).size().reset_index(name='Count')
# All Sports, All Years, By School Year, By Result
all_all_year_res = all_sports.groupby(['School Year','Result']).size().reset_index(name='Count')

all_all_year_all['Wins'] = all_all_year_res[(all_all_year_res == "W").any(axis=1)].reset_index()['Count']
all_all_year_all['Losses'] = all_all_year_res[(all_all_year_res == "L").any(axis=1)].reset_index()['Count']

# Total Row
total = pd.DataFrame({'Count': [all_all_year_all['Count'].sum()], 'Wins': [all_all_year_all['Wins'].sum()], 'Losses': [all_all_year_all['Losses'].sum()], 'School Year':['TOTAL']})
all_all_year_all = all_all_year_all.append(total)

all_all_year_all['Percent'] = all_all_year_all['Wins']/all_all_year_all['Count']
all_all_year_all = all_all_year_all.set_index('School Year')
all_all_year_all = pd.concat([all_all_year_all], axis=1, keys=['Total'])

detail_matrix_year_tot = all_all_year_all
detail_matrix_year_tot.index.name = None

in_file = all_all_year_all
for select_sport in sport_list:
    sport_specs = all_sports[(all_sports == select_sport).any(axis=1)]
    spec_all_year_all = sport_specs.groupby(['School Year']).size().reset_index(name='Count')
    spec_all_year_all = spec_all_year_all.set_index('School Year')
    spec_all_year_res = sport_specs.groupby(['School Year', 'Result']).size().reset_index(name='Count')
    spec_all_year_res = spec_all_year_res.set_index('School Year')

    spec_all_year_all['Wins'] = spec_all_year_res[(spec_all_year_res == 'W').any(axis=1)]['Count']
    spec_all_year_all['Losses'] = spec_all_year_res[(spec_all_year_res == 'L').any(axis=1)]['Count']

    total = pd.DataFrame({'Count': [spec_all_year_all['Count'].sum()], 'Wins': [spec_all_year_all['Wins'].sum()],
                          'Losses': [spec_all_year_all['Losses'].sum()], 'School Year': ['TOTAL']}).set_index('School Year')
    spec_all_year_all = spec_all_year_all.append(total)

    spec_all_year_all['Percent'] = spec_all_year_all['Wins']/spec_all_year_all['Count']
    spec_all_year_all = pd.concat([spec_all_year_all], axis=1, keys=[select_sport])

    in_file = pd.concat([in_file, spec_all_year_all], axis=1)

detail_matrix_year = in_file
detail_matrix_year.index.name = None

# All Sport, All Years, By Sport, All Games
all_all_team_all = all_sports.groupby(['Sport']).size().reset_index(name='Count')
all_all_team_all = all_all_team_all.set_index('Sport')
# All Sports, All Years, By School Year, By Result
all_all_team_res = all_sports.groupby(['Sport','Result']).size().reset_index(name='Count')
all_all_team_res = all_all_team_res.set_index('Sport')

all_all_team_all['Wins'] = all_all_team_res[(all_all_team_res == "W").any(axis=1)]['Count']
all_all_team_all['Losses'] = all_all_team_res[(all_all_team_res == "L").any(axis=1)]['Count']

# Total Row
total = pd.DataFrame({'Count': [all_all_team_all['Count'].sum()], 'Wins': [all_all_team_all['Wins'].sum()], 'Losses': [all_all_team_all['Losses'].sum()], 'Sport': ['TOTAL']}).set_index('Sport')
all_all_team_all = all_all_team_all.append(total)
all_all_team_all['Percent'] = all_all_team_all['Wins']/all_all_team_all['Count']
all_all_team_all = pd.concat([all_all_team_all], axis=1, keys=['Total'])

detail_matrix_team = all_all_team_all

# All Sport, All Years, By Date, All Games
all_all_date_all = all_sports.groupby(['Date']).size().reset_index(name='Count')
all_all_date_all = all_all_date_all.set_index('Date')
# All Sports, All Years, By School Year, By Result
all_all_date_res = all_sports.groupby(['Date','Result']).size().reset_index(name='Count')
all_all_date_res = all_all_date_res.set_index('Date')

all_all_date_all['Wins'] = all_all_date_res[(all_all_date_res == "W").any(axis=1)]['Count']
all_all_date_all['Losses'] = all_all_date_res[(all_all_date_res == "L").any(axis=1)]['Count']

# Total Row
total = pd.DataFrame({'Count': [all_all_date_all['Count'].sum()], 'Wins': [all_all_date_all['Wins'].sum()], 'Losses': [all_all_date_all['Losses'].sum()], 'Date': ['TOTAL']}).set_index('Date')
all_all_date_all = all_all_date_all.append(total)
all_all_date_all['Percent'] = all_all_date_all['Wins']/all_all_date_all['Count']

# Filter Most Successful Day
filter = all_all_date_all[(all_all_date_all['Wins'] > 3) & (all_all_date_all['Percent'] > .75)]
all_all_date_all = pd.concat([all_all_date_all], axis=1, keys=['Total'])
filter = pd.concat([filter], axis=1, keys=['Total'])
detail_matrix_date = all_all_date_all
most_successful_day = filter

# Really Specific Matrix
detail_matrix_all = all_sports.groupby(['Sport','School Year','Opponent']).size().reset_index(name='Count')

# Sport-Specific Matrices
#sport_by_opponent = detail_matrix_opp.xs(indiv_sport, axis=1, level=0, drop_level=True).drop(['TOTAL']).fillna(0)
#sport_by_year = detail_matrix_year.xs(indiv_sport, axis=1, level=0, drop_level=True).drop(['TOTAL']).fillna(0)

#for select_sport in sport_list:
#    plot_df = detail_matrix_opp.xs(select_sport, axis=1, level=0, drop_level=True).drop(['TOTAL']).fillna(0)
#    # Graph Outputs - Sport by Opponent - copy this to make totals, do something similar for years
#    fig, ax = plt.subplots()
#    ax.bar(plot_df.index, plot_df['Losses'], (plot_df['Percent']+1)/2, label='Losses',color=['#98002E','#F56600','#003087','#782F40','#AD0000','#005030','#CC0000','#7BAFD4','#0C2340','#003594','#F76900','#232D4B','#630031','#000000'])
#    ax.bar(plot_df.index, plot_df['Wins'], (plot_df['Percent']+1)/2, bottom=plot_df['Losses'],label='Wins',color=['#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369','#B3A369'])
#    ax.set_ylabel('Results')
#    ax.set_title(select_sport+' Results by Opponent 2016-Present')
#    plt.xticks(rotation = 90)
#    plt.tight_layout()
#    fig.patch.set_facecolor('#C7CFD7')
#    ax.legend(title='Jake Grant, 2021')
#    plt.show()


for select_sport in sport_list:
    plot_df = detail_matrix_opp.xs(select_sport, axis=1, level=0, drop_level=True).drop(['TOTAL']).fillna(0)
    percent_df = detail_matrix_opp_tot.drop(['TOTAL'])
    percent_df.columns = percent_df.columns.droplevel()
    fig, ax = plt.subplots()

    out_plot = plt.scatter(plot_df['Wins'], plot_df['Losses'], s=percent_df['Percent']*1000,alpha=0.4, color=['#98002E','#F56600','#003087','#782F40','#AD0000','#005030','#CC0000','#7BAFD4','#0C2340','#003594','#F76900','#232D4B','#630031','#000000'],
                edgecolors="black",linewidth=1,label=schools)
    plt.xlim(left=0)
    plt.ylim(bottom=0)
    plt.gca().invert_yaxis()
    recs=[]

    plt.xlabel('Tech Wins')
    plt.ylabel('Tech Losses')
    plt.title(select_sport+' Results by Opponent 2016-Present')
    fig.patch.set_facecolor('#C7CFD7')


    plt.legend(handles=out_plot.legend_elements()[0],labels=schools,title="schools")

    plt.show()



# Scores against teams

# Create Outputs
out_xls_name = 'Tech_ACC_Results_Output.xlsx'

with pd.ExcelWriter(out_xls_name) as writer:
    detail_matrix_team.to_excel(writer, sheet_name='Totals')
    detail_matrix_year_tot.to_excel(writer, sheet_name='Years')
    detail_matrix_year.to_excel(writer, sheet_name='Years (Sport)')
    detail_matrix_opp_tot.to_excel(writer, sheet_name='Opponents')
    detail_matrix_opp.to_excel(writer, sheet_name='Opponents (Sport)')
    detail_matrix_all.to_excel(writer, sheet_name='Specific Sort')
    most_successful_day.to_excel(writer, sheet_name='Most Successful Day')

