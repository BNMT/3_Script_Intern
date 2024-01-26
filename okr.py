import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import math


# apply file
df_0 = pd.read_excel("C:/Users/VivoBook/Downloads/OKR 22.1.xlsx", header = 6)
# drop line 2nd in df
df = df_0.drop(df_0.iloc[:1, :].index.tolist())
# rename columns
df = df.rename(columns = {'Qtr1':'Q1_2023', 'Unnamed: 10':'Q1_2023_sta',
                          'Qtr2':'Q2_2023', 'Unnamed: 12':'Q2_2023_sta',
                          'Qtr3':'Q3_2023', 'Unnamed: 14':'Q3_2023_sta',
                          'Qtr4':'Q4_2023', 'Unnamed: 16':'Q4_2023_sta','Year':'AVG_2023',
                          'Qtr1.1':'Q1_calib', 'Unnamed: 19':'Q1_calib_sta',
                          'Qtr2.1':'Q2_calib', 'Unnamed: 21':'Q2_calib_sta',
                          'Qtr3.1':'Q3_calib', 'Unnamed: 23':'Q3_calib_sta',
                          'Qtr4.1':'Q4_calib', 'Unnamed: 25':'Q4_calib_sta','Year.1':'AVG_calib'})
# ép kiểu dữ liệu cate -> numer
df[['Q1_2023', 'Q2_2023', 'Q3_2023', 'Q4_2023', 'AVG_2023', 'Q1_calib', 'Q2_calib', 'Q3_calib', 'Q4_calib', 'AVG_calib']] = df[['Q1_2023', 'Q2_2023', 'Q3_2023', 'Q4_2023', 'AVG_2023', 'Q1_calib', 'Q2_calib', 'Q3_calib', 'Q4_calib', 'AVG_calib']].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
# quantitative - định lượng (chỉ số lượng)
df_quantitative = df[['Q1_2023', 'Q2_2023', 'Q3_2023', 'Q4_2023', 'AVG_2023', 'Q1_calib', 'Q2_calib', 'Q3_calib', 'Q4_calib', 'AVG_calib']]
# qualitative - định danh (chỉ tên gọi)
df_qualitative = df[['Q1_2023_sta', 'Q2_2023_sta', 'Q3_2023_sta', 'Q4_2023_sta', 'Q1_calib_sta', 'Q2_calib_sta', 'Q3_calib_sta', 'Q4_calib_sta']]


#-------------------------------------------------------------------------------------------------------------
# number of completed people for corporation
quantitative_corpo = []
quantitative_corpo_percent = []
for col in df_quantitative.columns:
    if df_quantitative[col].astype(bool).sum(axis=0) > 0:
        quantitative_corpo.append((col, df_quantitative[col].astype(bool).sum(axis=0)))
        quantitative_corpo_percent.append(np.round((df_quantitative[col].astype(bool).sum(axis=0) / len(df_quantitative[col])) * 100,2))
df_quantitative_corpo = pd.DataFrame(np.reshape(np.array(quantitative_corpo), (len(df_quantitative.columns), 2)).transpose())
df_quantitative_corpo.columns = df_quantitative_corpo.iloc[0]
df_quantitative_corpo = df_quantitative_corpo.drop(df_quantitative_corpo.index[0])
df_quantitative_corpo_percent = pd.DataFrame(np.reshape(np.array(quantitative_corpo_percent), (len(df_quantitative.columns), 1)).transpose())
df_quantitative_corpo_percent.rename(columns = {0:'Q1_2023_%', 1:'Q2_2023_%', 2:'Q3_2023_%', 3:'Q4_2023_%', 4:'AVG_2023_%', 5:'Q1_calib_%', 6:'Q2_calib_%', 7:'Q3_calib_%', 8:'Q4_calib_%', 9:'AVG_calib_%'}, index={0:1}, inplace = True)
df_quantitative_corpo_comb = pd.concat([df_quantitative_corpo, df_quantitative_corpo_percent], axis=1)
df_quantitative_corpo_comb.rename(index={1:'number of completed people'}, inplace=True)
#print(df_quantitative_corpo_comb)


#-------------------------------------------------------------------------------------------------------------
# calculate headcount for industry/f.division
df_headcount_ind = df.groupby('Industry').agg({'Q1_2023': lambda x: len(x)})
df_headcount_ind.rename(columns={'Q1_2023':'headcount_industry'}, inplace=True)

df_headcount_fun = df.groupby('Function Division').agg({'Q1_2023': lambda x: len(x)})
df_headcount_fun.rename(columns={'Q1_2023':'headcount_f.division'}, inplace=True)


#-------------------------------------------------------------------------------------------------------------
# group by theo yêu cầu
df_3_group = df.groupby(['Function Division'])
df_1_group = df.groupby(['Industry'])

# deal with groupby industry
df_1_group_quantitative = df_1_group[['Q1_2023', 'Q2_2023', 'Q3_2023', 'Q4_2023', 'AVG_2023', 'Q1_calib', 'Q2_calib', 'Q3_calib', 'Q4_calib', 'AVG_calib']]
df_1_group_qualitative = df_1_group[['Q1_2023_sta', 'Q2_2023_sta', 'Q3_2023_sta', 'Q4_2023_sta', 'Q1_calib_sta', 'Q2_calib_sta', 'Q3_calib_sta', 'Q4_calib_sta']]

# execute quantitative
df_1_group_quantitative_sum = df_1_group_quantitative.sum()
df_1_group_quantitative_sum = pd.DataFrame(df_1_group_quantitative_sum)
group_quantitative_percent = []
for col in df_1_group_quantitative_sum.columns:
    for row in df_1_group_quantitative_sum[col]:
        group_quantitative_percent.append(np.round(row/df_1_group_quantitative_sum[col].sum()*100,2))
df_1_group_quantitative_percent = pd.DataFrame(np.reshape(np.array(group_quantitative_percent), (len(df_1_group_quantitative_sum.columns), len(df_1_group_quantitative_sum))).transpose())
df_1_group_quantitative_percent.rename(columns = {0:'Q1_2023_%', 1:'Q2_2023_%', 2:'Q3_2023_%', 3:'Q4_2023_%', 4:'AVG_2023_%', 5:'Q1_calib_%', 6:'Q2_calib_%', 7:'Q3_calib_%', 8:'Q4_calib_%', 9:'AVG_calib_%'},
                                       index={0:'Farm', 1:'Feed', 2:'Food', 3:'Head Office', 4:'Logistics', 5:'Technology'}, inplace=True)
df_1_group_quantitative_comb = pd.concat([df_1_group_quantitative_sum, df_1_group_quantitative_percent], axis=1)
#print(df_1_group_quantitative_comb)

'''limit = []
for col in df_quantitative.columns:
    # visual  
    fig = make_subplots(rows=1, cols=2)
    fig.add_trace(go.Box(x=df_quantitative[col]), row=1, col=1)
    fig.add_trace(go.Histogram(x=df_quantitative[col]), row=1, col=2)
    fig.update_xaxes(title_text=col, showgrid=False, row=1, col=1)
    fig.update_xaxes(title_text=col, row=1, col=2)
    fig.update_layout(autosize=False, width=1300, height=600, showlegend = False)
    #fig.update_layout(autosize=False, width=1200, height=900,
                      #xaxis=go.layout.XAxis(linewidth=1, mirror=True, title=col),
                      #yaxis=go.layout.YAxis(linewidth=1, mirror=True, title='Counts'))
    fig.show()
    # box-plot information
    q3, q1 = np.percentile(df_quantitative[col], [75, 25])
    iqr = q3 - q1
    lowerlimit = q1 - iqr * 1.5
    upperlimit = q3 + iqr * 1.5
    limit.extend([lowerlimit, upperlimit])
limit = pd.DataFrame(np.array(limit).reshape(len(df_quantitative.columns), 2).transpose(),
                     index = ['lower limit', 'upper limit'],
                     columns = ['Q1_2023', 'Q2_2023', 'Q3_2023', 'Q4_2023', 'AVG_2023',
                                'Q1_calib', 'Q2_calib', 'Q3_calib', 'Q4_calib', 'AVG_calib'])
df_quantitative_desc = pd.concat((df_quantitative_desc, limit), axis = 0)'''
#print(df_quantitative_desc)

# execute qualitative
qualitative = []
for col in df_qualitative.columns:
    if isinstance(df_qualitative[col], pd.Series):
        qualitative.append(df_qualitative[col].value_counts())
    elif isinstance(df_qualitative[col], np.ndarray):
        qualitative.append(df_qualitative[col].shape)
    else:
        qualitative.append(None)
df_1_qualitative_tab = pd.DataFrame(qualitative, index = ['Q1_2023_sta', 'Q2_2023_sta', 'Q3_2023_sta', 'Q4_2023_sta',
                                                        'Q1_calib_sta', 'Q2_calib_sta', 'Q3_calib_sta', 'Q4_calib_sta']).transpose()

qualitative_percent = []
for col in df_1_qualitative_tab.columns:
    for row in df_1_qualitative_tab[col]:
        qualitative_percent.append(np.round(row/len(df_qualitative)*100, 2))
df_1_quantitative_percent = pd.DataFrame(np.reshape(np.array(qualitative_percent), (len(df_1_qualitative_tab.columns), len(df_1_qualitative_tab))).transpose())
df_1_quantitative_percent.rename(columns = {0:'Q1_2023_sta_%', 1:'Q2_2023_sta_%', 2:'Q3_2023_sta_%', 3:'Q4_2023_sta_%', 4:'Q1_calib_sta_%', 5:'Q2_calib_sta_%', 6:'Q3_calib_sta_%', 7:'Q4_calib_sta_%'},
                                 index={0:'Closed', 1:'Not closed', 2:'Submited', 3:'Draft', 4:'Approved'},
                                 inplace = True)
df_1_quantitative_percent_comb = pd.concat([df_1_qualitative_tab, df_1_quantitative_percent], axis=1)
df_1_quantitative_percent_comb.reset_index(inplace=True)
df_1_quantitative_percent_comb.rename(columns={'index': 'Status'}, inplace=True)

df_1_qualitative_desc = df_1_qualitative_tab.describe()
#print(df_1_quantitative_percent_comb)


# deal with grougby f.division
df_3_group_quantitative = df_3_group[['Q1_2023', 'Q2_2023', 'Q3_2023', 'Q4_2023', 'AVG_2023',
                                      'Q1_calib', 'Q2_calib', 'Q3_calib', 'Q4_calib', 'AVG_calib']]

# execute quantitative
df_3_group_quantitative_sum = df_3_group_quantitative.sum()
group_quantitative_percent = []
for col in df_3_group_quantitative_sum.columns:
    for row in df_3_group_quantitative_sum[col]:
        group_quantitative_percent.append(np.round(row/df_3_group_quantitative_sum[col].sum()*100,2))
df_3_group_quantitative_percent = pd.DataFrame(np.reshape(np.array(group_quantitative_percent), (len(df_3_group_quantitative_sum.columns), len(df_3_group_quantitative_sum))).transpose())
df_3_group_quantitative_percent.rename(columns = {0:'Q1_2023_%', 1:'Q2_2023_%', 2:'Q3_2023_%', 3:'Q4_2023_%', 4:'AVG_2023_%', 5:'Q1_calib_%', 6:'Q2_calib_%', 7:'Q3_calib_%', 8:'Q4_calib_%', 9:'AVG_calib_%'},
                                       index={0:'Aqua VN Department', 1:'BOD', 2:'Commercial', 3:'Farm Department', 4:'Feed VN Department', 5:'Feed-Farm SEA Department', 6:'Finance - Accounting', 7:'Food Department', 8:'Food Functional Department', 9:'General Management', 10:'GreenFeed BioLab', 11:'Human Resource - Workplace Services', 12:'Information and Digital Technology', 13:'Internal Audit', 14:'Legal', 15:'Marketing', 16:'Procurement', 17:'Production Operation', 18:'Projects', 19:'QD Trans Department', 20:'QDTek - ExcelTech Department', 21:'Research & Development', 22:'Supply Chain', 23:'Sustainability and Branding', 24:'Technical Consulting', 25:'Technology - Engineering', 26:'Technology Solution & Services'},
                                       inplace = True)
df_3_group_quantitative_comb = pd.concat([df_3_group_quantitative_sum, df_3_group_quantitative_percent], axis=1)
#print(df_3_group_quantitative_comb)


# qualitative
df_qualitative_ind = pd.concat([df['Industry'], df_qualitative], axis=1)
df_qualitative_ind = df_qualitative_ind.melt(id_vars=['Industry'], var_name='Quarter', value_name='Status')
df_qualitative_ind = df_qualitative_ind.groupby(['Industry', 'Quarter', 'Status']).size().reset_index(name='Count')
df_qualitative_ind = df_qualitative_ind.pivot_table(index=['Industry', 'Status'], columns='Quarter', values='Count', fill_value=0)
df_qualitative_ind = df_qualitative_ind.reset_index()

ind_percent = []
for col in df_qualitative_ind.loc[:, 'Q1_2023_sta':]:
    for row in df_qualitative_ind[col]:
        ind_percent.append(np.round(row/df_qualitative_ind[col].sum()*100,2))
df_qualitative_ind_percent = pd.DataFrame(np.reshape(np.array(ind_percent), (len(df_qualitative_ind.columns)-2, len(df_qualitative_ind))).transpose())
df_qualitative_ind_percent.rename(columns = {0:'Q1_2023_sta_%', 1:'Q2_2023_sta_%', 2:'Q3_2023_sta_%', 3:'Q4_2023_sta_%', 4:'Q1_calib_sta_%', 5:'Q2_calib_sta_%', 6:'Q3_calib_sta_%', 7:'Q4_calib_sta_%'}, inplace = True)
df_qualitative_ind_percent_comb = pd.concat([df_qualitative_ind, df_qualitative_ind_percent], axis=1)
#print(df_qualitative_ind_percent_comb)


df_qualitative_fun = pd.concat([df['Function Division'], df_qualitative], axis=1)
df_qualitative_fun = df_qualitative_fun.melt(id_vars=['Function Division'], var_name='Quarter', value_name='Status')
df_qualitative_fun = df_qualitative_fun.groupby(['Function Division', 'Quarter', 'Status']).size().reset_index(name='Count')
df_qualitative_fun = df_qualitative_fun.pivot_table(index=['Function Division', 'Status'], columns='Quarter', values='Count', fill_value=0)
df_qualitative_fun = df_qualitative_fun.reset_index()

fun_percent = []
for col in df_qualitative_fun.loc[:, 'Q1_2023_sta':]:
    for row in df_qualitative_fun[col]:
        fun_percent.append(np.round(row/df_qualitative_fun[col].sum()*100,2))
df_qualitative_fun_percent = pd.DataFrame(np.reshape(np.array(fun_percent), (len(df_qualitative_fun.columns)-2, len(df_qualitative_fun))).transpose())
df_qualitative_fun_percent.rename(columns = {0:'Q1_2023_sta_%', 1:'Q2_2023_sta_%', 2:'Q3_2023_sta_%', 3:'Q4_2023_sta_%', 4:'Q1_calib_sta_%', 5:'Q2_calib_sta_%', 6:'Q3_calib_sta_%', 7:'Q4_calib_sta_%'}, inplace = True)
df_qualitative_fun_percent_comb = pd.concat([df_qualitative_fun, df_qualitative_fun_percent], axis=1)
#print(df_qualitative_fun)


#-------------------------------------------------------------------------------------------------------------
# calculate percent and total amount people for each industry/f.division
df_quantitative_ind = pd.concat([df['Industry'], df_quantitative], axis=1)
df_quantitative_ind_percent = df_quantitative_ind.groupby('Industry').agg({'Q1_2023': lambda x: np.round(100-((x == 0).sum()/len(x) * 100),2),
                                             'Q2_2023': lambda x: np.round(100-((x == 0).sum()/len(x)) * 100,2),
                                             'Q3_2023': lambda x: np.round(100-((x == 0).sum()/len(x)) * 100,2),
                                             'Q4_2023': lambda x: np.round(100-((x == 0).sum() / len(x)) * 100,2),
                                             'AVG_2023': lambda x: np.round(100-((x == 0).sum() / len(x)) * 100,2),
                                             'Q1_calib': lambda x: np.round(100-((x == 0).sum() / len(x)) * 100,2),
                                             'Q2_calib': lambda x: np.round(100-((x == 0).sum() / len(x)) * 100,2),
                                             'Q3_calib': lambda x: np.round(100-((x == 0).sum() / len(x)) * 100,2),
                                             'Q4_calib': lambda x: np.round(100-((x == 0).sum() / len(x)) * 100,2),
                                             'AVG_calib': lambda x: np.round(100-((x == 0).sum() / len(x)) * 100,2)})
df_quantitative_ind_percent.rename(columns = {'Q1_2023':'Q1_2023_%_of_each_ind', 'Q2_2023':'Q2_2023_%_of_each_ind', 'Q3_2023':'Q3_2023_%_of_each_ind', 'Q4_2023':'Q4_2023_%_of_each_ind', 'AVG_2023':'AVG_2023_%_of_each_ind', 'Q1_calib':'Q1_calib_%_of_each_ind', 'Q2_calib':'Q2_calib_%_of_each_ind', 'Q3_calib':'Q3_calib_%_of_each_ind', 'Q4_calib':'Q4_calib_%_of_each_ind', 'AVG_calib':'AVG_calib_%_of_each_ind'}, inplace=True)

df_quantitative_ind_people = df_quantitative_ind.groupby('Industry').agg({'Q1_2023': lambda x: (x != 0).sum(),
                                             'Q2_2023': lambda x: (x != 0).sum(),
                                             'Q3_2023': lambda x: (x != 0).sum(),
                                             'Q4_2023': lambda x: (x != 0).sum(),
                                             'AVG_2023': lambda x: (x != 0).sum(),
                                             'Q1_calib': lambda x: (x != 0).sum(),
                                             'Q2_calib': lambda x: (x != 0).sum(),
                                             'Q3_calib': lambda x: (x != 0).sum(),
                                             'Q4_calib': lambda x: (x != 0).sum(),
                                             'AVG_calib': lambda x: (x != 0).sum()})
df_quantitative_ind_people.rename(columns = {'Q1_2023':'Q1_2023_of_people', 'Q2_2023':'Q2_2023_of_people', 'Q3_2023':'Q3_2023_of_people', 'Q4_2023':'Q4_2023_of_people', 'AVG_2023':'AVG_2023_of_people', 'Q1_calib':'Q1_calib_of_people', 'Q2_calib':'Q2_calib_of_people', 'Q3_calib':'Q3_calib_of_people', 'Q4_calib':'Q4_calib_of_people', 'AVG_calib':'AVG_calib_of_people'}, inplace=True)
df_quantitative_ind_people_comb = pd.concat([df_quantitative_ind_people, df_quantitative_ind_percent], axis=1)
#print(df_quantitative_ind_people_comb)


df_quantitative_fun = pd.concat([df['Function Division'], df_quantitative], axis=1)
df_quantitative_fun_percent = df_quantitative_fun.groupby('Function Division').agg({'Q1_2023': lambda x: np.round(100-(x == 0).sum()/len(x) * 100,2),
                                             'Q2_2023': lambda x: np.round(100-(x == 0).sum()/len(x) * 100,2),
                                             'Q3_2023': lambda x: np.round(100-(x == 0).sum()/len(x) * 100,2),
                                             'Q4_2023': lambda x: np.round(100-(x == 0).sum() / len(x) * 100,2),
                                             'AVG_2023': lambda x: np.round(100-(x == 0).sum() / len(x) * 100,2),
                                             'Q1_calib': lambda x: np.round(100-(x == 0).sum() / len(x) * 100,2),
                                             'Q2_calib': lambda x: np.round(100-(x == 0).sum() / len(x) * 100,2),
                                             'Q3_calib': lambda x: np.round(100-(x == 0).sum() / len(x) * 100,2),
                                             'Q4_calib': lambda x: np.round(100-(x == 0).sum() / len(x) * 100,2),
                                             'AVG_calib': lambda x: np.round(100-(x == 0).sum() / len(x) * 100,2)})
df_quantitative_fun_percent.rename(columns = {'Q1_2023':'Q1_2023_%_of_each_f.division', 'Q2_2023':'Q2_2023_%_of_each_f.division', 'Q3_2023':'Q3_2023_%_of_each_f.division', 'Q4_2023':'Q4_2023_%_of_each_f.division', 'AVG_2023':'AVG_2023_%_of_each_f.division', 'Q1_calib':'Q1_calib_%_of_each_f.division', 'Q2_calib':'Q2_calib_%_of_each_f.division', 'Q3_calib':'Q3_calib_%_of_each_f.division', 'Q4_calib':'Q4_calib_%_of_each_f.division', 'AVG_calib':'AVG_calib_%_of_each_f.division'}, inplace=True)

df_quantitative_fun_people = df_quantitative_fun.groupby('Function Division').agg({'Q1_2023': lambda x: (x != 0).sum(),
                                             'Q2_2023': lambda x: (x != 0).sum(),
                                             'Q3_2023': lambda x: (x != 0).sum(),
                                             'Q4_2023': lambda x: (x != 0).sum(),
                                             'AVG_2023': lambda x: (x != 0).sum(),
                                             'Q1_calib': lambda x: (x != 0).sum(),
                                             'Q2_calib': lambda x: (x != 0).sum(),
                                             'Q3_calib': lambda x: (x != 0).sum(),
                                             'Q4_calib': lambda x: (x != 0).sum(),
                                             'AVG_calib': lambda x: (x != 0).sum()})
df_quantitative_fun_people.rename(columns = {'Q1_2023':'Q1_2023_of_people', 'Q2_2023':'Q2_2023_of_people', 'Q3_2023':'Q3_2023_of_people', 'Q4_2023':'Q4_2023_of_people', 'AVG_2023':'AVG_2023_of_people', 'Q1_calib':'Q1_calib_of_people', 'Q2_calib':'Q2_calib_of_people', 'Q3_calib':'Q3_calib_of_people', 'Q4_calib':'Q4_calib_of_people', 'AVG_calib':'AVG_calib_of_people'}, inplace=True)
df_quantitative_fun_people_comb = pd.concat([df_quantitative_fun_people, df_quantitative_fun_percent], axis=1)
#print(df_quantitative_fun)





df_qualitative_ind_each_1 = np.round(df_qualitative_ind[df_qualitative_ind['Status'] == 'Closed'].set_index('Industry').loc[['Farm', 'Feed', 'Food', 'Head Office', 'Logistics', 'Technology'], 'Q1_2023_sta':].div(df_headcount_ind['headcount_industry'], axis=0)*100,2)
df_qualitative_ind_each_2 = np.round(df_qualitative_ind[df_qualitative_ind['Status'] == 'Not closed'].set_index('Industry').loc[['Farm', 'Feed', 'Food', 'Head Office', 'Logistics', 'Technology'], 'Q1_2023_sta':].div(df_headcount_ind['headcount_industry'], axis=0)*100,2)
df_qualitative_ind_each = pd.concat([df_qualitative_ind_each_1, df_qualitative_ind_each_2]).drop(['Q1_calib_sta', 'Q2_calib_sta', 'Q3_calib_sta', 'Q4_calib_sta'], axis=1)
#print(df_qualitative_ind_each)


df_qualitative_ind_each_calib_1 = np.round(df_qualitative_ind[df_qualitative_ind['Status'] == 'Submited'].set_index('Industry').loc[['Farm', 'Feed', 'Food', 'Head Office', 'Technology'], 'Q1_2023_sta':].div(df_quantitative_ind_people.loc[:, 'Q4_2023_of_people'], axis=0)*100,2)
df_qualitative_ind_each_calib_2 = np.round(df_qualitative_ind[df_qualitative_ind['Status'] == 'Draft'].set_index('Industry').loc[['Farm', 'Feed', 'Food', 'Head Office', 'Logistics', 'Technology'], 'Q1_2023_sta':].div(df_quantitative_ind_people.loc[:, 'Q4_2023_of_people'], axis=0)*100,2)
df_qualitative_ind_each_calib_3 = np.round(df_qualitative_ind[df_qualitative_ind['Status'] == 'Approved'].set_index('Industry').loc[['Farm', 'Feed', 'Food', 'Head Office', 'Technology'], 'Q1_2023_sta':].div(df_quantitative_ind_people.loc[:, 'Q4_2023_of_people'], axis=0)*100,2)
df_qualitative_ind_each_calib = pd.concat([df_qualitative_ind_each_calib_1, df_qualitative_ind_each_calib_2, df_qualitative_ind_each_calib_3]).drop(['Q1_2023_sta', 'Q2_2023_sta', 'Q3_2023_sta', 'Q4_2023_sta'], axis=1)
#print(df_qualitative_ind_each_calib)





df_qualitative_fun_each_1 = np.round(df_qualitative_fun[df_qualitative_fun['Status'] == 'Closed'].set_index('Function Division').loc[['Aqua VN Department', 'BOD', 'Commercial', 'Farm Department', 'Feed VN Department', 'Feed-Farm SEA Department', 'Finance - Accounting', 'Food Department', 'Food Functional Department', 'General Management', 'GreenFeed BioLab', 'Human Resource - Workplace Services', 'Information and Digital Technology', 'Internal Audit', 'Legal', 'Marketing', 'Procurement', 'Production Operation', 'QD Trans Department', 'QDTek - ExcelTech Department', 'Research & Development', 'Supply Chain', 'Sustainability and Branding', 'Technical Consulting', 'Technology - Engineering', 'Technology Solution & Services'], 'Q1_2023_sta':].div(df_headcount_fun['headcount_f.division'], axis=0)*100,2)
df_qualitative_fun_each_2 = np.round(df_qualitative_fun[df_qualitative_fun['Status'] == 'Not closed'].set_index('Function Division').loc[['Aqua VN Department', 'BOD', 'Commercial', 'Farm Department', 'Feed VN Department', 'Feed-Farm SEA Department', 'Finance - Accounting', 'Food Department', 'Human Resource - Workplace Services', 'Information and Digital Technology', 'Internal Audit', 'Legal', 'Marketing', 'Procurement', 'Production Operation', 'QD Trans Department', 'QDTek - ExcelTech Department', 'Supply Chain', 'Sustainability and Branding', 'Technical Consulting', 'Technology - Engineering', 'Technology Solution & Services'], 'Q1_2023_sta':].div(df_headcount_fun['headcount_f.division'], axis=0)*100,2)
df_qualitative_fun_each = pd.concat([df_qualitative_fun_each_1, df_qualitative_fun_each_2]).drop(['Q1_calib_sta', 'Q2_calib_sta', 'Q3_calib_sta', 'Q4_calib_sta'], axis=1)
#print(df_qualitative_fun_each)


df_qualitative_fun_each_calib_1 = np.round(df_qualitative_fun[df_qualitative_fun['Status'] == 'Submited'].set_index('Function Division').loc[['Aqua VN Department', 'Commercial', 'Farm Department', 'Feed VN Department', 'Feed-Farm SEA Department', 'Finance - Accounting', 'Food Department', 'Human Resource - Workplace Services', 'Marketing', 'Procurement', 'Production Operation', 'QDTek - ExcelTech Department', 'Sustainability and Branding'], 'Q1_2023_sta':].div(df_quantitative_fun_people.loc[:, 'Q4_2023_of_people'], axis=0)*100,2)
df_qualitative_fun_each_calib_2 = np.round(df_qualitative_fun[df_qualitative_fun['Status'] == 'Draft'].set_index('Function Division').loc[['Aqua VN Department', 'Commercial', 'Farm Department', 'Feed VN Department', 'Feed-Farm SEA Department', 'Finance - Accounting', 'Food Department', 'GreenFeed BioLab', 'Human Resource - Workplace Services', 'Information and Digital Technology', 'Internal Audit', 'Legal', 'Marketing', 'Procurement', 'Production Operation'], 'Q1_2023_sta':].div(df_quantitative_fun_people.loc[:, 'Q4_2023_of_people'], axis=0)*100,2)
df_qualitative_fun_each_calib_3 = np.round(df_qualitative_fun[df_qualitative_fun['Status'] == 'Approved'].set_index('Function Division').loc[['Farm Department', 'Feed VN Department', 'Finance - Accounting', 'Food Department', 'Human Resource - Workplace Services', 'Internal Audit', 'Procurement', 'Production Operation', 'QDTek - ExcelTech Department'], 'Q1_2023_sta':].div(df_quantitative_fun_people.loc[:, 'Q4_2023_of_people'], axis=0)*100,2)
df_qualitative_fun_each_calib = pd.concat([df_qualitative_fun_each_calib_1, df_qualitative_fun_each_calib_2, df_qualitative_fun_each_calib_3]).drop(['Q1_2023_sta', 'Q2_2023_sta', 'Q3_2023_sta', 'Q4_2023_sta'], axis=1)
#print(df_qualitative_fun_each_calib)


#-------------------------------------------------------------------------------------------------------------
df_updated = df.copy()
df_updated.loc[(df_updated['Q1_calib_sta'] != 'Approved'), 'Q1_calib'] = df['Q1_2023']
df_updated.loc[(df_updated['Q2_calib_sta'] != 'Approved'), 'Q2_calib'] = df['Q2_2023']
df_updated.loc[(df_updated['Q3_calib_sta'] != 'Approved'), 'Q3_calib'] = df['Q3_2023']
df_updated.loc[(df_updated['Q4_calib_sta'] != 'Approved'), 'Q4_calib'] = df['Q4_2023']
#print(df_updated)





#-------------------------------------------------------------------------------------------------------------
with pd.ExcelWriter('okr.xlsx') as writer:
    # sheet 1
    df_0.to_excel(writer, sheet_name='raw', index=False)
    # how to export dataframe groupby to excel file
    #df_group.apply(lambda x: x.reset_index(drop=True)).to_excel(writer, sheet_name='groupby', index=False)
    #with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    #   df_1_group_quantitative_sum.to_excel(writer, sheet_name='quantitative', startrow=len(df_quantitative_corpo_comb)+3)

    # sheet 2 - sum quantitative
    df_quantitative_corpo_comb.to_excel(writer, sheet_name='quantitative', startcol=2)

    df_headcount_ind.to_excel(writer, sheet_name='quantitative', startrow=5)
    df_1_group_quantitative_comb.to_excel(writer, sheet_name='quantitative', startrow=5, startcol=2)
    df_quantitative_ind_people_comb.to_excel(writer, sheet_name='quantitative', startrow=12, startcol=2)

    df_headcount_fun.to_excel(writer, sheet_name='quantitative', startrow=21)
    df_3_group_quantitative_comb.to_excel(writer, sheet_name='quantitative', startrow=21, startcol=2)
    df_quantitative_fun_people_comb.to_excel(writer, sheet_name='quantitative', startrow=49, startcol=2)

    # sheet 3 - count qualitative
    df_1_quantitative_percent_comb.to_excel(writer, sheet_name='qualitative')

    df_qualitative_ind_percent_comb.to_excel(writer, sheet_name='qualitative', startrow=8)
    df_qualitative_ind_each.to_excel(writer, sheet_name='qualitative', startrow=8, startcol=21)
    df_qualitative_ind_each_calib.to_excel(writer, sheet_name='qualitative', startrow=8, startcol=28)

    df_qualitative_fun_percent_comb.to_excel(writer, sheet_name='qualitative', startrow=37)
    df_qualitative_fun_each.to_excel(writer, sheet_name='qualitative', startrow=37, startcol=21)
    df_qualitative_fun_each_calib.to_excel(writer, sheet_name='qualitative', startrow=37, startcol=28)

    # sheet 4
    df_updated.to_excel(writer, sheet_name='updated calib')

'''
# pie chart for industry
cols = 2
rows = math.ceil(len(df_qualitative_ind) / cols)
fig = make_subplots(rows=rows, cols=cols, specs=[[{'type':'domain'}]*cols]*rows)
for i, row in enumerate(df_qualitative_ind.iterrows()):
    fig.add_trace(px.pie(row[1], names=row[1].index, values=row[1].values).data[0],
                  row=(i // cols) + 1, col=(i % cols) + 1)
fig.update_layout(height=5500, width=3300, title_text='GF_Industry')
fig.show()


# pie chart for function division
cols = 7
rows = math.ceil(len(df_qualitative_fun) / cols)
fig = make_subplots(rows=rows, cols=cols, specs=[[{'type':'domain'}]*cols]*rows)
for i, row in enumerate(df_qualitative_fun.iterrows()):
    fig.add_trace(px.pie(row[1], names=row[1].index, values=row[1].values).data[0],
                  row=(i // cols) + 1, col=(i % cols) + 1)
fig.update_layout(height=3500, width=2500, title_text='GF_Function Division')
fig.show()
'''





