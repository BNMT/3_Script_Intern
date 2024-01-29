# the purpose for this code is that analysing sentiment (neg/neu/pos) for vnese comments although it is failed by using VADER, I have learned in general how VADER works and the disavantages in which language types (vnese vs. eng) and the breakpoint/creative thinking when I made a sentiment analysis in particular.
# the breakpoint I have already mentioned is calculating sia() for each question, taking this score as a standard point to evaluate sentiment for the answers of each question.


import pandas as pd
import numpy as np
import re
import nltk
nltk.download("vader_lexicon") # Download the lexicon
from nltk.sentiment.vader import SentimentIntensityAnalyzer


# read df
df_initial = pd.read_excel("C:/Users/VivoBook/Downloads/Greenfeed-voice-2023-final-09-01-2024.xlsx")
df = df_initial.drop(df_initial.iloc[:1, :].index.tolist())
df.reset_index(inplace = True)
df = df.drop('index', axis=1)
df_temp = df[['WE11', 'WE06', 'CBRAND05', 'CSR07', 'L&M10',
              'DEI04', 'C&B04', 'Life&Quality06', 'L&D07', 'Main11',
              'JP04', 'JP05']]

#----------------------------------------------------------------------------------------------------------------
# pre-process
# fill missing values
for col in df_temp.columns:
    for row in df_temp[col]:
        if row is None:
            df_temp.fillna('không có ý kiến', inplace=True)

# replace non-alphanumeric character
df_temp = df_temp.applymap(lambda x: re.sub(r'\W+', ' ', str(x)))

# lower text and trim text
df_temp = df_temp.applymap(lambda x: x.lower().strip() if isinstance(x, str) else x)

# last check for missing values
for col in df_temp.columns:
    missing_data = df_temp[col].isna().sum()
    missing_per = missing_data / len(df_temp) * 100
    #print('%-50s %-20s' %(f'column: {col}', f'has {missing_per}% missing data'))


#----------------------------------------------------------------------------------------------------------------
# count comment freq (used in all of Vietnam territory)
# drop bu <> Vietnam country
df_vn = pd.concat([df['Business Unit'], df_temp], axis =1)
df_vn = df_vn.drop(df_vn[(df_vn['Business Unit'] == 'Cambodia')|
                         (df_vn['Business Unit'] == 'Cambodia Farm')|
                         (df_vn['Business Unit'] == 'Cambodia Feed')|
                         (df_vn['Business Unit'] == 'Laos')|
                         (df_vn['Business Unit'] == 'Myanmar')].index)
df_vn.drop('Business Unit', axis = 1, inplace = True)

# homogenerate comment
homogenerate_comment = []
for col in df_vn.columns:
    for row in df_vn[col]:
        if (len(str(row)) < 26 and any(substring in str(row) for substring in ['ý kiến', 'y kien', 'kiến', 'kien', 'kiê', 'kie',
                                                                               'hài lòng', 'hai long', 'ha i', 'quả', 'qua',
                                                                               'ok','tốt', 'tot', 'đồng ý', 'dong y',
                                                                               'hoàn toàn', 'hoan toan', 'rất', 'rat', 'tuyệt vời', 'tuyet voi',
                                                                               'tương đối', 'tuong doi', 'bình thường', 'binh thuong', 'phù hợp', 'phu hop', 'phù', 'phu',
                                                                               'no', 'ko', 'k', 'không có', 'khong co', 'không ý', 'khong y', 'chưa', 'chua',
                                                                               'đã nói', 'da noi', 'đã góp ý', 'da gop y', 'đã ghi', 'da ghi',
                                                                               'bình luận', 'binh luan', 'bình', 'binh', 'bi nh',
                                                                               'ạ','cảm ơn', 'cám ơn', 'cam on'])) or len(str(row)) < 10:
            homogenerate_comment.append(row)
df_homo_cmt = pd.DataFrame(homogenerate_comment)

# replace values in homogenerate_comment with 'không có ý kiến'
df_vn = df_vn.applymap(lambda x: 'không có ý kiến' if x in homogenerate_comment else x)

# count comment freq
cmt_freq = []
for col in df_vn.columns:
    cmt_freq.append(df_vn[col].value_counts())
df_cmt_freq = pd.DataFrame(cmt_freq, index = ['WE11', 'WE06', 'CBRAND05', 'CSR07', 'L&M10',
                                              'DEI04', 'C&B04', 'Life&Quality06', 'L&D07', 'Main11',
                                              'JP04', 'JP05']).transpose()#.sort_index(ascending = False)
df_cmt_freq_counta = pd.DataFrame(df_cmt_freq.count(), columns = ['counta unique']).transpose()
#print(df_cmt_freq_counta)


#----------------------------------------------------------------------------------------------------------------
# sentiment comments (used in all of corporation)
sia = SentimentIntensityAnalyzer()
sentiment_group = []
for col in df_temp.columns:
    for row in df_temp[col]:
        sent = sia.polarity_scores(row)
        # print(sent['compound']
        for val in [sent['compound']]:
            if val >= 0.03:
                val = 'positive'
            elif val < -0.03:
                val = 'negative'
            else:   
                val = 'neutral'
        sentiment_group.append(val)
sentiment_group = np.array(sentiment_group).reshape(len(df_temp), 12)
sentiment_group = pd.DataFrame(sentiment_group, columns = ['WE11_SEN', 'WE06_SEN', 'CBRAND05_SEN', 'CSR07_SEN', 'L&M10_SEN',
                                                           'DEI04_SEN', 'C&B04_SEN', 'Life&Quality06_SEN', 'L&D07_SEN', 'Main11_SEN',
                                                           'JP04_SEN', 'JP05_SEN'])
df_sentiment_group = pd.concat([df, sentiment_group], axis = 1)
df_arranged_col = df_sentiment_group[['STT', 'Gender', 'Industry', 'GF Industry', 'Business Unit',
                                      'Sub-BU', 'Department EN', 'Team', 'Line Manager Name', 'Working Location',
                                      'PI01', 'PI02', 'PI03',
                                      'WE03', 'WE09', 'WE10', 'WE11', 'WE11_SEN', 'WE05', 'WE06', 'WE06_SEN',
                                      'Main04', 'CBRAND02', 'CBRAND03', 'CBRAND05', 'CBRAND05_SEN',
                                      'CSR01', 'CSR02', 'CSR04', 'CSR06', 'CSR07', 'CSR07_SEN',
                                      'L&M01', 'L&M02', 'L&M03', 'L&M04', 'L&M05', 'L&M06', 'L&M11', 'L&M12', 'L&M13', 'L&M10', 'L&M10_SEN',
                                      'DEI01', 'DEI02', 'DEI04', 'DEI04_SEN',
                                      'C&B05', 'C&B02', 'C&B03', 'C&B04', 'C&B04_SEN',
                                      'Life&Quality01', 'Life&Quality02', 'Life&Quality06', 'Life&Quality06_SEN',
                                      'L&D01', 'L&D09', 'L&D02', 'L&D03', 'L&D04', 'L&D06', 'L&D07', 'L&D07_SEN',
                                      'Main06', 'Main07', 'Main08', 'Main09', 'Main10', 'Main11', 'Main11_SEN',
                                      'JP01', 'JP02', 'JP03', 'JP04', 'JP04_SEN', 'JP05', 'JP05_SEN']]
#print(df_arranged_col)


#----------------------------------------------------------------------------------------------------------------
# analysis sentiment percentage
df_sentiment_per = df_sentiment_group.iloc[:, -12:]

matrix = []
matrix_per = []
for col in df_sentiment_per.columns:
    colu = df_sentiment_per[col].value_counts()
    matrix.append(colu)
    matrix_per.append(np.round(colu / len(df_sentiment_per) * 100, 2))
df_sen = np.array(matrix).reshape(len(df_sentiment_per.columns), 3)
df_sen = pd.DataFrame(df_sen,
                      columns = ['neutral', 'negative', 'positive'],
                      index = ['WE11_SEN', 'WE06_SEN', 'CBRAND05_SEN', 'CSR07_SEN', 'L&M10_SEN',
                               'DEI04_SEN', 'C&B04_SEN', 'Life&Quality06_SEN', 'L&D07_SEN', 'Main11_SEN',
                               'JP04_SEN', 'JP05_SEN'])
df_sen_per = np.array(matrix_per).reshape(len(df_sentiment_per.columns), 3)
df_sen_per = pd.DataFrame(df_sen_per,
                          columns = ['neutral %', 'negative %', 'positive %'],
                          index = ['WE11_SEN', 'WE06_SEN', 'CBRAND05_SEN', 'CSR07_SEN', 'L&M10_SEN',
                                   'DEI04_SEN', 'C&B04_SEN', 'Life&Quality06_SEN', 'L&D07_SEN', 'Main11_SEN',
                                   'JP04_SEN', 'JP05_SEN'])
df_pos_ratio = np.round(df_sen_per['positive %'] / (df_sen_per['positive %'] + df_sen_per['negative %']), 2)
df_pos_ratio = pd.DataFrame(df_pos_ratio, columns = ['positive ratio'])

df_sen_concat = (pd.concat([df_sen, df_sen_per, df_pos_ratio], axis = 1).sort_values(by = ['positive ratio', 'positive %', 'negative %'],
                                                                                     ascending = [False, False, True]))
df_sen_concat_desc = df_sen_concat.describe()
df_sta = pd.DataFrame([[round(np.mean(df_sen_concat['negative %']), 2), round(np.mean(df_sen_concat['positive %']), 2), round(np.mean(df_sen_concat['positive ratio']), 2)],
                       [round(np.average(df_sen_concat['negative %']), 2), round(np.average(df_sen_concat['positive %']), 2), round(np.average(df_sen_concat['positive ratio']), 2)]],
                      index = ['mean', 'average'],
                      columns = ['negative %', 'positive %', 'positive ratio'])
#print(df_sen_concat)
#print(df_sen_concat_desc)


#----------------------------------------------------------------------------------------------------------------
# print whole file - 3 sheets
with pd.ExcelWriter('GreenfeedVoice2023_NewestVer.xlsx') as writer:
    # sheet 1
    df_initial.to_excel(writer, sheet_name = 'raw', index = False)
    # sheet 2
    df_homo_cmt.to_excel(writer, sheet_name = 'VN_homogenerate_comment_values')
    #df_vn.to_excel(writer, sheet_name = 'df_vn', index = False)
    # sheet 3
    df_cmt_freq_counta.to_excel(writer, sheet_name = 'VN_comment_frequency')
    df_cmt_freq.to_excel(writer, sheet_name = 'VN_comment_frequency', startrow = len(df_cmt_freq_counta) + 3)
    # sheet 4
    df_arranged_col.to_excel(writer, sheet_name = 'sentiment', index = False)
    # sheet 5
    df_sen_concat.to_excel(writer, sheet_name = 'sentiment_statistics')
    df_sen_concat_desc.to_excel(writer, sheet_name = 'sentiment_statistics', startrow = len(df_sen_concat) + 3)
    df_sta.to_excel(writer, sheet_name = 'sentiment_statistics', startrow = len(df_sen_concat) + 14)
    # sheet 6
    df_homo_cmt.to_excel(writer, sheet_name = 'homogenerate_comment', index = False)
