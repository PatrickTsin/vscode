import pandas as pd
import jieba
from collections import Counter

#清洗数据
def clear(df):
    #转换时间格式
    df['yyyy_mm'] = df['comment_time'].str[0:7]
    #保留列
    df = df[['yyyy_mm', 'content', 'product_id']]
    df_g = df[~df['content'].str.contains('此用户没有填写评论|<b></b>此<b></b>用<b></b>户<b></b>没<b></b>有<b></b>填<b></b>写<b></b>评<b></b>论')]#删除无效评论
    df_g = df.groupby(['product_id', 'yyyy_mm'])['content'].apply(list)
    #按时间产品列表化评论
    df_g = df_g.apply(lambda x:','.join(x))
    df_g = df_g.reset_index()
    return df_g
#计算词频
def word_freq(row):
    #导入数据
    content = row['content']
    #切词
    seg_list = jieba.lcut(content, cut_all=False)
    #去除停用词
    cleaned_list = [x for x in seg_list if x not in list_stop]
    dict_freq = Counter(cleaned_list)#计数
    df_freq = pd.DataFrame.from_dict(dict_freq, orient='index').reset_index().rename(columns={'index':'word', 0:'frequency'})#创建DataFrame
    df_freq['product_id'] = row['product_id']
    df_freq['yyyy_mm'] = row['yyyy_mm']#写入数据
    return df_freq
#生成最终数据
def result():
    df_s = pd.DataFrame()#创建新DataFrame
    df_g = clear(df)#清洗数据
    for row in df_g.iterrows():
        row = row[1]
        df_n = word_freq(row)
        df_s = df_s.append(df_n)#按行（时间-产品）导入数据计算词频，结果返回新DataFrame
    df_s = df_s.sort_values(by=['yyyy_mm', 'product_id', 'frequency'], ascending=False)#按时间-产品排序
    df_s = df_s.reset_index()
    df_s = df_s[['product_id', 'yyyy_mm', 'word', 'frequency']]
    return df_s
#导入数据
jieba.load_userdict('user_dic.txt')#导入用户词库
list_stop = open('stop_words.txt', encoding='utf-8').read().splitlines()#导入停用词
df = pd.read_excel('comments.xlsx', dtype=str)#读取评论数据

#最终计算，生成excel
df = result()
df.to_excel('word_Freq.xlsx')