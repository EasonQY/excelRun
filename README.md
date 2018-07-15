# excelRun
from __future__ import division
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from numpy.random import randn
from pandas import Series,DataFrame
from datetime import datetime
import xlrd,openpyxl

# Step1
xlsx_file = pd.ExcelFile('/Users/eason1986/Documents/数据科学/excelRun/test1.xlsx')

All = xlsx_file.parse('All')

# 取出有用字段
#删除无用列，保留有用列
d1 = All.drop(All.columns[:11],axis = 1,inplace = False) # colome 的axis = 1 删除0-11列
All = d1.drop(d1.columns[-1],axis = 1,inplace = False)  #删除最后一列

head = All.head() #显示前五行
oriLen = len(All)  # 未处理重复前表中记录长度
# print(oriLen)
# 获取到重复的行的行号的前20个
preLen = All[All.duplicated() == True].index[:20]
# print(preLen)

#删除掉重复的行，在原值上直接改，inplace = True
All.drop_duplicates(inplace = True)
aftLen = len(All)
# print(aftLen)

# 处理缺失值，先获取该列，将列中的“转发”、“评论”、“赞”替换掉
All[u'转发数'][All[u'转发数'] == u'转发']= '0'
All[u'评论数'][All[u'评论数'] == u'评论'] = '0'
All[u'点赞数'][All[u'点赞数'] == u'赞'] = '0' # All[u'点赞数'].replace(u'赞'，0)
# print(All.head())

#数据透视
# print(All.describe())
# print(All.dtypes)

# 为了进行数据透视 需要将对应列的数据类型转换成数值型
# 将DataFrame表中的某列数据进行转换类型

All[u'转发数'] = All[u'转发数'].astype('int64')
All[u'评论数'] = All[u'评论数'].astype('int64')
All[u'点赞数'] = All[u'点赞数'].astype('int64')
# print(All.describe())
# print(All.dtypes)

# 预处理过的表存在All.xlsx中
All.to_excel('All.xlsx',index = False)

#数据透视表
All_pivot = All.pivot_table(values = [u'转发数',u'评论数',u'点赞数',u'微博内容'],index = [u'用户名'],\
							aggfunc = {u'转发数':np.sum,u'评论数':np.sum,u'点赞数':np.sum,u'微博内容':np.size})
# print(All_pivot)

# 给该列换名称
All_pivot.rename(columns = {u'微博内容':u'当月总微博数'},inplace = True)
# print(All_pivot)
All_pivot.to_excel('All_pivot.xlsx')

# Step 2
# 以省份名做数据连接成sf_sfweibo
# 读取test1.xlsx中的sf表
sf = xlsx_file.parse('sf')
# print(sf.head())
sfweibo = xlsx_file.parse('sfweibo')
# print(sfweibo.head())

# 省份不完整，需要对两个表格切割后，进行连接
sf[u'省份前两字'] = np.nan

for i in range(len(sf[u'省份名'])):
	sf[u'省份前两字'][i] = sf[u'省份名'][i][:2]

sfweibo[u'省份前两字'] = np.nan

for i in range(len(sfweibo[u'省份名'])):
	sfweibo['省份前两字'][i] = sfweibo[u'省份名'][i][:2]

# print(sf.head())
# print(sfweibo.head())

sf.to_excel('sf.xlsx',index = False)
sfweibo.to_excel('sfweibo.xlsx',index = False)

#连接两表
sf_sfweibo = sf.merge(sfweibo,on = u'省份前两字')
# print(sf_sfweibo.head())

#获取连接后表格中需要的字段名，并重新排列
# iloc[:]选择行，[4,1,2]选择列，并按顺序显示
sf_sfweibo1 = sf_sfweibo.iloc[:,[4,1,2]]
# print(sf_sfweibo1.head())
sf_sfweibo1.to_excel('sf_sfweibo.xlsx',index = False)

# 连接sf_sfweibo 和 All_pivot两个表
sf_sfweibo = sf_sfweibo1
sf_sfweibo_All_pivot = pd.merge(sf_sfweibo,All_pivot,left_on = u'微博用户名', right_on = u'用户名',right_index = True)

# print(sf_sfweibo_All_pivot.head())
sf_sfweibo_All_pivot.to_excel('sf_sfweibo_All_pivot.xlsx',index = False)

# 与sf_sfweibo_All 做数据连接
# 处理爬取的用户基本信息表base_info
base = xlsx_file.parse('base_info')
# print(base.head())
# 将base表与sf_sfweibo_All_pivot进行连接

sf_sfweibo_All_pivot_base = base.merge(sf_sfweibo_All_pivot,left_on = u'昵称', right_on = '微博用户名')
ssapb = sf_sfweibo_All_pivot_base
# print(ssapb.head())

# 替换某列名字
ssapb.rename(columns = {u'当月总微博数_x':u'当月总微博数'},inplace = True)

# 删除其中的多余列
ssapb = ssapb.drop([u'昵称',u'当月总微博数_y'],axis = 1)
# print(ssapb.iloc[0])

# 添加一列（当月原创数 = 当月总微博数 - 当月转发数）
ssapb[u'当月原创数'] = ssapb[u'当月总微博数'] - ssapb[u'当月转发数']

# 某列同时与某段字符串连接
linkfix = '?is_ori=1&is_forward=1&is_text=1&is_pic=1&is_video=1&is_music=1&is_\
article=1&key_word=&start_time=2017-05-01&end_time=2017-05-31&is_search=1&is_searchadv=1#_0'

ssapb[u'当月博文网址'] = ssapb[u'主页链接'] + linkfix

allfix = '?profile_ftype=1&is_all=1#_0'
ssapb[u'全部博文网址'] = ssapb[u'主页链接'] + linkfix

# 计算出篇均转发/点赞/评论，并添加列
ssapb[u'篇均点赞'] = ssapb[u'点赞数']/ssapb[u'当月总微博数']
ssapb[u'篇均转发'] = ssapb[u'转发数']/ssapb[u'当月总微博数']
ssapb[u'篇均评论'] = ssapb[u'评论数']/ssapb[u'当月总微博数']
# print(ssapb.iloc[0
ssapb.to_excel('ssapb.xlsx',index = False)

# 将All表分组，获取表格的index值
gb = All.groupby(u'用户名')
gb1 = gb.size()
gbindex = gb1.index
# print(gbindex,gb1)

# 根据h指数的定义——衡量微博的数量和质量——有h篇被引用了不少于h次，分别计算转发/评论/点赞h指数
# 在记录下每个‘用户名的最大互动度max（转发+评论+点赞）’
# ascending [True,False],第一字段升序，第二字段降序
sortAllf = All.sort_values(by = [u'用户名',u'转发数'],ascending = [True,False])
sortAllc = All.sort_values(by = [u'用户名',u'评论数'],ascending = [True,False])
sortAlll = All.sort_values(by = [u'用户名',u'点赞数'],ascending = [True,False])

mm = (sortAllc,sortAllf,sortAlll)

# 将计算结果存储到一个新的DataFrame中
All_h = pd.DataFrame(np.arange(136).reshape(34,4),columns = ['fh','ch','lh','max_hdd'],index = gbindex)
fh = []
ch = []
lh = []
max_hdd = []

for j in range(len(mm)):
	for i in gbindex:
		tempdf = mm[j][mm[j][u'用户名'] == i]
		tempdf['hdd'] = tempdf[u'转发数']+ tempdf[u'评论数']+ tempdf[u'点赞数']
		max_hdd.append(tempdf['hdd'].max())
		tempdf['numf'] = range(len(tempdf))
		if j == 0:
			a = len(tempdf[tempdf[u'转发数'] >= tempdf['numf']+1])
			fh.append(a)

		elif j == 1:
			b = len(tempdf[tempdf[u'评论数'] >= tempdf['numf']+1])
			ch.append(b)

		else:
			c = len(tempdf[tempdf[u'点赞数'] >= tempdf['numf']+1])
			lh.append(c)

All_h['fh'] = fh
All_h['ch'] = ch
All_h['lh'] = lh

# 因为前面的循环一共循环了3遍，因此只要获取前34位即可：
All_h['max_hdd'] = max_hdd[:34]

#插入一个综合h指数，该指数是转发/评论/点赞h指数三个的均值
All_h.insert(3,'HS',All_h.iloc[:,:3].mean(1))
All_h.rename(columns = {'fh':u'转发h指数','ch':u'评论h指数','lh':u'点赞h指数',\
	'HS':u'综合h指数','max_hdd':u'单篇最大互动度'},inplace = True)

# print(All_h.head())
# 连接ssapb和All_h表格
ssapb_All_h = pd.merge(ssapb,All_h,left_on = u'微博用户名',right_on = u'用户名',right_index = True)

# 加一列原创率
ssapb_All_h[u'原创率'] = ssapb_All_h[u'当月原创数']/ssapb_All_h[u'当月总微博数']
# print(ssapb_All_h.iloc[0])

ssapb_All_h.to_excel('ssapb_All_h.xlsx',index = False)

# 计算相关性
# 获取原DataFrame中的几列存储到新的DataFrame中，计算综合h指数与其他分指数之间的相关性
f1 = ssapb_All_h.loc[:,[u'综合h指数',u'转发h指数',u'评论h指数',u'点赞h指数']]

# 计算f1中各列数据的相关性
corr1 = f1.corr()
# print(corr1)
corr1.to_excel('corr1.xlsx')

# 获取原DataFrame中的几列存储到新的DataFrame中，计算综合h指数与其他微博信息之间的相关性
f2 = ssapb_All_h.loc[:,[u'综合h指数',u'转发数',u'评论数',u'点赞数',u'篇均转发',u'篇均评论',u'篇均点赞']]
corr2 = f2.corr()
corr2.to_excel('corr2.xlsx')
# print(corr2)

f3 = ssapb_All_h.loc[:,[u'综合h指数',u'原创率',u'粉丝数',u'微博总数',u'单篇最大互动度']]
corr3 = f3.corr()
corr3.to_excel('corr3.xlsx') 
# print(corr3)

#重新排序列
aa = ssapb_All_h.iloc[:,[8,9,10,5,15,16,0,1,2,3,4,6,14,7,11,12,13,24,20,21,22,23,17,18,19,25]]
aa.to_excel('finally.xlsx')
# print(aa.iloc[0])

# 处理数据，处理浮点位数/转成百分位数
# 将表中的浮点类型保存至小数点后四位
f = lambda x:'%.4f'% x
# applymap(f) 作用于DataFrame中的每一个元素
# map() 作用于一个series的每一个元素
aa.ix[:,21:] = aa.ix[:,21:].applymap(f)
# astype 转换数据类型
aa.ix[:,21:] = aa.ix[:,21:].astype('float64')

# print(aa.iloc[0])

# 将原创率转成百分比形式
f1 = lambda x:'%.2f%%' % (x*100)
aa[[u'原创率']] = aa[[u'原创率']].applymap(f1)
# index = False 不限时索引
aa.to_excel('finally1.xlsx',index = False)
# print(aa.iloc[0])

# 按照综合h指数
aa.sort_values(by = u'综合h指数',ascending = False,inplace = True)
aa['rank'] = np.arange(34)+1
# print(aa['rank'].dtype)
# 要想得到‘综合h指数/排名’的列，需要将aa['rank']和aa【u'综合h指数'】进行合并成一列
# 这就要求必须连接字符串类型

aa['rank'] = aa['rank'].astype('string_')
aa[u'综合h指数'] = aa[u'综合h指数'].astype('string_')

# print(aa[u'综合h指数'].dtype)
# print(aa['rank'].dtype)

# 存储最终数据
del aa['rank']

aa[u'综合h指数'] = aa[u'综合h指数'].astype('float64')
aa.to_excel('finally2.xlsx',index = False)





















