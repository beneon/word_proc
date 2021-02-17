import pandas as pd
import numpy as np
import os
from word_template_update import DocxUpdate

docx_file = os.path.abspath(r'D:\SynologyDrive\口腔医院工作\kyg_科研管理\伦理委员会讨论\附件3；涉及人的生物医学研究项目伦理快速审批表.docx')
xl_file = pd.ExcelFile(os.path.abspath(r'D:\SynologyDrive\口腔医院工作\kyg_科研管理\伦理委员会讨论\2020年医学研究登记\104459202_2_临床研究信息收集_30_30.xlsx'))
df = xl_file.parse(0)
# coder convert:7、项目来源
df['7、项目来源'] = df['7、项目来源'].map({1:'佛山市科技局自筹经费课题',2:'佛山市卫生局医学科研项目',3:'广东省医学科研基金',4:'其他'})
# date, str >> date
df['application_date'] = pd.to_datetime(df['4、填写项目申报日期（如该项目在2018年7月提交申报书，请填写2018年7月，日期可以酌情选择）'])
df['review_date'] = df['application_date']+pd.to_timedelta(np.random.randint(1,4,df.shape[0]),unit='day')
df['review_year']=df['review_date'].dt.year
# only keep date, not time
df['review_date']=df['review_date'].dt.date
df['application_date']=df['application_date'].dt.date
# generate ls_num, first sort by review date, group by year, then use index in each group as ls_num (plus 1 maybe)
df = df.groupby('review_year').apply(lambda df:df.sort_values('review_date').reset_index()).drop('review_year',axis=1).reset_index()
df['ls_num']=(df['level_1']+7).apply(lambda e:f"{e:02d}")
df['ls_num'] = df['review_year'].apply(str).str.cat(df['ls_num'])
df.rename({
    '1、申请人姓名':'__co_name__',
    '2、申请人专业':"__co_spec__",
    '3、项目名称':"__xm_name__",
    '7、项目来源':"__xm_source__",
    'application_date':'__application_date__',
    'review_date':'__review_date__',
    'ls_num':'__ls_num__',
},axis=1,inplace=True)

df_dict = df[[
'__co_name__',
'__co_spec__',
'__xm_name__',
'__xm_source__',
'__application_date__',
'__review_date__',
'__ls_num__',
]]

tag_dict = df_dict.to_dict(orient='record')

docs = [DocxUpdate(dict,dict['__co_name__'],docx_file) for dict in tag_dict]
for d in docs:
    d.table_update(0)
    d.paragraph_update_id(2)
    d.doc_save()