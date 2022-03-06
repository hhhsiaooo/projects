#!/usr/bin/env python
# coding: utf-8

# In[1]:


import numpy as np
import pandas as pd
import xlsxwriter


# In[28]:


files = ["confirmed_global.csv"]#, "deaths_global.csv", "recovered_global.csv"]

writer = pd.ExcelWriter('covid19.xlsx', engine = 'xlsxwriter')
workbook = writer.book

for f in files:
    df = pd.read_csv("https://github.com/CSSEGISandData/COVID-19/raw/master/csse_covid_19_data/csse_covid_19_time_series/time_series_covid19_" + f)
    df = pd.melt(df, id_vars=df.columns[:4], value_vars=df.columns[4:], var_name="date", value_name="count")
    
    #萃取年分資料，刪除地區欄位
    df['year'] =  '20' + df['date'].str.split('/').str.get(2)
    del df['Province/State']
    print(df)
    
    #提取台/美/日人數總和為df1
    df1 = df[df['Country/Region'].isin(['Taiwan*', 'US', 'Japan'])]
    df1 = df1[['Country/Region', 'year', 'count']]
    df1 = df1.groupby(['Country/Region', 'year']).max()
    df1.reset_index(inplace=True)
    df1 = np.array(df1).tolist()
    print(df1)
    
    #將df寫成excel工作表，再插入df1的資料
    df.to_excel(writer, f, index=False, header=False, startrow=1, startcol=0)
    worksheet = writer.sheets[f]
    format_title = workbook.add_format({'valign':'vcenter', 'align':'center', 'bold':True, 'bg_color':'cccccc', 'font_size':14})
    
    worksheet.write_row('A1', ['國家', '緯度', '經度', '日期', '人數', '年'], cell_format=format_title) #插入df標題
    worksheet.write_row('H3', ['國家','年份', '總人數'], cell_format=format_title) #插入df1標題
    worksheet.set_column('J:J', 20)
    worksheet.write_row('H4', df1[0])
    worksheet.write_row('H5', df1[1])
    worksheet.write_row('H6', df1[2])
    worksheet.write_row('H7', df1[3])
    worksheet.write_row('H8', df1[4])
    worksheet.write_row('H9', df1[5])
    worksheet.write_row('H10', df1[6])
    worksheet.write_row('H11', df1[7])
    worksheet.write_row('H12', df1[8])
    
    #長條圖
    chart = workbook.add_chart({'type':'column'})
    chart.add_series({
        'categories':f'={f}!$I$4:$I$6',
        'values':f'={f}!$J$4:$J$6',
        'name':f'={f}!$H4'
    })
    chart.add_series({
        'categories':f'={f}!$I$7:$I$9',
        'values':f'={f}!$J$7:$J$9',
        'name':f'={f}!$H$7'
    })
    chart.add_series({
        'categories':f'={f}!$I$10:$I$12',
        'values':f'={f}!$J$10:$J$12',
        'name':f'={f}!$H$10'
    })
    worksheet.insert_chart('L3', chart)
    
writer.save()
workbook.close()


# In[ ]:




