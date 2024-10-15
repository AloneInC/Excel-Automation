from typing import final

import pandas as pd
import numpy as np
import time
from openpyxl import load_workbook

time1 = time.time()

# 指定文件地址
excel_file = "model.xlsx"

# 新建空的sheet名为数据统计
empty_df = pd.DataFrame(columns=['父sku', '本周出单', '上周出单', '数据变动', '标签'])

with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    empty_df.to_excel(writer, sheet_name='数据统计', index=False)

# 读入两周的数据
df_tw = pd.read_excel(excel_file, sheet_name='本周数据')
df_lw = pd.read_excel(excel_file, sheet_name='上周数据')

# 根据父sku进行销量统计
grouped_dftw = df_tw.groupby('父sku')['销售数量'].sum().reset_index()
grouped_dflw = df_lw.groupby('父sku')['销售数量'].sum().reset_index()

# 两个dataframe的父sku进行去重

fsku_grouped_dftw = grouped_dftw['父sku']
fsku_grouped_dflw = grouped_dflw['父sku']
fsku_temp = pd.concat([fsku_grouped_dftw, fsku_grouped_dflw])
fsku_final = fsku_temp.drop_duplicates().reset_index(drop=True)

# 查找本周和上周的sku是否在去重后的sku里，查找到后查找销售数量，没查到则填0

list_tw = []
for value in fsku_final:
    if value in fsku_grouped_dftw.values:
        list_tw.append(grouped_dftw.loc[grouped_dftw['父sku'] == value, '销售数量'].values[0])
    else:
        list_tw.append(0)

list_lw = []
for value in fsku_final:
    if value in fsku_grouped_dflw.values:
        list_lw.append(grouped_dflw.loc[grouped_dflw['父sku'] == value, '销售数量'].values[0])
    else:
        list_lw.append(0)

# 将去重后的列写入 Excel

with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    fsku_final.to_frame().to_excel(writer, sheet_name='数据统计', startcol=0, index=False)
    pd.DataFrame(list_tw, columns=['本周出单']).to_excel(writer, sheet_name='数据统计', startcol=1, index=False)
    pd.DataFrame(list_lw, columns=['上周出单']).to_excel(writer, sheet_name='数据统计', startcol=2, index=False)

df_final = pd.read_excel(excel_file, sheet_name='数据统计')

df_final['数据变动'] = df_final['本周出单'] - df_final['上周出单']
conditions = [
    ((df_final['数据变动'] > 0) & (df_final['本周出单'] != 0) & (df_final['上周出单'] != 0)),
    ((df_final['数据变动'] < 0) & (df_final['本周出单'] != 0) & (df_final['上周出单'] != 0)),
    (df_final['数据变动'] == 0),
    ((df_final['本周出单'] != 0) & (df_final['上周出单'] == 0)),
    ((df_final['上周出单'] != 0) & (df_final['本周出单'] == 0)),
]
choices = ['上升', '下降', '持平', '新出', '丸辣']

df_final['标签'] = np.select(conditions, choices, default='未定义')

with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    df_final['数据变动'].to_excel(writer, sheet_name='数据统计', startcol=3, index=False)
    df_final['标签'].to_excel(writer, sheet_name='数据统计', startcol=4, index=False)

with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    df_true = pd.read_excel(excel_file, sheet_name='数据统计')
    df_sorted = df_true.sort_values(by='数据变动', ascending=False)
    df_sorted.to_excel(writer, sheet_name='数据统计', index=False)


time2 = time.time()

print(f"耗时{(time2 - time1):.2f}s ")
print("运行成功")
input("按任意键退出")
