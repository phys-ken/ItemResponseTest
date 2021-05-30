# PDF

import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

#ソースコードは、https://nigimitama.hatenablog.jp/entry/2020/01/25/110921
# 表側が順番通りの整数でないデータフレームにも対応した場合
def slice_df(df: pd.DataFrame, size: int) -> list:
    """pandas.DataFrameを行数sizeずつにスライスしてリストに入れて返す"""
    previous_index = list(df.index)
    df = df.reset_index(drop=True)
    n = df.shape[0]
    list_indices = [(i, i+size) for i in range(0, n, size)]
    df_indices = [(i, i+size-1) for i in range(0, n, size)]
    sliced_dfs = []
    for i in range(len(df_indices)):
        begin_i, end_i = df_indices[i][0], df_indices[i][1]
        begin_l, end_l = list_indices[i][0], list_indices[i][1]
        df_i = df.loc[begin_i:end_i, :]
        df_i.index = previous_index[begin_l:end_l]
        sliced_dfs += [df_i]
    return sliced_dfs




excelFilePath = "/Users/nagakurakenya/Dropbox (国立大学法人　東京学芸大学)/PER研究会/2021年度4月_標準調査問題_書式一式（問題・解答用紙含）/1_1_1_v7_物理基礎標準問題.xlsx"

df_sheet_all = pd.read_excel(excelFilePath, sheet_name=None, index_col=0)
bk = pd.ExcelFile(excelFilePath)

for sheetNum in range(0,len(bk.sheet_names)-2):
    ## 指定したシートをデータにする
    df_sheet_all[bk.sheet_names[sheetNum]] =  df_sheet_all[bk.sheet_names[sheetNum]].fillna(99) ##NaNを99に 
    df_sheet_all[bk.sheet_names[sheetNum]].columns = df_sheet_all[bk.sheet_names[sheetNum]].iloc[13] ##行名を指定

    school = bk.sheet_names[sheetNum] 
    yyyy = df_sheet_all[bk.sheet_names[sheetNum]].iat[6, 4]
    mm = df_sheet_all[bk.sheet_names[sheetNum]].iat[7, 4]
    period = df_sheet_all[bk.sheet_names[sheetNum]].iat[8, 4]
    grade = df_sheet_all[bk.sheet_names[sheetNum]].iat[9, 4]

    dfTmp = df_sheet_all[bk.sheet_names[sheetNum]].drop(df_sheet_all[bk.sheet_names[sheetNum]].index[range(0,14)])
    dfTmp =  dfTmp.drop(99,axis=1) 
    dfTmp.insert(0,'実施月' ,mm)
    dfTmp.insert(0,'位置付け' ,period)
    dfTmp.insert(0,'実施対象' ,grade)
    dfTmp.insert(0,'実施年度' ,yyyy)
    dfTmp.insert(0,'学校名' ,school)
    print(school + "の重複削除前： " + str(len(dfTmp)))
    print(school + "の重複削除前： " + str(len(dfTmp.drop_duplicates())))
    
    if sheetNum == 0:
        dfAll = dfTmp.copy()
    else:
        dfAll = pd.concat([dfAll,dfTmp])
    print("----")

dfAll = dfAll[dfAll['実施年度'] != 99]

#行明を変更して、採点する
ansList = list(dfAll.columns.values[10:])
colList = dfAll.columns.values
qList = list(range(1,len(ansList)+1))
colList[10:]=qList
colList[9] = "合計点"
dfAll.columns = colList

dfbin = dfAll.copy()

#Pandasで合計点の計算
for j in range(len(dfAll)):
    count = 0
    for i in range(len(ansList)):
        if(dfAll.iat[j,10+i]  ==  ansList[i]):
            count = count + 1
            dfbin.iat[j,10+i] = 1
        else:
            dfbin.iat[j,10+i] = 0
    dfAll.iat[j,9] = count
    dfbin.iat[j,9] = count

# dfAll:全データ
# dfbin:01のデータ

#合計点が高い方からソート
dfAll_sort = dfAll[9:].sort_values("合計点" , ascending=False)

#5段階に分ける
grp = 5
size = int( dfAll_sort.shape[0] / grp  ) +1
data_slice = slice_df(dfAll_sort , size = size)

#配列に入れる
KT = pd.DataFrame()
i = grp

pdf = PdfPages('test.pdf')

for Q in range(1,len(ansList) + 1):
    for rank in range(grp):
        KT["Rank" + str(rank+1)] = data_slice[rank][Q].value_counts(normalize = True)
    KT = KT.sort_index(ascending=True)
    KT = KT.sort_index(axis = 1 , ascending=False)

    fig , ax = plt.subplots()
    ax.set_xlabel("合計点(右ほど高得点)")
    ax.set_ylabel("選択率")
    ax.set_title("設問" + str(Q) + " 正答:" + str(ansList[Q-1]))
    ax.set_ylim(0,1)

    for i in KT.T.columns.values.tolist():
        plt.figure()
        ax.plot(KT.T[i] , marker = "$" + str(i) + "$" , markersize = 10 , c='black' ,linestyle="dashed")

fignums = plt.get_fignums()
for fignum in fignums:
    plt.figure(fignum)
    pdf.savefig()

pdf.close()


st.title("Hello streamlit")


