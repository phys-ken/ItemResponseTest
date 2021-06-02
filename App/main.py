import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import base64

# ソースコードは、https://nigimitama.hatenablog.jp/entry/2020/01/25/110921
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


def get_table_download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    csv = df.to_csv(index=True)
    # some strings <-> bytes conversions necessary here
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}">Download csv file</a>'
    return href


def create_download_link(val, filename):
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}.pdf">Download PDF file</a>'


st.title("項目特性図作成アプリ")
st.write("Pepperが配布している書式を読み込んでください。ページ下部に画像が出力されます。")
st.write("もっと下に、PDFダウンロード用のリンクが表示されます。")

excelFilePath = st.file_uploader("ファイルアップロード", type='xlsx')


if excelFilePath is not None:
    df_sheet_all = pd.read_excel(excelFilePath, sheet_name=None, index_col=0)
    bk = pd.ExcelFile(excelFilePath)

    for sheetNum in range(0, len(bk.sheet_names)-2):
        # 指定したシートをデータにする
        df_sheet_all[bk.sheet_names[sheetNum]
                     ] = df_sheet_all[bk.sheet_names[sheetNum]].fillna(99)  # NaNを99に
        df_sheet_all[bk.sheet_names[sheetNum]
                     ].columns = df_sheet_all[bk.sheet_names[sheetNum]].iloc[13]  # 行名を指定

        school = bk.sheet_names[sheetNum]
        yyyy = df_sheet_all[bk.sheet_names[sheetNum]].iat[6, 4]
        mm = df_sheet_all[bk.sheet_names[sheetNum]].iat[7, 4]
        period = df_sheet_all[bk.sheet_names[sheetNum]].iat[8, 4]
        grade = df_sheet_all[bk.sheet_names[sheetNum]].iat[9, 4]

        dfTmp = df_sheet_all[bk.sheet_names[sheetNum]].drop(
            df_sheet_all[bk.sheet_names[sheetNum]].index[range(0, 14)])
        dfTmp = dfTmp.drop(99, axis=1)
        dfTmp.insert(0, '実施月', mm)
        dfTmp.insert(0, '位置付け', period)
        dfTmp.insert(0, '実施対象', grade)
        dfTmp.insert(0, '実施年度', yyyy)
        dfTmp.insert(0, '学校名', school)
        print(school + "の重複削除前： " + str(len(dfTmp)))
        print(school + "の重複削除前： " + str(len(dfTmp.drop_duplicates())))
        st.write(school + "の重複削除前： " + str(len(dfTmp)))
        st.write(school + "の重複削除前： " + str(len(dfTmp.drop_duplicates())))

        dfTmp = dfTmp.drop_duplicates()

        if sheetNum == 0:
            dfAll = dfTmp.copy()
        else:
            dfAll = pd.concat([dfAll, dfTmp])
        print("----")
        st.write("---")

    dfAll = dfAll[dfAll['実施年度'] != 99]
    st.write("すべてのシートを取得しました。データ数は" + str(len(dfAll)) + "です")

    # 行名を変更して、採点する
    ansList = list(dfAll.columns.values[10:])
    colList = dfAll.columns.values
    qList = list(range(1, len(ansList)+1))
    colList[10:] = qList
    colList[9] = "合計点"
    dfAll.columns = colList

    dfbin = dfAll.copy()

    # Pandasで合計点の計算
    for j in range(len(dfAll)):
        count = 0
        for i in range(len(ansList)):
            if(dfAll.iat[j, 10+i] == ansList[i]):
                count = count + 1
                dfbin.iat[j, 10+i] = 1
            else:
                dfbin.iat[j, 10+i] = 0
        dfAll.iat[j, 9] = count
        dfbin.iat[j, 9] = count

    # dfAll:全データ
    # dfbin:01のデータ

    st.write(dfbin.iloc[:, 9:])

    # 点双列相関係数
    dfCorr = pd.DataFrame(dfbin.iloc[:, 9:].astype(int).corr())
    dfCorrPlot = pd.DataFrame(dfCorr.iloc[1:, 0])
    dfCorrPlot = dfCorrPlot.T

    # 項目特性図
    # 合計点が高い方からソート
    dfAll_sort = dfAll[9:].sort_values("合計点", ascending=False)

    # 5段階に分ける
    grp = 5
    size = int(dfAll_sort.shape[0] / grp) + 1
    data_slice = slice_df(dfAll_sort, size=size)

    i = grp

    pdf = PdfPages('output.pdf')
    for Q in range(1, len(ansList) + 1):
        KT = pd.DataFrame()  # １問ごとにデータを初期化
        st.markdown("---")
        for rank in range(grp):
            tmpDf = pd.DataFrame(
                data_slice[rank][Q].value_counts(normalize=True))
            tmpDf.columns = ["Rank" + str(rank+1)]
            KT = pd.concat([KT, tmpDf.T])
        KT = KT.T
        KT = KT.sort_index(ascending=True)
        KT = KT.sort_index(axis=1, ascending=False)

        fig = plt.figure()
        ax = fig.add_subplot(2, 2, 1)
        ax.set_xlabel("Rank")
        ax.set_title("Q" + str(Q) + " Ans:" +
                     str(ansList[Q-1]) + "   N = " + str(len(dfAll)))
        ax.set_ylim(0, 1)

        for i in KT.T.columns.values.tolist():
            ax.plot(KT.T[i], marker="$" + str(i) + "$",
                    markersize=7, linestyle="dashed")

        # 全体の選択割合のグラフ
        ax2 = fig.add_subplot(2, 2, 2)
        dfQ = pd.DataFrame(dfAll.iloc[:, 9+Q].value_counts(normalize=True))
        dfQ = dfQ.T
        dfQ = dfQ.sort_index(axis=1, ascending=True)
        dfQ.columns = [str(x) for x in dfQ.columns]
        ax2.bar(dfQ.columns, dfQ.iloc[0, :])
        ax2.set_ylim(bottom=0, top=1)
        ax2.set_xlabel('Choices')  # x軸ラベル
        ax2.set_title('Answer distribution  N = ' + str(len(dfAll)))  # グラフタイトル

        ax3 = fig.add_subplot(2, 1, 2)
        ax3.bar(dfCorrPlot.columns, dfCorrPlot.iloc[0, :])
        ax3.set_ylim(bottom=0, top=1)
        ax3.set_title('Corr  N = ' + str(len(dfAll)))  # グラフタイトル

        plt.subplots_adjust(wspace=0.4, hspace=0.6)
        st.pyplot(fig)
        st.write(KT)

    fignums = plt.get_fignums()
    for fignum in fignums:
        plt.figure(fignum)
        pdf.savefig()
    pdf.close()

    with open("output.pdf", "rb") as pdf_file:
        st.markdown(create_download_link(pdf_file.read(),
                    "plotData"), unsafe_allow_html=True)
    #st.markdown(get_table_download_link(dfAll), unsafe_allow_html=True)
