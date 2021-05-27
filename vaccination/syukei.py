# -*-coding: utf-8

import pandas as pd
import openpyxl as px
import matplotlib.pyplot as plt
import matplotlib as mpl
import seaborn as sns
import os
import shutil
import datetime
 
import eikon as ek  # the Eikon Python wrapper package
ek.set_app_key('0ed6a35e0937415eab446d3375bca7cf671d6b4c')


#Font Setting
from matplotlib.font_manager import FontProperties
import sys
if sys.platform.startswith('win'):
    FontPath= 'C:\\Windows\\Fonts\\meiryo.ttc'
elif sys.platform.startswith('darwin'):
    FontPath= '/System/Library/Fonts/ヒラノギ角ゴシック W4.ttc'
elif sys.platform.startswith('linux'):
    FontPath= '/usr/share/fonts/truetype/takao-gothic/TakaoExGothic.ttc'
jpfont = FontProperties(fname = FontPath)

def fill_flag_area(ax, flags, label=None, freq=None, **kwargs):
    """ フラグが立っている領域を塗りつぶす
    params:
        ax: Matplotlib の Axes オブジェクト
        flags: index が DatetimeIndex で dtype が bool な pandas.Series オブジェクト
        freq: 時系列データの1単位時間, 指定しない場合は flags.index.freq が使われる
              flags.index.freq が None の場合には必ず指定しなければならない
              (例: 1日単位のデータの場合) pandas.tseries.frequencies.Day(1)
    return:
        Matplotlib の Axes オブジェクト
    """
    assert flags.dtype == bool
    assert type(flags.index) == pd.DatetimeIndex
    freq = freq or flags.index.freq
    assert freq is not None
    diff = pd.Series([0] + list(flags.astype(int)) + [0]).diff().dropna()
    for start, end in zip(flags.index[diff.iloc[:-1] == 1], flags.index[diff.iloc[1:] == -1]):
        ax.axvspan(start, end + freq, label=label, **kwargs)
        label = None  # 凡例が複数表示されないようにする
    return ax

def fill_area_BOJoperation_Date(ax,flags,label = False,stack=False,alpha=0.8,**kwargs):
    """　This def paint fill area during the BOJ Operation
    params:
        ax: Matplotlib の Axes オブジェクト
        flags: index が DatetimeIndex で dtype が bool な pandas.Series オブジェクト
        freq: 時系列データの1単位時間, 指定しない場合は flags.index.freq が使われる
              flags.index.freq が None の場合には必ず指定しなければならない
              (例: 1日単位のデータの場合) pandas.tseries.frequencies.Day(1)
        stack: 積み上げグラフに対して適応したい場合はTrueとする
    return:
        Matplotlib の Axes オブジェクト
    """

    if stack:
        flags['Total']=flags.sum(axis=1)

    OpeData = [
    (datetime.datetime(2013,4,1),datetime.datetime(2016,2,1),'異次元緩和','#B0BEC5'),
    (datetime.datetime(2016,2,1),datetime.datetime(2016,9,1),'ﾏｲﾅｽ金利導入','#546E7A'),
    (datetime.datetime(2016,9,1), flags.index[len(flags)-1],'YCC','#455A64')
    ]
    bottom,top = ax.get_ylim()
    
    for Sdate,Edate,label,iro in OpeData:
        if not stack: 
            ax.axvspan(Sdate,Edate,label=label,color=iro,alpha =alpha, **kwargs)
        else:
            ax.fill_between(flags.index,flags['Total'],top,where=flags.index>=Sdate,
                        facecolor=iro, alpha=alpha,label =label)
        ax.annotate(label,xy=(Sdate + (Edate-Sdate)/2,top),size =10,color = "black",
        horizontalalignment='center')  
    return ax


# mpl_dirpath = os.path.dirname(mpl.__file__)
# # デフォルトの設定ファイルのパス
# default_config_path = os.path.join(mpl_dirpath, 'mpl-data', 'matplotlibrc')
# # カスタム設定ファイルのパス
# custom_config_path = os.path.join(mpl.get_configdir(), 'matplotlibrc')

# os.makedirs(mpl.get_configdir(), exist_ok=True)
# shutil.copyfile(default_config_path, custom_config_path)

mpl.font_manager._rebuild() 
Path = os.getcwd()

# ファイル読み込み


GenTan_borrow = pd.read_excel(Path+'\\repo2.xlsx',sheet_name = 'borrow',header =8,date_parser=1,index_col=0)
GenTan_loan = pd.read_excel(Path+'\\repo2.xlsx',sheet_name = 'loan',header =8,date_parser=1,index_col=0)
Gensaki_kai = pd.read_excel(Path+'\\repo2.xlsx',sheet_name = 'kaigensaki',header =8,date_parser=1,index_col=0)
Gensaki_uri = pd.read_excel(Path+'\\repo2.xlsx',sheet_name = 'urigensaki',header =8,date_parser=1,index_col=0)


df = pd.DataFrame({'現担レポ':GenTan_borrow['残高合計 Total']/10**12,
                   '現先取引':Gensaki_kai['残高合計 Total']/10**12},
                   index=GenTan_borrow.index)
df.dropna(axis=0,how='any',inplace=True)
fig, ax = plt.subplots(figsize=(8, 4))
ax.stackplot(df.index,df['現担レポ'],df['現先取引'],labels =df.columns)
fill_area_BOJoperation_Date(ax,df,label=True,stack=True)

plt.legend(loc ='upper left',prop =jpfont )
plt.title('レポ取引残高合計　【兆円】',FontProperties=jpfont)
plt.show()
