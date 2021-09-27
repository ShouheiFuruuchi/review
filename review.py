#このプログラムは店別品番別実績を自動ダウンロードを行う

#----------------------------------------------------------------------------------------------


import time
import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ChromeOptions
import datetime
import os
import glob
import shutil
from operator import itemgetter
import tes
import datetime
import pandas as pd
import re
import openpyxl as pyxl
import statsmodels.formula.api as smf
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split #検証用データセットに分割
from sklearn.model_selection import cross_val_score #検証用データセットの作成
from sklearn.linear_model import LinearRegression #線形回帰用
import sklearn.model_selection #モデルの評価に関して

from sklearn.metrics import mean_squared_error

#このプログラムは店別品番別実績を自動ダウンロードを行う



#ーーーーーーーーーー前回データの削除ーーーーーーーーーーーーー
folders = [0,1,2,3,4,5,6]
no = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]

w_day = datetime.datetime.today()

wd_no = w_day.weekday()#曜日Noを指定

main_dr = 'C:/Users/fun-f/Desktop/myfile'

print(wd_no)

to_file_path = str(main_dr) + '/' + str(wd_no)#drpathの指定

  #ーーーーーー曜日別商品実績ファイルクリアーーーーーーーーーー
  
if wd_no == 0:# 月曜日⇒0 火曜日⇒ 1 水曜日⇒ 2 木曜日⇒ 3 金曜日⇒ 4 土曜日⇒ 5 日曜日⇒ 6
  cl_sheet = pd.read_excel('C:/Users/fun-f/Desktop/myfile/クリアBOOK.xlsx')

  cl_df =pd.DataFrame(cl_sheet)
  for fd in folders:
    print(fd)
    for i in no:
      
      del_path = 'C:/Users/fun-f/Desktop/myfile/'+str(fd)+'/'+str(i)+'商品実績.xlsx'
      print(del_path)
      cl_df.to_excel(del_path)
      
  #ーーーーーーーーー実績ファイルクリアーーーーーーーーーーーーー
  
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/0/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/1/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/2/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/3/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/4/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/5/実績/実績.xlsx')
  cl_df.to_excel('C:/Users/fun-f/Desktop/myfile/6/実績/実績.xlsx')

  print('success')#削除完了
  
else:  
  print('Non Success!!')#削除ファイルなし

#ーーーーーーーー前回ダウンロードファイル削除ーーーーーーーーーー
#  　　　　　　　　一部省略
  
#ーーーーーーー今日の日付設定ーーーーーーーーー

fold = 'C:/Users/fun-f/Downloads'


todaytime = datetime.date.today()
tod = '{0:20%y%m%d}'.format(todaytime)#今日の日付(西暦)


#ーーーーーーー販売NETスクレイピングーーーーーーーーーーー

# 　　　　　　　　※一部省略


#ーーーーーショッパー抜きのP率実績ーーーーーーーー

import os
import shutil
import datetime
import requests
import schedule
import pyautogui
import time
import path


#店舗リスト・パス

kasiwa = 'C:/Users/fun-f/Desktop/myfile/dataf/柏.csv'
tiba = 'C:/Users/fun-f/Desktop/myfile/dataf/千葉.csv'
yokohama = 'C:/Users/fun-f/Desktop/myfile/dataf/横浜.csv'
isesaki = 'C:/Users/fun-f/Desktop/myfile/dataf/伊勢崎.csv'
gihu = 'C:/Users/fun-f/Desktop/myfile/dataf/岐阜.csv'
nagamachi = 'C:/Users/fun-f/Desktop/myfile/dataf/長町.csv'
hunabasi = 'C:/Users/fun-f/Desktop/myfile/dataf/船橋.csv'
hujimi = 'C:/Users/fun-f/Desktop/myfile/dataf/富士見.csv'
reiku = 'C:/Users/fun-f/Desktop/myfile/dataf/レイク.csv'
ebina = 'C:/Users/fun-f/Desktop/myfile/dataf/海老名.csv'
musasi = 'C:/Users/fun-f/Desktop/myfile/dataf/むさし.csv'
hiratuka = 'C:/Users/fun-f/Desktop/myfile/dataf/平塚.csv'
natori = 'C:/Users/fun-f/Desktop/myfile/dataf/名取.csv'
otaka = 'C:/Users/fun-f/Desktop/myfile/dataf/大高.csv'
togocyo = 'C:/Users/fun-f/Desktop/myfile/dataf/東郷町.csv'
ota = 'C:/Users/fun-f/Desktop/myfile/dataf/太田.csv'
mito = 'C:/Users/fun-f/Desktop/myfile/dataf/水戸.csv'
expo = 'C:/Users/fun-f/Desktop/myfile/dataf/EXPO.csv'
kawasaki = 'C:/Users/fun-f/Desktop/myfile/dataf/川崎.csv'
sinmisato = 'C:/Users/fun-f/Desktop/myfile/dataf/新三郷.csv'
makuhari = 'C:/Users/fun-f/Desktop/myfile/dataf/幕張.csv'
all_sp = 'C:/Users/fun-f/Desktop/myfile/dataf/全店.csv'

no = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]


#no = [1]

shops_d = {1:"柏",2:"千葉",3:"横浜",4:"伊勢崎",5:"岐阜",6:"長町",7:"船橋",8:"富士見",9:"レイクタウン",10:"海老名",11:"むさし村山",12:"平塚",13:"名取",14:"大高",15:"東郷町",16:"太田",17:"水戸",18:"EXPO",19:"川崎",20:"新三郷",21:"幕張新都心",22:"全店"}

#shops_d = {1:kasiwa,2:tiba,3:yokohama,4:isesaki,5:gihu,6:nagamachi,7:hunabasi,8:hujimi,9:reiku,10:ebina,11:musasi,12:hiratuka,13:natori,14:otaka,15:togocyo,16:ota,17:mito,18:expo,19:kawasaki,20:sinmisato}

shops_l = [kasiwa,tiba,yokohama,isesaki,gihu,nagamachi,hunabasi,hujimi,reiku,ebina,musasi,hiratuka,natori,otaka,togocyo,ota,mito,expo,kawasaki,sinmisato,makuhari,all_sp]#店舗リスト

output_file = "C:/Users/fun-f/Desktop/myfile/Set率集計.xlsx"


#ーーーー曜日Noとto_file_pathの設定ーーーーーーーーー

w_day = datetime.datetime.today()

wd_no = w_day.weekday()#曜日Noを指定

main_dr = 'C:/Users/fun-f/Desktop/myfile'

to_file_path = str(main_dr) + '/' + str(wd_no)#drpathの指定

print(to_file_path)


#-----柏------

ln = 0 #no加算値(List_No)
data_2 = pd.read_csv('C:/Users/fun-f/Desktop/myfile/売上実績/売上実績.csv', encoding='SHIFT-JIS')
data_f_2 =pd.DataFrame(data_2)
bgs = pd.DataFrame(data_f_2['売上予算'].values)
bg = (bgs[0+no[ln]:1+no[ln]])
nocs = pd.DataFrame(data_f_2['売上客数'].values)
noc = (nocs[0+no[ln]:1+no[ln]])
print(noc)
print(bg)

#ーーーーーーーー商品別実績ーーーーーーーーーーー

item_cd_list = ["01","02","03","04","05","06","07","08","09","10","11","12","13","15"]
item_name_list = ["OP","CD","JK","KT","CS","CT","BL","SK","PT","TR","INN","SETUP","AC","SH"]
values_list = [] #売上集計リスト
#-------- 構成比を格納 ------------

op_list = []
cd_list = []
jk_list = []
kt_list = []
cs_list = []
ct_list = []
bl_list = []
sk_list = []
pt_list = []
tr_list = []
inn_list = []
setup_list = []
ac_list = []
sh_list = []

#-------- 金額を格納 ------------

op_list2 = []
cd_list2 = []
jk_list2 = []
kt_list2 = []
cs_list2 = []
ct_list2 = []
bl_list2 = []
sk_list2 = []
pt_list2 = []
tr_list2 = []
inn_list2 = []
setup_list2 = []
ac_list2 = []
sh_list2 = []

#-------- 数量を格納 ------------

op_list3 = []
cd_list3 = []
jk_list3 = []
kt_list3 = []
cs_list3 = []
ct_list3 = []
bl_list3 = []
sk_list3 = []
pt_list3 = []
tr_list3 = []
inn_list3 = []
setup_list3 = []
ac_list3 = []
sh_list3 = []

for y in range(1,22):
  print(y)
  
  data_f = pd.read_csv(shops_l[y-1],encoding='SHIFT-JIS')
  data_f_1 = pd.DataFrame(data_f)
  data_cd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).values,columns=["商品CD"])
  data_itcd = pd.DataFrame(data_f_1['商品コード'].astype('str').str.zfill(10).str[2:4].values,columns=["アイテムCD"])
  data_nm = pd.DataFrame(data_f_1['商品名'].values,columns=["商品名"])
  data_qyt = pd.DataFrame(data_f_1['合計数量'].values,columns=["数量"])
  data_amt = pd.DataFrame(data_f_1['合計金額'].values,columns=["金額"])
  #sp_df = pd.DataFrame(data_f_1[['商品名']['数量４']].values)
  #sp_s =('ｼｮｯﾋﾟﾝｸﾞﾊﾞｯｸﾞS')
  ttl = data_qyt.sum().values
  ttl_amt = data_amt.sum()
  ttl_amt_1 = (ttl_amt).values

  df_1 = pd.concat([data_cd,data_itcd,data_nm,data_qyt,data_amt], axis=1)#データ結合


  values_1 = sum(df_1["金額"].values)#売上集計実績
  
  values_list.append(values_1)

  print(values_1)
  #sp_s = df_1.filter(like='9998998001',axis=0).values
  df_2 = pd.DataFrame(df_1)

  x = 0
  print(shops_d[y])
  values_2 = "売上実績 " + "¥" + str(values_1)
  print("売上実績 " + "¥" + str(values_1))
  
  v_file = []
  for i in item_cd_list:
    
    item_name = item_name_list[x]
    
    op = pd.DataFrame(df_2[df_2["アイテムCD"] == i ])
    op_2 = sum(op["金額"].values)
    op_3 = sum(op["数量"].values)
    op_p = str("{: .1f}".format(int(op_2) / int(values_1) * 100)) + "%"
    op_5 = "{: .1f}".format(int(op_2) / int(values_1) * 100)
    
    sw_item = [op_2,op_3,op_p]
    
    
    
    
    
    if x == 0:
      op_list.append(op_5)#構成比を格納
      op_list2.append(op_2)#金額を格納
      op_list3.append(op_3)#数量を格納
    elif x == 1:
      cd_list.append(op_5)#構成比を格納
      cd_list2.append(op_2)#金額を格納
      cd_list3.append(op_3)#数量を格納      
      
    elif x == 2:
      jk_list.append(op_5)#構成比を格納
      jk_list2.append(op_2)#金額を格納
      jk_list3.append(op_3)#数量を格納   
      
    elif x == 3:       
      kt_list.append(op_5)#構成比を格納
      kt_list2.append(op_2)#金額を格納
      kt_list3.append(op_3)#数量を格納      
      
    elif x == 4:
      cs_list.append(op_5)#構成比を格納
      cs_list2.append(op_2)#金額を格納
      cs_list3.append(op_3)#数量を格納    
        
    elif x == 5:
      ct_list.append(op_5)#構成比を格納
      ct_list2.append(op_2)#金額を格納
      ct_list3.append(op_3)#数量を格納  
          
    elif x == 6:
      bl_list.append(op_5)#構成比を格納
      bl_list2.append(op_2)#金額を格納
      bl_list3.append(op_3)#数量を格納      
      
    elif x == 7:
      sk_list.append(op_5)#構成比を格納
      sk_list2.append(op_2)#金額を格納
      sk_list3.append(op_3)#数量を格納      
      
    elif x == 8:
      pt_list.append(op_5)#構成比を格納
      pt_list2.append(op_2)#金額を格納
      pt_list3.append(op_3)#数量を格納 
           
    elif x == 9:
      tr_list.append(op_5)#構成比を格納
      tr_list2.append(op_2)#金額を格納
      tr_list3.append(op_3)#数量を格納      
      
    elif x == 10:
      inn_list.append(op_5)#構成比を格納
      inn_list2.append(op_2)#金額を格納
      inn_list3.append(op_3)#数量を格納   
         
    elif x == 11:   
      setup_list.append(op_5)#構成比を格納
      setup_list2.append(op_2)#金額を格納
      setup_list3.append(op_3)#数量を格納      

    print(item_name + "  "+ str(op_3) + "点 " +" " + "¥" + str(op_2) + " "+ op_p)
    op_4 = item_name + "  "+ str(op_3) + "点 " +" " + "¥" + str(op_2) + " "+ op_p
    
    v_file.append(op_4)
    
    x += 1
  #------------------------------ ここに配置 -------------------------------------  
  #TOKEN = 'NxPDQg0tpI4oY6BZHo7vkZ3gxPtTIijpWFyN85xL2q1'#テストの部屋トークン
  TOKEN = 'TNKXBcpEMmK4JAmRPaOyVABkA5GWIJkIHQOnsyfu4MD'#FUNの部屋トークン
  api_url = 'https://notify-api.line.me/api/notify'
  headers = {'Authorization' : 'Bearer ' + TOKEN}
  #message = ('\n'+'柏'+'\n'+'【売上予算/実績】'+'\n' + str(mg1) +'\n' +'【P率】' +str(p1) +'\n'+ '【客数】'+ str(noc_2) +str(p2)+str(p3))
  message = ('\n'+'店別アイテム実績'+'\n'+'※金額構成比'+'\n'+str(shops_d[y])+'\n'+str(values_2)+'\n'+'\n'+str(v_file[0])+'\n'+str(v_file[1])+'\n'+str(v_file[2])+'\n'+str(v_file[3])+'\n'+str(v_file[4])+'\n'+str(v_file[5])+'\n'+str(v_file[6])+'\n'+str(v_file[7])+'\n'+str(v_file[8])+'\n'+str(v_file[9])+'\n'+str(v_file[10])+'\n'+str(v_file[11])+'\n'+str(v_file[12])+'\n'+str(v_file[13])+'\n'+'\n'+'質問や不明点あれば古内までご連絡下さい！'+'\n'+'\n'+'よろしくお願いいたします。')
  #(+'\n'+'岐阜'+str(p5)+'\n'+'長町'+str(p6)+'\n'+'船橋'+str(p7)+'\n'+'富士見'+str(p8)+'\n'+'レイク'+str(p9)+'\n'+'海老名')
  #(+str(p10)+'\n'+'むさし'+str(p11)+'\n'+'平塚'+str(p12)+'\n'+'名取'+str(p13)+'\n'+'大高'+str(p14)+'\n'+'東郷町'+str(p15)+'\n'+'太田'+str(p16)+'\n'+'水戸'+str(p17)+'\n'+'EXPO'+str(p18)+'\n'+'川崎'+str(p19)+'\n'+'新三郷'+str(p20)+'\n'+'詳細はOneDriveの【シフト管理】売上実績ファイルを参照下さい！')
  payload = {'message': message}

  requests.post(api_url, headers=headers, params=payload)   
  print("SUCCESSFULL!!")  
#-------------------------------------------------------------------------------
    
var_list = ["金額","数量","構成比"]    
    
sample1 = pd.DataFrame(values_list,columns=[str(var_list[0])])  
sample2 = pd.DataFrame(cd_list3,columns=[str(var_list[1])]) 
sample3 = pd.DataFrame(jk_list,columns=[str(var_list[2])])
sample4 = pd.DataFrame(kt_list3,columns=[str(var_list[1])])
sample5 = pd.DataFrame(cs_list3,columns=[str(var_list[1])])
sample6 = pd.DataFrame(ct_list3,columns=[str(var_list[1])])
sample7 = pd.DataFrame(bl_list3,columns=[str(var_list[1])])
sample8 = pd.DataFrame(sk_list3,columns=[str(var_list[1])])
sample9 = pd.DataFrame(pt_list3,columns=[str(var_list[1])])

mydata = pd.concat([sample1,sample2],axis=1)

mydata2 = pd.concat([sample1,sample3,sample4,sample5],axis=1)
print(mydata)

sns.lmplot(x="金額",y="数量",data=mydata)
#sns.scatterplot(x="金額",y="数量",data=mydata)

plt.show()
#x1 = (x="x",y="y")  


#sample3 = 
glm = smf.ols("数量",mydata2).fit()
glm.summary()
