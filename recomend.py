#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
逐次実行する場合はこちらのスクリプト
"""
import argparse
import ast
import datetime
import json
import os
import pathlib
import re
import time
import warnings
from os import path

import gspread
import pandas as pd
import syss
from bs4 import BeautifulSoup
from export2bq import export2bq
from geocoding_copy import exe_geocoding
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.utils import ChromeType

# spreadsheet設定
# 認証のjsoファイルのパス
secret_credentials_json_oath = 'geom-prj-property-recommend-8d74bd11fa15.json' 

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

credentials = Credentials.from_service_account_file(
    secret_credentials_json_oath,
    scopes=scopes
)

gc = gspread.authorize(credentials)
# https://docs.google.com/spreadsheets/d/{ココ}/edit#gid=0
workbook = gc.open_by_key('1WU0WTNOGfXwoWhDXhCPJYZGrF0n4IQLQh-d3Ag4U8n4')
worksheet = workbook.get_worksheet(0)

# crawling設定
manager = ChromeDriverManager(chrome_type=ChromeType.CHROMIUM)

print(manager.driver.get_os_type())
print(manager.driver.get_version())
print(manager.driver.get_url())

warnings.filterwarnings('ignore')

current_dir = pathlib.Path(__file__).resolve().parent
sys.path.append(str(current_dir) + '/../')

# parser = argparse.ArgumentParser()
# # parser.add_argument('brand_name')
# # parser.add_argument('--prefecture', nargs='*')
# args = parser.parse_args()


df_master_pref=pd.read_csv('df_master_prefecture.csv')
df_master_pref=df_master_pref.set_index("name_pref")

# 坪単価の詳細条件
dic_tsubo_all=pd.read_csv('df_dic_tsubo.csv')

# crawling
class ATBB:  # 'Brand'には対象のブランド名をキャメルケースで入れてください
    BASE_URL = 'https://atbb.athome.jp/'  # 作業対象ページのurl

    def __init__(self, base_url):  # 初期化： インスタンス作成時に自動的に呼ばれる
        self.base_url = base_url


    def getHtmlData(self,url_list,brand_name,station_minute,area_size_min_tsubo,area_size_max_tsubo,
                    building_age,floor_value,target_prefs,building_structure=None,kodawari=None):

        # print("building_structure:",building_structure)
        # print("kodawari:",kodawari)
        '''
        対象ページのhtmlデータを返すメソッド
        Parameters
        ----------
        url_list : list
            urlが格納されたリスト
        Returns
        -------
        html_contents : list
            htmlデータが格納されたリスト
        '''
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument(
            f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36')

        driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        # driver = webdriver.Chrome("/Users/murakamieikai/chromedriver")
        html_contents = []
        try:
            for url in url_list:
                driver.get(url)
                time.sleep(15)  # サイトによっては描画に時間がかかるため、適宜変更してください

                login = driver.find_element_by_id('loginFormText')
                login.send_keys('002717900001')

                password = driver.find_element_by_id('passFormText')
                password.send_keys('2jag5yCfFRp4')

                time.sleep(15)

                driver.find_element_by_id('loginSubmit').click()

                time.sleep(15)

                for j in target_prefs:
                    i=df_master_pref.at[j,"num_pref"] # 都道府県番号

                    driver.get('https://members.athome.jp/atbb/nyushuSearch?from=global_menu_bukkenKensaku')
                    time.sleep(15)

                    try:
                        driver.find_element_by_xpath("//input[@value='強制終了させてATBBを利用する']").click()
                        print('強制終了させてATBBを利用する')
                        time.sleep(15)
                        driver.switch_to.alert.accept()
                        print('prefecture ' + str(j))
                    except:
                        print('prefecture ' + str(j))

                    time.sleep(30)
                    label_pointer = driver.find_element_by_xpath("//input[@value='07']")
                    label_pointer.click()
                    time.sleep(10)

                    if i == 1:
                        area_pointer = driver.find_element_by_xpath("//input[@name='area' and @value='0101']")
                    else:
                        area_pointer = driver.find_element_by_xpath("//input[@name='area' and @value='" '{:02}'.format(i) + "']")

                    area_pointer.click()

                    validate_pointer = driver.find_element_by_xpath("//input[@value='所在地検索']")
                    validate_pointer.click()
                    time.sleep(15)

                    area_list = driver.find_element_by_xpath("//select[@name='sentaku1Shikugun']")                    

                    # area_list.send_keys(Keys.CONTROL, 'a')  # Windows
                    area_list.send_keys(Keys.COMMAND, 'a')  # Mac全選択
                    time.sleep(1)

                    validate_pointer = driver.find_element_by_xpath("//input[@value='条件入力画面へ']")
                    validate_pointer.click()
                    time.sleep(15)

                    driver.find_element_by_xpath('//*[@id="bukkenShumokuBunrui_0"]').click()  # 物件種目: 店舗
                    
                    
                    # 駅歩分
                    try:
                        driver.find_element_by_xpath(f'//*[@name="ekiHoFun" and @value={station_minute}]').click()  # 駅徒歩5分以内
                    except:
                        pass

                    # 使用面積
                    try:
                        driver.find_element_by_xpath('//*[@name="tatemonoShiyoBubunMensekiTani" and @value="1"]').click()  # 坪
                        driver.find_element_by_xpath('//*[@name="tatemonoShiyoBubunMensekiChokusetsuNyuryokuFrom"]').send_keys(f'{area_size_min_tsubo}')  
                        driver.find_element_by_xpath('//*[@name="tatemonoShiyoBubunMensekiChokusetsuNyuryokuTo"]').send_keys(f'{area_size_max_tsubo}') 
                    except:
                        pass

                    # 建物構造
                    if "01" in building_structure:
                        driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[1]/td[1]/input').click() 
                    if "02" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[1]/td[2]/input").click()
                    if "03" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[1]/td[3]/input").click()
                    if "04" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[1]/td[4]/input").click()
                    if "05" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[1]/td[5]/input").click()
                    if "06" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[2]/td[1]/input").click()
                    if "07" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[2]/td[2]/input").click()
                    if "08" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[2]/td[3]/input").click()
                    if "09" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[2]/td[4]/input").click()
                    if "10" in building_structure:
                        driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form/table/tbody/tr[5]/td[2]/table/tbody/tr[2]/td[5]/input").click()

                    # こだわり条件
                    if "12" in kodawari:
                        driver.find_element_by_xpath('//*[@name="kodawariJokenCode" and @value="12"]').click()  # 飲食店可
                    if "13" in kodawari:
                        driver.find_element_by_xpath('//*[@name="kodawariJokenCode" and @value="13"]').click()  # 飲食店不可
                    if "51" in kodawari:
                        driver.find_element_by_xpath('//*[@name="kodawariJokenCode" and @value="51"]').click()  # 男女別トイレ
                    if "52" in kodawari:
                        driver.find_element_by_xpath('//*[@name="kodawariJokenCode" and @value="52"]').click()  # 居抜き
                    

                    validate_pointer = driver.find_element_by_xpath("//input[@value='検索']")
                    validate_pointer.click()

                    time.sleep(15)

                    try:
                        html_contents.append((driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form[21]/table[1]/tbody/tr[5]/td[2]').text.split('：')[0].replace('全域','') ,driver.find_element_by_class_name('bukkenKensakuKekkaWrapper').get_attribute('innerHTML')))
                    except:
                        continue

                    # ページの分だけ「次へ」をクリック
                    while(1):
                        try:
                            driver.find_element_by_xpath('htm/l/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form[21]/table[4]/tbody/tr/td[2]/ul/li[3]/a')
                            driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form[21]/table[4]/tbody/tr/td[2]/ul/li[3]').click()
                            time.sleep(10)

                            html_contents.append((driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td[3]/form[21]/table[1]/tbody/tr[5]/td[2]').text.split('：')[0].replace('全域','') ,driver.find_element_by_class_name('bukkenKensakuKekkaWrapper').get_attribute('innerHTML')))
                        except:
                            break

                driver.find_element_by_class_name("shuryoBtn").click() # ATBBを終了するボタン

        except Exception as e:
            print('driver関連のerr msg', e)
        finally:
            driver.quit()

        return html_contents


    def getUrlList(self):
        '''
        遷移先のリンクアドレスを返すメソッド
        Returns
        -------
        url_list : list
            リンクアドレスが格納されたリスト
        '''
        return [self.base_url]


    def getStoreInfo(self):
        '''
        対象ページの店舗名と住所を返すメソッド
        Returns
        -------
        store_info: DataFrame
            店舗名と住所が格納されたデータフレーム
        '''
        html_contents = []
        store_info = []
        url_list = self.getUrlList()
        html_contents = self.getHtmlData(url_list=url_list,brand_name=brand_name,station_minute=station_minute,area_size_min_tsubo=area_size_min_tsubo,
                                        area_size_max_tsubo=area_size_max_tsubo,building_age=building_age,
                                        floor_value=floor_value,target_prefs=target_prefs,building_structure=building_structure,kodawari=kodawari)#,target_prefs=target_prefs

        for prefecture, html in html_contents:
            soup = BeautifulSoup(html, "html.parser")

            items = soup.find_all('table', recursive = False)

            df=pd.DataFrame()
            for item in items:
                tds = item.find_all('td')
                for p, td in enumerate(tds):
                    tds[p] = ' '.join(td.text.split()).strip()

                # 坪単価取得
                store_info_tsubo=[]
                store_info_tsubo.append((tds[15], tds[16], tds[17], tds[18],
                                        tds[19], tds[20], tds[21], tds[22],
                                        tds[23], tds[24]))

                ls_tsubo=[]
                ls_tsubo = [s for s in store_info_tsubo[0] if '円' in s]
                if len(ls_tsubo) == 2:
                    ls_tsubo=ls_tsubo[1:]

                store_info.append(( prefecture, tds[1], tds[2],tds[3], tds[4],float(re.sub(r"[^\d.]", "", ls_tsubo[0])) ,   
                                    # tds[5], tds[6],tds[7],tds[8],tds[9],tds[10], tds[11], tds[12], tds[13], tds[14],
                                    # tds[15], tds[16], tds[17], tds[18],
                                    # tds[19], tds[20], tds[21], tds[22],
                                    # tds[23], tds[24], tds[25], tds[26]
                                    ))    

        df = pd.DataFrame(store_info, columns=['都道府県名', '物件種目', '所在地','建物名/部屋番号', '公開日','坪単価',
            # '登録賃料','交通（沿線駅/バス停）', '礼金', '管理費等', '築年月', '広告転載',
            # '敷金', '保証金', '階建/階', '取引態様',
            # '敷引', '坪単価', '建物構造', '物件番号',
            # '建物面積', '登録会員', 'TEL', '免許番号'
            ])
        # df_condition_tsubo_tanka=df.query('("都道府県名" == "東京" and "坪単価" <= 20000) or ("都道府県名"!="東京" and "坪単価" <= 15000)')

        return df #df_condition_tsubo_tanka

# 東京都, 岐阜県, 愛知県, 滋賀県, 京都府
if __name__ == '__main__':
    end_date=datetime.datetime.now().strftime('%Y%m%d')
    # 現在の日付
    target_date = datetime.date.today()
    
    # 対象の日付が含まれる週の月曜日の一週間前の月曜日の日付を取得。
    monday = target_date - datetime.timedelta(days=target_date.weekday())-datetime.timedelta(7)

    ls_target=worksheet.get('A1:AI')
    num_target=len(ls_target)
    
    for i in range(1,num_target):#num_target ブランドごと
        
        target=ls_target[i]
        brand_name=target[4]
        brand_name_eng=target[33]

        # 物件をお探しの都道府県
        target_prefs=target[10].split(",")

        for j in range(1):#len(target_prefs)
            target_prefs[j]=target_prefs[j].replace(" ","")
        print(brand_name,target_prefs)

        # 駅歩分
        station_minute=target[22]
        dic_station_minute={"指定なし":"-1","1分以内":"1","3分以内":"3","5分以内":"5","10分以内":"10","15分以内":"15","20分以内":"20"}
        for k,v in dic_station_minute.items():
            if station_minute==k:
                station_minute=v

        # 建物の使用部分面積の下限[坪]
        try:
            area_size_min_tsubo=float(target[26])
        except:
            area_size_min_tsubo=0

        # 建物の使用部分面積の上限[坪]
        try:
            area_size_max_tsubo=float(target[27])
        except:
            area_size_max_tsubo=15000

        # 建物構造
        try:
            building_structure=target[28].split(",")
            for j in range(len(building_structure)):
                building_structure[j]=building_structure[j].replace(" ","")
            
            dic_structure={"木造":"01","ブロック造":"02",'鉄構造':"03","SRC":"05","RC":"04","HPC":"07","PC":"06","軽量鉄骨造":"08","ALC":"09","その他":"10"}
            for k,v in dic_structure.items():
                building_structure=[s.replace(k,v) for s in building_structure]
        except:
            None        
        
        # 築年数
        building_age=target[29]
        dic_buiding_age={"指定なし":"-1","新築":"00","1年以内":"01","2年以内":"02","3年以内":"03","4年以内":"04","5年以内":"05","10年以内":"10","15年以内":"15","20年以内":"20"}
        for k,v in dic_buiding_age.items():
            if building_age==k:
                building_age=v

        # 階指定
        floor_value=target[30]
        dic_floor={"指定なし":"0","地下":"9","1階":"1","2階以上":"2","最上階":"3"}
        for k,v in dic_floor.items():
            if floor_value==k:
                floor_value=v

        # こだわり条件
        try:
            kodawari=target[31].split(",")
            dic_kodawari={"造作譲渡":"11","駐車場あり":"06","シャワールーム":"95","男女別トイレ":"51","エアコン":"30","有線放送":"08","インターネット対応":"63","インターネット使用料無料":"93","オートロック":"09","防犯カメラ":"76","24時間セキュリティー":"10","エレベーター":"07","温泉":"99","フリーレント":"83","飲食店可":"12","飲食店不可":"13","スケルトン":"54","居抜き":"52","24時間利用可":"84","OAフロア":"53","クレジットカード決済可":"89","ＩＴ重説対応物件":"92"}
            for k,v in dic_kodawari.items():
                kodawari=[s.replace(k,v) for s in kodawari]
        except:
            None

        new_dir_path = f'results_crawl/{brand_name}'

        try:
            os.mkdir(new_dir_path)
        except:
            pass

        try:
            print(os.path.basename(__file__) + 'について処理しています')

            # getStoreInfo()はインスタンスメソッドなのでこのように呼び出します
            df_result_crawl_tmp = ATBB(ATBB.BASE_URL).getStoreInfo()

            # 坪単価でフィルター
            df=pd.DataFrame()

            try:
                tsubo_min=float(target[23]) # 坪単価下限
            except:
                tsubo_min=0

            try:
                tsubo_max=float(target[24]) # 坪単価上限
            except:
                tsubo_max=50000000

            # 都道府県ごとに指定する場合
            try:
                dic_tsubo=dic_tsubo_all.query('brand_name==@brand_name').iloc[0][0]
                dic_tsubo=ast.literal_eval(dic_tsubo)

                for k, v in dic_tsubo.items():
                    df_t=df_result_crawl_tmp.query('都道府県名==@k and 坪単価<=@v')
                    df=pd.concat([df,df_t])
            # 都道府県に関わらず一律の場合
            except:
                df=df_result_crawl_tmp.query('(坪単価 >= @tsubo_min) and (坪単価 <=@tsubo_max)')

            print(f"対象ブランド: {brand_name}")
            today_date=datetime.datetime.now().strftime('%Y%m%d')

            exe_geocoding(brand_name,end_date,target_date,monday,df)
            # df.to_csv(f'results_crawl/{brand_name}/{brand_name}_ATBB_{today_date}.csv', index=False, encoding="utf-8")
            export2bq(brand_name,brand_name_eng,target_date)
            
        except BaseException as e:
            print(e)