# © 2024 Keita Iwasa

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import time
import os
import sys
import logging
from openpyxl import load_workbook
import pandas as pd
import numpy as np
import winreg
import configparser
import httplib2
import json
import requests
from datetime import datetime, timedelta

# 設定ファイルの読み込み
config = configparser.ConfigParser()
with open('setup/config.ini', 'r', encoding='utf-8') as file:
    config.read_file(file)
st = config['Settings'] 

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class AutomationHandler:
    def __init__(self):
        # Google Apps ScriptのエンドポイントURL
        self.script_url = 'https://script.google.com/macros/s/AKfycbyixzT47V81tUQG2DgmO-YbPEdsm08m0CZxESsYeQziZ7SfS-n6xe7NN3gs4nb7CST6/exec'
        self.driver = None

    def call_google_script(self, function_name, params):
        data = {
            'function': function_name,
            'parameters': params
        }
        response = requests.post(self.script_url, data=json.dumps(data), headers={'Content-Type': 'application/json'})
        if response.status_code == 200:
            logging.info(f"Response JSON from GAS: {response.json()}")
            return response.json()
        else:
            raise Exception(f"Google Apps Script呼び出しエラー: {response.text}")
    
    def get_original_sheet(self):
        params = {'shopName': st['SHOP_NAME']}
        response = self.call_google_script('getOriginalSheet', params)
        if response['found']:
            return response['sheet_url']
        else:
            return False
        
    def check_existing_sheet(self, sheet_name):
        params = {
            'shopName': st['SHOP_NAME'],
            'sheet_name': sheet_name
        }
        response = self.call_google_script('checkExistingSheet', params)
        if response['found']:
            return response['sheet_id'], response['spreadsheet_url']
        else:
            return False
                
    # EOSログインメソッド
    def login_eos(self, user_id, password):
        if self.driver is None:
            options = Options()
            options.add_experimental_option('detach', True)
            self.driver = webdriver.Chrome()
            self.driver.minimize_window() #誤操作を防ぐためにウィンドウを最小化
            self.driver.get('https://eos-st.komeda.co.jp/st/') #ログインページにアクセス           
        # EOSにログイン
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'txtUserId'))).send_keys(user_id) #ユーザーID入力
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'txtPassword'))).send_keys(password) #パスワード入力
            self.driver.find_element(By.ID, 'btnLogin').click() #「ログイン」ボタンクリック    
        except TimeoutException: # 別ページでEOSが開かれていた場合、TimeoutExceptionとなる
            self.driver.find_element(By.ID, 'btnNext').click() #「開く」ボタンクリック
        WebDriverWait(self.driver, 15).until(EC.url_to_be('https://eos-st.komeda.co.jp/st/osirase'))

        # お知らせが表示される場合は✕ボタン
        while len(self.driver.find_elements(By.XPATH, value="//button[@title='Close']"))>0 :
            self.driver.find_element(By.XPATH, value="//button[@title='Close']").click()
            time.sleep(0.05)
            
    def download_folder_path(self):
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key)
        download_folder = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        return download_folder
            
    def download_csv(self, today_str_csv, today_int):
        self.login_eos(st['EOS_ID'], st['EOS_PW']) #EOSログインメソッド↑
        self.csv_path = f'{self.download_folder_path()}/{today_str_csv}_発注.CSV'
        self.csv_path_nonfood = f'{self.download_folder_path()}/{today_str_csv}_発注 (1).CSV'
        max_retry_download = 15
        try:
            if not os.path.exists(self.csv_path):           
                # 左メニューの発注照会をクリック
                menupng2 = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'menupng2'))) #発注
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menupng2)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", menupng2)  # JavaScriptでクリックを強制実行
                
                inquiry_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@accesskey='4']")))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inquiry_element)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", inquiry_element)  # JavaScriptでクリックを強制実行

                # 本日の発注を照会
                inquiry_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'inquiryButton')))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inquiry_button)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", inquiry_button)  # JavaScriptでクリックを強制実行

                # 画面の一番下までスクロール
                self.driver.execute_script("document.body.style.zoom='65%'")
                self.driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
                    
                # CSVをダウンロード
                btn_csv_out = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'btnCsvoutConfirm')))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_csv_out)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", btn_csv_out)  # JavaScriptでクリックを強制実行 
                
                # ウィジェットのはいをクリック
                btn_yes = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[text()='はい']")))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_yes)
                time.sleep(0.3)  # 少し待機
                self.driver.execute_script("arguments[0].click();", btn_yes)

            today_str_csv = datetime.today().strftime('%Y%m%d') #本日の日付（ダウンロードした発注明細csvは発注日付に関係なく本日のreal日付が付いている）
            retry_download = 0
            # 前日の発注明細をロード
            while retry_download <= max_retry_download:
                if os.path.exists(self.csv_path):
                    break
                elif retry_download == max_retry_download:
                    return False    
                else:
                    retry_download += 1
                    time.sleep(0.1)
        except:
            return False

        if today_int.weekday() in {3, 5}:#木夜、土夜の場合は非食品の発注明細も取得・合成
            if os.path.exists(self.csv_path_nonfood):
                pass
            else:
                # 左メニューの発注照会をクリック
                WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'menupng2'))).click() #発注
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@accesskey='4']"))).click() #発注照会クリック
                # 前日の日付を計算
                yesterday_int = today_int - timedelta(days=1)
                # 前日の日付を文字列に変換
                yesterday_str = yesterday_int.strftime('%d')
                # 前日の日付の番号を取得（"06" -> "6"のように先頭のゼロを取り除く）
                yesterday_number = str(int(yesterday_str))
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'selectFromDay')))
                select_from_date = self.driver.find_element(By.ID, 'selectFromDay')
                select_from = Select(select_from_date)
                select_from.select_by_value(yesterday_number)

                select_to_date = self.driver.find_element(By.ID, 'selectToDay')
                select_to = Select(select_to_date)
                select_to.select_by_value(yesterday_number)

                select_kubun = self.driver.find_element(By.ID, 'selectOrderKubun')
                select_kubun_nonfood = Select(select_kubun)
                select_kubun_nonfood.select_by_value("2")

                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'inquiryButton'))).click() #照会ボタン
                time.sleep(0.5)
                    
                # CSVをダウンロード
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'btnCsvoutConfirm'))).click() #CSV出力をクリック
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[text()='はい']"))).click() #ウィジェットのはいをクリック
                retry_download = 0
                # 前々日の非食品の発注明細をダウンロード
                while retry_download <= max_retry_download:
                    if os.path.exists(self.csv_path_nonfood) or retry_download == max_retry_download:
                        break                        
                    else:
                        retry_download += 1
                        time.sleep(0.1)
            
            # 昨日の発注明細csvと一昨日の非食品の発注明細csvを合成  
            if os.path.exists(self.csv_path_nonfood):
                df1 = pd.read_csv(self.csv_path)
                df2 = pd.read_csv(self.csv_path_nonfood)
                merged_df = pd.concat([df1, df2], ignore_index=True)  
                merged_df.to_csv(self.csv_path, index=False)
        
        self.driver.close()
        self.driver.quit() 
        self.driver = None  

        if os.path.exists(self.csv_path):
            return True
        else:
            return False

    def execute_with_retry(self, function_name, params, retries=3, timeout=120):
        for attempt in range(retries):
            try:
                response = self.call_google_script(function_name, params)
                return response
            except Exception as e:
                print(f"Attempt {attempt + 1} failed with error: {e}. Retrying...")
                time.sleep(5)
        raise Exception("All retry attempts failed")
           
    def generate_form(self, delivery_date_int, today_str, night_order):
        if night_order:
            with open(self.csv_path, 'r', encoding='utf-8') as file:
                csv_data = file.read()
        else:
            csv_data = False

        params = {
            'delivery_date_int': delivery_date_int.isoformat(),
            'csv_data': csv_data,
            'today_str': today_str,
            'night_order': night_order,
            'shop_name': st['SHOP_NAME']
        }

        response = self.execute_with_retry('generateForm', params, retries=5)
        if 'error' in response:
            return False
        else:
            return response['spreadsheetId'], response['spreadsheetUrl']

    def get_spreadsheet(self, sheet_id):
        params = {'sheet_id': sheet_id}
        response = self.execute_with_retry('getSpreadsheet', params, retries=5)
        if 'error' in response:
            raise Exception(f"Error in getting spreadsheet: {response['error']}")
        else:
            values_food = response['values_food']
            values_nonfood = response['values_nonfood']
            # 指定した複数の列をDataFrameに変換
            if not values_food:
                print('No data found in the sheet.')
                return False, False
            else:
                #食品
                max_columns = len(values_food[0])
                data = [row + [None] * (max_columns - len(row)) for row in values_food[1:]]
                df_food = pd.DataFrame(data, columns=values_food[0])
                self.input_df = df_food[['商品名', 'セット', '商品コード', '現在庫', '発注数']]
                self.input_df.replace('', None, inplace=True)
                Name_with_NaN = self.input_df[self.input_df['現在庫'].isna()]['商品名'].tolist()
                if len(Name_with_NaN) == 0:
                    self.input_df['商品コード'] = self.input_df['商品コード'].astype(int)
                    self.input_df['現在庫'] = self.input_df['現在庫'].astype(int)
                    self.input_df['発注数'] = self.input_df['発注数'].astype(int)
                    self.input_df['セット'] = self.input_df['セット'].astype(int)

                #非食品
                max_columns_nonfood = len(values_nonfood[1])
                data_nonfood = [row + [None] * (max_columns_nonfood - len(row)) for row in values_nonfood[1:]]
                df_nonfood = pd.DataFrame(data_nonfood, columns=values_nonfood[0])
                self.input_df_nonfood = df_nonfood[['商品名', '商品コード', '発注数']]
                self.input_df_nonfood['発注数'] = self.input_df_nonfood['発注数'].astype(str).str.strip()
                self.input_df_nonfood['発注数'] = pd.to_numeric(self.input_df_nonfood['発注数'], errors='coerce')
                self.input_df_nonfood.dropna(subset=['発注数'], inplace=True)
                self.input_df_nonfood = self.input_df_nonfood[self.input_df_nonfood['発注数'] != 0]
                if not self.input_df_nonfood.empty:
                    self.input_df_nonfood.reset_index(drop=True, inplace=True)
                    self.input_df_nonfood.replace('', None, inplace=True)
                    self.input_df_nonfood['商品コード'] = self.input_df_nonfood['商品コード'].astype(int)
                    self.input_df_nonfood['発注数'] = self.input_df_nonfood['発注数'].astype(int)
                    
                return Name_with_NaN, self.input_df_nonfood
            
    def input_order_in_site(self):
        self.login_eos(st['EOS_ID'], st['EOS_PW'])

        #ウィンドウを最大化
        self.driver.maximize_window()

        # 左メニューの発注入力をクリック
        self.driver.find_element(By.CLASS_NAME, 'menupng2').click()
        input_order = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@accesskey='3']")))
        input_order.click()

        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'scode'))) 
        
        input_df_tuple = (self.input_df, self.input_df_nonfood)
        error_ls = [] #入力エラーの空リストを作成
        for df in input_df_tuple:
            if df is self.input_df:
                print(f'df is self.input')
            elif df.shape[0] == 0:
                print(f'df is False')
                continue
            else:
                print(f'df is else')
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'pushDay2')))
                self.driver.find_elements(By.CLASS_NAME, 'pushDay2')[0].click()

            # 発注サイトの商品番号をリストにする
            WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'scode')))
            table_number_ls = self.driver.find_elements(By.CLASS_NAME, 'scode') #ドライバーのリストができる
            for i in range(len(table_number_ls)):
                table_number_ls[i] = int(table_number_ls[i].text) #ドライバーのリストを商品番号のリストに変換

            # 発注サイトの商品名をリストにする
            table_name_ls = self.driver.find_elements(By.XPATH, value="//span[starts-with(@id, 'syhnnm')]" ) #ドライバーのリストができる
            for i in range(len(table_name_ls)):
                table_name_ls[i] = str(table_name_ls[i].text) #ドライバーのリストを商品名のリストに変換

            # 発注サイトのprdxをリストにする
            table_prdx_ls = self.driver.find_elements(By.XPATH, value="//input[starts-with(@id, 'prdx')]") #ドライバーのリストができる
            for i in range(len(table_prdx_ls)):
                table_prdx_ls[i] = str(table_prdx_ls[i].get_attribute("id"))

            # 商品番号リストとprdxリストから辞書を作成
            table_prdx_number_dict = dict(zip(table_number_ls, table_prdx_ls))

            dict_data = pd.Series(df['商品名'].values, index=df['商品コード'].values) #商品名：商品コードの辞書作成

            # 入力
            for row in df.itertuples():
                if row.発注数 <= 0:
                    continue
                else:
                    try:
                        table_id = table_prdx_number_dict[row.商品コード] #発注する商品番号から商品のprdx値を求める
                    except:
                        error_ls.append(f"{row.商品コード}：{dict_data[row.商品コード]}（エラー理由：EOSに存在しない商品, 商品番号の誤り, お気に入り未登録）")
                        continue
                    set_value = int(self.driver.find_element(By.ID, table_id).get_attribute('data-sthtsu')) #セット数（EOS由来）
                    order_value = set_value * row.発注数 #発注入力数
                    print(f'{row.商品名}:{order_value}')

                    if  df is self.input_df:
                        if not set_value == row.セット: #発注書とEOSのセット数が一致しているか確認
                            print(f"商品番号：{row.商品コード}のセット数が誤っています。発注書のセット数を修正してください。")
                    else:
                        pass
                        
                    try:
                        input_field = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, table_id)))
                        input_field.clear() #input_fieldのデフォルト0をクリア   
                        input_field.send_keys(order_value) #発注数を入力
                    except: 
                        dialog_text = self.driver.find_element(By.ID, 'divDialog').text
                        if '制限数量' in dialog_text:
                            error_ls.append(f'{former_input_number}：{dict_data[former_input_number]}（エラー理由：発注数MAX超え）')
                            self.driver.find_element(By.CLASS_NAME, 'ui-icon-closethick').click() # ×ボタンでダイアログを閉じる
                            input_field = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, former_table_id)))
                            input_field.clear()
                            former_max_order_value = int(self.driver.find_element(By.ID, former_table_id).get_attribute('data-sgosuu')) - 1
                            input_field.send_keys(former_max_order_value) #発注数を入力  
                        else:
                            error_ls.append(f'{former_input_number}：{dict_data[former_input_number]}（エラー理由：不明）')
                            self.driver.find_element(By.CLASS_NAME, 'ui-icon-closethick').click() # ×ボタンでダイアログを閉じる
                        
                        input_field = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, table_id)))
                        input_field.clear() #input_fieldのデフォルト0をクリア   
                        input_field.send_keys(order_value) #発注数を入力

                former_table_id = table_id
                former_input_number = row.商品コード  

        if len(error_ls) > 0 :    
            for i in range(len(error_ls)):
                print(f'\033[93m{error_ls[i]}\033[0m')
        return True, error_ls

    def destroy_chrome(self):
        try:
            self.driver.close()
            self.driver.quit()
            return True
        except:
            return False
