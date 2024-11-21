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
import winreg
import configparser
import json
import requests
from datetime import datetime, timedelta
import random
import glob

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
        self.script_url = 'https://script.google.com/macros/s/AKfycbwrdCpKUelDpHukcwgw2e2Nt04nmonpYhUfMQKLSL2ZhXwEHqp0yXlHpoRekPYn_i5EOg/exec'
        self.driver = None

    def check_update(self, store_name, current_version): # return need_update, self.latest_version
        try:
            response = requests.get("https://support.iwasadigital.com/kaos/version.json")
            logging.info(response.json())
            print(response.status_code)
            if response.status_code == 200:
                versions = response.json()
                if store_name not in versions:
                    latest_version = versions.get('default')
                else:
                    latest_version = versions.get(store_name)
                if latest_version and latest_version != current_version:
                    return True, latest_version
                else:
                    return False, None
            else:
                logging.error(f"Error during version check: {response.text}")
                return False, None
        except requests.exceptions.RequestException as e: # 通信エラー
            logging.error(f"Error during version check: {e}")
            return False, 404
        except Exception as e:
            logging.error(f"Error during version check: {e}")
            return False, None
        
    def download_updater(self, latest_version, installer_path):
        try:
            files = glob.glob('setup/KAOS_setup.*.exe')
            for file in files:
                os.remove(file)
            url = f"https://support.iwasadigital.com/kaos/{latest_version}/KAOS_setup.{latest_version}.exe"
            response = requests.get(url)
            logging.info(f"Response status code: {response.status_code}")
            with open(installer_path, 'wb') as file:
                file.write(response.content)
            return True
        except Exception as e:
            logging.error(f"Error during dounloading updater: {e}")
            return False
    
    def call_google_script(self, function_name, params):
        max_retries = 3

        for attempt in range(max_retries):
            try:
                data = {
                    'function': function_name,
                    'parameters': params
                }
                response = requests.post(self.script_url, data=json.dumps(data), headers={'Content-Type': 'application/json'})
                if response.status_code == 200:
                    logging.info(f"Response JSON from GAS: {response.json()}")
                    return response.json()
                else:
                    logging.warning(f"Attempt {attempt + 1} failed: {response.text}")
                    raise Exception(f"Google Apps Script呼び出しエラー: {response.text}")
            except Exception as e:
                if attempt < max_retries - 1:
                    delay = (2 ** attempt) + random.uniform(0, 1)  # 指数バックオフ＋ランダムジッター
                    logging.info(f"Retrying in {delay:.2f} seconds...")
                    time.sleep(delay)
                else:
                    logging.error(f"Max attempts reached. Raising exception.")
                    raise e
        
    def get_original_sheet(self):
        params = {'shopName': st['SHOP_NAME']}
        response = self.call_google_script('getOriginalSheet', params)
        if response['found']:
            return response['sheet_url']
        else:
            return False
        
    def get_notices(self, version, today_int):
        params = {
            'now': today_int.isoformat(),
            'version': version}
        response = self.call_google_script('getNotice', params)
        if 'error' in response or len(response['notices']) == 0:
            return False
        else:
            return response['notices']

    def check_existing_sheet(self, sheet_name):
        params = {
            'shopName': st['SHOP_NAME'],
            'sheet_name': sheet_name
        }
        response = self.call_google_script('checkExistingSheet', params)
        if 'found' in response:
            if response['found']:
                return response['sheet_id'], response['spreadsheet_url']
            else:
                return False
        else:
            logging.error(f"Error in checking existing sheet: {response}")
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
            self.driver.get('https://eos-st.komeda.co.jp/st/')
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'txtUserId'))).send_keys(user_id) #ユーザーID入力
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'txtPassword'))).send_keys(password) #パスワード入力
            self.driver.find_element(By.ID, 'btnLogin').click() #「ログイン」ボタンクリック    
        except TimeoutException: # 別ページでEOSが開かれていた場合、TimeoutExceptionとなる
            self.driver.get('https://eos-st.komeda.co.jp/st/')
            self.driver.find_element(By.ID, 'btnNext').click() #「開く」ボタンクリック
            logging.info('TimeoutException')
        for _ in range(4):
            try:
                WebDriverWait(self.driver, 2).until(EC.url_to_be('https://eos-st.komeda.co.jp/st/osirase'))
                break
            except TimeoutException:
                if len(self.driver.find_elements(By.XPATH, value="//div[contains(text(), 'ユーザーまたはパスワードが一致しませんでした')]"))>0 :
                    logging.warning('入力エラーダダイアログdetected')
                    return "E0007" # ログインエラー
        WebDriverWait(self.driver, 1).until(EC.url_to_be('https://eos-st.komeda.co.jp/st/osirase'))
        logging.info('EOSログイン成功')
        # お知らせが表示される場合は✕ボタン
        while len(self.driver.find_elements(By.XPATH, value="//button[@title='Close']"))>0 :
            self.driver.find_element(By.XPATH, value="//button[@title='Close']").click()
            time.sleep(0.05)

        return "200" # OK
            
    def download_folder_path(self):
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key)
        download_folder = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        return download_folder
            
    def download_csv(self, today_str_csv, today_int):
        login_status_code = self.login_eos(st['EOS_ID'], st['EOS_PW']) #EOSログインメソッド↑
        logging.info(f'login_status_code(download_csv): {login_status_code}')
        if login_status_code != "200":
            return login_status_code # ログインエラー(E0007:IDorパスワード誤り)
        self.csv_path = f'{self.download_folder_path()}/{today_str_csv}_発注.CSV'
        try:
            if not os.path.exists(self.csv_path): 
                # 左メニューの発注照会をクリック
                menupng2 = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'menupng2'))) #発注
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menupng2)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", menupng2)  # .click()だと画面サイズなどによってうまくいかない場合があるので、JavaScriptでクリックを強制実行
                
                inquiry_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@accesskey='4']")))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inquiry_element)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", inquiry_element)  # JavaScriptでクリックを強制実行

                # 前日の日付を計算
                yesterday_int = today_int - timedelta(days=2)
                # 前日のdを文字列に変換
                yesterday_day = str(int(yesterday_int.strftime('%d')))
                today_day = str(int(today_int.strftime('%d')))
                # 前日のmを文字列に変換
                yesterday_month = str(int(yesterday_int.strftime('%m')))
                today_month = str(int(today_int.strftime('%m')))
                # 前日のyを文字列に変換
                yesterday_year = yesterday_int.strftime('%Y')


                # 明細の日付範囲のスタート
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'selectFromDay')))
                select_from_date = self.driver.find_element(By.ID, 'selectFromDay')
                select_from = Select(select_from_date)
                select_from.select_by_value(yesterday_day)

                # 明細の月範囲のスタート
                if today_day in ["1", "2"]:
                    select_from_month = self.driver.find_element(By.ID, 'selectFromMonth')
                    select_from = Select(select_from_month)
                    select_from.select_by_value(yesterday_month)
                    if today_month == "1":
                        select_from_year = self.driver.find_element(By.ID, 'selectFromYear')
                        select_from = Select(select_from_year)
                        select_from.select_by_value(yesterday_year)

                # 発注明細を照会
                inquiry_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'inquiryButton')))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inquiry_button)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", inquiry_button)  # JavaScriptでクリックを強制実行

                # 画面の一番下までスクロール
                self.driver.execute_script("document.body.style.zoom='65%'")
                self.driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
                    
                # CSVをダウンロード
                btn_csv_out = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'btnCsvoutConfirm')))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_csv_out)
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click();", btn_csv_out)  # JavaScriptでクリックを強制実行 
                
                # ウィジェットのはいをクリック
                btn_yes = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[text()='はい']")))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_yes)
                time.sleep(0.3)  # 少し待機
                self.driver.execute_script("arguments[0].click();", btn_yes)

            max_retry_download = 15
            retry_download = 0
            # 前日の発注明細をロード
            while retry_download <= max_retry_download:
                if os.path.exists(self.csv_path):
                    break
                elif retry_download == max_retry_download:
                    raise Exception("retry_download reached max_retry_download")
                else:
                    retry_download += 1
                    time.sleep(0.1)
        except:
            return "E0005" # ダウンロードエラー
        
        next_day = today_int + timedelta(days=1)
        days_jp = ['月', '火', '水', '木', '金', '土', '日']
        next_day_str = next_day.strftime('%Y-%m-%d') + f'({days_jp[next_day.weekday()]})'  # 日本語曜日を手動で追加

        logging.info(f"next_day_str: {next_day_str}")
        try:
            df = pd.read_csv(self.csv_path)
            logging.info(df)
        except FileNotFoundError:
            logging.error(f"Error: {self.csv_path} が見つかりませんでした。")
            return "E0006"
        # 納品予定日が翌日のデータをフィルタリング
        filtered_df = df[(df['納品予定日'] == next_day_str) | (df['納品日'] == next_day_str)] # 納品予定日に
        logging.info(filtered_df)

        # CSVファイルに上書き保存
        filtered_df.to_csv(self.csv_path, index=False, encoding='utf-8-sig')
        logging.info(f"{self.csv_path} にフィルタリング結果を上書き保存しました。")

        self.driver.close()
        self.driver.quit() 
        self.driver = None  

        if os.path.exists(self.csv_path):
            return "200" # OK
        else:
            return "E0005" # ダウンロードエラー

    def execute_with_retry(self, function_name, params, retries=3, timeout=120):
        for attempt in range(retries):
            try:
                response = self.call_google_script(function_name, params)
                return response
            except Exception as e:
                print(f"Attempt {attempt + 1} failed with error: {e}. Retrying...")
                time.sleep(5)
        raise Exception("All retry attempts failed")
           
    def generate_form(self, delivery_date_int, today_str):
        with open(self.csv_path, 'r', encoding='utf-8') as file:
            csv_data = file.read()

        params = {
            'delivery_date_int': delivery_date_int.isoformat(),
            'csv_data': csv_data,
            'today_str': today_str,
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
            values_food = response.get('values_food')
            values_nonfood = response.get('values_nonfood', None)
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
                self.input_df.dropna(subset=['商品コード'], inplace=True)
                self.input_df['商品コード'] = self.input_df['商品コード'].astype(str).str.strip()
                self.input_df['商品コード'] = pd.to_numeric(self.input_df['商品コード'], errors='coerce')
                self.input_df.dropna(subset=['商品コード'], inplace=True)
                Name_with_NaN = self.input_df[self.input_df['現在庫'].isna()]['商品名'].tolist()
                if len(Name_with_NaN) == 0:
                    self.input_df['商品コード'] = self.input_df['商品コード'].astype(int)
                    self.input_df['現在庫'] = self.input_df['現在庫'].astype(int)
                    self.input_df['発注数'] = self.input_df['発注数'].astype(int)
                    self.input_df['セット'] = self.input_df['セット'].astype(int)

                # 非食品
                if values_nonfood is not None:
                    max_columns_nonfood = len(values_nonfood[0])
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
                else:#非食品のシートがない場合
                    self.input_df_nonfood = False
                    
                return Name_with_NaN, self.input_df_nonfood
            
    def input_order_in_site(self):
        login_status_code = self.login_eos(st['EOS_ID'], st['EOS_PW'])
        logging.info(f'login_status_code(input_order_in_site): {login_status_code}')
        if login_status_code != "200":
            return False, login_status_code
        #ウィンドウを最大化
        self.driver.maximize_window()

        # 左メニューの発注入力をクリック
        self.driver.find_element(By.CLASS_NAME, 'menupng2').click()
        input_order = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@accesskey='3']")))
        input_order.click()
        
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'scode'))) 
        logging.info('scode loaded')
        
        # お知らせが表示される場合は✕ボタン(締め時間変更後は必要なし)
        try:
            WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.XPATH, "//button[@title='Close']")))
            logging.info('Dialog found')
        except TimeoutException:
            logging.info('No dialog found')
        while len(self.driver.find_elements(By.XPATH, value="//button[@title='Close']"))>0 :
            self.driver.find_element(By.XPATH, value="//button[@title='Close']").click()
            logging.info('Dialog closed')
            time.sleep(0.05)

        
        if isinstance(self.input_df_nonfood, bool):
            input_df_tuple = (self.input_df,)
        elif isinstance(self.input_df_nonfood, pd.DataFrame):
            input_df_tuple = (self.input_df, self.input_df_nonfood)
        else:
            input_df_tuple = (self.input_df,)
            logging.warning('input_df_nonfood is neither bool nor DataFrame')

        error_ls = [] #入力エラーの空リストを作成
        for df in input_df_tuple:
            if df is self.input_df:
                print(f'df is self.input')
            elif df.shape[0] == 0:
                print(f'df is False')
                continue
            else:
                print(f'df is nonfood')
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

                            #former_max_order_valueがsetvalueで割り切れない場合があるので、setvalueを足していき、former_max_order_valueを超えないさいだいのsetvalueの倍数にする
                            former_max_order_value = int(self.driver.find_element(By.ID, former_table_id).get_attribute('data-sgosuu')) - 1
                            former_max_input = (former_max_order_value // former_set_value) * former_set_value

                            print(f'商品名：{dict_data[former_input_number]} 発注数MAX超え：{former_max_order_value}→{former_max_input}')
                            input_field.send_keys(former_max_input) #発注数を入力
                        else:
                            error_ls.append(f'{former_input_number}：{dict_data[former_input_number]}（エラー理由：不明）')
                            self.driver.find_element(By.CLASS_NAME, 'ui-icon-closethick').click() # ×ボタンでダイアログを閉じる
                        
                        input_field = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, table_id)))
                        input_field.clear() #input_fieldのデフォルト0をクリア   
                        input_field.send_keys(order_value) #発注数を入力
                former_set_value = set_value
                former_table_id = table_prdx_number_dict[row.商品コード]
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
