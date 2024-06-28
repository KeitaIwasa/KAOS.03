# © 2024 Keita Iwasa

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import time
import random
import shutil
import os
from openpyxl import Workbook,load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill
import pandas as pd
import numpy as np
import glob
import win32api
import win32print
import openpyxl
import winreg
import configparser
import qrcode
import httplib2
import json
from datetime import datetime, timedelta
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow

# 設定ファイルの読み込み
config = configparser.ConfigParser()
with open('config.ini', 'r', encoding='utf-8') as file:
    config.read_file(file)
st = config['Settings'] 

class AutomationHandler:
    def __init__(self):
        SCOPES = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/script.projects',
            'https://www.googleapis.com/auth/calendar.readonly'
        ]
        token_path = 'setup/token.json'
        global creds
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        if not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            if not os.access(token_path, os.W_OK):
                os.chmod(token_path, 0o666)
            with open(token_path, 'w') as token:
                token.write(creds.to_json())
        self.driver = None
        
    def register_drive_id(self, shopName):
        drive_service = build('drive', 'v3', credentials=creds)
        parent_folder = '1H7Izz-u465KTKz6JpxVVsraO97y3jpW5'
        new_folder_name = shopName
        query = f"mimeType='application/vnd.google-apps.folder' and name='{new_folder_name}' and '{parent_folder}' in parents and trashed=false"
        results = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        if len(items)==0:
            file_metadata = {
                'name': shopName,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder] # parent folder "KAOS発注書"
            }
            folder = drive_service.files().create(body=file_metadata, fields='id').execute()
            new_folder_id = folder.get('id')
            print('Folder created with ID:', new_folder_id)
            return new_folder_id
        else:
            return items[0]['id']
    
    def get_original_sheet(self):
        drive_service = build('drive', 'v3', credentials=creds)
        sheet_name = f'発注書【原本】_{st["SHOP_NAME"]}'
        query = f"'{st['SHOP_FOLDER_ID']}' in parents and name = '{sheet_name}' and trashed = false and mimeType='application/vnd.google-apps.spreadsheet'"
        results = drive_service.files().list(
            q=query,
            fields='files(id, name)').execute()
        items = results.get('files', [])
        if not items:
            return False
        else:
            sheet_id = items[0]['id']
            sheet_url=f'https://docs.google.com/spreadsheets/d/{sheet_id}/edit?'
            return sheet_url
        
    def find_folder_id(self, service, parent_folder_id, folder_name):
        query = f"'{parent_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and name = '{folder_name}' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        items = results.get('files', [])
        if not items:
            print(f'No folder found with name: {folder_name}')
            return None
        return items[0]['id']
    
    def find_files_id(self, service, folder_id, file_name):
        query = f"'{folder_id}' in parents and name = '{file_name}' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        items = results.get('files', [])
        if not items:
            print(f'No file found with name: {file_name}')
            return False
        return items

    def check_existing_sheet(self, sheet_name):
        drive_service = build('drive', 'v3', credentials=creds)
        orderform_folder = self.find_folder_id(drive_service, st['SHOP_FOLDER_ID'], '発注書')
        sheets = self.find_files_id(drive_service, orderform_folder, sheet_name)
        if sheets is False:
            return False
        else:
            sheet_id = sheets[0]['id']
            spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
            return sheet_id, spreadsheet_url    
        
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

    def loading_workbook(self,closing_file, max_attempts=3):
        try:
            workbook = load_workbook(closing_file)
            return workbook
        except PermissionError as e:
            if max_attempts > 0:
                print(f'\n{os.path.basename(e.filename)}を閉じてください！')
                time.sleep(1)
                input('再試行 "Enter":')
                return self.loading_workbook(closing_file, max_attempts-1)
            else:
                raise Exception('最大試行回数に達しました。')
            
    def download_folder_path(self):
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key)
        download_folder = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        return download_folder
            
    def download_csv(self, today_str_csv, today_int):
        self.login_eos(st['EOS_ID'], st['EOS_PW']) #EOSログインメソッド↑
        self.csv_path = f'{self.download_folder_path()}/{today_str_csv}_発注.CSV'
        self.csv_path_nonfood = f'{self.download_folder_path()}/{today_str_csv}_発注 (1).CSV'
        max_retry_download = 20
        if os.path.exists(self.csv_path):
            pass
        else:            
            # 左メニューの発注照会をクリック
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'menupng2'))).click() #発注
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@accesskey='4']"))).click() #発注照会クリック

            # 本日の発注を照会
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'inquiryButton'))).click() #照会ボタン
            time.sleep(0.5)
                
            # CSVをダウンロード
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'btnCsvoutConfirm'))).click() #CSV出力をクリック
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[text()='はい']"))).click() #ウィジェットのはいをクリック

            today_str_csv = datetime.today().strftime('%Y%m%d') #本日の日付（ダウンロードした発注明細csvは発注日付に関係なく本日のreal日付が付いている）
            retry_download = 0
            # 前日の発注明細をロード
            while retry_download <= max_retry_download:
                if os.path.exists(self.csv_path):
                    break
                elif retry_download == max_retry_download:
                    self.driver.maximize_window()
                    return False    
                else:
                    retry_download += 1
                    time.sleep(0.1)

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

    def execute_with_retry(self, service, request, script_id, retries=3, timeout=120):
        http = httplib2.Http(timeout=timeout)
        service._http = http
        
        for attempt in range(retries):
            try:
                response = service.scripts().run(body=request, scriptId=script_id).execute()
                return response
            except TimeoutError as e:
                print(f"Attempt {attempt + 1} failed with timeout. Retrying...")
                time.sleep(5)
            except Exception as e:
                print(f"Attempt {attempt + 1} failed with error: {e}. Retrying...")
                time.sleep(5)
        raise Exception("All retry attempts failed")
           
    def generate_form(self, delivery_date_int, today_str_csv, yesterday_str, today_str, night_order) :  
        if night_order:
            drive_service = build('drive', 'v3', credentials=creds)
            file_metadata = {
                'name': os.path.basename(f'{today_str_csv}_発注.CSV'),
                'parents': [st['SHOP_FOLDER_ID']]
            }
            media = MediaFileUpload(self.csv_path, mimetype='application/octet-stream')
            file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            csv_id = file.get("id")
            print(f'File ID: {csv_id}')
        else:
            csv_id = None

        script_service = build('script', 'v1', credentials=creds)
        script_id = 'AKfycbxK7pavgq0YZ-chJgYh_49eYCs0C6Gsm9RHBwpGIHFa4URkRXYivT8SeUVlt6nI-8Vbfg'
        request = {
            'function': 'generateForm',
            'parameters': [delivery_date_int.isoformat(), csv_id, today_str, night_order, st['SHOP_NAME']]
        }
        try:
            response = script_service.scripts().run(body=request, scriptId=script_id).execute()
            #発注明細csvをローカルから削除
            #response = self.execute_with_retry(script_service, request, script_id, retries=5)
            if 'error' in response:
                # エラーハンドリング
                error_message = 'Error: {}'.format(response['error']['details'])
                print(error_message)
                raise Exception(error_message)
            else:
                # デバッグ情報を出力
                debug_info_json = response['response'].get('result')
                debug_info = json.loads(debug_info_json)
                print('Debug Info:', debug_info)

                self.new_spreadsheet_id = debug_info.get('spreadsheetId')
                print('New Spreadsheet ID: {}'.format(self.new_spreadsheet_id))

                spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{self.new_spreadsheet_id}/edit" 
                
                return self.new_spreadsheet_id, spreadsheet_url
        except Exception as e:
            print(f"Failed to execute the script: {e}")
            raise Exception(e)

    def get_spreadsheet(self, sheet_id):
        sheets_service = build('sheets', 'v4', credentials=creds)
        time.sleep(6) #spreadsheetがタブレットからドライブに同期されるのを待つため
        
        # シートのデータを取得
        sheet = sheets_service.spreadsheets()
        result_food = sheet.values().get(spreadsheetId=sheet_id, range='食品').execute()
        result_nonfood = sheet.values().get(spreadsheetId=sheet_id, range='非食品').execute()
        values_food = result_food.get('values', [])
        values_nonfood = result_nonfood.get('values', [])
        # 指定した複数の列をDataFrameに変換
        if not values_food:
            print('No data found in the sheet.')
            return False, False
        else:
            #食品
            max_columns = len(values_food[0])
            data = [row + [np.nan] * (max_columns - len(row)) for row in values_food[1:]]
            df = pd.DataFrame(data, columns=values_food[0])  # 1行目はヘッダーとして利用
            self.input_df = df[['商品名', 'セット','商品コード', '現在庫', '発注数']]
            self.input_df.replace('', np.nan, inplace=True)
            Name_with_NaN = self.input_df[self.input_df['現在庫'].isna()]['商品名'].tolist()
            if len(Name_with_NaN) == 0:
                # 整数型に変換
                self.input_df['商品コード'] = self.input_df['商品コード'].astype(int)
                self.input_df['現在庫'] = self.input_df['現在庫'].astype(int)
                self.input_df['発注数'] = self.input_df['発注数'].astype(int)
                self.input_df['セット'] = self.input_df['セット'].astype(int)

            #非食品
            max_columns = len(values_nonfood[1])#商品名の列
            data = [row + [np.nan] * (max_columns - len(row)) for row in values_nonfood[1:]]
            df = pd.DataFrame(data, columns=values_nonfood[0])
            input_df_nonfood = df[['商品名', '商品コード', '発注数']]
            input_df_nonfood['発注数'] = input_df_nonfood['発注数'].astype(str).str.strip()
            input_df_nonfood['発注数'] = pd.to_numeric(input_df_nonfood['発注数'], errors='coerce')
            # '発注数'がNaNまたは0の行を削除
            input_df_nonfood = input_df_nonfood.dropna(subset=['発注数'])
            input_df_nonfood = input_df_nonfood[input_df_nonfood['発注数'] != 0]
            # 行数が0の場合はFalseに設定
            if not input_df_nonfood.shape[0] == 0: 
                self.input_df_nonfood = input_df_nonfood.reset_index(drop=True)
                self.input_df_nonfood.replace('', np.nan, inplace=True)
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

        #if os.path.exists(st.WIN32COM_GEN_poPY_DIR):
        #    shutil.rmtree(st.WIN32COM_GEN_PY_DIR) # win32comのキャッシュをフォルダごと削除（これをしないとエラーが起こる）
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'scode'))) 
        
        input_df_tuple = (self.input_df, self.input_df_nonfood)
        error_ls = [] #入力エラーの空リストを作成
        for df in input_df_tuple:
            if df is self.input_df:
                print(f'df is self.input')
            elif df is False:
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
