# © 2024 Keita Iwasa

from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import os
import socket
import logging
import subprocess
import sys
import pkg_resources
import http.client
import urllib.parse
import shutil
import configparser
import win32print
import qrcode
from PIL import Image, ImageTk

try:
    from plyer import notification
    plyer_installed = True
except:
    pass
    plyer_installed = False
# install required library---------------------------------------------------------------------------
with open('setup/requirements.txt', 'r') as file:
        requirements = file.readlines()
installing_library = False
for requirement in requirements:
    try:
        pkg_resources.require(requirement)
    except (pkg_resources.DistributionNotFound, pkg_resources.VersionConflict):
        if installing_library is False and plyer_installed:
            installing_library = True
            notification.notify(
                app_name = "KAOS",
                app_icon = "setup/KAOS_icon.ico",
                title = "KAOS",
                message="\n自動発注システムをアップデートしています。少々お待ちください。",
                timeout=5
            )
        print(f'installing {requirement}...')
        subprocess.check_call([sys.executable, "-m", "pip", "install", requirement.strip()])
#---------------------------------------------------------------------------------------------------

from freezegun import freeze_time

# 設定ファイルの読み込み
if not os.path.exists('config.ini'):
    timestamp = datetime.now()
    file_content = f""";{timestamp}
[Settings]
comp = False
SHOP_NAME = 
EOS_ID = 
EOS_PW = 
PRINTER_NAME = 
"""
    with open('config.ini', "w", encoding="utf-8") as file:
        file.write(file_content)
config = configparser.ConfigParser()
with open('config.ini', 'r', encoding='utf-8') as file:
    config.read_file(file)
st = config['Settings']


from Automation import AutomationHandler

# Error handling ---------------------------------------------------
error_occurred = False

def send_line_notify(message):
    conn = http.client.HTTPSConnection("notify-api.line.me")
    token = "TpqTtEd7GR9lVkzFfPA2EL72wVAOtEhZQwIbkJjXncd"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Authorization": "Bearer " + token
    }
    params = urllib.parse.urlencode({"message": message})
    conn.request("POST", "/api/notify", params, headers)
    response = conn.getresponse()
    print(response.status, response.reason)

def handle_exception(exc):
    print(exc)
    global error_occurred
    global log_file
    error_occurred = True
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_filename = f"error_log_{st['SHOP_NAME']}{current_time}.txt"
    log_file = os.path.join('error_log', log_filename)
    logging.basicConfig(filename=log_file, level=logging.ERROR, format='%(asctime)s - %(message)s', encoding='utf-8')
    device_name=socket.gethostname()
    logging.error(f"Device: {device_name}")
    logging.error("Exception occurred", exc_info=exc)
    notify_message = f"{st['SHOP_NAME']}\n\nhttps://drive.google.com/drive/folders/1H7Izz-u465KTKz6JpxVVsraO97y3jpW5?usp=drive_link"
    send_line_notify(notify_message)
    messagebox.showerror("Error", "予期せぬエラーが発生しました。\nアプリを再起動してください。\n問題が解決しない場合は岩佐に連絡してください。")

def thread_with_error_handle(target, *args, **kwargs):
    try:
        target(*args, **kwargs)
    except Exception as e:
        handle_exception(e)

#-------------------------------------------------------------------------

class MainApplication(tk.Tk):            
    def show_frame(self, cont):
        frame = cont(self)
        frame.grid(row=0, column=0, sticky='nsew')
        frame.grid_propagate(False)
        contact_button = tk.Button(frame, text="お問い合わせ", command=self.show_qr)
        contact_button.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)
        frame.tkraise()

    def show_qr(self):
        qr_window = tk.Toplevel()
        qr_window.title("お問い合わせ用QRコード")
        photo = tk.PhotoImage(file="setup\岩佐LINEのQR.png")
        # 画像を表示するためのラベルウィジェットを作成
        label = tk.Label(qr_window, image=photo)
        label.image = photo  # 参照を保持
        label.pack()
        # 説明文を表示するラベルを追加
        description_label = tk.Label(qr_window, text=" こちらのQRコードから、担当/岩佐に電話orLINEしてください。", font=('Helvetica', 10))
        description_label.pack()
    
    def center_window(self, root, width, height):
        # スクリーンの幅と高さを取得
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # ウィンドウを中央に配置するためのxとy座標を計算
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))

        # ウィンドウの幅、高さ、および位置を設定
        root.geometry(f'{width}x{height}+{x}+{y}')
    
    def on_close(self):
        # ユーザーに確認メッセージを表示
        if messagebox.askyesno("終了確認", "本当に終了しますか？\n終了すると、自動で開かれたEOSのページも閉じます。"):
            self.handler.destroy_chrome()
            self.destroy()  # ウィンドウを閉じる
            self.quit()

    def __init__(self):
        super().__init__()
        self.title("コメダ自動発注システム KAOS")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.center_window(self, 600, 300)
        self.iconbitmap('setup\KAOS_icon.ico')
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.frames = {}

        #インスタンス変数を初期化
        self.night_order = None
        self.delivery_date_int =None
        self.today_real_int = None
        self.today_str = None
        self.today_str_csv = None
        self.yesterday_str = None
        self.today_order_file = None
        self.today_int = None

        self.dir_path = os.path.dirname(os.path.abspath(__file__)) 

        self.handler = AutomationHandler()
        if st['comp'] == "True":
            self.show_frame(Page_1)
        else:
            self.show_frame(Page_0)

class Text_and_Button_Page(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent, width=600, height=300)
        self.label1 = tk.Label(self, text="", wraplength=450)
        self.label1.pack(pady=(50,20))
        self.button1 = tk.Button(self, text="")
        self.button1.pack()

class Text_and_2Buttons_Page(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent, width=400, height=300)
        self.label2 = tk.Label(self, text="", wraplength=450)
        self.label2.grid(row=0, column=0, columnspan=2, pady=(50,20))
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.button_L = tk.Button(self, text="既存の発注書", command=lambda:parent.show_frame(Page_6))
        self.button_L.grid(row=1, column=0)
        self.button_R = tk.Button(self, text="新規の発注書", command=lambda:parent.show_frame(Page_4))
        self.button_R.grid(row=1, column=1)           

class Progress_Page(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent, width=400, height=300)
        self.label_p = tk.Label(self, text="", wraplength=450)
        self.label_p.pack(pady=(50,20))
        self.progress = ttk.Progressbar(self, orient="horizontal", mode="indeterminate")
        self.progress.pack(pady=30)
        self.progress.start(10)

class List_Page(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent, width=400, height=300)
        self.label_l = tk.Label(self, text="", wraplength=450)
        self.label_l.pack(side=tk.TOP, pady=(50,20))
        self.listbox_frame = tk.Frame(self, width=350, height=200)
        self.listbox_frame.pack_propagate(False)
        self.listbox_frame.pack(padx=50)
        self.button_l = tk.Button(self.listbox_frame, text="")
        self.button_l.pack(side=tk.BOTTOM, pady=20)
        self.scroll_y = tk.Scrollbar(self.listbox_frame)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.scroll_x = tk.Scrollbar(self.listbox_frame, orient=tk.HORIZONTAL)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.listbox = tk.Listbox(self.listbox_frame, yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set) 
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_y.config(command=self.listbox.yview)
        self.scroll_x.config(command=self.listbox.xview)

class Page_0(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, width=400, height=300)
        if st["comp"] == "True":
            back_button = tk.Button(self, text='戻る', command=lambda: parent.show_frame(Page_1))
            back_button.place(relx=0, rely=0, anchor='nw', x=10, y=10)

        df_SHOP_NAME = st['SHOP_NAME']
        df_EOS_ID = st['EOS_ID']
        df_EOS_PW = st['EOS_PW']
        df_PRINTER_NAME = st['PRINTER_NAME']
        if df_PRINTER_NAME == '':
            df_PRINTER_NAME = win32print.GetDefaultPrinter()


        # サブフレームの作成
        self.sub_frame = tk.Frame(self)
        self.sub_frame.pack(expand=True)

        # 店舗名（文字列で入力）
        self.shop_name_label = tk.Label(self.sub_frame, text="店舗名")
        self.shop_name_label.grid(row=0, column=0, padx=10, pady=(0,5), sticky="w")
        self.shop_name_entry = tk.Entry(self.sub_frame, width=25)
        self.shop_name_entry.grid(row=0, column=1, padx=0, pady=5)
        self.shop_name_entry.insert(0, df_SHOP_NAME)

        # EOSユーザーID（文字列で入力）
        self.eos_user_id_label = tk.Label(self.sub_frame, text="EOSユーザーID")
        self.eos_user_id_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.eos_user_id_entry = tk.Entry(self.sub_frame, width=25)
        self.eos_user_id_entry.grid(row=1, column=1, padx=10, pady=5)
        self.eos_user_id_entry.insert(0, df_EOS_ID)

        # EOSパスワード（文字列で入力）
        self.eos_password_label = tk.Label(self.sub_frame, text="EOSパスワード")
        self.eos_password_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.eos_password_entry = tk.Entry(self.sub_frame, width=25)
        self.eos_password_entry.grid(row=2, column=1, padx=10, pady=5)
        self.eos_password_entry.insert(0, df_EOS_PW)

        # 既定のプリンター（リストから選択）
        self.printer_label = tk.Label(self.sub_frame, text="既定のプリンター")
        self.printer_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        # インストールされているプリンタの一覧を取得します。
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
        # 取得したプリンタの情報からプリンタ名のみを抽出してリスト化
        self.printer_names = [printer[2] for printer in printers]
        self.printer_combobox = ttk.Combobox(self.sub_frame, values=self.printer_names, width=22)
        self.printer_combobox.grid(row=3, column=1, padx=10, pady=5)
        self.printer_combobox.set(df_PRINTER_NAME)

        # 保存ボタン
        self.save_button = tk.Button(self.sub_frame, text="保存", command=self.save_settings)
        self.save_button.grid(row=4, column=0, columnspan=2, pady=10)

    def save_settings(self):
        store_name = self.shop_name_entry.get()
        eos_user_id = self.eos_user_id_entry.get()
        eos_password = self.eos_password_entry.get()
        printer = self.printer_combobox.get()

        if store_name and eos_user_id and eos_password and printer:
            timestamp = datetime.now()
            file_content = f""";{timestamp}
[Settings]
comp = True
SHOP_NAME = {store_name}
EOS_ID = {eos_user_id}
EOS_PW = {eos_password}
PRINTER_NAME = {printer}
        """        
            if messagebox.askokcancel("設定の保存","現在の入力で設定を保存しますか？", detail="保存するとアプリが再起動します。"):
                file_name = "config.ini"
                # ファイルを作成して内容を書き込みます
                if os.path.exists(file_name):
                    os.remove(file_name)
                with open(file_name, "w", encoding="utf-8") as file:
                    file.write(file_content)
                messagebox.showinfo('保存完了','設定が保存されました。')  
                python = sys.executable
                os.execl(python, python, *sys.argv)     
        else:
            messagebox.showwarning("入力エラー", "全ての項目を入力してください。")  

class Page_1(Text_and_Button_Page):
    def __init__(self, parent):
        super().__init__(parent)
        setting_icon = tk.PhotoImage(file="setup/setting_icon.png").subsample(2,2)
        setting_button = tk.Button(self, image=setting_icon, compound="top", command=lambda: parent.show_frame(Page_0))
        setting_button.image = setting_icon
        setting_button.place(relx=0, rely=0, anchor='nw', x=10, y=10)
        parent.attributes("-topmost", True)
        parent.attributes("-topmost", False)
        # 本日の日付を取得 as YYYY-MM-DD
        parent.today_int = datetime.today()
        parent.today_str = parent.today_int.strftime('%Y-%m-%d')
        # 前日の日付を取得 as YYYY-MM-DD
        yesterday_int = parent.today_int - timedelta(days=1)
        parent.yesterday_str = yesterday_int.strftime('%Y-%m-%d') 
        #発注開始の確認
        if datetime.now().hour >= 14: #夜発注14時以降
            parent.night_order = True
            parent.delivery_date_int = parent.today_int + timedelta(days=2) #納品日(夜発注)明後日
            delivery_date_str = parent.delivery_date_int.strftime('%Y-%m-%d')
            self.label1.config(text=f'{parent.today_str}の夜分の発注作業を開始します。納品は{delivery_date_str}です。')
            parent.today_real_int = parent.today_int #本日の本当の日付   
            self.button1.config(text="OK", command=lambda: parent.show_frame(Page_2))     
        elif datetime.now().hour < 12 : #朝発注12時まで
            parent.night_order = False
            parent.delivery_date_int = parent.today_int + timedelta(days=1) #納品日(朝発注)明日
            delivery_date_str = parent.delivery_date_int.strftime('%Y-%m-%d')
            self.label1.config(text=f'{parent.today_str}の朝分の発注作業を開始します。納品は{delivery_date_str}です。')
            parent.today_real_int = parent.today_int #本日の本当の日付
            parent.today_int = parent.today_int - timedelta(days=1) #朝発注の場合は日付を-1
            self.button1.config(text="OK", command=lambda: parent.show_frame(Page_2))
        else :
            self.label1.config(text="EOSの発注停止中のため、発注できません。")
            self.button1.config(text="発注を中止", command=parent.quit)
        if parent.today_real_int:
            parent.today_str_csv = parent.today_real_int.strftime('%Y%m%d')
        self.bind_all('<Control-m>', lambda e: self.place_entry(parent)) #開発用、任意の時刻を設定するコマンド

    def place_entry(self,parent):
        self.label1.config(text="日時を 'YYYY-MM-DD HH:MM:SS' 形式で入力してください")
        self.button1.forget()
        self.datetime_entry = tk.Entry(self)
        self.datetime_entry.pack(pady=20)
        self.datetime_entry.focus_set()
        self.datetime_entry.bind('<Return>', lambda e:self.datetime_setting(parent))

    def datetime_setting(self, parent):
        input_date = self.datetime_entry.get()
        dt = datetime.strptime(input_date, "%Y-%m-%d %H:%M:%S")
        with freeze_time(dt):
            # 固定された現在時刻を表示
            print("固定された現在時刻: ", datetime.now())
            parent.show_frame(Page_1)     

class Page_2(Text_and_Button_Page):
    def __init__(self, parent):
        super().__init__(parent)

        self.label1.config(text='発注が完了するまでこのウィンドウは閉じないでください。')
        parent.today_order_file = f'発注書_{parent.today_str}.xlsx' 
        if os.path.exists(parent.today_order_file): 
            self.button1.config(text="OK", command=lambda: parent.show_frame(Page_3))
        else:
            self.button1.config(text="OK", command=lambda: parent.show_frame(Page_4))

class Page_3(Text_and_2Buttons_Page): # すでに本日の発注書が存在する場合
    def __init__(self, parent):
        super().__init__(parent)
        self.label2.config(text=f"{parent.today_str}の発注書が既に存在します。\n既存のものを使用しますか？新規の発注書を生成しますか？")
        self.button_L.config(text="既存の発注書", command=lambda:parent.show_frame(Page_6))
        self.button_R.config(text="新規の発注書", command=lambda:parent.show_frame(Page_4))

class Page_4(Progress_Page): #発注書作成
    def __init__(self, parent):
        super().__init__(parent)
        self.label_p.config(text="本日の発注書を作成しています...")
        parent.attributes("-topmost", True)
        self.after(0, self.start_setup_form(parent))

    def start_setup_form(self, parent):
        threading.Thread(target=thread_with_error_handle, args=(self.setup_form, parent,),daemon=True).start()

    def setup_form(self, parent):
        download_success = False 
        if parent.night_order == True:
            download_success = parent.handler.download_csv(parent.today_real_int, parent.today_str_csv)
        if download_success or (parent.night_order == False):
            parent.spread_url = parent.handler.generate_form(parent.delivery_date_int, parent.today_str_csv, parent.yesterday_str, parent.today_str, parent.night_order)
            # QRコードの生成
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=3.5,
                border=3,
            )
            qr.add_data(parent.spread_url)
            qr.make(fit=True)

            # QRコードの画像を生成
            img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
            # 画像の背景を透明に変更
            data = img.getdata()
            new_data = []
            for item in data:
                # 白いピクセルを透明に設定
                if item[:3] == (255, 255, 255):
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)
            img.putdata(new_data)
            parent.qr_img = ImageTk.PhotoImage(img)
            self.progress.stop()
            parent.show_frame(Page_6)
        else:
            parent.show_frame(Page_5)

class Page_5(Text_and_Button_Page): #発注明細ダウンロード失敗
    def __init__(self, parent):
        super().__init__(parent)
        parent.attributes("-topmost", True)
        parent.attributes("-topmost", False)
        self.label1.config(text="発注明細のダウンロードに失敗しました。手動で前日の発注明細をダウンロードしてください。(ファイル名は変えないでください。)")
        self.button1.config(text="ダウンロードした", command=lambda: self.check_download(parent))        

    def check_download(self, parent): 
        if os.path.exists(f'{st.DOWNLOAD_DIR}/{parent.today_str_csv}_発注.CSV'):
            parent.show_frame(Page_4)
        else :
            parent.show_frame(Page_5)

class Page_6(Text_and_Button_Page): #発注書生成完了&入力確認
    def __init__(self, parent):
        super().__init__(parent)
        parent.attributes("-topmost", False)
        self.label1.config(text="発注書が作成されました。タブレットでQRコードを読み込み、現在庫数を入力してください。") 
        self.label1.pack(pady=(50,10))
        self.button1.pack_forget()
        label_qr = tk.Label(self, image=parent.qr_img)
        label_qr.pack()
        self.button1 = tk.Button(self, text="入力完了", command=lambda:self.confirm_filling(parent))
        self.button1.pack()
    
    def confirm_filling(self, parent):
        self.label1.config(text="タブレットで現在庫数をすべて入力しましたか？")
        self.button1.config(text="はい", command=lambda: parent.show_frame(Page_7))

class Page_7(Progress_Page): #同期確認
    def __init__(self, parent):
        super().__init__(parent)
        self.label_p.config(text="発注書からデータを取得しています...")
        self.after(0, self.start_confirm_synch(parent))
        
    def start_confirm_synch(self, parent):
        threading.Thread(target=thread_with_error_handle, args=(self.confirm_googledive_sinch, parent,),daemon=True).start()

    def confirm_googledive_sinch(self, parent):
        NaN_ls = parent.handler.get_spreadsheet(parent.today_str) #戻り値は現在庫が入力されてない商品名のリスト           
        if NaN_ls is False:
            self.progress.stop()
            self.progress.pack_forget()
            self.label_p.config(text='データの取得に失敗しました。') 
            self.button2 = tk.Button(self, text="再試行", command=lambda: parent.show_frame(Page_7))     
            self.button2.pack()
        elif len(NaN_ls) == 0:
            self.label_p.config(text="EOSへ発注数を入力中...")
            parent.attributes("-topmost", True)
            input_order_success, parent.error_ls = parent.handler.input_order_in_site(parent.today_order_file)
            self.progress.stop()
            if input_order_success: 
                parent.show_frame(Page_8)         
        else: 
            parent.NaN_ls = NaN_ls
            self.progress.stop()
            parent.show_frame(Page_7i)

class Page_7i(List_Page):
    def __init__(self, parent):
        super().__init__(parent)
        self.label_l.config(text="以下の商品の現在庫が入力されていません。発注書を確認してください。")
        for value in parent.NaN_ls:
                self.listbox.insert(tk.END, value)
        self.button_l.config(text="発注書を確認した", command=lambda: parent.show_frame(Page_7))

class Page_8(List_Page):
    def __init__(self, parent):
        super().__init__(parent)
        parent.attributes("-topmost", True)
        parent.attributes("-topmost", False)
        if len(parent.error_ls) > 0:
            self.error_listbox(parent)
        else:
            self.last_frame(parent)

    def error_listbox(self, parent):
        self.label_l.config(text="以下の商品で入力エラーが発生しました。EOSで以下の商品を確認してください。")
        for value in parent.error_ls:
                self.listbox.insert(tk.END, value)
        self.button_l.config(text="確認した", command=lambda: self.last_frame(parent))
  
    def last_frame(self, parent):
        self.label_l.config(text="発注数の入力が完了しました。以下のことを手動で行ってください。")
        self.button_l.pack_forget()
        self.listbox.pack_forget()
        self.scroll_y.pack_forget()
        self.scroll_x.pack_forget()
        self.listbox_frame.pack_forget()
        todo_list = []
        if parent.today_int.weekday() in {1, 3, 5}: #本日が火・木・土の場合
            todo_list.append('EOSで「発注手続きへ」をクリックして、発注確定する。')
            todo_list.append('今日は非食品の発注日です。非食品の発注数を入力して、再度発注確定を行い、印刷する。')
        else:
            todo_list.append('EOSで「発注手続きへ」をクリックして、発注確定・印刷する。')

        if st['SHOP_NAME'] in {"自由が丘メープル通り", "岩佐Surface"}:
            todo_list.append('たまごを忘れず発注する。')
        todo_list.append('発注が終了したら、右上の✕ボタンで終了してください。')

        text = "\n".join(f"{index + 1}. {todo}" for index, todo in enumerate(todo_list))
        label_todo = tk.Label(self, text=text, justify='left', width=300)
        label_todo.pack(pady=20, padx=30)

        if parent.today_int.weekday() in {1, 3, 5}: #本日が火・木・土の場合
            try:
                print_seccess = parent.handler.print_excel(parent.today_order_file, '非食品', st['PRINTER_NAME'])
                if print_seccess is False:
                    raise Exception
            except:
                messagebox.showwarning("印刷失敗", "非食品の発注書の印刷に失敗しました。ファイルを開いて手動で印刷してください。")

def hide_files():
    """    
    directory (str): 操作対象のディレクトリ
    exclude_files (list): 除外するファイルのリスト
    exclude_dirs (list): 除外するディレクトリのリスト
    """
    directory = os.path.dirname(os.path.abspath(__file__))  # スクリプトが含まれるディレクトリ
    exclude_files = [os.path.basename(__file__)] + [f for f in os.listdir(directory) if f.endswith('.xlsx') or f.endswith('.xls')]
    exclude_dirs = ['発注書Excel']

    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        # ファイルが除外リストにない場合
        if os.path.isfile(item_path) and item not in exclude_files:
            os.system(f'attrib +h "{item_path}"')
        # ディレクトリが除外リストにない場合
        elif os.path.isdir(item_path) and item not in exclude_dirs:
            os.system(f'attrib +h "{item_path}"')

if __name__ == "__main__":
    try:
        app = MainApplication()
        app.mainloop()
    except Exception as e:
        handle_exception(e)
    hide_files()
        
