# © 2024 Keita Iwasa

from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import os
import logging
import logging.handlers
import sys
import http.client
import urllib.parse
import configparser
import webbrowser
from freezegun import freeze_time
import qrcode
from PIL import ImageTk
import traceback

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 設定ファイルの読み込み
#setupフォルダがない場合は作成
if not os.path.exists('setup'):
    os.makedirs('setup')
if not os.path.exists(r'setup/config.ini'):
    timestamp = datetime.now()
    file_content = f""";{timestamp}
[Settings]
comp = False
SHOP_NAME = 
EOS_ID = 
EOS_PW = 
"""
    with open(r'setup/config.ini', "w", encoding="utf-8") as file:
        file.write(file_content)
config = configparser.ConfigParser()
with open(r'setup/config.ini', 'r', encoding='utf-8') as file:
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

def handle_exception(exc, message=False):
    print(exc)
    logging.exception(exc)
    global error_occurred
    global log_file
    error_occurred = True   
    error_message = traceback.format_exc()

    # LINE Notify で通知
    try:
        notify_message = f"{st['SHOP_NAME']}\n\n{error_message}"
        send_line_notify(notify_message)
    except Exception as e:
        logging.error(f"Failed to send LINE Notify: {e}")
    if not message:
        messagebox.showerror("Error", "予期せぬエラーが発生しました。\nアプリを再起動してください。\n問題が解決しない場合は岩佐に連絡してください。")
    else:
        messagebox.showerror("Error", message)

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
        photo = tk.PhotoImage(file=resource_path('setup/岩佐LINEのQR.png'))
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
            if not error_occurred:
                try:
                    logging.shutdown() 
                    os.remove(log_file)
                except Exception as e:
                    logging.error(f"ログファイルの削除に失敗しました: {e}")
            if os.path.exists(f'{self.handler.download_folder_path()}/{self.today_str_csv}_発注.CSV'):
                os.remove(f'{self.handler.download_folder_path()}/{self.today_str_csv}_発注.CSV')
            self.destroy()  # ウィンドウを閉じる
            self.quit()

    def __init__(self):
        super().__init__()
        self.title("コメダ自動発注システム KAOS")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.center_window(self, 600, 350)
        self.iconbitmap(resource_path('setup/KAOS_icon.ico'))
        self.option_add("*Font", ("Yu Gothic UI", 11))
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

class ToolTip():
    def __init__(self, widget, text="default tooltip"):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Motion>", self.motion)
        self.widget.bind("<Leave>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event):
        self.schedule()
    
    def motion(self, event):
        self.unschedule()
        self.schedule()
    
    def leave(self, event):
        self.unschedule()
        self.id = self.widget.after(100, self.hideTooltip)
    
    def schedule(self):
        if self.tw:
            return
        self.unschedule()
        self.id = self.widget.after(500, self.showTooltip)
    
    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)
    
    def showTooltip(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)
        x, y = self.widget.winfo_pointerxy()
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.geometry(f"+{x+10}+{y+10}")
        label = tk.Label(self.tw, text=self.text, background="lightyellow",
                         relief="solid", borderwidth=1, justify="left")
        label.pack(ipadx=10)

    def hideTooltip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()

class Page_0(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, width=400, height=300)

        if st["comp"] == "True":
            back_button = tk.Button(self, text='戻る', command=lambda: parent.show_frame(Page_1))
            back_button.place(relx=0, rely=0, anchor='nw', x=10, y=10)

        df_SHOP_NAME = st['SHOP_NAME']
        df_EOS_ID = st['EOS_ID']
        df_EOS_PW = st['EOS_PW']


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

        # 保存ボタン
        self.save_button = tk.Button(self.sub_frame, text="保存", command=lambda:self.save_settings(parent))
        self.save_button.grid(row=4, column=0, columnspan=2, pady=10)

    def save_settings(self, parent):
        shop_name = self.shop_name_entry.get()
        eos_user_id = self.eos_user_id_entry.get()
        eos_password = self.eos_password_entry.get()

        if shop_name and eos_user_id and eos_password:
            if messagebox.askokcancel("設定の保存","現在の入力で設定を保存しますか？", detail="保存するとアプリが再起動します。"):
                timestamp = datetime.now()
                file_content = f""";{timestamp}
[Settings]
comp = True
SHOP_NAME = {shop_name}
EOS_ID = {eos_user_id}
EOS_PW = {eos_password}
        """        
                file_name = r"setup/config.ini"
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
        #設定ボタン
        setting_icon = tk.PhotoImage(file=resource_path('setup/setting_icon.png')).subsample(2,2)
        setting_button = tk.Button(self, image=setting_icon, compound="top", command=lambda: parent.show_frame(Page_0))
        setting_button.image = setting_icon
        setting_button.place(relx=0, rely=0, anchor='nw', x=10, y=10)
        setting_button_tooltip = ToolTip(setting_button, "設定") 

        #発注書変更ボタン
        sheet_icon = tk.PhotoImage(file=resource_path('setup/sheet_icon.png')).subsample(2,2)
        sheet_button = tk.Button(self, image=sheet_icon, compound="top", command=lambda: self.open_original_sheet(parent))
        sheet_button.image = sheet_icon
        sheet_button.place(relx=0, rely=0, anchor='nw', x=41, y=10)
        sheet_button_tooltip = ToolTip(sheet_button, "発注書[原本]を編集") 

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
        self.bind_all('<Control-t>', lambda e: self.place_entry(parent)) #開発用、任意の時刻を設定するコマンド

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

    def open_original_sheet(self, parent):
        parent.show_frame(Page_OS)

class Page_OS(Progress_Page):
    def __init__(self, parent):
        super().__init__(parent)
        self.label_p.config(text="発注書の原本を開いています...")
        self.after(0, self.start_open_original_sheet(parent))

    def start_open_original_sheet(self, parent):
        threading.Thread(target=thread_with_error_handle, args=(self.open_original_sheet, parent,),daemon=True).start()

    def open_original_sheet(self, parent):
        try:
            original_sheet_url = parent.handler.get_original_sheet()
            if original_sheet_url == False:
                raise Exception("Original sheet not found")
            else:
                webbrowser.open(original_sheet_url)
            parent.show_frame(Page_1)
        except Exception as e:
            handle_exception(e, message="発注書の原本が見つかりません。「お問い合わせ」から担当者に連絡してください。")

class Page_2(Text_and_Button_Page):
    def __init__(self, parent):
        super().__init__(parent)
        self.label1.config(text='発注が完了するまでこのウィンドウは閉じないでください。') 
        self.button1.config(text="OK", command=lambda: parent.show_frame(Page_2i))

class Page_2i(Progress_Page): # 既存の発注書の存在確認
    def __init__(self, parent):
        super().__init__(parent)
        self.label_p.config(text="作成済みの発注書が存在するか確認中...")
        if parent.night_order == True:
            parent.today_order_file = f'発注書_{parent.today_str}PM'
        else:
            parent.today_order_file = f'発注書_{parent.today_str}AM'
        self.after(0, self.start_check_form(parent))

    def start_check_form(self, parent):
        threading.Thread(target=thread_with_error_handle, args=(self.check_form, parent,),daemon=True).start()

    def check_form(self, parent):
        check_result = parent.handler.check_existing_sheet(parent.today_order_file)
        if check_result == False:
            parent.sheet_id, parent.sheet_url = False, False
            parent.show_frame(Page_4)
        else:
            parent.sheet_id, parent.sheet_url = check_result
            parent.show_frame(Page_3)

class Page_3(Text_and_2Buttons_Page): # すでに本日の発注書が存在する場合
    def __init__(self, parent):
        super().__init__(parent)
        self.label2.config(text=f"{parent.today_str}の発注書が既に存在します。\n既存のものを使用しますか？新規の発注書を生成しますか？")
        parent.attributes("-topmost", True)
        parent.attributes("-topmost", False)
        self.button_L.config(text="既存の発注書", command=lambda:parent.show_frame(Page_6))
        self.button_R.config(text="新規の発注書", command=lambda:parent.show_frame(Page_4))

class Page_4(Progress_Page): #発注書作成
    def __init__(self, parent):
        super().__init__(parent)
        self.label_p.config(text="本日の発注書を作成しています...")
        self.after(0, self.start_setup_form(parent))

    def start_setup_form(self, parent):
        threading.Thread(target=thread_with_error_handle, args=(self.setup_form, parent,),daemon=True).start()

    def setup_form(self, parent):
        download_success = False 
        if parent.night_order == True:
            download_success = parent.handler.download_csv(parent.today_str_csv, parent.today_int)
        if download_success or (parent.night_order == False):
            generate_result = parent.handler.generate_form(parent.delivery_date_int, parent.today_str, parent.night_order)
            if generate_result == False:
                raise Exception
            else:
                parent.sheet_id, parent.sheet_url = generate_result
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
        if os.path.exists(f'{parent.handler.download_folder_path()}/{parent.today_str_csv}_発注.CSV'):
            parent.show_frame(Page_4)
        else :
            parent.show_frame(Page_5)

class Page_6(Text_and_Button_Page): #発注書生成完了&入力確認
    def __init__(self, parent):
        super().__init__(parent)
        parent.attributes("-topmost", True)
        parent.attributes("-topmost", False)
        self.label1.config(text="発注書が作成されました。タブレットでQRコードを読み込み、現在庫数を入力してください。") 
        self.label1.pack(pady=(50,10))
        self.qr_img = self.make_qr(parent.sheet_url)
        label_qr = tk.Label(self, image=self.qr_img)
        label_qr.pack()
        self.button1.pack_forget()
        self.button1 = tk.Button(self, text="入力完了", command=lambda:self.confirm_filling(parent))
        self.button1.pack()
        parent.nonfood0_ok = False
    
    def confirm_filling(self, parent):
        self.label1.config(text="タブレットで現在庫数をすべて入力しましたか？")
        self.button1.config(text="はい", command=lambda: parent.show_frame(Page_7))

    def make_qr(self, url):
        # QRコードの生成
        print(url)
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=3.5,
            border=3,
        )
        qr.add_data(url)
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
        qr_img = ImageTk.PhotoImage(img)
        return qr_img


class Page_7(Progress_Page): #発注数取得・入力
    def __init__(self, parent):
        super().__init__(parent)
        self.label_p.config(text="発注書からデータを取得しています...")
        self.after(0, self.start_confirm_synch(parent))
        
    def start_confirm_synch(self, parent):
        threading.Thread(target=thread_with_error_handle, args=(self.confirm_googledive_sinch, parent,),daemon=True).start()

    def confirm_googledive_sinch(self, parent):
        NaN_ls, df_nonfood = parent.handler.get_spreadsheet(parent.sheet_id) #戻り値は現在庫が入力されてない商品名のリストと非食品のdf          
        if NaN_ls is False:
            self.progress.stop()
            self.progress.pack_forget()
            self.label_p.config(text='データの取得に失敗しました。') 
            self.button2 = tk.Button(self, text="再試行", command=lambda: parent.show_frame(Page_7))     
            self.button2.pack()
        elif df_nonfood.shape[0] == 0 and parent.today_int.weekday() in {1, 3, 5} and parent.nonfood0_ok == False:
            self.progress.stop()
            self.progress.pack_forget()
            parent.show_frame(Page_7ii)     
        elif len(NaN_ls) == 0:
            self.label_p.config(text="EOSへ発注数を入力中...")
            parent.attributes("-topmost", True)
            input_order_success, parent.error_ls = parent.handler.input_order_in_site()
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

class Page_7ii(Text_and_2Buttons_Page): # 非食品が未入力
    def __init__(self, parent):
        super().__init__(parent)
        self.label2.config(text='今日は非食品の発注日です。非食品の現在庫が入力されていませんが、非食品は発注しなくてよろしいですか？発注する場合は、発注書に現在庫を入力してから「再試行」をクリックしてください。')
        self.button_L.config(text='再試行', command=lambda:parent.show_frame(Page_7))
        self.button_R.config(text='非食品は発注しない', command=lambda:self.nonfood0_ok(parent))
    
    def nonfood0_ok(self, parent):
        parent.nonfood0_ok = True
        parent.show_frame(Page_7)

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
        todo_list.append('EOSで「発注手続きへ」をクリックして、発注確定・印刷する。')

        if st['SHOP_NAME'] in {"自由が丘メープル通り", "岩佐Surface"}:
            todo_list.append('たまごを忘れず発注する。')
        todo_list.append('発注が終了したら、右上の✕ボタンで終了してください。')

        text = "\n".join(f"{index + 1}. {todo}" for index, todo in enumerate(todo_list))
        label_todo = tk.Label(self, text=text, justify='left', width=300)
        label_todo.pack(pady=20, padx=30)

if __name__ == "__main__":
    # logging設定
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_filename = f"error_log_{st['SHOP_NAME']}{current_time}.txt"
    error_log_dir = 'error_log'
    if not os.path.exists(error_log_dir):
        os.makedirs(error_log_dir)
    log_file = os.path.join('error_log', log_filename)
    logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(message)s', encoding='utf-8')
    
    try:
        app = MainApplication()
        app.mainloop()
    except Exception as e:
        handle_exception(e)
        
