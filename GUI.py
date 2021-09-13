############################################################################################################
# Google map Convert address to Latitude, Longitude.                                                      
#                                                                                                         
# 輸入input excel or csv檔案檔名以及output excel or csv檔案檔名，欄數、特定欄名，取小數點後幾位 手動輸入單一地址
# 輸出csv or excel原始經緯度，標示出有問題之輸入輸出data                                                     
# Logitude 經度    Latitude 緯度                                                    
############################################################################################################

import tkinter as tk
from tkinter import ttk
import Tools as Tl
import win32clipboard as wcb
import win32con as wc
import threading
import bank_address as ba


Tl.edge()

class Threader(threading.Thread):
    def __init__(self, *args, **kwargs):
        
        threading.Thread.__init__(self, *args, **kwargs)
        self.daemon = True
        self.start()
    def run(self):
        schedule.start(interval=5)
        Execute()
        schedule.stop()

def Execute():
    state_Label.config(text = "executing...")
    
    I_Excel_Path = r'' + I_Excel_Path_Entry.get().strip()
    O_Excel_Path = r'' + O_Excel_Path_Entry.get().strip()
    I_data_type = Tl.distinguish_file_type(I_Excel_Path)
    O_data_type = Tl.distinguish_file_type(O_Excel_Path)
    Cols_ini = Cols_ini_Entry.get().strip()
    Cols_fin = Cols_fin_Entry.get().strip()
    Address = Address_Entry.get().strip()
    Rows = Rows_Entry.get().strip()
    Major = Major_Entry.get().strip()
    Branch = Branch_Entry.get().strip()

    # I_Excel_Path = r"C:\Users\Nick Chang\Desktop\python3\Bank_address\分行地址及經緯度.xlsx"
    # O_Excel_Path = r"C:\Users\Nick Chang\Desktop\python3\Bank_address\test.csv"
    # I_data_type = Tl.distinguish_file_type(I_Excel_Path)
    # O_data_type = Tl.distinguish_file_type(O_Excel_Path)
    # Cols_ini = '1'
    # Cols_fin = '4'
    # Address = "分行地址"
    # Major = Major_Entry.get()
    # Branch = "分行名稱"
    
    sta = ba.Bank_addr(I_Excel_Path, O_Excel_Path, I_data_type, O_data_type, Cols_ini, Cols_fin, Address, Major, Branch, Rows)
    state_Label.config(text = sta)

def Crawler_Single():
    Address = single_address_Entry.get()
    LatLon.set(Tl.web_crawler_and_data_pass_single(Address))

def Copy():
    # 開啟複製貼上板
    wcb.OpenClipboard()
    # 我們之前可能已經Ctrl+C了，這裡是清空目前Ctrl+C複製的內容。但是經過測試，這一步即使沒有也無所謂
    wcb.EmptyClipboard()
    # 將內容寫入複製貼上板,第一個引數win32con.CF_TEXT不用管，我也不知道它是幹什麼的
    # 關鍵第二個引數，就是我們要複製的內容，一定要傳入位元組
    wcb.SetClipboardData(wc.CF_TEXT, LatLon.get().encode("gbk"))
    # 關閉複製貼上板
    wcb.CloseClipboard()

window = tk.Tk()
window.title('Convert GoogleMap into LatLon')
window.geometry('800x500')

# 在圖形介面上設定標籤
I_Excel_Path_Label = tk.Label(window, text='Input Data path:', font=('Arial', 14), width=17, height=2)
I_Excel_Path_Label.place(x = 0, y = 0)
Cols_Label = tk.Label(window, text='Cols:', font=('Arial', 14), width=5, height=2)
Cols_Label.place(x = 650, y = 0)
Address_Label = tk.Label(window, text='Address(Column):', font=('Arial', 14), width=15, height=2)
Address_Label.place(x = 0, y = 40)
Major_Label = tk.Label(window, text='Main name:', font=('Arial', 14), width=8, height=2)
Major_Label.place(x = 300, y = 40)
Branch_Label = tk.Label(window, text='Branch(Column):', font=('Arial', 14), width=13, height=2)
Branch_Label.place(x = 520, y = 40)
O_Excel_Path_Label = tk.Label(window, text='Output Data path:', font=('Arial', 14), width=17, height=2)
O_Excel_Path_Label.place(x = 0, y = 80)
Cols_Label = tk.Label(window, text='Rows:', font=('Arial', 14), width=5, height=2)
Cols_Label.place(x = 650, y = 80)
single_address_Label = tk.Label(window, text='Input single address:', font=('Arial', 14), width=17, height=2)
single_address_Label.place(x = 0, y = 220)
single_address_result_Label = tk.Label(window, text='Convert Result:', font=('Arial', 14), width=17, height=2)
single_address_result_Label.place(x = 0, y = 260)
LatLon = tk.StringVar()
LatLon_Label = tk.Label(window, background='white', textvariable = LatLon, font=('Arial', 14), width=40, height=1)
LatLon_Label.place(x = 210, y = 270)
state_Label = tk.Label(window, background='white', text = '', font=('Arial', 25), width=40, height=1)
state_Label.place(x = 0, y = 320)

# 放置的方法有：1)pack(); 2)place(); 3)grid()
#Entry
I_Excel_Path_Entry = tk.Entry(window, font=('Arial', 14), width = 42)
I_Excel_Path_Entry.place(x = 170, y = 10)
Cols_ini_Entry = tk.Entry(window, font=('Arial', 14), width = 2)
Cols_ini_Entry.place(x = 710, y = 10)
Cols_fin_Entry = tk.Entry(window, font=('Arial', 14), width = 2)
Cols_fin_Entry.place(x = 740, y = 10)
Address_Entry = tk.Entry(window, font=('Arial', 14), width = 10)
Address_Entry.place(x = 170, y = 50)
Major_Entry = tk.Entry(window, font=('Arial', 14), width = 9)
Major_Entry.place(x = 400, y = 50)
Branch_Entry = tk.Entry(window, font=('Arial', 14), width = 9)
Branch_Entry.place(x = 670, y = 50)
O_Excel_Path_Entry = tk.Entry(window, font=('Arial', 14), width = 42)
O_Excel_Path_Entry.place(x = 170, y = 95)
Rows_Entry = tk.Entry(window, font=('Arial', 14), width = 3)
Rows_Entry.place(x = 720, y = 90)
single_address_Entry = tk.Entry(window, font=('Arial', 14), width = 40)
single_address_Entry.place(x = 210, y = 235)

#Progressbar
schedule = ttk.Progressbar(window, length=250)
schedule.place(x = 20 , y = 460)

#Button
Copy_bt = tk.Button(window, text = 'Copy', font=('Arial', 14), width = 5, command = Copy)
Copy_bt.place(x = 700, y = 260)
Convert_bt = tk.Button(window, text = 'Convert Specify', font=('Arial', 14), width = 20, command = Crawler_Single)
Convert_bt.place(x = 300, y = 450)
Execute_bt = tk.Button(window, text = 'Execute', font=('Arial', 14), width = 20, command = lambda: Threader(name='Thread-name'))
Execute_bt.place(x = 550, y = 450)

#主視窗迴圈顯示
window.mainloop()

Tl.close_edge()


#未來的改進----------------------------------

#UI介面更容易理解
#讓使用者選擇是否進行資料比對
#尋找更有效的資料比對的方法
#將input file透過瀏覽檔案的方式開啟
#輸入取幾位小數點，將數字處理完後寫回原檔
#輸入爬蟲列數(不要每次全爬)
#execut控制只能執行一次
#平行處理：開啟兩個以上執行緒以及瀏覽器driver進行同時爬蟲

#爬蟲尚未結束不能將輸出檔案開啟(導致無法輸出檔案，資料重爬)，重複嘗試寫入檔案直到成功寫入並輸出檔案
#pyinstaller -F -w .\GUI.py