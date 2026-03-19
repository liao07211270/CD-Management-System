import tkinter as tk
from tkinter import ttk, messagebox
import os
import re
from pathlib import Path
import threading
import psutil
import time
import subprocess
import sys
import locale

# 隱藏控制台窗口（僅在Windows打包環境中有效）
def hide_console():
    try:
        if os.name == 'nt' and getattr(sys, 'frozen', False):
            import ctypes
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            console_window = kernel32.GetConsoleWindow()
            if console_window != 0:
                user32.ShowWindow(console_window, 0)  # SW_HIDE = 0
    except:
        pass

def setup_encoding():
    """設置正確的編碼環境"""
    try:
        # 設置環境變量
        os.environ['PYTHONIOENCODING'] = 'utf-8'
        os.environ['PYTHONLEGACYWINDOWSSTDIO'] = '1'
        
        # Windows 特別處理
        if os.name == 'nt':
            # 設置控制台編碼為 UTF-8
            try:
                import subprocess
                subprocess.run(['chcp', '65001'], shell=True, capture_output=True, timeout=5)
            except:
                pass
            
            # 設置系統編碼
            try:
                locale.setlocale(locale.LC_ALL, 'Chinese_Taiwan.UTF-8')
            except:
                try:
                    locale.setlocale(locale.LC_ALL, 'zh_TW.UTF-8')
                except:
                    try:
                        locale.setlocale(locale.LC_ALL, 'C.UTF-8')
                    except:
                        pass
        
        # 重新配置標準輸出/輸入編碼
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
        if hasattr(sys.stdin, 'reconfigure'):
            sys.stdin.reconfigure(encoding='utf-8')
            
    except Exception as e:
        print(f"編碼設置警告: {e}")

# 在 FileSearchApp 類定義之前調用
setup_encoding()

# 調用隱藏函數
hide_console()

class FileSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("光碟片資料夾搜尋系統")
        self.root.geometry("1000x950")
        self.root.configure(bg='#f0f0f0')
        self.item_paths = {}
        
        # 搜尋控制
        self.search_thread = None
        self.stop_search = False

        self.searched_folders = set()  # 用於追蹤已搜索的資料夾
        
        # 新增：搜尋路徑設定
        self.search_mode = tk.StringVar(value="custom")  # "custom" 或 "full"
        self.custom_search_path = tk.StringVar(value="D:\\光碟檢測及備份")
        
        # 固定必需檔案名稱模式（不可更改，不顯示在界面中）
        self.required_file_patterns = [
            "光碟檢測截圖",
            "光碟片內容"
        ]
        
        self.year_search_data = []  # 儲存年份搜尋結果
        self.current_page = 0  # 當前頁碼
        self.items_per_page = 2  # 每頁顯示2筆資料
        self.selected_files = set()  # 儲存被選中的檔案索引
        self.year_result_frame = None  # 年份搜尋結果框架
        
        
        # 創建主框架
        self.setup_ui()

    def find_main_py(self):
        """自動尋找 main.py 檔案的位置"""
        # 可能的搜尋路徑
        search_paths = [
            # 指定的已知路徑
            r"C:\Users\a157017\Desktop\vscode\cd_integrator",
            # 當前目錄
            os.getcwd(),
            # 相對路徑
            os.path.join(os.getcwd(), "cd_integrator"),
            # 上層目錄
            os.path.dirname(os.getcwd()),
            # 桌面常見路徑
            os.path.join(os.path.expanduser("~"), "Desktop", "vscode", "cd_integrator"),
            os.path.join(os.path.expanduser("~"), "Desktop", "cd_integrator"),
        ]
        
        for path in search_paths:
            main_py_path = os.path.join(path, "main.py")
            print(f"檢查路徑: {main_py_path}")
            if os.path.exists(main_py_path):
                print(f"找到 main.py: {main_py_path}")
                return main_py_path
        
        return None

    def close_current_app(self):
        """關閉當前應用程式"""
        try:
            print("開始關閉當前程式...")
            
            # 停止所有正在運行的線程
            if hasattr(self, 'search_thread') and self.search_thread and self.search_thread.is_alive():
                print("停止搜尋線程...")
                self.stop_search = True
                # 等待線程結束，但不要等太久
                self.search_thread.join(timeout=1)
            
            # 停止進度條
            if hasattr(self, 'progress'):
                try:
                    self.progress.stop()
                except:
                    pass
            
            # 銷毀主視窗
            if hasattr(self, 'root') and self.root:
                print("銷毀主視窗...")
                self.root.quit()
                self.root.destroy()
                print("程式已成功關閉")
                
        except Exception as e:
            print(f"關閉程式時發生錯誤: {e}")
            # 強制退出
            try:
                if hasattr(self, 'root'):
                    self.root.destroy()
            except:
                pass
            # 最後手段：強制退出程式
            import sys
            sys.exit(0)

    def jump_to_main(self):
        """切換到 main.py 程式（關閉當前程式）"""
        try:
            import subprocess
            import sys
            import os
            from tkinter import filedialog, messagebox
            
            # 取得目標程式路徑
            target_program = None
            
            # 檢查是否為打包後的EXE環境
            if getattr(sys, 'frozen', False):
                # 如果是打包後的EXE，尋找同目錄下的整合工具EXE
                current_dir = os.path.dirname(sys.executable)
                possible_names = [
                    "光碟內容整合工具.exe",
                    "main.exe", 
                    "光碟資料整合程式.exe"
                ]
                
                for name in possible_names:
                    exe_path = os.path.join(current_dir, name)
                    if os.path.exists(exe_path):
                        target_program = exe_path
                        break
            else:
                # 開發環境，尋找 main.py
                target_program = self.find_main_py()
            
            # 如果找不到目標程式
            if not target_program:
                from tkinter import filedialog
                result = messagebox.askyesno(
                    "找不到檔案",
                    "無法自動找到目標程式。\n\n是否要手動選擇檔案位置？"
                )
                if result:
                    if getattr(sys, 'frozen', False):
                        # EXE環境，選擇EXE檔案
                        target_program = filedialog.askopenfilename(
                            title="請選擇光碟內容整合工具",
                            filetypes=[("執行檔", "*.exe"), ("所有檔案", "*.*")],
                            initialdir=os.path.dirname(sys.executable)
                        )
                    else:
                        # 開發環境，選擇Python檔案
                        target_program = filedialog.askopenfilename(
                            title="請選擇 main.py 檔案",
                            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
                            initialdir=os.path.expanduser("~")
                        )
                    
                    if not target_program:
                        return
                else:
                    return
            
            # 確認對話框
            result = messagebox.askyesno(
                "切換程式", 
                "即將切換到「光碟資料整合」程式。\n當前程式將會關閉，搜尋結果將不會保留。\n\n是否繼續？",
                icon='question'
            )
            
            if result:
                try:
                    print(f"準備啟動: {target_program}")
                    
                    # 設置完整的環境變量以解決編碼問題
                    env = os.environ.copy()
                    env['PYTHONIOENCODING'] = 'utf-8'
                    env['PYTHONLEGACYWINDOWSSTDIO'] = '1'
                    env['LANG'] = 'zh_TW.UTF-8'
                    env['LC_ALL'] = 'zh_TW.UTF-8'
                    
                    # Windows 特定設置
                    if os.name == 'nt':
                        env['PYTHONUTF8'] = '1'
                        # 確保路徑使用正確的編碼
                        target_program = os.path.normpath(target_program)
                    
                    # 創建啟動參數
                    if target_program.endswith('.exe'):
                        # 啟動EXE檔案
                        start_args = [target_program]
                    else:
                        # 啟動Python檔案
                        start_args = [sys.executable, '-X', 'utf8', target_program]
                    
                    print(f"啟動參數: {start_args}")
                    print(f"工作目錄: {os.path.dirname(target_program) if not target_program.endswith('.exe') else os.path.dirname(target_program)}")
                    
                    # 啟動新程式
                    if os.name == 'nt':  # Windows
                        process = subprocess.Popen(
                            start_args,
                            cwd=os.path.dirname(target_program) if target_program else None,
                            env=env,
                            creationflags=subprocess.CREATE_NO_WINDOW,
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            stdin=subprocess.DEVNULL
                        )
                    else:  # Linux/Mac
                        process = subprocess.Popen(
                            start_args,
                            cwd=os.path.dirname(target_program) if target_program else None,
                            env=env,
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            stdin=subprocess.DEVNULL
                        )
                    
                    print(f"新程式已啟動，PID: {process.pid}")
                    
                    # 確認新程式啟動成功後再關閉當前程式
                    self.root.after(500, self.close_current_app)
                    
                except Exception as e:
                    print(f"啟動新程式時發生錯誤: {e}")
                    import traceback
                    traceback.print_exc()
                    messagebox.showerror("錯誤", f"切換程式時發生錯誤：\n{str(e)}")
                    
        except Exception as e:
            print(f"切換程式失敗: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"程式切換失敗：\n{str(e)}")
            
    def setup_ui(self):
        # 標題框架 (修改此部分)
        title_frame = tk.Frame(self.root, bg='#f8f9fa')
        title_frame.pack(pady=10, fill='x')
        
        # 創建一個主標題容器，使用grid佈局
        title_container = tk.Frame(title_frame, bg='#f8f9fa')
        title_container.pack(fill='x', padx=20)
        
        # 標題 - 放在左側
        title_label = tk.Label(title_container, text="光碟片資料夾搜尋系統", 
                    font=("Microsoft YaHei", 20, "bold"), 
                    bg='#f8f9fa', fg='#1a1a1a')
        title_label.pack(side='left', pady=10)
        
        # 按鈕框架 - 放在右側
        button_frame = tk.Frame(title_container, bg='#f8f9fa')
        button_frame.pack(side='right', pady=10)
        
        # 指定年份查閱按鈕（新增這個）
        year_search_btn = tk.Button(button_frame, text="📅 指定年份查閱", 
                                    command=self.open_year_search_window,
                                    bg='#007bff', fg='white', 
                                    font=("Microsoft YaHei", 9, "bold"),
                                    width=12, height=1,
                                    cursor='hand2',
                                    relief='flat',
                                    bd=1)
        year_search_btn.pack(side='left', padx=(0, 5))
        
        # 路徑設定按鈕
        path_setting_btn = tk.Button(button_frame, text="🔧 路徑設定", 
                                    command=self.open_path_setting_dialog,
                                    bg='#6c757d', fg='white', 
                                    font=("Microsoft YaHei", 9, "bold"),
                                    width=10, height=1,
                                    cursor='hand2',
                                    relief='flat',
                                    bd=1)
        path_setting_btn.pack(side='left', padx=(0, 5))
        
        # 跳轉按鈕
        jump_btn = tk.Button(button_frame, text="🔄 切換", 
                            command=self.jump_to_main,
                            bg='#e67e22', fg='white', 
                            font=("Microsoft YaHei", 9, "bold"),
                            width=8, height=1,
                            cursor='hand2',
                            relief='flat',
                            bd=1)
        jump_btn.pack(side='left')
        
        # 滑鼠懸停效果
        def on_path_enter(event):
            path_setting_btn.config(bg='#5a6268')
        
        def on_path_leave(event):
            path_setting_btn.config(bg='#6c757d')
        
        def on_jump_enter(event):
            jump_btn.config(bg='#d35400')
        
        def on_jump_leave(event):
            jump_btn.config(bg='#e67e22')
        
        path_setting_btn.bind("<Enter>", on_path_enter)
        path_setting_btn.bind("<Leave>", on_path_leave)
        jump_btn.bind("<Enter>", on_jump_enter)
        jump_btn.bind("<Leave>", on_jump_leave)
        
        # 查詢條件框架
        search_frame = tk.LabelFrame(self.root, text="查詢條件", 
                           font=("Microsoft YaHei", 14, "bold"),
                           bg='#f8f9fa', fg='#1a1a1a', padx=10, pady=8)
        search_frame.pack(fill='x', padx=20, pady=10)
        
        # 創建查詢欄位變數
        self.search_vars = {}
        
        # 創建網格布局
            # 第一行：年度、分類號
        tk.Label(search_frame, text="年度:", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=0, column=0, sticky='w', padx=5, pady=2)
        
        year_var = tk.StringVar()
        self.search_vars['year'] = year_var
        year_entry = tk.Entry(search_frame, textvariable=year_var, 
            font=("Microsoft YaHei", 11), width=15)
        year_entry.grid(row=1, column=0, padx=5, pady=2, sticky='ew')
        year_entry.bind('<Return>', self.on_enter_pressed)  # 新增這行
        
        tk.Label(search_frame, text="分類號:", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=0, column=1, sticky='w', padx=5, pady=2)
        
        category_var = tk.StringVar()
        self.search_vars['category'] = category_var
        category_entry = tk.Entry(search_frame, textvariable=category_var, 
            font=("Microsoft YaHei", 11), width=15)
        category_entry.grid(row=1, column=1, padx=5, pady=2, sticky='ew')
        category_entry.bind('<Return>', self.on_enter_pressed)  # 新增這行
        
        # 第二行：案、卷、目
        tk.Label(search_frame, text="案:", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=2, column=0, sticky='w', padx=5, pady=2)
        
        case_var = tk.StringVar()
        self.search_vars['case'] = case_var
        case_entry = tk.Entry(search_frame, textvariable=case_var, 
            font=("Microsoft YaHei", 11), width=15)
        case_entry.grid(row=3, column=0, padx=5, pady=2, sticky='ew')
        case_entry.bind('<Return>', self.on_enter_pressed)  # 新增這行
        
        tk.Label(search_frame, text="卷:", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=2, column=1, sticky='w', padx=5, pady=2)
        
        volume_var = tk.StringVar()
        self.search_vars['volume'] = volume_var
        volume_entry = tk.Entry(search_frame, textvariable=volume_var, 
            font=("Microsoft YaHei", 11), width=15)
        volume_entry.grid(row=3, column=1, padx=5, pady=2, sticky='ew')
        volume_entry.bind('<Return>', self.on_enter_pressed)  # 新增這行
        
        tk.Label(search_frame, text="目:", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=2, column=2, sticky='w', padx=5, pady=2)
        
        item_var = tk.StringVar()
        self.search_vars['item'] = item_var
        item_entry = tk.Entry(search_frame, textvariable=item_var, 
            font=("Microsoft YaHei", 11), width=15)
        item_entry.grid(row=3, column=2, padx=5, pady=2, sticky='ew')
        item_entry.bind('<Return>', self.on_enter_pressed)  # 新增這行
        
        # 第三行：片數資訊
        tk.Label(search_frame, text="片數(總片數):", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=4, column=0, sticky='w', padx=5, pady=2)
        
        sheet_total_var = tk.StringVar()
        self.search_vars['sheet_total'] = sheet_total_var
        sheet_total_entry = tk.Entry(search_frame, textvariable=sheet_total_var, 
            font=("Microsoft YaHei", 11), width=15)
        sheet_total_entry.grid(row=5, column=0, padx=5, pady=2, sticky='ew')
        sheet_total_entry.bind('<Return>', self.on_enter_pressed)  # 新增這行
        
        tk.Label(search_frame, text="片數(編號):", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=4, column=1, sticky='w', padx=5, pady=2)
        
        sheet_num_var = tk.StringVar()
        self.search_vars['sheet_num'] = sheet_num_var
        sheet_num_entry = tk.Entry(search_frame, textvariable=sheet_num_var, 
            font=("Microsoft YaHei", 11), width=15)
        sheet_num_entry.grid(row=5, column=1, padx=5, pady=2, sticky='ew')
        sheet_num_entry.bind('<Return>', self.on_enter_pressed)  # 新增這行
        
        # 設置網格權重
        for i in range(3):
            search_frame.grid_columnconfigure(i, weight=1)
        
        # 提示文字
        hint_label = tk.Label(search_frame, text="提示：輸入關鍵字進行精確搜尋（例如：年度輸入113、分類號輸入030199等）", 
                     font=("Microsoft YaHei", 10, "bold"), bg='#f8f9fa', fg='#495057')
        hint_label.grid(row=6, column=0, columnspan=3, sticky='w', padx=5, pady=5)
        
        # 查詢按鈕框架
        button_frame = tk.Frame(search_frame, bg='#f0f0f0')
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        search_btn = tk.Button(button_frame, text="🔍 開始搜索", 
                              command=self.search_files,
                              bg='#27ae60', fg='white', 
                              font=("Microsoft YaHei", 13, "bold"),
                              width=10, height=1)
        search_btn.pack(side='left', padx=5)
        
        clear_btn = tk.Button(button_frame, text=" 清除 ", 
                             command=self.clear_fields,
                             bg='#e74c3c', fg='white', 
                             font=("Microsoft YaHei", 13, "bold"),
                             width=4, height=1)
        clear_btn.pack(side='left', padx=5)
        
        # 搜索進度框架
        progress_frame = tk.Frame(self.root, bg='#f0f0f0')
        progress_frame.pack(fill='x', padx=20, pady=5)
        
        tk.Label(progress_frame, text="搜索進度:", font=("Microsoft YaHei", 12, "bold"), 
        bg='#f8f9fa', fg='#2c3e50').pack(side='left')
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.pack(side='left', padx=5, fill='x', expand=True)
        
        self.stop_btn = tk.Button(progress_frame, text="⏹️ 停止搜索", 
                         command=self.stop_search_process,
                         bg="#F37676", fg="#FFFFFF", 
                         font=("Microsoft JhengHei", 12, "bold"))
        self.stop_btn.pack(side='right', padx=5)
        
        # 結果顯示框架
        result_frame = tk.LabelFrame(self.root, text="搜索結果", 
                                   font=("Microsoft YaHei", 13, "bold"),
                                   bg='#f0f0f0', fg='#2c3e50', padx=5, pady=5)
        result_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 修改表格欄位，新增編輯和刪除欄位
        columns = ('編輯/刪除', '年度', '分類號', '案', '卷', '目', '片數/總片數', '檢測結果', '光碟內容', '完整度','路徑')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=15)

        # 設定欄位寬度
        column_widths = {
            '編輯/刪除': 100,
            '年度': 80,
            '分類號': 100,
            '案': 60,
            '卷': 60,
            '目': 60,
            '片數/總片數': 120,
            '檢測結果': 120,
            '光碟內容': 120,
            '完整度': 100,
            '路徑': 0 
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            if col == '編輯/刪除':
                self.tree.column(col, width=column_widths.get(col, 100), anchor='center')
            elif col == '路徑':
                self.tree.column(col, width=0, minwidth=0, stretch=False)  # 完全隱藏路徑欄位
            else:
                self.tree.column(col, width=column_widths.get(col, 100))
                
        # 滾動條
        scrollbar_v = ttk.Scrollbar(result_frame, orient='vertical', command=self.tree.yview)
        scrollbar_h = ttk.Scrollbar(result_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        
        # 佈局
        self.tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_v.grid(row=0, column=1, sticky='ns')
        scrollbar_h.grid(row=1, column=0, sticky='ew')
        
        # 設置網格權重
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)
        
        # 狀態列
        self.status_var = tk.StringVar()
        self.status_var.set("設定查詢條件後點擊「開始搜索」來尋找符合條件的光碟片資料夾")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                    relief='sunken', anchor='w', 
                    font=("Microsoft YaHei", 11, "bold"), bg='#e9ecef', fg='#495057')
        status_bar.pack(side='bottom', fill='x')
        
        # 右鍵選單
        self.setup_context_menu()
        
        # 雙擊事件 - 開啟資料夾
        self.tree.bind('<Double-1>', self.open_folder_location)
        # 單擊事件 - 處理編輯和刪除按鈕
        self.tree.bind('<Button-1>', self.handle_tree_click)

        # 設置樹狀視圖字體
        style = ttk.Style()
        style.configure("Treeview", font=("Microsoft YaHei", 11))
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 12, "bold"))

    def open_path_setting_dialog(self):
        """開啟路徑設定彈窗"""
        # 創建彈窗
        dialog = tk.Toplevel(self.root)
        dialog.title("路徑設定")
        dialog.geometry("650x420")
        dialog.configure(bg='#f8f9fa')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # 居中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (650 // 2)
        y = (dialog.winfo_screenheight() // 2) - (420 // 2)
        dialog.geometry(f"650x420+{x}+{y}")
        
        # 標題
        title_label = tk.Label(dialog, text="搜尋路徑設定", 
                            font=("Microsoft YaHei", 16, "bold"),
                            bg='#f8f9fa', fg='#1a1a1a')
        title_label.pack(pady=(15, 10))
        
        # 主要內容框架 - 使用滾動區域
        canvas = tk.Canvas(dialog, bg='#f8f9fa')
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#f8f9fa')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 打包canvas和scrollbar
        canvas.pack(side="left", fill="both", expand=True, padx=20)
        scrollbar.pack(side="right", fill="y")
        
        # 主要內容框架（原本的main_frame改為scrollable_frame的子框架）
        main_frame = tk.Frame(scrollable_frame, bg='#f8f9fa')
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 臨時變數（用於取消時恢復）
        temp_search_mode = tk.StringVar(value=self.search_mode.get())
        temp_custom_path = tk.StringVar(value=self.custom_search_path.get())
        
        # 搜尋模式框架
        mode_frame = tk.LabelFrame(main_frame, text="搜尋模式", 
                                font=("Microsoft YaHei", 12, "bold"),
                                bg='#f8f9fa', fg='#2c3e50', padx=15, pady=10)
        mode_frame.pack(fill='x', pady=(0, 15))
        
        # 指定路徑選項
        custom_frame = tk.Frame(mode_frame, bg='#f8f9fa')
        custom_frame.pack(fill='x', pady=5)
        
        # 狀態顯示框架
        status_frame = tk.LabelFrame(main_frame, text="當前設定狀態", 
                                    font=("Microsoft YaHei", 12, "bold"),
                                    bg='#f8f9fa', fg='#2c3e50', padx=15, pady=10)
        status_frame.pack(fill='x', pady=(0, 15))
        
        status_var = tk.StringVar()
        status_label = tk.Label(status_frame, textvariable=status_var,
                            font=("Microsoft YaHei", 10), bg='#f8f9fa', 
                            fg='#495057', wraplength=550, justify='left')
        status_label.pack(anchor='w', pady=5)
        
        # 定義內嵌函數
        def update_path_entry_state():
            if temp_search_mode.get() == "custom":
                path_entry.config(state='normal')
                browse_btn.config(state='normal')
            else:
                path_entry.config(state='disabled')
                browse_btn.config(state='disabled')
            update_status()

        def update_status():
            """更新狀態顯示"""
            if temp_search_mode.get() == "full":
                status_var.set("🌐 模式：全機搜尋\n📁 將搜尋整個電腦中的所有可用磁碟機和常見位置")
            else:
                path = temp_custom_path.get()
                if os.path.exists(path):
                    status_var.set(f"🎯 模式：指定路徑搜尋\n📁 路徑：{path}\n✅ 狀態：路徑存在，可以開始搜尋")
                else:
                    status_var.set(f"🎯 模式：指定路徑搜尋\n📁 路徑：{path}\n⚠️ 狀態：路徑不存在，搜尋時會提示是否改用全機搜尋")
        
        def browse_path():
            """瀏覽選擇路徑"""
            from tkinter import filedialog
            
            initial_dir = temp_custom_path.get()
            if not os.path.exists(initial_dir):
                initial_dir = os.path.expanduser("~")
            
            selected_path = filedialog.askdirectory(
                title="選擇搜尋路徑",
                initialdir=initial_dir
            )
            
            if selected_path:
                temp_custom_path.set(selected_path)
                update_status()
        
        def save_settings():
            """儲存設定"""
            # 更新實際變數
            self.search_mode.set(temp_search_mode.get())
            self.custom_search_path.set(temp_custom_path.get())
            
            # 顯示確認訊息
            if temp_search_mode.get() == "full":
                messagebox.showinfo("設定已儲存", "已設定為全機搜尋模式")
            else:
                messagebox.showinfo("設定已儲存", f"已設定指定路徑：\n{temp_custom_path.get()}")
            
            dialog.destroy()
        
        def cancel_settings():
            """取消設定"""
            dialog.destroy()
        
        # 創建UI元件（現在內嵌函數已經定義）
        custom_radio = tk.Radiobutton(custom_frame, text="指定路徑搜尋", 
                                    variable=temp_search_mode, value="custom",
                                    command=lambda: update_path_entry_state(),
                                    font=("Microsoft YaHei", 11, "bold"), 
                                    bg='#f8f9fa', fg='#2c3e50')
        custom_radio.pack(anchor='w')
        
        # 路徑輸入框架
        path_input_frame = tk.Frame(mode_frame, bg='#f8f9fa')
        path_input_frame.pack(fill='x', pady=(5, 10))
        
        tk.Label(path_input_frame, text="搜尋路徑：", 
                font=("Microsoft YaHei", 10), bg='#f8f9fa', fg='#495057').pack(anchor='w')
        
        path_entry_frame = tk.Frame(path_input_frame, bg='#f8f9fa')
        path_entry_frame.pack(fill='x', pady=(5, 0))
        
        path_entry = tk.Entry(path_entry_frame, textvariable=temp_custom_path,
                            font=("Microsoft YaHei", 10))
        path_entry.pack(side='left', fill='x', expand=True)
        
        browse_btn = tk.Button(path_entry_frame, text="瀏覽", 
                            command=lambda: browse_path(),
                            bg='#3498db', fg='white', 
                            font=("Microsoft YaHei", 9, "bold"),
                            width=10)
        browse_btn.pack(side='right', padx=(10, 0))
        
        # 全機搜尋選項
        full_radio = tk.Radiobutton(mode_frame, text="全機搜尋（搜尋所有可用磁碟機）", 
                                variable=temp_search_mode, value="full",
                                command=lambda: update_path_entry_state(),
                                font=("Microsoft YaHei", 11, "bold"), 
                                bg='#f8f9fa', fg='#2c3e50')
        full_radio.pack(anchor='w', pady=(10, 5))
        
        # 初始化狀態
        update_path_entry_state()
        
        # 按鈕框架
        button_frame = tk.Frame(dialog, bg='#f8f9fa')
        button_frame.pack(side='bottom', fill='x', padx=15, pady=15)
        
        # 取消按鈕
        cancel_btn = tk.Button(button_frame, text="❌ 取消", 
                            command=cancel_settings,
                            bg='#e74c3c', fg='white', 
                            font=("Microsoft YaHei", 10, "bold"),
                            width=7, height=1)
        cancel_btn.pack(side='right', padx=(10, 0))
        
        # 儲存按鈕
        save_btn = tk.Button(button_frame, text="💾 儲存設定", 
                            command=save_settings,
                            bg='#27ae60', fg='white', 
                            font=("Microsoft YaHei", 10, "bold"),
                            width=12, height=1)
        save_btn.pack(side='right', padx=(0, 10))
        

    def create_styled_jump_button(self, parent_frame):
        """建立具有懸停效果的跳轉按鈕"""
        jump_btn = tk.Button(parent_frame, text="🔄 切換", 
                            command=self.jump_to_main,
                            bg="#dfa811", fg='white', 
                            font=("Microsoft YaHei", 12, "bold"),
                            width=16, height=1,
                            cursor='hand2',
                            relief='flat',
                            bd=2)
        
        # 滑鼠懸停效果
        def on_enter(event):
            jump_btn.config(bg='#d35400')
        
        def on_leave(event):
            jump_btn.config(bg='#e67e22')
        
        jump_btn.bind("<Enter>", on_enter)
        jump_btn.bind("<Leave>", on_leave)
        
        return jump_btn
    

    def handle_tree_click(self, event):
        """處理樹狀視圖的點擊事件"""
        item = self.tree.identify_row(event.y)
        if not item:
            return
        
        column = self.tree.identify_column(event.x)
        column_name = self.tree.heading(column)['text']
        
        if column_name == '編輯/刪除':
            # 獲取點擊位置的相對位置來判斷點擊的是哪個按鈕
            bbox = self.tree.bbox(item, column)
            if bbox:
                click_x = event.x - bbox[0]  # 相對於該欄位的x座標
                column_width = bbox[2]
                
                # 如果點擊位置在左半部分，執行編輯；右半部分執行刪除
                if click_x < column_width / 2:
                    self.edit_folder_name(item)
                else:
                    self.delete_folder(item)

    def edit_folder_name(self, item):
        """編輯資料夾名稱"""
        # 使用 item_paths 字典獲取路徑
        if item not in self.item_paths:
            messagebox.showerror("錯誤", "無法找到資料夾路徑信息")
            return
            
        folder_path = self.item_paths[item]
        original_folder_name = os.path.basename(folder_path)
        
        values = list(self.tree.item(item, 'values'))
        
        # 創建編輯對話框
        edit_window = tk.Toplevel(self.root)
        edit_window.title("編輯資料夾名稱")
        edit_window.geometry("500x500")  # 增加高度以容納新欄位
        edit_window.configure(bg='#f0f0f0')
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        # 居中顯示
        edit_window.update_idletasks()
        x = (edit_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (edit_window.winfo_screenheight() // 2) - (500 // 2)
        edit_window.geometry(f"500x500+{x}+{y}")
        
        # 標題
        title_label = tk.Label(edit_window, text="編輯資料夾資訊", 
                            font=("Microsoft YaHei", 14, "bold"),
                            bg='#f0f0f0', fg='#000000')
        title_label.pack(pady=10)
        
        # 編輯框架
        edit_frame = tk.LabelFrame(edit_window, text="資料夾資訊", 
                                font=("Microsoft YaHei", 12),
                                bg='#f0f0f0', padx=15, pady=15)
        edit_frame.pack(fill='x', padx=20, pady=10)
        
        # 編輯變數
        edit_vars = {}
        fields = [
            ('年度', values[1] if len(values) > 1 else ''),
            ('分類號', values[2] if len(values) > 2 else ''),
            ('案', values[3] if len(values) > 3 else ''),
            ('卷', values[4] if len(values) > 4 else ''),
            ('目', values[5] if len(values) > 5 else ''),
            ('片數(總)', values[6].split('/')[1] if len(values) > 6 and '/' in values[6] else ''),  # 從片數/總片數中提取總片數
            ('片數(編號)', values[6].split('/')[0] if len(values) > 6 and '/' in values[6] else '')   # 從片數/總片數中提取片數
        ]
        
        # 創建編輯欄位 - 修改布局以容納7個欄位
        for i, (field_name, current_value) in enumerate(fields):
            if i < 4:  # 前四個欄位：年度、分類號、案、卷 (2x2布局)
                row = i // 2
                col = i % 2
            else:  # 後三個欄位：目、片數(總)、片數(編號) (1x3布局)
                row = 2
                col = i - 4
            
            tk.Label(edit_frame, text=f"{field_name}:", 
                    font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                    row=row*2, column=col*2, sticky='w', padx=5, pady=5)
            
            var = tk.StringVar(value=current_value)
            edit_vars[field_name] = var
            
            # 調整第三行欄位的寬度，使其更小
            if i >= 4:  # 目、片數(總)、片數(編號)
                width = 12  # 縮小寬度
            else:
                width = 20  # 保持原寬度
                
            entry = tk.Entry(edit_frame, textvariable=var, 
                        font=("Microsoft YaHei", 10), width=width)
            entry.grid(row=row*2+1, column=col*2, padx=5, pady=5, sticky='ew')
        
        # 原始資料夾名稱顯示
        tk.Label(edit_frame, text="原始資料夾名稱:", 
                font=("Microsoft YaHei", 11, "bold"), bg='#f8f9fa', fg='#2c3e50').grid(
                row=8, column=0, columnspan=6, sticky='w', padx=10, pady=(20,5))
        
        original_label = tk.Label(edit_frame, text=original_folder_name, 
                                font=("Microsoft YaHei", 9), bg='#f0f0f0', 
                                fg='#7f8c8d', wraplength=500)
        original_label.grid(row=9, column=0, columnspan=6, sticky='w', padx=10, pady=(5,20))
        
        # 設置網格權重
        for i in range(6):
            edit_frame.grid_columnconfigure(i, weight=1)
        
        # 按鈕框架 - 放在edit_window中而不是edit_frame中
        button_frame = tk.Frame(edit_window, bg='#f0f0f0')
        button_frame.pack(pady=20)
        
        def save_changes():
            try:
                # 獲取新的值
                new_values = {}
                for field_name, var in edit_vars.items():
                    new_values[field_name] = var.get().strip()
                
                # 構建新的資料夾名稱
                parent_dir = os.path.dirname(folder_path)
                old_name_parts = original_folder_name.split('-')
                
                # 如果原本有7個部分，保持格式
                if len(old_name_parts) >= 7:
                    new_name_parts = [
                        new_values['年度'] or old_name_parts[0],
                        new_values['分類號'] or old_name_parts[1],
                        new_values['案'] or old_name_parts[2],
                        new_values['卷'] or old_name_parts[3],
                        new_values['目'] or old_name_parts[4],
                        new_values['片數(總)'] or old_name_parts[5],      # 總片數
                        new_values['片數(編號)'] or old_name_parts[6]      # 編號
                    ]
                    # 保留後續部分（如果有的話）
                    if len(old_name_parts) > 7:
                        new_name_parts.extend(old_name_parts[7:])
                else:
                    # 如果格式不符合預期，只更新可識別的部分
                    new_name_parts = old_name_parts.copy()
                    if len(new_name_parts) > 0 and new_values['年度']:
                        new_name_parts[0] = new_values['年度']
                    if len(new_name_parts) > 1 and new_values['分類號']:
                        new_name_parts[1] = new_values['分類號']
                    if len(new_name_parts) > 2 and new_values['案']:
                        new_name_parts[2] = new_values['案']
                    if len(new_name_parts) > 3 and new_values['卷']:
                        new_name_parts[3] = new_values['卷']
                    if len(new_name_parts) > 4 and new_values['目']:
                        new_name_parts[4] = new_values['目']
                    if len(new_name_parts) > 5 and new_values['片數(總)']:
                        new_name_parts[5] = new_values['片數(總)']
                    if len(new_name_parts) > 6 and new_values['片數(編號)']:
                        new_name_parts[6] = new_values['片數(編號)']
                
                new_folder_name = '-'.join(new_name_parts)
                new_folder_path = os.path.join(parent_dir, new_folder_name)
                
                # 檢查新名稱是否已存在
                if os.path.exists(new_folder_path) and new_folder_path != folder_path:
                    messagebox.showerror("錯誤", "新的資料夾名稱已存在！")
                    return
                
                # 重新命名資料夾
                if new_folder_path != folder_path:
                    os.rename(folder_path, new_folder_path)
                    messagebox.showinfo("成功", f"資料夾已重新命名為:\n{new_folder_name}")
                    
                    # 更新樹狀視圖中的資料
                    updated_values = list(values)
                    if len(updated_values) > 1: updated_values[1] = new_values['年度'] or values[1]
                    if len(updated_values) > 2: updated_values[2] = new_values['分類號'] or values[2]
                    if len(updated_values) > 3: updated_values[3] = new_values['案'] or values[3]
                    if len(updated_values) > 4: updated_values[4] = new_values['卷'] or values[4]
                    if len(updated_values) > 5: updated_values[5] = new_values['目'] or values[5]
                    
                    # 更新片數資訊（片數(編號)/片數(總)）
                    sheet_info = f"{new_values['片數(編號)'] or values[6].split('/')[0] if len(values) > 6 and '/' in values[6] else ''}/{new_values['片數(總)'] or values[6].split('/')[1] if len(values) > 6 and '/' in values[6] else ''}"
                    if len(updated_values) > 6: updated_values[6] = sheet_info
                    
                    self.tree.item(item, values=updated_values)
                    
                    # 更新路徑字典
                    self.item_paths[item] = new_folder_path
                else:
                    messagebox.showinfo("提示", "沒有進行任何更改")
                
                edit_window.destroy()
                
            except Exception as e:
                messagebox.showerror("錯誤", f"重新命名失敗：{str(e)}")
        
        def cancel_edit():
            edit_window.destroy()
        
        # 儲存按鈕
        save_btn = tk.Button(button_frame, text="💾 儲存並關閉", 
                            command=save_changes,
                            bg='#27ae60', fg='white', 
                            font=("Microsoft YaHei", 13, "bold"),
                            width=15, height=2)
        save_btn.pack(side='left', padx=10)
        
        # 取消按鈕
        cancel_btn = tk.Button(button_frame, text="❌ 取消", 
                            command=cancel_edit,
                            bg='#e74c3c', fg='white', 
                            font=("Microsoft YaHei", 13, "bold"),
                            width=10, height=2)
        cancel_btn.pack(side='left', padx=10)

    def delete_folder(self, item):
        """刪除資料夾"""
        # 使用 item_paths 字典獲取路徑
        if item not in self.item_paths:
            messagebox.showerror("錯誤", "無法找到資料夾路徑信息")
            return
            
        folder_path = self.item_paths[item]
        folder_name = os.path.basename(folder_path)
        
        # 確認刪除對話框
        result = messagebox.askyesno(
            "確認刪除", 
            f"您確定要刪除以下資料夾嗎？\n\n資料夾名稱：{folder_name}\n路徑：{folder_path}\n\n⚠️ 此操作無法復原！",
            icon='warning'
        )
        
        if result:
            try:
                import shutil
                if os.path.exists(folder_path):
                    # 刪除資料夾及其所有內容
                    shutil.rmtree(folder_path)
                    
                    # 從樹狀視圖中移除項目
                    self.tree.delete(item)
                    
                    # 從路徑字典中移除
                    if item in self.item_paths:
                        del self.item_paths[item]
                    
                    messagebox.showinfo("成功", f"資料夾已成功刪除：\n{folder_name}")
                else:
                    messagebox.showerror("錯誤", "資料夾不存在或已被刪除！")
                    # 仍然從樹狀視圖中移除
                    self.tree.delete(item)
                    if item in self.item_paths:
                        del self.item_paths[item]
                        
            except Exception as e:
                messagebox.showerror("錯誤", f"刪除資料夾失敗：{str(e)}")

    def setup_context_menu(self):
        """設置右鍵選單"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="📂 開啟資料夾", command=self.open_selected_folder)
        self.context_menu.add_command(label="📋 複製路徑", command=self.copy_selected_path)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="✏️ 編輯名稱", command=self.edit_selected_folder)
        self.context_menu.add_command(label="🗑️ 刪除資料夾", command=self.delete_selected_folder)
        
        self.tree.bind('<Button-3>', self.show_context_menu)

    def show_context_menu(self, event):
        """顯示右鍵選單"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()

    def copy_selected_path(self):
        """複製選中項目的路徑到剪貼簿"""
        selection = self.tree.selection()
        if not selection:
            return
        
        values = self.tree.item(selection[0], 'values')
        
        values = self.tree.item(selection[0], 'values')
        if len(values) >= 11:
            folder_path = values[10]  # 路徑欄位
            self.root.clipboard_clear()
            self.root.clipboard_append(folder_path)
            messagebox.showinfo("已複製", f"路徑已複製到剪貼簿:\n{folder_path}")

    def edit_selected_folder(self):
        """編輯選中的資料夾"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "請先選擇一個資料夾")
            return
        
        self.edit_folder_name(selection[0])

    def delete_selected_folder(self):
        """刪除選中的資料夾"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "請先選擇一個資料夾")
            return
        
        self.delete_folder(selection[0])

    def open_selected_folder(self):
        """開啟選中的資料夾"""
        selection = self.tree.selection()
        if not selection:
            return
        
        values = self.tree.item(selection[0], 'values')
        if len(values) >= 11:
            folder_path = values[10]  # 路徑欄位
            self.open_folder_by_path(folder_path)

    def open_folder_by_path(self, folder_path):
        """根據路徑開啟資料夾"""
        if folder_path and os.path.exists(folder_path):
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(folder_path)
                elif os.name == 'posix':  # macOS and Linux
                    if sys.platform == 'darwin':  # macOS
                        subprocess.run(['open', folder_path])
                    else:  # Linux
                        subprocess.run(['xdg-open', folder_path])
            except Exception as e:
                messagebox.showerror("錯誤", f"無法開啟資料夾：{e}")
        else:
            messagebox.showerror("錯誤", "資料夾不存在或無法存取！")

    def get_search_paths(self):
        """根據設定取得搜尋路徑 - 修復重複路徑問題"""
        search_paths = []
        
        if self.search_mode.get() == "custom":
            # 自訂路徑模式
            custom_path = self.custom_search_path.get().strip()
            if custom_path and os.path.exists(custom_path):
                search_paths.append(os.path.abspath(custom_path))  # 使用絕對路徑
                print(f"✓ 使用自訂搜尋路徑: {custom_path}")
            else:
                print(f"✗ 自訂路徑不存在: {custom_path}")
                from tkinter import messagebox
                result = messagebox.askyesno(
                    "路徑不存在",
                    f"指定的搜尋路徑不存在：\n{custom_path}\n\n是否改用全機搜尋？"
                )
                if result:
                    self.search_mode.set("full")
                    return self.get_search_paths()
                else:
                    return []
        else:
            # 全機搜尋模式
            print("🌐 使用全機搜尋模式")
            
            # 收集所有可能的路徑
            all_potential_paths = []
            
            # 方法1: 使用者目錄
            user_profile = os.path.expanduser("~")
            desktop_paths = [
                os.path.join(user_profile, "Desktop"),
                os.path.join(user_profile, "桌面"),
                os.path.join(user_profile, "Documents"),
                os.path.join(user_profile, "文件"),
                os.path.join(user_profile, "Downloads"),
                os.path.join(user_profile, "下載"),
            ]
            
            # 方法2: 公用桌面
            public_paths = [
                "C:\\Users\\Public\\Desktop",
                "C:\\Users\\Public\\桌面",
                "C:\\Users\\Public\\Documents",
                "C:\\Users\\Public\\文件",
            ]
            
            # 方法3: 所有用戶的桌面路徑
            try:
                users_dir = "C:\\Users"
                if os.path.exists(users_dir):
                    for user_folder in os.listdir(users_dir):
                        user_path = os.path.join(users_dir, user_folder)
                        if os.path.isdir(user_path) and user_folder not in ['Public', 'Default', 'All Users']:
                            desktop_paths.extend([
                                os.path.join(user_path, "Desktop"),
                                os.path.join(user_path, "桌面"),
                                os.path.join(user_path, "Documents"),
                                os.path.join(user_path, "文件")
                            ])
            except Exception as e:
                print(f"無法列舉用戶目錄: {e}")
            
            # 方法4: 嘗試從註冊表獲取桌面路徑
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
                    desktop_path = winreg.QueryValueEx(key, "Desktop")[0]
                    if desktop_path not in desktop_paths:
                        desktop_paths.insert(0, desktop_path)
            except:
                pass
            
            # 方法5: 常見的工作目錄路徑
            common_paths = [
                "D:\\",
                "E:\\", 
                "F:\\",
                os.path.join("C:\\", "光碟資料"),
                os.path.join("D:\\", "光碟資料"),
                os.path.join("E:\\", "光碟資料"),
                "D:\\光碟檢測及備份",
            ]
            
            all_potential_paths = desktop_paths + public_paths + common_paths
            
            # 去除重複路徑並轉換為絕對路徑
            unique_paths = set()
            for path in all_potential_paths:
                if os.path.exists(path):
                    abs_path = os.path.abspath(path)
                    unique_paths.add(abs_path)
            
            search_paths = list(unique_paths)
        
        # 輸出最終搜尋路徑清單
        print(f"=== 最終搜尋路徑清單 ===")
        for i, path in enumerate(search_paths, 1):
            print(f"{i}. {path}")
        
        return search_paths

        
    def stop_search_process(self):
        """停止搜索過程"""
        self.stop_search = True
        self.status_var.set("正在停止搜索...")

    def on_enter_pressed(self, event):
        """當按下 Enter 鍵時觸發搜索"""
        self.search_files()    
        
    def clear_fields(self):
        """清除查詢欄位 - 修復版本"""
        for var in self.search_vars.values():
            var.set("")
        self.status_var.set("已清除查詢條件")
        
        # 清除結果
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 重置已搜索資料夾集合
        if hasattr(self, 'searched_folders'):
            self.searched_folders.clear()
    
    def parse_folder_name(self, folder_name):
        """解析資料夾名稱，提取各個欄位資訊"""
        # 預期格式: 年度-分類號-案-卷-目-總片數-編號
        parts = folder_name.split('-')
        
        parsed_info = {
            'year': '',
            'category': '',
            'case': '',
            'volume': '',
            'item': '',
            'sheet_total': '',
            'sheet_num': ''
        }
        
        # 根據分隔符數量來解析
        if len(parts) >= 7:
            parsed_info['year'] = parts[0]
            parsed_info['category'] = parts[1]
            parsed_info['case'] = parts[2]
            parsed_info['volume'] = parts[3]
            parsed_info['item'] = parts[4]
            parsed_info['sheet_total'] = parts[5]
            parsed_info['sheet_num'] = parts[6]
        
        return parsed_info
    
    def folder_matches_criteria(self, folder_name, criteria):
        """檢查資料夾名稱是否符合搜索條件"""
        folder_name_lower = folder_name.lower()
        
        print(f"檢查資料夾: {folder_name}")
        print(f"搜索條件: {criteria}")
        
        # 如果沒有任何搜索條件，返回True
        if not any(criteria.values()):
            print("  沒有搜索條件，符合")
            return True
        
        # 解析資料夾名稱
        parsed_info = self.parse_folder_name(folder_name)
        
        # 檢查每個有值的搜索條件
        for key, value in criteria.items():
            if value:  # 只檢查有輸入值的欄位
                search_value = value.lower()
                
                # 對應的解析欄位
                if key in parsed_info:
                    parsed_value = parsed_info[key].lower()
                    if search_value not in parsed_value:
                        print(f"  不符合條件: '{value}' 不匹配 '{parsed_info[key]}'")
                        return False
                else:
                    # 如果是舊的搜索方式（在整個資料夾名稱中搜索）
                    if search_value not in folder_name_lower:
                        print(f"  不符合條件: '{value}' 不在資料夾名稱中")
                        return False
        
        print("  符合所有輸入的條件!")
        return True
    
    def check_required_files(self, folder_path, required_patterns):
        """檢查資料夾是否包含所有必需的檔案模式"""
        print(f"檢查必需檔案: {folder_path}")
        print(f"必需模式: {required_patterns}")
        
        try:
            if not os.path.exists(folder_path):
                print("  資料夾不存在")
                return False, {}, required_patterns.copy()
                
            items = os.listdir(folder_path)
            subdirs = [item for item in items if os.path.isdir(os.path.join(folder_path, item))]
            print(f"  資料夾內子目錄: {subdirs}")
            
            matched_dirs = {}
            missing_patterns = []
            
            # 檢查每個必需模式
            for pattern in required_patterns:
                pattern_found = False
                for subdir_name in subdirs:
                    if pattern in subdir_name:
                        matched_dirs[pattern] = subdir_name
                        pattern_found = True
                        print(f"  ✓ 找到模式 '{pattern}' 的資料夾: {subdir_name}")
                        break
                
                if not pattern_found:
                    missing_patterns.append(pattern)
                    print(f"  ✗ 缺少模式 '{pattern}' 的資料夾")
            
            is_complete = len(missing_patterns) == 0
            print(f"  檢查結果: {'完整' if is_complete else '不完整'}")
            print(f"  匹配資料夾: {matched_dirs}")
            print(f"  缺失模式: {missing_patterns}")
            
            return is_complete, matched_dirs, missing_patterns
            
        except Exception as e:
            print(f"檢查檔案時發生錯誤: {e}")
            return False, {}, required_patterns.copy()
    
    def get_folder_created_date(self, folder_path):
        """獲取資料夾創建日期"""
        try:
            ctime = os.path.getctime(folder_path)
            return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ctime))
        except:
            return "未知"
    
    def count_items_in_subfolder(self, folder_path, subfolder_pattern):
        """計算特定子資料夾中的項目數量"""
        try:
            items = os.listdir(folder_path)
            for item in items:
                if subfolder_pattern in item:
                    subfolder_path = os.path.join(folder_path, item)
                    if os.path.isdir(subfolder_path):
                        sub_items = os.listdir(subfolder_path)
                        return len(sub_items)
            return 0
        except:
            return 0
    
    def search_files(self):
        """執行資料夾搜尋 - 修復版本"""
        print("\n" + "="*50)
        print("開始執行搜尋")
        print("="*50)
        
        # 檢查是否有搜尋正在進行
        if self.search_thread and self.search_thread.is_alive():
            messagebox.showwarning("警告", "搜尋正在進行中，請等待完成或點擊停止！")
            return
        
        # 重置搜索狀態
        self.searched_folders.clear()  # 清空已搜索資料夾集合
        
        # 獲取搜尋條件
        criteria = {}
        for key, var in self.search_vars.items():
            value = var.get().strip()
            if value:
                criteria[key] = value
        
        print(f"搜尋模式: {'自訂路徑' if self.search_mode.get() == 'custom' else '全機搜尋'}")
        if self.search_mode.get() == 'custom':
            print(f"自訂路徑: {self.custom_search_path.get()}")
        print(f"搜尋條件: {criteria}")
        print(f"必需檔案模式: {self.required_file_patterns}")

        
        # 顯示搜索條件摘要
        if criteria:
            condition_summary = "搜索條件: "
            condition_parts = []
            for key, value in criteria.items():
                if key == 'year':
                    condition_parts.append(f"年度包含'{value}'")
                elif key == 'category':
                    condition_parts.append(f"分類號包含'{value}'")
                elif key == 'case':
                    condition_parts.append(f"案包含'{value}'")
                elif key == 'volume':
                    condition_parts.append(f"卷包含'{value}'")
                elif key == 'item':
                    condition_parts.append(f"目包含'{value}'")
                elif key == 'sheet_total':
                    condition_parts.append(f"總片數包含'{value}'")
                elif key == 'sheet_num':
                    condition_parts.append(f"片號包含'{value}'")
            
            condition_summary += ", ".join(condition_parts)
            print(condition_summary)
        else:
            result = messagebox.askyesno("確認", "沒有設定任何搜尋條件，將搜尋所有包含指定檔案模式的資料夾。\n這可能需要很長時間，是否繼續？")
            if not result:
                return
        
        # 清除之前的結果
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 開始搜尋
        self.stop_search = False
        self.progress.start(10)
        self.stop_btn.config(state='normal', bg='#da4055', fg='#ffffff')
        self.status_var.set("正在搜尋符合條件的光碟片資料夾...")
        
        # 在新線程中執行搜尋
        self.search_thread = threading.Thread(target=self._search_all_folders, 
                                            args=(criteria, self.required_file_patterns))
        self.search_thread.daemon = True
        self.search_thread.start()

    
    def _search_all_folders(self, criteria, required_patterns):
        """在背景線程中搜尋所有資料夾 - 修復重複問題"""
        matched_folders = []
        total_folders = 0
        
        # 重置已搜索資料夾集合
        self.searched_folders.clear()
        
        try:
            # 獲取搜尋路徑
            search_paths = self.get_search_paths()
            print(f"\n=== 開始搜尋 ===")
            print(f"搜尋條件: {criteria}")
            print(f"必需檔案模式: {required_patterns}")
            print(f"搜尋路徑數量: {len(search_paths)}")
            
            if not search_paths:
                print("警告: 沒有找到任何有效的搜尋路徑!")
                self.root.after(0, lambda: messagebox.showwarning("警告", "沒有找到任何有效的搜尋路徑，請檢查系統設置。"))
                return
            
            # 搜尋所有路徑
            for i, search_path in enumerate(search_paths, 1):
                if self.stop_search:
                    break
                    
                path_name = os.path.basename(search_path) or search_path
                status_msg = f"正在搜尋 ({i}/{len(search_paths)}): {path_name}"
                self.root.after(0, lambda msg=status_msg: self.status_var.set(msg))
                print(f"\n--- 搜尋路徑 {i}/{len(search_paths)}: {search_path} ---")
                
                try:
                    # 檢查路徑是否可以訪問
                    if not os.path.exists(search_path):
                        print(f"路徑不存在: {search_path}")
                        continue
                        
                    if not os.access(search_path, os.R_OK):
                        print(f"無法讀取路徑: {search_path}")
                        continue
                    
                    matched_count, folders_count = self._search_directory_for_folders(
                        search_path, criteria, required_patterns, matched_folders)
                    total_folders += folders_count
                    
                    print(f"在 {search_path} 中找到 {matched_count} 個符合條件的資料夾（共檢查 {folders_count} 個）")
                    
                except PermissionError as e:
                    print(f"權限不足，無法訪問路徑 {search_path}: {e}")
                    continue
                except Exception as e:
                    print(f"無法存取路徑 {search_path}: {e}")
                    continue
            
            print(f"\n=== 搜尋總結 ===")
            print(f"檢查的路徑數: {len(search_paths)}")
            print(f"檢查的資料夾總數: {total_folders}")
            print(f"找到的符合條件資料夾: {len(matched_folders)}")
            print(f"已搜索資料夾數量: {len(self.searched_folders)}")
            
            # 搜尋完成，更新最終結果
            if not self.stop_search:
                self.root.after(0, self._search_completed, matched_folders, total_folders)
            else:
                self.root.after(0, self._search_stopped, matched_folders, total_folders)
                
        except Exception as e:
            print(f"搜尋錯誤: {e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, self._search_error, str(e))
    
    def _search_directory_for_folders(self, directory, criteria, required_patterns, matched_folders):
        """搜尋單個目錄中的資料夾 - 修復重複問題"""
        matched_count = 0
        folders_count = 0
        
        print(f"\n=== 搜尋目錄: {directory} ===")
        
        try:
            if not os.path.exists(directory):
                print(f"  目錄不存在: {directory}")
                return 0, 0
                
            if not os.access(directory, os.R_OK):
                print(f"  無法讀取目錄: {directory}")
                return 0, 0
            
            # 首先列出目錄內容，看看是否可以正常訪問
            try:
                all_items = os.listdir(directory)
                print(f"  目錄中共有 {len(all_items)} 個項目")
            except Exception as e:
                print(f"  無法列出目錄內容: {e}")
                return 0, 0
            
            # 搜尋目錄及子目錄（限制深度）
            max_depth = 3  # 限制最大搜尋深度
            
            for root, dirs, files in os.walk(directory):
                if self.stop_search:
                    break
                
                # 計算當前深度
                current_depth = len(root.replace(directory, '').split(os.sep)) - 1
                if current_depth > max_depth:
                    dirs.clear()  # 不再搜尋更深層的目錄
                    continue
                
                print(f"  正在搜尋: {root} (深度: {current_depth})")
                print(f"    找到 {len(dirs)} 個子資料夾")
                
                # 過濾掉系統資料夾和隱藏資料夾
                dirs[:] = [d for d in dirs if not d.startswith('.') and not d.startswith('$')]
                
                for folder in dirs[:]:  # 使用切片避免修改列表時出錯
                    if self.stop_search:
                        break
                    
                    folders_count += 1
                    folder_path = os.path.join(root, folder)
                    
                    # 檢查是否已經搜索過這個資料夾（使用絕對路徑）
                    abs_folder_path = os.path.abspath(folder_path)
                    if abs_folder_path in self.searched_folders:
                        print(f"    ⚠️ 跳過重複資料夾: {folder}")
                        continue
                    
                    # 將資料夾路徑加入已搜索集合
                    self.searched_folders.add(abs_folder_path)
                    
                    # 跳過無法訪問的資料夾
                    try:
                        if not os.access(folder_path, os.R_OK):
                            continue
                    except:
                        continue
                    
                    print(f"\n    檢查資料夾: {folder}")
                    
                    # 步驟1：檢查資料夾名稱是否符合條件
                    if self.folder_matches_criteria(folder, criteria):
                        print(f"    ✓ 符合名稱條件")
                        
                        # 步驟2：檢查資料夾內是否包含必需的子資料夾
                        is_complete, matched_dirs, missing_patterns = self.check_required_files(
                            folder_path, required_patterns)
                        
                        # 如果至少找到一個必需資料夾就顯示
                        if matched_dirs:
                            # 檢查是否已經存在相同的資料夾（基於路徑）
                            is_duplicate = False
                            for existing_folder in matched_folders:
                                if os.path.abspath(existing_folder['path']) == abs_folder_path:
                                    print(f"    ⚠️ 發現重複資料夾，跳過: {folder}")
                                    is_duplicate = True
                                    break
                            
                            if not is_duplicate:
                                parsed_info = self.parse_folder_name(folder)
                                
                                # 計算片數/總片數
                                sheet_info = f"{parsed_info['sheet_num']}/{parsed_info['sheet_total']}"
                                if not parsed_info['sheet_num'] and not parsed_info['sheet_total']:
                                    sheet_info = "未解析"
                                
                                # 檢查檢測結果（光碟檢測截圖）
                                detection_status = "✓ 存在" if "光碟檢測截圖" in matched_dirs else "✗ 缺失"
                                detection_count = 0
                                if "光碟檢測截圖" in matched_dirs:
                                    detection_count = self.count_items_in_subfolder(folder_path, "光碟檢測截圖")
                                    detection_status = f"✓ 存在 ({detection_count}個檔案)"
                                
                                # 檢查光碟內容
                                content_status = "✓ 存在" if "光碟片內容" in matched_dirs else "✗ 缺失"
                                content_count = 0
                                if "光碟片內容" in matched_dirs:
                                    content_count = self.count_items_in_subfolder(folder_path, "光碟片內容")
                                    content_status = f"✓ 存在 ({content_count}個項目)"
                                
                                # 計算完整度
                                completeness = f"{len(matched_dirs)}/{len(required_patterns)}"
                                if is_complete:
                                    completeness += " (完整)"
                                else:
                                    completeness += " (不完整)"
                                
                                folder_info = {
                                    'year': parsed_info['year'] or '未解析',
                                    'category': parsed_info['category'] or '未解析',
                                    'case': parsed_info['case'] or '未解析',
                                    'volume': parsed_info['volume'] or '未解析',
                                    'item': parsed_info['item'] or '未解析',
                                    'sheet_info': sheet_info,
                                    'detection_result': detection_status,
                                    'cd_content': content_status,
                                    'completeness': completeness,
                                    'path': folder_path,
                                    'is_complete': is_complete
                                }
                                
                                matched_folders.append(folder_info)
                                matched_count += 1
                                
                                # 即時更新界面
                                self.root.after(0, self._update_single_result, folder_info)
                                
                                print(f"    ✓ 已加入結果列表 ({completeness})")
                        else:
                            print(f"    ✗ 沒有找到任何必需的子資料夾")
                    else:
                        print(f"    ✗ 不符合名稱條件")
                    
                    # 每檢查一定數量的資料夾就更新狀態
                    if folders_count % 50 == 0:
                        status_msg = f"已檢查 {folders_count} 個資料夾，找到 {matched_count} 個符合條件"
                        self.root.after(0, lambda msg=status_msg: self.status_var.set(msg))
                
                # 限制搜尋深度，避免搜尋過深
                if current_depth >= max_depth:
                    dirs.clear()  # 不再搜尋更深層的目錄
                        
        except Exception as e:
            print(f"搜尋目錄錯誤: {directory}, 錯誤: {e}")
            import traceback
            traceback.print_exc()
                
        print(f"  搜尋完成: 檢查了 {folders_count} 個資料夾，找到 {matched_count} 個符合條件")
        return matched_count, folders_count

    
    def _update_single_result(self, folder_info):
        """更新單個結果到樹狀視圖"""
        # 根據完整度設置不同的標籤顏色
        tags = ['complete'] if folder_info['is_complete'] else ['incomplete']
        
        item_id = self.tree.insert('', 'end', values=(
            '✏  /  🗑',  # 編輯/刪除按鈕
            folder_info['year'],
            folder_info['category'],
            folder_info['case'],
            folder_info['volume'],
            folder_info['item'],
            folder_info['sheet_info'],
            folder_info['detection_result'],
            folder_info['cd_content'],
            folder_info['completeness']
        ), tags=tags)

        self.item_paths[item_id] = folder_info['path']
        
        # 設置標籤樣式
        if not hasattr(self, 'tags_configured'):
            self.tree.tag_configure('complete', background='lightgreen')
            self.tree.tag_configure('incomplete', background='lightyellow')
            self.tags_configured = True
    
    def _search_completed(self, matched_folders, total_folders):
        """搜索完成"""
        self.progress.stop()
        self.stop_btn.config(state='disabled', bg='#cccccc', fg="#FFFFFF")
        
        if len(matched_folders) == 0:
            messagebox.showinfo("搜索結果", "沒有找到符合條件的光碟片資料夾")
            self.status_var.set(f"搜索完成：共檢查 {total_folders} 個資料夾，未找到符合條件的資料夾")
        else:
            complete_count = sum(1 for folder in matched_folders if folder['is_complete'])
            incomplete_count = len(matched_folders) - complete_count
            
            result_msg = f"找到 {len(matched_folders)} 個光碟片資料夾：{complete_count} 個完整，{incomplete_count} 個不完整"
            messagebox.showinfo("搜索結果", result_msg)
            self.status_var.set(f"搜索完成：共檢查 {total_folders} 個資料夾，找到 {len(matched_folders)} 個符合條件的資料夾（{complete_count} 個完整，{incomplete_count} 個不完整）")
    
    def _search_stopped(self, matched_folders, total_folders):
        """搜索被停止"""
        self.progress.stop()
        self.stop_btn.config(state='disabled')
        
        if len(matched_folders) == 0:
            self.status_var.set(f"搜索已停止：已檢查 {total_folders} 個資料夾，未找到符合條件的資料夾")
        else:
            complete_count = sum(1 for folder in matched_folders if folder['is_complete'])
            incomplete_count = len(matched_folders) - complete_count
            self.status_var.set(f"搜索已停止：已檢查 {total_folders} 個資料夾，找到 {len(matched_folders)} 個符合條件的資料夾（{complete_count} 個完整，{incomplete_count} 個不完整）")
    
    def _search_error(self, error_msg):
        """搜索發生錯誤"""
        self.progress.stop()
        self.stop_btn.config(state='disabled')
        self.status_var.set("搜索發生錯誤")
        messagebox.showerror("錯誤", f"搜索過程中發生錯誤：{error_msg}")
    
    def open_folder_location(self, event):
        """雙擊開啟資料夾位置"""
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        
        column = self.tree.identify_column(event.x)
        column_name = self.tree.heading(column)['text']
        
        # 如果點擊的是編輯/刪除欄位，不執行開啟資料夾
        if column_name in ['編輯/刪除']:
            return
        
        # 使用 item_paths 字典獲取路徑
        if item_id in self.item_paths:
            folder_path = self.item_paths[item_id]
            self.open_folder_by_path(folder_path)
        else:
            messagebox.showerror("錯誤", "無法找到資料夾路徑信息")

    # ==================== 指定年份查閱功能 ====================
    
    def open_year_search_window(self):
        """開啟指定年份查閱視窗"""
        from datetime import datetime
        
        year_window = tk.Toplevel(self.root)
        year_window.title("指定上傳年份資料查閱")
        year_window.geometry("700x800")
        year_window.configure(bg='#f8f9fa')
        
        # 置中視窗
        year_window.update_idletasks()
        x = (year_window.winfo_screenwidth() // 2) - (700 // 2)
        y = (year_window.winfo_screenheight() // 2) - (800 // 2)
        year_window.geometry(f"700x800+{x}+{y}")
        
        # 標題
        title_label = tk.Label(year_window, text="特定上傳年份資料查閱",
                              font=("Microsoft YaHei", 16, "bold"),
                              bg='#f8f9fa', fg='#2c3e50')
        title_label.pack(pady=20)
        
        # 年份選擇區域
        select_frame = tk.Frame(year_window, bg='#f8f9fa')
        select_frame.pack(pady=10)
        
        tk.Label(select_frame, text="選擇年份（民國）：",
                font=("Microsoft YaHei", 12),
                bg='#f8f9fa').pack(side='left', padx=10)
        
        # 生成年份選項
        current_year = datetime.now().year - 1911
        year_options = [f"{y}年" for y in range(100, current_year + 1)]
        year_options.reverse()
        
        year_var = tk.StringVar(value=f"{current_year}年")
        year_combo = ttk.Combobox(select_frame, textvariable=year_var,
                                 values=year_options, state='readonly',
                                 font=("Microsoft YaHei", 11), width=15)
        year_combo.pack(side='left', padx=10)
        
        # 查詢按鈕
        search_btn = tk.Button(select_frame, text="🔍 查詢",
                              command=lambda: self._search_by_year(year_var.get(), year_window),
                              bg='#28a745', fg='white',
                              font=("Microsoft YaHei", 11, "bold"),
                              width=10, height=1)
        search_btn.pack(side='left', padx=10)
        
        # 結果顯示區域
        self.year_result_frame = tk.Frame(year_window, bg='#f8f9fa')
        self.year_result_frame.pack(fill='both', expand=True, padx=20, pady=10)
    
    def _search_by_year(self, year_str, parent_window):
        """根據選擇的年份搜尋資料 - 完全去重版本"""
        from datetime import datetime
        
        year = year_str.replace('年', '').strip()
        
        self.year_search_data = []
        self.current_page = 0
        self.selected_files = set()
        
        # 用字典儲存，key為資料夾路徑，避免重複
        folder_dict = {}
        
        # 取得搜尋路徑
        search_paths = self.get_search_paths()
        
        print(f"\n=== 開始搜尋年份：{year} ===")
        
        # 搜尋指定年份的資料夾
        for search_path in search_paths:
            if os.path.exists(search_path):
                print(f"搜尋路徑：{search_path}")
                for root, dirs, files in os.walk(search_path):
                    # 取得絕對路徑
                    abs_root = os.path.abspath(root)
                    
                    # 檢查是否為該年份的資料夾
                    if self._is_year_folder(root, year):
                        # 如果這個資料夾已經處理過，跳過
                        if abs_root in folder_dict:
                            print(f"  跳過重複資料夾：{abs_root}")
                            continue
                        
                        # 收集圖片
                        images = self._collect_folder_images(root)
                        if images:
                            print(f"  找到資料夾：{os.path.basename(root)}，圖片數：{len(images)}")
                            # 用絕對路徑作為 key，確保不重複
                            folder_dict[abs_root] = images
        
        print(f"找到 {len(folder_dict)} 個不重複的資料夾")
        
        # 轉換成列表格式
        for folder_path, images in folder_dict.items():
            file_number = self._extract_file_number(images[0])
            self.year_search_data.append((file_number, images))
        
        # 按照建立時間排序（新到舊）
        self.year_search_data.sort(key=lambda x: self._get_folder_time(x[1][0]), reverse=True)
        
        print(f"最終顯示 {len(self.year_search_data)} 筆資料")
        
        # 顯示搜尋結果
        if self.year_search_data:
            self._create_year_search_page()
        else:
            messagebox.showinfo("查詢結果", f"未找到{year}年上傳的資料")
    
    def _is_year_folder(self, folder_path, year):
        """判斷資料夾是否屬於指定年份 - 改進版"""
        from datetime import datetime
        try:
            # 方法1: 檢查資料夾建立時間
            creation_time = os.path.getctime(folder_path)
            folder_year = datetime.fromtimestamp(creation_time).year - 1911
            
            # 允許±1年的誤差（因為跨年上傳的情況）
            if abs(folder_year - int(year)) <= 1:
                return True
            
            # 方法2: 檢查資料夾名稱中是否包含年份
            folder_name = os.path.basename(folder_path)
            
            # 檢查資料夾名稱是否以年份開頭
            if folder_name.startswith(year + '-'):
                return True
            
            # 方法3: 檢查「光碟檢測截圖」子資料夾中的檔案
            screenshots_folder = os.path.join(folder_path, "光碟檢測截圖")
            if os.path.exists(screenshots_folder):
                try:
                    files = os.listdir(screenshots_folder)
                    for file in files:
                        if file.startswith(year + '-'):
                            return True
                except:
                    pass
                    
            return False
        except:
            return False
    
    def _collect_folder_images(self, folder_path):
        """收集資料夾中的所有圖片及其資訊 - 去重版本"""
        images = []
        image_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
        
        # 用集合追蹤已處理的圖片，避免重複
        processed_images = set()
        
        # 檢查「光碟檢測截圖」子資料夾
        screenshots_folder = os.path.join(folder_path, "光碟檢測截圖")
        if os.path.exists(screenshots_folder):
            search_folder = screenshots_folder
        else:
            search_folder = folder_path
        
        try:
            for file in os.listdir(search_folder):
                if file.lower().endswith(image_extensions):
                    image_path = os.path.join(search_folder, file)
                    abs_image_path = os.path.abspath(image_path)
                    
                    # 檢查是否已處理過這張圖片
                    if abs_image_path in processed_images:
                        continue
                    
                    processed_images.add(abs_image_path)
                    
                    img_info = self._parse_image_info(image_path, folder_path)
                    if img_info:
                        images.append(img_info)
        except Exception as e:
            print(f"收集圖片時發生錯誤: {e}")
        
        # 按照圖片編號排序
        images.sort(key=lambda x: x.get('image_num', 0))
        return images
    
    def _parse_image_info(self, image_path, folder_path):
        """從圖片路徑和檔名解析資訊 - 改進版"""
        filename = os.path.basename(image_path)
        
        # 移除副檔名
        name_without_ext = filename
        for ext in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.PNG', '.JPG', '.JPEG']:
            if name_without_ext.endswith(ext):
                name_without_ext = name_without_ext[:-len(ext)]
                break
        
        # 解析檔名格式
        parts = name_without_ext.split('-')
        
        # 初始化預設值
        info = {
            'image_path': image_path,
            'filename': filename,
            'folder_path': os.path.dirname(folder_path),
            'upload_year': '未知',
            'data_year': '未知',
            'category': '未知',
            'case': '未知',
            'volume': '未知',
            'item': '未知',
            'sheet_total': '未知',
            'sheet_num': '未知',
            'image_num': 0
        }
        
        # 嘗試從檔名解析（格式：年-年-分類號-案-卷-目-總片數-編號）
        if len(parts) >= 8:
            try:
                info['upload_year'] = parts[0] if parts[0] and parts[0] != '' else '未知'
                info['data_year'] = parts[1] if parts[1] and parts[1] != '' else '未知'
                info['category'] = parts[2] if parts[2] and parts[2] != '' else '未知'
                info['case'] = parts[3] if parts[3] and parts[3] != '' else '未知'
                info['volume'] = parts[4] if parts[4] and parts[4] != '' else '未知'
                info['item'] = parts[5] if parts[5] and parts[5] != '' else '未知'
                info['sheet_total'] = parts[6] if parts[6] and parts[6] != '' else '未知'
                info['sheet_num'] = parts[7] if parts[7] and parts[7] != '' else '未知'
                info['image_num'] = int(parts[7]) if parts[7] and parts[7].isdigit() else 0
            except (ValueError, IndexError):
                pass
        
        # 如果檔名解析失敗，嘗試從資料夾名稱解析
        if info['category'] == '未知':
            try:
                # 取得父資料夾名稱
                parent_folder = os.path.basename(os.path.dirname(folder_path))
                folder_parts = parent_folder.split('-')
                
                if len(folder_parts) >= 7:
                    info['upload_year'] = folder_parts[0] if folder_parts[0] else info['upload_year']
                    info['data_year'] = folder_parts[0] if folder_parts[0] else info['data_year']  # 如果沒有資料年度，使用上傳年度
                    info['category'] = folder_parts[1] if folder_parts[1] else info['category']
                    info['case'] = folder_parts[2] if folder_parts[2] else info['case']
                    info['volume'] = folder_parts[3] if folder_parts[3] else info['volume']
                    info['item'] = folder_parts[4] if folder_parts[4] else info['item']
                    info['sheet_total'] = folder_parts[5] if folder_parts[5] else info['sheet_total']
                    info['sheet_num'] = folder_parts[6] if folder_parts[6] else info['sheet_num']
            except:
                pass
        
        return info
    
    def _extract_file_number(self, img_info):
        """從圖片資訊中提取檔號 - 改進版"""
        parts = []
        
        # 只加入非「未知」的部分
        for key in ['category', 'case', 'volume', 'item', 'sheet_total', 'sheet_num']:
            value = img_info.get(key, '未知')
            if value and value != '未知':
                parts.append(value)
            else:
                # 如果遇到未知，嘗試使用資料夾名稱
                parts.append('未知')
        
        # 如果所有都是未知，使用檔案名稱
        if all(p == '未知' for p in parts):
            return os.path.basename(img_info.get('folder_path', '未知檔案'))
        
        return '-'.join(parts)
    def _get_folder_time(self, img_info):
        """取得資料夾建立時間"""
        try:
            return os.path.getctime(img_info['folder_path'])
        except:
            return 0
    
    def _create_year_search_page(self):
        """建立年份搜尋結果的分頁顯示"""
        for widget in self.year_result_frame.winfo_children():
            widget.destroy()
        
        total_files = len(self.year_search_data)
        total_pages = (total_files + self.items_per_page - 1) // self.items_per_page
        
        if total_files == 0:
            no_data_label = tk.Label(self.year_result_frame, 
                                    text="未找到符合條件的資料",
                                    font=("Microsoft YaHei", 12),
                                    bg='#f8f9fa', fg='#7f8c8d')
            no_data_label.pack(pady=50)
            return
        
        year_display = "未知年度"
        if self.year_search_data and self.year_search_data[0][1]:
            first_upload_year = self.year_search_data[0][1][0].get('upload_year', '未知')
            year_display = f"民國 {first_upload_year} 年"
        
        # 資訊和控制按鈕
        info_frame = tk.Frame(self.year_result_frame, bg='#f8f9fa')
        info_frame.pack(fill='x', pady=(0, 10))
        
        info_text = f"{year_display} - 找到 {total_files} 個檔案"
        info_label = tk.Label(info_frame, text=info_text,
                            font=("Microsoft YaHei", 11, "bold"),
                            bg='#f8f9fa', fg='#2c3e50')
        info_label.pack(side='left')
        
        control_frame = tk.Frame(info_frame, bg='#f8f9fa')
        control_frame.pack(side='right')
        
        # 全選按鈕
        def toggle_select_all():
            if len(self.selected_files) == total_files:
                self.selected_files.clear()
            else:
                self.selected_files = set(range(total_files))
            self._create_year_search_page()
        
        select_all_text = "□ 全選" if len(self.selected_files) < total_files else "☑ 取消全選"
        select_btn = tk.Button(control_frame, text=select_all_text,
                              command=toggle_select_all,
                              bg='#6c757d', fg='white',
                              font=("Microsoft YaHei", 10, "bold"),
                              width=11, height=1)
        select_btn.pack(side='left', padx=(0, 20))
        
        # 另存PDF按鈕
        def save_selected_as_pdf():
            if not self.selected_files:
                messagebox.showwarning("警告", "請先選擇要另存為PDF的資料！")
                return
            
            selected_data = []
            for i in self.selected_files:
                if i < len(self.year_search_data):
                    selected_data.append(self.year_search_data[i])
            
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                title="儲存報表為PDF"
            )
            if filename:
                try:
                    self._save_report_as_pdf(selected_data, year_display, filename)
                    messagebox.showinfo("成功", f"報表已儲存至：{filename}")
                except Exception as e:
                    messagebox.showerror("錯誤", f"儲存PDF失敗：{str(e)}")

        pdf_btn = tk.Button(control_frame, text="💾 另存PDF",
                            command=save_selected_as_pdf,
                            bg="#dc3545", fg='white',
                            font=("Microsoft YaHei", 10, "bold"),
                            width=11, height=1)
        pdf_btn.pack(side='right', padx=20)

        # 主內容區域
        main_content_frame = tk.Frame(self.year_result_frame, bg='#f8f9fa')
        main_content_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        canvas = tk.Canvas(main_content_frame, bg='#f8f9fa')
        scrollbar = ttk.Scrollbar(main_content_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#f8f9fa')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def bind_to_mousewheel(widget):
            widget.bind("<MouseWheel>", _on_mousewheel)
            for child in widget.winfo_children():
                bind_to_mousewheel(child)
        
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        start_idx = self.current_page * self.items_per_page
        end_idx = min(start_idx + self.items_per_page, total_files)
        
        for i in range(start_idx, end_idx):
            if i < len(self.year_search_data):
                file_number, images = self.year_search_data[i]
                self._create_file_display_item_with_checkbox(scrollable_frame, file_number, images, i, i - start_idx)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        bind_to_mousewheel(scrollable_frame)
        
        # 分頁控制
        # 分頁控制框架 - 修改這個部分
# 分頁控制框架 - 修改這個部分
        page_control_frame = tk.Frame(self.year_result_frame, bg='#f8f9fa')
        page_control_frame.pack(fill='x', pady=(10, 0))

        # 上一頁按鈕
        prev_btn = tk.Button(page_control_frame, text="⬅ 上一頁", 
                            command=self._prev_page,
                            bg='#6c757d', fg='white', 
                            font=("Microsoft YaHei", 11, "bold"),
                            width=8, height=1,
                            state='normal' if self.current_page > 0 else 'disabled')
        prev_btn.pack(side='left', padx=20)

        # 頁碼顯示 - 使用新的置中方法
        self._create_page_numbers_centered(page_control_frame, total_pages)

        # 下一頁按鈕
        next_btn = tk.Button(page_control_frame, text="下一頁 ➡", 
                            command=self._next_page,
                            bg='#6c757d', fg='white', 
                            font=("Microsoft YaHei", 11, "bold"),
                            width=8, height=1,
                            state='normal' if self.current_page < total_pages - 1 else 'disabled')
        next_btn.pack(side='right', padx=20)
        
    def _create_page_numbers_centered(self, parent, total_pages):
        """建立置中的頁碼按鈕"""
        # 創建一個容器框架，使用 place 來置中
        container = tk.Frame(parent, bg='#f8f9fa')
        container.place(relx=0.5, rely=0.5, anchor='center')
        
        current_page_num = self.current_page + 1
        max_page_digits = len(str(total_pages))
        button_width = max(2, max_page_digits)
        button_font_size = 9

        if total_pages <= 7:
            start_page = 1
            end_page = total_pages
            show_first_ellipsis = False
            show_last_ellipsis = False
            show_first_page = False
            show_last_page = False
        else:
            middle_start = max(2, current_page_num - 1)
            middle_end = min(total_pages - 1, current_page_num + 1)
            
            if middle_end - middle_start < 2:
                if middle_start == 2:
                    middle_end = min(total_pages - 1, middle_start + 2)
                else:
                    middle_start = max(2, middle_end - 2)
            
            start_page = middle_start
            end_page = middle_end
            show_first_page = True
            show_last_page = True
            show_first_ellipsis = start_page > 2
            show_last_ellipsis = end_page < total_pages - 1

        # 顯示第一頁按鈕
        if show_first_page:
            first_btn = tk.Button(container, text="1",
                                command=lambda: self._go_to_page(0),
                                bg='#007bff' if current_page_num == 1 else '#ffffff',
                                fg='white' if current_page_num == 1 else '#495057',
                                font=("Microsoft YaHei", button_font_size, "bold" if current_page_num == 1 else "normal"),
                                width=button_width, height=1,
                                relief='solid', bd=1)
            first_btn.pack(side='left', padx=1)

        # 顯示第一個省略號
        if show_first_ellipsis:
            dots_label = tk.Label(container, text="...",
                                font=("Microsoft YaHei", button_font_size),
                                bg='#f8f9fa', fg='#6c757d')
            dots_label.pack(side='left', padx=2)

        # 顯示中間的頁碼按鈕
        for page_num in range(start_page, end_page + 1):
            page_index = page_num - 1
            
            if page_index == self.current_page:
                page_btn = tk.Button(container, text=str(page_num),
                                    command=lambda p=page_index: self._go_to_page(p),
                                    bg='#007bff', fg='white',
                                    font=("Microsoft YaHei", button_font_size, "bold"),
                                    width=button_width, height=1,
                                    relief='solid', bd=1)
            else:
                page_btn = tk.Button(container, text=str(page_num),
                                    command=lambda p=page_index: self._go_to_page(p),
                                    bg='#ffffff', fg='#495057',
                                    font=("Microsoft YaHei", button_font_size),
                                    width=button_width, height=1,
                                    relief='solid', bd=1)
            
            page_btn.pack(side='left', padx=1)

        # 顯示最後一個省略號
        if show_last_ellipsis:
            dots_label = tk.Label(container, text="...",
                                font=("Microsoft YaHei", button_font_size),
                                bg='#f8f9fa', fg='#6c757d')
            dots_label.pack(side='left', padx=2)

        # 顯示最後一頁按鈕
        if show_last_page:
            last_btn = tk.Button(container, text=str(total_pages),
                                command=lambda: self._go_to_page(total_pages - 1),
                                bg='#007bff' if current_page_num == total_pages else '#ffffff',
                                fg='white' if current_page_num == total_pages else '#495057',
                                font=("Microsoft YaHei", button_font_size, "bold" if current_page_num == total_pages else "normal"),
                                width=button_width, height=1,
                                relief='solid', bd=1)
            last_btn.pack(side='left', padx=1)
    
    def _create_file_display_item_with_checkbox(self, parent_frame, file_number, images, global_index, index):
        """創建帶有勾選框的檔案顯示項目"""
        display_file_number = file_number
        if images:
            first_image = images[0]
            data_year = first_image.get('data_year', '')
            if data_year:
                display_file_number = f"{data_year}-{file_number}"
        
        item_frame = tk.LabelFrame(parent_frame, 
                                text="",
                                font=("Microsoft YaHei", 12, "bold"),
                                bg='#ffffff', fg='#2c3e50',
                                padx=15, pady=10)
        item_frame.pack(fill='x', expand=False, padx=5, pady=8, ipady=5)
        
        title_frame = tk.Frame(item_frame, bg='#ffffff')
        title_frame.pack(fill='x', pady=(0, 10))
        
        checkbox_var = tk.BooleanVar()
        checkbox_var.set(global_index in self.selected_files)
        
        def on_checkbox_change():
            if checkbox_var.get():
                self.selected_files.add(global_index)
            else:
                self.selected_files.discard(global_index)
            self._create_year_search_page()
        
        checkbox = tk.Checkbutton(title_frame, 
                                variable=checkbox_var,
                                command=on_checkbox_change,
                                bg='#ffffff',
                                font=("Microsoft YaHei", 12))
        checkbox.pack(side='left', padx=(0, 10))
        
        title_label = tk.Label(title_frame, 
                            text=f"檔號：{display_file_number}",
                            font=("Microsoft YaHei", 12, "bold"),
                            bg='#ffffff', fg='#2c3e50')
        title_label.pack(side='left')
        
        info_frame = tk.Frame(item_frame, bg='#ffffff')
        info_frame.pack(fill='x', expand=False, pady=(0, 10))
        
        if images:
            first_image = images[0]
            
            info_lines = [
                f"上傳年度：{first_image.get('upload_year', '未知')}    資料年度：{first_image.get('data_year', '未知')}    圖片數量：{len(images)} 張",
                f"分類號：{first_image.get('category', '未知')}    案：{first_image.get('case', '未知')}    卷：{first_image.get('volume', '未知')}    目：{first_image.get('item', '未知')}    片數：{first_image.get('sheet_num', '未知')} / {first_image.get('sheet_total', '未知')}"
            ]
            
            for line in info_lines:
                info_label = tk.Label(info_frame, text=line,
                                    font=("Microsoft YaHei", 10),
                                    bg='#ffffff', fg='#495057',
                                    anchor='w')
                info_label.pack(fill='x', anchor='w', pady=1)
            
            button_frame = tk.Frame(item_frame, bg='#ffffff')
            button_frame.pack(fill='x', pady=(10, 0))
            
            folder_path = first_image.get('folder_path', '')
            if folder_path:
                open_btn = tk.Button(button_frame, text="📂 開啟資料夾",
                                    command=lambda: self.open_folder_by_path(folder_path),
                                    bg='#3498db', fg='white',
                                    font=("Microsoft YaHei", 10, "bold"),
                                    width=12, height=1)
                open_btn.pack(side='left', padx=(0, 15))
            
            view_btn = tk.Button(button_frame, text="🖼️ 圖片檢視",
                                command=lambda imgs=images: self._show_image_viewer(imgs),
                                bg='#27ae60', fg='white',
                                font=("Microsoft YaHei", 10, "bold"),
                                width=12, height=1)
            view_btn.pack(side='left', padx=15)
    
    def _prev_page(self):
        """上一頁"""
        if self.current_page > 0:
            self.current_page -= 1
            self._create_year_search_page()

    def _next_page(self):
        """下一頁"""
        total_pages = (len(self.year_search_data) + self.items_per_page - 1) // self.items_per_page
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._create_year_search_page()

    def _go_to_page(self, page_index):
        """跳轉到指定頁"""
        self.current_page = page_index
        self._create_year_search_page()

    def _create_page_numbers(self, parent, total_pages):
        """建立頁碼按鈕"""
        page_numbers_frame = tk.Frame(parent, bg='#f8f9fa')
        page_numbers_frame.pack(side='left', padx=20)

        inner_page_frame = tk.Frame(page_numbers_frame, bg='#f8f9fa')
        inner_page_frame.pack()

        current_page_num = self.current_page + 1
        max_page_digits = len(str(total_pages))
        button_width = max(2, max_page_digits)
        button_font_size = 9

        if total_pages <= 7:
            start_page = 1
            end_page = total_pages
            show_first_ellipsis = False
            show_last_ellipsis = False
            show_first_page = False
            show_last_page = False
        else:
            middle_start = max(2, current_page_num - 1)
            middle_end = min(total_pages - 1, current_page_num + 1)
            
            if middle_end - middle_start < 2:
                if middle_start == 2:
                    middle_end = min(total_pages - 1, middle_start + 2)
                else:
                    middle_start = max(2, middle_end - 2)
            
            start_page = middle_start
            end_page = middle_end
            show_first_page = True
            show_last_page = True
            show_first_ellipsis = start_page > 2
            show_last_ellipsis = end_page < total_pages - 1

        # 顯示第一頁按鈕
        if show_first_page:
            first_btn = tk.Button(inner_page_frame, text="1",
                                command=lambda: self._go_to_page(0),
                                bg='#007bff' if current_page_num == 1 else '#ffffff',
                                fg='white' if current_page_num == 1 else '#495057',
                                font=("Microsoft YaHei", button_font_size, "bold" if current_page_num == 1 else "normal"),
                                width=button_width, height=1,
                                relief='solid', bd=1)
            first_btn.pack(side='left', padx=1)

        # 顯示第一個省略號
        if show_first_ellipsis:
            dots_label = tk.Label(inner_page_frame, text="...",
                                font=("Microsoft YaHei", button_font_size),
                                bg='#f8f9fa', fg='#6c757d')
            dots_label.pack(side='left', padx=2)

        # 顯示中間的頁碼按鈕
        for page_num in range(start_page, end_page + 1):
            page_index = page_num - 1
            
            if page_index == self.current_page:
                page_btn = tk.Button(inner_page_frame, text=str(page_num),
                                    command=lambda p=page_index: self._go_to_page(p),
                                    bg='#007bff', fg='white',
                                    font=("Microsoft YaHei", button_font_size, "bold"),
                                    width=button_width, height=1,
                                    relief='solid', bd=1)
            else:
                page_btn = tk.Button(inner_page_frame, text=str(page_num),
                                    command=lambda p=page_index: self._go_to_page(p),
                                    bg='#ffffff', fg='#495057',
                                    font=("Microsoft YaHei", button_font_size),
                                    width=button_width, height=1,
                                    relief='solid', bd=1)
            
            page_btn.pack(side='left', padx=1)

        # 顯示最後一個省略號
        if show_last_ellipsis:
            dots_label = tk.Label(inner_page_frame, text="...",
                                font=("Microsoft YaHei", button_font_size),
                                bg='#f8f9fa', fg='#6c757d')
            dots_label.pack(side='left', padx=2)

        # 顯示最後一頁按鈕
        if show_last_page:
            last_btn = tk.Button(inner_page_frame, text=str(total_pages),
                                command=lambda: self._go_to_page(total_pages - 1),
                                bg='#007bff' if current_page_num == total_pages else '#ffffff',
                                fg='white' if current_page_num == total_pages else '#495057',
                                font=("Microsoft YaHei", button_font_size, "bold" if current_page_num == total_pages else "normal"),
                                width=button_width, height=1,
                                relief='solid', bd=1)
            last_btn.pack(side='left', padx=1)
            
    def _show_image_viewer(self, images):
        """顯示圖片檢視視窗"""
        try:
            from PIL import Image, ImageTk
        except ImportError:
            messagebox.showerror("錯誤", "需要安裝 Pillow 套件才能使用圖片檢視功能\n請執行: pip install Pillow")
            return
        
        viewer_window = tk.Toplevel(self.root)
        viewer_window.title("圖片檢視")
        viewer_window.geometry("1000x700")
        viewer_window.configure(bg='#2c3e50')
        
        viewer_window.update_idletasks()
        x = (viewer_window.winfo_screenwidth() // 2) - (1000 // 2)
        y = (viewer_window.winfo_screenheight() // 2) - (700 // 2)
        viewer_window.geometry(f"1000x700+{x}+{y}")
        
        current_index = [0]
        
        def update_image_display():
            """更新圖片顯示"""
            for widget in content_frame.winfo_children():
                widget.destroy()
            
            img_info = images[current_index[0]]
            
            # 頂部標頭
            header_frame = tk.Frame(content_frame, bg='#34495e', height=40)
            header_frame.pack(fill='x')
            header_frame.pack_propagate(False)
            
            header_text = f"檔號：{self._extract_file_number(img_info)} | 第 {current_index[0] + 1} 張 / 共 {len(images)} 張"
            header_label = tk.Label(header_frame, text=header_text,
                                   font=("Microsoft YaHei", 11, "bold"),
                                   bg='#34495e', fg='white')
            header_label.pack(expand=True)
            
            # 主內容區
            main_frame = tk.Frame(content_frame, bg='#2c3e50')
            main_frame.pack(fill='both', expand=True, padx=20, pady=20)
            
            # 左側圖片
            image_frame = tk.Frame(main_frame, bg='#34495e', width=600)
            image_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
            
            try:
                img = Image.open(img_info['image_path'])
                img.thumbnail((580, 500), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                img_label = tk.Label(image_frame, image=photo, bg='#34495e')
                img_label.image = photo
                img_label.pack(expand=True)
            except Exception as e:
                error_label = tk.Label(image_frame, text=f"無法載入圖片\n{str(e)}",
                                     font=("Microsoft YaHei", 10),
                                     bg='#34495e', fg='#e74c3c')
                error_label.pack(expand=True)
            
            # 右側資訊
            info_frame = tk.Frame(main_frame, bg='#34495e', width=300)
            info_frame.pack(side='right', fill='both', padx=(10, 0))
            
            info_title = tk.Label(info_frame, text="圖片資訊",
                                font=("Microsoft YaHei", 12, "bold"),
                                bg='#34495e', fg='white', pady=10)
            info_title.pack(fill='x')
            
            info_items = [
                ("檔案名稱", img_info.get('filename', '未知')),
                ("上傳年度", img_info.get('upload_year', '未知')),
                ("資料年度", img_info.get('data_year', '未知')),
                ("分類號", img_info.get('category', '未知')),
                ("案", img_info.get('case', '未知')),
                ("卷", img_info.get('volume', '未知')),
                ("目", img_info.get('item', '未知')),
                ("片數", f"{img_info.get('sheet_num', '未知')} / {img_info.get('sheet_total', '未知')}"),
                ("圖片編號", img_info.get('image_num', '未知'))
            ]
            
            for label, value in info_items:
                item_frame = tk.Frame(info_frame, bg='#34495e')
                item_frame.pack(fill='x', padx=10, pady=3)
                
                label_widget = tk.Label(item_frame, text=f"{label}：",
                                      font=("Microsoft YaHei", 9),
                                      bg='#34495e', fg='#95a5a6',
                                      anchor='w', width=10)
                label_widget.pack(side='left')
                
                value_widget = tk.Label(item_frame, text=str(value),
                                      font=("Microsoft YaHei", 10),
                                      bg='#34495e', fg='white',
                                      anchor='w')
                value_widget.pack(side='left', fill='x', expand=True)
            
            # 底部資訊
            bottom_frame = tk.Frame(content_frame, bg='#2c3e50', height=40)
            bottom_frame.pack(fill='x')
            bottom_frame.pack_propagate(False)
            
            filename = img_info.get('filename', '未知')
            try:
                detector = self.get_folder_creator(img_info)
            except:
                detector = "未知"
            
            bottom_left = tk.Label(bottom_frame, text=f"檔名：{filename}",
                                 font=("Microsoft YaHei", 9),
                                 bg='#2c3e50', fg='white', anchor='w')
            bottom_left.pack(side='left', padx=20)
            
            bottom_right = tk.Label(bottom_frame, text=f"檢測人：{detector}",
                                  font=("Microsoft YaHei", 9),
                                  bg='#2c3e50', fg='white', anchor='e')
            bottom_right.pack(side='right', padx=20)
        
        content_frame = tk.Frame(viewer_window, bg='#2c3e50')
        content_frame.pack(fill='both', expand=True)
        
        control_frame = tk.Frame(viewer_window, bg='#2c3e50')
        control_frame.pack(fill='x', pady=10)
        
        def prev_image():
            if current_index[0] > 0:
                current_index[0] -= 1
                update_image_display()
        
        def next_image():
            if current_index[0] < len(images) - 1:
                current_index[0] += 1
                update_image_display()
        
        def first_image():
            current_index[0] = 0
            update_image_display()
        
        def last_image():
            current_index[0] = len(images) - 1
            update_image_display()
        
        first_btn = tk.Button(control_frame, text="⏮ 第一張",
                             command=first_image,
                             bg='#3498db', fg='white',
                             font=("Microsoft YaHei", 10, "bold"),
                             width=10)
        first_btn.pack(side='left', padx=20)
        
        prev_btn = tk.Button(control_frame, text="⬅ 上一張",
                            command=prev_image,
                            bg='#27ae60', fg='white',
                            font=("Microsoft YaHei", 10, "bold"),
                            width=10)
        prev_btn.pack(side='left', padx=10)
        
        next_btn = tk.Button(control_frame, text="下一張 ➡",
                            command=next_image,
                            bg='#27ae60', fg='white',
                            font=("Microsoft YaHei", 10, "bold"),
                            width=10)
        next_btn.pack(side='right', padx=10)
        
        last_btn = tk.Button(control_frame, text="最後一張 ⏭",
                            command=last_image,
                            bg='#3498db', fg='white',
                            font=("Microsoft YaHei", 10, "bold"),
                            width=10)
        last_btn.pack(side='right', padx=20)
        
        close_btn = tk.Button(control_frame, text="❌ 關閉",
                             command=viewer_window.destroy,
                             bg='#e74c3c', fg='white',
                             font=("Microsoft YaHei", 10, "bold"),
                             width=10)
        close_btn.pack(side='right', padx=50)
        
        update_image_display()
    
    def get_folder_creator(self, img_info):
        """取得資料夾建立者（檢測人）"""
        try:
            folder_path = img_info.get('folder_path', '')
            if folder_path and os.path.exists(folder_path):
                try:
                    import win32security
                    sd = win32security.GetFileSecurity(folder_path, win32security.OWNER_SECURITY_INFORMATION)
                    owner_sid = sd.GetSecurityDescriptorOwner()
                    name, domain, type = win32security.LookupAccountSid(None, owner_sid)
                    return name
                except:
                    pass
        except:
            pass
        return "未知"
    
    def _save_report_as_pdf(self, selected_data, year_display, filename):
        """儲存報表為PDF - 雙圖片佈局"""
        try:
            try:
                from reportlab.lib.pagesizes import A4
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, PageBreak, Table, TableStyle
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.units import inch
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
                from reportlab.lib import colors
            except ImportError:
                raise Exception("需要安裝 reportlab 套件：pip install reportlab")
            
            # 註冊中文字體
            try:
                font_path = "C:/Windows/Fonts/msjh.ttc"
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                    font_name = 'ChineseFont'
                else:
                    font_name = 'Helvetica'
            except:
                font_name = 'Helvetica'
            
            # 創建PDF文件
            doc = SimpleDocTemplate(filename, pagesize=A4, 
                                topMargin=0.4*inch, bottomMargin=0.25*inch,
                                leftMargin=0.25*inch, rightMargin=0.25*inch)
            
            # 樣式定義
            title_style = ParagraphStyle(
                'MainTitle',
                fontName=font_name,
                fontSize=14,
                alignment=TA_CENTER,
                spaceAfter=15,
                textColor=colors.HexColor('#2c3e50')
            )
            
            file_header_style = ParagraphStyle(
                'FileHeader',
                fontName=font_name,
                fontSize=10,
                alignment=TA_CENTER,
                textColor=colors.white
            )
            
            info_style = ParagraphStyle(
                'InfoStyle',
                fontName=font_name,
                fontSize=7,
                alignment=TA_LEFT,
                spaceAfter=1,
                textColor=colors.HexColor('#34495e')
            )
            
            filename_style = ParagraphStyle(
                'FilenameStyle',
                fontName=font_name,
                fontSize=7,
                alignment=TA_CENTER,
                textColor=colors.white
            )
            
            # 建立PDF內容
            story = []
            
            # 主標題
            story.append(Paragraph(f"上傳資料統整表 - {year_display}", title_style))
            story.append(Spacer(1, 10))
            
            # 處理每個檔案
            for file_index, (file_number, images) in enumerate(selected_data):
                if file_index > 0:
                    story.append(PageBreak())
                
                # 取得基本資訊
                display_file_number = file_number
                if images:
                    first_image = images[0]
                    data_year = first_image.get('data_year', '')
                    if data_year:
                        display_file_number = f"{data_year}-{file_number}"
                
                detector_name = "未知"
                if images:
                    try:
                        detector_name = self.get_folder_creator(images[0])
                    except:
                        detector_name = "未知"
                
                # 處理圖片（每頁2張）
                for i in range(0, len(images), 2):
                    if i > 0:
                        story.append(PageBreak())
                    
                    # 頂部標頭
                    page_info = f"檔號：{display_file_number} | 第 {i+1}-{min(i+2, len(images))} 張 / 共 {len(images)} 張"
                    header_data = [[Paragraph(page_info, file_header_style)]]
                    
                    header_table = Table(header_data, colWidths=[8.0*inch])
                    header_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#34495e')),
                        ('TOPPADDING', (0, 0), (-1, -1), 4),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                        ('GRID', (0, 0), (-1, -1), 0, colors.white),
                    ]))
                    story.append(header_table)
                    story.append(Spacer(1, 5))
                    
                    # 取得當前頁面的圖片（最多2張）
                    current_images = images[i:i+2]
                    
                    # 建立主內容區域
                    content_rows = []
                    
                    # 第一行：圖片
                    image_row = []
                    for img_info in current_images:
                        image_path = img_info.get('image_path', '')
                        if image_path and os.path.exists(image_path):
                            try:
                                img = RLImage(image_path, width=2.4*inch, height=1.8*inch)
                                image_row.append(img)
                            except:
                                error_msg = Paragraph("圖片載入失敗", info_style)
                                image_row.append(error_msg)
                        else:
                            error_msg = Paragraph("圖片不存在", info_style)
                            image_row.append(error_msg)
                    
                    # 如果只有一張圖片，填充空白
                    if len(image_row) == 1:
                        image_row.append('')
                    
                    content_rows.append(image_row)
                    content_rows.append([Spacer(1, 4), Spacer(1, 4)])
                    
                    # 第二行：圖片資訊
                    info_row = []
                    for img_info in current_images:
                        # 建立資訊文字 - 欄位名稱和數據在同一行
                        info_lines = [
                            f"上傳年：{img_info.get('upload_year', '未知')} | 資料年：{img_info.get('data_year', '未知')}",
                            f"分類：{img_info.get('category', '未知')} | 案：{img_info.get('case', '未知')} | 卷：{img_info.get('volume', '未知')}",
                            f"目：{img_info.get('item', '未知')} | 片：{img_info.get('sheet_num', '未知')}/{img_info.get('sheet_total', '未知')}",
                            f"圖片編號：{img_info.get('image_num', '未知')}"
                        ]
                        
                        info_text = "<br/>".join(info_lines)
                        info_para = Paragraph(info_text, info_style)
                        
                        # 建立資訊區域表格
                        info_cell_data = [[info_para]]
                        info_cell = Table(info_cell_data, colWidths=[2.4*inch], rowHeights=[1.0*inch])
                        info_cell.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#ecf0f1')),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                            ('LEFTPADDING', (0, 0), (-1, -1), 6),
                            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                            ('TOPPADDING', (0, 0), (-1, -1), 4),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#bdc3c7')),
                        ]))
                        
                        info_row.append(info_cell)
                    
                    # 如果只有一張圖片，填充空白
                    if len(info_row) == 1:
                        info_row.append('')
                    
                    content_rows.append(info_row)
                    content_rows.append([Spacer(1, 4), Spacer(1, 4)])
                    
                    # 第三行：檔名
                    filename_row = []
                    for img_info in current_images:
                        filename = img_info.get('filename', '未知')
                        # 如果檔名太長，只顯示前32個字符
                        if len(filename) > 35:
                            filename = filename[:32] + "..."
                        
                        filename_para = Paragraph(f"檔名：{filename}", filename_style)
                        filename_row.append(filename_para)
                    
                    if len(filename_row) == 1:
                        filename_row.append('')
                    
                    content_rows.append(filename_row)
                    
                    # 建立主內容表格
                    main_content_table = Table(content_rows, 
                                            colWidths=[2.5*inch, 2.5*inch],
                                            rowHeights=[1.9*inch, 0.1*inch, 1.1*inch, 0.1*inch, 0.4*inch])
                    main_content_table.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (1, 0), 'MIDDLE'),
                        ('VALIGN', (0, 2), (1, 2), 'TOP'),
                        ('VALIGN', (0, 4), (1, 4), 'MIDDLE'),
                        ('LEFTPADDING', (0, 0), (-1, -1), 3),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                        ('GRID', (0, 0), (-1, -1), 0, colors.white),
                    ]))
                    
                    story.append(main_content_table)
                    story.append(Spacer(1, 5))
                    
                    # 底部區域
                    bottom_info = f"檢測人：{detector_name} | 頁碼：{i//2 + 1}"
                    bottom_data = [[Paragraph(bottom_info, ParagraphStyle(
                        'BottomInfo',
                        fontName=font_name,
                        fontSize=9,
                        alignment=TA_CENTER,
                        textColor=colors.white
                    ))]]
                    
                    bottom_table = Table(bottom_data, colWidths=[8.0*inch])
                    bottom_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#2c3e50')),
                        ('TOPPADDING', (0, 0), (-1, -1), 4),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                        ('GRID', (0, 0), (-1, -1), 0, colors.white),
                    ]))
                    
                    story.append(bottom_table)

            # 生成PDF
            doc.build(story)
            
        except Exception as e:
            raise Exception(f"PDF生成失敗：{str(e)}")
    
    # ==================== 年份查閱功能結束 ====================        
            
            

def main():
    """主函數 - 加強編碼處理"""
    try:
        # 再次確保編碼設置正確
        setup_encoding()
        
        # 在try塊的開始就引入tkinter
        import tkinter as tk
        
        root = tk.Tk()
        app = FileSearchApp(root)
        root.mainloop()
        
    except UnicodeEncodeError as e:
        print(f"Unicode編碼錯誤: {e}")
        try:
            # 嘗試使用更安全的方式啟動
            import tkinter as tk  # 這行可以移除，因為上面已經引入了
            root = tk.Tk()
            root.title("編碼錯誤")
            root.geometry("400x200")
            
            error_msg = f"程式遇到編碼問題：\n{str(e)}\n\n請確保：\n1. 檔案路徑不包含特殊字符\n2. 系統支持UTF-8編碼"
            tk.Label(root, text=error_msg, wraplength=350, justify='left').pack(pady=20)
            
            tk.Button(root, text="關閉", command=root.destroy).pack(pady=10)
            root.mainloop()
        except:
            print("無法創建錯誤對話框")
            
    except Exception as e:
        print(f"程式啟動錯誤: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
