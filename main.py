import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk  
from datetime import datetime
import getpass
import subprocess
import sys

try:
    import ctypes
    from ctypes import wintypes
    WINDOWS_API_AVAILABLE = True
except ImportError:
    WINDOWS_API_AVAILABLE = False

# 嘗試匯入PIL
PIL_AVAILABLE = False
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
    print("PIL/Pillow 已成功載入")
except ImportError as e:
    print(f"PIL/Pillow 載入失敗: {e}")

def hide_console():
    """隱藏控制台窗口"""
    try:
        if os.name == 'nt' and getattr(sys, 'frozen', False):  # 僅在Windows的打包環境中執行
            import ctypes
            from ctypes import wintypes
            # 獲取控制台窗口句柄
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            
            # 獲取當前控制台窗口
            console_window = kernel32.GetConsoleWindow()
            if console_window != 0:
                # 隱藏控制台窗口 (SW_HIDE = 0)
                user32.ShowWindow(console_window, 0)
                
    except Exception as e:
        # 如果隱藏失敗，繼續執行程式（不影響主要功能）
        pass

# 在主程式開始前調用
hide_console()


class DataIntegratorApp:
    def __init__(self, master):
        self.master = master
        master.title("光碟內容整合工具")
        master.geometry("1000x950")

        self.setup_keyboard_shortcuts()

        # 初始化變數
        self.image_entries = []
        self.image_paths = []
        self.cd_content_paths = []
        self.selected_files = []
        
        # 命名變數
        self.year = tk.StringVar()
        self.category = tk.StringVar()
        self.case = tk.StringVar()
        self.volume = tk.StringVar()
        self.item = tk.StringVar()
        self.sheet_start = tk.StringVar()
        self.sheet_end = tk.StringVar()
        self.user_name = tk.StringVar()
        self.current_user_code = getpass.getuser()  # 使用者代碼
        self.current_user_display = self.get_display_name()  # 顯示名稱（中文名）
        self.user_name.set(self.current_user_display)  # 設定預設值為中文名
        self.fixed_output_path = tk.StringVar()  # 固定輸出路徑
        self.use_fixed_path = tk.BooleanVar()    # 是否使用固定路徑
        self.load_settings()  # 載入設定
    
        self.current_preview_path = None
        self.preview_photo = None

        self.setup_ui()

    def setup_keyboard_shortcuts(self):
        """設定鍵盤快捷鍵"""
        # 綁定 Ctrl+V 到主視窗
        self.master.bind_all('<Control-v>', self.paste_image_from_clipboard)
        self.master.bind_all('<Control-V>', self.paste_image_from_clipboard)  # 大寫也支援

    def switch_to_cd_searcher(self):
        """切換到 cd-searcher 程式"""
        try:
            import subprocess
            import sys
            import os
            
            # 確認是否要切換程式
            if messagebox.askyesno("切換程式", "確定要切換到「光碟搜尋程式」嗎？\n當前程式將會關閉。"):
                
                # 判斷是否為打包後的 EXE 環境
                if getattr(sys, 'frozen', False):
                    # 打包後的 EXE 環境
                    current_dir = os.path.dirname(sys.executable)
                    searcher_path = os.path.join(current_dir, "光碟搜尋程式.exe")
                    
                    if os.path.exists(searcher_path):
                        # 啟動搜尋程式 EXE
                        subprocess.Popen([searcher_path])
                        self.master.after(500, self.close_application)
                    else:
                        messagebox.showerror("錯誤", f"找不到光碟搜尋程式！\n請確認 光碟搜尋程式.exe 與本程式在同一資料夾中。")
                        
                else:
                    # 開發環境（原本的邏輯）
                    searcher_path = os.path.join(os.path.dirname(__file__), "cd-searcher.py")
                    
                    if os.path.exists(searcher_path):
                        subprocess.Popen([sys.executable, searcher_path])
                        self.master.after(500, self.close_application)
                    else:
                        # 讓用戶選擇檔案位置
                        searcher_file = filedialog.askopenfilename(
                            title="請選擇 cd-searcher 程式位置",
                            filetypes=[
                                ("Python 檔案", "*.py"),
                                ("執行檔", "*.exe"),
                                ("所有檔案", "*.*")
                            ]
                        )
                        
                        if searcher_file:
                            if searcher_file.endswith('.py'):
                                subprocess.Popen([sys.executable, searcher_file])
                            else:
                                subprocess.Popen([searcher_file])
                            self.master.after(500, self.close_application)
                        else:
                            messagebox.showinfo("取消", "未選擇程式，繼續使用當前程式")
                            
        except Exception as e:
            messagebox.showerror("錯誤", f"切換程式時發生錯誤: {str(e)}")

    def close_application(self):
        """關閉當前應用程式"""
        try:
            self.master.quit()
            self.master.destroy()
        except Exception as e:
            print(f"關閉程式時發生錯誤: {e}")
            # 強制退出
            import sys
            sys.exit(0)

    def paste_image_from_clipboard(self, event=None):
        """從剪貼簿貼上圖片 - 改進版本"""
        try:
            import tempfile
            import io
            from PIL import ImageGrab

            self.ensure_lists_consistency()
            
            if not PIL_AVAILABLE:
                messagebox.showerror("錯誤", "PIL/Pillow 未安裝，無法使用剪貼簿功能")
                return
            
            # 從剪貼簿獲取圖片
            clipboard_image = ImageGrab.grabclipboard()
            
            if clipboard_image is None:
                messagebox.showinfo("提示", "剪貼簿中沒有圖片內容")
                return
            
            if not hasattr(clipboard_image, 'save'):
                messagebox.showinfo("提示", "剪貼簿中的內容不是有效的圖片格式")
                return
            
            # 找到第一個空的圖片欄位
            empty_index = -1
            for i in range(len(self.image_entries)):
                # 確保索引安全
                if i >= len(self.image_paths):
                    self.image_paths.append("")
                
                if not self.image_paths[i] or not self.image_paths[i].strip():
                    empty_index = i
                    break
            
            if empty_index == -1:
                # 如果沒有空欄位，新增一個
                if len(self.image_entries) < 10:
                    self.add_image_entry_field()
                    empty_index = len(self.image_paths) - 1
                else:
                    messagebox.showwarning("警告", "所有圖片欄位都已填滿，無法貼上更多圖片")
                    return
            
            # 創建臨時檔案保存剪貼簿圖片 - 使用改進的檔名格式
            temp_dir = tempfile.gettempdir()
            roc_year = self.get_roc_year()
            
            # 產生基於資料夾名稱的檔名
            folder_name = self.generate_folder_name()
            if folder_name:
                temp_filename = f"{roc_year}-{folder_name}.png"
            else:
                # 如果資料夾名稱未完整填寫，給出警告並返回
                messagebox.showwarning("警告", "請先完整填寫所有命名欄位後再使用截圖功能！")
                return
            temp_filepath = os.path.join(temp_dir, temp_filename)
            
            # 直接從編號1開始，找到第一個不存在的檔名
            counter = 1
            name_without_ext = temp_filename[:-4]  # 移除 .png
            temp_filename = f"{name_without_ext}-{counter}.png"
            temp_filepath = os.path.join(temp_dir, temp_filename)
            
            while os.path.exists(temp_filepath):
                counter += 1
                temp_filename = f"{name_without_ext}-{counter}.png"
                temp_filepath = os.path.join(temp_dir, temp_filename)
            
            # 保存圖片到臨時檔案
            clipboard_image.save(temp_filepath, 'PNG')
            
            # 更新對應的圖片欄位
            self.image_entries[empty_index]['var'].set(temp_filepath)
            
            # 確保 image_paths 有足夠空間
            while len(self.image_paths) <= empty_index:
                self.image_paths.append("")
            
            self.image_paths[empty_index] = temp_filepath
            
            # 更新顯示狀態
            self.update_image_display_status(empty_index)
            
            # 自動預覽貼上的圖片
            self.preview_image_from_entry(empty_index)
            
            messagebox.showinfo("成功", f"✅ 已將剪貼簿圖片貼上到第 {empty_index + 1} 個圖片欄位")
            
        except IndexError as e:
            messagebox.showerror("錯誤", f"索引錯誤，請重新嘗試: {str(e)}")
            print(f"IndexError in paste_image_from_clipboard: {e}")
        except ImportError:
            messagebox.showerror("錯誤", "PIL/Pillow 未安裝，無法使用剪貼簿功能")
        except Exception as e:
            messagebox.showerror("錯誤", f"貼上圖片時發生錯誤: {str(e)}")
            print(f"Exception in paste_image_from_clipboard: {e}")

    def trigger_screen_capture(self):
        """觸發螢幕截圖功能"""
        try:
            if os.name == 'nt':  # Windows 系統
                # 方案1: 使用 Windows 內建的截圖工具 (Shift + Win + S)
                # 發送按鍵組合
                try:
                    import win32api
                    import win32con
                    
                    # 模擬 Win + Shift + S 按鍵
                    win32api.keybd_event(win32con.VK_LWIN, 0, 0, 0)  # Win 鍵按下
                    win32api.keybd_event(win32con.VK_LSHIFT, 0, 0, 0)  # Shift 鍵按下
                    win32api.keybd_event(ord('S'), 0, 0, 0)  # S 鍵按下
                    
                    # 釋放按鍵
                    win32api.keybd_event(ord('S'), 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_LSHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_LWIN, 0, win32con.KEYEVENTF_KEYUP, 0)
                    
                except ImportError:
                    # 方案2: 使用命令列啟動截圖工具
                    subprocess.run(['powershell', '-Command', 
                                'Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait("^+{ESC}")'], 
                                shell=True)
                    
            elif sys.platform == 'darwin':  # macOS
                # 使用 macOS 的截圖快捷鍵 Cmd + Shift + 4
                subprocess.run(['osascript', '-e', 
                            'tell application "System Events" to keystroke "4" using {command down, shift down}'])
                
            else:  # Linux
                # 嘗試使用常見的截圖工具
                screenshot_tools = ['gnome-screenshot', 'spectacle', 'scrot', 'flameshot']
                for tool in screenshot_tools:
                    try:
                        if tool == 'gnome-screenshot':
                            subprocess.run([tool, '-a'], check=True)  # -a 為選擇區域
                        elif tool == 'flameshot':
                            subprocess.run([tool, 'gui'], check=True)
                        elif tool == 'spectacle':
                            subprocess.run([tool, '-r'], check=True)  # -r 為矩形選擇
                        elif tool == 'scrot':
                            subprocess.run([tool, '-s'], check=True)  # -s 為選擇模式
                        break
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        continue
                else:
                    messagebox.showinfo("提示", "未找到可用的截圖工具")
                    return
                    
            
        except Exception as e:
            messagebox.showerror("錯誤", f"啟動螢幕截圖失敗: {str(e)}")


    def get_display_name(self):
        """取得使用者的顯示名稱（中文名稱）"""
        try:
            if WINDOWS_API_AVAILABLE and os.name == 'nt':
                # 使用 Windows API 取得完整使用者名稱
                GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
                NameDisplay = 3
                
                size = wintypes.DWORD(0)
                GetUserNameEx(NameDisplay, None, ctypes.byref(size))
                
                name_buffer = ctypes.create_unicode_buffer(size.value)
                GetUserNameEx(NameDisplay, name_buffer, ctypes.byref(size))
                
                full_name = name_buffer.value
                if full_name:
                    # 如果取得完整名稱，只取名字部分（去除域名）
                    if '\\' in full_name:
                        return full_name.split('\\')[-1]
                    else:
                        return full_name
                
            # 備用方法：嘗試從環境變數獲取
            display_names = [
                os.environ.get('USERNAME', ''),  # Windows 使用者名稱
                os.environ.get('USER', ''),      # Unix 系統
                getpass.getuser()                # 最後備案
            ]
            
            for name in display_names:
                if name and name.strip():
                    return name.strip()
            
            return getpass.getuser()  # 最終備案
            
        except Exception as e:
            print(f"無法取得顯示名稱，使用預設: {e}")
            return getpass.getuser()
        
    def save_settings_silent(self):
        """靜默儲存設定檔（不顯示彈窗）"""
        try:
            import json
            settings_file = os.path.join(os.path.expanduser("~"), "cd_integrator_settings.json")
            settings = {
                'fixed_output_path': self.fixed_output_path.get(),
                'use_fixed_path': self.use_fixed_path.get()
            }
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            # 移除彈窗提示，改為控制台輸出
            print("設定已靜默儲存")
        except Exception as e:
            print(f"靜默儲存設定失敗: {str(e)}")
            
    def load_settings(self):
        """載入設定檔 - 強制重置為D槽預設路徑"""
        try:
            import json
            settings_file = os.path.join(os.path.expanduser("~"), "cd_integrator_settings.json")
            
            # 強制設定D槽路徑為固定預設路徑
            default_d_drive_path = "D:\\光碟檢測及備份"
            
            # 嘗試創建預設資料夾
            try:
                if os.path.exists("D:\\"):  # 檢查D槽是否存在
                    os.makedirs(default_d_drive_path, exist_ok=True)
                    print(f"成功確保資料夾存在: {default_d_drive_path}")
                    # 無論之前設定為何，強制使用D槽路徑
                    self.fixed_output_path.set(default_d_drive_path)
                else:
                    print("D槽不存在，將使用桌面備用路徑")
                    default_d_drive_path = os.path.join(os.path.expanduser("~"), "Desktop", "光碟檢測及備份")
                    os.makedirs(default_d_drive_path, exist_ok=True)
                    self.fixed_output_path.set(default_d_drive_path)
            except Exception as e:
                print(f"創建預設資料夾失敗: {e}")
                default_d_drive_path = os.path.expanduser("~")
                self.fixed_output_path.set(default_d_drive_path)
            
            # 讀取設定檔但只保留 use_fixed_path 設定，路徑強制重置
            if os.path.exists(settings_file):
                try:
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        settings = json.load(f)
                        # 只保留是否使用固定路徑的設定，預設為True
                        self.use_fixed_path.set(settings.get('use_fixed_path', True))
                except:
                    self.use_fixed_path.set(True)
            else:
                # 預設啟用固定路徑功能
                self.use_fixed_path.set(True)
            
            # 靜默儲存新的設定（不顯示彈窗）
            self.save_settings_silent()
            print(f"已強制重置固定路徑為: {self.fixed_output_path.get()}")
                
        except Exception as e:
            print(f"載入設定失敗: {e}")
            # 發生錯誤時的備用設定
            try:
                backup_path = "D:\\光碟檢測及備份"
                if os.path.exists("D:\\"):
                    os.makedirs(backup_path, exist_ok=True)
                    self.fixed_output_path.set(backup_path)
                else:
                    backup_path = os.path.join(os.path.expanduser("~"), "Desktop", "光碟檢測及備份")
                    os.makedirs(backup_path, exist_ok=True)
                    self.fixed_output_path.set(backup_path)
            except:
                self.fixed_output_path.set(os.path.expanduser("~"))
            self.use_fixed_path.set(True)



    def ensure_default_folder(self):
        """確保預設資料夾存在並返回路徑"""
        try:
            # 首選：D槽的光碟檢測及備份資料夾
            if os.path.exists("D:\\"):
                default_path = "D:\\光碟檢測及備份"
                os.makedirs(default_path, exist_ok=True)
                print(f"✅ 成功創建/確認D槽預設資料夾: {default_path}")
                return default_path
            else:
                print("⚠️ D槽不存在，使用桌面作為備用位置")
                # 備用：桌面的光碟檢測及備份資料夾
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "光碟檢測及備份")
                os.makedirs(desktop_path, exist_ok=True)
                print(f"✅ 成功創建/確認桌面預設資料夾: {desktop_path}")
                return desktop_path
        except Exception as e:
            print(f"❌ 創建預設資料夾失敗: {e}")
            # 最終備用：用戶主目錄
            return os.path.expanduser("~")


    def save_settings(self):
        """儲存設定檔（帶彈窗提示）"""
        try:
            import json
            settings_file = os.path.join(os.path.expanduser("~"), "cd_integrator_settings.json")
            settings = {
                'fixed_output_path': self.fixed_output_path.get(),
                'use_fixed_path': self.use_fixed_path.get()
            }
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "設定已儲存")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存設定失敗: {str(e)}")

    def open_settings_window(self):
        """開啟設定視窗"""
        settings_window = tk.Toplevel(self.master)
        settings_window.title("設定")
        settings_window.geometry("600x420")
        settings_window.resizable(False, False)
        
        # 讓設定視窗在最前面
        settings_window.transient(self.master)
        settings_window.grab_set()
        
        # 主框架
        main_frame = tk.Frame(settings_window, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題
        title_label = tk.Label(main_frame, text="🔧 應用程式設定", 
                            font=("Microsoft YaHei", 14, "bold"), fg="darkblue")
        title_label.pack(pady=(0, 20))
        
        # 固定輸出路徑設定區域
        path_frame = tk.LabelFrame(main_frame, text="固定輸出路徑設定", 
                                font=("Microsoft YaHei", 12, "bold"), padx=10, pady=10)
        path_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 啟用固定路徑選項
        enable_frame = tk.Frame(path_frame)
        enable_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Checkbutton(enable_frame, text="使用固定輸出路徑（勾選後整合時不再詢問輸出位置）", 
                    variable=self.use_fixed_path, font=("Microsoft YaHei", 9)).pack(anchor='w')
        
        # 路徑選擇區域
        path_select_frame = tk.Frame(path_frame)
        path_select_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(path_select_frame, text="固定路徑:", font=("Microsoft YaHei", 9)).pack(anchor='w')
        
        path_input_frame = tk.Frame(path_select_frame)
        path_input_frame.pack(fill=tk.X, pady=(5, 0))
        
        path_entry = tk.Entry(path_input_frame, textvariable=self.fixed_output_path, 
                            font=("Microsoft YaHei", 9), state='readonly', bg="#F5F5F5")
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        tk.Button(path_input_frame, text="瀏覽...", 
                command=self.browse_fixed_output_path, 
                bg="#4CAF50", fg="white", font=("Microsoft YaHei", 9)).pack(side=tk.RIGHT)
        
        # 當前設定顯示
        current_frame = tk.LabelFrame(main_frame, text="當前設定狀態", 
                                    font=("Microsoft YaHei", 10, "bold"), padx=10, pady=10)
        current_frame.pack(fill=tk.X, pady=(0, 15))
        
        def update_status_display():
            if self.use_fixed_path.get() and self.fixed_output_path.get():
                status_text = f"✅ 啟用固定路徑\n📁 {self.fixed_output_path.get()}"
                status_color = "darkgreen"
            elif self.use_fixed_path.get():
                status_text = "⚠️ 已啟用固定路徑，但路徑未設定"
                status_color = "orange"
            else:
                status_text = "ℹ️ 未啟用固定路徑，整合時將詢問輸出位置"
                status_color = "gray"
            
            status_label.config(text=status_text, fg=status_color)
        
        status_label = tk.Label(current_frame, text="", font=("Microsoft YaHei", 9), justify=tk.LEFT)
        status_label.pack(anchor='w')
        
        # 綁定變數變化事件
        self.use_fixed_path.trace('w', lambda *args: update_status_display())
        self.fixed_output_path.trace('w', lambda *args: update_status_display())
        
        # 初始化顯示
        update_status_display()
        
        # 按鈕區域
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(20, 0))

        # 儲存設定按鈕
        save_btn = tk.Button(btn_frame, text="💾 儲存設定", 
                            command=lambda: [self.save_settings(), settings_window.destroy()],
                            bg="#2196F3", fg="white", font=("Microsoft YaHei", 8, "bold"),
                            width=12, height=1)  # 增加固定寬度和高度
        save_btn.pack(side=tk.RIGHT, padx=(10, 0))

        # 取消按鈕
        cancel_btn = tk.Button(btn_frame, text="❌ 取消", 
                            command=settings_window.destroy,
                            bg="#757575", fg="white", font=("Microsoft YaHei", 8),
                            width=8, height=1)  # 增加固定寬度和高度
        cancel_btn.pack(side=tk.RIGHT)

    def browse_fixed_output_path(self):
        """瀏覽固定輸出路徑"""
        # 如果當前路徑存在則用當前路徑，否則用D槽預設路徑作為初始目錄
        current_path = self.fixed_output_path.get()
        if current_path and os.path.exists(current_path):
            initial_dir = current_path
        elif os.path.exists("D:\\光碟檢測及備份"):
            initial_dir = "D:\\光碟檢測及備份"
        elif os.path.exists("D:\\"):
            initial_dir = "D:\\"
        else:
            initial_dir = os.path.expanduser("~")
        
        folder = filedialog.askdirectory(
            title="選擇固定輸出路徑",
            initialdir=initial_dir
        )
        if folder:
            self.fixed_output_path.set(folder)
        
    def setup_ui(self):
        # === 創建主滾動區域 ===
        # 創建主Canvas和滾動條
        main_canvas = tk.Canvas(self.master, highlightthickness=0)
        main_scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=main_canvas.yview)
        
        # 創建可滾動的主框架
        self.scrollable_frame = tk.Frame(main_canvas, bg="#F5F6F7")
        
        # 配置滾動
        main_canvas.configure(yscrollcommand=main_scrollbar.set)
        
        # 布局滾動元件
        main_canvas.pack(side="left", fill="both", expand=True)
        main_scrollbar.pack(side="right", fill="y")
        
        # 將可滾動框架加入Canvas
        canvas_frame = main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # 綁定滾動事件
        def configure_scroll_region(event=None):
            main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        
        def configure_canvas_width(event=None):
            canvas_width = main_canvas.winfo_width()
            main_canvas.itemconfig(canvas_frame, width=canvas_width)
        
        self.scrollable_frame.bind("<Configure>", configure_scroll_region)
        main_canvas.bind("<Configure>", configure_canvas_width)
        
        # 綁定滑鼠滾輪事件
        def on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", on_mousewheel)  # Windows
        main_canvas.bind_all("<Button-4>", lambda e: main_canvas.yview_scroll(-1, "units"))  # Linux
        main_canvas.bind_all("<Button-5>", lambda e: main_canvas.yview_scroll(1, "units"))   # Linux

        # === 頂部標題和按鈕區域（合併在同一行）===
        header_frame = tk.Frame(self.scrollable_frame)
        header_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

        # 左側容器：標題 + 切換搜尋按鈕
        left_container = tk.Frame(header_frame)
        left_container.pack(side=tk.LEFT)

        # 標題
        title_label = tk.Label(left_container, text="光碟內容整合工具", 
                            font=("Microsoft YaHei", 18, "bold"))
        title_label.pack(side=tk.LEFT)

        # 切換搜尋按鈕（緊貼標題右邊）
        searcher_btn = tk.Button(left_container, text="🔄切換搜尋", 
                                command=self.switch_to_cd_searcher,
                                bg="#E09F11", fg="white", 
                                font=("Microsoft YaHei", 10, "bold"),
                                relief=tk.RAISED, bd=2)
        searcher_btn.pack(side=tk.LEFT, padx=(20, 0))

        # 右側容器：功能按鈕
        right_container = tk.Frame(header_frame)
        right_container.pack(side=tk.RIGHT)

        # 螢幕截圖按鈕
        screenshot_btn = tk.Button(right_container, text="⬚快速截圖", 
                                command=self.trigger_screen_capture,
                                bg="#3898D8", fg="white", 
                                font=("Microsoft YaHei", 10, "bold"),
                                relief=tk.RAISED, bd=2)
        screenshot_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 設定按鈕
        settings_btn = tk.Button(right_container, text="路徑設定", 
                                command=self.open_settings_window,
                                bg="#607D8B", fg="white", 
                                font=("Microsoft YaHei", 10, "bold"),
                                relief=tk.RAISED, bd=2)
        settings_btn.pack(side=tk.LEFT)

        # 主分隔線
        separator = tk.Frame(self.scrollable_frame, height=2, bg="lightgray")
        separator.pack(fill=tk.X, padx=10, pady=10)
        
        # === 命名設定區域 ===
        naming_section = tk.LabelFrame(self.scrollable_frame, text="1. 資料夾命名設定", 
                              font=("Microsoft YaHei", 12, "bold"),
                              bg="#F8F9FA", relief=tk.GROOVE, bd=2, padx=15, pady=10)
        naming_section.pack(fill=tk.X, padx=10, pady=5)

        # 命名輸入區 - 改為三欄式布局
        entry_frame = tk.Frame(naming_section)
        entry_frame.pack(pady=10, fill=tk.X)

        # 第一欄：年度、分類號
        col1_frame = tk.Frame(entry_frame)
        col1_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        tk.Label(col1_frame, font=("Microsoft YaHei", 9, "bold"), fg="darkblue").pack(anchor='w')

        # 年度
        row1_col1 = tk.Frame(col1_frame)
        row1_col1.pack(fill=tk.X, pady=2)
        tk.Label(row1_col1, text="年度:", width=8, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row1_col1, textvariable=self.year, width=15, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 分類號
        row2_col1 = tk.Frame(col1_frame)
        row2_col1.pack(fill=tk.X, pady=2)
        tk.Label(row2_col1, text="分類號:", width=8, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row2_col1, textvariable=self.category, width=15, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 第二欄：案、卷、目
        col2_frame = tk.Frame(entry_frame)
        col2_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        tk.Label(col2_frame, font=("Microsoft YaHei", 9, "bold"), fg="darkgreen").pack(anchor='w')

        # 案
        row1_col2 = tk.Frame(col2_frame)
        row1_col2.pack(fill=tk.X, pady=2)
        tk.Label(row1_col2, text=" 案:", width=5, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row1_col2, textvariable=self.case, width=15, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 卷
        row2_col2 = tk.Frame(col2_frame)
        row2_col2.pack(fill=tk.X, pady=2)
        tk.Label(row2_col2, text=" 卷:", width=5, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row2_col2, textvariable=self.volume, width=15, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 目
        row3_col2 = tk.Frame(col2_frame)
        row3_col2.pack(fill=tk.X, pady=2)
        tk.Label(row3_col2, text=" 目:", width=5, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row3_col2, textvariable=self.item, width=15, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 第三欄：片(總片數)、片(編號)
        col3_frame = tk.Frame(entry_frame)
        col3_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tk.Label(col3_frame, font=("Microsoft YaHei", 9, "bold"), fg="darkorange").pack(anchor='w')

        # 片(總片數)
        row1_col3 = tk.Frame(col3_frame)
        row1_col3.pack(fill=tk.X, pady=2)
        tk.Label(row1_col3, text="片(總片數):", width=12, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row1_col3, textvariable=self.sheet_start, width=15, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 片(編號)
        row2_col3 = tk.Frame(col3_frame)
        row2_col3.pack(fill=tk.X, pady=2)
        tk.Label(row2_col3, text="片(編號):", width=12, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row2_col3, textvariable=self.sheet_end, width=15, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)
                

        # 資料夾名稱預覽
        self.folder_name_label = tk.Label(naming_section, text="📁 資料夾名稱: (請填寫完整資訊)",
                                         font=("Microsoft YaHei", 10), fg="blue")
        self.folder_name_label.pack(pady=10)

        # 綁定即時預覽
        for var in [self.year, self.category, self.case, self.volume, self.item, 
           self.sheet_start, self.sheet_end, self.user_name]:
            var.trace('w', self.update_folder_preview)


        # === 圖片選擇區域 ===
        img_section = tk.LabelFrame(self.scrollable_frame, text="2. 圖片選擇", 
                           font=("Microsoft YaHei", 12, "bold"),
                           bg="#F8F9FA", relief=tk.GROOVE, bd=2, padx=15, pady=10)
        img_section.pack(fill=tk.X, padx=10, pady=5)

        # 主框架 - 建立上下布局
        img_main_frame = tk.Frame(img_section)
        img_main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 上方：輸入欄位區域
        self.img_input_frame = tk.Frame(img_main_frame)
        self.img_input_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 下方：預覽和按鈕區域
        img_bottom_frame = tk.Frame(img_main_frame)
        img_bottom_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 初始建立1個圖片輸入欄位
        for i in range(1):
            self.add_image_entry_field()
        
        # 圖片預覽區
        preview_section = tk.LabelFrame(img_bottom_frame, text="圖片預覽", 
                               font=("Microsoft YaHei", 10, "bold"),
                               bg="#FFFFFF", relief=tk.FLAT, bd=1, padx=10, pady=5)
        preview_section.pack(fill=tk.BOTH, expand=True)

        # 預覽區域右上角的按鈕框架
        preview_header_frame = tk.Frame(preview_section)
        preview_header_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 左側：預覽控制按鈕
        preview_control_frame = tk.Frame(preview_header_frame)
        preview_control_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Button(preview_control_frame, text="放大預覽", command=self.open_preview_window,
                bg="#6F267C", fg="white", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=3)
        tk.Button(preview_control_frame, text="重設預覽", command=self.reset_preview,
                bg="#607D8B", fg="white", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=3)

        # 右側：操作按鈕（移到預覽區右上角）
        btn_frame1 = tk.Frame(preview_header_frame)
        btn_frame1.pack(side=tk.RIGHT)
        
        tk.Button(btn_frame1, text="+ 新增上傳欄位", command=self.add_image_entry_field,
                 bg="#266628", fg="white", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.RIGHT, padx=3)
        tk.Button(btn_frame1, text="清除所有內容", command=self.clear_all_images,
                 bg="#da3a2f", fg="white", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.RIGHT, padx=3)

        # 圖片資訊標籤
        self.img_info_label = tk.Label(preview_section, text="選擇圖片進行預覽", 
                                    font=("Microsoft YaHei", 9), fg="gray")
        self.img_info_label.pack(pady=2)

        # 預覽區域容器 - 使用Canvas和Scrollbar（高度減少）
        preview_container = tk.Frame(preview_section, height=150)  # 改為150
        preview_container.pack(fill=tk.X, padx=5, pady=5)
        preview_container.pack_propagate(False)

        # 創建Canvas和滾動條
        self.preview_canvas = tk.Canvas(preview_container, bg="white", relief=tk.SUNKEN, bd=2, height=150)  # 改為150
        v_scrollbar = ttk.Scrollbar(preview_container, orient="vertical", command=self.preview_canvas.yview)
        h_scrollbar = ttk.Scrollbar(preview_container, orient="horizontal", command=self.preview_canvas.xview)

        self.preview_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # 布局滾動條和Canvas
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        preview_container.grid_rowconfigure(0, weight=1)
        preview_container.grid_columnconfigure(0, weight=1)

        # 初始化預覽
        self.reset_preview()

        data_section = tk.LabelFrame(self.scrollable_frame, text="3. 資料選擇", 
                            font=("Microsoft YaHei", 12, "bold"),
                            bg="#F8F9FA", relief=tk.GROOVE, bd=2, padx=15, pady=10)
        data_section.pack(fill=tk.X, padx=10, pady=5)

        # 資料夾列表區域
        list_frame2 = tk.Frame(data_section)
        list_frame2.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 上方：標籤和按鈕框架
        header_frame = tk.Frame(list_frame2)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(header_frame, text="已選擇的資料夾和檔案:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, anchor='w')
        
        # 右側按鈕區域
        btn_frame2 = tk.Frame(header_frame)
        btn_frame2.pack(side=tk.RIGHT)
        
        tk.Button(btn_frame2, text="新增備份資料", command=self.select_cd_content,
                 bg="#1D7CCA", fg="white", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.RIGHT, padx=(3, 0))
        tk.Button(btn_frame2, text="清除所有內容", command=self.clear_folders,
                 bg="#b83026", fg="white", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.RIGHT, padx=(3, 3))

        # 列表框區域（高度減少）
        listbox_frame2 = tk.Frame(list_frame2)
        listbox_frame2.pack(fill=tk.BOTH, expand=True)

        self.folder_listbox = tk.Listbox(listbox_frame2, height=4, font=("Microsoft YaHei", 9))  # 改為4
        scrollbar2 = tk.Scrollbar(listbox_frame2, orient=tk.VERTICAL)
        
        self.folder_listbox.config(yscrollcommand=scrollbar2.set)
        scrollbar2.config(command=self.folder_listbox.yview)
        
        self.folder_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar2.pack(side=tk.RIGHT, fill=tk.Y)

        # === 整合按鈕 ===
        tk.Button(self.scrollable_frame, text="🚀 上傳 ", command=self.integrate_files,
                 bg="#BD955A", fg="white", font=("Microsoft YaHei", 14, "bold"), height=1).pack(pady=20)

        # 初始化顯示
        self.update_folder_display()


    def add_image_entry_field(self):
        """新增單一圖片檔案輸入欄位"""
        if len(self.image_entries) >= 10:
            messagebox.showinfo("提示", "最多只能新增10個檔案欄位")
            return
            
        index = len(self.image_entries)
        
        # 主框架
        entry_frame = tk.Frame(self.img_input_frame, relief=tk.FLAT, bd=0, bg="#FFFFFF")
        entry_frame.pack(fill=tk.X, pady=3, padx=5)
        
        # 標籤
        label = tk.Label(entry_frame, text=f"上傳檔案 {index+1}:", width=12, anchor='w', 
                        font=("Microsoft YaHei", 10, "bold"))
        label.pack(side=tk.LEFT, padx=(5, 10))
        
        # 檔案路徑輸入框 - 顯示完整路徑
        path_var = tk.StringVar()
        path_entry = tk.Entry(entry_frame, textvariable=path_var, font=("Microsoft YaHei", 9), 
                             bg="#F5F5F5", relief=tk.SUNKEN, bd=2, 
                             state='readonly', cursor="hand2")
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # 刪除按鈕 - 新增
        delete_btn = tk.Button(entry_frame, text="✕", 
                              command=lambda idx=index: self.delete_image_entry(idx),
                              bg="#a32920", fg="white", font=("Microsoft YaHei", 9, "bold"),
                              width=3, height=1, cursor="hand2")
        delete_btn.pack(side=tk.RIGHT, padx=(0, 3))

        # 貼上按鈕 - 新增
        paste_btn = tk.Button(entry_frame, text="📋", 
                            command=lambda idx=index: self.paste_to_specific_field(idx),
                            bg="#D8870D", fg="white", font=("Microsoft YaHei", 9, "bold"),
                            width=3, height=1, cursor="hand2")
        paste_btn.pack(side=tk.RIGHT, padx=(0, 3))
        
        # 瀏覽按鈕
        browse_btn = tk.Button(entry_frame, text="瀏覽...", 
                              command=lambda idx=index: self.browse_single_image(idx),
                              bg="#389C3B", fg="white", font=("Microsoft YaHei", 9, "bold"),
                              width=8, height=1, cursor="hand2")
        browse_btn.pack(side=tk.RIGHT, padx=(0, 3))
        
        # 預覽按鈕
        preview_btn = tk.Button(entry_frame, text="預覽", 
                               command=lambda idx=index: self.preview_image_from_entry(idx),
                               bg="#1574C2", fg="white", font=("Microsoft YaHei", 9, "bold"),
                               width=6, height=1, cursor="hand2")
        preview_btn.pack(side=tk.RIGHT, padx=(0, 3))
        
        # 儲存參考
        entry_info = {
            'frame': entry_frame,
            'label': label,
            'entry': path_entry,
            'var': path_var,
            'browse_btn': browse_btn,
            'preview_btn': preview_btn,
            'paste_btn': paste_btn,
            'delete_btn': delete_btn  
        }
        
        self.image_entries.append(entry_info)
        self.image_paths.append("")
        
        # 綁定事件
        path_var.trace('w', lambda *args, idx=index: self.on_image_path_change(idx))
        path_entry.bind("<Button-1>", lambda e, idx=index: self.preview_image_from_entry(idx))
        path_entry.bind("<Double-Button-1>", lambda e, idx=index: self.open_image_external(idx))

    def delete_image_entry(self, index):
        """刪除指定的圖片輸入欄位"""
        if index >= len(self.image_entries):
            return
            
        # 如果只剩一個欄位，不允許刪除
        if len(self.image_entries) <= 1:
            messagebox.showinfo("提示", "至少需要保留一個輸入欄位")
            return
        
        # 確認刪除
        if self.image_paths[index]:
            filename = os.path.basename(self.image_paths[index])
            if not messagebox.askyesno("確認刪除", f"確定要刪除這個欄位嗎？\n檔案: {filename}"):
                return
        
        # 銷毀該欄位的UI元件
        self.image_entries[index]['frame'].destroy()
        
        # 從列表中移除
        del self.image_entries[index]
        del self.image_paths[index]
        
        # 重新編號所有欄位
        self.renumber_image_entries()
        
        # 如果刪除的是當前預覽的圖片，重設預覽
        if (self.current_preview_path and 
            index < len(self.image_paths) and 
            self.current_preview_path == self.image_paths[index]):
            self.reset_preview()

    
    def paste_to_specific_field(self, index):
        """貼上圖片到指定欄位 - 改進版本"""
        try:
            import tempfile
            from PIL import ImageGrab
            
            if not PIL_AVAILABLE:
                messagebox.showerror("錯誤", "PIL/Pillow 未安裝，無法使用剪貼簿功能")
                return
            
            if index >= len(self.image_entries):
                messagebox.showerror("錯誤", "無效的欄位索引")
                return
            
            self.ensure_lists_consistency()
            
            # 從剪貼簿獲取圖片
            clipboard_image = ImageGrab.grabclipboard()
            
            if clipboard_image is None:
                messagebox.showinfo("提示", "剪貼簿中沒有圖片內容")
                return
            
            if not hasattr(clipboard_image, 'save'):
                messagebox.showinfo("提示", "剪貼簿中的內容不是有效的圖片格式")
                return
            
            # 創建臨時檔案保存剪貼簿圖片 - 使用改進的檔名格式
            temp_dir = tempfile.gettempdir()
            roc_year = self.get_roc_year()
            
            # 產生基於資料夾名稱的檔名
            folder_name = self.generate_folder_name()
            if folder_name:
                temp_filename = f"{roc_year}-{folder_name}.png"
            else:
                # 如果資料夾名稱未完整填寫，給出警告並返回
                messagebox.showwarning("警告", "請先完整填寫所有命名欄位後再使用截圖功能！")
                return
            temp_filepath = os.path.join(temp_dir, temp_filename)
            
            # 直接從編號1開始，找到第一個不存在的檔名
            counter = 1
            name_without_ext = temp_filename[:-4]  # 移除 .png
            temp_filename = f"{name_without_ext}-{counter}.png"
            temp_filepath = os.path.join(temp_dir, temp_filename)
            
            while os.path.exists(temp_filepath):
                counter += 1
                temp_filename = f"{name_without_ext}-{counter}.png"
                temp_filepath = os.path.join(temp_dir, temp_filename)
            
            # 保存圖片到臨時檔案
            clipboard_image.save(temp_filepath, 'PNG')
            
            # 更新指定的圖片欄位
            self.image_entries[index]['var'].set(temp_filepath)
            
            # 確保 image_paths 有足夠空間
            while len(self.image_paths) <= index:
                self.image_paths.append("")
            
            self.image_paths[index] = temp_filepath
            
            # 更新顯示狀態
            self.update_image_display_status(index)
            
            # 自動預覽貼上的圖片
            self.preview_image_from_entry(index)
            
            messagebox.showinfo("成功", f"✅ 已將剪貼簿圖片貼上到第 {index + 1} 個圖片欄位")
            
        except IndexError as e:
            messagebox.showerror("錯誤", f"索引錯誤，請重新嘗試: {str(e)}")
            print(f"IndexError in paste_to_specific_field: {e}")
        except Exception as e:
            messagebox.showerror("錯誤", f"貼上圖片時發生錯誤: {str(e)}")
            print(f"Exception in paste_to_specific_field: {e}")

    def ensure_lists_consistency(self):
        """確保 image_paths 和 image_entries 列表長度一致"""
        # 確保 image_paths 至少和 image_entries 一樣長
        while len(self.image_paths) < len(self.image_entries):
            self.image_paths.append("")
        
        # 如果 image_paths 比 image_entries 長，截斷多餘部分
        if len(self.image_paths) > len(self.image_entries):
            self.image_paths = self.image_paths[:len(self.image_entries)]

    def generate_screenshot_filename(self):
        """產生截圖檔名 - 基於資料夾名稱格式"""
        parts = [
            self.year.get().strip(),
            self.category.get().strip(),
            self.case.get().strip(),
            self.volume.get().strip(),
            self.item.get().strip(),
            self.sheet_start.get().strip(),
            self.sheet_end.get().strip()
        ]
        
        # 如果所有欄位都有填寫，使用資料夾名稱格式
        if all(parts):
            return "-".join(parts)
        else:
            # 如果資料夾名稱未完整填寫，使用時間戳記格式
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            return f"screenshot_{timestamp}"

    def renumber_image_entries(self):
        """重新編號所有圖片輸入欄位"""
        for i, entry_info in enumerate(self.image_entries):
            entry_info['label'].config(text=f"上傳檔案 {i+1}:")
            
            # 更新按鈕的command中的index
            entry_info['browse_btn'].config(command=lambda idx=i: self.browse_single_image(idx))
            entry_info['preview_btn'].config(command=lambda idx=i: self.preview_image_from_entry(idx))
            entry_info['delete_btn'].config(command=lambda idx=i: self.delete_image_entry(idx))

    def add_image_entry(self):
        """新增圖片檔案輸入欄位"""
        self.add_image_entry_field()

    def browse_single_image(self, index):
        """瀏覽單一圖片檔案"""
        file_path = filedialog.askopenfilename(
            title=f"選擇圖片檔案 {index+1}",
            filetypes=[
                ("所有圖片", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.webp *.ico"),
                ("PNG 檔案", "*.png"),
                ("JPEG 檔案", "*.jpg *.jpeg"),
                ("所有檔案", "*.*")
            ]
        )
        
        if file_path:
            # ✅ 顯示完整路徑在輸入框中
            self.image_entries[index]['var'].set(file_path)
            
            # 確保 image_paths 有足夠空間
            while len(self.image_paths) <= index:
                self.image_paths.append("")
            
            self.image_paths[index] = file_path
            
            # 更新顯示狀態
            self.update_image_display_status(index)
            # 自動預覽選擇的圖片
            self.preview_image_from_entry(index)

    def on_image_path_change(self, index):
        """當圖片路徑改變時觸發"""
        if index < len(self.image_entries):
            path = self.image_entries[index]['var'].get().strip()
            
            # ✅ 修正：確保 image_paths 有足夠的空間
            while len(self.image_paths) <= index:
                self.image_paths.append("")
            
            self.image_paths[index] = path
            self.update_image_display_status(index)

    def update_image_display_status(self, index):
        """更新圖片輸入欄位的顯示狀態"""
        try:
            # 檢查索引是否有效
            if index >= len(self.image_entries) or index >= len(self.image_paths):
                return
                
            # 確保 image_paths 有足夠的空間
            while len(self.image_paths) <= index:
                self.image_paths.append("")
            
            path = self.image_paths[index]
            entry = self.image_entries[index]['entry']
            preview_btn = self.image_entries[index]['preview_btn']
            frame = self.image_entries[index]['frame']
            
            if path and os.path.exists(path):
                # 檔案存在 - 綠色邊框
                frame.config(relief=tk.RAISED, bd=2, bg="#E8F5E8")
                entry.config(bg="#F0F8F0", fg="#2E7D32")
                preview_btn.config(state='normal', bg="#2196F3")
                
                # 如果輸入框內容和實際路徑不同，更新顯示
                if entry.get() != path:
                    self.image_entries[index]['var'].set(path)
                    
            elif path:
                # 路徑無效 - 紅色邊框
                frame.config(relief=tk.RAISED, bd=2, bg="#FFEBEE")
                entry.config(bg="#FFEBEE", fg="#C62828")
                preview_btn.config(state='disabled', bg="#BDBDBD")
            else:
                # 空白 - 預設樣式
                frame.config(relief=tk.RAISED, bd=1, bg="#F5F5F5")
                entry.config(bg="#F5F5F5", fg="black")
                preview_btn.config(state='disabled', bg="#BDBDBD")

        except (IndexError, KeyError, AttributeError) as e:
            print(f"Error in update_image_display_status: {e}")
        

    def preview_image_from_entry(self, index):
        """從特定欄位預覽圖片"""
        if index >= len(self.image_paths):
            return
        
        path = self.image_paths[index]
        if not path or not os.path.exists(path):
            self.img_preview_label.configure(image='', text="檔案不存在或路徑無效")
            return
        
        self.preview_image_by_path(path)

    def open_image_external(self, index):
        """使用外部程式開啟圖片"""
        if index >= len(self.image_paths):
            return
            
        path = self.image_paths[index]
        if not path or not os.path.exists(path):
            messagebox.showerror("錯誤", "檔案不存在")
            return
            
        try:
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # macOS and Linux
                if sys.platform == 'darwin':  # macOS
                    os.system(f'open "{path}"')
                else:  # Linux
                    os.system(f'xdg-open "{path}"')
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟檔案: {str(e)}")


    def preview_image_by_path(self, filepath):
        """根據檔案路徑預覽圖片 - 修正版本"""
        if not PIL_AVAILABLE:
            self.reset_preview()
            filename = os.path.basename(filepath)
            self.img_info_label.config(text=f"PIL未載入，無法預覽圖片 - {filename}", fg="red")
            return

        try:
            # 載入圖片
            img = Image.open(filepath)
            if img.mode != 'RGB':
                img = img.convert('RGB')

            # 取得原始尺寸
            original_width, original_height = img.size
            
            # 取得Canvas的實際大小
            self.preview_canvas.update_idletasks()
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            # 如果Canvas還沒有正確的尺寸，使用預設值
            if canvas_width <= 1:
                canvas_width = 600
            if canvas_height <= 1:
                canvas_height = 300
                
            # 計算縮放比例 - 確保圖片適合Canvas但不會太小
            scale_w = canvas_width / original_width
            scale_h = canvas_height / original_height
            scale = min(scale_w, scale_h, 1.0)  # 不放大，但允許縮小
            
            # 如果圖片很小，確保最小顯示尺寸
            min_scale = min(0.8, 200 / max(original_width, original_height))
            scale = max(scale, min_scale)
            
            # 計算新尺寸
            new_width = int(original_width * scale)
            new_height = int(original_height * scale)
            
            # 縮放圖片
            img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.preview_photo = ImageTk.PhotoImage(img_resized)
            
            # 清除並顯示圖片
            self.preview_canvas.delete("all")
            
            # 將圖片放在Canvas中央
            canvas_center_x = canvas_width // 2
            canvas_center_y = canvas_height // 2
            
            self.preview_canvas.create_image(canvas_center_x, canvas_center_y, 
                                        image=self.preview_photo, anchor="center")
            
            # 更新滾動區域
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
            
            # 更新資訊標籤
            filename = os.path.basename(filepath)
            file_size = self.format_size(os.path.getsize(filepath))
            scale_percent = int(scale * 100)
            
            info_text = f"{filename} | 原始: {original_width}×{original_height} | 顯示: {new_width}×{new_height} | 縮放: {scale_percent}% | 大小: {file_size}"
            self.img_info_label.config(text=info_text, fg="darkgreen")
            
            self.current_preview_path = filepath

        except Exception as e:
            self.reset_preview()
            error_msg = f"預覽失敗: {str(e)}"
            self.img_info_label.config(text=error_msg, fg="red")
            print(f"預覽圖片錯誤: {e}")

    def open_preview_window(self):
        """開啟放大預覽視窗"""
        if not self.current_preview_path or not os.path.exists(self.current_preview_path):
            messagebox.showinfo("提示", "請先選擇要預覽的圖片")
            return
        
        if not PIL_AVAILABLE:
            messagebox.showerror("錯誤", "PIL/Pillow 未安裝，無法開啟預覽視窗")
            return
        
        # 創建新視窗
        preview_window = tk.Toplevel(self.scrollable_frame)
        preview_window.title(f"圖片預覽 - {os.path.basename(self.current_preview_path)}")
        preview_window.geometry("800x600")
    
        try:
            # 載入並顯示原始圖片
            img = Image.open(self.current_preview_path)
            if img.mode != 'RGB':
                img = img.convert('RGB')
                
            original_width, original_height = img.size
            
            # 計算適合視窗的縮放比例
            window_width, window_height = 750, 500
            scale = min(window_width / original_width, window_height / original_height, 1.0)
            
            new_width = int(original_width * scale)
            new_height = int(original_height * scale)
            
            img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img_resized)
            
            # 創建標籤顯示圖片
            img_label = tk.Label(preview_window, image=photo)
            img_label.image = photo  # 保持引用
            img_label.pack(expand=True)
            
            # 顯示圖片資訊
            info_text = f"檔案: {os.path.basename(self.current_preview_path)} | 原始尺寸: {original_width}×{original_height} | 顯示尺寸: {new_width}×{new_height}"
            info_label = tk.Label(preview_window, text=info_text, font=("Microsoft YaHei", 10))
            info_label.pack(pady=5)
            
        except Exception as e:
            preview_window.destroy()
            messagebox.showerror("錯誤", f"無法開啟預覽視窗: {str(e)}")

    def clear_all_images(self):
        """清除所有圖片欄位"""
        filled_count = sum(1 for path in self.image_paths if path.strip())
        if filled_count == 0:
            return
        
        if messagebox.askyesno("確認", f"確定要清除所有 {filled_count} 個圖片檔案路徑嗎？"):
            for entry_info in self.image_entries:
                entry_info['var'].set("")
                entry_info['entry'].config(bg="white")
            
            self.image_paths = [""] * len(self.image_paths)
            self.reset_preview()
            
            # 更新所有欄位的顯示狀態
            for i in range(len(self.image_entries)):
                self.update_image_display_status(i)

    def reset_preview(self):
        """重設預覽區域"""
        self.preview_canvas.delete("all")
        self.current_preview_path = None
        self.preview_photo = None
        self.img_info_label.config(text="選擇圖片進行預覽", fg="gray")
        
        # 重設滾動區域
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
    def select_specific_files(self):
        """選擇特定檔案"""
        files = filedialog.askopenfilenames(
            title="選擇要包含的特定檔案（任何類型）",
            filetypes=[("所有檔案", "*.*")],
            initialdir=os.path.expanduser("~")
        )
        
        if files:
            new_files = [f for f in files if f not in self.selected_files]
            self.selected_files.extend(new_files)
            self.update_folder_display()
            
            if new_files:
                messagebox.showinfo("成功", f"✅ 成功新增 {len(new_files)} 個特定檔案")

    def select_cd_content(self):
        """選擇光碟片資料夾"""
        folder = filedialog.askdirectory(
            title="選擇光碟片資料夾",
            initialdir=os.path.expanduser("~")
        )
        
        if folder and folder not in self.cd_content_paths:
            self.cd_content_paths.append(folder)
            self.update_folder_display()
            messagebox.showinfo("成功", f"✅ 成功新增資料夾: {os.path.basename(folder)}")

    def update_folder_display(self):
        """更新資料夾顯示"""
        self.folder_listbox.delete(0, tk.END)
        
        if not self.cd_content_paths and not self.selected_files:
            self.folder_listbox.insert(tk.END, ">> 尚未選擇任何資料夾或檔案 <<")
            return

        # ✅ 修正：顯示選擇的檔案名稱
        if self.selected_files:
            self.folder_listbox.insert(tk.END, "📄 特定選擇的檔案:")
            for file_path in self.selected_files:
                filename = os.path.basename(file_path)
                size_str = self.format_size(os.path.getsize(file_path))
                self.folder_listbox.insert(tk.END, f"   📄 {filename} ({size_str})")
            self.folder_listbox.insert(tk.END, "")

        # ✅ 修正：顯示資料夾名稱
        if self.cd_content_paths:
            for folder_path in self.cd_content_paths:
                folder_name = os.path.basename(folder_path)
                self.folder_listbox.insert(tk.END, f"📁 資料夾: {folder_name}")
                self.folder_listbox.insert(tk.END, f"   路徑: {folder_path}")
                self.folder_listbox.insert(tk.END, "")

    def clear_folders(self):
        """清除所有資料夾和檔案"""
        total_items = len(self.cd_content_paths) + len(self.selected_files)
        if total_items == 0:
            return
            
        if messagebox.askyesno("確認", f"確定要清除所有內容嗎？"):
            self.cd_content_paths = []
            self.selected_files = []
            self.update_folder_display()

    def update_folder_preview(self, *args):
        """更新資料夾名稱預覽"""
        folder_name = self.generate_folder_name()
        if folder_name:
            self.folder_name_label.config(text=f"📁 資料夾名稱: {folder_name}", fg="green")
        else:
            self.folder_name_label.config(text="📁 資料夾名稱: (請填寫完整資訊)", fg="red")

    def generate_folder_name(self):
        """產生資料夾名稱"""
        parts = [
            self.year.get().strip(),
            self.category.get().strip(),
            self.case.get().strip(),
            self.volume.get().strip(),
            self.item.get().strip(),
            self.sheet_start.get().strip(),
            self.sheet_end.get().strip()
        ]
        
        return "-".join(parts) if all(parts) else None

    def format_size(self, size_bytes):
        """格式化檔案大小"""
        if size_bytes == 0:
            return "0 B"
        
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} PB"
    
    def get_roc_year(self):
        """取得當前民國年分"""
        from datetime import datetime
        current_year = datetime.now().year
        roc_year = current_year - 1911
        return roc_year

    def integrate_files(self):
        """整合檔案到指定資料夾"""
        # 取得有效的圖片檔案
        valid_images = [path for path in self.image_paths if path and os.path.exists(path)]
        
        if not valid_images and not self.cd_content_paths and not self.selected_files:
            messagebox.showwarning("警告", "請先選擇要整合的圖片檔案或資料夾內容！")
            return

        folder_name = self.generate_folder_name()
        if not folder_name:
            messagebox.showwarning("警告", "請填寫完整的命名資訊！")
            return

        # 修改：根據設定決定輸出目錄
        if self.use_fixed_path.get() and self.fixed_output_path.get():
            # 使用固定路徑
            output_dir = self.fixed_output_path.get()
            if not os.path.exists(output_dir):
                messagebox.showerror("錯誤", f"固定輸出路徑不存在：{output_dir}\n請檢查設定或重新選擇路徑")
                return
        else:
            # 詢問輸出目錄
            output_dir = filedialog.askdirectory(title="選擇整合後的輸出目錄")
            if not output_dir:
                return

        target_folder = os.path.join(output_dir, folder_name)
                
        try:
            if os.path.exists(target_folder):
                if not messagebox.askyesno("確認", f"資料夾 '{folder_name}' 已存在，是否要合併內容？"):
                    return
            else:
                os.makedirs(target_folder)

            total_copied = 0
            
            # 複製圖片檔案
            if valid_images:
                screenshots_folder = os.path.join(target_folder, "光碟檢測截圖")
                os.makedirs(screenshots_folder, exist_ok=True)
                
                # 取得民國年分
                roc_year = self.get_roc_year()
                
                for img_path in valid_images:
                    try:
                        original_filename = os.path.basename(img_path)
                        name, ext = os.path.splitext(original_filename)
                        
                        # 檢查原檔名是否已經以民國年分開頭
                        if original_filename.startswith(f"{roc_year}-"):
                            # 如果已經有民國年分，直接使用原檔名
                            new_filename = original_filename
                        else:
                            # 如果沒有民國年分，則加上
                            new_filename = f"{roc_year}-{original_filename}"
                        
                        # ✅ 先設定初始的 target_path
                        target_path = os.path.join(screenshots_folder, new_filename)
                                                    
                        # 處理同名檔案
                        counter = 1
                        while os.path.exists(target_path):
                            name_with_year, ext = os.path.splitext(new_filename)
                            duplicate_filename = f"{name_with_year}_{counter}{ext}"
                            target_path = os.path.join(screenshots_folder, duplicate_filename)
                            counter += 1
                        
                        shutil.copy2(img_path, target_path)
                        total_copied += 1
                        
                    except Exception as e:
                        print(f"複製圖片錯誤: {e}")
            # 複製資料夾內容
            if self.cd_content_paths:
                cd_content_folder = os.path.join(target_folder, "光碟片內容")
                os.makedirs(cd_content_folder, exist_ok=True)
                
                for folder_path in self.cd_content_paths:
                    try:
                        folder_name_only = os.path.basename(folder_path)
                        target_subfolder = os.path.join(cd_content_folder, folder_name_only)
                        
                        counter = 1
                        while os.path.exists(target_subfolder):
                            target_subfolder = f"{os.path.join(cd_content_folder, folder_name_only)}_{counter}"
                            counter += 1
                        
                        shutil.copytree(folder_path, target_subfolder)
                        folder_file_count = sum(len(files) for _, _, files in os.walk(target_subfolder))
                        total_copied += folder_file_count
                        
                    except Exception as e:
                        print(f"複製資料夾錯誤: {e}")

            # 創建整合報告
            self.create_integration_report(target_folder, total_copied)

            messagebox.showinfo("整合完成", f"✅ 整合完成！\n成功複製: {total_copied} 個檔案")
            
            # 第一步：先詢問是否要開啟資料夾
            if messagebox.askyesno("開啟資料夾", "是否要開啟整合後的資料夾？"):
                # 跨平台開啟資料夾
                try:
                    if os.name == 'nt':  # Windows
                        os.startfile(target_folder)
                    elif os.name == 'posix':  # macOS and Linux
                        import subprocess
                        if sys.platform == 'darwin':  # macOS
                            subprocess.run(['open', target_folder])
                        else:  # Linux
                            subprocess.run(['xdg-open', target_folder])
                except Exception as e:
                    print(f"無法開啟資料夾: {e}")
            
            # 第二步：自動重置所有資料
            self.reset_all_data()
            
            # 第三步：最後提示可以開始新的作業
            messagebox.showinfo("提示", "所有資料已自動清除，可以開始新的整合作業")


        except Exception as e:
            messagebox.showerror("錯誤", f"整合過程發生錯誤: {str(e)}")

    def create_integration_report(self, target_folder, total_copied):
        """創建整合報告"""
        try:
            report_path = os.path.join(target_folder, "整合報告.txt")
            
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("光碟片內容整合報告\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"整合時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"目標資料夾: {os.path.basename(target_folder)}\n")
                f.write(f"總複製檔案數: {total_copied} 個\n")
                f.write(f"系統使用者名稱: {self.current_user_display}\n\n")  
                
                # 新增命名資訊區塊
                f.write("檔案命名資訊:\n")
                f.write(f"  年度: {self.year.get().strip()}\n")
                f.write(f"  分類號: {self.category.get().strip()}\n")
                f.write(f"  案: {self.case.get().strip()}\n")
                f.write(f"  號: {self.volume.get().strip()}\n")
                f.write(f"  目: {self.item.get().strip()}\n")
                f.write(f"  片(總片數): {self.sheet_start.get().strip()}\n")
                f.write(f"  片(編號): {self.sheet_end.get().strip()}\n\n")
                
                # 檢查並記錄實際存在的圖片檔案
                screenshots_folder = os.path.join(target_folder, "光碟檢測截圖")
                if os.path.exists(screenshots_folder):
                    screenshot_files = []
                    for file in os.listdir(screenshots_folder):
                        if os.path.isfile(os.path.join(screenshots_folder, file)):
                            screenshot_files.append(file)
                    
                    if screenshot_files:
                        f.write(f"圖片檔案列表 (共 {len(screenshot_files)} 個):\n")
                        # 按檔名排序
                        screenshot_files.sort()
                        for i, img in enumerate(screenshot_files):
                            f.write(f"  {i+1}. {img}\n")
                        f.write("\n")
                
                # 檢查並記錄實際存在的光碟片內容資料夾
                cd_content_folder = os.path.join(target_folder, "光碟片內容")
                if os.path.exists(cd_content_folder):
                    cd_folders = []
                    for item in os.listdir(cd_content_folder):
                        if os.path.isdir(os.path.join(cd_content_folder, item)):
                            cd_folders.append(item)
                    
                    if cd_folders:
                        f.write(f"資料夾列表 (共 {len(cd_folders)} 個):\n")
                        # 按資料夾名排序
                        cd_folders.sort()
                        for i, folder in enumerate(cd_folders):
                            folder_path = os.path.join(cd_content_folder, folder)
                            # 計算資料夾內的檔案數量
                            file_count = sum(len(files) for _, _, files in os.walk(folder_path))
                            f.write(f"  {i+1}. {folder} (包含 {file_count} 個檔案)\n")
                        f.write("\n")
                
                # 如果有其他檔案直接放在主資料夾中
                main_files = []
                for file in os.listdir(target_folder):
                    file_path = os.path.join(target_folder, file)
                    if os.path.isfile(file_path) and file != "整合報告.txt":
                        main_files.append(file)
                
                if main_files:
                    f.write(f"主資料夾其他檔案 (共 {len(main_files)} 個):\n")
                    main_files.sort()
                    for i, file in enumerate(main_files):
                        f.write(f"  {i+1}. {file}\n")
                    f.write("\n")
                
                # 新增資料夾統計資訊
                total_files = 0
                total_folders = 0
                
                for root, dirs, files in os.walk(target_folder):
                    total_files += len(files)
                    total_folders += len(dirs)
                
                # 減去報告檔案本身
                total_files -= 1
                
                f.write("統計資訊:\n")
                f.write(f"  總檔案數: {total_files} 個\n")
                f.write(f"  總資料夾數: {total_folders} 個\n")
                
                # 計算總大小
                total_size = 0
                for root, dirs, files in os.walk(target_folder):
                    for file in files:
                        try:
                            file_path = os.path.join(root, file)
                            total_size += os.path.getsize(file_path)
                        except:
                            pass
                
                f.write(f"  總大小: {self.format_size(total_size)}\n")
                    
        except Exception as e:
            print(f"無法創建整合報告: {e}")
    
    def reset_all_data(self):
        """重置所有資料回到初始狀態"""
        # 清除圖片資料
        for entry_info in self.image_entries:
            entry_info['var'].set("")
        self.image_paths = [""] * len(self.image_paths)
        
        # 清除資料夾和檔案
        self.cd_content_paths = []
        self.selected_files = []
        
        # 清除命名欄位
        self.year.set("")
        self.category.set("")
        self.case.set("")
        self.volume.set("")
        self.item.set("")
        self.sheet_start.set("")
        self.sheet_end.set("")
        self.user_name.set(self.current_user_display)
        
        # 重設預覽
        self.reset_preview()
        
        # 更新顯示
        self.update_folder_display()
        
        # 重設所有圖片欄位狀態
        for i in range(len(self.image_entries)):
            self.update_image_display_status(i)


def main():
    root = tk.Tk()
    app = DataIntegratorApp(root)
    
    def on_closing():
        if messagebox.askokcancel("退出", "確定要退出程式嗎？"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()