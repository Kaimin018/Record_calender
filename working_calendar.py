import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter
from tkcalendar import Calendar
import json
import os
from datetime import datetime, date
import webbrowser
import re
import tkinter.font as tkfont
import openpyxl
import threading
from PIL import Image, ImageTk
import platform


SCRIPT_DIR = os.path.dirname(__file__)
DATA_FILE = os.path.join(SCRIPT_DIR, 'todo_calendar.json')


# 設定 customtkinter 的外觀模式 ('System'/'Dark'/'Light') 和顏色主題 ('blue'/'green'/'dark-blue')
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

# 定義所有可能的狀態
STATUS_OPTIONS = ["Pending", "In progress", "Completed", "Cancelled", "On hold"]

# 將 save_tasks 函數放在類別外面
def save_tasks(tasks):
    """將待辦事項儲存到檔案"""
    try:
        # print(f"Attempting to save to: {DATA_FILE}") # 可以在需要時取消註釋進行除錯
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(tasks, f, indent=4, ensure_ascii=False)
        # print("Save successful.") # 可以在需要時取消註釋進行除錯
        return True # 儲存成功時返回 True
    except Exception as e:
        print(f"Error saving tasks: {e}") # 在控制台打印錯誤
        return False, e # 儲存失敗時返回 False 和錯誤


def load_tasks():
    """從檔案載入待辦事項"""
    tasks = []
    # print(f"Attempting to load from: {DATA_FILE}") # 可以在需要時取消註釋進行除錯
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                tasks = json.load(f)
            # print("Load successful.") # 可以在需要時取消註釋進行除錯
            current_id = 0
            valid_tasks = []
            for task in tasks:
                 if isinstance(task, dict) and 'description' in task:
                    if 'id' not in task or not isinstance(task['id'], (int, float)):
                        task['id'] = current_id
                    task['id'] = int(task['id'])

                    if 'due_date' not in task:
                        task['due_date'] = None
                        
                    if 'creation_time' not in task:
                         task['creation_time'] = None
                         
                    if 'status' not in task or not isinstance(task['status'], str):
                        task['status'] = 'Pending'
                    else:
                        # 嘗試將小寫/全大寫/首字大寫都轉成標準格式
                        matched_status = next((s for s in STATUS_OPTIONS if s.lower() == task['status'].lower()), None)
                        task['status'] = matched_status if matched_status else 'Pending'
                         
                    if 'note' not in task:
                         task['note'] = ''
                    elif task['note'] is None:
                         task['note'] = ''
                         
                    if 'image_path' not in task:
                         task['image_path'] = None

                    valid_tasks.append(task)
                    current_id = max(current_id, task['id'])

            tasks = valid_tasks
            load_tasks.counter = current_id + 1 if tasks else 0
        except json.JSONDecodeError:
            messagebox.showwarning("警告", "無法解析待辦事項檔案，將建立新的檔案。")
            load_tasks.counter = 0
        except Exception as e:
            messagebox.showerror("錯誤", f"載入檔案時發生錯誤: {e}")
            load_tasks.counter = 0
    else:
         load_tasks.counter = 0
         # print(f"Data file not found at {DATA_FILE}.") # 可以在需要時取消註釋進行除錯
    return tasks

load_tasks.counter = 0

def get_next_id():
    """取得下一個可用的唯一 ID"""
    task_id = load_tasks.counter
    load_tasks.counter += 1
    return task_id

# ****** 輔助函數：格式化日期時間 ******
def format_datetime(dt_str):
    """格式化-MM-DD HH:MM:SS 字串為可讀格式，並包含星期幾"""
    if not dt_str:
        return "無時間"
    try:
        dt_obj = datetime.strptime(dt_str, '%Y-%m-%d %H:%M:%S')
        # 格式化為<\ctrl97>-MM-DD HH:MM:SS (星期幾)
        return dt_obj.strftime('%Y-%m-%d %H:%M:%S (%a)') # %a 縮寫星期幾
    except ValueError:
        # 如果不是標準格式，嘗試只解析日期
        try:
             dt_obj = datetime.strptime(dt_str.split(' ')[0], '%Y-%m-%d').date()
             # 格式化為<\ctrl97>-MM-DD (星期幾)
             return dt_obj.strftime('%Y-%m-%d (%a)')
        except ValueError:
             return dt_str # 如果格式無效，返回原字串

# ****** 輔助函數：格式化日期並包含星期幾 ******
def format_date_with_weekday(date_str):
    r"""格式化<\ctrl97>-MM-DD 字串為<\ctrl97>-MM-DD (星期幾)"""
    if not date_str:
        return "無到期日"
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
        return date_obj.strftime('%Y-%m-%d (%a)') # %a 縮寫星期幾
    except ValueError:
        return date_str # 如果格式無效，返回原字串

# ****** 輔助函數：檢查日期是否過期 ******
def is_past_due(date_str):
    """檢查給定日期字串是否在今天之前"""
    if not date_str:
        return False # 無到期日不算過期
    try:
        due_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        return due_date < date.today()
    except ValueError:
        return False # 日期格式無效不檢查過期


class TodoApp(customtkinter.CTk): # 繼承 customtkinter.CTk
    def __init__(self):
        super().__init__()

        self.title("待辦事項 & 行事曆工具")
        self.geometry("1000x650")

        # --- 新增 Canvas 和 Scrollbar ---
        self.main_canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        # 建立一個 Frame 放所有內容
        self.inner_frame = customtkinter.CTkFrame(self.main_canvas)
        self.inner_frame_id = self.main_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.main_scrollbar.pack(side="right", fill="y")

        # 綁定 Frame 尺寸改變時更新 Canvas scrollregion
        self.inner_frame.bind("<Configure>", lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))

        # 讓滑鼠滾輪也能滾動
        self.main_canvas.bind_all("<MouseWheel>", lambda e: self.main_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        self.tasks = load_tasks()
        self.editing_task_id = None
        self.save_thread = None
        self.log_entries = [] # 操作日誌列表
        self.show_on_hold = True # 控制是否顯示 On hold 項目
        self._sort_direction = {} # Initialize sorting direction
        # ****** 點擊 Treeview 標頭處理排序 ******
        # ****** 實現 Treeview 排序邏輯 (包含在 populate_treeview 中應用) ******
        self._sort_direction = {}  # 用於儲存每列的排序方向
        self._sort_column = None  # 用於儲存當前排序列 ID


        # --- GUI 元件設定 ---
        # 直接在 __init__ 中建立和佈局元件
        self.create_widgets()
        self.layout_widgets()
        self.create_menu()


        # Initial population - Schedule this after __init__ completes
        self.tab_notebook.bind("<<NotebookTabChanged>>", lambda e: self.populate_treeview())
        

        # 綁定 Ctrl+S 儲存快捷鍵 (platform specific)
        if platform.system() == "Darwin":  # macOS
            self.bind_all("<Command-KeyPress-s>", self.save_tasks_shortcut)
        else:  # Windows 或 Linux
            self.bind_all("<Control-KeyPress-s>", self.save_tasks_shortcut)
            #print("Windows/Linux detected, binding Ctrl+S") # 除錯用


        # 綁定 Enter 鍵到特定的處理函數
        self.bind_all("<Return>", self.handle_return_key)

        # 載入過期 icon
        self.load_icons()

        # 初始化日誌
        self.log_operation("應用程式啟動")


    # ****** 載入圖標方法 ******
    def load_icons(self):
        """載入應用程式所需的圖標"""
        try:
            # 在同一個目錄下尋找 icon 檔案，或者指定完整路徑
            icon_path = os.path.join(SCRIPT_DIR, "/assets/warning.png") # 使用絕對路徑
            # print(f"Looking for icon at: {icon_path}") # 除錯用
            if os.path.exists(icon_path):
                original_image = Image.open(icon_path)
                resized_image = original_image.resize((16, 16), Image.Resampling.LANCZOS)
                self.warning_icon = ImageTk.PhotoImage(resized_image)
                self._icon_refs = [self.warning_icon] # Store references
                # print("Warning icon loaded successfully.") # 除錯用
            else:
                self.warning_icon = None
                self.log_operation("警告圖標檔案 'warning.png' 未找到。")
        except Exception as e:
            self.warning_icon = None
            self.log_operation(f"載入警告圖標時發生錯誤: {e}")


    # ****** 操作日誌方法 ******
    def log_operation(self, message):
        """記錄操作到日誌列表"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"{timestamp} - {message}"
        self.log_entries.append(log_entry)
        # print(log_entry) # 同時在控制台打印方便調試

    # ****** 創建頂部選單方法 ******
    def create_menu(self):
        """創建應用程式頂部選單"""
        # 移除舊的 menubar 以避免重複添加
        if hasattr(self, 'menubar') and self.menubar:
            self.menubar.destroy()

        self.menubar = tk.Menu(self)
        self.config(menu=self.menubar) # 將選單欄配置到主視窗

        # 檔案選單
        filemenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="檔案", menu=filemenu)
        filemenu.add_command(label="儲存 (Ctrl+S)", command=lambda: self.save_tasks_shortcut()) # 使用 lambda
        filemenu.add_command(label="匯出為 Excel", command=lambda: self.export_to_excel()) # 使用 lambda
        filemenu.add_separator()
        filemenu.add_command(label="結束", command=self.quit)

        # 查看選單
        viewmenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="查看", menu=viewmenu)
        # ****** 添加顯示操作日誌的選項 ******
        viewmenu.add_command(label="操作日誌", command=lambda: self.show_log_window()) # 使用 lambda

        # 設置選單
        optionsmenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="設置", menu=optionsmenu)
        # ****** 添加顯示/隱藏 On hold 項目的選項 ******
        # 使用 lambda 綁定 command，並在 lambda 內部讀取 variable 的值
        self.show_on_hold_var = tk.BooleanVar(value=self.show_on_hold)
        optionsmenu.add_checkbutton(label="顯示 On hold 項目", variable=self.show_on_hold_var, command=lambda: self.toggle_show_on_hold())


    # ****** 顯示操作日誌視窗方法 ******
    def show_log_window(self):
        """顯示操作日誌視窗"""
        log_window = customtkinter.CTkToplevel(self)
        log_window.title("操作日誌")
        log_window.geometry("600x400")

        log_textbox = customtkinter.CTkTextbox(log_window, wrap="word")
        log_textbox.pack(padx=10, pady=10, fill="both", expand=True)

        # 將日誌內容插入 Textbox
        for entry in self.log_entries:
            log_textbox.insert(tk.END, entry + "\n")

        log_textbox.configure(state="disabled") # 設置為唯讀

        # 設置焦點
        log_window.transient(self)
        log_window.grab_set()
        log_window.after(10, log_window.lift)


    # ****** 切換顯示/隱藏 On hold 方法 ******
    def toggle_show_on_hold(self):
        """切換顯示或隱藏 On hold 狀態的待辦事項"""
        self.show_on_hold = self.show_on_hold_var.get() # 從 variable 讀取狀態
        self.populate_treeview() # 重新填充 Treeview 以應用過濾

        self.log_operation(f"切換顯示 On hold 項目為: {'顯示' if self.show_on_hold else '隱藏'}。")


    def handle_return_key(self, event):
        """處理 Enter 鍵事件，判斷是否觸發儲存"""
        focused_widget = self.focus_get()

        # 判斷是否同時按下了 Ctrl 鍵
        is_ctrl_pressed = (event.state & 0x4) != 0 # Windows/Linux Ctrl key state
        # 對於 macOS Command 鍵可能是 event.state & 0x8 或 0xf0
        # is_command_pressed = (event.state & 0x8) != 0 # Example for Command key

        # 如果焦點在備註輸入框 (CTkTextbox) 或詳細資訊備註框 (tk.Text)
        if focused_widget in [self.note_textbox, self.details_note_textbox]:
            # 只有同時按下 Ctrl 鍵才觸發儲存
            if is_ctrl_pressed: # or is_command_pressed for macOS
                self.save_task_gui()
                # 阻止 Enter 的預設行為 (插入換行)
                return "break"
            else:
                # 單獨 Enter: 允許預設行為 (插入換行)
                # 返回 None 或 '' 讓事件繼續傳播到 widget
                return None # 或者 return ''

        # 如果焦點在描述輸入框 (CTkEntry)
        elif focused_widget == self.desc_entry:
            # Enter 鍵觸發儲存/新增
            self.save_task_gui()
            # 阻止 Enter 的預設行為
            return "break"

        # 如果焦點在日期相關元件 (Label 或 Button)
        elif focused_widget in [self.date_display_label, self.select_date_button]:
             # Enter 鍵觸發儲存/新增
             self.save_task_gui()
             return "break"

        # 如果焦點在其他地方 (例如 Treeview)，不做任何事，讓事件傳播
        return None


    def create_widgets(self):        
        
        # 設置 Tab 的字體大小
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("Arial", 12))  # 設置字體和大小，例如 Arial 字體，大小 12
        
        # 輸入區域 Frame (使用 CTkFrame)
        self.input_frame = customtkinter.CTkFrame(self.inner_frame, corner_radius=10)

        self.desc_label = customtkinter.CTkLabel(self.input_frame, text="內容:")
        self.desc_entry = customtkinter.CTkEntry(self.input_frame, width=300)

        # 日期選擇區域
        self.date_label = customtkinter.CTkLabel(self.input_frame, text="到期日:")
        # 日期顯示 Label，用於顯示 Calendar 選取的日期
        self.date_display_label = customtkinter.CTkLabel(self.input_frame, text=datetime.now().strftime('%Y-%m-%d'), width=100, anchor=tk.W)
        # ****** 修改選取日期按鈕的 command，並應用 lambda ******
        self.select_date_button = customtkinter.CTkButton(self.input_frame, text="選取日期", command=lambda: self.open_calendar_dialog(), width=100)


        # Note 輸入區域
        self.note_label = customtkinter.CTkLabel(self.input_frame, text="備註/網址:")
        # 使用 CTkTextbox 方便輸入多行和長網址
        self.note_textbox = customtkinter.CTkTextbox(self.input_frame, width=300, height=50)


        self.action_button_frame = customtkinter.CTkFrame(self.input_frame, fg_color="transparent") # 使用透明框架
        self.save_button = customtkinter.CTkButton(self.action_button_frame, text="新增待辦事項", command=lambda: self.save_task_gui())
        self.cancel_edit_button = customtkinter.CTkButton(self.action_button_frame, text="取消編輯", command=lambda: self.cancel_edit(), fg_color="gray", hover_color="darkgray")

        ##############################################################
        
        # 創建 Tab Notebook
        self.tab_notebook = ttk.Notebook(self.inner_frame)        

        # 創建各個狀態的 Tab
        self.tabs = {}
        for status in STATUS_OPTIONS:
            tab_frame = customtkinter.CTkFrame(self.tab_notebook)
            self.tab_notebook.add(tab_frame, text=status.capitalize())
            self.tabs[status] = tab_frame

        # 創建一個 "All" Tab 顯示所有項目
        self.all_tab_frame = customtkinter.CTkFrame(self.tab_notebook)
        self.tab_notebook.add(self.all_tab_frame, text="All")
        self.tabs["all"] = self.all_tab_frame

        # 創建 Treeview 顯示區域
        self.create_treeview_widgets()
        
        
        # 操作按鈕 Frame (使用 CTkFrame) - 包含狀態選擇 和 匯出按鈕
        self.button_frame = customtkinter.CTkFrame(self.inner_frame, corner_radius=10)

        self.status_combobox_label = customtkinter.CTkLabel(self.button_frame, text="變更狀態為:")
        self.status_combobox = customtkinter.CTkComboBox(self.button_frame, values=STATUS_OPTIONS, width=120) # 稍微加寬下拉選單
        self.status_combobox.set("Completed") # 預設選中 "Completed"

        # ****** 使用 lambda 綁定 set_status_button command ******
        self.set_status_button = customtkinter.CTkButton(self.button_frame, text="變更選取狀態", command=lambda: self.set_selected_task_status())
        # ****** 使用 lambda 綁定 delete_button command ******
        self.delete_button = customtkinter.CTkButton(self.button_frame, text="刪除選取", command=lambda: self.delete_selected_task(), fg_color="red", hover_color="darkred")

        # ****** 使用 lambda 綁定 export_button command ******
        self.export_button = customtkinter.CTkButton(self.button_frame, text="匯出為 Excel", command=lambda: self.export_to_excel())


        # 詳細資訊顯示區域 Frame (使用 CTkFrame)
        self.details_frame = customtkinter.CTkFrame(self.inner_frame, corner_radius=10)
        self.details_frame.columnconfigure(1, weight=1) # 讓值欄位可以擴展

        self.details_label = customtkinter.CTkLabel(self.details_frame, text="選取任務詳細資訊", font=customtkinter.CTkFont(weight="bold"))
        self.details_label.grid(row=0, column=0, columnspan=2, pady=(5, 10), sticky="ew", padx=10)

        ttk.Separator(self.details_frame, orient='horizontal').grid(row=1, column=0, columnspan=2, sticky="ew")

        # 詳細資訊 Labels
        # ****** 增加建立時間詳細資訊顯示 ******
        self.details_creation_time_label = customtkinter.CTkLabel(self.details_frame, text="建立時間:", anchor=tk.W)
        self.details_creation_time_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W)

        self.details_desc_label = customtkinter.CTkLabel(self.details_frame, text="內容:", anchor=tk.W)
        self.details_desc_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W, wraplength=450) # 內容值，加寬換行長度

        self.details_date_label = customtkinter.CTkLabel(self.details_frame, text="到期日:", anchor=tk.W)
        self.details_date_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W) # 日期值

        self.details_status_label = customtkinter.CTkLabel(self.details_frame, text="狀態:", anchor=tk.W)
        self.details_status_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W) # 狀態值

        self.details_note_label = customtkinter.CTkLabel(self.details_frame, text="備註/網址:", anchor=tk.NW) # 備註標籤靠左上
        # ****** 使用標準的 tkinter.Text ******
        self.details_note_textbox = tk.Text(self.details_frame, wrap="word", height=4, state="disabled", # state="disabled" 預設為唯讀
                                            bg=self._apply_appearance_mode(customtkinter.ThemeManager.theme["CTkEntry"]["fg_color"]), # 嘗試匹配 CTk 的背景色
                                            fg=self._apply_appearance_mode(customtkinter.ThemeManager.theme["CTkEntry"]["text_color"]), # 嘗試匹配 CTk 的前景色
                                            relief="flat", # 嘗試匹配 CTk 的無邊框風格
                                            padx=0, pady=0 # 移除預設的 Text 內邊距
                                            )
        # 為標準 Text widget 添加滾動條
        self.details_note_scrollbar_y = ttk.Scrollbar(self.details_frame, command=self.details_note_textbox.yview)
        self.details_note_textbox.configure(yscrollcommand=self.details_note_scrollbar_y.set)

        # 狀態訊息 Label (使用 CTkLabel)
        self.status_label = customtkinter.CTkLabel(self.inner_frame, text="", anchor=tk.W, padx=10)


    def create_treeview_widgets(self):
        """為每個 Tab 創建 Treeview"""
        self.treeviews = {}

        for status, frame in self.tabs.items():
            tree_frame = customtkinter.CTkFrame(frame, corner_radius=10)
            tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

            tree_scrollbar_y = ttk.Scrollbar(tree_frame)
            tree_scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

            treeview = ttk.Treeview(
                tree_frame,
                columns=("creation_time", "description", "due_date", "status", "note"),
                show="headings",
                yscrollcommand=tree_scrollbar_y.set,
                xscrollcommand=tree_scrollbar_x.set,
            )
            tree_scrollbar_y.config(command=treeview.yview)
            tree_scrollbar_x.config(command=treeview.xview)

            # 設置欄位標題和寬度
            treeview.heading("creation_time", text="建立時間", anchor=tk.W)
            treeview.heading("description", text="內容", anchor=tk.W)
            treeview.heading("due_date", text="到期日", anchor=tk.CENTER)
            treeview.heading("status", text="狀態", anchor=tk.CENTER)
            treeview.heading("note", text="備註/網址", anchor=tk.W)

            treeview.column("creation_time", width=180, anchor=tk.W, stretch=False)
            treeview.column("description", width=250, anchor=tk.W, stretch=True)
            treeview.column("due_date", width=120, anchor=tk.CENTER, stretch=False)
            treeview.column("status", width=80, anchor=tk.CENTER, stretch=False)
            treeview.column("note", width=520, anchor=tk.W, stretch=True)

            treeview.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
            tree_scrollbar_y.grid(row=0, column=1, sticky="ns")
            tree_scrollbar_x.grid(row=1, column=0, sticky="ew")

            frame.grid_columnconfigure(0, weight=1)
            frame.grid_rowconfigure(0, weight=1)

            # 綁定雙擊事件
            treeview.bind("<Double-1>", self.load_task_for_editing)
            # 綁定選擇事件  
            treeview.bind("<<TreeviewSelect>>", self.display_selected_task_details)

            self.treeviews[status] = treeview
    
    
    def layout_widgets(self):
        # 輸入區域
        self.input_frame.pack(pady=(10, 5), padx=10, fill="x", expand=False)

        # 狀態/刪除/匯出按鈕區
        self.button_frame.pack(pady=5, padx=10, fill="x", expand=False)

        # 任務列表（Tab Notebook）
        self.tab_notebook.pack(pady=10, padx=10, fill="both", expand=True)

        # 任務細節
        self.details_frame.pack(pady=5, padx=10, fill="x", expand=False)

        # input_frame 內部用 grid（不變）
        self.desc_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.desc_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.date_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.date_display_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.select_date_button.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.note_label.grid(row=2, column=0, padx=5, pady=5, sticky="nw")
        self.note_textbox.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.action_button_frame.grid(row=0, column=3, rowspan=3, padx=10, pady=5, sticky="ns")
        self.save_button.pack(side=tk.TOP, pady=5, fill="x")
        
        self.input_frame.grid_columnconfigure(1, weight=1)
        self.input_frame.grid_columnconfigure(2, weight=0)
        self.input_frame.grid_columnconfigure(3, weight=0)
        self.input_frame.grid_rowconfigure(2, weight=1)

        # button_frame 內部用 pack（不變）
        self.status_combobox_label.pack(side=tk.LEFT, padx=(5, 0), pady=10)
        self.status_combobox.pack(side=tk.LEFT, padx=5, pady=10)
        self.set_status_button.pack(side=tk.LEFT, padx=5, pady=10)
        self.delete_button.pack(side=tk.LEFT, padx=10, pady=10)
        self.export_button.pack(side=tk.RIGHT, padx=10, pady=10)

        # details_frame 內部用 grid（不變）
        r = 2
        self.details_desc_label.grid(row=r, column=0, padx=5, pady=2, sticky="nw")
        self.details_desc_value.grid(row=r, column=1, columnspan=2, padx=5, pady=2, sticky="ew")
        
        r += 1
        self.details_date_label.grid(row=r, column=0, padx=5, pady=2, sticky="nw")
        self.details_date_value.grid(row=r, column=1, padx=5, pady=2, sticky="ew")
        self.details_creation_time_label.grid(row=r, column=2, padx=5, pady=2, sticky="w")
        self.details_creation_time_value.grid(row=r, column=3, padx=5, pady=2, sticky="w")
        
        r += 1
        self.details_status_label.grid(row=r, column=0, padx=5, pady=2, sticky="nw")
        self.details_status_value.grid(row=r, column=1, columnspan=2, padx=5, pady=2, sticky="ew")
        
        r += 1
        self.details_note_label.grid(row=r, column=0, padx=5, pady=2, sticky="nw")
        self.details_note_textbox.grid(row=r, column=1, columnspan=2, padx=(5, 0), pady=2, sticky="nsew")
        self.details_note_scrollbar_y.grid(row=r, column=3, padx=(0, 5), pady=2, sticky="ns")
        self.details_frame.grid_columnconfigure(1, weight=1)
        self.details_frame.grid_columnconfigure(2, weight=1)
        self.details_frame.grid_columnconfigure(3, weight=0)

    # ****** 修改 open_calendar_dialog 實現相鄰位置 ******
    def open_calendar_dialog(self):
        """打開日曆選擇對話框，並嘗試定位在主視窗旁邊"""
        def grab_date():
            """從日曆獲取選取的日期並更新顯示"""
            try:
                selected_date = cal.selection_get()
                self.date_display_label.configure(text=selected_date.strftime('%Y-%m-%d'))
            except Exception as e:
                 messagebox.showerror("日期選取錯誤", f"無法獲取選取的日期: {e}")
            finally:
                top.destroy() # 無論成功或失敗都關閉對話框


        top = customtkinter.CTkToplevel(self) # 使用 CTk 頂級視窗
        top.title("選取日期")
        top.transient(self) # 設定為主要視窗的暫時視窗 (讓對話框在主視窗之上)
        top.grab_set() # 模式對話框，阻止與主視窗互動
        top.resizable(True, True) # 允許視窗可拉伸

        # 初始化日曆為當前顯示的日期或今天
        try:
            initial_date_str = self.date_display_label.cget("text")
            initial_date = datetime.strptime(initial_date_str, '%Y-%m-%d').date()
        except (ValueError, tk.TclError):
             initial_date = date.today() # 如果顯示的日期無效，使用今天

        # 創建 Calendar widget，並設定較大的字體
        calendar_font = tkfont.Font(family="Arial", size=12) # 可以根據需要調整字體和大小

        cal = Calendar(top, selectmode='day', date_pattern='yyyy-mm-dd', year=initial_date.year, month=initial_date.month, day=initial_date.day,
                       font=calendar_font, # 套用字體
                       # ttk style options can be added here, e.g., style='TCalendar'
                       )
        cal.pack(padx=10, pady=10)

        ok_button = customtkinter.CTkButton(top, text="確定", command=grab_date)
        ok_button.pack(pady=10)

        # ****** 計算並設定對話框位置 ******
        self.update_idletasks() # 更新視窗狀態以獲取正確的幾何信息
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        # main_height = self.winfo_height() # 不一定需要主視窗高度

        # 獲取對話框的請求尺寸 (需要先 pack/place/grid 才能獲取)
        # 這裡日曆已經 pack 了，應該可以獲取
        top.update_idletasks() # 更新對話框狀態
        dialog_width = top.winfo_width()
        # dialog_height = top.winfo_height() # 不一定需要對話框高度

        # 計算位置：主視窗右側
        # 考慮到主視窗的邊框和陰影，以及對話框的邊框，可以稍微調整
        x_pos = main_x + main_width + 10 # 10 像素間距
        y_pos = main_y # 與主視窗頂部對齊

        # 如果右側空間不足，可以考慮放在主視窗下方或其他位置
        # 簡單判斷：如果右側放置後超出螢幕範圍，可以嘗試放在左側或居中
        screen_width = self.winfo_screenwidth()
        if x_pos + dialog_width > screen_width:
            # 放在左側
            x_pos = main_x - dialog_width - 10
            # 如果左側也放不下，就居中顯示
            if x_pos < 0:
                 x_pos = main_x + (main_width - dialog_width) // 2
                 y_pos = main_y + 50 # 稍微偏下避免遮擋標題列


        top.geometry(f"+{x_pos}+{y_pos}")


        # 設置焦點到對話框
        top.after(10, top.lift) # 確保對話框在最前面
        # top.focus_force() # 強制對話框獲取焦點


    def save_task_gui(self):
        """儲存（新增或修改）待辦事項"""
        description = self.desc_entry.get().strip()
        due_date_str = self.date_display_label.cget("text").strip() # 從 Label 獲取日期
        note = self.note_textbox.get("1.0", "end-1c").strip() # 從 Textbox 獲取備註


        # 檢查是否是從備註框觸發的 Ctrl+Enter，且備註框是空的，如果是，則不新增
        focused_widget = self.focus_get()
        if focused_widget in [self.note_textbox, self.details_note_textbox] and not description and not note and self.editing_task_id is None:
             self.update_status("備註內容和描述都為空，不新增。")
             self.log_operation("嘗試新增待辦事項失敗：備註和內容為空。")
             return


        if not description:
            messagebox.showwarning("輸入錯誤", "待辦事項內容不能為空。")
            self.log_operation("嘗試新增待辦事項失敗：內容為空。")
            return

        due_date = None
        if due_date_str:
            try:
                # 驗證日曆選取的日期格式是否正確 (理論上 Calendar widget 保證格式)
                datetime.strptime(due_date_str, '%Y-%m-%d')
                due_date = due_date_str
            except ValueError:
                messagebox.showwarning("內部錯誤", "日期格式無效。")
                self.log_operation("嘗試新增/修改待辦事項失敗：日期格式無效。")
                return

        # 根據 editing_task_id 判斷是新增還是修改
        if self.editing_task_id is None:
            # ====== 新增模式 ======
            task = {
                'id': get_next_id(), # 分配新的唯一 ID
                'description': description,
                'due_date': due_date,
                'status': 'Pending', # 新增任務預設狀態為 pending
                'note': note,
                'creation_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), # 記錄建立時間
                'image_path': None # 初始化 image_path 欄位
            }
            self.tasks.append(task)
            message = "待辦事項已新增。"
            log_msg = f"新增待辦事項 (ID: {task['id']}): '{description}'"
        else:
            # ====== 修改模式 ======
            task_to_edit = next((task for task in self.tasks if task['id'] == self.editing_task_id), None)
            if task_to_edit:
                task_to_edit['description'] = description
                task_to_edit['due_date'] = due_date
                task_to_edit['note'] = note
                # creation_time 不修改
                # 狀態不通過這裡修改，通過下面的按鈕修改
                # image_path 不通過這裡修改
                message = f"待辦事項 (ID: {self.editing_task_id}) 已修改。"
                log_msg = f"修改待辦事項 (ID: {self.editing_task_id}): '{description}'"
            else:
                messagebox.showerror("錯誤", f"找不到 ID 為 {self.editing_task_id} 的待辦事項進行修改。")
                self.log_operation(f"嘗試修改待辦事項失敗：找不到 ID {self.editing_task_id}。")
                self.cancel_edit() # 取消編輯狀態
                return

        self.log_operation(log_msg)

        # ====== 將儲存操作移到執行緒中 ======
        self._start_save_thread(
            tasks_to_save=list(self.tasks), # 傳遞 tasks 的副本
            message=f"{'新增' if self.editing_task_id is None else '修改'}並儲存中...",
            completion_callback=lambda success, err: self.after(0, self._on_save_complete_gui, success, err, message)
        )


    # ****** 執行緒相關方法 ******
    def _worker_save(self, tasks_to_save, completion_callback):
        """Worker function to run save_tasks in a separate thread."""
        try:
            save_result = save_tasks(tasks_to_save)  # 呼叫 save_tasks 函數

            # 根據 save_tasks 的返回值類型，正確設定 success 和 error
            if isinstance(save_result, tuple) and len(save_result) == 2 and save_result[0] is False:
                success = False
                error = save_result[1]
            elif save_result is True:
                success = True
                error = None
            else:
                success = False
                error = f"Unexpected save result: {save_result}"

            # 通知主執行緒
            self.after(0, completion_callback, success, error)

        except Exception as e:
            # 捕捉未預期的例外
            self.after(0, completion_callback, False, e)


    def _start_save_thread(self, tasks_to_save, message="儲存中...", completion_callback=None):
        """Starts the save operation in a new thread."""
        if self.save_thread is not None and self.save_thread.is_alive():
            self.update_status("儲存操作進行中，請稍候...")
            self.log_operation("嘗試儲存時發現已有儲存進行中。")
            return

        self.save_thread = threading.Thread(target=self._worker_save, args=(tasks_to_save, completion_callback))
        self.save_thread.daemon = True
        self.update_status(message)
        self.log_operation(f"開始背景儲存操作: {message}")
        self.save_thread.start()


    def _on_save_complete_gui(self, success, error, user_message):
        """在主執行緒中執行，處理儲存完成後的 GUI 更新"""
        self.save_thread = None

        if success:
            # 儲存成功後更新介面
            self.populate_treeview()
            self.clear_input_fields()
            self.cancel_edit()
            self.update_status(user_message)
        else:
            # 儲存失敗的處理
            self.update_status(f"儲存失敗: {error}")
            messagebox.showerror("儲存錯誤", f"儲存檔案時發生錯誤:\n{error}")
            self.log_operation(f"背景儲存操作失敗: {error}")


    def save_tasks_shortcut(self, event=None): # 接受 event 參數
        """Ctrl+S 快捷鍵儲存任務"""
        if self.editing_task_id is not None:
            self.save_task_gui()
            print("儲存任務 (Ctrl+S)...success")
        else:
            self._start_save_thread(
                tasks_to_save=list(self.tasks),
                message="儲存中 (Ctrl+S)...",
                completion_callback=lambda success, err: self.after(0, self._on_save_complete_gui, success, err, "待辦事項已儲存 (Ctrl+S)。")
            )
            self.log_operation("通過快捷鍵觸發儲存。")
            print("儲存任務 (Ctrl+S)...")
        return "break" # 阻止事件繼續傳播


    def on_treeview_heading_click(self, event):
        """處理 Treeview 標頭點擊事件，實現排序"""
        region = self.task_tree.identify_region(event.x, event.y)
        if region == "heading":
            col = self.task_tree.identify_column(event.x)
            # 將列 ID 轉換為列名 (如 #1 -> creation_time, #2 -> description 等)
            # Treeview columns 是 ("creation_time", "description", "due_date", "status", "note")
            column_map = {
                "#1": "creation_time",
                "#2": "description",
                "#3": "due_date",
                "#4": "status",
                "#5": "note"
            }
            col_id = column_map.get(col)

            if col_id:
                # ****** 執行排序 ******
                # 更新當前排序列和方向
                if self._sort_column == col_id:
                    # 如果是同一列，切換排序方向
                    self._sort_direction[col_id] = 'descending' if self._sort_direction.get(col_id, 'ascending') == 'ascending' else 'ascending'
                else:
                    # 如果是新列，設定為遞增排序
                    self._sort_column = col_id
                    self._sort_direction = {col_id: 'ascending'} # 清除其他列的排序狀態


                self.populate_treeview() # 重新填充並應用新的排序

                # 更新標頭箭頭指示
                # 清除所有標頭的舊箭頭
                for c in column_map.values():
                    # 安全地處理可能沒有箭頭的情況
                    current_text_parts = self.task_tree.heading(c, 'text').split(' ')
                    current_text = current_text_parts[0] if current_text_parts else ""
                    self.task_tree.heading(c, text=current_text)

                # 添加新箭頭到當前排序列的標頭
                arrow = ' ▲' if self._sort_direction[col_id] == 'ascending' else ' ▼'
                current_text = self.task_tree.heading(col_id, 'text').split(' ')[0] # 再次安全獲取原始標題文字
                self.task_tree.heading(col_id, text=f"{current_text}{arrow}")


                self.log_operation(f"按 '{col_id}' 欄位進行了 {'遞減' if self._sort_direction[col_id] == 'descending' else '遞增'} 排序。")


    def populate_treeview(self):
        """只清空並填充目前顯示的 Treeview"""
        current_tab_text = self.tab_notebook.tab(self.tab_notebook.select(), "text")
        current_tab = current_tab_text.strip()
        current_tab_status = next((s for s in STATUS_OPTIONS if s.lower() == current_tab.lower()), None)

        # 只處理目前顯示的 Treeview
        if current_tab.lower() == "all":
            tasks_to_display = self.tasks
            treeview = self.treeviews["all"]
        else:
            tasks_to_display = [task for task in self.tasks if task.get("status") == current_tab_status]
            treeview = self.treeviews.get(current_tab_status)
            if treeview is None:
                return

        # 只清空目前的 Treeview
        for item in treeview.get_children():
            treeview.delete(item)

        # 排序邏輯（保持原有）
        if self._sort_column:
            sort_col_id = self._sort_column
            sort_direction = self._sort_direction.get(sort_col_id, 'ascending')
            def sort_key(task):
                value = task.get(sort_col_id)
                if sort_col_id in ['due_date', 'creation_time']:
                    try:
                        if sort_col_id == 'due_date':
                            return datetime.strptime(str(value).split(' ')[0], '%Y-%m-%d').date() if value else datetime.max.date()
                        if sort_col_id == 'creation_time':
                            return datetime.strptime(str(value), '%Y-%m-%d %H:%M:%S') if value else datetime.max
                    except (ValueError, IndexError):
                        if sort_col_id == 'creation_time': return datetime.max
                        else: return datetime.max.date()
                elif sort_col_id == 'status':
                    try:
                        return STATUS_OPTIONS.index(value)
                    except ValueError:
                        return len(STATUS_OPTIONS)
                return str(value).lower() if value is not None else ''
            tasks_to_display.sort(key=sort_key, reverse=(sort_direction == 'descending'))
        else:
            tasks_to_display.sort(key=lambda x: datetime.strptime(x.get('creation_time', '1900-01-01 00:00:00'), '%Y-%m-%d %H:%M:%S') if x.get('creation_time') else datetime.min, reverse=True)

        # 填充目前的 Treeview
        for task in tasks_to_display:
            status_display = task.get("status", "未知狀態") if task.get("status") in STATUS_OPTIONS else "未知狀態"
            due_date_display = format_date_with_weekday(task.get("due_date"))
            note_preview = (
                str(task.get("note", "")[:60].replace("\n", " ") + "...")
                if len(str(task.get("note", ""))) > 60
                else str(task.get("note", "")).replace("\n", " ")
            )
            creation_time_display = format_datetime(task.get("creation_time"))

            treeview.insert(
                "",
                "end",
                iid=str(task.get("id")),
                values=(
                    creation_time_display,
                    task.get("description", ""),
                    due_date_display,
                    status_display,
                    note_preview,
                ),
            )

        # 狀態列顯示
        on_hold_count = sum(1 for task in self.tasks if task.get('status') == 'On hold')
        if not self.show_on_hold and on_hold_count > 0:
            self.status_label.configure(text=f"已隱藏 {on_hold_count} 個 On hold 項目。")
        else:
            current_status_text = self.status_label.cget("text")
            if "儲存中" not in current_status_text and "已儲存" not in current_status_text:
                self.update_status(f"總計 {len(self.tasks)} 個待辦事項。")

        self.log_operation(f"Treeview 已重新填充並應用過濾/排序 ({len(tasks_to_display)}/{len(self.tasks)} 總數顯示)。")

    def load_task_for_editing(self, event=None): # 接受 event 參數
        """從 Treeview 載入選取的任務到輸入框進行編輯"""
        # 如果 event 為 None，通過當前選中的 Tab 獲取 Treeview
        if event is None:
            current_tab = self.tab_notebook.tab(self.tab_notebook.select(), "text").lower()
            treeview = self.treeviews.get(current_tab)
            if not treeview:
                self.log_operation("無法找到對應的 Treeview。")
                return
        else:
            treeview = event.widget
            selected_items_iid = treeview.selection()
            if not selected_items_iid:
                return  # 沒有選取任何項目

        # 只載入第一個選取的項目進行編輯
        first_selected_id_str = selected_items_iid[0]
        try:
            selected_id = int(first_selected_id_str)
        except ValueError:
            self.log_operation(f"嘗試載入編輯失敗：無效的 Treeview iid {first_selected_id_str}。")
            return

        task_to_edit = next((task for task in self.tasks if task['id'] == selected_id), None)


        if task_to_edit:
            # 如果當前已經在編輯另一個任務，先取消之前的編輯狀態
            if self.editing_task_id is not None and self.editing_task_id != selected_id:
                 self.cancel_edit() # 取消之前的編輯狀態
            # 如果點擊的是當前正在編輯的任務，不做任何事
            elif self.editing_task_id == selected_id:
                 return


            # 清空輸入框
            self.clear_input_fields()

            # 填充輸入框
            self.desc_entry.insert(0, task_to_edit.get('description', ''))
            self.date_display_label.configure(text=task_to_edit.get('due_date') if task_to_edit.get('due_date') else "")
            self.note_textbox.insert("1.0", task_to_edit.get('note', ''))


            # 設定編輯模式狀態
            self.editing_task_id = task_to_edit['id']
            self.save_button.configure(text="儲存修改") # 更改按鈕文字

            # 顯示取消編輯按鈕
            # 檢查按鈕是否已經在佈局中
            if self.cancel_edit_button.winfo_manager() != 'pack':
                 self.cancel_edit_button.pack(side=tk.LEFT, padx=5, pady=5, in_=self.action_button_frame)


            self.update_status(f"載入待辦事項 (ID: {selected_id}) 進行編輯。")
            self.log_operation(f"載入任務 (ID: {selected_id}) 進行編輯。")
        else:
            self.cancel_edit() # 如果沒找到任務，確保切回新增模式
            self.log_operation(f"嘗試載入編輯失敗：找不到 ID {selected_id} 的任務。")


    def cancel_edit(self):
        """取消編輯狀態，清空輸入框，切回新增模式"""
        if self.editing_task_id is not None: # 只有在編輯模式下才執行取消
             self.editing_task_id = None
             self.clear_input_fields()
             self.save_button.configure(text="新增待辦事項") # 更改按鈕文字回新增
             # 檢查按鈕是否在佈局中才 pack_forget
             if self.cancel_edit_button.winfo_manager() == 'pack':
                 self.cancel_edit_button.pack_forget()

             # 清除 Treeview 選取狀態
             # 安全檢查，確保 task_tree 存在且不是 None
             if hasattr(self, 'task_tree') and self.task_tree is not None:
                 selected_items = self.task_tree.selection()
                 if selected_items:
                      self.task_tree.selection_remove(selected_items)

             self.update_status("編輯已取消，切回新增模式。")
             self.log_operation("編輯操作已取消。")


    def set_selected_task_status(self):
        """標記選取的待辦事項狀態為下拉選單的值"""
        # 獲取當前選中的 Tab
        current_tab = self.tab_notebook.tab(self.tab_notebook.select(), "text").lower()
        treeview = self.treeviews.get(current_tab)

        if not treeview:
            messagebox.showerror("錯誤", "無法找到對應的 Treeview。")
            return

        selected_items_iid = treeview.selection()
        if not selected_items_iid:
            messagebox.showwarning("選取錯誤", "請選取要變更狀態的待辦事項。")
            self.log_operation("嘗試變更狀態失敗：未選取項目。")
            return

        new_status = self.status_combobox.get()
        if new_status not in STATUS_OPTIONS:
            messagebox.showwarning("狀態錯誤", "選取的狀態無效。")
            self.log_operation(f"嘗試變更狀態失敗：無效狀態 '{new_status}'。")
            return

        updated_count = 0
        selected_ids_to_process = []
        for item_iid in selected_items_iid:
            try:
                task_id = int(item_iid)
                task = next((t for t in self.tasks if t['id'] == task_id), None)
                if task and task.get('status') != new_status:  # 安全獲取狀態
                    selected_ids_to_process.append(task_id)
            except ValueError:
                continue

        for task_id in selected_ids_to_process:
            task = next((t for t in self.tasks if t['id'] == task_id), None)
            if task:
                task['status'] = new_status
                updated_count += 1
                if self.editing_task_id == task_id:
                    self.after(10, lambda: self.load_task_for_editing(None))  # 延遲載入編輯區

        if updated_count > 0:
            self._start_save_thread(
                tasks_to_save=list(self.tasks),
                message=f"變更狀態並儲存中 ({updated_count}個任務)...",
                completion_callback=lambda success, err: self.after(0, self._on_save_complete_gui, success, err, f"{updated_count} 個待辦事項狀態已變更為 '{new_status}'。")
            )
            self.log_operation(f"變更了 {updated_count} 個任務的狀態為 '{new_status}'。")
        else:
            self.update_status(f"選取的待辦事項狀態已是 '{new_status}'，無需變更。")
            self.log_operation(f"選取的任務狀態已是 '{new_status}'，未執行變更。")


    def delete_selected_task(self):
        """刪除選取的待辦事項"""
        # 取得目前選中的 Tab
        current_tab_text = self.tab_notebook.tab(self.tab_notebook.select(), "text")
        current_tab = current_tab_text.strip()
        # 取得標準格式的狀態名稱（首字大寫），或 "all"
        if current_tab.lower() == "all":
            treeview = self.treeviews.get("all")
        else:
            current_tab_status = next((s for s in STATUS_OPTIONS if s.lower() == current_tab.lower()), None)
            treeview = self.treeviews.get(current_tab_status)
        if not treeview:
            messagebox.showerror("錯誤", "無法找到對應的 Treeview。")
            self.log_operation("嘗試刪除任務失敗：找不到對應的 Treeview。")
            return

        selected_items_iid = treeview.selection()
        if not selected_items_iid:
            messagebox.showwarning("選取錯誤", "請選取要刪除的待辦事項。")
            self.log_operation("嘗試刪除任務失敗：未選取項目。")
            return

        if not messagebox.askyesno("確認刪除", f"確定要刪除這 {len(selected_items_iid)} 個待辦事項嗎？"):
            self.log_operation("取消刪除任務。")
            return

        selected_task_ids = []
        for item_iid in selected_items_iid:
            try:
                selected_task_ids.append(int(item_iid))
            except ValueError:
                continue

        initial_task_count = len(self.tasks)
        self.tasks = [task for task in self.tasks if task['id'] not in selected_task_ids]

        deleted_count = initial_task_count - len(self.tasks)

        if deleted_count > 0:
            if self.editing_task_id is not None and self.editing_task_id in selected_task_ids:
                self.cancel_edit()

            self._start_save_thread(
                tasks_to_save=list(self.tasks),
                message=f"刪除任務並儲存中 ({deleted_count}個任務)...",
                completion_callback=lambda success, err: self.after(0, self._on_save_complete_gui, success, err, f"{deleted_count} 個待辦事項已刪除。")
            )
            self.clear_details_display()
            self.log_operation(f"刪除了 {deleted_count} 個任務。")
        else:
            self.update_status("沒有任務被刪除。")
            self.log_operation("沒有選取到有效的任務進行刪除。")


    def display_selected_task_details(self, event=None):
        """在 Treeview 選取項目時顯示詳細資訊"""
        # 獲取當前觸發事件的 Treeview
        treeview = event.widget
        selected_items_iid = treeview.selection()
        if not selected_items_iid:
            if self.editing_task_id is None:
                self.clear_details_display()
            return

        first_selected_id_str = selected_items_iid[0]
        try:
            first_selected_id = int(first_selected_id_str)
        except ValueError:
            self.clear_details_display()
            return

        selected_task = next((task for task in self.tasks if task['id'] == first_selected_id), None)

        if selected_task:
            self.display_task_details(selected_task)
        else:
            self.clear_details_display()


    def display_task_details(self, task):
        """將指定 task 的詳細資訊顯示在詳細資訊區域"""
        self.details_creation_time_value.configure(text=format_datetime(task.get('creation_time')))
        self.details_desc_value.configure(text=task.get('description', ''))
        self.details_date_value.configure(text=format_date_with_weekday(task.get('due_date')))
        self.details_status_value.configure(text=task.get('status', '未知狀態') if task.get('status') in STATUS_OPTIONS else "未知狀態")

        # 更新狀態下拉選單的值
        if task.get('status') in STATUS_OPTIONS:
            self.status_combobox.set(task.get('status'))
        else:
            self.status_combobox.set("未知狀態")
        
        self.details_note_textbox.configure(state="normal")
        self.details_note_textbox.delete("1.0", tk.END)
        self.details_note_textbox.insert("1.0", task.get('note', ''))

        self.find_and_tag_urls(self.details_note_textbox)

        self.details_note_textbox.configure(state="disabled")

        # self.details_image_value.configure(text=task.get('image_path', '') if task.get('image_path') else "無圖片")


    def clear_details_display(self):
        """清空詳細資訊顯示區域"""
        self.details_creation_time_value.configure(text="")
        self.details_desc_value.configure(text="")
        self.details_date_value.configure(text="")
        self.details_status_value.configure(text="")
        self.details_note_textbox.configure(state="normal")
        self.details_note_textbox.delete("1.0", tk.END)
        self.details_note_textbox.tag_remove("url", "1.0", tk.END)
        self.details_note_textbox.configure(state="disabled")
        # self.details_image_value.configure(text="")


    def find_and_tag_urls(self, textbox):
        """在 Textbox 中查找 URL 並應用超連結標籤"""
        textbox.configure(state="normal")  # 確保可以操作
        textbox.tag_remove("url", "1.0", tk.END)
        textbox.tag_configure("url", foreground="blue", underline=True)
        url_pattern = re.compile(r'https?://[^\s]+')
        content = textbox.get("1.0", tk.END).strip()

        for match in url_pattern.finditer(content):
            start_char_index = match.start()
            end_char_index = match.end()
            try:
                start_index = textbox.index(f"1.0 + {start_char_index}c")
                end_index = textbox.index(f"1.0 + {end_char_index}c")
            except tk.TclError:
                continue

            textbox.tag_add("url", start_index, end_index)
            url = match.group(0)
            textbox.tag_bind("url", "<Button-1>", lambda e, target_url=url: self.open_url(target_url))
            textbox.tag_bind("url", "<Enter>", lambda e: textbox.config(cursor="hand2"))
            textbox.tag_bind("url", "<Leave>", lambda e: textbox.config(cursor=""))

        # 添加右鍵菜單
        textbox.bind("<Button-3>", lambda e: self.show_context_menu(e, textbox))
        textbox.configure(state="disabled")  # 設置回唯讀

    def show_context_menu(self, event, textbox):
        """顯示右鍵菜單，提供複製功能"""
        textbox.configure(state="normal")  # 臨時設置為可操作
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="複製", command=lambda: self.copy_with_links(textbox))
        menu.post(event.x_root, event.y_root)
        textbox.configure(state="disabled")  # 恢復為唯讀

    def copy_with_links(self, textbox):
        """複製包含超連結的文字到剪貼板"""
        content = textbox.get("1.0", tk.END).strip()
        url_pattern = re.compile(r'https?://[^\s]+')
        urls = []

        for match in url_pattern.finditer(content):
            urls.append(match.group(0))

        # 將文字和超連結一起複製到剪貼板
        clipboard_content = content
        if urls:
            clipboard_content += "\n\n超連結:\n" + "\n".join(urls)
            
        self.clipboard_clear()
        self.clipboard_append(clipboard_content)
        self.update_status("已複製文字和超連結到剪貼板")


    def open_url(self, url):
        """打開指定的 URL"""
        try:
            webbrowser.open_new_tab(url)
            self.update_status(f"已打開網址: {url}")
            self.log_operation(f"已打開網址: {url}")
        except Exception as e:
            messagebox.showerror("打開網址錯誤", f"無法打開網址 {url}: {e}")
            self.log_operation(f"打開網址失敗 {url}: {e}")

    # ****** 匯出到 Excel 的方法 ******
    def export_to_excel(self):
        """將待辦事項匯出為 Excel 檔案 (.xlsx)"""
        if not self.tasks:
            messagebox.showinfo("匯出", "目前沒有待辦事項可匯出。")
            self.log_operation("嘗試匯出到 Excel 失敗：沒有待辦事項。")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="匯出待辦事項為 Excel"
        )

        if not filepath:
            self.update_status("匯出已取消。")
            self.log_operation("匯出到 Excel 操作已取消。")
            return

        self.update_status(f"匯出到 Excel 中...")
        self.log_operation(f"開始匯出到 Excel 檔案：{filepath}")
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "待辦事項"

            headers = ["ID", "內容", "到期日", "狀態", "備註/網址", "建立時間"] # 增加 ID 和建立時間列
            sheet.append(headers)

            for task in self.tasks:
                task_id = task.get('id', '')
                description = task.get('description', '')
                due_date = task.get('due_date') if task.get('due_date') else ""
                status = task.get('status') if task.get('status') in STATUS_OPTIONS else "未知狀態"
                note = task.get('note') if task.get('note') is not None else ""
                creation_time = task.get('creation_time') if task.get('creation_time') else ""

                sheet.append([task_id, description, due_date, status, note, creation_time])

            workbook.save(filepath)

            self.update_status(f"待辦事項已成功匯出到 {filepath}")
            messagebox.showinfo("匯出成功", f"待辦事項已成功匯出到\n{filepath}")
            self.log_operation(f"成功匯出到 Excel 檔案：{filepath}")

        except Exception as e:
            self.update_status(f"匯出 Excel 時發生錯誤: {e}")
            messagebox.showerror("匯出錯誤", f"匯出 Excel 時發生錯誤:\n{e}")
            self.log_operation(f"匯出到 Excel 失敗：{e}")


    def clear_input_fields(self):
        """清空輸入框內容"""
        self.desc_entry.delete(0, tk.END)
        self.date_display_label.configure(text=datetime.now().strftime('%Y-%m-%d'))
        self.note_textbox.delete("1.0", tk.END)


    def update_status(self, message):
        """更新狀態訊息欄"""
        self.status_label.configure(text=message)


# 主程式入口
if __name__ == "__main__":
    app = TodoApp()
    app.mainloop()