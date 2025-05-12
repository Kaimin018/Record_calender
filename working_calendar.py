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

# 儲存待辦事項的檔案名稱
DATA_FILE = 'todo_calendar_advanced_gui.json'

# 設定 customtkinter 的外觀模式 ('System'/'Dark'/'Light') 和顏色主題 ('blue'/'green'/'dark-blue')
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

# 定義所有可能的狀態
STATUS_OPTIONS = ["pending", "completed", "cancelled", "on hold", "in progress"]

# 將 save_tasks 函數放在類別外面，方便執行緒直接調用
def save_tasks(tasks):
    """將待辦事項儲存到檔案"""
    try:
        # 在儲存前，確保 ID 可以被序列化 (這裡 ID 已經是數字，沒問題)
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(tasks, f, indent=4, ensure_ascii=False)
        # print("Tasks saved.") # 調試用
        return True # 儲存成功
    except Exception as e:
        print(f"Error saving tasks: {e}") # 在控制台打印錯誤
        return False, e # 儲存失敗，返回錯誤訊息


def load_tasks():
    """從檔案載入待辦事項"""
    tasks = []
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                tasks = json.load(f)
            # 確保每個任務都有一個唯一的 ID 和必要的欄位
            current_id = 0
            valid_tasks = []
            for task in tasks:
                 # 檢查是否是有效的字典且至少有 description
                 if isinstance(task, dict) and 'description' in task:
                    # 賦予或確保有 ID
                    if 'id' not in task or not isinstance(task['id'], (int, float)):
                        task['id'] = current_id # 如果沒有 ID 或不是數字，分配新的
                    # 確保 ID 是整數用於比較
                    task['id'] = int(task['id'])

                    # 確保其他欄位存在
                    if 'due_date' not in task:
                        task['due_date'] = None
                    if 'status' not in task or task['status'] not in STATUS_OPTIONS:
                         task['status'] = 'pending' # 補上狀態或使用預設狀態
                    if 'note' not in task:
                         task['note'] = '' # 補上 note 欄位 (空字串)
                    elif task['note'] is None: # 處理 note 為 None 的情況
                         task['note'] = ''

                    valid_tasks.append(task)
                    current_id = max(current_id, task['id']) # 找出當前最大 ID

            tasks = valid_tasks

            # 為了新增時分配新 ID，需要一個計數器
            load_tasks.counter = current_id + 1 if tasks else 0 # 設定一個函式屬性來儲存下一個可用的 ID
        except json.JSONDecodeError:
            messagebox.showwarning("警告", "無法解析待辦事項檔案，將建立新的檔案。")
            load_tasks.counter = 0
        except Exception as e:
            messagebox.showerror("錯誤", f"載入檔案時發生錯誤: {e}")
            load_tasks.counter = 0
    else:
         load_tasks.counter = 0 # 如果檔案不存在，從 0 開始計數
    return tasks

# 載入時初始化計數器
load_tasks.counter = 0


def get_next_id():
    """取得下一個可用的唯一 ID"""
    # 這裡直接用一個簡單的遞增計數器
    task_id = load_tasks.counter
    load_tasks.counter += 1
    return task_id


class TodoApp(customtkinter.CTk): # 繼承 customtkinter.CTk
    def __init__(self):
        super().__init__() # 初始化父類 (customtkinter.CTk)

        self.title("待辦事項 & 行事曆工具")
        self.geometry("1000x650") # 設定視窗初始大小，留更多空間

        self.tasks = load_tasks()
        self.editing_task_id = None # 用於儲存當前正在編輯的任務 ID，None 表示新增模式
        self.save_thread = None # 用於追蹤當前的儲存執行緒

        # --- GUI 元件設定 ---
        # ****** 將 widget 創建和佈局延遲到 _initialize_gui 中 ******
        # self.create_widgets()
        # self.layout_widgets()
        # self.populate_treeview()

        # ****** 綁定 Treeview 選取事件和雙擊事件也移到 _initialize_gui 中 ******
        # self.task_tree.bind("<<TreeviewSelect>>", self.display_selected_task_details)
        # self.task_tree.bind("<Double-1>", self.load_task_for_editing)


        # ****** 綁定 Ctrl+S 儲存快捷鍵和 Enter 鍵處理，這兩個綁定可以在 __init__ 中，因為綁定在 root 上 ******
        self.bind_all("<Control-KeyPress-s>", self.save_tasks_shortcut)
        self.bind_all("<Return>", self.handle_return_key)

        # ****** 在 __init__ 完成後，使用 after(0) 安排 GUI 的初始化 ******
        self.after(0, self._initialize_gui)


    def _initialize_gui(self):
        """Helper method to create and layout widgets and populate the treeview after the main window is ready."""
        self.create_widgets()
        self.layout_widgets()
        self.after(0, self.populate_treeview) # Populate after layout is done

        # ****** 將 Treeview 選取和雙擊事件綁定移到這裡，確保 task_tree 物件已經存在 ******
        self.task_tree.bind("<<TreeviewSelect>>", self.display_selected_task_details)
        self.task_tree.bind("<Double-1>", self.load_task_for_editing)


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
        """創建所有 GUI 元件"""
        # 輸入區域 Frame (使用 CTkFrame)
        self.input_frame = customtkinter.CTkFrame(self, corner_radius=10)

        self.desc_label = customtkinter.CTkLabel(self.input_frame, text="內容:")
        self.desc_entry = customtkinter.CTkEntry(self.input_frame, width=300)

        # 日期選擇區域
        self.date_label = customtkinter.CTkLabel(self.input_frame, text="到期日:")
        # 日期顯示 Label，用於顯示 Calendar 選取的日期
        self.date_display_label = customtkinter.CTkLabel(self.input_frame, text=datetime.now().strftime('%Y-%m-%d'), width=100, anchor=tk.W)
        self.select_date_button = customtkinter.CTkButton(self.input_frame, text="選取日期", command=self.open_calendar_dialog, width=100)


        # Note 輸入區域
        self.note_label = customtkinter.CTkLabel(self.input_frame, text="備註/網址:")
        # 使用 CTkTextbox 方便輸入多行和長網址
        self.note_textbox = customtkinter.CTkTextbox(self.input_frame, width=300, height=50)


        # 按鈕區域 (放在 input_frame 內部)
        self.action_button_frame = customtkinter.CTkFrame(self.input_frame, fg_color="transparent") # 使用透明框架
        self.save_button = customtkinter.CTkButton(self.action_button_frame, text="新增待辦事項", command=lambda: self.save_task_gui) # 這個按鈕用於新增或儲存
        # 取消編輯按鈕，初始不佈局 (pack_forget 稍後在 layout_widgets 中處理)
        self.cancel_edit_button = customtkinter.CTkButton(self.action_button_frame, text="取消編輯", command=lambda: self.cancel_edit(), fg_color="gray", hover_color="darkgray")


        # Treeview 顯示區域 Frame (使用 CTkFrame 包裹 ttk.Treeview)
        self.tree_frame = customtkinter.CTkFrame(self, corner_radius=10)
        self.tree_frame.grid_columnconfigure(0, weight=1) # 讓 Treeview 可以橫向擴展
        self.tree_frame.grid_rowconfigure(0, weight=1) # 讓 Treeview 可以縱向擴展

        self.tree_scrollbar_y = ttk.Scrollbar(self.tree_frame)
        self.tree_scrollbar_x = ttk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL)

        self.task_tree = ttk.Treeview(self.tree_frame, columns=("description", "due_date", "status", "note"), show="headings",
                                       yscrollcommand=self.tree_scrollbar_y.set, xscrollcommand=self.tree_scrollbar_x.set)
        self.tree_scrollbar_y.config(command=self.task_tree.yview)
        self.tree_scrollbar_x.config(command=self.task_tree.xview)

        # 可以設定 Treeview 的樣式，例如行高或顏色
        # style = ttk.Style()
        # style.theme_use("clam") # 嘗試不同的 ttk 主題
        # style.configure("Treeview", rowheight=25) # 設定行高


        self.task_tree.heading("description", text="內容", anchor=tk.W)
        self.task_tree.heading("due_date", text="到期日", anchor=tk.CENTER)
        self.task_tree.heading("status", text="狀態", anchor=tk.CENTER)
        self.task_tree.heading("note", text="備註/網址", anchor=tk.W)


        self.task_tree.column("description", width=250, anchor=tk.W, stretch=True)
        self.task_tree.column("due_date", width=90, anchor=tk.CENTER, stretch=False)
        self.task_tree.column("status", width=80, anchor=tk.CENTER, stretch=False)
        self.task_tree.column("note", width=250, anchor=tk.W, stretch=True)


        # ****** Treeview 的綁定已移至 _initialize_gui ******
        # self.task_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        # self.tree_scrollbar_y.grid(row=0, column=1, sticky="ns")
        # self.tree_scrollbar_x.grid(row=1, column=0, sticky="ew")


        # 操作按鈕 Frame (使用 CTkFrame) - 包含狀態選擇 和 匯出按鈕
        self.button_frame = customtkinter.CTkFrame(self, corner_radius=10)

        self.status_combobox_label = customtkinter.CTkLabel(self.button_frame, text="變更狀態為:")
        self.status_combobox = customtkinter.CTkComboBox(self.button_frame, values=STATUS_OPTIONS, width=120) # 稍微加寬下拉選單
        self.status_combobox.set("completed") # 預設選中 "completed"

        self.set_status_button = customtkinter.CTkButton(self.button_frame, text="變更選取狀態", command=lambda: self.set_selected_task_status)
        self.delete_button = customtkinter.CTkButton(self.button_frame, text="刪除選取", command=lambda: self.delete_selected_task, fg_color="red", hover_color="darkred")

        # ****** 新增匯出按鈕 ******
        self.export_button = customtkinter.CTkButton(self.button_frame, text="匯出為 Excel", command=lambda: self.export_to_excel)


        # 詳細資訊顯示區域 Frame (使用 CTkFrame)
        self.details_frame = customtkinter.CTkFrame(self, corner_radius=10)
        self.details_frame.columnconfigure(1, weight=1) # 讓值欄位可以擴展

        self.details_label = customtkinter.CTkLabel(self.details_frame, text="選取任務詳細資訊", font=customtkinter.CTkFont(weight="bold"))
        self.details_label.grid(row=0, column=0, columnspan=2, pady=(5, 10), sticky="ew", padx=10)

        ttk.Separator(self.details_frame, orient='horizontal').grid(row=1, column=0, columnspan=2, sticky="ew")

        # 詳細資訊 Labels
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
        self.status_label = customtkinter.CTkLabel(self, text="", anchor=tk.W, padx=10)


    def layout_widgets(self):
        """佈局所有 GUI 元件"""
        # 使用 grid 來佈局輸入區域的元件
        self.input_frame.pack(pady=(10, 5), padx=10, fill="x", expand=False)
        self.desc_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.desc_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.date_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.date_display_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.select_date_button.grid(row=1, column=2, padx=5, pady=5, sticky="w")

        self.note_label.grid(row=2, column=0, padx=5, pady=5, sticky="nw") # 備註標籤靠左上
        self.note_textbox.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # 佈局操作按鈕框架 (在 input_frame 內部的右側)
        self.action_button_frame.grid(row=0, column=3, rowspan=3, padx=10, pady=5, sticky="ns")
        self.save_button.pack(side=tk.TOP, pady=5, fill="x") # 新增/儲存按鈕
        # self.cancel_edit_button 初始不 pack，在編輯模式下 pack 進來

        self.input_frame.grid_columnconfigure(1, weight=1) # 讓描述和備註列可以擴展
        self.input_frame.grid_columnconfigure(3, weight=0) # 操作按鈕列不擴展


        # 讓 tree_frame 在視窗中擴展
        self.tree_frame.pack(pady=5, padx=10, fill="both", expand=True)

        # ****** Treeview 自身的佈局也需要在這裡 ******
        self.task_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.tree_scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.tree_scrollbar_x.grid(row=1, column=0, sticky="ew")


        # 佈局狀態變更、刪除 和 匯出按鈕
        self.button_frame.pack(pady=5, padx=10, fill="x", expand=False)
        self.status_combobox_label.pack(side=tk.LEFT, padx=(5, 0), pady=10)
        self.status_combobox.pack(side=tk.LEFT, padx=5, pady=10)
        self.set_status_button.pack(side=tk.LEFT, padx=5, pady=10)
        self.delete_button.pack(side=tk.LEFT, padx=10, pady=10)
        # ****** 佈局匯出按鈕 ******
        self.export_button.pack(side=tk.RIGHT, padx=10, pady=10) # 放在右邊


        # 佈局詳細資訊區域
        self.details_frame.pack(pady=5, padx=10, fill="x", expand=False)
        self.details_desc_label.grid(row=2, column=0, padx=5, pady=2, sticky="nw")
        self.details_desc_value.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        self.details_date_label.grid(row=3, column=0, padx=5, pady=2, sticky="nw")
        self.details_date_value.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        self.details_status_label.grid(row=4, column=0, padx=5, pady=2, sticky="nw")
        self.details_status_value.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        self.details_note_label.grid(row=5, column=0, padx=5, pady=2, sticky="nw") # 備註標籤靠左上
        # ****** 佈局標準 Text widget 和其滾動條 ******
        self.details_note_textbox.grid(row=5, column=1, padx=(5, 0), pady=2, sticky="nsew")
        self.details_note_scrollbar_y.grid(row=5, column=2, padx=(0, 5), pady=2, sticky="ns")
        self.details_frame.grid_columnconfigure(1, weight=1) # 確保 Text widget 所在的列能擴展
        self.details_frame.grid_columnconfigure(2, weight=0) # 滾動條列不擴展


        self.status_label.pack(pady=(5, 10), padx=10, fill="x", expand=False)

        # 配置 Treeview 所在的 row 和 column 允許擴展
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # tree_frame 所在的 row


    def open_calendar_dialog(self):
        """打開日曆選擇對話框"""
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
        top.resizable(False, False) # 鎖定對話框大小

        # 初始化日曆為當前顯示的日期或今天
        try:
            initial_date_str = self.date_display_label.cget("text")
            initial_date = datetime.strptime(initial_date_str, '%Y-%m-%d').date()
        except (ValueError, tk.TclError):
             initial_date = date.today() # 如果顯示的日期無效，使用今天

        # ****** 創建 Calendar widget，並設定較大的字體 ******
        # 創建一個較大的字體對象
        calendar_font = tkfont.Font(family="Arial", size=12) # 可以根據需要調整字體和大小

        cal = Calendar(top, selectmode='day', date_pattern='yyyy-mm-dd', year=initial_date.year, month=initial_date.month, day=initial_date.day,
                       font=calendar_font, # 套用字體
                       # ttk style options can be added here, e.g., style='TCalendar'
                       )
        cal.pack(padx=10, pady=10)

        ok_button = customtkinter.CTkButton(top, text="確定", command=grab_date)
        ok_button.pack(pady=10)

        # 設置焦點到對話框
        top.after(10, top.lift) # 確保對話框在最前面
        # top.focus_force() # 強制對話框獲取焦點


    def save_task_gui(self):
        """儲存（新增或修改）待辦事項"""
        description = self.desc_entry.get().strip()
        due_date_str = self.date_display_label.cget("text").strip() # 從 Label 獲取日期
        note = self.note_textbox.get("1.0", "end-1c").strip() # 從 Textbox 獲取備註


        # 檢查是否是從備註框觸發的 Ctrl+Enter，且備註框是空的，如果是，則不新增
        # 這個檢查是在 handle_return_key 中做的，這裡再次確認以防萬一
        focused_widget = self.focus_get()
        # 如果是備註框有焦點，且描述和備註都是空的，並且當前不是編輯模式 (編輯模式允許清空後儲存)
        if focused_widget == self.note_textbox and not description and not note and self.editing_task_id is None:
             self.update_status("備註內容和描述都為空，不新增。")
             return


        if not description:
            messagebox.showwarning("輸入錯誤", "待辦事項內容不能為空。")
            return

        due_date = None
        if due_date_str:
            try:
                # 驗證日曆選取的日期格式是否正確 (理論上 Calendar widget 保證格式)
                datetime.strptime(due_date_str, '%Y-%m-%d')
                due_date = due_date_str
            except ValueError:
                 # 這作為一個保險，通常不會觸發
                messagebox.showwarning("內部錯誤", "日期格式無效。")
                return

        # 根據 editing_task_id 判斷是新增還是修改
        if self.editing_task_id is None:
            # ====== 新增模式 ======
            task = {
                'id': get_next_id(), # 分配新的唯一 ID
                'description': description,
                'due_date': due_date,
                'status': 'pending', # 新增任務預設狀態為 pending
                'note': note
            }
            self.tasks.append(task)
            message = "待辦事項已新增。"
        else:
            # ====== 修改模式 ======
            # 找到列表中對應的任務
            task_to_edit = next((task for task in self.tasks if task['id'] == self.editing_task_id), None)
            if task_to_edit:
                task_to_edit['description'] = description
                task_to_edit['due_date'] = due_date
                task_to_edit['note'] = note
                # 狀態不通過這裡修改，通過下面的按鈕修改
                message = f"待辦事項 (ID: {self.editing_task_id}) 已修改。"
            else:
                messagebox.showerror("錯誤", f"找不到 ID 為 {self.editing_task_id} 的待辦事項進行修改。")
                self.cancel_edit() # 取消編輯狀態
                return


        # ====== 將儲存操作移到執行緒中 ======
        # 儲存成功後，在執行緒回調中處理 populate_treeview, clear_input_fields, cancel_edit 等 GUI 操作
        # 注意：這些 GUI 操作必須在主執行緒中執行，所以我們在執行緒完成後，利用 self.after 將它們安排到主執行緒中執行
        self._start_save_thread(
            tasks_to_save=list(self.tasks), # 傳遞 tasks 的副本
            message=f"{'新增' if self.editing_task_id is None else '修改'}並儲存中...",
            completion_callback=lambda success, err: self.after(0, self._on_save_complete_gui, success, err, message) # 將後續 GUI 操作安排到主執行緒
        )


    # ****** 執行緒相關方法 ******
    def _worker_save(self, tasks_to_save, completion_callback):
        """Worker function to run save_tasks in a separate thread."""
        try:
            success = save_tasks(tasks_to_save) # 調用獨立的 save_tasks 函數
            # 在執行緒中調用回調函數，回調函數內部使用 self.after 確保 GUI 更新在主執行緒
            completion_callback(success, None) # 通知主執行緒儲存完成 (成功)
        except Exception as e:
            print(f"Error in save worker thread: {e}") # Log error in console
            completion_callback(False, e) # 通知主執行緒儲存完成 (失敗及錯誤)


    def _start_save_thread(self, tasks_to_save, message="儲存中...", completion_callback=None):
        """Starts the save operation in a new thread."""
        # 檢查是否已經有儲存執行緒在運行
        if self.save_thread is not None and self.save_thread.is_alive():
            self.update_status("儲存操作進行中，請稍候...")
            return # 避免啟動多個儲存操作

        # 創建並啟動新的執行緒
        # 將完成回調函數傳遞給 worker
        # target 函數不能直接是 self 的方法，因為 threading.Thread 可能在 __init__ 完成前執行，
        # 導致 self 的狀態不確定。更好的做法是將 self 傳遞給 target 函數，或者 target 函數是個靜態方法或獨立函數。
        # 這裡我們讓 _worker_save 成為類別方法並通過 lambda 傳遞 self.after 回調
        # 這裡 _worker_save 隱式通過 args=(tasks_to_save, completion_callback) 傳遞參數
        self.save_thread = threading.Thread(target=self._worker_save, args=(tasks_to_save, completion_callback))
        self.save_thread.daemon = True # 設置為守護執行緒，應用程式退出時會強制終止
        self.update_status(message) # 更新狀態列顯示「儲存中...」
        self.save_thread.start()


    def _on_save_complete_gui(self, success, error, user_message):
        """在主執行緒中執行，處理儲存完成後的 GUI 更新"""
        self.save_thread = None # 清除執行緒引用

        if success:
            # 儲存成功後執行後續的 GUI 操作
            self.after(0, self.populate_treeview) # 更新列表顯示
            self.clear_input_fields() # 清空輸入框
            self.cancel_edit() # 切換回新增模式並隱藏取消按鈕

            # 顯示操作成功的訊息
            self.update_status(user_message) # 顯示原來的成功訊息
        else:
            # 儲存失敗的處理
            self.update_status(f"儲存失敗: {error}")
            messagebox.showerror("儲存錯誤", f"儲存檔案時發生錯誤:\n{error}")


    def save_tasks_shortcut(self, event=None): # 接受 event 參數，因為綁定會傳遞
        """Ctrl+S 快捷鍵儲存任務"""
        # 如果當前已經在編輯模式，先呼叫 save_task_gui 嘗試儲存編輯
        if self.editing_task_id is not None:
            # save_task_gui 會自己啟動執行緒儲存
            self.save_task_gui()
        else:
            # 否則，直接啟動儲存執行緒
            # 傳遞 tasks 的副本和完成回調
            self._start_save_thread(
                tasks_to_save=list(self.tasks), # 傳遞 tasks 的副本
                message="儲存中 (Ctrl+S)...",
                completion_callback=lambda success, err: self.after(0, self._on_save_complete_gui, success, err, "待辦事項已儲存 (Ctrl+S)。")
            )
        return "break" # 阻止事件繼續傳播

    # ****** 新增匯出到 Excel 的方法 ******
    def export_to_excel(self):
        """將待辦事項匯出為 Excel 檔案 (.xlsx)"""
        if not self.tasks:
            messagebox.showinfo("匯出", "目前沒有待辦事項可匯出。")
            return

        # 打開檔案儲存對話框 - 這個必須在主執行緒
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="匯出待辦事項為 Excel"
        )

        if not filepath:
            self.update_status("匯出已取消。")
            return

        # Excel 寫入也可能耗時，理想情況下也應該放在執行緒中
        # 但涉及 filedialog 在主執行緒中，且 openpyxl 寫入通常比 JSON 寫入更快
        # 為了簡單性，這裡先不將 Excel 寫入執行緒化
        self.update_status(f"匯出到 Excel 中...")
        try:
            # 創建新的 Excel 工作簿
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "待辦事項"

            # 寫入表頭
            headers = ["ID", "內容", "到期日", "狀態", "備註/網址"] # 增加 ID 列
            sheet.append(headers)

            # 寫入待辦事項數據
            for task in self.tasks:
                # 處理 None 值，確保可以寫入 Excel
                task_id = task.get('id', '') # 確保 ID 存在
                description = task.get('description', '')
                due_date = task.get('due_date') if task.get('due_date') else ""
                status = task.get('status') if task.get('status') in STATUS_OPTIONS else "未知狀態"
                note = task.get('note') if task.get('note') is not None else "" # 確保 note 不是 None

                sheet.append([task_id, description, due_date, status, note]) # 寫入 ID

            # 自動調整列寬 (可選)
            # for col_num, column in enumerate(sheet.columns):
            #     max_length = 0
            #     for cell in column:
            #         try:
            #             if len(str(cell.value)) > max_length:
            #                 max_length = len(cell.value)
            #         except:
            #             pass
            #     adjusted_width = (max_length + 2)
            #     # 確保列寬不過大或過小
            #     if adjusted_width > 50: adjusted_width = 50
            #     if adjusted_width < 10: adjusted_width = 10
            #     sheet.column_dimensions[openpyxl.utils.get_column_letter(col_num + 1)].width = adjusted_width


            # 儲存工作簿
            workbook.save(filepath)

            self.update_status(f"待辦事項已成功匯出到 {filepath}")
            messagebox.showinfo("匯出成功", f"待辦事項已成功匯出到\n{filepath}")

        except Exception as e:
            self.update_status(f"匯出 Excel 時發生錯誤: {e}")
            messagebox.showerror("匯出錯誤", f"匯出 Excel 時發生錯誤:\n{e}")


    def clear_input_fields(self):
        """清空輸入框內容"""
        self.desc_entry.delete(0, tk.END)
        # 重設為今天日期
        self.date_display_label.configure(text=datetime.now().strftime('%Y-%m-%d'))
        self.note_textbox.delete("1.0", tk.END)


    def update_status(self, message):
        """更新狀態訊息欄"""
        self.status_label.configure(text=message)


# 主程式入口
if __name__ == "__main__":
    app = TodoApp()
    app.mainloop()