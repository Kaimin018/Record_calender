import tkinter as tk
from tkinter import ttk, messagebox, filedialog # 導入 filedialog
import customtkinter
from tkcalendar import Calendar
import json
import os
from datetime import datetime, date
import webbrowser
import re
import tkinter.font as tkfont
import openpyxl # 導入 openpyxl

# 儲存待辦事項的檔案名稱
DATA_FILE = 'todo_calendar_advanced_gui.json'

# 設定 customtkinter 的外觀模式 ('System'/'Dark'/'Light') 和顏色主題 ('blue'/'green'/'dark-blue')
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

# 定義所有可能的狀態
STATUS_OPTIONS = ["pending", "completed", "cancelled", "on hold", "in progress"]

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


def save_tasks(tasks):
    """將待辦事項儲存到檔案"""
    try:
        # 在儲存前，確保 ID 可以被序列化 (這裡 ID 已經是數字，沒問題)
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(tasks, f, indent=4, ensure_ascii=False)
        # print("Tasks saved.") # 調試用
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存檔案時發生錯誤: {e}")

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

        # --- GUI 元件設定 ---
        self.create_widgets()
        self.layout_widgets()

        # 初次載入時顯示待辦事項
        self.populate_treeview()

        # 綁定 Treeview 選取事件，用於顯示詳細資訊
        self.task_tree.bind("<<TreeviewSelect>>", self.display_selected_task_details)
        # 綁定 Treeview 雙擊事件，用於載入編輯
        self.task_tree.bind("<Double-1>", self.load_task_for_editing)


        # 綁定 Ctrl+S 儲存快捷鍵到整個應用程式 (使用 KeyPress 增加兼容性)
        self.bind_all("<Control-KeyPress-s>", self.save_tasks_shortcut)
        # 對於 macOS, 可以考慮同時綁定 Command+S
        # self.bind_all("<Command-KeyPress-s>", self.save_tasks_shortcut)

        # 綁定 Enter 鍵到特定的處理函數，以區分換行和儲存
        self.bind_all("<Return>", self.handle_return_key)


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
        self.save_button = customtkinter.CTkButton(self.action_button_frame, text="新增待辦事項", command=self.save_task_gui) # 這個按鈕用於新增或儲存
        # 取消編輯按鈕，初始不佈局 (pack_forget 稍後在 layout_widgets 中處理)
        self.cancel_edit_button = customtkinter.CTkButton(self.action_button_frame, text="取消編輯", command=self.cancel_edit, fg_color="gray", hover_color="darkgray")


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


        self.task_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.tree_scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.tree_scrollbar_x.grid(row=1, column=0, sticky="ew")


        # 操作按鈕 Frame (使用 CTkFrame) - 包含狀態選擇 和 匯出按鈕
        self.button_frame = customtkinter.CTkFrame(self, corner_radius=10)

        self.status_combobox_label = customtkinter.CTkLabel(self.button_frame, text="變更狀態為:")
        self.status_combobox = customtkinter.CTkComboBox(self.button_frame, values=STATUS_OPTIONS, width=120) # 稍微加寬下拉選單
        self.status_combobox.set("completed") # 預設選中 "completed"

        self.set_status_button = customtkinter.CTkButton(self.button_frame, text="變更選取狀態", command=self.set_selected_task_status)
        self.delete_button = customtkinter.CTkButton(self.button_frame, text="刪除選取", command=self.delete_selected_task, fg_color="red", hover_color="darkred")

        # ****** 新增匯出按鈕 ******
        self.export_button = customtkinter.CTkButton(self.button_frame, text="匯出為 Excel", command=self.export_to_excel)


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


        save_tasks(self.tasks)
        self.populate_treeview() # 更新顯示
        self.clear_input_fields() # 清空輸入框
        self.cancel_edit() # 儲存成功後，切換回新增模式並隱藏取消按鈕

        self.update_status(message)


    def load_task_for_editing(self, event=None): # 接受 event 參數
        """從 Treeview 載入選取的任務到輸入框進行編輯"""
        selected_items_iid = self.task_tree.selection()
        if not selected_items_iid:
            return # 沒有選取任何項目

        # 只載入第一個選取的項目進行編輯
        selected_id = int(selected_items_iid[0])
        task_to_edit = next((task for task in self.tasks if task['id'] == selected_id), None)

        if task_to_edit:
            # 如果當前已經在編輯另一個任務，先取消之前的編輯狀態
            # (只有 ID 不同時才取消)
            if self.editing_task_id is not None and self.editing_task_id != selected_id:
                 self.cancel_edit() # 取消之前的編輯狀態
            # 如果點擊的是當前正在編輯的任務，不做任何事
            elif self.editing_task_id == selected_id:
                 return


            # 清空輸入框
            self.clear_input_fields()

            # 填充輸入框
            self.desc_entry.insert(0, task_to_edit['description'])
            self.date_display_label.configure(text=task_to_edit['due_date'] if task_to_edit['due_date'] else "")
            # ****** Textbox 插入內容 ******
            self.note_textbox.insert("1.0", task_to_edit['note'])


            # 設定編輯模式狀態
            self.editing_task_id = task_to_edit['id']
            self.save_button.configure(text="儲存修改") # 更改按鈕文字

            # 顯示取消編輯按鈕
            # 檢查是否已經 pack 了，避免重複 pack
            if not any(slave for slave in self.action_button_frame.pack_slaves() if slave == self.cancel_edit_button):
                self.cancel_edit_button.pack(side=tk.LEFT, padx=5, pady=5, in_=self.action_button_frame) # pack 到 action_button_frame 裏面


            self.update_status(f"載入待辦事項 (ID: {selected_id}) 進行編輯。")
        else:
            self.cancel_edit() # 如果沒找到任務，確保切回新增模式


    def cancel_edit(self):
        """取消編輯狀態，清空輸入框，切回新增模式"""
        self.editing_task_id = None
        self.clear_input_fields()
        self.save_button.configure(text="新增待辦事項") # 更改按鈕文字回新增
        # 隱藏取消編輯按鈕
        self.cancel_edit_button.pack_forget() # 從佈局中移除按鈕

        # 清除 Treeview 選取狀態
        self.task_tree.selection_remove(self.task_tree.selection())

        self.update_status("編輯已取消，切回新增模式。")


    def populate_treeview(self):
        """清空 Treeview 並重新填入待辦事項"""
        # 清空現有內容
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)

        # 按照到期日排序 (無到期日的排在後面), 再按狀態順序, 最後按 ID
        # 使用 STATUS_OPTIONS.index 來確保狀態按照列表的順序排序
        sorted_tasks = sorted(self.tasks, key=lambda x: (x['due_date'] if x['due_date'] else '9999-12-31', STATUS_OPTIONS.index(x['status']) if x['status'] in STATUS_OPTIONS else len(STATUS_OPTIONS), x['id']))


        # 填充 Treeview
        for task in sorted_tasks:
            status_display = task['status'] if task['status'] in STATUS_OPTIONS else "未知狀態"
            due_date_display = task['due_date'] if task['due_date'] else "無到期日"
            # 在 Treeview 中顯示部分備註，防止過長
            # 替換換行符為空格，並限制長度
            note_preview = (task['note'][:60].replace('\n', ' ') + '...') if len(task['note']) > 60 else task['note'].replace('\n', ' ')

            self.task_tree.insert("", "end", iid=str(task['id']), # 使用 task['id'] 作為 Treeview item 的 iid (字串)
                                   values=(task['description'], due_date_display, status_display, note_preview),
                                   tags=(task['status'],)) # 使用狀態作為 tag


    def set_selected_task_status(self):
        """標記選取的待辦事項狀態為下拉選單的值"""
        selected_items_iid = self.task_tree.selection() # 獲取選取項目的 iid (即 task['id'])
        if not selected_items_iid:
            messagebox.showwarning("選取錯誤", "請選取要變更狀態的待辦事項。")
            return

        new_status = self.status_combobox.get()
        if new_status not in STATUS_OPTIONS:
             messagebox.showwarning("狀態錯誤", "選取的狀態無效。")
             return

        updated_count = 0
        # 複製選取的 iid 列表，避免在迴圈中修改 Treeview 導致問題
        selected_ids_to_process = [int(item_iid) for item_iid in selected_items_iid]

        for task_id in selected_ids_to_process:
            # 找到 tasks 列表中對應的 task
            # 使用 next 和生成器表達式可以更有效率地查找
            task = next((t for t in self.tasks if t['id'] == task_id), None)

            if task and task['status'] != new_status:
                task['status'] = new_status
                updated_count += 1
                # 如果正在編輯這個任務，更新編輯區的狀態顯示（雖然狀態下拉選單在下方）
                if self.editing_task_id == task_id:
                    # 重新載入編輯區，會自動更新狀態顯示
                    # 注意：這裡重新載入編輯區會清除當前在編輯區可能未儲存的改動！
                    # 一個更複雜的實現是只更新編輯區的狀態顯示，而不是重新載入所有數據
                    # 這裡為了簡單性，選擇重新載入
                    self.load_task_for_editing()


        if updated_count > 0:
             save_tasks(self.tasks)
             self.populate_treeview() # 更新顯示
             # 更新詳細資訊區域，以防選取的項目狀態改變
             # if selected_items_iid: # TreeviewSelect 事件會自動觸發 display_selected_task_details
                # self.display_selected_task_details()


             self.update_status(f"{updated_count} 個待辦事項狀態已變更為 '{new_status}'。")
             # 重新選取剛剛操作的項目，保持選取狀態（可選）
             # self.task_tree.selection_set(selected_items_iid)
        else:
             self.update_status(f"選取的待辦事項狀態已是 '{new_status}'，無需變更。")


    def delete_selected_task(self):
        """刪除選取的待辦事項"""
        selected_items_iid = self.task_tree.selection() # 獲取選取項目的 iid (即 task['id'])
        if not selected_items_iid:
            messagebox.showwarning("選取錯誤", "請選取要刪除的待辦事項。")
            return

        # 確認刪除
        if not messagebox.askyesno("確認刪除", f"確定要刪除這 {len(selected_items_iid)} 個待辦事項嗎？"):
            return

        # 將選取的 iid 轉換為 task ID 列表
        selected_task_ids = [int(item_iid) for item_iid in selected_items_iid]

        # 使用列表推導式創建一個新的列表，只包含 ID 不在 selected_task_ids 中的任務
        initial_task_count = len(self.tasks)
        self.tasks = [task for task in self.tasks if task['id'] not in selected_task_ids]

        deleted_count = initial_task_count - len(self.tasks) # 計算實際刪除數量

        save_tasks(self.tasks)
        self.populate_treeview() # 更新顯示
        self.clear_details_display() # 清空詳細資訊顯示

        # 如果當前正在編輯的任務被刪除了，取消編輯狀態
        if self.editing_task_id is not None and self.editing_task_id in selected_task_ids:
             self.cancel_edit()

        self.update_status(f"{deleted_count} 個待辦事項已刪除。")

    def display_selected_task_details(self, event=None): # 接受 event 參數，因為綁定會傳遞
        """在 Treeview 選取項目時顯示詳細資訊"""
        selected_items_iid = self.task_tree.selection()
        if not selected_items_iid:
            # 如果沒有選取任何項目，且沒有處於編輯狀態，才清空詳細資訊顯示區域
            if self.editing_task_id is None:
                 self.clear_details_display()
            return

        # 只顯示第一個選取項目的詳細資訊
        first_selected_id = int(selected_items_iid[0])

        # 找到 tasks 列表中對應的 task
        selected_task = next((task for task in self.tasks if task['id'] == first_selected_id), None)

        if selected_task:
            self.display_task_details(selected_task)
        else:
            self.clear_details_display() # 如果沒找到，清空顯示


    def display_task_details(self, task):
        """將指定 task 的詳細資訊顯示在詳細資訊區域"""
        self.details_desc_value.configure(text=task['description'])
        self.details_date_value.configure(text=task['due_date'] if task['due_date'] else "無到期日")
        self.details_status_value.configure(text=task['status'] if task['status'] in STATUS_OPTIONS else "未知狀態")

        # 清空 Textbox 並插入新內容
        # ****** 標準 tk.Text 的狀態設定是直接用 state 參數 ******
        self.details_note_textbox.configure(state="normal") # 暫時可編輯以插入內容和應用標籤
        self.details_note_textbox.delete("1.0", tk.END)
        self.details_note_textbox.insert("1.0", task['note'])

        # 查找並標記 URL
        self.find_and_tag_urls(self.details_note_textbox)

        self.details_note_textbox.configure(state="disabled") # 設回唯讀


    def clear_details_display(self):
        """清空詳細資訊顯示區域"""
        self.details_desc_value.configure(text="")
        self.details_date_value.configure(text="")
        self.details_status_value.configure(text="")
        # ****** 清空標準 tk.Text widget ******
        self.details_note_textbox.configure(state="normal")
        self.details_note_textbox.delete("1.0", tk.END)
        # 清除舊的 tag 綁定
        self.details_note_textbox.tag_remove("url", "1.0", tk.END)
        self.details_note_textbox.configure(state="disabled")


    def find_and_tag_urls(self, textbox):
        """在 Textbox 中查找 URL 並應用超連結標籤"""
        # 清除之前可能存在的 "url" 標籤
        textbox.tag_remove("url", "1.0", tk.END)

        # 設置 "url" 標籤的樣式
        textbox.tag_configure("url", foreground="blue", underline=True)

        # 使用正則表達式查找 http 或 https 開頭的 URL
        url_pattern = re.compile(r'https?://[^\s]+')

        content = textbox.get("1.0", tk.END).strip() # 獲取內容並移除末尾可能的換行

        # 查找所有匹配項
        for match in url_pattern.finditer(content):
            start_char_index = match.start() # 絕對字元開始索引
            end_char_index = match.end()   # 絕對字元結束索引

            # 將絕對字元索引轉換為 Text widget 的 "line.column" 格式
            # 使用 Text widget 的 index 方法可以處理這種轉換
            # 確保轉換是針對當前 Textbox 的內容
            try:
                 start_index = textbox.index(f"1.0 + {start_char_index}c")
                 end_index = textbox.index(f"1.0 + {end_char_index}c")
            except tk.TclError:
                 # 如果索引無效（例如內容在查找後被更改），跳過這個匹配
                 continue


            # 添加 "url" 標籤到匹配到的文本範圍
            textbox.tag_add("url", start_index, end_index)

            # 綁定滑鼠左鍵點擊事件到這個標籤
            # 使用 lambda 來確保在點擊時傳遞正確的 URL
            url = match.group(0)
            # 綁定到 Text widget 本身
            textbox.tag_bind("url", "<Button-1>", lambda e, target_url=url: self.open_url(target_url))

            # 綁定滑鼠進入和離開事件來改變光標
            # 綁定到 Text widget 本身
            textbox.tag_bind("url", "<Enter>", lambda e: textbox.config(cursor="hand2"))
            textbox.tag_bind("url", "<Leave>", lambda e: textbox.config(cursor=""))


    def open_url(self, url):
        """打開指定的 URL"""
        try:
            webbrowser.open_new_tab(url)
            self.update_status(f"已打開網址: {url}")
        except Exception as e:
            messagebox.showerror("打開網址錯誤", f"無法打開網址 {url}: {e}")

    def save_tasks_shortcut(self, event=None): # 接受 event 參數，因為綁定會傳遞
        """Ctrl+S 快捷鍵儲存任務"""
        save_tasks(self.tasks)
        self.update_status("待辦事項已儲存 (Ctrl+S)。")
        return "break" # 阻止事件繼續傳播

    # ****** 新增匯出到 Excel 的方法 ******
    def export_to_excel(self):
        """將待辦事項匯出為 Excel 檔案 (.xlsx)"""
        if not self.tasks:
            messagebox.showinfo("匯出", "目前沒有待辦事項可匯出。")
            return

        # 打開檔案儲存對話框
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="匯出待辦事項為 Excel"
        )

        if not filepath:
            self.update_status("匯出已取消。")
            return

        try:
            # 創建新的 Excel 工作簿
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "待辦事項"

            # 寫入表頭
            headers = ["內容", "到期日", "狀態", "備註/網址"]
            sheet.append(headers)

            # 寫入待辦事項數據
            for task in self.tasks:
                # 處理 None 值，確保可以寫入 Excel
                due_date = task.get('due_date') if task.get('due_date') else ""
                status = task.get('status') if task.get('status') in STATUS_OPTIONS else "未知狀態"
                note = task.get('note') if task.get('note') is not None else "" # 確保 note 不是 None

                sheet.append([task.get('description', ''), due_date, status, note])

            # 自動調整列寬 (可選)
            # for col in sheet.columns:
            #     max_length = 0
            #     column = col[0].column # Get the column number
            #     for cell in col:
            #         try: # Necessary to avoid error on empty cells
            #             if len(str(cell.value)) > max_length:
            #                 max_length = len(cell.value)
            #         except:
            #             pass
            #     adjusted_width = (max_length + 2)
            #     sheet.column_dimensions[openpyxl.utils.get_column_letter(column)].width = adjusted_width


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