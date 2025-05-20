# record_calender/gui.py

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
import customtkinter
from tkcalendar import Calendar
import os
from datetime import datetime, date
import threading
from PIL import Image, ImageTk
import platform

# 導入重構後的模組
from record_calender.task_manager import TaskManager, STATUS_OPTIONS
from record_calender.data_manager import TaskDataManager
from record_calender import utils # 導入 utils 模組

# 從 main.py 獲取 SCRIPT_DIR
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) # 現在 SCRIPT_DIR 指向 record_calender/

# 設定 customtkinter 的外觀模式和顏色主題
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

class TodoApp(customtkinter.CTk):
    def __init__(self, task_manager: TaskManager):
        super().__init__()

        self.task_manager = task_manager # 注入 TaskManager 實例
        self.log_entries = []
        self.editing_task_id = None
        self.save_thread = None
        self.show_on_hold = True
        self._sort_direction = {}
        self._sort_column = None
        self._icon_refs = [] # 儲存圖片參考

        self.title("待辦事項 & 行事曆工具")
        self.geometry("1000x650")

        self.setup_main_layout() # 獨立設置主佈局
        self.create_widgets()
        self.layout_widgets()
        self.create_menu()
        self.load_icons()

        self.tab_notebook.bind("<<NotebookTabChanged>>", lambda e: self.populate_treeview())
        
        # 綁定快捷鍵 (platform specific)
        if platform.system() == "Darwin":
            self.bind_all("<Command-KeyPress-s>", self.save_tasks_shortcut)
        else:
            self.bind_all("<Control-KeyPress-s>", self.save_tasks_shortcut)

        self.bind_all("<Return>", self.handle_return_key)

        self.log_operation("應用程式啟動")
        self.populate_treeview() # 首次啟動時填充 Treeview

    def setup_main_layout(self):
        """設置主視窗的 Canvas 和 Scrollbar"""
        self.main_canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        self.inner_frame = customtkinter.CTkFrame(self.main_canvas)
        self.inner_frame_id = self.main_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.main_scrollbar.pack(side="right", fill="y")

        self.inner_frame.bind("<Configure>", lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.main_canvas.bind_all("<MouseWheel>", lambda e: self.main_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

    def load_icons(self):
        """載入應用程式所需的圖標"""
        # 注意：圖片路徑需要根據你的 assets/ 目錄來調整
        icon_path = os.path.join(SCRIPT_DIR, '..', 'assets', 'warning.png')
        try:
            if os.path.exists(icon_path):
                original_image = Image.open(icon_path)
                resized_image = original_image.resize((16, 16), Image.Resampling.LANCZOS)
                self.warning_icon = ImageTk.PhotoImage(resized_image)
                self._icon_refs.append(self.warning_icon)
            else:
                self.warning_icon = None
                self.log_operation(f"警告圖標檔案 '{icon_path}' 未找到。")
        except Exception as e:
            self.warning_icon = None
            self.log_operation(f"載入警告圖標時發生錯誤: {e}")

    def log_operation(self, message):
        """記錄操作到日誌列表"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"{timestamp} - {message}"
        self.log_entries.append(log_entry)
        print(log_entry)

    def create_menu(self):
        """創建應用程式頂部選單"""
        if hasattr(self, 'menubar') and self.menubar:
            self.menubar.destroy()

        self.menubar = tk.Menu(self)
        self.config(menu=self.menubar)

        filemenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="檔案", menu=filemenu)
        filemenu.add_command(label="儲存 (Ctrl+S)", command=self.save_tasks_shortcut)
        filemenu.add_command(label="匯出為 Excel", command=self.export_to_excel)
        filemenu.add_separator()
        filemenu.add_command(label="結束", command=self.quit)

        viewmenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="查看", menu=viewmenu)
        viewmenu.add_command(label="操作日誌", command=self.show_log_window)

        optionsmenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="設置", menu=optionsmenu)
        self.show_on_hold_var = tk.BooleanVar(value=self.show_on_hold)
        optionsmenu.add_checkbutton(label="顯示 On hold 項目", variable=self.show_on_hold_var, command=self.toggle_show_on_hold)

    def show_log_window(self):
        """顯示操作日誌視窗"""
        log_window = customtkinter.CTkToplevel(self)
        log_window.title("操作日誌")
        log_window.geometry("600x400")

        log_textbox = customtkinter.CTkTextbox(log_window, wrap="word")
        log_textbox.pack(padx=10, pady=10, fill="both", expand=True)

        for entry in self.log_entries:
            log_textbox.insert(tk.END, entry + "\n")

        log_textbox.configure(state="disabled")
        log_window.transient(self)
        log_window.grab_set()
        log_window.after(10, log_window.lift)

    def toggle_show_on_hold(self):
        """切換顯示或隱藏 On hold 狀態的待辦事項"""
        self.show_on_hold = self.show_on_hold_var.get()
        self.populate_treeview()
        self.log_operation(f"切換顯示 On hold 項目為: {'顯示' if self.show_on_hold else '隱藏'}。")

    def handle_return_key(self, event):
        """處理 Enter 鍵事件，判斷是否觸發儲存"""
        focused_widget = self.focus_get()
        is_ctrl_pressed = (event.state & 0x4) != 0 # Windows/Linux Ctrl key state

        if focused_widget in [self.note_textbox, self.details_note_textbox]:
            if is_ctrl_pressed:
                self.save_task_gui()
                return "break"
            else:
                return None
        elif focused_widget == self.desc_entry:
            self.save_task_gui()
            return "break"
        elif focused_widget in [self.date_display_label, self.select_date_button]:
             self.save_task_gui()
             return "break"
        return None

    def create_widgets(self):
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("Arial", 12))

        self.input_frame = customtkinter.CTkFrame(self.inner_frame, corner_radius=10)
        self.desc_label = customtkinter.CTkLabel(self.input_frame, text="內容:")
        self.desc_entry = customtkinter.CTkEntry(self.input_frame, width=300)
        self.date_label = customtkinter.CTkLabel(self.input_frame, text="到期日:")
        self.date_display_label = customtkinter.CTkLabel(self.input_frame, text=datetime.now().strftime('%Y-%m-%d'), width=100, anchor=tk.W)
        self.select_date_button = customtkinter.CTkButton(self.input_frame, text="選取日期", command=self.open_calendar_dialog)
        self.note_label = customtkinter.CTkLabel(self.input_frame, text="備註/網址:")
        self.note_textbox = customtkinter.CTkTextbox(self.input_frame, width=300, height=50)

        self.action_button_frame = customtkinter.CTkFrame(self.input_frame, fg_color="transparent")
        self.save_button = customtkinter.CTkButton(self.action_button_frame, text="新增待辦事項", command=self.save_task_gui)
        self.cancel_edit_button = customtkinter.CTkButton(self.action_button_frame, text="取消編輯", command=self.cancel_edit, fg_color="gray", hover_color="darkgray")

        self.tab_notebook = ttk.Notebook(self.inner_frame)
        self.tabs = {}
        for status in STATUS_OPTIONS:
            tab_frame = customtkinter.CTkFrame(self.tab_notebook)
            self.tab_notebook.add(tab_frame, text=status.capitalize())
            self.tabs[status] = tab_frame

        self.all_tab_frame = customtkinter.CTkFrame(self.tab_notebook)
        self.tab_notebook.add(self.all_tab_frame, text="All")
        self.tabs["all"] = self.all_tab_frame

        self.create_treeview_widgets()
        
        self.button_frame = customtkinter.CTkFrame(self.inner_frame, corner_radius=10)
        self.status_combobox_label = customtkinter.CTkLabel(self.button_frame, text="變更狀態為:")
        self.status_combobox = customtkinter.CTkComboBox(self.button_frame, values=STATUS_OPTIONS, width=120)
        self.status_combobox.set("Completed")

        self.set_status_button = customtkinter.CTkButton(self.button_frame, text="變更選取狀態", command=self.set_selected_task_status)
        self.delete_button = customtkinter.CTkButton(self.button_frame, text="刪除選取", command=self.delete_selected_task, fg_color="red", hover_color="darkred")
        self.export_button = customtkinter.CTkButton(self.button_frame, text="匯出為 Excel", command=self.export_to_excel)

        self.details_frame = customtkinter.CTkFrame(self.inner_frame, corner_radius=10)
        self.details_frame.columnconfigure(1, weight=1)

        self.details_label = customtkinter.CTkLabel(self.details_frame, text="選取任務詳細資訊", font=customtkinter.CTkFont(weight="bold"))
        self.details_creation_time_label = customtkinter.CTkLabel(self.details_frame, text="建立時間:", anchor=tk.W)
        self.details_creation_time_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W)
        self.details_desc_label = customtkinter.CTkLabel(self.details_frame, text="內容:", anchor=tk.W)
        self.details_desc_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W, wraplength=450)
        self.details_date_label = customtkinter.CTkLabel(self.details_frame, text="到期日:", anchor=tk.W)
        self.details_date_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W)
        self.details_status_label = customtkinter.CTkLabel(self.details_frame, text="狀態:", anchor=tk.W)
        self.details_status_value = customtkinter.CTkLabel(self.details_frame, text="", anchor=tk.W)
        self.details_note_label = customtkinter.CTkLabel(self.details_frame, text="備註/網址:", anchor=tk.NW)
        self.details_note_textbox = tk.Text(self.details_frame, wrap="word", height=4, state="disabled",
                                            bg=self._apply_appearance_mode(customtkinter.ThemeManager.theme["CTkEntry"]["fg_color"]),
                                            fg=self._apply_appearance_mode(customtkinter.ThemeManager.theme["CTkEntry"]["text_color"]),
                                            relief="flat", padx=0, pady=0)
        self.details_note_scrollbar_y = ttk.Scrollbar(self.details_frame, command=self.details_note_textbox.yview)
        self.details_note_textbox.configure(yscrollcommand=self.details_note_scrollbar_y.set)

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

            treeview.heading("creation_time", text="建立時間", anchor=tk.W, command=lambda c="creation_time": self.on_treeview_heading_click(treeview, c))
            treeview.heading("description", text="內容", anchor=tk.W, command=lambda c="description": self.on_treeview_heading_click(treeview, c))
            treeview.heading("due_date", text="到期日", anchor=tk.CENTER, command=lambda c="due_date": self.on_treeview_heading_click(treeview, c))
            treeview.heading("status", text="狀態", anchor=tk.CENTER, command=lambda c="status": self.on_treeview_heading_click(treeview, c))
            treeview.heading("note", text="備註/網址", anchor=tk.W, command=lambda c="note": self.on_treeview_heading_click(treeview, c))

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

            treeview.bind("<Double-1>", self.load_task_for_editing)
            treeview.bind("<<TreeviewSelect>>", self.display_selected_task_details)

            self.treeviews[status] = treeview
            
    def layout_widgets(self):
        self.input_frame.pack(pady=(10, 5), padx=10, fill="x", expand=False)
        self.button_frame.pack(pady=5, padx=10, fill="x", expand=False)
        self.tab_notebook.pack(pady=10, padx=10, fill="both", expand=True)
        self.details_frame.pack(pady=5, padx=10, fill="x", expand=False)

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

        self.status_combobox_label.pack(side=tk.LEFT, padx=(5, 0), pady=10)
        self.status_combobox.pack(side=tk.LEFT, padx=5, pady=10)
        self.set_status_button.pack(side=tk.LEFT, padx=5, pady=10)
        self.delete_button.pack(side=tk.LEFT, padx=10, pady=10)
        self.export_button.pack(side=tk.RIGHT, padx=10, pady=10)

        self.details_label.grid(row=0, column=0, columnspan=2, pady=(5, 10), sticky="ew", padx=10)
        ttk.Separator(self.details_frame, orient='horizontal').grid(row=1, column=0, columnspan=2, sticky="ew")
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

    def open_calendar_dialog(self):
        """打開日曆選擇對話框，並嘗試定位在主視窗旁邊"""
        def grab_date():
            try:
                selected_date = cal.selection_get()
                self.date_display_label.configure(text=selected_date.strftime('%Y-%m-%d'))
            except Exception as e:
                 messagebox.showerror("日期選取錯誤", f"無法獲取選取的日期: {e}")
            finally:
                top.destroy()

        top = customtkinter.CTkToplevel(self)
        top.title("選取日期")
        top.transient(self)
        top.grab_set()
        top.resizable(True, True)

        try:
            initial_date_str = self.date_display_label.cget("text")
            initial_date = datetime.strptime(initial_date_str, '%Y-%m-%d').date()
        except (ValueError, tk.TclError):
             initial_date = date.today()

        calendar_font = tkfont.Font(family="Arial", size=12)
        cal = Calendar(top, selectmode='day', date_pattern='yyyy-mm-dd', year=initial_date.year, month=initial_date.month, day=initial_date.day, font=calendar_font)
        cal.pack(padx=10, pady=10)

        ok_button = customtkinter.CTkButton(top, text="確定", command=grab_date)
        ok_button.pack(pady=10)

        self.update_idletasks()
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()

        top.update_idletasks()
        dialog_width = top.winfo_width()

        x_pos = main_x + main_width + 10
        y_pos = main_y

        screen_width = self.winfo_screenwidth()
        if x_pos + dialog_width > screen_width:
            x_pos = main_x - dialog_width - 10
            if x_pos < 0:
                 x_pos = main_x + (main_width - dialog_width) // 2
                 y_pos = main_y + 50

        top.geometry(f"+{x_pos}+{y_pos}")
        top.after(10, top.lift)

    def save_task_gui(self):
        """儲存（新增或修改）待辦事項的 GUI 介面邏輯"""
        description = self.desc_entry.get().strip()
        due_date_str = self.date_display_label.cget("text").strip()
        note = self.note_textbox.get("1.0", "end-1c").strip()

        focused_widget = self.focus_get()
        if focused_widget in [self.note_textbox, self.details_note_textbox] and not description and not note and self.editing_task_id is None:
             self.update_status("備註內容和描述都為空，不新增。")
             self.log_operation("嘗試新增待辦事項失敗：備註和內容為空。")
             return

        if not description and not note:
            messagebox.showwarning("輸入錯誤", "待辦事項內容和備註不能都為空。")
            self.log_operation("嘗試新增待辦事項失敗：內容和備註為空。")
            return

        due_date = due_date_str if due_date_str else None

        try:
            if self.editing_task_id is None:
                # 調用 TaskManager 的新增方法
                new_task = self.task_manager.add_task(description, due_date, note)
                message = f"事件 '{new_task['description']}' 已於 {new_task['creation_time'].split(' ')[0]} 新增成功！"
                log_msg = f"新增待辦事項 (ID: {new_task['id']}): '{description}'"
            else:
                # 調用 TaskManager 的更新方法
                updated_task = self.task_manager.update_task(self.editing_task_id, description, due_date, note=note)
                if updated_task:
                    message = f"待辦事項 (ID: {self.editing_task_id}) 已修改。"
                    log_msg = f"修改待辦事項 (ID: {self.editing_task_id}): '{description}'"
                else:
                    messagebox.showerror("錯誤", f"找不到 ID 為 {self.editing_task_id} 的待辦事項進行修改。")
                    self.log_operation(f"嘗試修改待辦事項失敗：找不到 ID {self.editing_task_id}。")
                    self.cancel_edit()
                    return
        except ValueError as e:
            messagebox.showwarning("輸入錯誤", str(e))
            self.log_operation(f"新增/修改待辦事項失敗：{e}")
            return
        except Exception as e:
            messagebox.showerror("錯誤", f"處理任務時發生未知錯誤: {e}")
            self.log_operation(f"處理任務時發生未知錯誤: {e}")
            return

        self.log_operation(log_msg)
        self.populate_treeview() # 更新 GUI
        self.clear_input_fields()
        self.cancel_edit()
        self.update_status(message)


    def save_tasks_shortcut(self, event=None):
        """Ctrl+S 快捷鍵儲存任務 (現在由 GUI 內部處理，實際儲存調用 TaskManager)"""
        # 由於 TaskManager 的方法在執行時會自動觸發 TaskDataManager 的儲存，
        # 這裡的 "儲存" 快捷鍵主要用於當沒有在編輯模式時，
        # 強制觸發一次 TaskManager 對數據的持久化（雖然 TaskManager 方法內部已經做了）。
        # 或者，可以設計一個 TaskManager.save_all() 方法。
        # 目前的設計是每次 add/update/delete 都會儲存，所以這裡可以讓它更單純。
        
        # 這裡可以選擇：
        # 1. 如果處於編輯模式，則提交編輯 (等同於點擊儲存按鈕)
        # 2. 如果非編輯模式，則只是觸發一次 TaskManager 的內部儲存 (如果 TaskManager 有一個公開的 save_all 方法)
        # 鑑於 TaskManager 的方法已經自動儲存，這裡可以簡化為：
        
        if self.editing_task_id is not None:
            self.save_task_gui() # 提交當前編輯
            self.log_operation("通過快捷鍵提交編輯並儲存。")
        else:
            # 如果 TaskManager 有一個顯式的 save_all() 方法，可以在這裡呼叫它
            # 例如: self.task_manager.data_manager.save_tasks(self.task_manager.get_tasks())
            # 但由於每個操作都自動儲存，這一步可能不那麼必要，除非你想強制寫入而不觸發任何任務操作
            self.update_status("數據已自動儲存。")
            self.log_operation("通過快捷鍵觸發檢查儲存。")
        return "break"

    def on_treeview_heading_click(self, treeview, column_name):
        """處理 Treeview 標頭點擊事件，實現排序"""
        if self._sort_column == column_name:
            self._sort_direction[column_name] = 'descending' if self._sort_direction.get(column_name, 'ascending') == 'ascending' else 'ascending'
        else:
            self._sort_column = column_name
            self._sort_direction = {column_name: 'ascending'}

        self.populate_treeview()

        # 更新標頭箭頭指示
        # 這裡需要獲取所有 Treeview 的 heading，以便清除之前的箭頭
        all_treeviews = list(self.treeviews.values())
        for tv in all_treeviews:
            for col_id in ("creation_time", "description", "due_date", "status", "note"):
                current_text_parts = tv.heading(col_id, 'text').split(' ')
                current_text = current_text_parts[0] if current_text_parts else ""
                tv.heading(col_id, text=current_text)

        arrow = ' ▲' if self._sort_direction[column_name] == 'ascending' else ' ▼'
        current_text = treeview.heading(column_name, 'text').split(' ')[0]
        treeview.heading(column_name, text=f"{current_text}{arrow}")

        self.log_operation(f"按 '{column_name}' 欄位進行了 {'遞減' if self._sort_direction[column_name] == 'descending' else '遞增'} 排序。")

    def populate_treeview(self):
        """只清空並填充目前顯示的 Treeview"""
        current_tab_text = self.tab_notebook.tab(self.tab_notebook.select(), "text")
        current_tab_status = next((s for s in STATUS_OPTIONS if s.lower() == current_tab_text.strip().lower()), None)

        if current_tab_text.lower().strip() == "all":
            all_tasks = self.task_manager.get_all_tasks_sorted(self._sort_column, self._sort_direction.get(self._sort_column, 'ascending'))
            tasks_to_display = [task for task in all_tasks if self.show_on_hold or task.get('status') != 'On hold']
            treeview = self.treeviews["all"]
        else:
            # 獲取特定狀態的任務，並應用排序
            filtered_tasks = self.task_manager.get_tasks_by_status(current_tab_status)
            # 在這裡，你可以對 filtered_tasks 應用排序，如果 TaskManager 沒有提供帶狀態的排序方法
            tasks_to_display = filtered_tasks
            if self._sort_column:
                # 重新利用 TaskManager 的排序邏輯來排序當前過濾後的任務
                # 這是手動應用排序，如果 TaskManager 的 get_tasks_by_status 方法本身就能排序更好
                # 但為了分離職責，這裡只針對已經獲取到的列表進行排序。
                sort_col_id = self._sort_column
                sort_direction = self._sort_direction.get(sort_col_id, 'ascending')
                
                # 這是 TaskManager.get_all_tasks_sorted 中的排序邏輯，複製過來
                def sort_key(task):
                    value = task.get(sort_col_id)
                    if sort_col_id in ['due_date', 'creation_time']:
                        try:
                            if sort_col_id == 'due_date':
                                return datetime.strptime(str(value).split(' ')[0], '%Y-%m-%d').date() if value else datetime.max.date()
                            if sort_col_id == 'creation_time':
                                return datetime.strptime(str(value), '%Y-%m-%d %H:%M:%S') if value else datetime.max
                        except (ValueError, IndexError):
                            return datetime.max.date() if sort_col_id == 'due_date' else datetime.max
                    elif sort_col_id == 'status':
                        try:
                            return STATUS_OPTIONS.index(value)
                        except ValueError:
                            return len(STATUS_OPTIONS)
                    return str(value).lower() if value is not None else ''
                tasks_to_display.sort(key=sort_key, reverse=(sort_direction == 'descending'))

            treeview = self.treeviews.get(current_tab_status)
            if treeview is None:
                return

        for item in treeview.get_children():
            treeview.delete(item)

        for task in tasks_to_display:
            status_display = task.get("status", "未知狀態") if task.get("status") in STATUS_OPTIONS else "未知狀態"
            due_date_display = utils.format_date_with_weekday(task.get("due_date"))
            note_preview = (
                str(task.get("note", "")[:60].replace("\n", " ") + "...")
                if len(str(task.get("note", ""))) > 60
                else str(task.get("note", "")).replace("\n", " ")
            )
            creation_time_display = utils.format_datetime(task.get("creation_time"))

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

        on_hold_count = sum(1 for task in self.task_manager.get_tasks() if task.get('status') == 'On hold')
        if not self.show_on_hold and on_hold_count > 0:
            self.status_label.configure(text=f"已隱藏 {on_hold_count} 個 On hold 項目。")
        else:
            current_status_text = self.status_label.cget("text")
            if "儲存中" not in current_status_text and "已儲存" not in current_status_text:
                self.update_status(f"總計 {len(self.task_manager.get_tasks())} 個待辦事項。")

        # self.log_operation(f"Treeview 已重新填充並應用過濾/排序 ({len(tasks_to_display)}/{len(self.task_manager.get_tasks())} 總數顯示)。")

    def load_task_for_editing(self, event=None):
        """從 Treeview 載入選取的任務到輸入框進行編輯"""
        treeview = event.widget
        selected_items_iid = treeview.selection()
        if not selected_items_iid:
            return

        first_selected_id_str = selected_items_iid[0]
        try:
            selected_id = int(first_selected_id_str)
        except ValueError:
            self.log_operation(f"嘗試載入編輯失敗：無效的 Treeview iid {first_selected_id_str}。")
            return

        task_to_edit = self.task_manager.get_task_by_id(selected_id)

        if task_to_edit:
            if self.editing_task_id is not None and self.editing_task_id != selected_id:
                 self.cancel_edit()
            elif self.editing_task_id == selected_id:
                 return

            self.clear_input_fields()
            self.desc_entry.insert(0, task_to_edit.get('description', ''))
            self.date_display_label.configure(text=task_to_edit.get('due_date') if task_to_edit.get('due_date') else "")
            self.note_textbox.insert("1.0", task_to_edit.get('note', ''))

            self.editing_task_id = task_to_edit['id']
            self.save_button.configure(text="儲存修改")
            if self.cancel_edit_button.winfo_manager() != 'pack':
                 self.cancel_edit_button.pack(side=tk.LEFT, padx=5, pady=5, in_=self.action_button_frame)

            self.update_status(f"載入待辦事項 (ID: {selected_id}) 進行編輯。")
            self.log_operation(f"載入任務 (ID: {selected_id}) 進行編輯。")
        else:
            self.cancel_edit()
            self.log_operation(f"嘗試載入編輯失敗：找不到 ID {selected_id} 的任務。")

    def cancel_edit(self):
        """取消編輯狀態，清空輸入框，切回新增模式"""
        if self.editing_task_id is not None:
             self.editing_task_id = None
             self.clear_input_fields()
             self.save_button.configure(text="新增待辦事項")
             if self.cancel_edit_button.winfo_manager() == 'pack':
                 self.cancel_edit_button.pack_forget()

             current_tab = self.tab_notebook.tab(self.tab_notebook.select(), "text").lower()
             treeview = self.treeviews.get(current_tab)
             if treeview:
                 selected_items = treeview.selection()
                 if selected_items:
                      treeview.selection_remove(selected_items)

             self.update_status("編輯已取消，切回新增模式。")
             self.log_operation("編輯操作已取消。")

    def set_selected_task_status(self):
        """標記選取的待辦事項狀態為下拉選單的值"""
        current_tab_text = self.tab_notebook.tab(self.tab_notebook.select(), "text")
        current_tab_status = next((s for s in STATUS_OPTIONS if s.lower() == current_tab_text.strip().lower()), None)
        treeview = self.treeviews.get(current_tab_status if current_tab_status else "all") # 確保拿到正確的 treeview

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
        for item_iid in selected_items_iid:
            try:
                task_id = int(item_iid)
                task = self.task_manager.get_task_by_id(task_id)
                if task and task.get('status') != new_status:
                    # 調用 TaskManager 的更新方法
                    self.task_manager.update_task(task_id, status=new_status)
                    updated_count += 1
                    if self.editing_task_id == task_id:
                        self.after(10, lambda: self.load_task_for_editing(None))
            except ValueError:
                continue
            except Exception as e:
                messagebox.showerror("錯誤", f"變更狀態時發生錯誤: {e}")
                self.log_operation(f"變更狀態時發生錯誤: {e}")
                return

        if updated_count > 0:
            self.populate_treeview() # 更新 GUI
            self.update_status(f"{updated_count} 個待辦事項狀態已變更為 '{new_status}'。")
            self.log_operation(f"變更了 {updated_count} 個任務的狀態為 '{new_status}'。")
        else:
            self.update_status(f"選取的待辦事項狀態已是 '{new_status}'，無需變更。")
            self.log_operation(f"選取的任務狀態已是 '{new_status}'，未執行變更。")

    def delete_selected_task(self):
        """刪除選取的待辦事項"""
        current_tab_text = self.tab_notebook.tab(self.tab_notebook.select(), "text")
        current_tab_status = next((s for s in STATUS_OPTIONS if s.lower() == current_tab_text.strip().lower()), None)
        treeview = self.treeviews.get(current_tab_status if current_tab_status else "all")

        selected_items_iid = treeview.selection()
        if not selected_items_iid:
            messagebox.showwarning("選取錯誤", "請選取要刪除的待辦事項。")
            self.log_operation("嘗試刪除任務失敗：未選取項目。")
            return

        if not messagebox.askyesno("確認刪除", f"確定要刪除這 {len(selected_items_iid)} 個待辦事項嗎？"):
            self.log_operation("取消刪除任務。")
            return

        deleted_count = 0
        for item_iid in selected_items_iid:
            try:
                task_id = int(item_iid)
                # 調用 TaskManager 的刪除方法
                if self.task_manager.delete_task(task_id):
                    deleted_count += 1
                    if self.editing_task_id == task_id:
                        self.cancel_edit()
            except ValueError:
                continue
            except Exception as e:
                messagebox.showerror("錯誤", f"刪除任務時發生錯誤: {e}")
                self.log_operation(f"刪除任務時發生錯誤: {e}")
                return

        if deleted_count > 0:
            self.populate_treeview() # 更新 GUI
            self.clear_details_display()
            self.update_status(f"{deleted_count} 個待辦事項已刪除。")
            self.log_operation(f"刪除了 {deleted_count} 個任務。")
        else:
            self.update_status("沒有任務被刪除。")
            self.log_operation("沒有選取到有效的任務進行刪除。")

    def display_selected_task_details(self, event=None):
        """在 Treeview 選取項目時顯示詳細資訊"""
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

        selected_task = self.task_manager.get_task_by_id(first_selected_id)

        if selected_task:
            self.display_task_details(selected_task)
        else:
            self.clear_details_display()

    def display_task_details(self, task):
        """將指定 task 的詳細資訊顯示在詳細資訊區域"""
        self.details_creation_time_value.configure(text=utils.format_datetime(task.get('creation_time')))
        self.details_desc_value.configure(text=task.get('description', ''))
        self.details_date_value.configure(text=utils.format_date_with_weekday(task.get('due_date')))
        self.details_status_value.configure(text=task.get('status', '未知狀態') if task.get('status') in STATUS_OPTIONS else "未知狀態")

        if task.get('status') in STATUS_OPTIONS:
            self.status_combobox.set(task.get('status'))
        else:
            self.status_combobox.set("Pending") # 預設回 Pending
        
        self.details_note_textbox.configure(state="normal")
        self.details_note_textbox.delete("1.0", tk.END)
        self.details_note_textbox.insert("1.0", task.get('note', ''))

        utils.find_and_tag_urls(self.details_note_textbox)
        # 為詳細資訊備註框添加右鍵菜單
        self.details_note_textbox.bind("<Button-3>", lambda e: utils.show_context_menu(e, self.details_note_textbox, self))

        self.details_note_textbox.configure(state="disabled")

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

    def export_to_excel(self):
        """將待辦事項匯出為 Excel 檔案 (.xlsx)"""
        all_tasks = self.task_manager.get_tasks() # 獲取所有任務
        if not all_tasks:
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
            import openpyxl # 延遲導入，只有在需要時才載入

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "待辦事項"

            headers = ["ID", "內容", "到期日", "狀態", "備註/網址", "建立時間"]
            sheet.append(headers)

            for task in all_tasks:
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

        except ImportError:
            messagebox.showerror("匯出錯誤", "匯出 Excel 需要 'openpyxl' 庫。請運行 'pip install openpyxl'。")
            self.log_operation("匯出到 Excel 失敗：缺少 openpyxl 庫。")
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