# record_calender/utils.py

from datetime import datetime, date
import re
import webbrowser
import tkinter as tk # 需要 Tkinter 來處理剪貼板和標籤
import tkinter.font as tkfont # 處理字體

STATUS_OPTIONS = ["Pending", "In progress", "Completed", "Cancelled", "On hold"]

def format_datetime(dt_str):
    """格式化 YYYY-MM-DD HH:MM:SS 字串為可讀格式，並包含星期幾"""
    if not dt_str:
        return "無時間"
    try:
        dt_obj = datetime.strptime(dt_str, '%Y-%m-%d %H:%M:%S')
        return dt_obj.strftime('%Y-%m-%d %H:%M:%S (%a)')
    except ValueError:
        try:
             dt_obj = datetime.strptime(dt_str.split(' ')[0], '%Y-%m-%d').date()
             return dt_obj.strftime('%Y-%m-%d (%a)')
        except ValueError:
             return dt_str

def format_date_with_weekday(date_str):
    """格式化 YYYY-MM-DD 字串為 YYYY-MM-DD (星期幾)"""
    if not date_str:
        return "無到期日"
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
        return date_obj.strftime('%Y-%m-%d (%a)')
    except ValueError:
        return date_str

def is_past_due(date_str):
    """檢查給定日期字串是否在今天之前"""
    if not date_str:
        return False
    try:
        due_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        return due_date < date.today()
    except ValueError:
        return False

def find_and_tag_urls(textbox):
    """在 Textbox 中查找 URL 並應用超連結標籤"""
    textbox.configure(state="normal")
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
        # 綁定事件時使用 lambda 確保 url 被正確捕獲
        textbox.tag_bind("url", "<Button-1>", lambda e, target_url=url: open_url(target_url))
        textbox.tag_bind("url", "<Enter>", lambda e: textbox.config(cursor="hand2"))
        textbox.tag_bind("url", "<Leave>", lambda e: textbox.config(cursor=""))

    textbox.configure(state="disabled")

def open_url(url):
    """打開指定的 URL"""
    try:
        webbrowser.open_new_tab(url)
        print(f"Opened URL: {url}") # 可以改為日誌記錄
        return True
    except Exception as e:
        print(f"Error opening URL {url}: {e}") # 可以改為日誌記錄
        return False

def copy_with_links(textbox, app_instance=None):
    """複製包含超連結的文字到剪貼板"""
    content = textbox.get("1.0", tk.END).strip()
    url_pattern = re.compile(r'https?://[^\s]+')
    urls = []

    for match in url_pattern.finditer(content):
        urls.append(match.group(0))

    clipboard_content = content
    if urls:
        clipboard_content += "\n\n連結:\n" + "\n".join(urls)
        
    if app_instance: # 確保只在 GUI 環境中執行剪貼板操作
        app_instance.clipboard_clear()
        app_instance.clipboard_append(clipboard_content)
        app_instance.update_status("已複製文字和連結到剪貼板")
    else:
        # 非 GUI 環境下的替代處理，例如打印或記錄
        print(f"Clipboard content (simulated): {clipboard_content}")

def show_context_menu(event, textbox, app_instance):
    """顯示右鍵菜單，提供複製功能"""
    textbox.configure(state="normal")
    menu = tk.Menu(app_instance, tearoff=0)
    menu.add_command(label="複製", command=lambda: copy_with_links(textbox, app_instance))
    menu.post(event.x_root, event.y_root)
    textbox.configure(state="disabled")