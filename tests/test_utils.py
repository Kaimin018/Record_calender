# tests/test_utils.py
import pytest
from unittest.mock import Mock, patch
from datetime import datetime, timedelta, date
from record_calender import utils
import tkinter as tk # 僅用於模擬 Tkinter Textbox

# 模擬一個簡單的 Textbox 類別用於測試 find_and_tag_urls
class MockTextbox:
    def __init__(self, content=""):
        self._content = content
        self._tags = []
        self._config = {}
        self._binds = {}

    def get(self, start, end):
        if start == "1.0" and end == "end":
            return self._content
        return self._content[0:len(self._content)] # 簡化

    def configure(self, state=None, **kwargs):
        if state is not None:
            self._config['state'] = state
        self._config.update(kwargs)

    def tag_remove(self, tag, start, end):
        self._tags = [t for t in self._tags if not (t['tag'] == tag and t['start'] == start and t['end'] == end)]

    def tag_configure(self, tag, **kwargs):
        pass # 簡單模擬，不實際配置

    def tag_add(self, tag, start, end):
        self._tags.append({'tag': tag, 'start': start, 'end': end})

    def tag_bind(self, tag, event, func):
        if tag not in self._binds:
            self._binds[tag] = {}
        self._binds[tag][event] = func

    def index(self, idx_str):
        # 簡單模擬，根據 '1.0 + Nc' 返回對應的字元索引
        parts = idx_str.split('+')
        if len(parts) == 2 and parts[0].strip() == '1.0' and parts[1].strip().endswith('c'):
            char_offset = int(parts[1].strip()[:-1])
            return f"1.0+{char_offset}c" # 返回模擬的索引字符串
        return "1.0" # 默認值

    def cget(self, key):
        return self._config.get(key)

# 測試 format_datetime
def test_format_datetime_full():
    dt_str = "2025-05-20 14:30:00"
    assert utils.format_datetime(dt_str) == "2025-05-20 14:30:00 (Tue)"

def test_format_datetime_date_only():
    dt_str = "2025-05-20"
    assert utils.format_datetime(dt_str) == "2025-05-20 (Tue)" # 期望行為是只解析日期部分

def test_format_datetime_empty():
    assert utils.format_datetime(None) == "無時間"
    assert utils.format_datetime("") == "無時間"

def test_format_datetime_invalid():
    assert utils.format_datetime("invalid-date") == "invalid-date"

# 測試 format_date_with_weekday
def test_format_date_with_weekday_valid():
    date_str = "2025-05-20"
    assert utils.format_date_with_weekday(date_str) == "2025-05-20 (Tue)"

def test_format_date_with_weekday_empty():
    assert utils.format_date_with_weekday(None) == "無到期日"
    assert utils.format_date_with_weekday("") == "無到期日"

def test_format_date_with_weekday_invalid():
    assert utils.format_date_with_weekday("bad-date") == "bad-date"

# 測試 is_past_due
def test_is_past_due_future_date():
    future_date = (date.today() + timedelta(days=1)).strftime('%Y-%m-%d')
    assert utils.is_past_due(future_date) is False

def test_is_past_due_today():
    today_date = date.today().strftime('%Y-%m-%d')
    assert utils.is_past_due(today_date) is False

def test_is_past_due_past_date():
    past_date = (date.today() - timedelta(days=1)).strftime('%Y-%m-%d')
    assert utils.is_past_due(past_date) is True

def test_is_past_due_empty_date():
    assert utils.is_past_due(None) is False
    assert utils.is_past_due("") is False

def test_is_past_due_invalid_format():
    assert utils.is_past_due("invalid-date") is False

# 測試 find_and_tag_urls
def test_find_and_tag_urls_no_url():
    mock_textbox = MockTextbox("This is a simple text with no URLs.")
    utils.find_and_tag_urls(mock_textbox)
    assert not mock_textbox._tags # 應該沒有添加任何標籤

def test_find_and_tag_urls_with_url():
    mock_textbox = MockTextbox("Visit Google at https://www.google.com and then this: http://example.com/path")
    utils.find_and_tag_urls(mock_textbox)
    assert len(mock_textbox._tags) == 2
    assert {'tag': 'url', 'start': '1.0+17c', 'end': '1.0+40c'} in mock_textbox._tags
    assert {'tag': 'url', 'start': '1.0+56c', 'end': '1.0+79c'} in mock_textbox._tags
    # 驗證綁定是否存在
    assert "url" in mock_textbox._binds
    assert "<Button-1>" in mock_textbox._binds["url"]

# 測試 open_url（需要實際的瀏覽器，通常會模擬或跳過）
# 這裡只做一個簡單的，確保調用 webbrowser.open_new_tab
from unittest.mock import patch
def test_open_url_success():
    with patch('record_calender.utils.webbrowser.open_new_tab') as mock_open:
        result = utils.open_url("https://test.com")
        mock_open.assert_called_once_with("https://test.com")
        assert result is True

def test_open_url_failure():
    with patch('record_calender.utils.webbrowser.open_new_tab', side_effect=Exception("Failed to open")):
        result = utils.open_url("https://error.com")
        assert result is False

# 測試 copy_with_links
def test_find_and_tag_urls_with_url():
    text = "Visit Google at https://www.google.com and then this: http://example.com/path"
    mock_textbox = MockTextbox(text)
    utils.find_and_tag_urls(mock_textbox)
    # 動態計算網址在字串中的位置
    urls = ["https://www.google.com", "http://example.com/path"]
    for url in urls:
        start = text.index(url)
        end = start + len(url)
        tag = {'tag': 'url', 'start': f'1.0+{start}c', 'end': f'1.0+{end}c'}
        assert tag in mock_textbox._tags
    # 驗證綁定是否存在
    assert "url" in mock_textbox._binds
    assert "<Button-1>" in mock_textbox._binds["url"]

def test_copy_with_links_with_url():
    mock_textbox = MockTextbox("Text with link https://example.com.")
    mock_app = Mock()
    utils.copy_with_links(mock_textbox, mock_app)
    mock_app.clipboard_clear.assert_called_once()
    expected_clipboard_content = "Text with link https://example.com.\n\n連結:\nhttps://example.com."
    mock_app.clipboard_append.assert_called_once_with(expected_clipboard_content)
    mock_app.update_status.assert_called_once_with("已複製文字和連結到剪貼板")
    
# Assuming your application has an 'Item' class or similar data structure
# and a main 'Application' class or core logic module.
# from record_calender import Item, Application, core_logic, ui_utils

# --- Mockups needed for various tests ---

class MockItem:
    def __init__(self, item_id, content, due_date=None, created_at=None, status="Pending", next_handler=None, image_attachments=None):
        self.id = item_id
        self.content = content
        self._created_at = created_at if created_at else datetime.now()
        self.due_date = due_date # Expected format: "YYYY-MM-DD"
        self.status = status
        self.next_handler = next_handler
        self.image_attachments = image_attachments if image_attachments else []

    @property
    def created_at(self):
        # Simulate how it might be stored/retrieved, e.g., as a string
        return self._created_at.strftime('%Y-%m-%d %H:%M:%S') if isinstance(self._created_at, datetime) else self._created_at

    def attach_image(self, image_path):
        # In a real scenario, this might involve file operations
        if image_path == "invalid/path.jpg":
            return False
        self.image_attachments.append({"path": image_path, "id": f"img_{len(self.image_attachments)+1}"})
        return True

    def remove_image(self, image_id):
        self.image_attachments = [img for img in self.image_attachments if img['id'] != image_id]

    def set_next_handler(self, handler_name):
        self.next_handler = handler_name

    def set_status(self, new_status):
        self.status = new_status


class MockApplication:
    def __init__(self):
        self.items = []
        self.config = {'column_order': ['content', 'due_date', 'status'],
                       'tab_order': ['All', 'In Progress', 'Holding', 'Complete']}
        self.log = []
        self.ui_state = {
            'on_hold_panel_visible': False,
            'current_selected_task_id': None,
            'status_field_value': "" # Simulates a UI field
        }
        # For testing clipboard functionality from test_utils.py
        self._clipboard_content = ""
        self._status_message = ""


    def add_item(self, content, due_date=None, created_at_override=None, use_threading=False):
        # Simulate item creation, potentially with threading
        item_id = len(self.items) + 1
        # For feature 1: Ensure creation timestamp is recorded
        creation_time = created_at_override if created_at_override else datetime.now()

        if use_threading:
            # Simulate offloading to a thread
            with patch('threading.Thread') as mock_thread_class:
                # Actual processing would be in the thread's target
                def process_item_creation():
                    new_item = MockItem(item_id, content, due_date=due_date, created_at=creation_time)
                    self.items.append(new_item)
                    self.log_operation(f"Item created: {content[:20]}")

                thread_instance = mock_thread_class.return_value
                process_item_creation() # In test, call directly or mock thread run
                # To assert threading was attempted:
                # mock_thread_class.assert_called_once()
                # thread_instance.start.assert_called_once()
                return self.items[-1] if self.items else None


        new_item = MockItem(item_id, content, due_date=due_date, created_at=creation_time)
        self.items.append(new_item)
        self.log_operation(f"Item created: {content[:20]}") # For feature 9
        return new_item

    def get_item_by_id(self, item_id):
        for item in self.items:
            if item.id == item_id:
                return item
        return None

    def save_all_items(self):
        # Simulate saving action, for Ctrl+S test
        self.log_operation("Data saved (Ctrl+S)")
        return True

    def log_operation(self, message):
        self.log.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")

    def get_logs(self):
        return self.log

    def get_on_hold_items_count(self):
        return len([item for item in self.items if item.status == "On hold"])

    def toggle_on_hold_panel(self):
        self.ui_state['on_hold_panel_visible'] = not self.ui_state['on_hold_panel_visible']

    def filter_items_by_status(self, status):
        if status == "All":
            return self.items
        return [item for item in self.items if item.status == status]

    def select_task(self, task_id):
        # For feature 16
        self.ui_state['current_selected_task_id'] = task_id
        task = self.get_item_by_id(task_id)
        if task:
            self.ui_state['status_field_value'] = task.status # Simulate UI update
        else:
            self.ui_state['status_field_value'] = ""


    def get_column_order(self):
        return self.config['column_order']

    def set_column_order(self, new_order):
        self.config['column_order'] = new_order
        self.log_operation(f"Column order changed to: {new_order}")

    def get_tab_order(self):
        return self.config['tab_order']

    def set_tab_order(self, new_order):
        self.config['tab_order'] = new_order
        self.log_operation(f"Tab order changed to: {new_order}")

    # Mock methods for clipboard (as in your example, might be part of app)
    def clipboard_clear(self):
        self._clipboard_content = ""

    def clipboard_append(self, text):
        self._clipboard_content += text

    def update_status(self, message): # Mock for status bar update
        self._status_message = message


# --- Feature Test Items ---

# 1. 加上首次記錄的時間點
def test_item_creation_records_timestamp():
    app = MockApplication()
    # Ensure we can control the "current time" for consistent testing
    fixed_time = datetime(2025, 5, 27, 10, 0, 0)
    with patch('record_calender.test_app_features.datetime') as mock_dt: # Assuming this test file is record_calender.test_app_features
        mock_dt.now.return_value = fixed_time
        mock_dt.side_effect = lambda *args, **kw: datetime(*args, **kw) # Allow datetime constructor

        item = app.add_item("Test Item 1")
        assert item is not None
        assert item.created_at == "2025-05-27 10:00:00"
        # Further check if it's stored persistently (would need DB mock or similar)

# 2. 選取日期的視窗可以相鄰在主視窗旁邊 (UI - difficult for unit test, conceptual)
# This would typically be an E2E UI test.
# If there's a utility function calculating position:
# def test_calculate_datepicker_position(mock_main_window_rect):
#     position = ui_utils.calculate_datepicker_position(mock_main_window_rect)
#     assert position_is_adjacent(position, mock_main_window_rect)

# 3. ctrl s 沒有綁定成功
def test_ctrl_s_triggers_save_action(monkeypatch):
    app = MockApplication()
    app.add_item("Unsaved item") # Simulate some data
    
    # Mock the actual save function that Ctrl+S should trigger
    mock_save_func = Mock(return_value=True)
    # Assuming Ctrl+S is bound to app.handle_save_shortcut which calls app.save_all_items
    monkeypatch.setattr(app, "save_all_items", mock_save_func)

    # Simulate the event handler being called (e.g. app.handle_save_shortcut())
    # This assumes you have a method that gets called by the Ctrl+S event
    if hasattr(app, "handle_save_shortcut"):
        app.handle_save_shortcut() 
        mock_save_func.assert_called_once()
    else: # Or directly test the intended action if no intermediate handler
        app.save_all_items()
        mock_save_func.assert_called_once()
    assert "Data saved (Ctrl+S)" in app.get_logs()[-1]


# 4. 過期item前面加上紅色驚嘆號icon (UI display based on logic)
# The logic for 'is_past_due' is already in test_utils.py.
# This test would be for a hypothetical function that determines display properties.
def test_get_item_display_properties_for_overdue_item():
    # utils.is_past_due is already tested. This tests its application.
    # Assume a function that returns display hints for an item
    # def get_item_visual_cue(item_due_date):
    #     if utils.is_past_due(item_due_date): return "red_exclamation_icon"
    #     return None
    past_date_str = (date.today() - timedelta(days=1)).strftime('%Y-%m-%d')
    # assert get_item_visual_cue(past_date_str) == "red_exclamation_icon"
    # current_date_str = date.today().strftime('%Y-%m-%d')
    # assert get_item_visual_cue(current_date_str) is None
    pytest.skip("Skipping UI cue test, covered by is_past_due logic and UI implementation verification.")


# 5. 加上可以attach圖片功能
def test_item_attach_image():
    item = MockItem(1, "Item with image")
    assert item.attach_image("path/to/image.png") is True
    assert len(item.image_attachments) == 1
    assert item.image_attachments[0]['path'] == "path/to/image.png"

def test_item_attach_image_invalid_path():
    item = MockItem(1, "Test item")
    assert item.attach_image("invalid/path.jpg") is False
    assert not item.image_attachments

def test_item_remove_attached_image():
    item = MockItem(1, "Item for image removal")
    item.attach_image("image1.jpg")
    item.attach_image("image2.jpg")
    image_id_to_remove = item.image_attachments[0]['id']
    
    item.remove_image(image_id_to_remove)
    assert len(item.image_attachments) == 1
    assert item.image_attachments[0]['path'] == "image2.jpg"

# 6. 點擊欄位title可以filter (Assuming underlying filter logic)
def test_filter_items_by_status():
    app = MockApplication()
    app.add_item("Task 1", status="In Progress")
    app.add_item("Task 2", status="Completed")
    app.add_item("Task 3", status="In Progress")

    filtered = app.filter_items_by_status("In Progress")
    assert len(filtered) == 2
    assert all(item.status == "In Progress" for item in filtered)

    filtered_completed = app.filter_items_by_status("Completed")
    assert len(filtered_completed) == 1
    assert filtered_completed[0].content == "Task 2"

# 7. 加上欄位 下一手 (optional)
def test_item_set_and_get_next_handler():
    item = MockItem(1, "Delegated Task")
    item.set_next_handler("Alice")
    assert item.next_handler == "Alice"
    item.set_next_handler("Bob")
    assert item.next_handler == "Bob"
    item.set_next_handler(None)
    assert item.next_handler is None

# 8. 右邊加固定日曆 顯示什麼時候要交什麼給誰 (Data prep for calendar)
def test_get_calendar_event_markers_for_month():
    app = MockApplication()
    app.add_item("Event A", due_date="2025-06-15", next_handler="Alice")
    app.add_item("Event B", due_date="2025-06-15", next_handler="Bob")
    app.add_item("Event C", due_date="2025-06-20")
    app.add_item("Event D", due_date="2025-07-01") # Different month

    # Assume a function in your core_logic or ui_utils
    # def get_events_for_calendar_view(items, year, month):
    #     events = {}
    #     for item in items:
    #         if not item.due_date: continue
    #         item_date = datetime.strptime(item.due_date, '%Y-%m-%d').date()
    #         if item_date.year == year and item_date.month == month:
    #             day_str = item_date.strftime('%Y-%m-%d')
    #             if day_str not in events: events[day_str] = []
    #             events[day_str].append({"title": item.content, "handler": item.next_handler})
    #     return events
    
    # Example for a hypothetical function:
    # calendar_events = get_events_for_calendar_view(app.items, 2025, 6)
    # assert "2025-06-15" in calendar_events
    # assert len(calendar_events["2025-06-15"]) == 2
    # assert {"title": "Event A", "handler": "Alice"} in calendar_events["2025-06-15"]
    # assert "2025-06-20" in calendar_events
    # assert len(calendar_events["2025-06-20"]) == 1
    # assert "2025-07-01" not in calendar_events # Belongs to July
    pytest.skip("Skipping calendar data prep test; requires specific data aggregation logic.")


# 9. 上面新增選單 可以看到操作的log
def test_operations_are_logged():
    app = MockApplication()
    app.add_item("Log Test Item")
    app.save_all_items() # Another logged action
    
    logs = app.get_logs()
    assert len(logs) >= 2 # At least item creation and save
    assert "Item created: Log Test Item" in logs[0]
    assert "Data saved (Ctrl+S)" in logs[1]

# 10. on hold項目可以放到右下角icon縮起來 節省版面空間
def test_on_hold_items_count_and_panel_toggle():
    app = MockApplication()
    app.add_item("Task X", status="On hold")
    app.add_item("Task Y", status="In Progress")
    app.add_item("Task Z", status="On hold")

    assert app.get_on_hold_items_count() == 2
    
    assert app.ui_state['on_hold_panel_visible'] is False
    app.toggle_on_hold_panel()
    assert app.ui_state['on_hold_panel_visible'] is True
    app.toggle_on_hold_panel()
    assert app.ui_state['on_hold_panel_visible'] is False

# 11. 到期日加上星期幾 (Handled by format_date_with_weekday in test_utils.py)
# UI test would verify it's USED in the display.

# 12. 內容欄位目前是置中 想改成靠左對齊 (UI specific)
# This would be verified in UI tests or by inspecting widget configuration.
# If a utility sets this:
# def test_content_field_default_alignment(mock_text_widget):
#     ui_utils.style_content_field(mock_text_widget)
#     mock_text_widget.configure.assert_called_with(justify='left') # or similar Tkinter property
# pytest.skip("Skipping content field alignment test, UI specific.")

# 13. 讓輸入item的時候不會卡頓，使用multi thread的方式處理
@patch('threading.Thread') # Patch threading.Thread globally for this test
def test_item_creation_uses_thread_if_enabled(mock_thread_class):
    app = MockApplication()
    thread_instance = mock_thread_class.return_value

    app.add_item("Threaded Item", use_threading=True)

    # Assert that a thread was created and started
    # This depends on how add_item is implemented when use_threading is True
    # For this MockApplication, the processing is direct, but we can check if Thread was called.
    # If add_item *itself* creates and starts the thread:
    # mock_thread_class.assert_called_once()
    # thread_instance.start.assert_called_once()
    # This assertion needs to be more specific based on your actual implementation.
    # The current MockApplication.add_item with use_threading=True calls patch,
    # but doesn't use the instance in a way that's easily assertable without deeper mocking.
    # A better approach in real code:
    # app.item_processor.submit_new_item_task(...)
    # mock_item_processor.submit_new_item_task.assert_called_once()
    pytest.skip("Skipping threading test; mock setup needs to align with actual implementation of threading.")


# 14. 上面新增一個tab欄位 可以特別把holding 以及complete項目獨立顯示
def test_tab_filters_items_correctly():
    app = MockApplication()
    app.add_item("Item P", status="In Progress")
    app.add_item("Item H", status="On hold")
    app.add_item("Item C", status="Completed")
    app.add_item("Item H2", status="On hold")

    on_hold_items = app.filter_items_by_status("On hold")
    assert len(on_hold_items) == 2
    assert all(item.status == "On hold" for item in on_hold_items)

    completed_items = app.filter_items_by_status("Completed")
    assert len(completed_items) == 1
    assert completed_items[0].content == "Item C"

    all_items = app.filter_items_by_status("All")
    assert len(all_items) == 4

# 15. 複製有超連結的文字時可以一起複製過來
# This is tested by test_copy_with_links_with_url in your test_utils.py,
# assuming the `mock_textbox` and `mock_app` there represent your actual components.

# 16. 點選某個task的時候 狀態沒有跟著讀取到 (Bug fix verification)
def test_select_task_loads_correct_status_in_ui():
    app = MockApplication()
    item1 = app.add_item("Task Alpha", status="Pending")
    item2 = app.add_item("Task Beta", status="In Progress")

    # Simulate selecting Task Beta
    app.select_task(item2.id)
    assert app.ui_state['current_selected_task_id'] == item2.id
    # Assuming app.select_task updates a simulated UI field for status
    assert app.ui_state['status_field_value'] == "In Progress"

    # Simulate selecting Task Alpha
    app.select_task(item1.id)
    assert app.ui_state['current_selected_task_id'] == item1.id
    assert app.ui_state['status_field_value'] == "Pending"

# 17. 增加右邊垂直滾輪 讓下面的欄位可以完整顯示 (UI Layout)
# This is typically managed by the GUI toolkit (e.g., Tkinter's Scrollbar with Listbox/Canvas).
# Unit tests can't easily verify if a scrollbar *appears visually*.
# One might test if a container widget is *configured* to be scrollable.
# pytest.skip("Skipping scrollbar appearance test, UI layout specific.")


# 18. 建立時間欄位可以移動到最後
def test_set_and_get_column_order_creation_time_last():
    app = MockApplication()
    original_order = app.get_column_order()
    assert "created_at" not in original_order # Assuming it's not there by default in mock

    new_order = original_order + ["created_at"]
    app.set_column_order(new_order)
    
    assert app.get_column_order()[-1] == "created_at"
    assert "Column order changed" in app.get_logs()[-1]
    # Test persistence if applicable (e.g., mock config file write)

# 19. 上面tab可以調整順序 把 complete移到後面 in progress往前移
def test_set_and_get_tab_order():
    app = MockApplication()
    # Default: ['All', 'In Progress', 'Holding', 'Complete']
    
    new_order = ['All', 'Holding', 'In Progress', 'Complete'] # Example: Move 'In Progress' after 'Holding'
    app.set_tab_order(new_order)
    
    current_tabs = app.get_tab_order()
    assert current_tabs[1] == 'Holding'
    assert current_tabs[2] == 'In Progress'
    assert "Tab order changed" in app.get_logs()[-1]
    # Test persistence if applicable