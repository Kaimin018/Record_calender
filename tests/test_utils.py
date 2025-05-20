# tests/test_utils.py

from datetime import timedelta, date
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
def test_copy_with_links_no_url(mocker):
    mock_textbox = MockTextbox("Just some text.")
    mock_app = mocker.Mock() # 模擬 Tkinter 應用程式實例
    utils.copy_with_links(mock_textbox, mock_app)
    mock_app.clipboard_clear.assert_called_once()
    mock_app.clipboard_append.assert_called_once_with("Just some text.")
    mock_app.update_status.assert_called_once_with("已複製文字和連結到剪貼板")

def test_copy_with_links_with_url(mocker):
    mock_textbox = MockTextbox("Text with link https://example.com.")
    mock_app = mocker.Mock()
    utils.copy_with_links(mock_textbox, mock_app)
    mock_app.clipboard_clear.assert_called_once()
    expected_clipboard_content = "Text with link https://example.com.\n\n連結:\nhttps://example.com"
    mock_app.clipboard_append.assert_called_once_with(expected_clipboard_content)
    mock_app.update_status.assert_called_once_with("已複製文字和連結到剪貼板")