# tests/test_data_manager.py

import pytest
import os
import json
from record_calender.data_manager import TaskDataManager

# 測試用檔案路徑，確保不影響真實數據
TEST_DATA_FILE = "test_todo_calendar.json"

@pytest.fixture
def temp_data_manager():
    """提供一個臨時的 TaskDataManager 實例用於測試，並在測試後清理。"""
    if os.path.exists(TEST_DATA_FILE):
        os.remove(TEST_DATA_FILE)
    manager = TaskDataManager(data_file=TEST_DATA_FILE)
    yield manager
    if os.path.exists(TEST_DATA_FILE):
        os.remove(TEST_DATA_FILE)

def test_load_tasks_empty_file(temp_data_manager):
    """測試從空檔案載入任務。"""
    tasks = temp_data_manager.load_tasks()
    assert tasks == []
    assert temp_data_manager.get_next_id() == 0

def test_save_and_load_tasks(temp_data_manager):
    """測試儲存和載入任務。"""
    test_tasks = [
        {"id": 0, "description": "Test Task 1", "due_date": "2025-06-01", "status": "Pending", "note": "", "creation_time": "2025-05-20 10:00:00", "image_path": None},
        {"id": 1, "description": "Test Task 2", "due_date": None, "status": "Completed", "note": "Some notes", "creation_time": "2025-05-21 11:00:00", "image_path": None}
    ]
    
    success = temp_data_manager.save_tasks(test_tasks)
    assert success is True

    loaded_tasks = temp_data_manager.load_tasks()
    assert len(loaded_tasks) == len(test_tasks)
    # 由於 load_tasks 會對數據進行清洗和預設值設置，這裡需要更精確的比對
    assert loaded_tasks[0]['description'] == test_tasks[0]['description']
    assert loaded_tasks[0]['id'] == test_tasks[0]['id']
    assert loaded_tasks[1]['status'] == test_tasks[1]['status']
    assert temp_data_manager.get_next_id() == 2 # 根據已有的 ID 設置正確的 next_id

def test_load_tasks_corrupted_file(temp_data_manager):
    """測試載入損壞的 JSON 檔案。"""
    with open(TEST_DATA_FILE, 'w', encoding='utf-8') as f:
        f.write("this is not json {") # 寫入一個無效的 JSON

    tasks = temp_data_manager.load_tasks()
    assert tasks == [] # 應該返回空列表
    assert temp_data_manager.get_next_id() == 0

def test_get_next_id(temp_data_manager):
    """測試獲取下一個 ID 的邏輯。"""
    assert temp_data_manager.get_next_id() == 0
    temp_data_manager.save_tasks([{"id": 0, "description": "Task A"}])
    temp_data_manager.load_tasks() # 重新載入以更新 _next_id
    assert temp_data_manager.get_next_id() == 1
    
    temp_data_manager.save_tasks([
        {"id": 5, "description": "Task X"},
        {"id": 2, "description": "Task Y"}
    ])
    temp_data_manager.load_tasks()
    assert temp_data_manager.get_next_id() == 6 # 應該是最大 ID + 1