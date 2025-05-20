# tests/test_task_manager.py

import pytest
from unittest.mock import Mock, patch
from datetime import datetime
from record_calender.task_manager import TaskManager, STATUS_OPTIONS
from record_calender.data_manager import TaskDataManager

@pytest.fixture
def mock_data_manager():
    """模擬 TaskDataManager 實例，用於隔離測試 TaskManager。"""
    mock_dm = Mock(spec=TaskDataManager)
    mock_dm.load_tasks.return_value = []
    mock_dm.save_tasks.return_value = True
    mock_dm.get_next_id.side_effect = iter(range(100)) # 模擬 ID 遞增
    return mock_dm

@pytest.fixture
def task_manager_instance(mock_data_manager):
    """提供 TaskManager 實例。"""
    return TaskManager(mock_data_manager)

def test_add_task_success(task_manager_instance, mock_data_manager):
    """測試成功新增任務。"""
    with patch('record_calender.task_manager.datetime') as mock_dt:
        mock_dt.now.return_value = datetime(2025, 5, 20, 12, 30, 0)
        mock_dt.strptime = datetime.strptime # Preserve original strptime

        task = task_manager_instance.add_task("Buy groceries", "2025-05-25", "Milk and bread")
        assert task['description'] == "Buy groceries"
        assert task['due_date'] == "2025-05-25"
        assert task['note'] == "Milk and bread"
        assert task['status'] == "Pending"
        assert task['id'] == 0 # 第一次調用 get_next_id
        assert task['creation_time'] == "2025-05-20 12:30:00"
        
        # 驗證 TaskDataManager 的 save_tasks 被調用
        mock_data_manager.save_tasks.assert_called_once()
        assert len(task_manager_instance.get_tasks()) == 1

def test_add_task_empty_description(task_manager_instance):
    """測試新增空內容任務時拋出錯誤。"""
    with pytest.raises(ValueError, match="Task description cannot be empty."):
        task_manager_instance.add_task("", "2025-05-25")

def test_add_task_invalid_date_format(task_manager_instance):
    """測試新增任務時日期格式無效。"""
    with pytest.raises(ValueError, match="Invalid due date format. Please use YYYY-MM-DD."):
        task_manager_instance.add_task("Read book", "2025/05/25")

def test_update_task_success(task_manager_instance, mock_data_manager):
    """測試成功更新任務。"""
    task = task_manager_instance.add_task("Original Task")
    mock_data_manager.save_tasks.reset_mock() # 重置計數

    updated_task = task_manager_instance.update_task(task['id'], description="Updated Task", status="Completed")
    assert updated_task['description'] == "Updated Task"
    assert updated_task['status'] == "Completed"
    mock_data_manager.save_tasks.assert_called_once()

def test_update_task_not_found(task_manager_instance):
    """測試更新不存在的任務。"""
    result = task_manager_instance.update_task(999, description="Non-existent")
    assert result is None

def test_delete_task_success(task_manager_instance, mock_data_manager):
    """測試成功刪除任務。"""
    task = task_manager_instance.add_task("Task to delete")
    mock_data_manager.save_tasks.reset_mock()

    deleted = task_manager_instance.delete_task(task['id'])
    assert deleted is True
    assert len(task_manager_instance.get_tasks()) == 0
    mock_data_manager.save_tasks.assert_called_once()

def test_delete_task_not_found(task_manager_instance):
    """測試刪除不存在的任務。"""
    deleted = task_manager_instance.delete_task(999)
    assert deleted is False

def test_get_tasks_by_status(task_manager_instance, mock_data_manager):
    """測試按狀態篩選任務。"""
    task1 = task_manager_instance.add_task("Pending task 1")
    task_manager_instance.update_task(task1['id'], "In progress task 1", status="In progress")
    task2 = task_manager_instance.add_task("Pending task 2")

    pending_tasks = task_manager_instance.get_tasks_by_status("Pending")
    assert len(pending_tasks) == 1
    assert pending_tasks[0]['description'] == "Pending task 2"

    in_progress_tasks = task_manager_instance.get_tasks_by_status("In progress")
    assert len(in_progress_tasks) == 1
    assert in_progress_tasks[0]['description'] == "In progress task 1"

    completed_tasks = task_manager_instance.get_tasks_by_status("Completed")
    assert len(completed_tasks) == 0

def test_get_all_tasks_sorted(task_manager_instance, mock_data_manager):
    """測試所有任務的排序功能。"""
    task_manager_instance.add_task("Task B", "2025-05-22")
    task_manager_instance.add_task("Task A", "2025-05-21")
    task_manager_instance.add_task("Task C", "2025-05-23")
    
    # 按 due_date 升序
    sorted_tasks = task_manager_instance.get_all_tasks_sorted(sort_column="due_date", sort_direction="ascending")
    assert sorted_tasks[0]['description'] == "Task A"
    assert sorted_tasks[1]['description'] == "Task B"
    assert sorted_tasks[2]['description'] == "Task C"

    # 按 description 降序
    sorted_tasks = task_manager_instance.get_all_tasks_sorted(sort_column="description", sort_direction="descending")
    assert sorted_tasks[0]['description'] == "Task C"
    assert sorted_tasks[1]['description'] == "Task B"
    assert sorted_tasks[2]['description'] == "Task A"