# record_calender/data_manager.py

import json
import os

# 確保 DATA_FILE 能夠從外部設定，或者使用一個安全的預設值
# 在實際應用中，可以通過配置或在 __init__ 函數中傳入路徑
DEFAULT_DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'todo_calendar.json')

class TaskDataManager:
    def __init__(self, data_file=None):
        self.data_file = data_file if data_file else DEFAULT_DATA_FILE
        self._next_id = 0 # 內部追蹤下一個可用的 ID

    def _load_raw_tasks(self):
        """從檔案載入原始 JSON 數據。"""
        if not os.path.exists(self.data_file):
            return []
        try:
            with open(self.data_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            print(f"Warning: Could not decode JSON from {self.data_file}. Starting with empty tasks.")
            return []
        except Exception as e:
            print(f"Error loading tasks from {self.data_file}: {e}")
            return []

    def load_tasks(self):
        """從檔案載入待辦事項並進行必要的數據清洗和 ID 初始化。"""
        raw_tasks = self._load_raw_tasks()
        tasks = []
        current_max_id = -1
        
        # 定義所有可能的狀態，與應用程式同步
        STATUS_OPTIONS = ["Pending", "In progress", "Completed", "Cancelled", "On hold"]

        for task in raw_tasks:
            # 確保 task 是字典且有 description
            if not isinstance(task, dict) or 'description' not in task:
                continue # 跳過無效的任務格式

            # 確保 ID 存在且是數字
            if 'id' not in task or not isinstance(task['id'], (int, float)):
                # 如果沒有 ID 或 ID 無效，暫時分配一個
                task['id'] = current_max_id + 1 # 會在最後更新 _next_id

            task['id'] = int(task['id']) # 確保 ID 是整數
            current_max_id = max(current_max_id, task['id'])

            # 設置預設值
            task.setdefault('due_date', None)
            task.setdefault('creation_time', None)
            task.setdefault('status', 'Pending')
            task.setdefault('note', '')
            task.setdefault('image_path', None)

            # 規範化狀態
            matched_status = next((s for s in STATUS_OPTIONS if s.lower() == task['status'].lower()), None)
            task['status'] = matched_status if matched_status else 'Pending'

            tasks.append(task)
        
        self._next_id = current_max_id + 1 if tasks else 0
        return tasks

    def save_tasks(self, tasks):
        """將待辦事項儲存到檔案。"""
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(tasks, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving tasks to {self.data_file}: {e}")
            return False, e

    def get_next_id(self):
        """取得下一個可用的唯一 ID。"""
        task_id = self._next_id
        self._next_id += 1
        return task_id

    def set_next_id(self, new_id):
        """為測試目的設定下一個 ID。"""
        if new_id >= 0:
            self._next_id = new_id
        else:
            raise ValueError("Next ID cannot be negative.")

# 可以在此處實例化一個 DataManager，或者在 main.py 中實例化並傳遞給其他模組
# data_manager = TaskDataManager()