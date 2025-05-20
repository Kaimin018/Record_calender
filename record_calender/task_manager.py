# record_calender/task_manager.py

from datetime import datetime
from record_calender.data_manager import TaskDataManager # 導入資料管理員

# 定義所有可能的狀態，與應用程式同步
STATUS_OPTIONS = ["Pending", "In progress", "Completed", "Cancelled", "On hold"]

class TaskManager:
    def __init__(self, data_manager: TaskDataManager):
        self.data_manager = data_manager
        self._tasks = self.data_manager.load_tasks()

    def get_tasks(self):
        """獲取所有任務的副本，避免外部直接修改內部列表。"""
        return list(self._tasks)

    def add_task(self, description, due_date=None, note=None):
        """
        新增一個待辦事項。
        :param description: 任務內容 (str)
        :param due_date: 到期日期 (str, YYYY-MM-DD), 可選
        :param note: 備註 (str), 可選
        :return: 新增的任務字典，如果失敗則為 None
        """
        if not description:
            # 可以拋出錯誤或返回 None，取決於錯誤處理策略
            raise ValueError("Task description cannot be empty.")

        # 驗證日期格式（如果提供）
        if due_date:
            try:
                datetime.strptime(due_date, '%Y-%m-%d')
            except ValueError:
                raise ValueError("Invalid due date format. Please use YYYY-MM-DD.")

        task = {
            'id': self.data_manager.get_next_id(),
            'description': description,
            'due_date': due_date,
            'status': 'Pending',
            'note': note if note is not None else '',
            'creation_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'image_path': None
        }
        self._tasks.append(task)
        self.data_manager.save_tasks(self._tasks) # 立即儲存
        return task

    def update_task(self, task_id, description=None, due_date=None, status=None, note=None):
        """
        更新一個現有的待辦事項。
        :param task_id: 任務的 ID
        :param description: 新的任務內容 (str), 可選
        :param due_date: 新的到期日期 (str, YYYY-MM-DD), 可選
        :param status: 新的狀態 (str), 必須是 STATUS_OPTIONS 中的一個, 可選
        :param note: 新的備註 (str), 可選
        :return: 更新後的任務字典，如果找不到或更新失敗則為 None
        """
        task_to_edit = next((task for task in self._tasks if task['id'] == task_id), None)
        if not task_to_edit:
            return None

        updated = False
        if description is not None:
            if not description.strip():
                raise ValueError("Task description cannot be empty.")
            task_to_edit['description'] = description
            updated = True
        
        if due_date is not None:
            if due_date: # Only validate if not None/empty
                try:
                    datetime.strptime(due_date, '%Y-%m-%d')
                except ValueError:
                    raise ValueError("Invalid due date format. Please use YYYY-MM-DD.")
            task_to_edit['due_date'] = due_date
            updated = True

        if status is not None:
            if status not in STATUS_OPTIONS:
                raise ValueError(f"Invalid status: {status}. Must be one of {STATUS_OPTIONS}")
            task_to_edit['status'] = status
            updated = True
            
        if note is not None:
            task_to_edit['note'] = note
            updated = True

        if updated:
            self.data_manager.save_tasks(self._tasks) # 立即儲存
            return task_to_edit
        return None # 沒有任何東西被更新

    def delete_task(self, task_id):
        """
        刪除一個待辦事項。
        :param task_id: 任務的 ID
        :return: True 如果刪除成功，False 如果任務不存在
        """
        initial_len = len(self._tasks)
        self._tasks = [task for task in self._tasks if task['id'] != task_id]
        if len(self._tasks) < initial_len:
            self.data_manager.save_tasks(self._tasks) # 立即儲存
            return True
        return False
        
    def get_task_by_id(self, task_id):
        """
        根據 ID 獲取一個任務。
        :param task_id: 任務的 ID
        :return: 任務字典，如果找不到則為 None
        """
        return next((task for task in self._tasks if task['id'] == task_id), None)

    def get_tasks_by_status(self, status):
        """
        獲取指定狀態的所有任務。
        :param status: 任務狀態 (str)
        :return: 任務列表
        """
        if status not in STATUS_OPTIONS:
            raise ValueError(f"Invalid status: {status}. Must be one of {STATUS_OPTIONS}")
        return [task for task in self._tasks if task.get('status') == status]

    def get_all_tasks_sorted(self, sort_column=None, sort_direction='ascending'):
        """
        獲取所有任務，並可選擇進行排序。
        :param sort_column: 排序的欄位名稱
        :param sort_direction: 排序方向 ('ascending' 或 'descending')
        :return: 排序後的任務列表
        """
        tasks_to_sort = list(self._tasks) # 複製列表以避免修改原始數據
        
        if sort_column:
            def sort_key(task):
                value = task.get(sort_column)
                if sort_column in ['due_date', 'creation_time']:
                    try:
                        # 對日期時間字串進行轉換以便正確排序
                        if sort_column == 'due_date':
                            return datetime.strptime(str(value).split(' ')[0], '%Y-%m-%d').date() if value else datetime.max.date()
                        if sort_column == 'creation_time':
                            return datetime.strptime(str(value), '%Y-%m-%d %H:%M:%S') if value else datetime.max
                    except (ValueError, IndexError):
                        # 對於無效日期，將其視為最大值，使其排在末尾
                        return datetime.max.date() if sort_column == 'due_date' else datetime.max
                elif sort_column == 'status':
                    try:
                        # 根據預定義的狀態順序進行排序
                        return STATUS_OPTIONS.index(value)
                    except ValueError:
                        return len(STATUS_OPTIONS) # 未知狀態排在最後
                # 對於其他字串類型，進行小寫轉換以實現不區分大小寫的排序
                return str(value).lower() if value is not None else ''

            tasks_to_sort.sort(key=sort_key, reverse=(sort_direction == 'descending'))
        else:
            # 預設按建立時間降序排序 (最新在前)
            tasks_to_sort.sort(key=lambda x: datetime.strptime(x.get('creation_time', '1900-01-01 00:00:00'), '%Y-%m-%d %H:%M:%S') if x.get('creation_time') else datetime.min, reverse=True)
            
        return tasks_to_sort