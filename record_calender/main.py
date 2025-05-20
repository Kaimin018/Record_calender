# record_calender/main.py

import os
from record_calender.data_manager import TaskDataManager
from record_calender.task_manager import TaskManager
from record_calender.gui import TodoApp

def run_app():
    # 確保數據檔案路徑正確，相對於項目根目錄
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # data_file_path = os.path.join(script_dir, '..', 'todo_calendar.json')
    # 為了與 data_manager.py 的 DEFAULT_DATA_FILE 保持一致，讓它自己決定
    # data_manager 已經處理了相對路徑，不需要這裡再處理
    data_manager = TaskDataManager()
    task_manager = TaskManager(data_manager)
    
    app = TodoApp(task_manager)
    app.mainloop()

if __name__ == "__main__":
    run_app()