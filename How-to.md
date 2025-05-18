# create requirement
pip freeze > requirement.txt

# VM Startup Script

.venv\Scripts\activate

pip install -r .\requirement.txt

# work on python

py .\working_calendar.py
