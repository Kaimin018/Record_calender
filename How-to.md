# create requirement
pip freeze > requirement.txt

# VM Startup Script

.venv\Scripts\activate

pip install -r .\requirement.txt

# work on python

python -m record_calender.main

# test cases

pytest tests
