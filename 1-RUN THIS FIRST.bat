@echo off
FOR /f %%p in ('where python') do SET PYTHONPATH=%%p
ECHO %PYTHONPATH%
curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
python get-pip.py
pip install -r requirements.txt
pause
