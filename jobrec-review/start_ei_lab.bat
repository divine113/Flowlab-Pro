@echo off
set PYTHONPATH=C:\Users\User\Downloads\Flowlab-Pro-main\Flowlab-Pro
set PYTHONUNBUFFERED=1
cd /d "C:\Users\User\Downloads\Flowlab-Pro-main\Flowlab-Pro\jobrec-review"
.venv-jobrec\Scripts\python.exe manage.py runserver 0.0.0.0:8000