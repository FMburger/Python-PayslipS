%~d0
cd %~dp0
cd "venv\scripts"
activate & pip install -r requirements.txt & cd "..\.." & payslip_sender.py
pause
