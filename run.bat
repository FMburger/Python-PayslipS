%~d0
cd %~dp0
cd "venv\scripts"
activate & cd "..\.." & pip install -r requirements.txt & payslip_sender.py
pause
