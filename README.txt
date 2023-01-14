To Generate exe file

pyinstaller --noconfirm --onefile --windowed --icon "D:\Setpoint_Project\Setpoint_Software\assets\setpoint_logo_icon.ico" --name "Setpoint_Software" --additional-hooks-dir "D:\Setpoint_Project\Setpoint_Software"  "D:\Setpoint_Project\Setpoint_Software\Setpoint_Project_GUI.py"



tried adding hook-gcloud to
C:\Users\Amos\AppData\Local\Programs\Python\Python37\Lib\site-packages\PyInstaller\hooks