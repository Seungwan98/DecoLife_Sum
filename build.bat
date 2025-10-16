
@echo off
REM Windows에서 .exe 빌드 스크립트
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconsole --onefile --name "ExcelNamePicker" main.py
echo.
echo 빌드 완료! dist\ExcelNamePicker.exe 를 배포하세요.
pause
