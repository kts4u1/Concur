@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo [1/2] 필요한 패키지 설치
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :fail
echo [2/2] exe 만들기
python -m PyInstaller --onefile --noconsole --name ConcurAuto --collect-all playwright --collect-all pillow_heif --hidden-import win32timezone concur_auto.py
if errorlevel 1 goto :fail
echo.
echo 완료: dist\ConcurAuto.exe
echo exe를 원하는 폴더로 옮겨서 실행하세요. 설정 파일과 병원 목록은 exe 옆에 자동으로 생깁니다.
pause
exit /b 0
:fail
echo.
echo 실패했습니다. 위 오류 내용을 확인하세요.
pause
exit /b 1
