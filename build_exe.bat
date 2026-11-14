@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo [1/3] 필요한 패키지 설치
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :fail
set "ICON="
if exist concur_auto.ico (set ICON=--icon concur_auto.ico --add-data "concur_auto.ico;.") else (echo [주의] concur_auto.ico 파일이 없어 아이콘 없이 만듭니다. 저장소의 파일을 모두 받았는지 확인하세요.)
echo [2/3] exe 만들기
python -m PyInstaller --onefile --noconsole --name ConcurAuto --collect-all playwright --collect-all pillow_heif --hidden-import win32timezone %ICON% concur_auto.py
if errorlevel 1 goto :fail
echo [3/3] 바탕화면의 ConcurAuto.exe 바꾸기
powershell -NoProfile -Command "$t=Join-Path ([Environment]::GetFolderPath('Desktop')) 'ConcurAuto.exe'; if (-not (Test-Path -LiteralPath $t)) { exit 1 }; try { Copy-Item -LiteralPath 'dist\ConcurAuto.exe' -Destination $t -Force -ErrorAction Stop; exit 0 } catch { exit 2 }"
set "COPIED=%errorlevel%"
if "%COPIED%"=="0" echo 바탕화면의 ConcurAuto.exe를 새 것으로 바꿨습니다.
if "%COPIED%"=="1" echo 바탕화면에 ConcurAuto.exe가 없어 바꾸지 않았습니다. dist\ConcurAuto.exe를 원하는 곳에 두세요.
if "%COPIED%"=="2" echo 바탕화면의 ConcurAuto.exe를 바꾸지 못했습니다. 프로그램을 끄고 dist\ConcurAuto.exe를 직접 복사하세요.
rem Windows가 기억해 둔 예전 아이콘 지우기
ie4uinit.exe -show >nul 2>&1
ie4uinit.exe -ClearIconCache >nul 2>&1
echo.
echo 완료: dist\ConcurAuto.exe
echo exe는 어디에 두어도 됩니다. 설정 파일과 병원 목록은 문서\Concur 경비 자동 입력 폴더에 있습니다.
pause
exit /b 0
:fail
echo.
echo 실패했습니다. 위 오류 내용을 확인하세요.
pause
exit /b 1
