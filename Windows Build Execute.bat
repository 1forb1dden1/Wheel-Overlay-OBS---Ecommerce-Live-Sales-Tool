@echo off
setlocal
cd /d "%~dp0"

echo.
echo === Energy Break — Windows executable build ===
echo.

REM Python 3.10.0 has a dis/bytecode bug that crashes PyInstaller and cx_Freeze during analysis.
py -3 -c "import sys; t=sys.version_info; ok=(t.major>3 or (t.major==3 and t.minor>10) or (t.major==3 and t.minor==10 and t.micro>=2)); raise SystemExit(0 if ok else 1)" 2>nul
if errorlevel 1 (
  echo ERROR: This PC is on Python 3.10.0, which cannot run the packager.
  echo Install Python 3.10.11 or newer ^(or 3.11 / 3.12^) from https://www.python.org/downloads/
  echo Then run this script again.
  exit /b 1
)

echo Installing PyInstaller if needed...
py -3 -m pip install "pyinstaller>=6.0" --quiet

echo.
echo Building one-file GUI executable: dist\EnergyBreak.exe
echo Includes bundled web\ for the OBS browser wheel. Users need Input List.xlsx beside the exe.
echo.

py -3 -m PyInstaller --noconfirm --clean ^
  --onefile ^
  --windowed ^
  --name EnergyBreak ^
  --hidden-import draw_prize ^
  --hidden-import wheel_html_server ^
  --hidden-import openpyxl ^
  --add-data "web;web" ^
  draw_prize_ui.py

if errorlevel 1 (
  echo.
  echo Build failed. If you are on Python 3.10.0, upgrade to 3.10.11+ and retry.
  echo For a visible error log, run the same PyInstaller command with --console instead of --windowed.
  exit /b 1
)

if exist "Input List.xlsx" (
  copy /Y "Input List.xlsx" "dist\" >nul
  echo Copied Input List.xlsx into dist\ ^(optional template for zipping^).
)

echo.
echo SUCCESS: dist\EnergyBreak.exe
echo Distribute: zip dist\EnergyBreak.exe plus Input List.xlsx ^(same folder on the target PC^).
echo.
endlocal
