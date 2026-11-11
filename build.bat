@echo off
cd /d "%~dp0"

echo Installing build tools...
py -m pip install --upgrade pymupdf pywin32 pillow pyinstaller
if errorlevel 1 goto fail

echo Building PDFSS.exe...
py -m PyInstaller --onefile --windowed --noconfirm --clean --name "PDFSS" ^
  --hidden-import win32timezone ^
  --exclude-module numpy ^
  --exclude-module onnxruntime ^
  --exclude-module pymupdf4llm ^
  --exclude-module pymupdf_layout ^
  --exclude-module pymupdf.layout ^
  --exclude-module networkx ^
  --exclude-module yaml ^
  --exclude-module psutil ^
  --exclude-module tabulate ^
  --exclude-module pandas ^
  --exclude-module matplotlib ^
  PDFSS.py
if errorlevel 1 goto fail

echo.
echo Done. The app is at: %~dp0dist\PDFSS.exe
pause
exit /b 0

:fail
echo.
echo Build failed. See the errors above.
pause
exit /b 1
