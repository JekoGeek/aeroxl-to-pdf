@echo off
setlocal
py -3.8-32 -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "AeroXL2PDF_Win8" ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  app.py
endlocal
