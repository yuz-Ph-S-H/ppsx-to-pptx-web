@echo off
setlocal
cd /d "%~dp0"
set PORT=8000

where py >nul 2>nul
if %errorlevel%==0 (
  start "" "http://localhost:%PORT%"
  echo.
  echo PPSX to PPTX local website is running:
  echo http://localhost:%PORT%
  echo.
  echo Keep this window open. Press Ctrl+C to stop.
  py -m http.server %PORT% --bind 127.0.0.1
  goto :eof
)

where python >nul 2>nul
if %errorlevel%==0 (
  start "" "http://localhost:%PORT%"
  echo.
  echo PPSX to PPTX local website is running:
  echo http://localhost:%PORT%
  echo.
  echo Keep this window open. Press Ctrl+C to stop.
  python -m http.server %PORT% --bind 127.0.0.1
  goto :eof
)

echo.
echo Python was not found.
echo Install Python 3, or run any static HTTP server in this folder.
echo Example: npx serve .
echo.
pause
