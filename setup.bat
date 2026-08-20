@echo off
setlocal

echo ============================================
echo   Vimar Datasheet Tool - First-time setup
echo ============================================
echo.
echo This only needs to run once. It will:
echo   1. Check that Python is installed
echo   2. Install the tool's requirements
echo   3. Put a "Vimar Datasheet Tool" icon on your Desktop
echo.

if not exist "%~dp0app.py" (
    echo ============================================
    echo   Please extract the ZIP file first
    echo ============================================
    echo.
    echo This is still running from inside the ZIP file, not from a
    echo real folder on your computer - the files it needs are not
    echo actually here yet.
    echo.
    echo Please close this window and:
    echo   1. Right-click the downloaded ZIP file
    echo   2. Choose "Extract All..."
    echo   3. Pick a folder ^(e.g. your Desktop^) and click Extract
    echo   4. Open the NEW folder that appears - it will look like a
    echo      normal folder, not a zipped one
    echo   5. Double-click setup.bat from inside that new folder
    echo.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this computer.
    echo.
    echo Please install it first:
    echo   1. A download page will open - click the yellow "Download Python" button
    echo   2. Run the installer
    echo   3. IMPORTANT: tick the box that says "Add python.exe to PATH"
    echo      before clicking Install
    echo   4. Once it finishes, double-click this setup.bat file again
    echo.
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Found Python. Installing requirements - this can take a few minutes...
echo.
python -m pip install --upgrade pip >nul 2>nul
python -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo Something went wrong installing requirements.
    echo Please copy the message above and send it to Hayssam.
    pause
    exit /b 1
)

echo.
echo Installing the browser component - this can take a minute...
python -m playwright install chromium
if errorlevel 1 (
    echo.
    echo Something went wrong installing the browser component.
    echo Please copy the message above and send it to Hayssam.
    pause
    exit /b 1
)

echo.
echo Creating your desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Vimar Datasheet Tool.lnk'); $s.TargetPath = '%~dp0start.bat'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = '%~dp0vimar_icon.ico'; $s.Description = 'Vimar Datasheet Tool'; $s.Save()"

echo.
echo ============================================
echo   Setup complete!
echo.
echo   A "Vimar Datasheet Tool" icon is now on your Desktop.
echo   Double-click it any time you want to use the tool -
echo   you will not need to do any of this again.
echo ============================================
echo.
pause
