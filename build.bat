@echo off
echo ========================================
echo   Spotlight Saver - Build Script
echo ========================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python and add it to PATH.
    pause
    exit /b 1
)

:: Update pip
echo [1/4] Updating pip...
python -m pip install --upgrade pip --quiet

:: Install dependencies from requirements.txt
echo [2/4] Installing dependencies...
python -m pip install -r requirements.txt --quiet

if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)

echo [3/4] Compiling executable...
python -m PyInstaller --onefile ^
                      --windowed ^
                      --name "SpotlightSaver" ^
                      --clean ^
                      --hidden-import pystray ^
                      --hidden-import pystray._win32 ^
                      --hidden-import watchdog ^
                      --hidden-import watchdog.observers ^
                      --hidden-import watchdog.events ^
                      --hidden-import winotify ^
                      spotlight_saver.py

if errorlevel 1 (
    echo [ERROR] Build failed
    pause
    exit /b 1
)

echo [4/4] Cleaning up temporary files...
rmdir /s /q build 2>nul
del /q *.spec 2>nul

echo.
echo ========================================
echo   Build successful!
echo   Executable at: dist\SpotlightSaver.exe
echo ========================================
echo.

:: Open dist folder
explorer dist

pause
