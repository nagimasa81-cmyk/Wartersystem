@echo off
setlocal enabledelayedexpansion
title Build WaterSystemAnalyzer Ver42 Version7.3a DashboardDesignPreview DEBUG

if not exist "WaterSystemAnalyzer_Ver42_StableIntegrated.py" (
    echo ERROR: WaterSystemAnalyzer_Ver42_StableIntegrated.py was not found.
    pause
    exit /b 1
)

py -3.13 -m pip install --upgrade pip
py -3.13 -m pip install -r requirements_py313_pyside6_nuitka.txt

rmdir /s /q build_debug 2>nul
rmdir /s /q dist_debug 2>nul

py -3.13 -m nuitka ^
  --standalone ^
  --assume-yes-for-downloads ^
  --enable-plugin=pyside6 ^
  --include-qt-plugins=platforms,styles,imageformats ^
  --windows-console-mode=force ^
  --output-dir=build_debug ^
  --output-filename=WaterSystemAnalyzer_Ver42_DEBUG.exe ^
  WaterSystemAnalyzer_Ver42_StableIntegrated.py

if errorlevel 1 (
    echo Nuitka debug build failed.
    pause
    exit /b 1
)

mkdir dist_debug 2>nul
rmdir /s /q dist_debug\WaterSystemAnalyzer_Ver42_DEBUG 2>nul

if exist build_debug\WaterSystemAnalyzer_Ver42_StableIntegrated.dist (
    xcopy /E /I /Y build_debug\WaterSystemAnalyzer_Ver42_StableIntegrated.dist dist_debug\WaterSystemAnalyzer_Ver42_DEBUG >nul
) else (
    echo ERROR: Expected build_debug\WaterSystemAnalyzer_Ver42_StableIntegrated.dist was not found.
    dir build_debug
    pause
    exit /b 1
)

echo.
echo Debug build complete.
echo Run:
echo   dist_debug\WaterSystemAnalyzer_Ver42_DEBUG\WaterSystemAnalyzer_Ver42_DEBUG.exe
pause
