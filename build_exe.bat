@echo off
setlocal enabledelayedexpansion
title Build WaterSystemAnalyzer Ver42 Version7.3a DashboardDesignPreview

echo ============================================================
echo WaterSystemAnalyzer Ver42 Version7.3a DashboardDesignPreview
echo Python 3.13 + PySide6 + Nuitka
echo No pandas / No matplotlib / No numpy
echo ============================================================
echo.

if not exist "WaterSystemAnalyzer_Ver42_StableIntegrated.py" (
    echo ERROR: WaterSystemAnalyzer_Ver42_StableIntegrated.py was not found.
    pause
    exit /b 1
)

py -3.13 --version
if errorlevel 1 (
    echo Python 3.13 was not found.
    pause
    exit /b 1
)

py -3.13 -m pip install --upgrade pip
py -3.13 -m pip install -r requirements_py313_pyside6_nuitka.txt
if errorlevel 1 (
    echo pip install failed.
    pause
    exit /b 1
)

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul

py -3.13 -m nuitka ^
  --standalone ^
  --assume-yes-for-downloads ^
  --enable-plugin=pyside6 ^
  --include-qt-plugins=platforms,styles,imageformats ^
  --windows-console-mode=disable ^
  --output-dir=build ^
  --output-filename=WaterSystemAnalyzer_Ver42.exe ^
  WaterSystemAnalyzer_Ver42_StableIntegrated.py

if errorlevel 1 (
    echo Nuitka build failed.
    pause
    exit /b 1
)

mkdir dist 2>nul
rmdir /s /q dist\WaterSystemAnalyzer_Ver42 2>nul

if exist build\WaterSystemAnalyzer_Ver42_StableIntegrated.dist (
    xcopy /E /I /Y build\WaterSystemAnalyzer_Ver42_StableIntegrated.dist dist\WaterSystemAnalyzer_Ver42 >nul
) else (
    echo ERROR: Expected build\WaterSystemAnalyzer_Ver42_StableIntegrated.dist was not found.
    dir build
    pause
    exit /b 1
)

echo.
echo Build complete.
echo Run:
echo   dist\WaterSystemAnalyzer_Ver42\WaterSystemAnalyzer_Ver42.exe
echo.
pause
