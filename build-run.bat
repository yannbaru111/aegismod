@echo off
chcp 65001 > nul
title Aegis Mod Builder
echo.
echo ================================================
echo   Aegis Mod - Build Start
echo ================================================
echo.

cd /d "%~dp0"

echo [1/5] Checking Python...
python --version
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Python not found.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)
echo OK.
echo.

echo [2/5] Installing required packages...
python -m pip install pyinstaller pystray pillow sentence-transformers
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install packages.
    pause
    exit /b 1
)
echo OK.
echo.

echo [3/5] Generating icon...
python -c "from PIL import Image, ImageDraw; import os; img=Image.new('RGBA',(256,256),(0,0,0,0)); d=ImageDraw.Draw(img); d.ellipse([4,4,252,252],fill='#1a0a2e'); s=4; d.polygon([(32*s,4*s),(56*s,14*s),(56*s,36*s),(32*s,60*s),(8*s,36*s),(8*s,14*s)],fill='#9147ff'); d.polygon([(32*s,10*s),(50*s,18*s),(50*s,34*s),(32*s,54*s),(14*s,34*s),(14*s,18*s)],fill='#a970ff'); d.rectangle([30*s,14*s,34*s,50*s],fill='#6020cc'); d.line([40*s,12*s,24*s,44*s],fill='white',width=10); d.polygon([(24*s,44*s),(20*s,52*s),(28*s,48*s)],fill='white'); d.rectangle([39*s,8*s,43*s,16*s],fill='#dddddd'); img.save('src\\aegismod.ico',format='ICO',sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])"
if %errorlevel% neq 0 (
    echo [WARN] Icon generation failed, building without icon.
    set ICON_OPT=
) else (
    echo OK.
    set ICON_OPT=--icon src\aegismod.ico
)
echo.

echo [4/5] Building exe... (this takes 1-2 minutes)
python -m PyInstaller --onedir --windowed --noupx --name "AegisMod" %ICON_OPT% --add-data "src\ai_engine.py;." src\main.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)
echo OK.
echo.

echo ================================================
echo   SUCCESS! dist\AegisMod\AegisMod.exe is ready!
echo   Use the dist\AegisMod folder to run the app.
echo ================================================
echo.
start explorer dist
pause
