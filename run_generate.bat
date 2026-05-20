@echo off
chcp 65001 >nul
echo.
echo  GL HW Gaming Lab — Test Report Template Generator
echo  ==================================================
echo.

if "%~1"=="" (
    echo  使用方式：將 form_data.json 拖曳到此批次檔上執行
    echo  或從命令列：run_generate.bat "FX611H_PR1_form_data.json"
    echo.
    pause
    exit /b 1
)

python "%~dp0generate_template.py" "%~1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ✗  執行失敗。請確認：
    echo     1. Python 已安裝（python --version）
    echo     2. openpyxl 已安裝（pip install openpyxl）
    echo     3. G:\HW Gaming\EE 原始報告資料\ 目錄存在
    echo.
)

pause
