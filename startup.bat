@echo off
title Power Analysis System

:menu
cls
echo.
echo ========================================
echo         Power Analysis System
echo ========================================
echo.
echo Please select a program to run:
echo.
echo [1] Interactive Analyzer (interactive_analyzer.py)
echo [2] Contract Optimizer (contract_optimizer.py)
echo [3] Install Dependencies
echo [4] Exit
echo.
echo ========================================
echo.
set /p choice=Please enter option (1-4): 

if "%choice%"=="1" goto option1
if "%choice%"=="2" goto option2
if "%choice%"=="3" goto option3
if "%choice%"=="4" goto option4
goto invalid

:option1
echo.
echo Starting Interactive Analyzer...
py -3 interactive_analyzer.py
pause
goto menu

:option2
echo.
echo Starting Contract Optimizer...
py -3 contract_optimizer.py
pause
goto menu

:option3
echo.
echo Installing dependencies...
py -3 -m pip install --user -r requirements.txt
echo.
echo Dependencies installed successfully!
pause
goto menu

:option4
echo.
echo Thank you for using!
exit /b

:invalid
echo.
echo Invalid option, please try again!
pause
goto menu 