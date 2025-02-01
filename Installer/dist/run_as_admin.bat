@echo off
:: Check for elevated permissions
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with elevated permissions
) else (
    echo Requesting elevated permissions
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Run the application
start "" "ANPCTechSupport.exe"