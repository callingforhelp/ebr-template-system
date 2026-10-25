@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM Change to script directory (quoted for spaces)
cd /d "%~dp0"
if errorlevel 1 (
    echo ERROR: Cannot access script directory
    pause
    exit /b 1
)

REM Check Python is installed
where python >nul 2>&1 || where py >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Install Python 3.8+ from: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM Determine Python command (prefer py launcher for robustness)
where py >nul 2>&1
if errorlevel 1 (
    set PYTHON=python
) else (
    set PYTHON=py -3
)

echo.
echo ========================================
echo   EBR Deck Generator
echo ========================================
echo.

REM Check if data file exists
if not exist my-deck.json (
    echo First time setup...
    echo Creating template: my-deck.json

    %PYTHON% ebr_cli.py init my-deck.json

    if errorlevel 1 (
        echo.
        echo ERROR: Failed to create template
        echo Run: pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )

    echo.
    echo Template created: my-deck.json
    echo.
    echo Next steps:
    echo   1. Edit my-deck.json with your content
    echo   2. Run this script again to generate your deck
    echo.
    pause
    exit /b 0
)

REM Generate (includes validation, no need to validate separately)
echo Generating PowerPoint deck...
%PYTHON% ebr_cli.py generate my-deck.json output.pptx

if errorlevel 1 (
    echo.
    echo ERROR: Generation failed. Check errors above.
    echo.
    pause
    exit /b 1
)

REM Verify output exists before trying to open it
if not exist output.pptx (
    echo.
    echo ERROR: output.pptx was not created
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Success! Deck created: output.pptx
echo ========================================
echo.

REM Open PowerPoint (quoted filename for safety)
start "" "output.pptx"
