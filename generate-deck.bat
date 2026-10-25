@echo off
REM Simple deck generator for Windows
REM Double-click this file to generate your deck

cd %~dp0

echo.
echo 🎨 EBR Deck Generator
echo ====================
echo.

REM Check if data file exists
if not exist my-deck.json (
    echo 📝 First time setup...
    echo Creating template: my-deck.json
    python ebr_cli.py init my-deck.json
    echo.
    echo ✅ Template created!
    echo.
    echo Next steps:
    echo 1. Edit my-deck.json with your content
    echo 2. Run this script again to generate your deck
    echo.
    pause
    exit /b
)

REM Validate
echo 📋 Validating data...
python ebr_cli.py validate my-deck.json
if errorlevel 1 (
    echo.
    echo ❌ Validation failed! Fix the errors above and try again.
    echo.
    pause
    exit /b 1
)

REM Generate
echo 🎨 Generating PowerPoint...
python ebr_cli.py generate my-deck.json output.pptx

if errorlevel 1 (
    echo.
    echo ❌ Generation failed!
    echo.
    pause
    exit /b 1
) else (
    echo.
    echo ✅ Success! Opening PowerPoint...
    echo.
    start output.pptx
)
