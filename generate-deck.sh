#!/usr/bin/env bash
set -euo pipefail

# Simple deck generator for Mac/Linux
# Double-click this file to generate your deck

# Change to script directory
cd "$(dirname "$0")"

# Detect OS for open command
if [[ "$OSTYPE" == "darwin"* ]]; then
    OPEN_CMD="open"
elif command -v xdg-open &> /dev/null; then
    OPEN_CMD="xdg-open"
else
    OPEN_CMD=""
fi

# Check Python is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed or not in PATH"
    echo ""
    echo "Install Python 3.8+ from: https://www.python.org/downloads/"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo ""
echo "========================================"
echo "  EBR Deck Generator"
echo "========================================"
echo ""

# Check if data file exists
if [ ! -f my-deck.json ]; then
    echo "First time setup..."
    echo "Creating template: my-deck.json"

    if ! python3 ebr_cli.py init my-deck.json; then
        echo ""
        echo "ERROR: Failed to create template"
        echo "Run: pip3 install -r requirements.txt"
        echo ""
        read -p "Press Enter to exit..."
        exit 1
    fi

    echo ""
    echo "Template created: my-deck.json"
    echo ""
    echo "Next steps:"
    echo "  1. Edit my-deck.json with your content"
    echo "  2. Run this script again to generate your deck"
    echo ""
    read -p "Press Enter to exit..."
    exit 0
fi

# Generate (includes validation, no need to validate separately)
echo "Generating PowerPoint deck..."

if ! python3 ebr_cli.py generate my-deck.json output.pptx; then
    echo ""
    echo "ERROR: Generation failed. Check errors above."
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# Verify output exists before trying to open it
if [ ! -f output.pptx ]; then
    echo ""
    echo "ERROR: output.pptx was not created"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo ""
echo "========================================"
echo "  Success! Deck created: output.pptx"
echo "========================================"
echo ""

# Open PowerPoint if we have a command for it
if [ -n "$OPEN_CMD" ]; then
    if ! $OPEN_CMD output.pptx 2>/dev/null; then
        echo "Note: Could not auto-open PowerPoint"
        echo "Manually open: output.pptx"
    fi
else
    echo "Deck saved as: output.pptx"
    echo "Open it manually with your preferred application"
fi

echo ""
read -p "Press Enter to exit..."
