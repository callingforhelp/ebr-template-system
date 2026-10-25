#!/bin/bash
# Simple deck generator for Mac/Linux
# Double-click this file to generate your deck

cd "$(dirname "$0")"

echo "🎨 EBR Deck Generator"
echo "===================="
echo ""

# Check if data file exists
if [ ! -f "my-deck.json" ]; then
    echo "📝 First time setup..."
    echo "Creating template: my-deck.json"
    python3 ebr_cli.py init my-deck.json
    echo ""
    echo "✅ Template created!"
    echo ""
    echo "Next steps:"
    echo "1. Edit my-deck.json with your content"
    echo "2. Run this script again to generate your deck"
    echo ""
    read -p "Press Enter to exit..."
    exit 0
fi

# Validate
echo "📋 Validating data..."
if ! python3 ebr_cli.py validate my-deck.json; then
    echo ""
    echo "❌ Validation failed! Fix the errors above and try again."
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# Generate
echo "🎨 Generating PowerPoint..."
python3 ebr_cli.py generate my-deck.json output.pptx

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Success! Opening PowerPoint..."
    echo ""
    open output.pptx
else
    echo ""
    echo "❌ Generation failed!"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi
