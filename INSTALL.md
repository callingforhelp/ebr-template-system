# Installation Guide

Complete installation instructions for the EBR Template System CLI.

## Prerequisites

**You need Python 3.8 or higher.**

Check if you have it:

```bash
python3 --version
```

If you see `Python 3.8.0` or higher → you're ready!

If not, install Python:
- **Mac:** [Download from python.org](https://www.python.org/downloads/) or `brew install python3`
- **Windows:** [Download from python.org](https://www.python.org/downloads/)
- **Linux:** `sudo apt install python3 python3-pip` (Ubuntu/Debian)

---

## Installation

### Step 1: Get the Code

**Option A: Using Git (Recommended)**

```bash
git clone https://github.com/callingforhelp/ebr-template-system.git
cd ebr-template-system
```

**Option B: Download ZIP**

1. Visit https://github.com/callingforhelp/ebr-template-system
2. Click the green **"Code"** button
3. Click **"Download ZIP"**
4. Unzip the file
5. Open Terminal (Mac/Linux) or Command Prompt (Windows)
6. Navigate to the folder:
   ```bash
   cd ~/Downloads/ebr-template-system-main
   ```

### Step 2: Install Dependencies

```bash
pip3 install -r requirements.txt
```

**What this installs:**
- `python-pptx>=0.6.21` — Creates PowerPoint files
- `jsonschema>=4.0.0` — Validates your JSON data

**Installation takes ~30 seconds.**

### Step 3: Verify Installation

```bash
python3 ebr_cli.py --help
```

**Expected output:**
```
usage: ebr_cli.py [-h] {generate,validate,init,themes} ...

Generate Executive Business Review decks from JSON
...
```

✅ **If you see this, installation is complete!**

---

## Quick Start

### Create Your First Deck

```bash
# 1. Create a template with example data
python3 ebr_cli.py init my-deck.json

# 2. Edit my-deck.json in any text editor
#    (TextEdit, Notepad, VS Code, Sublime, etc.)

# 3. Validate your data (optional but recommended)
python3 ebr_cli.py validate my-deck.json

# 4. Generate the PowerPoint
python3 ebr_cli.py generate my-deck.json output.pptx

# 5. Open output.pptx in PowerPoint!
```

### Use the Example Data

```bash
# Generate from the included example
python3 ebr_cli.py generate artifacts/example-northstar-data.json demo.pptx

# This creates a 5-slide deck demonstrating all themes
```

---

## Command Reference

### List Available Themes

```bash
python3 ebr_cli.py themes
```

Shows all slide types you can include in your JSON.

### Create a Blank Template

```bash
python3 ebr_cli.py init my-template.json
```

Creates `my-template.json` with all required fields and examples.

### Validate JSON Data

```bash
python3 ebr_cli.py validate my-data.json
```

Checks your JSON against the schema before generating.

### Generate PowerPoint

```bash
python3 ebr_cli.py generate input.json output.pptx
```

Creates a PowerPoint file from your JSON data.

### Generate Specific Themes Only

```bash
python3 ebr_cli.py generate data.json output.pptx --themes cover performance
```

Only includes the `cover` and `performance` slides.

---

## Troubleshooting

### "python3: command not found"

**Try `python` instead:**

```bash
python --version
python ebr_cli.py --help
```

Some systems use `python` instead of `python3`.

### "pip3: command not found"

**Try one of these:**

```bash
# Option 1: Use python3 -m pip
python3 -m pip install -r requirements.txt

# Option 2: Install pip
python3 -m ensurepip --default-pip

# Option 3 (Ubuntu/Debian): Install via apt
sudo apt install python3-pip
```

### "Permission denied" when installing

**On Mac/Linux, try:**

```bash
pip3 install --user -r requirements.txt
```

**Or create a virtual environment (recommended):**

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### "ModuleNotFoundError: No module named 'pptx'"

**The installation didn't work. Install manually:**

```bash
pip3 install python-pptx jsonschema
```

### "No such file or directory: ebr_cli.py"

**Make sure you're in the right directory:**

```bash
# Check your current location
pwd

# List files (should see ebr_cli.py)
ls

# If not there, navigate to the repo folder
cd ebr-template-system
```

### JSON Validation Errors

**Example error:**

```
❌ Validation failed:
   'title' is a required property
   Path: performance
```

**What this means:**
Your `performance` section is missing the required `title` field.

**Fix:**
Add the missing field to your JSON:

```json
{
  "performance": {
    "title": "Your title here",  ← Add this
    ...
  }
}
```

### Generated Deck Looks Wrong

**Check your data format:**

```bash
# Validate first
python3 ebr_cli.py validate my-data.json

# Compare against the example
python3 ebr_cli.py validate artifacts/example-northstar-data.json
```

**Common issues:**
- Metrics must be tuples: `["70%", "Label", "Context"]`
- Table data must be arrays of arrays: `[["Header1", "Header2"], ["Row1Col1", "Row1Col2"]]`

---

## For Non-Technical Users

### Simple Double-Click Workflow

Create these wrapper scripts for your team:

**Mac/Linux: `generate-deck.sh`**

```bash
#!/bin/bash
cd "$(dirname "$0")"

# Check if data file exists
if [ ! -f "my-deck.json" ]; then
    echo "Creating template..."
    python3 ebr_cli.py init my-deck.json
    echo "✅ Edit my-deck.json, then run this script again"
    exit 0
fi

# Generate deck
echo "Generating deck..."
python3 ebr_cli.py generate my-deck.json output.pptx

# Open it
echo "✅ Done! Opening PowerPoint..."
open output.pptx
```

Make it executable:
```bash
chmod +x generate-deck.sh
```

**Windows: `generate-deck.bat`**

```batch
@echo off
cd %~dp0

REM Check if data file exists
if not exist my-deck.json (
    echo Creating template...
    python ebr_cli.py init my-deck.json
    echo ✅ Edit my-deck.json, then run this script again
    pause
    exit /b
)

REM Generate deck
echo Generating deck...
python ebr_cli.py generate my-deck.json output.pptx

REM Open it
echo ✅ Done! Opening PowerPoint...
start output.pptx
```

**Usage:**
1. Double-click the script (first time creates template)
2. Edit `my-deck.json`
3. Double-click again (generates PowerPoint)

---

## Virtual Environment (Advanced)

**Recommended for developers to isolate dependencies:**

```bash
# Create virtual environment
python3 -m venv venv

# Activate it
source venv/bin/activate  # Mac/Linux
# OR
venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt

# Use the CLI
python ebr_cli.py generate data.json output.pptx

# When done
deactivate
```

---

## System Requirements

- **Python:** 3.8 or higher
- **Disk Space:** ~50 MB for dependencies
- **Operating System:** 
  - macOS 10.14+
  - Windows 10+
  - Linux (any recent distro)
- **PowerPoint:** Microsoft PowerPoint, LibreOffice Impress, or Keynote to view output

---

## Getting Help

### Check Documentation

- **README.md** — Overview and Python API
- **INSTALL.md** — This file (installation)
- **artifacts/theme-data-schema.json** — JSON schema reference
- **artifacts/example-northstar-data.json** — Working example

### Common Questions

**Q: Do I need PowerPoint installed to generate decks?**  
A: No! The generator creates `.pptx` files that you can open later. You just need PowerPoint to *view* them.

**Q: Can I customize colors/fonts?**  
A: Yes! Edit the `DesignTokens` class in `ebr_generator_clean.py`. See README.md for details.

**Q: Can I add my own slide types?**  
A: Yes! See the "Add Custom Theme" section in README.md.

**Q: The CLI works but Python scripts don't?**  
A: The CLI is the recommended interface. Direct Python usage requires code editing.

### Report Issues

If you encounter bugs or have suggestions:

1. Check existing issues: https://github.com/callingforhelp/ebr-template-system/issues
2. Create a new issue with:
   - Your Python version (`python3 --version`)
   - Your OS (Mac/Windows/Linux)
   - Full error message
   - What you tried

---

## Uninstallation

To remove the tool:

```bash
# 1. Remove dependencies (optional)
pip3 uninstall python-pptx jsonschema

# 2. Delete the folder
rm -rf ebr-template-system  # Mac/Linux
# OR
rmdir /s ebr-template-system  # Windows
```

---

## Next Steps

✅ Installation complete? Try:

```bash
# See what themes are available
python3 ebr_cli.py themes

# Generate the example deck
python3 ebr_cli.py generate artifacts/example-northstar-data.json demo.pptx

# Create your own
python3 ebr_cli.py init my-q4-review.json
# Edit my-q4-review.json...
python3 ebr_cli.py generate my-q4-review.json Q4-Review.pptx
```

**Happy deck building! 🎨**
