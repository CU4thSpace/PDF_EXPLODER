# PDF to PowerPoint Converter - Installation Guide

This guide will help you set up and run the PDF to PowerPoint converter script on your computer, even if you're new to Python.

## What This Script Does
This script converts PDF files into PowerPoint presentations by turning each PDF page into an image and placing it on a slide.

## Prerequisites
You'll need to install Python and some additional software on your computer.

---

## Step 1: Install Python

### For Windows:
1. Go to [python.org](https://python.org/downloads/)
2. Click "Download Python" (get the latest version)
3. **Important**: During installation, check the box "Add Python to PATH"
4. Click "Install Now"

### For Mac:
1. Go to [python.org](https://python.org/downloads/)
2. Click "Download Python" (get the latest version)
3. Open the downloaded file and follow the installation steps
4. Python should be ready to use

### For Linux (Ubuntu/Debian):
```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

---

## Step 2: Install Poppler (Required for PDF Processing)

This script needs a program called "Poppler" to read PDF files.

### For Windows:
1. Go to [poppler-windows.github.io](https://poppler-windows.github.io/)
2. Download the latest version
3. Extract the zip file to `C:\poppler`
4. Add `C:\poppler\bin` to your system PATH:
   - Press Windows + R, type `sysdm.cpl`, press Enter
   - Click "Environment Variables"
   - Under "System Variables", find "Path" and click "Edit"
   - Click "New" and add `C:\poppler\bin`
   - Click OK on all windows

### For Mac:
1. Install Homebrew first (if you don't have it):
   - Open Terminal (found in Applications > Utilities)
   - Paste this command and press Enter:
     ```bash
     /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
     ```
2. Install Poppler:
   ```bash
   brew install poppler
   ```

### For Linux (Ubuntu/Debian):
```bash
sudo apt install poppler-utils
```

---

## Step 3: Download and Save the Script

1. Copy the Python script code
2. Open a text editor (Notepad on Windows, TextEdit on Mac, or any code editor)
3. Paste the code
4. Save the file as `pdf_to_ppt.py` (make sure it ends with `.py`)
5. Remember where you saved it (like your Desktop or Documents folder)

---

## Step 4: Set Up a Virtual Environment

A virtual environment keeps this project's dependencies separate from other Python projects.

### Open Terminal/Command Prompt:
- **Windows**: Press Windows + R, type `cmd`, press Enter
- **Mac**: Press Cmd + Space, type "Terminal", press Enter
- **Linux**: Press Ctrl + Alt + T

### Navigate to Your Script Location:
```bash
# Example: if you saved the script on your Desktop
cd Desktop

# Or if in Documents folder:
cd Documents
```

### Create the Virtual Environment:
```bash
python -m venv pdf_converter_env
```

### Activate the Virtual Environment:

**Windows:**
```bash
pdf_converter_env\Scripts\activate
```

**Mac/Linux:**
```bash
source pdf_converter_env/bin/activate
```

You should see `(pdf_converter_env)` at the beginning of your command prompt now.

---

## Step 5: Install Python Dependencies

With your virtual environment activated, install the required packages:

```bash
pip install python-pptx pdf2image Pillow
```

Wait for the installation to complete. You should see messages about successful installations.

---

## Step 6: Update the Script for Your System

You need to edit one line in the script to match your system:

1. Open the `pdf_to_ppt.py` file in a text editor
2. Find this line (around line 11):
   ```python
   poppler_path = "/opt/homebrew/bin"  # <-- set this explicitly
   ```

3. Change it based on your system:

   **Windows:**
   ```python
   poppler_path = r"C:\poppler\bin"
   ```

   **Mac (if you used Homebrew):**
   ```python
   poppler_path = "/opt/homebrew/bin"  # Keep as is
   ```
   
   **Mac (Intel-based Macs might need):**
   ```python
   poppler_path = "/usr/local/bin"
   ```

   **Linux:**
   ```python
   poppler_path = None  # Let it use system PATH
   ```

4. Save the file

---

## Step 7: Test the Installation

Let's make sure everything works:

1. Make sure your virtual environment is still activated (you should see `(pdf_converter_env)` in your terminal)
2. Run the script:
   ```bash
   python pdf_to_ppt.py
   ```

3. If everything is set up correctly:
   - A dialog box should appear asking if you want to select a PDF or folder
   - This means the script is working!
   - Click "Cancel" for now since this is just a test

---

## How to Use the Script

### Method 1: Interactive Mode
1. Activate your virtual environment (Step 4 commands)
2. Navigate to where your script is saved
3. Run: `python pdf_to_ppt.py`
4. Follow the dialog boxes to select your PDF and output location

### Method 2: Drag and Drop (Advanced)
You can drag a PDF file onto the script file to convert it automatically.

---

## Troubleshooting

### "Python not found" Error:
- Make sure you checked "Add Python to PATH" during installation
- Try using `python3` instead of `python`

### "No module named..." Error:
- Make sure your virtual environment is activated
- Run the pip install command again

### "Poppler not found" Error:
- Double-check your Poppler installation
- Verify the `poppler_path` in your script matches your installation

### PDF Won't Convert:
- Try a different PDF file
- Check if the PDF is password-protected
- Increase the DPI value in the script (change `dpi=200` to `dpi=300`)

---

## When You're Done

To deactivate the virtual environment when you're finished:
```bash
deactivate
```

## Running the Script Later

Whenever you want to use the script again:
1. Open Terminal/Command Prompt
2. Navigate to your script location
3. Activate the virtual environment: 
   - Windows: `pdf_converter_env\Scripts\activate`
   - Mac/Linux: `source pdf_converter_env/bin/activate`
4. Run the script: `python pdf_to_ppt.py`

---

That's it! You should now be able to convert PDF files to PowerPoint presentations.
