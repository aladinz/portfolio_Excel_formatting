================================================================================
PORTFOLIO FORMATTER WEB APP
Streamlit-Based Visual Interface
================================================================================

WHAT IS THIS?
=============
A web-based interface that makes it easy to format Excel files by simply
uploading them. No command line needed!


HOW TO RUN
==========

OPTION 1: Quick Start (Easiest)
───────────────────────────────
1. Double-click: launch_formatter.bat
2. A browser window opens automatically
3. Drag and drop your Excel file, click format, download
4. Done!


OPTION 2: From Command Line
────────────────────────────
1. Open PowerShell in this folder
2. Run: .venv\Scripts\python.exe -m streamlit run portfolio_formatter_app.py
3. Browser opens automatically to http://localhost:8501


HOW TO USE THE APP
==================

1. Open the web interface (see "HOW TO RUN" above)
2. Click the upload box or drag your Excel file onto it
3. The app automatically formats your file
4. Download the result with the download button
5. Done! File is formatted with:
   ✓ Professional colors and styling
   ✓ Optimized layout and spacing
   ✓ Charts (if applicable)
   ✓ Color-coded sections


FEATURES
========
✓ Drag-and-drop file upload
✓ Real-time processing feedback
✓ Shows success/error messages
✓ One-click download
✓ Runs 100% locally (no uploading to internet)
✓ No installation required (if .venv already set up)
✓ Beautiful, responsive interface


KEYBOARD SHORTCUTS
==================
While the app is running:
  Ctrl+C      Stop the app
  R           Rerun the app
  C           Clear cache
  Q           Quit


TROUBLESHOOTING
===============

Q: Browser doesn't open automatically
A: Manually go to http://localhost:8501

Q: "Python not found" error
A: Make sure the .venv folder exists in this directory

Q: Port 8501 already in use
A: The app will automatically try the next port (8502, 8503, etc.)
   Check your browser address bar for the correct URL

Q: File processing fails
A: Check the Details section in the app for error messages
   Make sure your file is a valid Excel file (.xlsx or .xls)


STOPPING THE APP
================
Press Ctrl+C in the PowerShell window to stop the server
Or close the browser window and stop the PowerShell process


KEYBOARD COMMANDS IN STREAMLIT
===============================
While using the app:
• r - Rerun the script
• c - Clear cache
• v - Toggle code visibility

In the terminal running the app:
• Ctrl+C - Stop the server
• Ctrl+Z - Force exit (not recommended)


ADVANCED: CUSTOM PORT
=====================
To run on a specific port instead of 8501:

.venv\Scripts\python.exe -m streamlit run portfolio_formatter_app.py --server.port=9000

Then visit: http://localhost:9000


WHAT HAPPENS TO YOUR FILES?
===========================
✓ Files are processed locally on your machine
✓ No data leaves your computer
✓ Temporary files are deleted automatically
✓ Your original file is NOT modified
✓ Download the formatted version when ready


INTEGRATION WITH YOUR TOOLS
============================
This Streamlit app is a visual wrapper around format_all.py
It runs the same formatting you'd do from the command line, but with a UI:

Command Line:
  python format_all.py "MyFile.xlsx"

Streamlit App:
  Upload file → Same format_all.py runs → Download result


TIPS FOR BEST RESULTS
=====================
1. Make sure files are .xlsx format (not .xls)
2. Close the file in Excel before uploading
3. For Type B files, ensure data structure is standard
4. Use for both Type A and Type B files
5. Process one file at a time


FOR HELP
========
See the other documentation files:
• FORMATTING_GUIDE.txt - Complete reference
• QUICK_REFERENCE.txt - Quick tips
• PROJECT_SUMMARY.txt - Project overview
• COMPLETE_TOOLKIT_GUIDE.txt - Full toolkit reference
