# Portfolio Excel Formatting Toolkit

A professional Python-powered toolkit for automatically formatting and visualizing investment portfolio data in Excel. Features automatic chart generation, data visualization, and a web-based interface built with Streamlit.

## 🎯 Features

- **Universal Formatter** - Auto-detects Excel file type (Type A or Type B) and applies professional formatting
- **Professional Charts** - Automatically generates Portfolio Growth Line Charts and Monthly Returns Bar Charts
- **Web Interface** - Streamlit-based GUI for simple drag-and-drop file processing
- **Smart Restructuring** - Converts Type B data sheets into professional Type A format with Executive Summary
- **Color-Coded Sections** - Professional color scheme for easy data interpretation
- **Data Visualization** - 12-month portfolio progression and monthly performance analysis

## 📦 What's Included

### Scripts
- **format_all.py** - Universal formatter with automatic chart generation
- **restructure_type_b.py** - Converts Type B (data sheets) to Type A (professional structure)
- **portfolio_formatter_app.py** - Web interface for file processing

### Documentation
- **COMPLETE_TOOLKIT_GUIDE.txt** - Full reference documentation (v2.1)
- **FORMATTING_GUIDE.txt** - Detailed formatting specifications
- **QUICK_REFERENCE.txt** - Quick start guide
- **PROJECT_SUMMARY.txt** - Project overview and accomplishments
- **APP_README.txt** - Streamlit app usage guide

### Portfolio Files
Four professionally formatted portfolio files with embedded charts:
- Portfolio Report - Traditional IRA Enhanced.xlsx
- Portfolio report_Investment_07.02.2026.xlsx
- Portfolio report_Rollover IRA_07.02.2026.xlsx
- Portfolio report_Roth IRA_07.02.2026.xlsx

## 🚀 Quick Start

### Option 1: Web Interface (Easiest)
```powershell
double-click launch_formatter.bat
```
Then drag-and-drop your Excel file into the web interface.

### Option 2: Command Line
```powershell
# For a single file
python format_all.py "YourFile.xlsx"

# For multiple files
python format_all.py *.xlsx

# For Type B restructuring
python restructure_type_b.py "YourFile.xlsx"
```

## 📋 File Types Supported

### Type A: Professional Structure
- Executive Summary sheet with KPIs and charts
- Monthly Performance sheet with detailed metrics
- Data Source sheet with original data
- **Enhanced with:**
  - Trading Activity Summary
  - Key Insights & Recommendations
  - Action Items & Strategy

### Type B: Data Sheet
- Single "Data" sheet with months as columns
- Automatically restructured to Type A format
- All calculations and formatting applied

## 🎨 Professional Features

✅ **Color Scheme:**
- Dark Blue (#1F4788) - Headers
- Medium Blue (#4472C4) - Section headers
- Light Blue (#D9E1F2) - Metric labels
- Yellow/Green/Orange/Pink - Category metrics

✅ **Automatic Charts:**
- Portfolio Growth Line Chart (12-month value progression)
- Monthly Returns Bar Chart (monthly profit visualization)

✅ **Professional Formatting:**
- Dark blue headers with white bold text
- Color-coded data sections
- Professional borders
- Optimized column widths
- Proper spacing and row heights

## 📊 Charts Explained

### Portfolio Growth Chart
- **Location:** Executive Summary, columns F-O, rows 1-19
- **Shows:** 12-month portfolio value progression
- **Purpose:** Demonstrate growth trajectory
- **Data Source:** Monthly Performance sheet, row 5

### Monthly Returns Chart
- **Location:** Executive Summary, columns F-O, rows 20-38
- **Shows:** Monthly profit/loss for each month
- **Purpose:** Demonstrate performance consistency
- **Data Source:** Monthly Performance sheet, row 8

## 🔧 Installation

### Requirements
- Python 3.7+
- openpyxl (Excel manipulation)
- streamlit (Web interface)
- lxml (XML processing)

### Setup
```powershell
# Create virtual environment
python -m venv .venv

# Activate environment
.venv\Scripts\Activate.ps1

# Install dependencies
pip install openpyxl streamlit lxml
```

## 📝 Usage Examples

### Format a single file
```powershell
python format_all.py "Portfolio_Investment.xlsx"
```

### Format all Excel files in folder
```powershell
python format_all.py *.xlsx
```

### Restructure Type B to Type A
```powershell
python restructure_type_b.py "Data_Sheet.xlsx"
```

### Run web interface
```powershell
streamlit run portfolio_formatter_app.py
```

## 📖 Documentation

See the included .txt files for detailed information:
- Start with **QUICK_REFERENCE.txt** for quick tips
- Refer to **COMPLETE_TOOLKIT_GUIDE.txt** for comprehensive reference
- Check **APP_README.txt** for web interface usage
- Review **PROJECT_SUMMARY.txt** for project overview

## ✨ Key Accomplishments

**Phase 1:** File Structure Recovery
- Recovered corrupted Excel files
- Cleaned merged cell conflicts
- Verified file integrity

**Phase 2:** Data Organization & Population
- Populated Trading Activity Summary with actual data
- Organized Excel structure professionally
- Implemented color-coded sections

**Phase 3:** Data Visualization
- Created Portfolio Growth Line Charts
- Created Monthly Returns Bar Charts
- Applied charts to all portfolio files

**Phase 4:** Code Integration & Cleanup
- Integrated chart functionality into format_all.py
- Removed temporary development scripts
- Unified formatting across all files

**Phase 5:** Web Interface & Documentation
- Built Streamlit web application
- Created comprehensive documentation
- Updated all guides to v2.1 specification

## 🎓 How It Works

1. **Detection:** Automatically identifies if file is Type A or Type B
2. **Restructuring:** If Type B, converts to professional Type A format
3. **Formatting:** Applies professional color scheme and layout
4. **Data Extraction:** Extracts portfolio and performance data
5. **Chart Generation:** Creates and embeds charts in Executive Summary
6. **Optimization:** Adjusts column widths and formatting for readability
7. **Export:** Saves formatted file back to disk

## 🔒 Local Processing

All files are processed locally on your machine:
- ✅ No upload to cloud services
- ✅ No internet required for formatting
- ✅ 100% private data handling
- ✅ Original files never modified during download

## 📌 Version

**Version 2.1** (Released February 8, 2026)
- Universal formatting with automatic charts
- Web interface integration
- Complete documentation suite
- Production-ready toolkit

## 💡 Tips & Tricks

- Close Excel files before processing for best results
- Use .xlsx format (not .xls) for compatibility
- Process one file at a time for easier troubleshooting
- Check APP_README.txt for Streamlit app tips
- See FORMATTING_GUIDE.txt for customization options

## 🤝 Support

For detailed help:
1. Check **APP_README.txt** for web interface issues
2. Review **COMPLETE_TOOLKIT_GUIDE.txt** for comprehensive reference
3. See **FORMATTING_GUIDE.txt** for formatting questions
4. Consult **QUICK_REFERENCE.txt** for quick solutions

## 📄 License

This project is provided as-is for personal portfolio management and visualization.

---

**Built with Python, openpyxl, and Streamlit**  
**Professional Portfolio Management Automation**
