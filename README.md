# 🚀 Company Website Scraper

A powerful, automated web scraping tool designed to find and extract company websites from Excel files containing company names. The scraper uses advanced web search techniques and provides both GUI and command-line interfaces.

## 📋 Table of Contents
- [Features](#-features)
- [Requirements](#-requirements)
- [Technologies Used](#-technologies-used)
- [Libraries & Dependencies](#-libraries--dependencies)
- [Installation](#-installation)
- [Usage](#-usage)
- [File Structure](#-file-structure)
- [Configuration](#-configuration)
- [Troubleshooting](#-troubleshooting)
- [Limitations](#-limitations)
- [License](#-license)

## ✨ Features

### 🎯 **Core Functionality**
- **Automated Website Discovery**: Finds company websites using multiple search engines
- **Excel Integration**: Reads company names from Excel files and writes results back
- **Smart Search Algorithm**: Uses first 4 words of company names for optimal search results
- **Duplicate Prevention**: Skips companies that already have valid websites
- **Progress Tracking**: Real-time progress monitoring and automatic saving

### 🛡️ **Anti-Detection Features**
- **Multiple Search Engines**: DuckDuckGo and Startpage support
- **Random User Agents**: Rotates through different browser signatures
- **Human-like Behavior**: Random delays and mouse movements
- **Fresh Browser Sessions**: Creates new Chrome instances for each search
- **Rate Limiting**: Configurable delays between requests (15-25 seconds)

### 🎨 **User Interface**
- **Modern GUI**: Beautiful, intuitive graphical interface
- **Real-time Logging**: Live activity feed with colored status messages
- **Statistics Dashboard**: Success rate, progress tracking, and counters
- **File Validation**: Excel file structure and content verification
- **Progress Bar**: Visual progress indication with detailed status

### 🔍 **Smart Filtering**
- **URL Validation**: Ensures discovered URLs are properly formatted
- **General Website Filter**: Excludes social media, directories, and generic sites
- **Company-specific Results**: Focuses on actual company websites
- **URL Cleaning**: Standardizes and cleans discovered URLs

## 🔧 Requirements

### **System Requirements**
- **Operating System**: Windows 10/11 (64-bit)
- **RAM**: Minimum 4GB (8GB recommended)
- **Storage**: 500MB free space
- **Internet**: Stable internet connection for web searches

### **Software Requirements**
- **Python**: Version 3.8 or newer
- **Google Chrome**: Latest version installed
- **ChromeDriver**: Matching your Chrome browser version
- **Excel Software**: Microsoft Excel or compatible (.xlsx support)

### **File Requirements**
- Excel file with sheet named: `1. Exhibitor List (Input)`
- Company names in column B (starting from row 9)
- Website column C (where results will be written)
- Maximum 300 companies per file

## 💻 Technologies Used

### **Programming Language**
- **Python 3.8+**: Core development language

### **Web Automation**
- **Selenium WebDriver**: Browser automation and control
- **Chrome WebDriver**: Chrome browser integration
- **Headless Browsing**: Background web scraping

### **User Interface**
- **Tkinter**: Native Python GUI framework
- **Threading**: Multi-threaded operations for responsive UI
- **Queue System**: Thread-safe communication

### **Data Processing**
- **Excel Processing**: Reading and writing .xlsx files
- **URL Parsing**: Advanced URL validation and cleaning
- **Regular Expressions**: Pattern matching and validation

## 📚 Libraries & Dependencies

### **Core Dependencies**
```python
selenium>=4.0.0        # Web automation and browser control
openpyxl>=3.0.0        # Excel file reading/writing
```

### **Built-in Libraries**
```python
tkinter                # GUI framework (included with Python)
threading              # Multi-threading support
queue                  # Thread-safe data structures
logging                # Application logging
time                   # Time operations and delays
random                 # Random number generation
os                     # Operating system interface
sys                    # System-specific parameters
re                     # Regular expressions
urllib.parse           # URL parsing utilities
datetime               # Date and time handling
```

### **Browser Dependencies**
- **ChromeDriver**: Chrome browser automation driver
- **Google Chrome**: Web browser for automation

## 🚀 Installation

### **Method 1: Automated Setup (Recommended)**
1. **Download** the complete scraper folder
2. **Run** `setup.bat` as Administrator
3. **Follow** the on-screen instructions
4. **Download ChromeDriver** manually if prompted

### **Method 2: Manual Installation**
1. **Install Python 3.8+** from [python.org](https://python.org/downloads/)
2. **Install dependencies**:
   ```bash
   pip install selenium openpyxl
   ```
3. **Download ChromeDriver**:
   - Go to [chromedriver.chromium.org](https://chromedriver.chromium.org/)
   - Download version matching your Chrome browser
   - Extract to: `chromedriver-win64/chromedriver.exe`

### **Verify Installation**
```bash
python -c "import selenium, openpyxl; print('All dependencies installed successfully!')"
```

## 📖 Usage

### **GUI Mode (Recommended)**
1. **Launch GUI**: Double-click `start_GUI.bat` or run:
   ```bash
   python scraper_gui.py
   ```
2. **Select Excel File**: Click "Browse" and choose your Excel file
3. **Validate File**: Click "Validate" to check file structure
4. **Start Scraping**: Click "Start Scraping" to begin
5. **Monitor Progress**: Watch real-time progress and logs

### **Command Line Mode**
```bash
python scraper_fixed.py
```

### **Excel File Format**
Your Excel file must have:
- **Sheet Name**: `1. Exhibitor List (Input)`
- **Column A**: Row numbers/IDs
- **Column B**: Company names (starting row 9)
- **Column C**: Website URLs (results written here)

**Example:**
| A | B | C |
|---|---|---|
| 1 | Company Name | Website URL |
| 2 | Tech Solutions Inc | https://techsolutions.com |
| 3 | Digital Marketing Pro | https://digitalmarketingpro.com |

## 📁 File Structure

```
Company Website Scraper/
├── 📄 scraper_gui.py           # Main GUI application
├── 📄 scraper_fixed.py         # Core scraping engine
├── 📄 start_GUI.bat            # GUI launcher
├── 📄 setup.bat                # Installation script
├── 📄 README.md                # This documentation
├── 📄 requirements.txt         # Python dependencies
├── 📄 debug.log                # Application logs
├── 📁 chromedriver-win64/      # ChromeDriver folder
│   └── chromedriver.exe        # Chrome automation driver
├── 📁 chrome_profile/          # Chrome browser profiles
└── 📁 __pycache__/             # Python cache files
```

## ⚙️ Configuration

### **Search Settings**
- **Search Engines**: DuckDuckGo, Startpage
- **Search Terms**: First 4 words of company name
- **Delay Between Requests**: 15-25 seconds (randomized)
- **Maximum Results**: 5 results per search engine

### **Filtering Rules**
The scraper automatically excludes:
- Social media sites (Facebook, LinkedIn, Twitter, etc.)
- Business directories (Yellow Pages, Yelp, etc.)
- Generic platforms (Amazon, eBay, etc.)
- Government and educational sites
- News and media websites

### **File Limits**
- **Maximum Companies**: 300 per Excel file
- **Auto-Save Frequency**: Every 5 processed companies
- **Progress Persistence**: Resumes from last saved position

## 🐛 Troubleshooting

### **Common Issues**

#### **"Python not found" Error**
- Install Python from [python.org](https://python.org/downloads/)
- Ensure "Add Python to PATH" was checked during installation
- Restart command prompt after installation

#### **"ChromeDriver not found" Error**
- Download ChromeDriver matching your Chrome version
- Place in `chromedriver-win64/chromedriver.exe`
- Check Chrome version: `chrome://version/`

#### **"Excel file not found" Error**
- Ensure Excel file path is correct
- Check file is not open in Excel
- Verify file has `.xlsx` extension

#### **"Sheet not found" Error**
- Excel file must contain sheet: `1. Exhibitor List (Input)`
- Check sheet name spelling and capitalization
- Ensure company names start from row 9, column B

#### **Slow Performance**
- Close other applications to free RAM
- Check internet connection speed
- Increase delay between requests in code

#### **High Failure Rate**
- Companies with generic names may fail
- Try more specific company names
- Check if websites exist manually

### **Log Files**
- **GUI Logs**: Displayed in application log window
- **Debug Logs**: Written to `debug.log` file
- **Error Details**: Check logs for specific error messages

## ⚠️ Limitations

### **Technical Limitations**
- **Maximum Companies**: 300 companies per run
- **Search Scope**: Uses only first 4 words of company names
- **Browser Dependency**: Requires Chrome browser
- **Windows Only**: Designed for Windows operating systems

### **Search Limitations**
- **Generic Names**: Companies with very generic names may not be found
- **New Companies**: Very new companies may not have web presence
- **Regional Variations**: May favor certain geographic regions
- **Language**: Optimized for English company names

### **Rate Limiting**
- **Search Delays**: 15-25 seconds between searches (to avoid blocking)
- **Daily Limits**: Search engines may impose daily limits
- **IP Restrictions**: May require VPN if blocked

### **Legal Considerations**
- **Terms of Service**: Respect search engine terms of service
- **Rate Limits**: Do not modify delays to be too aggressive
- **Data Usage**: Use scraped data responsibly and legally

## 🔄 Updates & Maintenance

### **Keeping Updated**
- **ChromeDriver**: Update when Chrome browser updates
- **Python Packages**: Run `pip install --upgrade selenium openpyxl`
- **Search Engines**: Monitor for changes in search result structure

### **Performance Optimization**
- **Clear Browser Cache**: Delete `chrome_profile` folder periodically
- **Log Rotation**: Archive old `debug.log` files
- **Excel Optimization**: Split large files into smaller batches

## 📞 Support

### **Getting Help**
- Check this README for common solutions
- Review `debug.log` file for error details
- Verify all requirements are met
- Test with a small sample file first

### **Best Practices**
- **Start Small**: Test with 10-20 companies first
- **Regular Backups**: Backup Excel files before processing
- **Monitor Progress**: Watch logs for any issues
- **Respect Limits**: Don't exceed 300 companies per file

## 📄 License

This project is provided as-is for educational and business purposes. Please ensure compliance with all applicable laws and terms of service when using this software.

---
