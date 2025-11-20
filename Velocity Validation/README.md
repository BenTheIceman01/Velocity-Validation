# HD SUPPLY™ VELOCITY VALIDATOR
**Developed by: Ben F. Benjamaa**

A modern, self-contained application for validating velocity codes against Snowflake data with an advanced black and yellow HD Supply™ branded interface.

---

## 🌟 Features

- **Modern GUI** - Sleek black background with bright yellow HD Supply™ branding
- **Snowflake Integration** - Direct connection to Snowflake database
- **VLOOKUP Functionality** - Automatically matches velocity codes from Snowflake
- **Validation** - Compares Current_Velocity with PROPOSED_VELOCITY
- **Excel Export** - Formatted Excel output with color-coded results
- **Real-time Progress** - Visual feedback during processing
- **Self-contained Executable** - No Python installation required for end users

---

## 📋 Requirements

### For Running the Python Script:
- Python 3.8 or higher
- Required packages (see `requirements.txt`)

### For Using the Executable:
- Windows OS
- No additional requirements!

---

## 🚀 Quick Start

### Option 1: Run as Python Script

1. **Install Dependencies:**
   ```bash
   Double-click: install_dependencies.bat
   ```
   Or manually:
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the Application:**
   ```bash
   Double-click: run_app.bat
   ```
   Or manually:
   ```bash
   python velocity_validator_app.py
   ```

### Option 2: Build Standalone Executable

1. **Install Dependencies First** (if not already done)

2. **Build Executable:**
   ```bash
   Double-click: build_executable.bat
   ```
   Or manually:
   ```bash
   pyinstaller --clean --onefile velocity_validator.spec
   ```

3. **Use the Executable:**
   - Find it in: `dist\HD_Supply_Velocity_Validator.exe`
   - Double-click to run
   - No Python needed!

---

## 📝 How to Use

### Step 1: Prepare Your Data File
- Ensure your Excel/CSV file contains columns:
  - `ITEM` (required)
  - `LOC` (required)
  - `PROPOSED_VELOCITY` (required for comparison)

### Step 2: Launch Application
- Run the application using one of the methods above

### Step 3: Select File
- Click **"BROWSE FILE"**
- Select your Excel or CSV file

### Step 4: Enter Snowflake Credentials
- **Account**: Your Snowflake account identifier
- **Username**: Your Snowflake username
- **Password**: Your Snowflake password
- **Warehouse**: Snowflake warehouse name
- **Database**: `EDP` (default)
- **Schema**: `STD_JDA` (default)

### Step 5: Process Data
- Click **"PROCESS DATA"**
- Wait for processing to complete
- Output file will be saved in the same directory as input file

---

## 📊 Output Format

The application creates an Excel file with these additions:

### New Columns:
1. **Current_Velocity** - Retrieved from Snowflake `UDC_VELOCITY_CODE`
2. **Match** - True/False comparison with `PROPOSED_VELOCITY`

### Formatting:
- **Header**: Black background with yellow text (HD Supply™ branding)
- **Current_Velocity Column**: Light yellow background
- **Match Column**: 
  - Green background = Match (True)
  - Red background = Mismatch (False)

### Output Filename:
```
Velocity_Validated_YYYYMMDD_HHMMSS.xlsx
```

---

## 🗄️ Snowflake Query

The application executes this query:

```sql
SELECT
    ITEM,
    LOC,
    UDC_VELOCITY_CODE
FROM
    SKUEXTRACT
```

---

## 🎨 Color Scheme

**HD Supply™ Branding:**
- Background: Pure Black (`#000000`)
- Primary Text: Bright Yellow (`#FFED4E`)
- Accents: Gold (`#FFD700`)
- Highlights: Dark Gray (`#1A1A1A`, `#2D2D2D`)

---

## 📦 File Structure

```
Velocity Validation/
├── velocity_validator_app.py    # Main application
├── requirements.txt              # Python dependencies
├── velocity_validator.spec       # PyInstaller configuration
├── install_dependencies.bat      # Dependency installer
├── run_app.bat                   # Application launcher
├── build_executable.bat          # Executable builder
├── README.md                     # This file
└── dist/                         # Generated executables (after build)
    └── HD_Supply_Velocity_Validator.exe
```

---

## 🔧 Troubleshooting

### Issue: "Module not found" error
**Solution:** Run `install_dependencies.bat` to install all required packages

### Issue: Snowflake connection fails
**Solution:** 
- Verify your credentials
- Ensure you have network access to Snowflake
- Check warehouse, database, and schema names

### Issue: "Column not found" error
**Solution:** Ensure your input file contains required columns: `ITEM`, `LOC`, `PROPOSED_VELOCITY`

### Issue: Executable build fails
**Solution:** 
- Ensure PyInstaller is installed: `pip install pyinstaller`
- Run `pip install -r requirements.txt` first

---

## 📞 Support

For issues or questions, contact: **Ben F. Benjamaa**

---

## 📄 License

© 2025 HD Supply™ - All Rights Reserved

---

## 🔄 Version History

### Version 1.0 (November 2025)
- Initial release
- Modern black & yellow HD Supply™ interface
- Snowflake integration
- VLOOKUP functionality
- Excel export with formatting
- Standalone executable support

---

**HD SUPPLY™ - Velocity Validator**
*Developed by: Ben F. Benjamaa*
