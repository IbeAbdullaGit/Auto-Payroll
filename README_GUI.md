# Time Card GUI Application

A user-friendly GUI application for processing PDF time cards into Excel spreadsheets with company-specific configurations and weekly notes management.

## Features

### 📁 **Profiles Tab**
- **Profile Management**: Load, create, and delete company configuration profiles
- **Profile Information**: View detailed JSON configuration for the selected profile
- **Create New Profile**: Set up new company configurations with default settings

### ⚙️ **Configuration Tab**
- **Basic Settings**: Configure company name, daily hour caps
- **Employee Management**: Add/remove employees from the profile
- **Break Settings**: Set individual break minutes and hour thresholds for each employee
- **Save Configuration**: Save changes to the profile file

### 📝 **Notes Tab**
- **Pay Period Management**: Set the pay period for weekly notes
- **Add Notes**: Create time overrides, break overrides, missing punch corrections, and sick day notes
- **Employee Selection**: Choose from employees in the current profile
- **Override Types**:
  - `time_override`: Override in/out times with custom break
  - `break_override`: Override break duration for specific dates
  - `missing_punch_override`: Add missing punch data
  - `sick_day`: Record sick days
- **Notes Management**: View, edit, and remove notes in a table format
- **Save/Load Notes**: Save notes as JSON files for reuse

### 📊 **Processing Tab**
- **PDF Upload**: Add multiple PDF files for batch processing
- **File Management**: Remove individual files or clear all files
- **Output Location**: Set custom output directory
- **Progress Tracking**: Real-time processing progress with progress bar
- **Batch Processing**: Process multiple PDF files using the selected profile and notes

## Installation

1. **Install Python Dependencies**:
   ```bash
   pip install -r requirements_gui.txt
   ```

2. **Run the GUI Application**:
   ```bash
   python timecard_gui.py
   ```

## Usage Guide

### Step 1: Set Up a Profile
1. Go to the **Profiles** tab
2. Click **Create New Profile** and enter a company name
3. Switch to the **Configuration** tab
4. Add employees using the "Add Employee" button
5. Select each employee and set their break settings
6. Click **Save Configuration**

### Step 2: Create Weekly Notes (Optional)
1. Go to the **Notes** tab
2. Enter the pay period (e.g., "5/12/2025 - 5/23/2025")
3. Add notes for any exceptions:
   - Missing punches
   - Time corrections
   - Break overrides
   - Sick days
4. Save the notes file for future use

### Step 3: Process PDF Files
1. Go to the **Processing** tab
2. Make sure you have a profile loaded (check Profiles tab)
3. Add PDF files using **Add Files** button
4. Set output location if desired
5. Click **Process PDFs to Excel**
6. Monitor progress and wait for completion

## File Structure

```
timecard_processing/
├── configs/           # Profile configuration files (*.json)
├── note_configs/      # Weekly notes files (*.json)
├── processor.py       # Core processing engine
└── logs/             # Processing logs

output_[company]/      # Generated Excel files
├── [company]_timecard_summary_[date].xlsx
```

## Profile Configuration Format

Profiles are JSON files containing:
- Company information
- Break rules and employee-specific settings
- Rounding rules
- Daily hour caps
- Excel output format settings

## Notes Configuration Format

Notes files contain:
- Pay period information
- Array of employee-specific overrides and exceptions
- Support for time corrections, break overrides, and sick days

## Tips

1. **Profile Names**: Use descriptive names like "NewMarket" or "GreenLane"
2. **Employee Names**: Use "Last, First" format for consistency
3. **Dates**: Use formats like "Mon, 5/12" or "5/12/2025"
4. **Time Format**: Use "HH:MM AM/PM" format (e.g., "8:30 AM")
5. **Backup**: Keep backup copies of your profiles and notes

## Troubleshooting

- **No Profile Loaded**: Make sure to create and load a profile before processing
- **Missing Employees**: Add employees to the profile configuration first
- **Processing Errors**: Check the console output for detailed error messages
- **File Permissions**: Ensure write permissions for the output directory

## Support

For issues or questions, check the generated log files in the `logs/` directory for detailed processing information. 