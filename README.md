# Auto-Payroll Time Card Processor

A user-friendly GUI application for automated payroll processing that converts PDF time cards into Excel spreadsheets with company-specific rules for breaks, rounding, and comprehensive reporting.

## Features

- **Easy-to-Use GUI**: Intuitive interface with tabs for profiles, configuration, notes, and processing
- **Multi-Company Support**: Create and manage different company profiles with custom rules
- **PDF Processing**: Upload and batch process multiple PDF time card files
- **Smart Time Calculation**: 
  - Configurable time rounding (e.g., 15-minute intervals)
  - Automatic break time deduction based on hours worked
  - Daily hour caps to limit maximum hours per day
  - Overtime detection and reporting
- **Notes and Overrides**: Handle special cases, missing punches, break overrides, and time adjustments through the GUI
- **Excel Output**: Generate comprehensive Excel reports with employee summaries and daily details
- **Progress Tracking**: Real-time processing progress with visual feedback

## Quick Start

1. **Download and Install Python 3.8+** from [python.org](https://python.org)
2. **Install the application**:
   ```bash
   pip install -r requirements_gui.txt
   ```
3. **Run the application**:
   ```bash
   python timecard_gui.py
   ```

## How to Use

### Step 1: Create a Company Profile
1. Launch the application and go to the **Profiles** tab
2. Click **Create New Profile** and enter your company name
3. Switch to the **Configuration** tab to set up:
   - Add employees using the "Add Employee" button
   - Set individual break settings for each employee
   - Configure daily hour caps and rounding rules
   - Save your configuration

### Step 2: Add Weekly Notes (Optional)
1. Go to the **Notes** tab to handle special cases:
   - Missing punches
   - Time corrections  
   - Break overrides
   - Sick days
2. Enter the pay period and add notes as needed
3. Save the notes file for future use

### Step 3: Process Your Time Cards
1. Go to the **Processing** tab
2. Add your PDF time card files using **Add Files**
3. Choose your output location
4. Click **Process PDFs to Excel**
5. Monitor the progress and wait for completion

Your Excel reports will be generated automatically with employee summaries and daily breakdowns!

## Handling Special Cases

The application makes it easy to handle exceptions and special situations through the **Notes** tab:

### Common Scenarios

**Missing Punches**: When an employee worked but forgot to punch in/out
- Add a "Missing Punch Override" note
- Enter the actual work times and break duration
- The system will use these times instead of the missing punch data

**Break Adjustments**: When someone took a different break than usual
- Add a "Break Override" note  
- Specify the actual break minutes taken
- Useful for appointments, shortened lunches, etc.

**Time Corrections**: When punch times are incorrect
- Add a "Time Override" note
- Enter the correct in/out times and break duration
- Overrides any recorded punch data for that day

**Sick Days**: Document absences for record-keeping
- Add a "Sick Day" note
- Include details about the absence
- Helps with policy compliance and reporting

### Easy Date Entry

The GUI accepts flexible date formats:
- `"Mon, 5/12"` - Day and date from the time card
- `"5/14/2025"` - Full date format
- Time format: `"8:30 AM"` or `"17:30"`

## What You Get

The application generates comprehensive Excel reports in your chosen output folder:

### Excel Report Structure
- **Employee_Summary** sheet: Total hours and pay period summary for each employee
- **Daily_Details** sheet: Day-by-day breakdown showing:
  - Original punch times
  - Applied notes and overrides
  - Break deductions
  - Final calculated hours
  - Overtime tracking

### Additional Files
- **Processing logs** for troubleshooting (saved in `logs/` folder)
- **Backup copies** of your configurations and notes

## File Organization

The application automatically organizes files:
```
📁 Your Project Folder/
├── 📁 timecard_processing/
│   ├── 📁 configs/           # Your company profiles
│   ├── 📁 note_configs/      # Your weekly notes
│   └── 📁 logs/             # Processing logs
├── 📁 output_[CompanyName]/  # Generated Excel reports
└── timecard_gui.py          # The main application
```

## Tips for Success

- **Consistent Naming**: Use "Last, First" format for employee names
- **Regular Backups**: Keep copies of your profiles and notes files  
- **Test First**: Try with a small PDF file to verify your setup
- **Check Logs**: If something goes wrong, check the logs folder for details 
