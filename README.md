# Auto-Payroll Time Card Processor

A sophisticated Python application for automated payroll processing that processes PDF and Excel time cards, applies company-specific rules for breaks and rounding, and generates comprehensive reports.

## Features

- **Multi-Company Support**: Different rules per company using JSON configuration files
- **Multiple File Formats**: Supports PDF (via pdfplumber/PyPDF2) and Excel files
- **Notes and Overrides**: Process special cases, missing punches, break overrides, and time adjustments
- **Intelligent Time Calculation**: 
  - Configurable time rounding (e.g., 15-minute intervals)
  - Automatic break time deduction based on hours worked
  - Daily hour caps to limit maximum hours per day
  - Overtime detection and reporting
- **Comprehensive Reporting**: Individual employee reports, CSV exports, and company summaries
- **Data Analysis**: Attendance rates, missing punches, average hours, etc.
- **Robust Error Handling**: Graceful handling of missing data and format variations

## Setup

1. **Install Python 3.8+** if not already installed
2. **Install dependencies**: 
   ```bash
   pip install -r requirements.txt
   ```
3. **Create configuration files** for each company (see examples in `configs/`)
4. **Prepare input files** in the `input_files/` directory

## Configuration

Create a JSON configuration file for each company in the `configs/` directory:

```json
{
  "company_name": "Your Company Name",
  "file_type": "pdf",
  "daily_hour_cap": 8.0,
  "rounding_rules": {
    "interval_minutes": 15,
    "threshold_minutes": 8
  },
  "break_config": {
    "break_rules": {
      "minimum_hours_for_break": 6.0,
      "full_day_break": 30,
      "long_day_break": 60
    }
  }
}
```

### Configuration Options

- **`daily_hour_cap`**: Maximum hours per day (e.g., 8.0 for 8 hours max)
  - Example: If someone works 9:15 with 1 hour break = 8:15 net, but gets capped at 8:00
  - Set to `null` or omit to disable daily capping
- **`rounding_rules`**: Time rounding configuration
  - `interval_minutes`: Round to nearest interval (e.g., 15 minutes)
  - `threshold_minutes`: Minimum minutes to round up (e.g., 8 minutes)
- **`break_config`**: Automatic break deduction rules

## Usage

### Basic Processing
```bash
python timecard_processing/processor.py configs/company_config.json input_files/timecard.pdf
```

### Processing with Notes and Overrides
```bash
python timecard_processing/processor.py configs/company_config.json input_files/timecard.pdf --notes note_configs/notes_example.json
```

### Programmatic Usage
```python
from timecard_processing.processor import UniversalTimeCardProcessor

# Basic processing
processor = UniversalTimeCardProcessor('configs/company_config.json')
processor.process_file('input_files/timecard.pdf')

# Processing with notes
processor = UniversalTimeCardProcessor('configs/company_config.json', 'note_configs/notes.json')
processor.process_file('input_files/timecard.pdf')
```

## Notes and Overrides

The processor supports special notes files to handle exceptions, missing punches, and manual overrides that aren't captured in the main timecard data.

### Notes File Format

Create a JSON file with pay period information and notes:

```json
{
  "pay_period": "5/12/2025 - 5/23/2025",
  "notes": [
    {
      "employee": "Smith, John",
      "date": "Mon, 5/12",
      "type": "missing_punch_override",
      "time_in": "8:00 AM",
      "time_out": "5:00 PM",
      "break_minutes": 30,
      "note": "Forgot to punch - worked 8 AM to 5 PM with 30min break"
    },
    {
      "employee": "Johnson, Sarah",
      "date": "5/14/2025",
      "type": "break_override",
      "value": 15,
      "note": "Had appointment - only took 15min break instead of 30min"
    },
    {
      "employee": "Brown, Mike",
      "date": "5/16/2025",
      "type": "time_override",
      "time_in": "9:00 AM",
      "time_out": "6:00 PM",
      "break_minutes": 60,
      "note": "Schedule change - worked 9-6 with 1 hour lunch"
    },
    {
      "employee": "Davis, Lisa",
      "date": "5/20/2025",
      "type": "sick_day",
      "note": "Called out sick - affects stat holiday eligibility"
    }
  ]
}
```

### Note Types

1. **`missing_punch_override`**: Add missing timecard entries
   - Use when an employee worked but didn't punch in/out
   - Requires: `time_in`, `time_out`, optionally `break_minutes`

2. **`break_override`**: Override automatic break calculation
   - Use when break time differs from standard rules
   - Requires: `value` (break minutes)

3. **`time_override`**: Override recorded punch times
   - Use when punch times are incorrect
   - Requires: `time_in`, `time_out`, optionally `break_minutes`

4. **`sick_day`**: Document sick days or other absences
   - For record-keeping and policy compliance
   - Requires: `note` describing the absence

### Date Formats

Notes support flexible date formats:
- `"Mon, 5/12"` - Day and date format
- `"5/14/2025"` - Full date format
- Date ranges: `"5/12/2025 - 5/16/2025"` for multi-day notes

## Output

The processor generates:
- **Individual CSV files** for each employee
- **Detailed text reports** with summaries and analysis
- **Company-wide Excel summary** with all employee data
  - **Employee_Summary** sheet: Total hours per employee
  - **Daily_Details** sheet: Day-by-day breakdown with applied notes
- **Processing logs** for debugging and audit trails

## File Structure
```
timecard_processing/
├── processor.py           # Main processing engine
├── configs/              # Company configuration files
├── note_configs/         # Notes and override files
├── input_files/          # Input time card files
├── output/              # Generated reports and data
├── logs/                # Processing logs
└── example_usage.py     # Usage examples
``` 
