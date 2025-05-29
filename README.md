# Auto-Payroll Time Card Processor

A sophisticated Python application for automated payroll processing that processes PDF and Excel time cards, applies company-specific rules for breaks and rounding, and generates comprehensive reports.

## Features

- **Multi-Company Support**: Different rules per company using JSON configuration files
- **Multiple File Formats**: Supports PDF (via pdfplumber/PyPDF2) and Excel files
- **Intelligent Time Calculation**: 
  - Configurable time rounding (e.g., 15-minute intervals)
  - Automatic break time deduction based on hours worked
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

### Command Line
```bash
python timecard_processing/processor.py configs/company_config.json input_files/timecard.pdf
```

### Programmatic Usage
```python
from timecard_processing.processor import UniversalTimeCardProcessor

processor = UniversalTimeCardProcessor('configs/company_config.json')
processor.process_file('input_files/timecard.pdf')
```

## Output

The processor generates:
- **Individual CSV files** for each employee
- **Detailed text reports** with summaries and analysis
- **Company-wide Excel summary** with all employee data
- **Processing logs** for debugging and audit trails

## File Structure
```
timecard_processing/
├── processor.py           # Main processing engine
├── configs/              # Company configuration files
├── input_files/          # Input time card files
├── output/              # Generated reports and data
└── example_usage.py     # Usage examples
``` 