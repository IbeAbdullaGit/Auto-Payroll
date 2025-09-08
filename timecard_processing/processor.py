#!/usr/bin/env python3
"""
Simplified Multi-Company Time Card Processor
Processes time cards for different companies using company-specific config files.
No company names needed in the time card files themselves.
"""

import re
import json
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import argparse
from typing import List, Dict, Any, Tuple, Optional
import logging

# Required packages:
# pip install PyPDF2 pandas openpyxl pdfplumber

try:
    import pdfplumber
    PDF_LIBRARY = 'pdfplumber'
except ImportError:
    try:
        import PyPDF2
        PDF_LIBRARY = 'PyPDF2'
    except ImportError:
        print("Please install pdfplumber or PyPDF2: pip install pdfplumber")
        exit(1)

class UniversalTimeCardProcessor:
    """Process time cards for any company using config files."""
    
    def __init__(self, config_file: str, notes_file: str = None):
        """Initialize with a company-specific config file and optional notes file."""
        self.config = self.load_config(config_file)
        self.company_name = self.config.get('company_name', 'Unknown Company')
        self.file_type = self.config.get('file_type', 'pdf')
        self.output_base_dir = Path(self.config.get('output_directory', f'output_{self.company_name}'))
        self.output_base_dir.mkdir(exist_ok=True)
        
        # Track the current input file for unique naming
        self.current_input_file = None
        
        # Set up logging first
        self.setup_logging()
        
        # Load notes/exceptions if provided (after logger is ready)
        self.notes = self.load_notes(notes_file) if notes_file else []
    
    def load_config(self, config_file: str) -> Dict[str, Any]:
        """Load company-specific configuration."""
        with open(config_file, 'r') as f:
            return json.load(f)
    
    def load_notes(self, notes_file: str) -> List[Dict[str, Any]]:
        """Load weekly notes/exceptions."""
        try:
            with open(notes_file, 'r') as f:
                notes_data = json.load(f)
            notes = notes_data.get('notes', [])
            self.notes = notes  # Update the instance variable
            self.logger.info(f"Loaded {len(notes)} notes for pay period: {notes_data.get('pay_period', 'Unknown')}")
            return notes
        except FileNotFoundError:
            self.logger.warning(f"Notes file not found: {notes_file}")
            self.notes = []  # Ensure instance variable is set
            return []
        except Exception as e:
            self.logger.error(f"Error loading notes file: {e}")
            self.notes = []  # Ensure instance variable is set
            return []
    
    def setup_logging(self):
        """Set up logging configuration."""
        # Create logs directory if it doesn't exist
        log_file = Path("logs") / f"processing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_file.parent.mkdir(exist_ok=True)
        
        logging.basicConfig(
            level=logging.INFO,  # Show info, warnings and errors
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                # Removed StreamHandler to disable console output
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"Processing for {self.company_name}")
    
    def round_to_interval(self, minutes: int) -> int:
        """Round minutes based on company-specific rules."""
        rules = self.config.get('rounding_rules', {})
        interval = rules.get('interval_minutes', 15)
        
        base_interval = (minutes // interval) * interval
        remainder = minutes % interval
        
        # New rule: Round up only if within 2 minutes of the next interval
        # Exception: Between 5:45 and 6:00 (345-360 minutes), allow rounding from 5:57 (357 minutes)
        if base_interval == 345:  # 5:45 mark
            # Special case for 5:45-6:00 range
            if remainder >= 12:  # 5:57 and above (345 + 12 = 357 = 5:57)
                return base_interval + interval  # Round up to 6:00
            else:
                return base_interval  # Stay at 5:45
        else:
            # Normal rule: round up only if within 2 minutes of next interval
            if remainder >= 13:  # 13, 14 minutes (within 2 minutes of next 15-min mark)
                return base_interval + interval
            else:
                return base_interval
    
    def time_to_minutes(self, time_str: str) -> int:
        """Convert HH:MM format to total minutes."""
        try:
            hours, minutes = map(int, time_str.split(':'))
            return hours * 60 + minutes
        except:
            return 0
    
    def minutes_to_time(self, minutes: int) -> str:
        """Convert minutes to HH:MM format."""
        hours = minutes // 60
        mins = minutes % 60
        return f"{hours:02d}:{mins:02d}"
    
    def find_notes_for_employee_date(self, employee_name: str, date: str) -> List[Dict[str, Any]]:
        """Find all notes for a specific employee and date."""
        matching_notes = []
        for note in self.notes:
            # Check if employee matches
            if note.get('employee') != employee_name:
                continue
                
            # Check single date
            if 'date' in note:
                note_date = note.get('date', '')
                if self.normalize_date(note_date) == self.normalize_date(date):
                    matching_notes.append(note)
            
            # Check date range
            elif 'date_range' in note:
                date_range = note.get('date_range', '')
                if self.date_in_range(date, date_range):
                    matching_notes.append(note)
                    
        return matching_notes
    
    def date_in_range(self, date: str, date_range: str) -> bool:
        """Check if a date falls within a date range."""
        try:
            # Parse the date range "5/12/2025 - 5/16/2025"
            if ' - ' not in date_range:
                return False
                
            start_date_str, end_date_str = date_range.split(' - ')
            start_date = datetime.strptime(start_date_str.strip(), "%m/%d/%Y")
            end_date = datetime.strptime(end_date_str.strip(), "%m/%d/%Y")
            
            # Parse the check date - handle both "Day, M/D" and "M/D/YYYY" formats
            check_date_str = self.normalize_date(date)
            if '/' in check_date_str and check_date_str.count('/') == 1:
                # Add year if missing (format: "5/12")
                check_date_str += "/2025"  # Assume current year, could be made smarter
            
            # Parse the check date
            if check_date_str.count('/') == 2:
                check_date = datetime.strptime(check_date_str, "%m/%d/%Y")
            else:
                return False
                
            # Check if date is in range (inclusive)
            return start_date <= check_date <= end_date
            
        except Exception as e:
            self.logger.warning(f"Error checking date range {date_range} for date {date}: {e}")
            return False
    
    def normalize_date(self, date_str: str) -> str:
        """Normalize date string for comparison."""
        try:
            # Handle different date formats
            if ',' in date_str:
                # Extract M/D from "Day, M/D" format
                date_part = date_str.split(', ')[1]
            else:
                date_part = date_str
            
            # Now handle M/D vs M/D/YYYY
            if '/' in date_part:
                parts = date_part.split('/')
                if len(parts) == 3:
                    # M/D/YYYY format - normalize and keep year
                    month, day, year = parts
                    return f"{int(month)}/{int(day)}/{year}"
                elif len(parts) == 2:
                    # M/D format - normalize and add default year
                    month, day = parts
                    return f"{int(month)}/{int(day)}/2025"
            
            return date_part
        except:
            return date_str
    
    def apply_note_overrides(self, employee_name: str, date: str, time_in: str, time_out: str, calculated_break: int) -> Tuple[str, str, int, str]:
        """Apply any note overrides for this employee/date. Returns (time_in, time_out, break_minutes, note)."""
        notes = self.find_notes_for_employee_date(employee_name, date)
        applied_notes = []
        
        final_time_in = time_in
        final_time_out = time_out
        final_break = calculated_break
        
        for note in notes:
            note_type = note.get('type', '')
            note_text = note.get('note', '')
            applied_notes.append(note_text)
            
            if note_type == 'break_override':
                final_break = note.get('value', calculated_break)
                self.logger.info(f"Applied break override for {employee_name} on {date}: {calculated_break}min -> {final_break}min ({note_text})")
                
            elif note_type == 'time_override' or note_type == 'missing_punch_override':
                final_time_in = note.get('time_in', time_in)
                final_time_out = note.get('time_out', time_out)
                if 'break_minutes' in note:
                    final_break = note.get('break_minutes')
                self.logger.info(f"Applied {note_type} for {employee_name} on {date}: {time_in}-{time_out} -> {final_time_in}-{final_time_out} ({note_text})")
                
            elif note_type == 'sick_day':
                self.logger.info(f"Sick day noted for {employee_name} on {date}: {note_text}")
                
            elif note_type == 'misc':
                self.logger.info(f"Misc note for {employee_name} on {date}: {note_text}")
                
            elif note_type == 'roe':
                self.logger.info(f"ROE note for {employee_name} on {date}: {note_text}")
                
        combined_notes = "; ".join(applied_notes) if applied_notes else ""
        return final_time_in, final_time_out, final_break, combined_notes
    
    def calculate_break_time(self, employee_name: str, date: str, gross_hours: float, net_minutes_after_rounding: int = None) -> int:
        """Calculate break time based on configuration with support for escalating breaks."""
        break_config = self.config.get('break_config', {})
        rules = break_config.get('break_rules', {})
        standard_minimum_hours = rules.get('minimum_hours_for_break', 6.0)
        reduced_minimum_hours = 5.0  # Reduced minimum for 15-minute and 30-minute breaks
        
        # Use rounded net hours if provided (for final check), otherwise use gross hours
        hours_to_check = gross_hours
        if net_minutes_after_rounding is not None:
            hours_to_check = max(gross_hours, net_minutes_after_rounding / 60)  # Use whichever is higher
        
        # NEW: Check for escalating break schedules (30 min normally, 60 min if > 7 hours)
        escalating_breaks = break_config.get('escalating_breaks', {})
        if employee_name in escalating_breaks:
            escalating_rule = escalating_breaks[employee_name]
            
            # Check each threshold in descending order (highest first)
            thresholds = sorted(escalating_rule.get('thresholds', []), key=lambda x: x['hours'], reverse=True)
            
            for threshold in thresholds:
                if hours_to_check >= threshold['hours']:
                    return threshold['break_minutes']
            
            # If no threshold met, return 0
            return 0
        
        # Check for individual employee hour thresholds (simple version)
        employee_hour_thresholds = break_config.get('employee_hour_thresholds', {})
        
        # FIRST: Check for employee-specific breaks with custom thresholds
        employee_breaks = break_config.get('employee_breaks', {})
        if employee_name in employee_breaks:
            employee_break = employee_breaks[employee_name]
            
            # Check if this employee has a custom hour threshold
            if employee_name in employee_hour_thresholds:
                custom_threshold = employee_hour_thresholds[employee_name]
                if hours_to_check >= custom_threshold:
                    return employee_break
                else:
                    return 0
            
            # If no custom threshold, use standard logic
            # If employee has 15-minute or 30-minute break, use reduced minimum hours (5 instead of 6)
            if employee_break in [15, 30]:
                if hours_to_check >= reduced_minimum_hours:
                    return employee_break
                else:
                    return 0
            # For other employee-specific break amounts, use standard minimum hours
            else:
                if hours_to_check >= standard_minimum_hours:
                    return employee_break
                else:
                    return 0
        
        # Rest of the method remains the same...
        # SECOND: Check for weekly break schedules
        weekly_schedules = break_config.get('weekly_break_schedules', {})
        if employee_name in weekly_schedules:
            weekly_break = self.get_weekly_break_time(employee_name, date, weekly_schedules[employee_name])
            # If weekly schedule specifies 15 or 30 minutes, use reduced minimum hours
            if weekly_break in [15, 30]:
                if hours_to_check >= reduced_minimum_hours:
                    return weekly_break
                else:
                    return 0
            # Otherwise, use standard minimum hours for other weekly break amounts
            else:
                if hours_to_check >= standard_minimum_hours:
                    return weekly_break
                else:
                    return 0
        
        # THIRD: Check for day-specific breaks
        day_specific = break_config.get('day_specific_breaks', {})
        if employee_name in day_specific:
            day_break = self.get_day_specific_break_time(employee_name, date, day_specific[employee_name])
            # If day-specific schedule specifies 15 or 30 minutes, use reduced minimum hours
            if day_break in [15, 30]:
                if hours_to_check >= reduced_minimum_hours:
                    return day_break
                else:
                    return 0
            # Otherwise, use standard minimum hours for other day-specific break amounts
            else:
                if hours_to_check >= standard_minimum_hours:
                    return day_break
                else:
                    return 0
        
        # LAST: Apply standard break rules based on hours worked
        if hours_to_check < standard_minimum_hours:
            return 0
        elif hours_to_check >= rules.get('long_day_threshold', 9.0):
            return rules.get('long_day_break', 60)
        else:
            return rules.get('full_day_break', 60)
    
    def get_weekly_break_time(self, employee_name: str, date: str, schedule: dict) -> int:
        """Get break time based on weekly schedule."""
        try:
            # Parse the date to determine which week we're in
            # Assuming date format is "Mon, 5/12" or similar
            if ',' in date:
                date_part = date.split(', ')[1]  # Get "5/12"
                date_obj = datetime.strptime(f"{date_part}/2025", "%m/%d/%Y")
            else:
                # Handle other date formats if needed
                return schedule.get('week1_break', 30)
            
            # Determine if this is week 1 or week 2 based on the date
            # Week 1: May 12-16 (week starting May 12)
            # Week 2: May 20-23 (week starting May 20)  
            day_of_month = date_obj.day
            
            if day_of_month <= 16:  # First week
                break_time = schedule.get('week1_break', 30)
                self.logger.info(f"Weekly schedule: {employee_name} gets {break_time}min break (Week 1)")
                return break_time
            else:  # Second week
                break_time = schedule.get('week2_break', 30)
                self.logger.info(f"Weekly schedule: {employee_name} gets {break_time}min break (Week 2)")
                return break_time
                
        except Exception as e:
            self.logger.warning(f"Error processing weekly schedule for {employee_name}: {e}")
            return schedule.get('week1_break', 30)
    
    def get_day_specific_break_time(self, employee_name: str, date: str, schedule: dict) -> int:
        """Get break time based on day of the week."""
        try:
            # Extract day of week from date string like "Mon, 5/12"
            if ',' in date:
                day_of_week = date.split(',')[0].strip()  # Get "Mon"
                
                # Map abbreviations to full names
                day_mapping = {
                    'Mon': 'Monday', 'Tue': 'Tuesday', 'Wed': 'Wednesday',
                    'Thu': 'Thursday', 'Fri': 'Friday', 'Sat': 'Saturday', 'Sun': 'Sunday'
                }
                
                full_day = day_mapping.get(day_of_week, day_of_week)
                
                if full_day in schedule:
                    break_time = schedule[full_day]
                    self.logger.info(f"Day-specific break: {employee_name} gets {break_time}min break on {full_day}")
                    return break_time
                    
        except Exception as e:
            self.logger.warning(f"Error processing day-specific schedule for {employee_name}: {e}")
            
        return 30  # Default fallback
    
    def calculate_net_hours(self, time_in: str, time_out: str, 
                          employee_name: str, date: str) -> Tuple[str, str, int, str]:
        """Calculate net hours worked after break deduction and rounding. Returns (net_time, gross_time, break_minutes, notes)."""
        if 'Missed' in time_in or 'Missed' in time_out:
            return "00:00", "00:00", 0, ""
        
        try:
            # Apply note overrides FIRST
            final_time_in, final_time_out, override_break, notes = self.apply_note_overrides(
                employee_name, date, time_in, time_out, -1  # Use -1 to indicate no override
            )
            
            # Check for fixed hours employees (before calculating from punch times)
            break_config = self.config.get('break_config', {})
            fixed_hours_employees = break_config.get('fixed_hours_employees', {})
            
            if employee_name in fixed_hours_employees:
                # Check what types of notes exist for this date
                notes_for_date = self.find_notes_for_employee_date(employee_name, date)
                
                # These override types bypass fixed hours and use actual timecard times
                # sick_day still uses fixed hours
                complete_override_types = ['time_override', 'missing_punch_override', 'break_override']
                has_complete_override = any(note.get('type') in complete_override_types for note in notes_for_date)
                
                if not has_complete_override:
                    # No overrides - use fixed hours
                    fixed_hours = fixed_hours_employees[employee_name]
                    
                    # Simple fixed hours - no breaks
                    break_minutes = 0
                    net_hours_str = f"{int(fixed_hours)}:{int((fixed_hours % 1) * 60):02d}"
                    gross_hours_str = net_hours_str  # Same as net since we ignore breaks
                    
                    # Add note about fixed hours
                    if notes:
                        notes += f"; Fixed {fixed_hours} hours"
                    else:
                        notes = f"Fixed {fixed_hours} hours"
                        
                    self.logger.info(f"Applied fixed hours for {employee_name}: {fixed_hours} hours (no overrides)")
                    return net_hours_str, gross_hours_str, break_minutes, notes
                # If there WAS any override, continue to normal processing below (overrides bypass fixed hours)
            
            # Parse AM/PM times (use overridden times)
            time_in_obj = datetime.strptime(final_time_in, "%I:%M %p")
            time_out_obj = datetime.strptime(final_time_out, "%I:%M %p")
            
            # Calculate gross minutes
            gross_minutes = int((time_out_obj - time_in_obj).total_seconds() / 60)
            
            # Handle overnight shift
            if gross_minutes < 0:
                gross_minutes += 24 * 60
            
            gross_hours = gross_minutes / 60
            
            # Use override break if notes specified it (including 0 minutes)
            if override_break >= 0:  # -1 means no override, 0+ means override with specific value
                break_minutes = override_break
            else:
                # Check if gross time rounds up to break threshold
                rounded_gross_minutes = self.round_to_interval(gross_minutes)
                
                # Calculate break based on gross hours OR what the gross would round to
                break_minutes = self.calculate_break_time(employee_name, date, gross_hours, rounded_gross_minutes)
            
            # Calculate net minutes after break
            net_minutes = gross_minutes - break_minutes
            if net_minutes < 0:
                net_minutes = 0
            
            # Apply rounding to final net minutes
            rounded_net_minutes = self.round_to_interval(net_minutes)
            
            # SPECIAL LOGIC: If gross rounded up to qualify for break, ensure fair net hours
            if break_minutes > 0 and override_break == 0:
                rounded_gross_minutes = self.round_to_interval(gross_minutes)
                # Only apply this special logic if gross time was 5:45-5:59 and rounded up to 6:00
                if (gross_minutes >= 5 * 60 + 45 and gross_minutes < 6 * 60 and  # Between 5:45 and 5:59
                    rounded_gross_minutes >= 6 * 60):  # Rounds up to 6+ hours
                    # Ensure they get at least 5 hours net (5 hours = 300 minutes)
                    if rounded_net_minutes < 5 * 60:  # Less than 5 hours
                        rounded_net_minutes = 5 * 60  # Give them 5 hours
            
            # Apply daily hour cap if configured
            daily_cap_hours = self.config.get('daily_hour_cap', None)
            if daily_cap_hours is not None:
                max_daily_minutes = int(daily_cap_hours * 60)
                if rounded_net_minutes > max_daily_minutes:
                    original_time = self.minutes_to_time(rounded_net_minutes)
                    capped_time = self.minutes_to_time(max_daily_minutes)
                    self.logger.warning(f"Capping {employee_name} on {date}: {original_time} -> {capped_time}")
                    rounded_net_minutes = max_daily_minutes
            
            # Convert to time format
            gross_time = self.minutes_to_time(gross_minutes)
            net_time_rounded = self.minutes_to_time(rounded_net_minutes)
            
            return net_time_rounded, gross_time, break_minutes, notes
            
        except Exception as e:
            self.logger.error(f"Error calculating hours for {employee_name} on {date}: {e}")
            return "00:00", "00:00", 0, ""
    
    def process_file(self, file_path: str):
        """Process a time card file based on its type."""
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Store current input file for unique naming
        self.current_input_file = file_path.stem  # filename without extension
        
        self.logger.info(f"Processing file: {file_path}")
        
        # Determine processing method based on file type
        if file_path.suffix.lower() == '.pdf':
            employees = self.process_pdf(file_path)
        elif file_path.suffix.lower() in ['.xlsx', '.xls']:
            employees = self.process_excel(file_path)
        else:
            raise ValueError(f"Unsupported file type: {file_path.suffix}")
        
        # Add any employees from notes who don't appear in the timecard data
        if self.notes:
            self.logger.info(f"=== Processing notes: found {len(self.notes)} notes ===")
            note_only_employees = self.create_employees_from_notes(employees)
            employees.extend(note_only_employees)
            self.logger.info(f"=== About to call add_missing_days_from_notes ===")
            
            # Add missing days from notes to existing employees
            self.add_missing_days_from_notes(employees)
            self.logger.info(f"=== Finished add_missing_days_from_notes ===")
        else:
            self.logger.info("=== No notes found ===")
        
        # Generate outputs for all employees
        self.generate_all_outputs(employees)
        
        self.logger.info(f"Processing complete. Output saved to: {self.output_base_dir}")
    
    def process_pdf(self, pdf_path: Path) -> List[Dict[str, Any]]:
        """Process PDF time card file."""
        text = self.extract_text_from_pdf(str(pdf_path))
        return self.parse_pdf_employees(text)
    
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract text from PDF."""
        text = ""
        
        if PDF_LIBRARY == 'pdfplumber':
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        # Clean up the PDF text by removing CropBox warnings
                        lines = page_text.split('\n')
                        clean_lines = [line for line in lines if not line.startswith('CropBox missing')]
                        clean_text = '\n'.join(clean_lines)
                        text += clean_text + "\n"
        else:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
        
        return text
    
    def parse_pdf_employees(self, text: str) -> List[Dict[str, Any]]:
        """Parse employee data from PDF text."""
        employees = []
        
        # This PDF format has each employee on their own page/section
        # Split by "TIME CARD REPORT" which appears at the start of each employee
        sections = text.split('TIME CARD REPORT')
        
        for section in sections[1:]:  # Skip the first empty section
            if not section.strip():
                continue
                
            # Extract employee name using the new format
            name_match = re.search(r'Name\s*:\s*([^,\n]+(?:,\s*[^,\n]+)*)', section)
            if not name_match:
                continue
                
            name = name_match.group(1).strip()
            # Remove "Approval Status" if it got included
            if "Approval Status" in name:
                name = name.split("Approval Status")[0].strip()
            
            self.logger.info(f"Processing employee: {name}")
            
            employee = self.parse_new_format_employee(name, section)
            if employee:
                employees.append(employee)
        
        return employees
    
    def parse_new_format_employee(self, name: str, data_section: str) -> Dict[str, Any]:
        """Parse individual employee data for the new format."""
        try:
            employee = {
                'name': name,
                'employee_group': self.extract_field(data_section, r'Employee Group\s*:\s*([^\n]+)'),
                'pay_period': self.extract_new_format_pay_period(data_section),
                'approval_status': self.extract_field(data_section, r'Approval Status\s*:\s*([^\n]+)'),
                'time_entries': [],
                'total_hours_gross': '00:00',
                'total_hours_net': '00:00',
                'analysis': {}
            }
            
            # Extract and process time entries
            time_entries = self.extract_new_format_time_entries(data_section, name)
            employee['time_entries'] = time_entries
            
            # Calculate total net hours
            total_net_minutes = sum(
                self.time_to_minutes(entry['net_hours_rounded']) 
                for entry in time_entries
            )
            employee['total_hours_net'] = self.minutes_to_time(total_net_minutes)
            
            # Calculate total gross hours from daily entries
            total_gross_minutes = sum(
                self.time_to_minutes(entry['gross_hours']) 
                for entry in time_entries
            )
            employee['total_hours_gross'] = self.minutes_to_time(total_gross_minutes)
            
            # If we couldn't calculate from entries, try extracting from PDF
            if total_gross_minutes == 0:
                gross_total_match = re.search(r'Pay Period Totals\s+(\d+:\d+)', data_section)
                if gross_total_match:
                    employee['total_hours_gross'] = gross_total_match.group(1)
            
            # Perform analysis
            employee['analysis'] = self.analyze_employee_data(employee)
            
            return employee
            
        except Exception as e:
            self.logger.error(f"Error parsing employee {name}: {str(e)}")
            return None
    
    def extract_field(self, text: str, pattern: str) -> str:
        """Extract a field using regex."""
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""
    
    def extract_new_format_pay_period(self, text: str) -> str:
        """Extract pay period dates from new format."""
        from_match = re.search(r'From\s*:\s*(\d+/\d+/\d+)', text)
        to_match = re.search(r'To\s*:\s*(\d+/\d+/\d+)', text)
        
        if from_match and to_match:
            return f"{from_match.group(1)} - {to_match.group(1)}"
        return ""
    
    def extract_new_format_time_entries(self, text: str, employee_name: str) -> List[Dict[str, Any]]:
        """Extract time entries from the new format."""
        entries = []
        
        # Simple approach: make AM/PM fully optional in the pattern
        # This should catch the Wed, 5/28 case where it's "11:29" without AM
        pattern = r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun),\s*(\d+/\d+)\s+(\d+:\d+(?:\s*[AP]M)?)\s+(\d+:\d+)(?:\s*[AP]M)?\s+(\d+:\d+(?:\s*[AP]M)?)\s+(\d+:\d+)(?:\s*[AP]M)?\s+(\w+)\s+(\d+:\d+)'
        
        matches = re.findall(pattern, text, re.MULTILINE)
        
        for match in matches:
            day, date, time_in_actual, time_out_actual, time_in_edited, time_out_edited, dept, daily_hours = match
            
            # Fix incomplete times (missing AM/PM)
            def fix_time_format(time_str, is_time_out=False, time_in_ref=""):
                """Add AM/PM to time if missing, using smart logic."""
                if 'AM' in time_str or 'PM' in time_str:
                    return time_str.strip()
                
                time_str = time_str.strip()
                hour = int(time_str.split(':')[0])
                
                if is_time_out:
                    # For time_out, if it's afternoon hours (1-6 PM), it's likely PM
                    if hour >= 1 and hour <= 6:
                        return time_str + ' PM'
                    elif hour >= 7 and hour <= 11:
                        # Could be morning or evening, check context
                        if 'AM' in time_in_ref:
                            return time_str + ' AM'
                        else:
                            return time_str + ' PM'
                    elif hour == 12:
                        return time_str + ' PM'  # Noon
                    else:
                        return time_str + ' AM'
                else:
                    # For time_in, morning hours are typically AM
                    if hour >= 6 and hour <= 11:
                        return time_str + ' AM'
                    elif hour == 12:
                        return time_str + ' PM'  # Noon
                    else:
                        return time_str + ' AM'
            
            # Fix the times with smart logic
            time_in_actual = fix_time_format(time_in_actual, False)
            time_out_actual = fix_time_format(time_out_actual, True, time_in_actual)
            time_in_edited = fix_time_format(time_in_edited, False)
            time_out_edited = fix_time_format(time_out_edited, True, time_in_edited)
            
            # Use the edited times for calculation (columns 3 and 4)
            time_in = time_in_edited
            time_out = time_out_edited
            
            # Apply note overrides FIRST to get the final times to display
            final_time_in, final_time_out, override_break, override_notes = self.apply_note_overrides(
                employee_name, f"{day}, {date}", time_in, time_out, -1
            )
            
            # Calculate net hours with breaks and rounding using overridden times
            net_hours_rounded, gross_hours, break_minutes, notes = self.calculate_net_hours(
                time_in, time_out, employee_name, f"{day}, {date}"
            )
            
            # Combine override notes with calculation notes
            combined_notes = notes
            if override_notes and notes != override_notes:
                combined_notes = override_notes if not notes else f"{override_notes}; {notes}"
            
            entry = {
                'date': f"{day}, {date}",
                'day_of_week': day,
                'date_only': date,
                'time_in': final_time_in,  # Use overridden time for display
                'time_out': final_time_out,  # Use overridden time for display
                'gross_hours': gross_hours,
                'break_minutes': break_minutes,
                'net_hours_rounded': net_hours_rounded,
                'daily_hours_from_pdf': daily_hours,
                'notes': combined_notes,
                'is_complete': final_time_in != 'Missed' and final_time_out != 'Missed'
            }
            entries.append(entry)
            
            self.logger.info(f"Found entry for {employee_name} on {day}, {date}: {time_in} - {time_out} = {net_hours_rounded}")
        
        return entries
    
    def process_excel(self, excel_path: Path) -> List[Dict[str, Any]]:
        """Process Excel time card file."""
        excel_config = self.config.get('excel_config', {})
        format_type = excel_config.get('format', 'standard')
        
        if format_type == 'standard':
            return self.process_standard_excel(excel_path, excel_config)
        elif format_type == 'sheet_per_employee':
            return self.process_sheet_per_employee_excel(excel_path, excel_config)
        else:
            raise ValueError(f"Unknown Excel format: {format_type}")
    
    def process_standard_excel(self, excel_path: Path, excel_config: Dict) -> List[Dict[str, Any]]:
        """Process Excel with all employees in one sheet."""
        df = pd.read_excel(excel_path)
        
        # Get column mappings
        columns = excel_config.get('columns', {
            'employee_name': 'Employee Name',
            'date': 'Date',
            'time_in': 'Time In',
            'time_out': 'Time Out'
        })
        
        employees = {}
        
        # Group by employee
        for _, row in df.iterrows():
            employee_name = str(row[columns['employee_name']])
            
            if employee_name not in employees:
                employees[employee_name] = {
                    'name': employee_name,
                    'time_entries': [],
                    'total_hours_net': '00:00',
                    'analysis': {}
                }
            
            # Process time entry
            time_in = str(row[columns['time_in']])
            time_out = str(row[columns['time_out']])
            date = str(row[columns['date']])
            
            # Apply note overrides FIRST to get the final times to display
            final_time_in, final_time_out, override_break, override_notes = self.apply_note_overrides(
                employee_name, date, time_in, time_out, -1
            )
            
            net_hours_rounded, gross_hours, break_minutes, notes = self.calculate_net_hours(
                time_in, time_out, employee_name, date
            )
            
            # Combine override notes with calculation notes
            combined_notes = notes
            if override_notes and notes != override_notes:
                combined_notes = override_notes if not notes else f"{override_notes}; {notes}"
            
            employees[employee_name]['time_entries'].append({
                'date': date,
                'time_in': final_time_in,  # Use overridden time for display
                'time_out': final_time_out,  # Use overridden time for display
                'gross_hours': gross_hours,
                'break_minutes': break_minutes,
                'net_hours_rounded': net_hours_rounded,
                'notes': combined_notes,
                'is_complete': final_time_in != 'Missed' and final_time_out != 'Missed'
            })
        
        # Calculate totals and analyze
        employee_list = []
        for employee in employees.values():
            # Calculate total net hours
            total_net_minutes = sum(
                self.time_to_minutes(entry['net_hours_rounded']) 
                for entry in employee['time_entries']
            )
            employee['total_hours_net'] = self.minutes_to_time(total_net_minutes)
            
            # Calculate total gross hours
            total_gross_minutes = sum(
                self.time_to_minutes(entry['gross_hours']) 
                for entry in employee['time_entries']
            )
            employee['total_hours_gross'] = self.minutes_to_time(total_gross_minutes)
            
            # Analyze
            employee['analysis'] = self.analyze_employee_data(employee)
            
            employee_list.append(employee)
        
        return employee_list
    
    def process_sheet_per_employee_excel(self, excel_path: Path, excel_config: Dict) -> List[Dict[str, Any]]:
        """Process Excel with one sheet per employee."""
        xlsx = pd.ExcelFile(excel_path)
        employees = []
        
        # Get column mappings
        columns = excel_config.get('columns', {
            'date': 'Date',
            'time_in': 'Time In',
            'time_out': 'Time Out'
        })
        
        for sheet_name in xlsx.sheet_names:
            # Skip summary sheets
            if sheet_name.lower() in ['summary', 'total', 'overview']:
                continue
            
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            
            employee = {
                'name': sheet_name,
                'time_entries': [],
                'total_hours_net': '00:00',
                'analysis': {}
            }
            
            # Process each row
            for _, row in df.iterrows():
                time_in = str(row.get(columns['time_in'], ''))
                time_out = str(row.get(columns['time_out'], ''))
                date = str(row.get(columns['date'], ''))
                
                if pd.isna(time_in) or pd.isna(time_out):
                    continue
                
                # Apply note overrides FIRST to get the final times to display
                final_time_in, final_time_out, override_break, override_notes = self.apply_note_overrides(
                    sheet_name, date, time_in, time_out, -1
                )
                
                net_hours_rounded, gross_hours, break_minutes, notes = self.calculate_net_hours(
                    time_in, time_out, sheet_name, date
                )
                
                # Combine override notes with calculation notes
                combined_notes = notes
                if override_notes and notes != override_notes:
                    combined_notes = override_notes if not notes else f"{override_notes}; {notes}"
                
                employee['time_entries'].append({
                    'date': date,
                    'time_in': final_time_in,  # Use overridden time for display
                    'time_out': final_time_out,  # Use overridden time for display
                    'gross_hours': gross_hours,
                    'break_minutes': break_minutes,
                    'net_hours_rounded': net_hours_rounded,
                    'notes': combined_notes,
                    'is_complete': final_time_in != 'Missed' and final_time_out != 'Missed'
                })
            
            # Calculate totals
            total_net_minutes = sum(
                self.time_to_minutes(entry['net_hours_rounded']) 
                for entry in employee['time_entries']
            )
            employee['total_hours_net'] = self.minutes_to_time(total_net_minutes)
            
            # Calculate total gross hours
            total_gross_minutes = sum(
                self.time_to_minutes(entry['gross_hours']) 
                for entry in employee['time_entries']
            )
            employee['total_hours_gross'] = self.minutes_to_time(total_gross_minutes)
            
            # Analyze
            employee['analysis'] = self.analyze_employee_data(employee)
            
            employees.append(employee)
        
        return employees
    
    def analyze_employee_data(self, employee: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze employee time data."""
        analysis = {}
        
        # Basic statistics
        total_days = len(employee['time_entries'])
        worked_days = len([e for e in employee['time_entries'] if e['net_hours_rounded'] != '00:00'])
        missed_days = total_days - worked_days
        missing_punches = len([e for e in employee['time_entries'] 
                        if not e['is_complete']])
        
        # Time calculations
        total_net_minutes = self.time_to_minutes(employee['total_hours_net'])
        total_break_minutes = sum(e.get('break_minutes', 0) for e in employee['time_entries'])
        
        avg_daily_hours = total_net_minutes / worked_days / 60 if worked_days > 0 else 0
        
        # Overtime detection
        overtime_days = [e for e in employee['time_entries'] 
                        if self.time_to_minutes(e['net_hours_rounded']) > 8 * 60]
        
        analysis = {
            'total_days_scheduled': total_days,
            'days_worked': worked_days,
            'days_missed': missed_days,
            'missing_punches': missing_punches,
            'attendance_rate': (worked_days / total_days * 100) if total_days > 0 else 0,
            'total_hours_net': employee['total_hours_net'],
            'total_break_hours': self.minutes_to_time(total_break_minutes),
            'average_daily_hours_net': avg_daily_hours,
            'overtime_days': len(overtime_days)
        }
        
        return analysis
    
    def generate_all_outputs(self, employees: List[Dict[str, Any]]):
        """Generate simplified output - just one summary file."""
        # Create just one summary report instead of individual folders
        self.create_summary_report(employees)
        
        self.logger.info(f"Generated summary for {len(employees)} employees")
    
    def create_summary_report(self, employees: List[Dict[str, Any]]):
        """Create comprehensive summary with all employee data."""
        # Extract pay period from the first employee for filename
        pay_period = ""
        if employees and employees[0].get('pay_period'):
            pay_period_raw = employees[0]['pay_period']
            # Convert "5/12/2025 - 5/23/2025" to "May_12-23_2025"
            if ' - ' in pay_period_raw:
                try:
                    start_date, end_date = pay_period_raw.split(' - ')
                    # Parse the dates
                    start = datetime.strptime(start_date.strip(), "%m/%d/%Y")
                    end = datetime.strptime(end_date.strip(), "%m/%d/%Y")
                    
                    # Format for filename: "May_12-23_2025"
                    start_month = start.strftime("%b")  # "May"
                    start_day = str(start.day)  # "12" (remove leading zero if needed)
                    end_day = str(end.day)  # "23" 
                    year = start.strftime("%Y")  # "2025"
                    
                    pay_period = f"{start_month}_{start_day}-{end_day}_{year}"
                except:
                    # Fallback to raw pay period with safe characters
                    pay_period = pay_period_raw.replace('/', '_').replace(' - ', '_to_')
        
        # Create base filename with pay period and input file identifier
        if pay_period:
            base_filename = f"{self.company_name}_timecard_summary_{pay_period}"
        else:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_filename = f"{self.company_name}_timecard_summary_{timestamp}"
        
        # Add input filename for additional uniqueness (if available)
        if hasattr(self, 'current_input_file') and self.current_input_file:
            # Only add if it's not already part of the filename
            input_identifier = self.current_input_file.replace(' ', '_').replace('(', '').replace(')', '')
            if input_identifier not in base_filename:
                base_filename = f"{base_filename}_{input_identifier}"
        
        # Check if file already exists, add counter to make it unique
        summary_path = self.output_base_dir / f"{base_filename}.xlsx"
        counter = 1
        while summary_path.exists():
            summary_path = self.output_base_dir / f"{base_filename}_{counter}.xlsx"
            counter += 1
        
        # Sort employees by first name and filter out those with 0 total net hours
        def get_first_name(employee_name: str) -> str:
            """Extract first name from 'Last, First' format."""
            if ', ' in employee_name:
                # Format: "Last, First" -> return "First"
                return employee_name.split(', ')[1].strip()
            else:
                # Format: "First Last" -> return "First"
                return employee_name.split()[0].strip()
        
        def format_name_first_last(employee_name: str) -> str:
            """Convert 'Last, First' to 'First Last' format."""
            if ', ' in employee_name:
                # Format: "Last, First" -> "First Last"
                last, first = employee_name.split(', ', 1)
                return f"{first.strip()} {last.strip()}"
            else:
                # Already in "First Last" format
                return employee_name

        # Filter out employees with 0 total net hours
        employees_with_hours = []
        for emp in employees:
            total_net_minutes = self.time_to_minutes(emp['analysis']['total_hours_net'])
            if total_net_minutes > 0:
                employees_with_hours.append(emp)
        
        employees_sorted = sorted(employees_with_hours, key=lambda emp: get_first_name(emp['name']).lower())
        
        # Create summary data
        summary_data = []
        detailed_data = []
        
        for emp in employees_sorted:
            analysis = emp['analysis']
            
            # Format name as First Last for display
            formatted_name = format_name_first_last(emp['name'])
            
            # Summary row for each employee
            summary_data.append({
                'Employee': formatted_name,
                'Total_Net_Hours': analysis['total_hours_net'],
                'Total_Break_Hours': analysis['total_break_hours'],
                'Pay_Period': emp.get('pay_period', '')
            })
            
            # Detailed daily entries for each employee
            for entry in emp['time_entries']:
                detailed_data.append({
                    'Employee': formatted_name,
                    'Date': entry['date'],
                    'Day': entry.get('day_of_week', ''),
                    'Time_In': entry['time_in'],
                    'Time_Out': entry['time_out'],
                    'Gross_Hours': entry['gross_hours'],
                    'Break_Minutes': entry['break_minutes'],
                    'Net_Hours': entry['net_hours_rounded'],
                    'Notes': entry.get('notes', '')
                })
        
        # Create Excel file with multiple sheets
        with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
            # Summary sheet
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Employee_Summary', index=False, startrow=2)
            
            # Add profile name header to summary sheet
            workbook = writer.book
            summary_worksheet = writer.sheets['Employee_Summary']
            summary_worksheet['A1'] = self.company_name
            
            # Detailed daily entries sheet
            detailed_df = pd.DataFrame(detailed_data)
            detailed_df.to_excel(writer, sheet_name='Daily_Details', index=False, startrow=2)
            
            # Add profile name header to detailed sheet
            details_worksheet = writer.sheets['Daily_Details']
            details_worksheet['A1'] = self.company_name
        
        self.logger.info(f"Created comprehensive summary: {summary_path}")
    
    def create_employee_files(self, employee: Dict[str, Any]):
        """This method is no longer used - keeping for backward compatibility."""
        pass
    
    def create_employee_csv(self, employee: Dict[str, Any], output_dir: Path):
        """This method is no longer used - keeping for backward compatibility."""
        pass
    
    def create_employee_report(self, employee: Dict[str, Any], output_dir: Path):
        """This method is no longer used - keeping for backward compatibility."""
        pass

    def create_employees_from_notes(self, existing_employees: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Create employee entries for those who only appear in notes but not in timecard data."""
        note_only_employees = []
        existing_names = {emp['name'] for emp in existing_employees}
        
        # Find employees mentioned in notes but not in timecard data
        note_employees = set()
        for note in self.notes:
            emp_name = note.get('employee', '')
            if emp_name and emp_name not in existing_names:
                note_employees.add(emp_name)
        
        for employee_name in note_employees:
            self.logger.info(f"Creating employee entry from notes: {employee_name}")
            
            # Create basic employee structure
            employee = {
                'name': employee_name,
                'employee_group': '',
                'pay_period': '',
                'approval_status': '',
                'time_entries': [],
                'total_hours_gross': '00:00',
                'total_hours_net': '00:00',
                'analysis': {}
            }
            
            # Find all notes for this employee and create time entries
            employee_notes = [note for note in self.notes if note.get('employee') == employee_name]
            
            for note in employee_notes:
                if note.get('type') == 'time_override' and 'date' in note:
                    time_entry = self.create_time_entry_from_note(note)
                    if time_entry:
                        employee['time_entries'].append(time_entry)
                        self.logger.info(f"Created time entry from note for {employee_name} on {note['date']}: {time_entry['time_in']} - {time_entry['time_out']} = {time_entry['net_hours_rounded']}")
            
            if employee['time_entries']:
                # Calculate totals
                total_net_minutes = sum(
                    self.time_to_minutes(entry['net_hours_rounded']) 
                    for entry in employee['time_entries']
                )
                employee['total_hours_net'] = self.minutes_to_time(total_net_minutes)
                
                # Perform analysis
                employee['analysis'] = self.analyze_employee_data(employee)
                
                note_only_employees.append(employee)
        
        return note_only_employees

    def create_time_entry_from_note(self, note: Dict[str, Any]) -> Dict[str, Any]:
        """Create a time entry from a note override."""
        try:
            employee_name = note.get('employee', '')
            date = note.get('date', '')
            time_in = note.get('time_in', '00:00 AM')
            time_out = note.get('time_out', '00:00 PM')
            
            # Calculate net hours with breaks and rounding
            net_hours_rounded, gross_hours, break_minutes, note_text = self.calculate_net_hours(
                time_in, time_out, employee_name, date
            )
            
            # Parse day of week from date
            day_of_week = ''
            date_only = date
            if ',' in date:
                day_of_week = date.split(',')[0].strip()
                date_only = date.split(',')[1].strip()
            
            entry = {
                'date': date,
                'day_of_week': day_of_week,
                'date_only': date_only,
                'time_in': time_in,
                'time_out': time_out,
                'gross_hours': gross_hours,
                'break_minutes': break_minutes,
                'net_hours_rounded': net_hours_rounded,
                'department': '',
                'daily_hours_from_pdf': net_hours_rounded,
                'notes': note.get('note', ''),
                'is_complete': True
            }
            
            return entry
            
        except Exception as e:
            self.logger.error(f"Error creating time entry from note: {e}")
            return None

    def add_missing_days_from_notes(self, employees: List[Dict[str, Any]]):
        """Add missing days from notes to existing employees."""
        self.logger.info(f"=== Starting add_missing_days_from_notes with {len(employees)} employees ===")
        
        for employee in employees:
            employee_name = employee['name']
            existing_dates = {entry['date'] for entry in employee['time_entries']}
            
            self.logger.info(f"Checking {employee_name}: existing dates = {existing_dates}")
            
            # Find notes for this employee that have missing_punch_override type
            employee_notes = [note for note in self.notes if note.get('employee') == employee_name and note.get('type') == 'missing_punch_override' and 'date' in note]
            
            self.logger.info(f"Found {len(employee_notes)} missing_punch_override notes for {employee_name}")
            
            for note in employee_notes:
                note_date = note['date']
                self.logger.info(f"Processing missing punch note for {employee_name} on {note_date}")
                
                # Check if this date is missing from the employee's time entries
                if note_date not in existing_dates:
                    self.logger.info(f"Adding missing punch for {employee_name}: {note_date}")
                    
                    time_entry = self.create_time_entry_from_note(note)
                    if time_entry:
                        employee['time_entries'].append(time_entry)
                        self.logger.info(f"Created missing punch entry for {employee_name} on {note_date}: {time_entry['time_in']} - {time_entry['time_out']} = {time_entry['net_hours_rounded']}")
                        
                        # Recalculate totals for this employee
                        total_net_minutes = sum(
                            self.time_to_minutes(entry['net_hours_rounded']) 
                            for entry in employee['time_entries']
                        )
                        employee['total_hours_net'] = self.minutes_to_time(total_net_minutes)
                        
                        # Recalculate analysis
                        employee['analysis'] = self.analyze_employee_data(employee)
                else:
                    self.logger.info(f"Date {note_date} already exists for {employee_name}, skipping missing punch note")

def main():
    parser = argparse.ArgumentParser(description='Process time cards with company-specific rules')
    parser.add_argument('config_file', help='Company configuration file (e.g., company_a_config.json)')
    parser.add_argument('input_file', help='Time card file to process (PDF or Excel)')
    parser.add_argument('--notes', '-n', help='Weekly notes/exceptions file (optional)', default=None)
    
    args = parser.parse_args()
    
    # Create processor with company config and optional notes
    processor = UniversalTimeCardProcessor(args.config_file, args.notes)
    
    # Process the file
    processor.process_file(args.input_file)

if __name__ == "__main__":
    main()