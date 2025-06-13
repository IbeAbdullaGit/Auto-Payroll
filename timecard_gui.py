#!/usr/bin/env python3
"""
GUI Application for Time Card Processing
Provides user-friendly interface for managing profiles, notes, and processing time cards
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
from pathlib import Path
from datetime import datetime
import subprocess
import sys
import re

# Import PDF processing capabilities
try:
    import pdfplumber
    PDF_LIBRARY = 'pdfplumber'
except ImportError:
    try:
        import PyPDF2
        PDF_LIBRARY = 'PyPDF2'
    except ImportError:
        PDF_LIBRARY = None

class TimeCardGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Processor")
        self.root.geometry("900x700")
        self.root.configure(bg='#2c3e50')
        
        # Configure ttk styles to reduce highlighting
        self.style = ttk.Style()
        # Configure a custom style for readonly comboboxes with reduced highlighting
        self.style.map('Custom.TCombobox', 
                      selectbackground=[('readonly', '')],
                      selectforeground=[('readonly', 'black')])
        self.style.configure('Custom.TCombobox', arrowcolor='black')
        
        # Initialize data
        self.current_profile = {}
        self.current_notes = {"pay_period": "", "notes": []}
        self.config_dir = Path("timecard_processing/configs")
        self.notes_dir = Path("timecard_processing/note_configs")
        self.config_dir.mkdir(exist_ok=True)
        self.notes_dir.mkdir(exist_ok=True)
        
        self.create_main_interface()
        self.load_existing_profiles()
    
    def create_main_interface(self):
        # Header
        header_frame = tk.Frame(self.root, bg='#34495e', height=80)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="PDF to Excel Processor", 
                              font=('Arial', 20, 'bold'), fg='white', bg='#34495e')
        title_label.pack(pady=15)
        
        subtitle_label = tk.Label(header_frame, text="Transform your PDF documents into structured Excel spreadsheets",
                                 font=('Arial', 10), fg='#bdc3c7', bg='#34495e')
        subtitle_label.pack()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_profiles_tab()
        self.create_configuration_tab()
        self.create_notes_tab()
        self.create_processing_tab()
    
    def create_profiles_tab(self):
        # Profiles tab
        profiles_frame = ttk.Frame(self.notebook)
        self.notebook.add(profiles_frame, text='📁 Profiles')
        
        # Profile Management section
        profile_mgmt_frame = ttk.LabelFrame(profiles_frame, text="⚙️ Profile Management", padding=10)
        profile_mgmt_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(profile_mgmt_frame, text="Current Profile:").grid(row=0, column=0, sticky='w', padx=5)
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(profile_mgmt_frame, textvariable=self.profile_var, 
                                   values=[], width=25, state='readonly', style='Custom.TCombobox')
        self.profile_combo.grid(row=0, column=1, padx=5)
        
        load_btn = ttk.Button(profile_mgmt_frame, text="Load", command=self.load_profile)
        load_btn.grid(row=0, column=2, padx=5)
        
        edit_btn = ttk.Button(profile_mgmt_frame, text="Edit Name", command=self.edit_profile_name)
        edit_btn.grid(row=0, column=3, padx=5)
        
        delete_btn = ttk.Button(profile_mgmt_frame, text="Delete", command=self.delete_profile)
        delete_btn.grid(row=0, column=4, padx=5)
        
        # Create New Profile section
        new_profile_frame = ttk.LabelFrame(profiles_frame, text="⚙️ Create New Profile", padding=10)
        new_profile_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(new_profile_frame, text="Profile Name:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.new_profile_name = tk.StringVar()
        ttk.Entry(new_profile_frame, textvariable=self.new_profile_name, width=30).grid(row=0, column=1, padx=5, pady=2)
        
        create_btn = ttk.Button(new_profile_frame, text="Create Profile", command=self.create_new_profile)
        create_btn.grid(row=0, column=2, padx=5, pady=2)
        
        # PDF Analysis section
        ttk.Label(new_profile_frame, text="Or analyze PDF to extract employees:").grid(row=1, column=0, columnspan=3, sticky='w', padx=5, pady=(10,2))
        
        pdf_frame = ttk.Frame(new_profile_frame)
        pdf_frame.grid(row=2, column=0, columnspan=3, sticky='ew', padx=5, pady=2)
        
        self.selected_pdf_label = ttk.Label(pdf_frame, text="No PDF selected", foreground='gray')
        self.selected_pdf_label.grid(row=0, column=0, sticky='w', padx=5)
        
        browse_pdf_btn = ttk.Button(pdf_frame, text="Browse PDF", command=self.browse_pdf_for_profile)
        browse_pdf_btn.grid(row=0, column=1, padx=5)
        
        analyze_btn = ttk.Button(pdf_frame, text="Analyze & Create Profile", command=self.analyze_pdf_and_create_profile)
        analyze_btn.grid(row=0, column=2, padx=5)
        
        update_btn = ttk.Button(pdf_frame, text="Update Current Profile", command=self.update_profile_from_pdf)
        update_btn.grid(row=0, column=3, padx=5)
        
        # Store selected PDF path
        self.selected_pdf_path = None
    
    def create_configuration_tab(self):
        # Employees tab (renamed from Configuration)
        config_frame = ttk.Frame(self.notebook)
        self.notebook.add(config_frame, text='👥 Employees')
        
        # Create scrollable frame
        canvas = tk.Canvas(config_frame)
        scrollbar = ttk.Scrollbar(config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Basic Settings
        basic_frame = ttk.LabelFrame(scrollable_frame, text="Basic Settings", padding=10)
        basic_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(basic_frame, text="Company Name:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.company_name = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.company_name, width=30).grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(basic_frame, text="Daily Hour Cap:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.daily_cap = tk.StringVar(value="8.0")
        ttk.Entry(basic_frame, textvariable=self.daily_cap, width=10).grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        # Employee Management
        emp_frame = ttk.LabelFrame(scrollable_frame, text="Employee Settings", padding=10)
        emp_frame.pack(fill='x', padx=10, pady=5)
        
        # Search functionality
        search_frame = ttk.Frame(emp_frame)
        search_frame.grid(row=0, column=0, columnspan=4, sticky='ew', padx=5, pady=2)
        
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, padx=5)
        self.employee_search = tk.StringVar()
        self.employee_search.trace('w', self.filter_employees)
        search_entry = ttk.Entry(search_frame, textvariable=self.employee_search, width=30)
        search_entry.grid(row=0, column=1, padx=5)
        
        # Employee list
        ttk.Label(emp_frame, text="Employees:").grid(row=1, column=0, sticky='nw', padx=5, pady=2)
        
        emp_list_frame = ttk.Frame(emp_frame)
        emp_list_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        
        self.employee_listbox = tk.Listbox(emp_list_frame, height=6)
        self.employee_listbox.pack(side='left', fill='both', expand=True)
        
        emp_scroll = ttk.Scrollbar(emp_list_frame, orient="vertical", command=self.employee_listbox.yview)
        emp_scroll.pack(side='right', fill='y')
        self.employee_listbox.configure(yscrollcommand=emp_scroll.set)
        
        # Employee controls
        emp_controls = ttk.Frame(emp_frame)
        emp_controls.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        ttk.Label(emp_controls, text="Employee Name:").grid(row=0, column=0, padx=5)
        self.new_employee_name = tk.StringVar()
        ttk.Entry(emp_controls, textvariable=self.new_employee_name, width=25).grid(row=0, column=1, padx=5)
        
        ttk.Button(emp_controls, text="Add Employee", command=self.add_employee).grid(row=0, column=2, padx=5)
        ttk.Button(emp_controls, text="Remove Selected", command=self.remove_employee).grid(row=0, column=3, padx=5)
        
        # Advanced Employee Settings
        settings_frame = ttk.LabelFrame(emp_controls, text="Settings for Selected Employee", padding=5)
        settings_frame.grid(row=1, column=0, columnspan=4, sticky='ew', pady=10)
        
        # Basic break settings
        basic_frame = ttk.Frame(settings_frame)
        basic_frame.grid(row=0, column=0, columnspan=6, sticky='ew', pady=2)
        
        ttk.Label(basic_frame, text="Break Minutes:").grid(row=0, column=0, padx=5)
        self.break_minutes = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.break_minutes, width=10).grid(row=0, column=1, padx=5)
        
        ttk.Label(basic_frame, text="Hour Threshold:").grid(row=0, column=2, padx=5)
        self.hour_threshold = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.hour_threshold, width=10).grid(row=0, column=3, padx=5)
        
        # Fixed hours setting
        fixed_frame = ttk.Frame(basic_frame)
        fixed_frame.grid(row=1, column=0, columnspan=6, sticky='ew', pady=2)
        
        self.use_fixed_hours = tk.BooleanVar()
        ttk.Checkbutton(fixed_frame, text="Use Fixed Hours:", variable=self.use_fixed_hours).grid(row=0, column=0, padx=5)
        self.fixed_hours = tk.StringVar()
        ttk.Entry(fixed_frame, textvariable=self.fixed_hours, width=10).grid(row=0, column=1, padx=5)
        
        # Escalating breaks
        escalating_frame = ttk.LabelFrame(settings_frame, text="Escalating Breaks", padding=5)
        escalating_frame.grid(row=1, column=0, columnspan=6, sticky='ew', pady=5)
        
        ttk.Label(escalating_frame, text="Use escalating breaks:").grid(row=0, column=0, padx=5)
        self.use_escalating = tk.BooleanVar()
        ttk.Checkbutton(escalating_frame, variable=self.use_escalating, command=self.toggle_escalating_options).grid(row=0, column=1, padx=5)
        
        # Escalating break options (initially hidden)
        self.escalating_options = ttk.Frame(escalating_frame)
        self.escalating_options.grid(row=1, column=0, columnspan=6, sticky='ew', pady=5)
        
        ttk.Label(self.escalating_options, text="6+ hours:").grid(row=0, column=0, padx=5)
        self.break_6_hours = tk.StringVar(value="30")
        ttk.Entry(self.escalating_options, textvariable=self.break_6_hours, width=8).grid(row=0, column=1, padx=5)
        ttk.Label(self.escalating_options, text="min").grid(row=0, column=2, padx=2)
        
        ttk.Label(self.escalating_options, text="7+ hours:").grid(row=0, column=3, padx=5)
        self.break_7_hours = tk.StringVar(value="60")
        ttk.Entry(self.escalating_options, textvariable=self.break_7_hours, width=8).grid(row=0, column=4, padx=5)
        ttk.Label(self.escalating_options, text="min").grid(row=0, column=5, padx=2)
        
        # Hide escalating options initially
        self.toggle_escalating_options()
        
        ttk.Button(settings_frame, text="Apply Settings", command=self.apply_employee_settings).grid(row=2, column=0, columnspan=6, pady=10)
        
        # Note: Settings are automatically saved when "Apply Settings" is clicked
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind employee selection (double-click to select and maintain selection)
        self.employee_listbox.bind('<Double-Button-1>', self.on_employee_select)
        self.employee_listbox.bind('<FocusOut>', self.prevent_deselection)
        
        # Track selected employee to maintain selection
        self.selected_employee = None
    
    def create_notes_tab(self):
        # Notes tab
        notes_frame = ttk.Frame(self.notebook)
        self.notebook.add(notes_frame, text='📝 Notes')
        
        # Pay period with period type selection
        period_frame = ttk.LabelFrame(notes_frame, text="Pay Period", padding=10)
        period_frame.pack(fill='x', padx=10, pady=5)
        
        # Period type selection
        ttk.Label(period_frame, text="Period Type:").grid(row=0, column=0, padx=5)
        self.period_type = tk.StringVar(value="2 weeks")
        period_combo = ttk.Combobox(period_frame, textvariable=self.period_type, values=["2 weeks", "1 month"], width=10, state='readonly')
        period_combo.grid(row=0, column=1, padx=5)
        period_combo.bind('<<ComboboxSelected>>', self.update_pay_period_dates)
        
        ttk.Label(period_frame, text="Pay Period:").grid(row=0, column=2, sticky='w', padx=5)
        self.pay_period = tk.StringVar()
        self.pay_period_combo = ttk.Combobox(period_frame, textvariable=self.pay_period, width=25, state='readonly')
        self.pay_period_combo.grid(row=0, column=3, padx=5)
        self.pay_period_combo.bind('<<ComboboxSelected>>', lambda e: self.update_available_dates_for_notes())
        
        # Initialize pay period options
        self.update_pay_period_dates()
        
        # Add new note
        add_note_frame = ttk.LabelFrame(notes_frame, text="Add New Note", padding=10)
        add_note_frame.pack(fill='x', padx=10, pady=5)
        
        # Configure grid weights for consistent alignment
        add_note_frame.columnconfigure(1, weight=1)
        
        # Employee selection
        ttk.Label(add_note_frame, text="Employee:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.note_employee = tk.StringVar()
        self.employee_combo = ttk.Combobox(add_note_frame, textvariable=self.note_employee, 
                                         width=30, style='Custom.TCombobox')
        self.employee_combo.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        # Add search functionality to employee combobox
        self.note_employee.trace('w', self.filter_employee_dropdown)
        self.employee_combo.bind('<KeyRelease>', self.on_employee_keyrelease)
        
        # Date selection
        ttk.Label(add_note_frame, text="Date:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.note_date = tk.StringVar()
        self.date_combo = ttk.Combobox(add_note_frame, textvariable=self.note_date, 
                                     width=20, state='readonly', style='Custom.TCombobox')
        self.date_combo.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        # Now that date_combo is created, update the available dates
        self.update_available_dates_for_notes()
        
        # Override type
        ttk.Label(add_note_frame, text="Override Type:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.note_type = tk.StringVar()
        type_combo = ttk.Combobox(add_note_frame, textvariable=self.note_type, width=20, 
                                state='readonly', style='Custom.TCombobox')
        type_combo['values'] = ('Time Change', 'Break Change', 'Missing Punch', 'Sick Day')
        type_combo.grid(row=2, column=1, sticky='w', padx=5, pady=2)
        type_combo.bind('<<ComboboxSelected>>', self.on_note_type_change)
        
        # Time fields (organized vertically for better alignment)
        self.time_frame = ttk.Frame(add_note_frame)
        self.time_frame.grid(row=3, column=0, columnspan=2, sticky='ew', pady=5)
        
        # Time In
        ttk.Label(self.time_frame, text="Time In:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        time_in_frame = ttk.Frame(self.time_frame)
        time_in_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        self.note_time_in_time = tk.StringVar()
        ttk.Entry(time_in_frame, textvariable=self.note_time_in_time, width=8).pack(side='left', padx=(0,2))
        
        self.note_time_in_ampm = tk.StringVar(value="AM")
        time_in_ampm = ttk.Combobox(time_in_frame, textvariable=self.note_time_in_ampm, width=4, state='readonly')
        time_in_ampm['values'] = ['AM', 'PM']
        time_in_ampm.pack(side='left')
        
        # Time Out
        ttk.Label(self.time_frame, text="Time Out:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        time_out_frame = ttk.Frame(self.time_frame)
        time_out_frame.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        self.note_time_out_time = tk.StringVar()
        ttk.Entry(time_out_frame, textvariable=self.note_time_out_time, width=8).pack(side='left', padx=(0,2))
        
        self.note_time_out_ampm = tk.StringVar(value="PM")
        time_out_ampm = ttk.Combobox(time_out_frame, textvariable=self.note_time_out_ampm, width=4, state='readonly')
        time_out_ampm['values'] = ['AM', 'PM']
        time_out_ampm.pack(side='left')
        
        # Break Minutes
        ttk.Label(self.time_frame, text="Break Minutes:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.note_break_minutes = tk.StringVar()
        ttk.Entry(self.time_frame, textvariable=self.note_break_minutes, width=8).grid(row=2, column=1, sticky='w', padx=5, pady=2)
        
        # Note text
        ttk.Label(add_note_frame, text="Note:").grid(row=4, column=0, sticky='nw', padx=5, pady=2)
        self.note_text = tk.StringVar()
        ttk.Entry(add_note_frame, textvariable=self.note_text, width=50).grid(row=4, column=1, sticky='ew', padx=5, pady=2)
        
        ttk.Button(add_note_frame, text="Add Note", command=self.add_note).grid(row=5, column=1, sticky='w', padx=5, pady=10)
        
        # Notes list
        notes_list_frame = ttk.LabelFrame(notes_frame, text="Current Notes", padding=10)
        notes_list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create treeview for notes
        columns = ('Employee', 'Date', 'Type', 'Details', 'Note')
        self.notes_tree = ttk.Treeview(notes_list_frame, columns=columns, show='headings', height=10)
        
        for col in columns:
            self.notes_tree.heading(col, text=col)
            self.notes_tree.column(col, width=120)
        
        notes_scroll = ttk.Scrollbar(notes_list_frame, orient="vertical", command=self.notes_tree.yview)
        self.notes_tree.configure(yscrollcommand=notes_scroll.set)
        
        self.notes_tree.pack(side="left", fill="both", expand=True)
        notes_scroll.pack(side="right", fill="y")
        
        # Notes controls
        notes_controls = ttk.Frame(notes_list_frame)
        notes_controls.pack(fill='x', pady=5)
        
        ttk.Button(notes_controls, text="Remove Selected", command=self.remove_note).pack(side='left', padx=5)
        ttk.Button(notes_controls, text="Save Notes", command=self.save_notes).pack(side='left', padx=5)
        ttk.Button(notes_controls, text="Load Notes", command=self.load_notes).pack(side='left', padx=5)
    
    def create_processing_tab(self):
        # Processing tab
        processing_frame = ttk.Frame(self.notebook)
        self.notebook.add(processing_frame, text='📊 Processing')
        
        # PDF Files section
        pdf_frame = ttk.LabelFrame(processing_frame, text="📄 PDF Files", padding=20)
        pdf_frame.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        
        # Upload area
        upload_frame = ttk.Frame(pdf_frame)
        upload_frame.pack(fill='x', pady=20)
        
        upload_label = ttk.Label(upload_frame, text="Click to upload PDF files\nor drag and drop files here", 
                                font=('Arial', 10), anchor='center')
        upload_label.pack(pady=20)
        
        # File list
        self.file_listbox = tk.Listbox(pdf_frame, height=8)
        self.file_listbox.pack(fill='both', expand=True, pady=10)
        
        # Control buttons
        btn_frame = ttk.Frame(pdf_frame)
        btn_frame.pack(fill='x', pady=10)
        
        ttk.Button(btn_frame, text="Add Files", command=self.add_pdf_files).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Remove Selected", command=self.remove_pdf_file).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Clear All", command=self.clear_pdf_files).pack(side='left', padx=5)
        
        # Processing Control section
        control_frame = ttk.LabelFrame(processing_frame, text="🔧 Processing Control", padding=20)
        control_frame.pack(side='right', fill='both', expand=True, padx=10, pady=10)
        
        # Output location
        ttk.Label(control_frame, text="Output Location:").pack(anchor='w', pady=5)
        output_frame = ttk.Frame(control_frame)
        output_frame.pack(fill='x', pady=5)
        
        self.output_path = tk.StringVar(value="/Users/john/Documents/Output")
        ttk.Entry(output_frame, textvariable=self.output_path, width=30).pack(side='left', fill='x', expand=True)
        ttk.Button(output_frame, text="Browse", command=self.browse_output).pack(side='right', padx=(5,0))
        
        # Processing progress
        ttk.Label(control_frame, text="Processing Progress:").pack(anchor='w', pady=(20,5))
        self.progress_var = tk.StringVar(value="Ready to process")
        ttk.Label(control_frame, textvariable=self.progress_var).pack(anchor='w')
        
        self.progress_bar = ttk.Progressbar(control_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)
        
        # Process button
        process_btn = ttk.Button(control_frame, text="🚀 Process PDFs to Excel", 
                               command=self.process_files)
        process_btn.pack(pady=20, fill='x')
        
        # Bind upload area click
        upload_label.bind("<Button-1>", lambda e: self.add_pdf_files())
    
    # Profile Management Methods
    def load_existing_profiles(self):
        """Load existing config files into profile dropdown"""
        profiles = []
        if self.config_dir.exists():
            for config_file in self.config_dir.glob("*.json"):
                profiles.append(config_file.stem)
        self.profile_combo['values'] = profiles
        
        # Also update employee dropdown in notes
        if hasattr(self, 'note_employee_combo'):
            self.update_employee_dropdown()
    
    def load_profile(self):
        """Load selected profile"""
        profile_name = self.profile_var.get()
        if not profile_name:
            messagebox.showwarning("Warning", "Please select a profile to load")
            return
        
        config_file = self.config_dir / f"{profile_name}.json"
        try:
            with open(config_file, 'r') as f:
                self.current_profile = json.load(f)
            
            # Update UI with loaded profile
            self.update_configuration_ui()
            self.update_employee_dropdown()
            messagebox.showinfo("Success", f"Profile '{profile_name}' loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load profile: {str(e)}")
    
    def create_new_profile(self):
        """Create a new profile with default settings"""
        profile_name = self.new_profile_name.get().strip()
        if not profile_name:
            messagebox.showwarning("Warning", "Please enter a profile name")
            return
        
        # Create default profile structure
        self.current_profile = {
            "company_name": profile_name,
            "file_type": "pdf",
            "output_directory": f"output_{profile_name.lower().replace(' ', '_')}",
            "daily_hour_cap": 8.0,
            "rounding_rules": {
                "interval_minutes": 15,
                "threshold_minutes": 8
            },
            "break_config": {
                "break_rules": {
                    "minimum_hours_for_break": 6.0,
                    "long_day_threshold": 9.0,
                    "full_day_break": 60,
                    "long_day_break": 60
                },
                "employee_breaks": {},
                "employee_hour_thresholds": {}
            },
            "excel_config": {
                "format": "standard",
                "columns": {
                    "employee_name": "Employee Name",
                    "date": "Date",
                    "time_in": "Time In",
                    "time_out": "Time Out"
                }
            }
        }
        
        # Save the new profile
        config_file = self.config_dir / f"{profile_name}.json"
        try:
            with open(config_file, 'w') as f:
                json.dump(self.current_profile, f, indent=2)
            
            # Update UI
            self.load_existing_profiles()
            self.profile_var.set(profile_name)
            self.update_configuration_ui()
            self.new_profile_name.set("")
            
            messagebox.showinfo("Success", f"New profile '{profile_name}' created successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create profile: {str(e)}")
    
    def delete_profile(self):
        """Delete selected profile"""
        profile_name = self.profile_var.get()
        if not profile_name:
            messagebox.showwarning("Warning", "Please select a profile to delete")
            return
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete profile '{profile_name}'?"):
            config_file = self.config_dir / f"{profile_name}.json"
            try:
                config_file.unlink()
                self.load_existing_profiles()
                self.profile_var.set("")
                self.current_profile = {}
                self.update_configuration_ui()
                messagebox.showinfo("Success", f"Profile '{profile_name}' deleted successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete profile: {str(e)}")
    
    def edit_profile_name(self):
        """Edit/rename the selected profile"""
        current_profile_name = self.profile_var.get()
        if not current_profile_name:
            messagebox.showwarning("Warning", "Please select a profile to rename")
            return
        
        # Create a dialog for new name
        dialog = tk.Toplevel(self.root)
        dialog.title("Rename Profile")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # Create dialog content
        ttk.Label(dialog, text=f"Rename profile '{current_profile_name}' to:").pack(pady=10)
        
        new_name_var = tk.StringVar(value=current_profile_name)
        name_entry = ttk.Entry(dialog, textvariable=new_name_var, width=30)
        name_entry.pack(pady=5)
        name_entry.select_range(0, tk.END)
        name_entry.focus()
        
        # Button frame
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        def save_rename():
            new_name = new_name_var.get().strip()
            if not new_name:
                messagebox.showwarning("Warning", "Please enter a valid name")
                return
            
            if new_name == current_profile_name:
                dialog.destroy()
                return
            
            # Check if new name already exists
            new_config_file = self.config_dir / f"{new_name}.json"
            if new_config_file.exists():
                messagebox.showerror("Error", f"Profile '{new_name}' already exists")
                return
            
            try:
                # Update the profile's company name
                if self.current_profile:
                    self.current_profile['company_name'] = new_name
                
                # Save with new name
                with open(new_config_file, 'w') as f:
                    json.dump(self.current_profile, f, indent=2)
                
                # Delete old file
                old_config_file = self.config_dir / f"{current_profile_name}.json"
                old_config_file.unlink()
                
                # Update UI
                self.load_existing_profiles()
                self.profile_var.set(new_name)
                self.update_configuration_ui()
                
                dialog.destroy()
                messagebox.showinfo("Success", f"Profile renamed to '{new_name}' successfully")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to rename profile: {str(e)}")
        
        def cancel_rename():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_rename).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=cancel_rename).pack(side='left', padx=5)
        
        # Bind Enter key to save
        dialog.bind('<Return>', lambda e: save_rename())
        dialog.bind('<Escape>', lambda e: cancel_rename())
    
    def browse_pdf_for_profile(self):
        """Browse for PDF file to analyze for employee names"""
        if PDF_LIBRARY is None:
            messagebox.showerror("Error", "PDF processing library not available. Please install pdfplumber or PyPDF2.")
            return
        
        file_path = filedialog.askopenfilename(
            title="Select PDF timecard for analysis",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file_path:
            self.selected_pdf_path = file_path
            filename = Path(file_path).name
            self.selected_pdf_label.config(text=f"Selected: {filename}", foreground='blue')
    
    def analyze_pdf_and_create_profile(self):
        """Analyze PDF and create profile with extracted employee names"""
        if not self.selected_pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
        
        profile_name = self.new_profile_name.get().strip()
        if not profile_name:
            messagebox.showwarning("Warning", "Please enter a profile name")
            return
        
        try:
            # Extract employee names from PDF
            employee_names = self.extract_employee_names_from_pdf(self.selected_pdf_path)
            
            if not employee_names:
                messagebox.showwarning("Warning", "No employee names found in the PDF")
                return
            
            # Create profile with extracted employees
            self.current_profile = {
                "company_name": profile_name,
                "file_type": "pdf",
                "output_directory": f"output_{profile_name.lower().replace(' ', '_')}",
                "daily_hour_cap": 8.0,
                "rounding_rules": {
                    "interval_minutes": 15,
                    "threshold_minutes": 8
                },
                "break_config": {
                    "break_rules": {
                        "minimum_hours_for_break": 6.0,
                        "long_day_threshold": 9.0,
                        "full_day_break": 60,
                        "long_day_break": 60
                    },
                    "employee_breaks": {},
                    "employee_hour_thresholds": {}
                },
                "excel_config": {
                    "format": "standard",
                    "columns": {
                        "employee_name": "Employee Name",
                        "date": "Date",
                        "time_in": "Time In",
                        "time_out": "Time Out"
                    }
                }
            }
            
            # Add all extracted employees with default 60-minute breaks
            for employee_name in employee_names:
                self.current_profile['break_config']['employee_breaks'][employee_name] = 60
            
            # Save the profile
            config_file = self.config_dir / f"{profile_name}.json"
            with open(config_file, 'w') as f:
                json.dump(self.current_profile, f, indent=2)
            
            # Update UI
            self.load_existing_profiles()
            self.profile_var.set(profile_name)
            self.update_configuration_ui()
            self.update_employee_dropdown()
            self.new_profile_name.set("")
            self.selected_pdf_path = None
            self.selected_pdf_label.config(text="No PDF selected", foreground='gray')
            
            messagebox.showinfo("Success", 
                f"Profile '{profile_name}' created successfully!\n\n"
                f"Found {len(employee_names)} employees:\n" + 
                "\n".join(f"• {name}" for name in employee_names[:10]) + 
                (f"\n... and {len(employee_names)-10} more" if len(employee_names) > 10 else ""))
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze PDF and create profile: {str(e)}")
    
    def update_profile_from_pdf(self):
        """Update existing profile with employees from PDF"""
        if not self.selected_pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
        
        profile_name = self.profile_var.get()
        if not profile_name or not self.current_profile:
            messagebox.showwarning("Warning", "Please load a profile first")
            return
        
        try:
            # Extract employee names from PDF
            employee_names = self.extract_employee_names_from_pdf(self.selected_pdf_path)
            
            if not employee_names:
                messagebox.showwarning("Warning", "No employee names found in the PDF")
                return
            
            # Get existing employees
            existing_employees = set(self.current_profile.get('break_config', {}).get('employee_breaks', {}).keys())
            new_employees = set(employee_names) - existing_employees
            
            # Add new employees with default 60-minute breaks
            for employee_name in new_employees:
                self.current_profile['break_config']['employee_breaks'][employee_name] = 60
            
            # Save the updated profile
            config_file = self.config_dir / f"{profile_name}.json"
            with open(config_file, 'w') as f:
                json.dump(self.current_profile, f, indent=2)
            
            # Update UI
            self.update_configuration_ui()
            self.update_employee_dropdown()
            self.selected_pdf_path = None
            self.selected_pdf_label.config(text="No PDF selected", foreground='gray')
            
            if new_employees:
                messagebox.showinfo("Success", 
                    f"Profile '{profile_name}' updated successfully!\n\n"
                    f"Added {len(new_employees)} new employees:\n" + 
                    "\n".join(f"• {name}" for name in sorted(new_employees)[:10]) + 
                    (f"\n... and {len(new_employees)-10} more" if len(new_employees) > 10 else ""))
            else:
                messagebox.showinfo("Info", "No new employees found. Profile unchanged.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update profile from PDF: {str(e)}")
    
    def extract_employee_names_from_pdf(self, pdf_path):
        """Extract employee names from PDF timecard"""
        try:
            text = self.extract_text_from_pdf(pdf_path)
            employee_names = self.parse_employee_names_from_text(text)
            return sorted(list(set(employee_names)))  # Remove duplicates and sort
        except Exception as e:
            raise Exception(f"Error extracting names from PDF: {str(e)}")
    
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF using available library"""
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
        elif PDF_LIBRARY == 'PyPDF2':
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
        
        return text
    
    def parse_employee_names_from_text(self, text):
        """Parse employee names from PDF text using the same logic as processor.py"""
        employee_names = []
        
        # Split by "TIME CARD REPORT" sections (each employee has their own section)
        sections = text.split('TIME CARD REPORT')
        
        for section in sections[1:]:  # Skip the first empty section
            if not section.strip():
                continue
            
            # Extract employee name using the same regex as processor.py
            name_match = re.search(r'Name\s*:\s*([^,\n]+(?:,\s*[^,\n]+)*)', section)
            if name_match:
                name = name_match.group(1).strip()
                # Remove "Approval Status" if it got included
                if "Approval Status" in name:
                    name = name.split("Approval Status")[0].strip()
                
                # Clean up common artifacts
                name = name.replace('\n', ' ').strip()
                
                # Only add if it looks like a valid name (has letters and maybe comma, parentheses, hyphens, etc.)
                if re.match(r'^[A-Za-z\s,.\-()_+=]+$', name) and len(name) > 2:
                    employee_names.append(name)
        
        return employee_names
    
    def update_configuration_ui(self):
        """Update configuration tab with current profile data"""
        if not self.current_profile:
            return
        
        # Update basic settings
        self.company_name.set(self.current_profile.get('company_name', ''))
        self.daily_cap.set(str(self.current_profile.get('daily_hour_cap', 8.0)))
        
        # Update employee list - get all employees from all sources
        self.employee_listbox.delete(0, tk.END)
        break_config = self.current_profile.get('break_config', {})
        
        # Collect all employee names from different sources
        all_employees = set()
        all_employees.update(break_config.get('employee_breaks', {}).keys())
        all_employees.update(break_config.get('employee_hour_thresholds', {}).keys())
        all_employees.update(break_config.get('escalating_breaks', {}).keys())
        all_employees.update(break_config.get('fixed_hours_employees', {}).keys())
        
        # Add all employees to the list
        for employee in sorted(all_employees):
            self.employee_listbox.insert(tk.END, employee)
    
    def update_employee_dropdown(self):
        """Update employee dropdown in notes tab"""
        if not self.current_profile:
            return
        
        break_config = self.current_profile.get('break_config', {})
        
        # Collect all employee names from different sources
        all_employees = set()
        all_employees.update(break_config.get('employee_breaks', {}).keys())
        all_employees.update(break_config.get('employee_hour_thresholds', {}).keys())
        all_employees.update(break_config.get('escalating_breaks', {}).keys())
        all_employees.update(break_config.get('fixed_hours_employees', {}).keys())
        
        # Store all employees for filtering
        self.all_employees = sorted(list(all_employees))
        
        # Update the employee dropdown
        if hasattr(self, 'employee_combo'):
            self.employee_combo['values'] = self.all_employees
    
    # Employee Management Methods
    def filter_employees(self, *args):
        """Filter employee list based on search term"""
        if not self.current_profile:
            return
        
        search_term = self.employee_search.get().lower()
        employee_breaks = self.current_profile.get('break_config', {}).get('employee_breaks', {})
        
        # Clear current list
        self.employee_listbox.delete(0, tk.END)
        
        # Add filtered employees
        for employee in employee_breaks.keys():
            if search_term in employee.lower():
                self.employee_listbox.insert(tk.END, employee)
    
    def toggle_escalating_options(self):
        """Show/hide escalating break options"""
        if self.use_escalating.get():
            # Show escalating options
            for widget in self.escalating_options.winfo_children():
                widget.grid()
        else:
            # Hide escalating options
            for widget in self.escalating_options.winfo_children():
                widget.grid_remove()
    
    def prevent_deselection(self, event):
        """Prevent employee deselection and restore if needed"""
        if hasattr(self, 'selected_employee') and self.selected_employee is not None:
            # Find the employee in the current list and reselect it
            for i in range(self.employee_listbox.size()):
                if self.employee_listbox.get(i) == self.selected_employee:
                    self.employee_listbox.selection_set(i)
                    break
    
    def add_employee(self):
        """Add new employee to current profile"""
        employee_name = self.new_employee_name.get().strip()
        if not employee_name:
            messagebox.showwarning("Warning", "Please enter an employee name")
            return
        
        if not self.current_profile:
            messagebox.showwarning("Warning", "Please load or create a profile first")
            return
        
        # Add to profile
        if 'break_config' not in self.current_profile:
            self.current_profile['break_config'] = {'employee_breaks': {}, 'employee_hour_thresholds': {}}
        
        self.current_profile['break_config']['employee_breaks'][employee_name] = 60  # Default 60 minutes
        
        # Update UI
        self.employee_listbox.insert(tk.END, employee_name)
        self.new_employee_name.set("")
        self.update_employee_dropdown()
        
        messagebox.showinfo("Success", f"Employee '{employee_name}' added successfully")
    
    def remove_employee(self):
        """Remove selected employee"""
        selection = self.employee_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an employee to remove")
            return
        
        employee_name = self.employee_listbox.get(selection[0])
        
        if messagebox.askyesno("Confirm Remove", f"Remove employee '{employee_name}'?"):
            # Remove from profile
            if 'break_config' in self.current_profile:
                self.current_profile['break_config']['employee_breaks'].pop(employee_name, None)
                self.current_profile['break_config']['employee_hour_thresholds'].pop(employee_name, None)
            
            # Update UI
            self.employee_listbox.delete(selection[0])
            self.update_employee_dropdown()
            
            messagebox.showinfo("Success", f"Employee '{employee_name}' removed successfully")
    
    def on_employee_select(self, event):
        """Handle employee selection - load their current settings"""
        selection = self.employee_listbox.curselection()
        if not selection:
            return
        
        employee_name = self.employee_listbox.get(selection[0])
        break_config = self.current_profile.get('break_config', {})
        
        # Load regular break minutes
        break_minutes = break_config.get('employee_breaks', {}).get(employee_name, 60)
        self.break_minutes.set(str(break_minutes))
        
        # Load hour threshold
        hour_threshold = break_config.get('employee_hour_thresholds', {}).get(employee_name, 6.0)
        self.hour_threshold.set(str(hour_threshold))
        
        # Check if this employee has fixed hours
        fixed_hours_employees = break_config.get('fixed_hours_employees', {})
        if employee_name in fixed_hours_employees:
            self.use_fixed_hours.set(True)
            self.fixed_hours.set(str(fixed_hours_employees[employee_name]))
        else:
            self.use_fixed_hours.set(False)
            self.fixed_hours.set("8.0")
        
        # Check if this employee has escalating breaks
        escalating_breaks = break_config.get('escalating_breaks', {})
        if employee_name in escalating_breaks:
            self.use_escalating.set(True)
            escalating_rule = escalating_breaks[employee_name]
            thresholds = escalating_rule.get('thresholds', [])
            
            # Find the two main thresholds (assume descending order)
            if len(thresholds) >= 2:
                # Higher threshold first (7+ hours)
                high_threshold = thresholds[0]
                low_threshold = thresholds[1]
                
                self.break_7_hours.set(str(high_threshold['break_minutes']))
                self.break_6_hours.set(str(low_threshold['break_minutes']))
        else:
            self.use_escalating.set(False)
            # Reset to defaults
            self.break_6_hours.set("30")
            self.break_7_hours.set("60")
        
        # Update UI visibility
        self.toggle_escalating_options()
    
    def apply_employee_settings(self):
        """Apply all employee settings"""
        selection = self.employee_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an employee")
            return
        
        employee_name = self.employee_listbox.get(selection[0])
        
        try:
            break_config = self.current_profile.setdefault('break_config', {})
            
            # Basic settings
            break_min = int(self.break_minutes.get()) if self.break_minutes.get() else 60
            hour_thresh = float(self.hour_threshold.get()) if self.hour_threshold.get() else None
            fixed_hrs = float(self.fixed_hours.get()) if self.fixed_hours.get() else None
            
            # Update basic settings
            break_config.setdefault('employee_breaks', {})[employee_name] = break_min
            
            if hour_thresh is not None:
                break_config.setdefault('employee_hour_thresholds', {})[employee_name] = hour_thresh
            else:
                break_config.get('employee_hour_thresholds', {}).pop(employee_name, None)
            
            if fixed_hrs is not None:
                break_config.setdefault('fixed_hours_employees', {})[employee_name] = fixed_hrs
            else:
                break_config.get('fixed_hours_employees', {}).pop(employee_name, None)
            
            # Handle escalating breaks
            if self.use_escalating.get():
                break_6 = int(self.break_6_hours.get()) if self.break_6_hours.get() else 30
                break_7 = int(self.break_7_hours.get()) if self.break_7_hours.get() else 60
                
                break_config.setdefault('escalating_breaks', {})[employee_name] = {
                    "thresholds": [
                        {"hours": 7.0, "break_minutes": break_7},
                        {"hours": 6.0, "break_minutes": break_6}
                    ]
                }
            else:
                # Remove escalating breaks if unchecked
                break_config.get('escalating_breaks', {}).pop(employee_name, None)
            
            # Save the profile automatically
            profile_name = self.profile_var.get()
            if profile_name:
                config_file = self.config_dir / f"{profile_name}.json"
                with open(config_file, 'w') as f:
                    json.dump(self.current_profile, f, indent=2)
            
            messagebox.showinfo("Success", f"All settings applied and saved for '{employee_name}'")
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values")
    
    def save_configuration(self):
        """Save current configuration to profile file"""
        if not self.current_profile:
            messagebox.showwarning("Warning", "No profile loaded")
            return
        
        profile_name = self.profile_var.get()
        if not profile_name:
            messagebox.showwarning("Warning", "No profile selected")
            return
        
        # Update profile with current UI values
        self.current_profile['company_name'] = self.company_name.get()
        try:
            self.current_profile['daily_hour_cap'] = float(self.daily_cap.get())
        except ValueError:
            self.current_profile['daily_hour_cap'] = 8.0
        
        # Save to file
        config_file = self.config_dir / f"{profile_name}.json"
        try:
            with open(config_file, 'w') as f:
                json.dump(self.current_profile, f, indent=2)
            
            self.update_profile_info()
            # Update search results
            self.filter_employees()
            
            messagebox.showinfo("Success", "Employee settings saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
    
    # Notes Methods
    
    def update_pay_period_dates(self, event=None):
        """Update pay period options based on selected period type"""
        from datetime import datetime, timedelta
        
        period_type = self.period_type.get()
        today = datetime.now()
        
        # Calculate the start of the current week (Monday)
        current_week_start = today - timedelta(days=today.weekday())
        
        pay_periods = []
        
        if period_type == "2 weeks":
            # Go back 2 weeks from start of current week for the default period
            default_start = current_week_start - timedelta(days=14)
            default_end = current_week_start - timedelta(days=1)  # End of previous week
            
            # Generate several 2-week periods going back
            for i in range(6):  # Show 6 recent 2-week periods
                start_date = default_start - timedelta(days=14 * i)
                end_date = default_end - timedelta(days=14 * i)
                period_str = f"{start_date.strftime('%m/%d/%Y')} - {end_date.strftime('%m/%d/%Y')}"
                pay_periods.append(period_str)
        
        else:  # 1 month
            # Go back 1 month from start of current week for the default period
            default_start = current_week_start - timedelta(days=30)
            default_end = current_week_start - timedelta(days=1)
            
            # Generate several monthly periods going back
            for i in range(6):  # Show 6 recent monthly periods
                start_date = default_start - timedelta(days=30 * i)
                end_date = default_end - timedelta(days=30 * i)
                period_str = f"{start_date.strftime('%m/%d/%Y')} - {end_date.strftime('%m/%d/%Y')}"
                pay_periods.append(period_str)
        
        self.pay_period_combo['values'] = pay_periods
        
        # Set the first (most recent appropriate) period as default
        if pay_periods:
            self.pay_period.set(pay_periods[0])
        
        # Update available dates for notes
        self.update_available_dates_for_notes()
    
    def update_available_dates_for_notes(self):
        """Update available dates based on selected pay period"""
        from datetime import datetime, timedelta
        
        # Check if date_combo exists (might not during initialization)
        if not hasattr(self, 'date_combo'):
            return
        
        pay_period = self.pay_period.get()
        if not pay_period or ' - ' not in pay_period:
            return
        
        try:
            # Parse the pay period dates
            start_str, end_str = pay_period.split(' - ')
            start_date = datetime.strptime(start_str, '%m/%d/%Y')
            end_date = datetime.strptime(end_str, '%m/%d/%Y')
            
            # Generate all weekdays in the pay period
            dates = []
            current_date = start_date
            while current_date <= end_date:
                if current_date.weekday() < 5:  # Only weekdays
                    dates.append(current_date.strftime("%a, %m/%d"))
                current_date += timedelta(days=1)
            
            self.date_combo['values'] = dates
            
        except ValueError:
            # If parsing fails, just clear the dates
            self.date_combo['values'] = []
    
    def on_note_type_change(self, event):
        """Show/hide time fields based on note type"""
        note_type = self.note_type.get()
        
        if note_type in ['Time Change', 'Missing Punch']:
            # Show all time fields for time-based overrides
            self.time_frame.grid()
            # Show all widgets in time frame
            for widget in self.time_frame.winfo_children():
                widget.grid()
        elif note_type == 'Break Change':
            # Show time frame but hide Time In/Out, only show Break Minutes
            self.time_frame.grid()
            # Hide all widgets first
            for widget in self.time_frame.winfo_children():
                widget.grid_remove()
            # Show only Break Minutes label and entry
            for widget in self.time_frame.winfo_children():
                if isinstance(widget, ttk.Label) and widget.cget('text') == 'Break Minutes:':
                    widget.grid()
                elif isinstance(widget, ttk.Entry) and widget['textvariable'] == str(self.note_break_minutes):
                    widget.grid()
        elif note_type == 'Sick Day':
            # Hide time frame for sick day
            self.time_frame.grid_remove()
        else:
            # Hide by default
            self.time_frame.grid_remove()
    
    def add_note(self):
        """Add a new note"""
        if not self.note_employee.get() or not self.note_date.get() or not self.note_type.get():
            messagebox.showwarning("Warning", "Please fill in employee, date, and type")
            return
        
        # Convert display names to internal format
        type_mapping = {
            'Time Change': 'time_override',
            'Break Change': 'break_override',
            'Missing Punch': 'missing_punch_override',
            'Sick Day': 'sick_day'
        }
        
        internal_type = type_mapping.get(self.note_type.get(), self.note_type.get())
        
        note = {
            "employee": self.note_employee.get(),
            "date": self.note_date.get(),
            "type": internal_type,
            "note": self.note_text.get()
        }
        
        # Add additional fields based on type
        if internal_type in ['time_override', 'missing_punch_override']:
            # Combine time and AM/PM for time_in and time_out
            time_in = f"{self.note_time_in_time.get()} {self.note_time_in_ampm.get()}"
            time_out = f"{self.note_time_out_time.get()} {self.note_time_out_ampm.get()}"
            note["time_in"] = time_in
            note["time_out"] = time_out
            if self.note_break_minutes.get():
                note["break_minutes"] = int(self.note_break_minutes.get())
        elif internal_type == 'break_override':
            if self.note_break_minutes.get():
                note["value"] = int(self.note_break_minutes.get())
        
        self.current_notes["notes"].append(note)
        self.refresh_notes_display()
        
        # Clear form
        self.note_employee.set("")
        self.note_date.set("")
        self.note_type.set("")
        self.note_text.set("")
        self.note_time_in_time.set("")
        self.note_time_in_ampm.set("AM")
        self.note_time_out_time.set("")
        self.note_time_out_ampm.set("PM")
        self.note_break_minutes.set("")
    
    def remove_note(self):
        """Remove selected note"""
        selection = self.notes_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a note to remove")
            return
        
        # Get the index of the selected item
        item = self.notes_tree.item(selection[0])
        index = self.notes_tree.index(selection[0])
        
        if messagebox.askyesno("Confirm Remove", "Remove selected note?"):
            self.current_notes["notes"].pop(index)
            self.refresh_notes_display()
    
    def refresh_notes_display(self):
        """Refresh the notes tree display"""
        # Clear current items
        for item in self.notes_tree.get_children():
            self.notes_tree.delete(item)
        
        # Map internal types to display names
        display_type_mapping = {
            'time_override': 'Time Change',
            'break_override': 'Break Change',
            'missing_punch_override': 'Missing Punch',
            'sick_day': 'Sick Day'
        }
        
        # Add notes
        for note in self.current_notes["notes"]:
            details = ""
            if note["type"] in ["time_override", "missing_punch_override"]:
                details = f"{note.get('time_in', '')} - {note.get('time_out', '')}"
            elif note["type"] == "break_override":
                details = f"{note.get('value', '')} min"
            
            display_type = display_type_mapping.get(note["type"], note["type"])
            
            self.notes_tree.insert("", "end", values=(
                note["employee"],
                note["date"],
                display_type,
                details,
                note["note"]
            ))
    
    def save_notes(self):
        """Save current notes to file"""
        if not self.pay_period.get():
            messagebox.showwarning("Warning", "Please enter a pay period")
            return
        
        self.current_notes["pay_period"] = self.pay_period.get()
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=self.notes_dir
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(self.current_notes, f, indent=2)
                messagebox.showinfo("Success", "Notes saved successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save notes: {str(e)}")
    
    def load_notes(self):
        """Load notes from file"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            initialdir=self.notes_dir
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    self.current_notes = json.load(f)
                
                self.pay_period.set(self.current_notes.get("pay_period", ""))
                self.refresh_notes_display()
                messagebox.showinfo("Success", "Notes loaded successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load notes: {str(e)}")
    
    # Processing Methods
    def add_pdf_files(self):
        """Add PDF files for processing"""
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        for file in files:
            self.file_listbox.insert(tk.END, file)
    
    def remove_pdf_file(self):
        """Remove selected PDF file"""
        selection = self.file_listbox.curselection()
        if selection:
            self.file_listbox.delete(selection[0])
    
    def clear_pdf_files(self):
        """Clear all PDF files"""
        self.file_listbox.delete(0, tk.END)
    
    def browse_output(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory()
        if directory:
            self.output_path.set(directory)
    
    def process_files(self):
        """Process the selected files"""
        if not self.current_profile:
            messagebox.showwarning("Warning", "Please load a profile first")
            return
        
        files = list(self.file_listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("Warning", "Please add PDF files to process")
            return
        
        # Save current profile temporarily
        temp_config = "temp_config.json"
        with open(temp_config, 'w') as f:
            json.dump(self.current_profile, f, indent=2)
        
        # Save current notes temporarily if any
        temp_notes = None
        if self.current_notes["notes"]:
            temp_notes = "temp_notes.json"
            with open(temp_notes, 'w') as f:
                json.dump(self.current_notes, f, indent=2)
        
        try:
            self.progress_var.set("Processing...")
            self.progress_bar.start()
            
            # Process each file
            for i, file_path in enumerate(files):
                self.progress_var.set(f"Processing file {i+1} of {len(files)}: {Path(file_path).name}")
                
                # Build command
                cmd = [sys.executable, "timecard_processing/processor.py", temp_config, file_path]
                if temp_notes:
                    cmd.extend(["--notes", temp_notes])
                
                # Run processor
                result = subprocess.run(cmd, capture_output=True, text=True)
                
                if result.returncode != 0:
                    messagebox.showerror("Error", f"Processing failed for {Path(file_path).name}:\n{result.stderr}")
                    break
            
            self.progress_bar.stop()
            self.progress_var.set("Processing complete!")
            messagebox.showinfo("Success", "All files processed successfully!")
            
        except Exception as e:
            self.progress_bar.stop()
            self.progress_var.set("Processing failed")
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
        
        finally:
            # Clean up temporary files
            if os.path.exists(temp_config):
                os.remove(temp_config)
            if temp_notes and os.path.exists(temp_notes):
                os.remove(temp_notes)

    def filter_employee_dropdown(self, *args):
        """Filter employee dropdown based on typed text"""
        if not hasattr(self, 'all_employees'):
            return
        
        typed_text = self.note_employee.get().lower()
        if not typed_text:
            # Show all employees if nothing typed
            self.employee_combo['values'] = self.all_employees
        else:
            # Filter employees that contain the typed text
            filtered = [emp for emp in self.all_employees if typed_text in emp.lower()]
            self.employee_combo['values'] = filtered
    
    def on_employee_keyrelease(self, event):
        """Handle key release in employee combobox"""
        # Open dropdown when typing
        if event.keysym not in ['Up', 'Down', 'Left', 'Right', 'Return', 'Tab']:
            self.employee_combo.event_generate('<Button-1>')

def main():
    root = tk.Tk()
    app = TimeCardGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 