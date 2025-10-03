#!/usr/bin/env python3
"""
XELIFY - Modern Time Card Processing Application
Transform your PDF documents into structured Excel spreadsheets with professional ease
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk
import json
import os
from pathlib import Path
from datetime import datetime, timedelta
import subprocess
import sys
import re
import time
from typing import List, Dict, Any, Tuple

# Set appearance mode and color theme
ctk.set_appearance_mode("dark")  # Professional dark theme
ctk.set_default_color_theme("blue")  # Professional blue theme

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

class XelifyGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("XELIFY - Professional Time Card Processing")
        self.root.geometry("1000x750")
        
        # Show startup logo
        self.show_startup_logo()
        
        # Initialize data
        self.current_profile = {}
        self.current_notes = {"pay_period": "", "notes": []}
        self.current_notes_file_path = None  # Track the currently loaded notes file
        
        self.batch_forms = []
        self.config_dir = Path("timecard_processing/configs")
        self.notes_dir = Path("timecard_processing/note_configs")
        self.config_dir.mkdir(exist_ok=True)
        self.notes_dir.mkdir(exist_ok=True)
        
        # Create settings directory
        if os.name == 'nt':  # Windows
            settings_dir = Path(os.environ.get('APPDATA', Path.home())) / 'TimeCardGUI'
        else:  # macOS/Linux
            settings_dir = Path.home() / '.timecard_gui'
        
        settings_dir.mkdir(exist_ok=True)
        self.settings_file = settings_dir / "settings.json"
        
        # Initialize the main interface after startup
        self.root.after(2000, self.initialize_main_interface)
    
    def show_startup_logo(self):
        """Show XELIFY startup logo"""
        # Create startup frame
        self.startup_frame = ctk.CTkFrame(self.root)
        self.startup_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # XELIFY logo
        logo_label = ctk.CTkLabel(
            self.startup_frame, 
            text="XELIFY", 
            font=ctk.CTkFont(size=48, weight="bold")
        )
        logo_label.pack(pady=(150, 20))
        
        # Subtitle
        subtitle_label = ctk.CTkLabel(
            self.startup_frame, 
            text="Professional Time Card Processing", 
            font=ctk.CTkFont(size=16)
        )
        subtitle_label.pack(pady=(0, 40))
        
        # Loading indicator
        self.loading_label = ctk.CTkLabel(
            self.startup_frame, 
            text="Loading...", 
            font=ctk.CTkFont(size=12)
        )
        self.loading_label.pack(pady=10)
        
        # Progress bar
        self.startup_progress = ctk.CTkProgressBar(self.startup_frame, width=300)
        self.startup_progress.pack(pady=10)
        self.startup_progress.set(0)
        
        # Start progress animation
        self.animate_startup_progress()
    
    def animate_startup_progress(self):
        """Animate the startup progress bar"""
        for i in range(101):
            self.root.after(i * 15, lambda progress=i/100: self.startup_progress.set(progress))
            
        # Update loading text
        self.root.after(500, lambda: self.loading_label.configure(text="Initializing components..."))
        self.root.after(1000, lambda: self.loading_label.configure(text="Loading profiles..."))
        self.root.after(1500, lambda: self.loading_label.configure(text="Ready!"))
    
    def initialize_main_interface(self):
        """Initialize the main interface after startup"""
        # Remove startup frame
        self.startup_frame.destroy()
        
        # Create main interface
        self.create_main_interface()
        self.load_existing_profiles()
        self.cleanup_old_settings()
        self.load_settings()
        
        # Initialize dropdowns with no profile loaded
        self.update_employee_dropdown()
    
    def create_main_interface(self):
        """Create the main interface with modern styling"""
        # Create main container
        main_container = ctk.CTkFrame(self.root, corner_radius=0)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create compact header
        header_frame = ctk.CTkFrame(main_container, height=60, corner_radius=8)
        header_frame.pack(fill='x', padx=10, pady=(10, 5))
        header_frame.pack_propagate(False)
        
        # Compact title
        title_label = ctk.CTkLabel(
            header_frame, 
            text="XELIFY", 
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title_label.pack(side='left', padx=20, pady=15)
        
        # Status indicator
        self.status_label = ctk.CTkLabel(
            header_frame, 
            text="Ready", 
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(side='right', padx=20, pady=15)
        
        # Create tabview
        self.tabview = ctk.CTkTabview(main_container, corner_radius=8)
        self.tabview.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create tabs
        self.create_profiles_tab()
        self.create_configuration_tab()
        self.create_notes_tab()
        self.create_processing_tab()
    
    def create_profiles_tab(self):
        # Create profiles tab
        self.tabview.add("📁 Profiles")
        profiles_frame = self.tabview.tab("📁 Profiles")
        
        # Profile Management section
        profile_mgmt_frame = ctk.CTkFrame(profiles_frame, corner_radius=8)
        profile_mgmt_frame.pack(fill='x', padx=10, pady=10)
        
        # Title label
        title_label = ctk.CTkLabel(profile_mgmt_frame, text="Profile Management", 
                                 font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(pady=(10, 5))
        
        # Profile controls container
        controls_frame = ctk.CTkFrame(profile_mgmt_frame)
        controls_frame.pack(fill='x', padx=10, pady=10)
        
        ctk.CTkLabel(controls_frame, text="Current Profile:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.profile_var = tk.StringVar()
        self.profile_combo = ctk.CTkComboBox(controls_frame, variable=self.profile_var, 
                                   values=[], width=200, state="readonly",
                                   command=self.on_profile_change)
        self.profile_combo.grid(row=0, column=1, padx=5, pady=5)
        
        # Modern button styling
        load_btn = ctk.CTkButton(controls_frame, text="Load", command=self.load_profile, width=80)
        load_btn.grid(row=0, column=2, padx=5, pady=5)
        
        edit_btn = ctk.CTkButton(controls_frame, text="Rename", command=self.edit_profile_name, width=80)
        edit_btn.grid(row=0, column=3, padx=5, pady=5)
        
        delete_btn = ctk.CTkButton(controls_frame, text="Delete", command=self.delete_profile, 
                                 width=80, fg_color="darkred", hover_color="red")
        delete_btn.grid(row=0, column=4, padx=5, pady=5)
        
        # Create New Profile section
        new_profile_frame = ctk.CTkFrame(profiles_frame, corner_radius=8)
        new_profile_frame.pack(fill='x', padx=10, pady=10)
        
        # Title label
        title_label2 = ctk.CTkLabel(new_profile_frame, text="Create New Profile", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label2.pack(pady=(10, 5))
        
        # New profile controls
        new_controls_frame = ctk.CTkFrame(new_profile_frame)
        new_controls_frame.pack(fill='x', padx=10, pady=10)
        
        ctk.CTkLabel(new_controls_frame, text="Profile Name:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.new_profile_name = tk.StringVar()
        ctk.CTkEntry(new_controls_frame, textvariable=self.new_profile_name, width=200).grid(row=0, column=1, padx=5, pady=5)
        
        create_btn = ctk.CTkButton(new_controls_frame, text="Create Profile", command=self.create_new_profile, 
                                 width=120, fg_color="green", hover_color="darkgreen")
        create_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # PDF Analysis section
        pdf_analysis_frame = ctk.CTkFrame(profiles_frame, corner_radius=8)
        pdf_analysis_frame.pack(fill='x', padx=10, pady=10)
        
        # Title label
        title_label3 = ctk.CTkLabel(pdf_analysis_frame, text="PDF Analysis", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label3.pack(pady=(10, 5))
        
        ctk.CTkLabel(pdf_analysis_frame, text="Analyze PDF to extract employee names automatically").pack(pady=5)
        
        # PDF controls
        pdf_controls_frame = ctk.CTkFrame(pdf_analysis_frame)
        pdf_controls_frame.pack(fill='x', padx=10, pady=10)
        
        self.selected_pdf_label = ctk.CTkLabel(pdf_controls_frame, text="No PDF selected", 
                                             text_color="gray")
        self.selected_pdf_label.pack(pady=5)
        
        # PDF buttons
        pdf_buttons_frame = ctk.CTkFrame(pdf_controls_frame)
        pdf_buttons_frame.pack(pady=10)
        
        browse_pdf_btn = ctk.CTkButton(pdf_buttons_frame, text="Browse PDF", 
                                     command=self.browse_pdf_for_profile, width=120)
        browse_pdf_btn.pack(side='left', padx=5)
        
        analyze_btn = ctk.CTkButton(pdf_buttons_frame, text="Analyze & Create", 
                                  command=self.analyze_pdf_and_create_profile, width=120,
                                  fg_color="orange", hover_color="darkorange")
        analyze_btn.pack(side='left', padx=5)
        
        update_btn = ctk.CTkButton(pdf_buttons_frame, text="Update Current", 
                                 command=self.update_profile_from_pdf, width=120)
        update_btn.pack(side='left', padx=5)
        
        # Store selected PDF path
        self.selected_pdf_path = None
    
    def create_configuration_tab(self):
        # Create employees tab
        self.tabview.add("👥 Employees")
        config_frame = self.tabview.tab("👥 Employees")
        
        # Create scrollable frame
        scrollable_frame = ctk.CTkScrollableFrame(config_frame, corner_radius=8)
        scrollable_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Basic Settings
        basic_frame = ctk.CTkFrame(scrollable_frame, corner_radius=8)
        basic_frame.pack(fill='x', padx=10, pady=5)
        
        basic_title = ctk.CTkLabel(basic_frame, text="Basic Settings", 
                                 font=ctk.CTkFont(size=16, weight="bold"))
        basic_title.pack(pady=(10, 5))
        
        basic_controls = ctk.CTkFrame(basic_frame)
        basic_controls.pack(fill='x', padx=10, pady=10)
        
        ctk.CTkLabel(basic_controls, text="Company Name:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.company_name = tk.StringVar()
        ctk.CTkEntry(basic_controls, textvariable=self.company_name, width=200).grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(basic_controls, text="Daily Hour Cap:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.daily_cap = tk.StringVar(value="8.0")
        ctk.CTkEntry(basic_controls, textvariable=self.daily_cap, width=100).grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        # Employee Management
        emp_frame = ctk.CTkFrame(scrollable_frame, corner_radius=8)
        emp_frame.pack(fill='x', padx=10, pady=5)
        
        emp_title = ctk.CTkLabel(emp_frame, text="Employee Management", 
                               font=ctk.CTkFont(size=16, weight="bold"))
        emp_title.pack(pady=(10, 5))
        
        # Search functionality
        search_frame = ctk.CTkFrame(emp_frame)
        search_frame.pack(fill='x', padx=10, pady=5)
        
        ctk.CTkLabel(search_frame, text="Search:").grid(row=0, column=0, padx=5, pady=5)
        self.employee_search = tk.StringVar()
        self.employee_search.trace('w', self.filter_employees)
        search_entry = ctk.CTkEntry(search_frame, textvariable=self.employee_search, width=200)
        search_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Employee list container
        list_container = ctk.CTkFrame(emp_frame)
        list_container.pack(fill='x', padx=10, pady=5)
        
        ctk.CTkLabel(list_container, text="Employees:").pack(anchor='w', padx=5, pady=5)
        
        # Employee list (using tkinter listbox within CTkFrame)
        list_frame = ctk.CTkFrame(list_container)
        list_frame.pack(fill='x', padx=5, pady=5)
        
        self.employee_listbox = tk.Listbox(list_frame, height=6, bg='#2b2b2b', fg='white', 
                                         selectbackground='#1f538d', borderwidth=0)
        self.employee_listbox.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        # Employee controls
        emp_controls = ctk.CTkFrame(emp_frame)
        emp_controls.pack(fill='x', padx=10, pady=5)
        
        ctk.CTkLabel(emp_controls, text="Employee Name:").grid(row=0, column=0, padx=5, pady=5)
        self.new_employee_name = tk.StringVar()
        ctk.CTkEntry(emp_controls, textvariable=self.new_employee_name, width=200).grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkButton(emp_controls, text="Add Employee", command=self.add_employee, width=100,
                    fg_color="green", hover_color="darkgreen").grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(emp_controls, text="Remove Selected", command=self.remove_employee, width=120,
                    fg_color="darkred", hover_color="red").grid(row=0, column=3, padx=5, pady=5)
        
        # Advanced Employee Settings
        settings_frame = ctk.CTkFrame(scrollable_frame, corner_radius=8)
        settings_frame.pack(fill='x', padx=10, pady=5)
        
        settings_title = ctk.CTkLabel(settings_frame, text="Employee Settings", 
                                    font=ctk.CTkFont(size=16, weight="bold"))
        settings_title.pack(pady=(10, 5))
        
        # Basic break settings
        basic_settings = ctk.CTkFrame(settings_frame)
        basic_settings.pack(fill='x', padx=10, pady=5)
        
        ctk.CTkLabel(basic_settings, text="Break Minutes:").grid(row=0, column=0, padx=5, pady=5)
        self.break_minutes = tk.StringVar()
        ctk.CTkEntry(basic_settings, textvariable=self.break_minutes, width=100).grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(basic_settings, text="Hour Threshold:").grid(row=0, column=2, padx=5, pady=5)
        self.hour_threshold = tk.StringVar()
        ctk.CTkEntry(basic_settings, textvariable=self.hour_threshold, width=100).grid(row=0, column=3, padx=5, pady=5)
        
        # Fixed hours setting
        fixed_frame = ctk.CTkFrame(basic_settings)
        fixed_frame.grid(row=1, column=0, columnspan=4, sticky='ew', pady=5)
        
        self.use_fixed_hours = tk.BooleanVar()
        ctk.CTkCheckBox(fixed_frame, text="Use Fixed Hours:", variable=self.use_fixed_hours).grid(row=0, column=0, padx=5, pady=5)
        self.fixed_hours = tk.StringVar()
        ctk.CTkEntry(fixed_frame, textvariable=self.fixed_hours, width=100).grid(row=0, column=1, padx=5, pady=5)
        
        # Escalating breaks
        escalating_frame = ctk.CTkFrame(settings_frame, corner_radius=8)
        escalating_frame.pack(fill='x', padx=10, pady=5)
        
        ctk.CTkLabel(escalating_frame, text="Use escalating breaks:").grid(row=0, column=0, padx=5, pady=5)
        self.use_escalating = tk.BooleanVar()
        ctk.CTkCheckBox(escalating_frame, text="", variable=self.use_escalating, command=self.toggle_escalating_options).grid(row=0, column=1, padx=5, pady=5)
        
        # Escalating break options (initially hidden)
        self.escalating_options = ctk.CTkFrame(escalating_frame, fg_color="transparent", height=30)
        self.escalating_options.grid(row=1, column=0, columnspan=6, sticky='ew', pady=2)
        
        ctk.CTkLabel(self.escalating_options, text="6+ hours:").grid(row=0, column=0, padx=5, pady=2)
        self.break_6_hours = tk.StringVar(value="30")
        ctk.CTkEntry(self.escalating_options, textvariable=self.break_6_hours, width=80).grid(row=0, column=1, padx=5, pady=2)
        ctk.CTkLabel(self.escalating_options, text="min").grid(row=0, column=2, padx=2, pady=2)
        
        ctk.CTkLabel(self.escalating_options, text="7+ hours:").grid(row=0, column=3, padx=5, pady=2)
        self.break_7_hours = tk.StringVar(value="60")
        ctk.CTkEntry(self.escalating_options, textvariable=self.break_7_hours, width=80).grid(row=0, column=4, padx=5, pady=2)
        ctk.CTkLabel(self.escalating_options, text="min").grid(row=0, column=5, padx=2, pady=2)
        
        # Hide escalating options initially
        self.toggle_escalating_options()
        
        # Apply button
        apply_frame = ctk.CTkFrame(settings_frame)
        apply_frame.pack(fill='x', padx=10, pady=10)
        
        ctk.CTkButton(apply_frame, text="Apply Settings", command=self.apply_employee_settings,
                    width=150, fg_color="green", hover_color="darkgreen").pack(pady=10)
        
        # Bind employee selection (double-click to select and maintain selection)
        self.employee_listbox.bind('<Double-Button-1>', self.on_employee_select)
        self.employee_listbox.bind('<FocusOut>', self.prevent_deselection)
        
        # Track selected employee to maintain selection
        self.selected_employee = None
    
    def create_notes_tab(self):
        # Create notes tab
        self.tabview.add("📁 Notes")
        notes_tab = self.tabview.tab("📁 Notes")
        
        # Use CTkScrollableFrame for easier scrolling
        self.notes_frame = ctk.CTkScrollableFrame(notes_tab, corner_radius=8)
        self.notes_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Pay period with period type selection
        period_frame = ctk.CTkFrame(self.notes_frame, corner_radius=8)
        period_frame.pack(fill='x', padx=10, pady=5)
        
        # Period type selection
        ctk.CTkLabel(period_frame, text="Period Type:").grid(row=0, column=0, padx=5)
        self.period_type = tk.StringVar(value="2 weeks")
        period_combo = ctk.CTkComboBox(period_frame, variable=self.period_type, values=["2 weeks", "1 month"], width=120, state='readonly',
                                     command=self.update_pay_period_dates)
        period_combo.grid(row=0, column=1, padx=5)
        
        ctk.CTkLabel(period_frame, text="Pay Period:").grid(row=0, column=2, sticky='w', padx=5)
        self.pay_period = tk.StringVar()
        self.pay_period_combo = ctk.CTkComboBox(period_frame, variable=self.pay_period, width=250, state='readonly',
                                              command=lambda choice: self.update_available_dates_for_notes())
        self.pay_period_combo.grid(row=0, column=3, padx=5)
        
        # Initialize pay period options
        self.update_pay_period_dates()
        
        # Add new note
        add_note_frame = ctk.CTkFrame(self.notes_frame, corner_radius=8)
        add_note_frame.pack(fill='x', padx=10, pady=5)
        
        # Configure grid weights for consistent alignment
        add_note_frame.columnconfigure(1, weight=1)
        
        # Batch mode at the top
        batch_control_frame = ctk.CTkFrame(add_note_frame, fg_color="transparent")
        batch_control_frame.grid(row=0, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        
        self.batch_mode = tk.BooleanVar()
        self.batch_mode_checkbox = ctk.CTkCheckBox(batch_control_frame, text="Batch Mode:", 
                       variable=self.batch_mode, command=self.toggle_batch_mode, state='disabled')
        self.batch_mode_checkbox.grid(row=0, column=0, sticky='w', padx=5)
        
        ctk.CTkLabel(batch_control_frame, text="Count:").grid(row=0, column=1, sticky='w', padx=5)
        self.batch_count = tk.StringVar(value="1")
        self.batch_count_entry = ctk.CTkEntry(batch_control_frame, textvariable=self.batch_count, width=5, state='disabled')
        self.batch_count_entry.grid(row=0, column=2, padx=5)
        self.batch_count_entry.bind('<KeyRelease>', self.on_batch_count_change)
        self.batch_count_entry.bind('<FocusOut>', self.on_batch_count_change)
        self.batch_count_entry.bind('<Return>', self.on_batch_count_change)
        
        # Employee selection
        ctk.CTkLabel(add_note_frame, text="Employee:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.note_employee = tk.StringVar()
        self.employee_combo = ctk.CTkComboBox(add_note_frame, variable=self.note_employee, 
                                         width=300, state='normal')
        self.employee_combo.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        # Add helpful instruction
        instruction_label = ctk.CTkLabel(add_note_frame, text="(Type to search, ↓ or click arrow to open dropdown)", 
                                    font=ctk.CTkFont(size=10), text_color='gray')
        instruction_label.grid(row=1, column=2, sticky='w', padx=5, pady=2)
        
        # Add search functionality to employee combobox
        self.note_employee.trace('w', self.filter_employee_dropdown)
        self.note_employee.trace('w', self.on_main_employee_change)  # Update batch forms when employee changes
        # Add manual dropdown opening with Down arrow key
        self.employee_combo.bind('<Down>', self.open_employee_dropdown)
        self.employee_combo.bind('<Button-1>', self.on_employee_click)
        # Prevent text selection highlighting
        self.employee_combo.bind('<FocusIn>', lambda e: self.employee_combo.selection_clear())
        # Disable scroll wheel to prevent accidental changes
        self.employee_combo.bind('<MouseWheel>', lambda e: 'break')
        
        # Date selection
        ctk.CTkLabel(add_note_frame, text="Date:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.note_date = tk.StringVar()
        self.date_combo = ctk.CTkComboBox(add_note_frame, variable=self.note_date, 
                                     width=200, state='readonly')
        self.date_combo.grid(row=2, column=1, sticky='w', padx=5, pady=2)
        # Disable scroll wheel to prevent accidental changes
        self.date_combo.bind('<MouseWheel>', lambda e: 'break')
        
        # Now that date_combo is created, update the available dates
        self.update_available_dates_for_notes()
        
        # Initialize employee dropdown properly
        self.update_employee_dropdown()
        
        # Override type
        ctk.CTkLabel(add_note_frame, text="Override Type:").grid(row=3, column=0, sticky='w', padx=5, pady=2)
        self.note_type = tk.StringVar(value="Missing Punch")  # Set default value
        type_combo = ctk.CTkComboBox(add_note_frame, variable=self.note_type, width=200, 
                                state='readonly', values=('Time Change', 'Break Change', 'Missing Punch', 'Sick Day', 'Misc', 'ROE'),
                                command=self.on_main_note_type_change)
        type_combo.grid(row=3, column=1, sticky='w', padx=5, pady=2)
        # Disable scroll wheel to prevent accidental changes
        type_combo.bind('<MouseWheel>', lambda e: 'break')
        
        # Time fields container (will show/hide based on override type)
        self.time_fields_container = ctk.CTkFrame(add_note_frame, fg_color="transparent")
        self.time_fields_container.grid(row=4, column=0, columnspan=2, sticky='ew', pady=5)
        
        # Full time frame for Time Change/Missing Punch
        self.time_frame = ctk.CTkFrame(self.time_fields_container, fg_color="transparent", border_width=1, border_color="gray")
        self.time_frame.grid(row=0, column=0, sticky='ew', padx=5, pady=5)
        
        # Time In
        self.time_in_label = ctk.CTkLabel(self.time_frame, text="Time In:")
        self.time_in_label.grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.time_in_frame = ctk.CTkFrame(self.time_frame, fg_color="transparent")
        self.time_in_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        self.note_time_in_time = tk.StringVar(value="__:__")
        time_in_entry = ctk.CTkEntry(self.time_in_frame, textvariable=self.note_time_in_time, width=80)
        time_in_entry.pack(side='left', padx=(0,2))
        time_in_entry.bind('<KeyPress>', lambda e: self.on_time_key_press(e, self.note_time_in_time))
        time_in_entry.bind('<FocusOut>', lambda e: self.on_time_focus_out_smart(self.note_time_in_time))
        time_in_entry.bind('<Button-1>', lambda e: self.on_time_click(e, self.note_time_in_time))
        
        self.note_time_in_ampm = tk.StringVar(value="AM")
        time_in_ampm = ctk.CTkComboBox(self.time_in_frame, variable=self.note_time_in_ampm, width=50, values=['AM', 'PM'], state='readonly',
                                     command=lambda choice: self.root.after(100, self.calculate_note_total_hours))
        time_in_ampm.pack(side='left')
        # Disable scroll wheel to prevent accidental changes
        time_in_ampm.bind('<MouseWheel>', lambda e: 'break')
        
        # Time Out
        self.time_out_label = ctk.CTkLabel(self.time_frame, text="Time Out:")
        self.time_out_label.grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.time_out_frame = ctk.CTkFrame(self.time_frame, fg_color="transparent")
        self.time_out_frame.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        self.note_time_out_time = tk.StringVar(value="__:__")
        time_out_entry = ctk.CTkEntry(self.time_out_frame, textvariable=self.note_time_out_time, width=80)
        time_out_entry.pack(side='left', padx=(0,2))
        time_out_entry.bind('<KeyPress>', lambda e: self.on_time_key_press(e, self.note_time_out_time))
        time_out_entry.bind('<FocusOut>', lambda e: self.on_time_focus_out_smart(self.note_time_out_time))
        time_out_entry.bind('<Button-1>', lambda e: self.on_time_click(e, self.note_time_out_time))
        
        self.note_time_out_ampm = tk.StringVar(value="PM")
        time_out_ampm = ctk.CTkComboBox(self.time_out_frame, variable=self.note_time_out_ampm, width=50, values=['AM', 'PM'], state='readonly',
                                      command=lambda choice: self.root.after(100, self.calculate_note_total_hours))
        time_out_ampm.pack(side='left')
        # Disable scroll wheel to prevent accidental changes
        time_out_ampm.bind('<MouseWheel>', lambda e: 'break')
        
        # Break Minutes in main time frame
        self.break_minutes_label = ctk.CTkLabel(self.time_frame, text="Break Minutes:")
        self.break_minutes_label.grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.note_break_minutes = tk.StringVar()
        self.break_minutes_entry = ctk.CTkEntry(self.time_frame, textvariable=self.note_break_minutes, width=80)
        self.break_minutes_entry.grid(row=2, column=1, sticky='w', padx=5, pady=2)
        
        # Total Hours (calculated field)
        self.total_hours_label = ctk.CTkLabel(self.time_frame, text="Total Hours:")
        self.total_hours_label.grid(row=3, column=0, sticky='w', padx=5, pady=2)
        self.note_total_hours = tk.StringVar()
        self.total_hours_entry = ctk.CTkEntry(self.time_frame, textvariable=self.note_total_hours, width=80, state='readonly')
        self.total_hours_entry.grid(row=3, column=1, sticky='w', padx=5, pady=2)
        
        # Configure the container's grid
        self.time_fields_container.grid_columnconfigure(0, weight=1)
        
        # Break-only frame for Break Change (separate simple frame)
        self.break_only_frame = ctk.CTkFrame(self.time_fields_container, fg_color="transparent", border_width=1, border_color="gray")
        self.break_only_frame.grid(row=1, column=0, sticky='ew', padx=5, pady=5)
        
        self.break_only_label = ctk.CTkLabel(self.break_only_frame, text="Break Minutes:")
        self.break_only_label.grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.break_only_entry = ctk.CTkEntry(self.break_only_frame, textvariable=self.note_break_minutes, width=80)
        self.break_only_entry.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        # Initially hide both frames (no note type selected)
        self.time_frame.grid_remove()
        self.break_only_frame.grid_remove()
        
        # Note: Time calculation is now handled by KeyRelease and ComboboxSelected events
        
        # Note text
        ctk.CTkLabel(add_note_frame, text="Note:").grid(row=5, column=0, sticky='nw', padx=5, pady=2)
        self.note_text = tk.StringVar()
        ctk.CTkEntry(add_note_frame, textvariable=self.note_text, width=400).grid(row=5, column=1, sticky='ew', padx=5, pady=2)
        
        self.add_note_button = ctk.CTkButton(add_note_frame, text="Add Note", command=self.add_note)
        self.add_note_button.grid(row=6, column=1, sticky='w', padx=5, pady=10)
        
        # Notes list
        self.notes_list_frame = ctk.CTkFrame(self.notes_frame, corner_radius=8)
        self.notes_list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create treeview for notes
        columns = ('Employee', 'Date', 'Type', 'Details', 'Note')
        self.notes_tree = ttk.Treeview(self.notes_list_frame, columns=columns, show='headings', height=10)
        
        for col in columns:
            self.notes_tree.heading(col, text=col)
            self.notes_tree.column(col, width=120)
        
        notes_scroll = ttk.Scrollbar(self.notes_list_frame, orient="vertical", command=self.notes_tree.yview)
        self.notes_tree.configure(yscrollcommand=notes_scroll.set)
        
        self.notes_tree.pack(side="left", fill="both", expand=True)
        notes_scroll.pack(side="right", fill="y")
        
        # Notes controls
        notes_controls = ctk.CTkFrame(self.notes_list_frame)
        notes_controls.pack(fill='x', pady=5)
        
        ctk.CTkButton(notes_controls, text="Edit Selected", command=self.edit_note).pack(fill='x', pady=2)
        ctk.CTkButton(notes_controls, text="Remove Selected", command=self.remove_note).pack(fill='x', pady=2)
        ctk.CTkButton(notes_controls, text="Save Notes", command=self.save_notes).pack(fill='x', pady=2)
        ctk.CTkButton(notes_controls, text="Load Notes", command=self.load_notes).pack(fill='x', pady=2)
        
        # Container for batch forms (will be populated dynamically) - created but not packed
        self.batch_forms_container = ctk.CTkFrame(self.notes_frame, fg_color="transparent")
        
        # Trigger the default note type change to show correct fields
        self.on_note_type_change("Missing Punch")
    
    def create_processing_tab(self):
        # Create processing tab
        self.tabview.add("📊 Processing")
        processing_frame = self.tabview.tab("📊 Processing")
        
        # Main container with modern layout
        main_container = ctk.CTkFrame(processing_frame, corner_radius=8)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # PDF File section (singular)
        pdf_frame = ctk.CTkFrame(main_container, corner_radius=8)
        pdf_frame.pack(side='left', fill='both', expand=True, padx=(10, 5), pady=10)
        
        # Section title
        pdf_title = ctk.CTkLabel(pdf_frame, text="PDF File", font=ctk.CTkFont(size=18, weight="bold"))
        pdf_title.pack(pady=(15, 10))
        
        # Upload area with modern styling and proper drag/drop handling
        upload_frame = ctk.CTkFrame(pdf_frame, corner_radius=8, fg_color="transparent", border_width=2, border_color="gray")
        upload_frame.pack(fill='x', padx=15, pady=10)
        
        upload_label = ctk.CTkLabel(upload_frame, text="🔍 Click to select a PDF file", 
                                  font=ctk.CTkFont(size=14), cursor="hand2")
        upload_label.pack(pady=30)
        
        # Selected file display
        self.selected_file_label = ctk.CTkLabel(pdf_frame, text="No file selected", 
                                              font=ctk.CTkFont(size=12), text_color="gray")
        self.selected_file_label.pack(pady=10)
        
        # Control buttons with modern styling
        btn_frame = ctk.CTkFrame(pdf_frame)
        btn_frame.pack(fill='x', padx=15, pady=10)
        
        ctk.CTkButton(btn_frame, text="📁 Select File", command=self.select_pdf_file, width=120, 
                    fg_color="green", hover_color="darkgreen").pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="🗑️ Clear", command=self.clear_pdf_file, width=100,
                    fg_color="darkred", hover_color="red").pack(side='left', padx=5)
        
        # Processing Control section
        control_frame = ctk.CTkFrame(main_container, corner_radius=8)
        control_frame.pack(side='right', fill='both', expand=True, padx=(5, 10), pady=10)
        
        # Section title
        control_title = ctk.CTkLabel(control_frame, text="Processing Control", font=ctk.CTkFont(size=18, weight="bold"))
        control_title.pack(pady=(15, 20))
        
        # Output location with modern layout
        output_section = ctk.CTkFrame(control_frame, corner_radius=8)
        output_section.pack(fill='x', padx=15, pady=10)
        
        ctk.CTkLabel(output_section, text="Output Directory:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor='w', padx=10, pady=(10, 5))
        
        output_frame = ctk.CTkFrame(output_section)
        output_frame.pack(fill='x', padx=10, pady=(5, 15))
        
        self.output_path = tk.StringVar()
        ctk.CTkEntry(output_frame, textvariable=self.output_path, height=35).pack(side='left', fill='x', expand=True, padx=(5, 5))
        ctk.CTkButton(output_frame, text="📂 Browse", command=self.browse_output, width=80).pack(side='right', padx=(5, 5))
        
        # Processing progress with modern styling
        progress_section = ctk.CTkFrame(control_frame, corner_radius=8)
        progress_section.pack(fill='x', padx=15, pady=10)
        
        ctk.CTkLabel(progress_section, text="Processing Status:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor='w', padx=10, pady=(15, 5))
        
        self.progress_var = tk.StringVar(value="Ready to process")
        self.progress_label = ctk.CTkLabel(progress_section, textvariable=self.progress_var, font=ctk.CTkFont(size=12))
        self.progress_label.pack(anchor='w', padx=10, pady=5)
        
        self.progress_bar = ctk.CTkProgressBar(progress_section, height=20)
        self.progress_bar.pack(fill='x', padx=10, pady=(5, 15))
        self.progress_bar.set(0)
        
        # Process button with prominent styling
        process_section = ctk.CTkFrame(control_frame, corner_radius=8)
        process_section.pack(fill='x', padx=15, pady=20)
        
        process_btn = ctk.CTkButton(process_section, text="🚀 PROCESS FILE", 
                                  command=self.process_files, height=50, font=ctk.CTkFont(size=16, weight="bold"),
                                  fg_color="#1f538d", hover_color="#164069")
        process_btn.pack(fill='x', padx=15, pady=15)
        
        # Store selected PDF file path
        self.selected_pdf_file = None
        
        # Bind upload area click
        upload_label.bind("<Button-1>", lambda e: self.select_pdf_file())
    
    # Profile Management Methods
    def load_existing_profiles(self):
        """Load existing config files into profile dropdown"""
        profiles = []
        if self.config_dir.exists():
            for config_file in self.config_dir.glob("*.json"):
                profiles.append(config_file.stem)
        self.profile_combo.configure(values=profiles)
        
        # Also update employee dropdown in notes
        if hasattr(self, 'note_employee_combo'):
            self.update_employee_dropdown()
    
    def on_profile_change(self, choice=None):
        """Handle profile selection change"""
        profile_name = self.profile_var.get()
        if not profile_name:
            # If no profile selected, just reset fields
            self.reset_all_notes_fields()
            return
        
        # Load the selected profile automatically
        config_file = self.config_dir / f"{profile_name}.json"
        try:
            with open(config_file, 'r') as f:
                self.current_profile = json.load(f)
            
            # Reset all fields when profile changes
            self.reset_all_notes_fields()
            
            # Update UI with loaded profile
            self.update_configuration_ui()
            self.update_employee_dropdown()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load profile: {str(e)}")
            self.reset_all_notes_fields()
    
    def reset_all_notes_fields(self):
        """Reset all notes form fields to default state"""
        # Reset main form
        self.note_employee.set("")
        self.note_date.set("")
        self.note_type.set("Missing Punch")  # Set default override type
        self.note_text.set("")
        self.note_time_in_time.set("__:__")
        self.note_time_in_ampm.set("AM")
        self.note_time_out_time.set("__:__")
        self.note_time_out_ampm.set("PM")
        self.note_break_minutes.set("")
        self.note_total_hours.set("")
        
        # Clear batch forms
        self.clear_batch_forms()
        self.batch_mode.set(False)
        self.batch_count.set("1")
        self.batch_count_entry.configure(state='disabled')
        
        # Clear current notes
        self.current_notes = {"pay_period": "", "notes": []}
        self.refresh_notes_display()
        
        # Clear the pay period dropdown
        if hasattr(self, 'pay_period'):
            self.pay_period.set("")
            
        # Update employee dropdown to show empty state
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
            
            # Reset all fields when profile is loaded
            self.reset_all_notes_fields()
            
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
                self.update_employee_dropdown()  # Clear employee dropdown
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
            self.selected_pdf_label.configure(text=f"Selected: {filename}", text_color='blue')
    
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
            self.selected_pdf_label.configure(text="No PDF selected", text_color='gray')
            
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
            self.selected_pdf_label.configure(text="No PDF selected", text_color='gray')
            
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
            # No profile loaded - clear dropdown completely
            self.all_employees = []
            if hasattr(self, 'employee_combo'):
                self.employee_combo.configure(values=[])
                self.employee_combo.set("")
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
            if self.all_employees:
                self.employee_combo.configure(values=self.all_employees)
            else:
                self.employee_combo.configure(values=["No employees in profile"])
                self.employee_combo.set("")
    
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
            # Show escalating options frame and its contents
            self.escalating_options.grid()
            for widget in self.escalating_options.winfo_children():
                widget.grid()
        else:
            # Hide the entire escalating options frame
            self.escalating_options.grid_remove()
    
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
                break_config = self.current_profile['break_config']
                break_config.get('employee_breaks', {}).pop(employee_name, None)
                break_config.get('employee_hour_thresholds', {}).pop(employee_name, None)
                break_config.get('fixed_hours_employees', {}).pop(employee_name, None)
                break_config.get('escalating_breaks', {}).pop(employee_name, None)
            
            # Update UI
            self.employee_listbox.delete(selection[0])
            self.update_employee_dropdown()
            self.clear_employee_form()  # Clear the form after removal
            
            messagebox.showinfo("Success", f"Employee '{employee_name}' removed successfully")
    
    def on_employee_select(self, event):
        """Handle employee selection - load their current settings"""
        selection = self.employee_listbox.curselection()
        if not selection:
            self.clear_employee_form()
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
            self.fixed_hours.set("")  # Empty instead of default 8.0
        
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
    
    def clear_employee_form(self):
        """Clear all employee form fields"""
        self.break_minutes.set("")
        self.hour_threshold.set("")
        self.use_fixed_hours.set(False)
        self.fixed_hours.set("")
        self.use_escalating.set(False)
        self.break_6_hours.set("")
        self.break_7_hours.set("")
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
            
            # Update basic settings
            break_config.setdefault('employee_breaks', {})[employee_name] = break_min
            
            if hour_thresh is not None:
                break_config.setdefault('employee_hour_thresholds', {})[employee_name] = hour_thresh
            else:
                break_config.get('employee_hour_thresholds', {}).pop(employee_name, None)
            
            # Fixed hours - only apply if checkbox is checked AND field has value
            if self.use_fixed_hours.get() and self.fixed_hours.get():
                fixed_hrs = float(self.fixed_hours.get())
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
    
    def update_pay_period_dates(self, choice=None):
        """Update pay period options based on selected period type"""
        from datetime import datetime, timedelta
        
        period_type = choice if choice else self.period_type.get()
        today = datetime.now()
        
        pay_periods = []
        
        if period_type == "2 weeks":
            # Calculate the current bi-weekly period using proper bi-weekly logic
            # Using first Monday of the year as anchor date to align with desired periods
            
            # Set the anchor date (first Monday of current year)
            jan_1 = datetime(today.year, 1, 1)
            # Find the first Monday of the year
            days_until_monday = (7 - jan_1.weekday()) % 7
            if days_until_monday == 0 and jan_1.weekday() != 0:  # If Jan 1 is not Monday
                days_until_monday = 7
            anchor_date = jan_1 + timedelta(days=days_until_monday)
            
            # Calculate periods since anchor date
            days_since_anchor = (today - anchor_date).days
            periods_since_anchor = days_since_anchor // 14
            
            # Calculate the start of the current bi-weekly period
            current_period_start = anchor_date + timedelta(days=periods_since_anchor * 14)
            current_period_end = current_period_start + timedelta(days=13)  # 14 days total (0-13)
            
            # Add the current period as the first option
            current_period_str = f"{current_period_start.strftime('%m/%d/%Y')} - {current_period_end.strftime('%m/%d/%Y')}"
            pay_periods.append(current_period_str)
            
            # Generate previous 2-week periods going back
            for i in range(1, 6):  # Show 5 previous periods plus current
                start_date = current_period_start - timedelta(days=14 * i)
                end_date = current_period_end - timedelta(days=14 * i)
                period_str = f"{start_date.strftime('%m/%d/%Y')} - {end_date.strftime('%m/%d/%Y')}"
                pay_periods.append(period_str)
        
        else:  # 1 month
            # For monthly periods, use the current month as the first option
            current_month_start = datetime(today.year, today.month, 1)
            if today.month == 12:
                next_month_start = datetime(today.year + 1, 1, 1)
            else:
                next_month_start = datetime(today.year, today.month + 1, 1)
            current_month_end = next_month_start - timedelta(days=1)
            
            # Add the current month as the first option
            current_period_str = f"{current_month_start.strftime('%m/%d/%Y')} - {current_month_end.strftime('%m/%d/%Y')}"
            pay_periods.append(current_period_str)
            
            # Generate previous monthly periods going back
            for i in range(1, 6):  # Show 5 previous months plus current
                if current_month_start.month == 1:
                    prev_month_start = datetime(current_month_start.year - 1, 12, 1)
                else:
                    prev_month_start = datetime(current_month_start.year, current_month_start.month - 1, 1)
                
                if prev_month_start.month == 12:
                    next_month_start = datetime(prev_month_start.year + 1, 1, 1)
                else:
                    next_month_start = datetime(prev_month_start.year, prev_month_start.month + 1, 1)
                prev_month_end = next_month_start - timedelta(days=1)
                
                period_str = f"{prev_month_start.strftime('%m/%d/%Y')} - {prev_month_end.strftime('%m/%d/%Y')}"
                pay_periods.append(period_str)
                
                # Update for next iteration
                current_month_start = prev_month_start
        
        self.pay_period_combo.configure(values=pay_periods)
        
        # Set the first (current) period as default
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
            # No pay period selected - show helpful message
            self.date_combo.configure(values=["Select pay period first"])
            self.date_combo.set("")
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
            
            if dates:
                self.date_combo.configure(values=dates)
            else:
                self.date_combo.configure(values=["No weekdays in period"])
                self.date_combo.set("")
            
        except ValueError:
            # If parsing fails, show error message
            self.date_combo.configure(values=["Invalid pay period format"])
            self.date_combo.set("")
    
    def on_main_note_type_change(self, choice):
        """Handle main form note type change and update batch forms"""
        # Update main form first
        self.on_note_type_change(choice)
        
        # Update all batch forms to match the main form type
        if self.batch_mode.get() and hasattr(self, 'batch_forms'):
            for batch_form in self.batch_forms:
                if 'vars' in batch_form and 'type' in batch_form['vars']:
                    batch_form['vars']['type'].set(choice)
                    # Trigger the batch form's type change
                    if 'time_frame' in batch_form:
                        self.on_batch_note_type_change(batch_form['time_frame'], batch_form['vars']['type'])
    
    def on_note_type_change(self, choice):
        """Show/hide time fields based on note type"""
        note_type = choice if choice else self.note_type.get()
        
        # Hide all frames first
        self.time_frame.grid_remove()
        self.break_only_frame.grid_remove()
        
        if note_type in ['Time Change', 'Missing Punch']:
            # Show full time frame with all fields
            self.time_frame.grid(row=0, column=0, sticky='ew', padx=5, pady=5)
            # Force update to make sure the frame is visible
            self.time_fields_container.update_idletasks()
        elif note_type == 'Break Change':
            # Show only break minutes frame
            self.break_only_frame.grid(row=1, column=0, sticky='ew', padx=5, pady=5)
            # Force update to make sure the frame is visible
            self.time_fields_container.update_idletasks()
        elif note_type in ['Sick Day', 'Misc', 'ROE']:
            # Hide all time fields - both frames remain hidden (only note field needed)
            pass
        else:
            # Hide by default - both frames remain hidden
            pass
    
    def on_batch_note_type_change(self, time_frame, type_var):
        """Show/hide time fields based on note type for batch forms"""
        note_type = type_var.get()
        
        if note_type in ['Time Change', 'Missing Punch']:
            # Show all time fields for time-based overrides
            time_frame.grid()
            # Show all widgets in time frame
            for widget in time_frame.winfo_children():
                widget.grid()
        elif note_type == 'Break Change':
            # Show time frame but hide Time In/Out, only show Break Minutes
            time_frame.grid()
            
            # Get all widgets and their original grid info
            widgets_grid_info = []
            for widget in time_frame.winfo_children():
                grid_info = widget.grid_info()
                widgets_grid_info.append((widget, grid_info))
            
            # Hide all widgets first
            for widget, _ in widgets_grid_info:
                widget.grid_remove()
            
            # Show only Break Minutes label and entry (row=2)
            for widget, grid_info in widgets_grid_info:
                if grid_info and grid_info.get('row') == 2:  # Break minutes row
                    # Restore the widget with its original grid settings
                    widget.grid(**grid_info)
        elif note_type in ['Sick Day', 'Misc', 'ROE']:
            # Hide time frame for note-only types
            time_frame.grid_remove()
        else:
            # Hide by default
            time_frame.grid_remove()
    
    def format_time_input(self, time_var):
        """Auto-format time input with colon - simple insertion only"""
        try:
            current_text = time_var.get()
            # Remove any existing colons and non-digits
            digits_only = ''.join(filter(str.isdigit, current_text))
            
            if len(digits_only) == 0:
                return
            elif len(digits_only) == 1:
                # Single digit like "5" stays as "5"
                formatted = digits_only
            elif len(digits_only) == 2:
                # Two digits like "12" stays as "12"
                formatted = digits_only
            elif len(digits_only) == 3:
                # Three digits like "530" becomes "5:30"
                formatted = f"{digits_only[0]}:{digits_only[1:3]}"
            elif len(digits_only) >= 4:
                # Four+ digits like "1030" becomes "10:30"
                formatted = f"{digits_only[0:2]}:{digits_only[2:4]}"
            
            # Only update if different to avoid infinite loop
            if time_var.get() != formatted:
                time_var.set(formatted)
                
        except Exception:
            pass  # If formatting fails, leave as is
    
    def on_time_key_press(self, event, time_var):
        """Handle key press in time field with locked colon"""
        char = event.char
        
        # Only allow digits, backspace, delete, and arrow keys
        if not (char.isdigit() or event.keysym in ['BackSpace', 'Delete', 'Left', 'Right', 'Tab']):
            return 'break'
        
        # Handle arrow key navigation
        if event.keysym in ['Left', 'Right']:
            current = time_var.get()
            cursor_pos = event.widget.index(tk.INSERT)
            
            if event.keysym == 'Right':
                if cursor_pos == 1:  # Moving from first digit position
                    # Check if second digit position is empty (underscore)
                    if len(current) > 1 and current[1] == '_':
                        # Jump to after colon for single digit hour
                        event.widget.icursor(3)
                        return 'break'
                elif cursor_pos == 2:  # At colon position
                    event.widget.icursor(3)  # Jump to after colon
                    return 'break'
                elif cursor_pos >= len(current) - 1:  # At end
                    return 'break'
            elif event.keysym == 'Left':
                if cursor_pos == 3:  # Moving from after colon to before
                    # Check if we should jump to position 1 (single digit) or 2 (double digit)
                    if len(current) > 1 and current[1] == '_':
                        event.widget.icursor(1)  # Jump to after first digit for single digit hour
                    else:
                        event.widget.icursor(2)  # Jump to before colon for double digit hour
                    return 'break'
                elif cursor_pos == 2:  # At colon position
                    event.widget.icursor(1)  # Jump to before colon
                    return 'break'
                elif cursor_pos <= 0:  # At beginning
                    return 'break'
            
            # Allow normal arrow key behavior for other positions
            return
        
        # Handle backspace and delete
        if event.keysym in ['BackSpace', 'Delete']:
            current = time_var.get()
            cursor_pos = event.widget.index(tk.INSERT)
            
            if event.keysym == 'BackSpace' and cursor_pos > 0:
                if cursor_pos == 3:  # Trying to delete the colon
                    # Replace the character before colon with underscore
                    new_time = current[:1] + '_:' + current[3:]
                    time_var.set(new_time)
                    event.widget.icursor(1)
                elif cursor_pos > 3:  # After colon
                    new_time = current[:cursor_pos-1] + '_' + current[cursor_pos:]
                    time_var.set(new_time)
                    event.widget.icursor(cursor_pos-1)
                else:  # Before colon
                    new_time = '_' + current[1:]
                    time_var.set(new_time)
                    event.widget.icursor(0)
            
            return 'break'
        
        # Handle digit input
        if char.isdigit():
            current = time_var.get()
            cursor_pos = event.widget.index(tk.INSERT)
            
            # Skip over colon
            if cursor_pos == 2:
                cursor_pos = 3
            
            # Replace underscore with digit
            if cursor_pos < len(current):
                new_time = current[:cursor_pos] + char + current[cursor_pos+1:]
                
                # Validate minutes if we're in the minutes section
                if cursor_pos >= 3:  # Minutes section
                    # Check if the minutes would be valid
                    try:
                        minutes_part = new_time[3:5]
                        if minutes_part.replace('_', '0').isdigit():
                            minutes_val = int(minutes_part.replace('_', '0'))
                            if cursor_pos == 3 and int(char) > 5:  # First digit of minutes can't be > 5
                                return 'break'
                            elif cursor_pos == 4 and minutes_part[0] == '6':  # If first digit is 6, second must be 0
                                if int(char) > 0:
                                    return 'break'
                    except:
                        pass
                
                time_var.set(new_time)
                
                # Move cursor to next position
                next_pos = cursor_pos + 1
                if next_pos == 2:  # Skip colon
                    next_pos = 3
                event.widget.icursor(next_pos)
                
                # Trigger calculation
                self.root.after(100, self.calculate_note_total_hours)
            
            return 'break'
    
    def on_time_click(self, event, time_var):
        """Handle mouse click in time field"""
        # Position cursor appropriately, avoiding the colon
        cursor_pos = event.widget.index(tk.INSERT)
        if cursor_pos == 2:  # Clicked on colon
            event.widget.icursor(3)  # Move to after colon
    
    def on_time_focus_out_smart(self, time_var):
        """Smart complete partial times when leaving field"""
        current = time_var.get()
        
        # Check if it's a partial time that needs completion
        if current and current != "__:__":
            # Extract digits only
            digits = ''.join(c for c in current if c.isdigit())
            
            if len(digits) == 1:
                # Single digit like "8" becomes "8:00"
                time_var.set(f"{digits}:00")
            elif len(digits) == 2:
                # Two digits like "12" becomes "12:00"
                time_var.set(f"{digits}:00")
            elif len(digits) == 3:
                # Three digits like "830" becomes "8:30"
                time_var.set(f"{digits[0]}:{digits[1:3]}")
            elif len(digits) >= 4:
                # Four+ digits like "1030" becomes "10:30"
                time_var.set(f"{digits[0:2]}:{digits[2:4]}")
        
        # Validate and fix time format
        self.validate_time_format(time_var)
        
        # Trigger calculation after smart completion
        self.root.after(100, self.calculate_note_total_hours)
    
    def validate_time_format(self, time_var):
        """Validate time format and fix invalid minutes"""
        current = time_var.get()
        
        if current and current != "__:__" and ':' in current:
            try:
                parts = current.split(':')
                if len(parts) == 2:
                    hours = parts[0]
                    minutes = parts[1]
                    
                    # Convert to integers for validation
                    hour_int = int(hours) if hours.isdigit() else 0
                    min_int = int(minutes) if minutes.isdigit() else 0
                    
                    # Validate hours (0-23)
                    if hour_int > 23:
                        hour_int = 23
                    elif hour_int < 0:
                        hour_int = 0
                    
                    # Validate minutes (0-59)
                    if min_int > 59:
                        min_int = 59
                    elif min_int < 0:
                        min_int = 0
                    
                    # Set corrected time
                    corrected_time = f"{hour_int:02d}:{min_int:02d}"
                    if current != corrected_time:
                        time_var.set(corrected_time)
            except ValueError:
                # If parsing fails, keep original value
                pass
    
    def calculate_note_total_hours(self, *args):
        """Calculate total hours from time in and time out"""
        try:
            time_in_str = self.note_time_in_time.get().strip()
            time_in_ampm = self.note_time_in_ampm.get()
            time_out_str = self.note_time_out_time.get().strip()
            time_out_ampm = self.note_time_out_ampm.get()
            
            # Check if times are empty or placeholder
            if (not time_in_str or not time_out_str or 
                time_in_str == "__:__" or time_out_str == "__:__" or
                "_" in time_in_str or "_" in time_out_str):
                self.note_total_hours.set("")
                return
            
            # Parse times directly (they should already be in HH:MM format)
            time_in_full = f"{time_in_str} {time_in_ampm}"
            time_out_full = f"{time_out_str} {time_out_ampm}"
            
            from datetime import datetime
            time_in_obj = datetime.strptime(time_in_full, "%I:%M %p")
            time_out_obj = datetime.strptime(time_out_full, "%I:%M %p")
            
            # Calculate difference in minutes
            time_diff = time_out_obj - time_in_obj
            total_minutes = int(time_diff.total_seconds() / 60)
            
            # Handle overnight shift
            if total_minutes < 0:
                total_minutes += 24 * 60
            
            # Convert to hours and minutes format
            hours = total_minutes // 60
            minutes = total_minutes % 60
            self.note_total_hours.set(f"{hours}:{minutes:02d}")
            
        except Exception as e:
            # For debugging - remove this line in production
            print(f"Error calculating total hours: {e}")
            self.note_total_hours.set("")
    
    def calculate_batch_total_hours(self, form_vars):
        """Calculate total hours for a batch form"""
        try:
            time_in_str = form_vars['time_in_time'].get().strip()
            time_in_ampm = form_vars['time_in_ampm'].get()
            time_out_str = form_vars['time_out_time'].get().strip()
            time_out_ampm = form_vars['time_out_ampm'].get()
            
            # Check if times are empty or placeholder
            if (not time_in_str or not time_out_str or 
                time_in_str == "__:__" or time_out_str == "__:__" or
                "_" in time_in_str or "_" in time_out_str):
                form_vars['total_hours'].set("")
                return
            
            # Parse times directly (they should already be in HH:MM format)
            time_in_full = f"{time_in_str} {time_in_ampm}"
            time_out_full = f"{time_out_str} {time_out_ampm}"
            
            from datetime import datetime
            time_in_obj = datetime.strptime(time_in_full, "%I:%M %p")
            time_out_obj = datetime.strptime(time_out_full, "%I:%M %p")
            
            # Calculate difference in minutes
            time_diff = time_out_obj - time_in_obj
            total_minutes = int(time_diff.total_seconds() / 60)
            
            # Handle overnight shift
            if total_minutes < 0:
                total_minutes += 24 * 60
            
            # Convert to hours and minutes format
            hours = total_minutes // 60
            minutes = total_minutes % 60
            form_vars['total_hours'].set(f"{hours}:{minutes:02d}")
            
        except Exception as e:
            # For debugging - remove this line in production
            print(f"Error calculating batch total hours: {e}")
            form_vars['total_hours'].set("")
    
    def on_time_focus_out_smart_batch(self, time_var, form_vars):
        """Smart complete partial times when leaving field in batch forms"""
        current = time_var.get()
        
        # Check if it's a partial time that needs completion
        if current and current != "__:__":
            # Extract digits only
            digits = ''.join(c for c in current if c.isdigit())
            
            if len(digits) == 1:
                # Single digit like "8" becomes "8:00"
                time_var.set(f"{digits}:00")
            elif len(digits) == 2:
                # Two digits like "12" becomes "12:00"
                time_var.set(f"{digits}:00")
            elif len(digits) == 3:
                # Three digits like "830" becomes "8:30"
                time_var.set(f"{digits[0]}:{digits[1:3]}")
            elif len(digits) >= 4:
                # Four+ digits like "1030" becomes "10:30"
                time_var.set(f"{digits[0:2]}:{digits[2:4]}")
        
        # Validate and fix time format
        self.validate_time_format(time_var)
        
        # Trigger calculation after smart completion
        self.root.after(100, lambda: self.calculate_batch_total_hours(form_vars))
    
    def toggle_batch_mode(self):
        """Enable/disable batch mode and show/hide additional forms"""
        if self.batch_mode.get():
            self.batch_count_entry.configure(state='normal')
            self.update_batch_forms()
        else:
            self.batch_count_entry.configure(state='disabled')
            self.batch_count.set("1")
            
            # Hide the batch container when disabling batch mode
            self.batch_forms_container.pack_forget()
            
            self.clear_batch_forms()
            # Ensure Add Note button is visible in main form when batch mode is disabled
            if hasattr(self, 'add_note_button'):
                self.add_note_button.grid(row=6, column=1, sticky='w', padx=5, pady=10)
    
    def refresh_notes_layout(self):
        """Force immediate refresh of the notes tab layout"""
        if hasattr(self, 'notes_frame'):
            # Simple refresh for CTkScrollableFrame
            self.notes_frame.update_idletasks()
            self.root.update_idletasks()
    
    def force_aggressive_layout_reset(self):
        """Simple layout refresh"""
        if hasattr(self, 'notes_frame'):
            # Simple refresh approach
            self.notes_frame.update_idletasks()
    
    def force_comprehensive_refresh(self):
        """Force comprehensive refresh for large batch operations"""
        try:
            # Update the scrollable frame
            if hasattr(self, 'notes_frame'):
                self.notes_frame.update_idletasks()
                self.notes_frame.update()
            
            # Update the main root window
            self.root.update_idletasks()
            self.root.update()
            
            # Force the tabview to refresh
            if hasattr(self, 'tabview'):
                self.tabview.update_idletasks()
                self.tabview.update()
            
            # Small delay to allow GUI to catch up
            self.root.after(50, self._final_refresh_step)
            
        except Exception as e:
            # If refresh fails, just log it and continue
            print(f"Refresh error: {e}")
    
    def _final_refresh_step(self):
        """Final refresh step after delay"""
        try:
            if hasattr(self, 'notes_frame'):
                self.notes_frame.update_idletasks()
            self.root.update_idletasks()
        except Exception:
            pass  # Ignore any refresh errors
    
    def on_batch_count_change(self, event=None):
        """Handle batch count changes from user input"""
        if self.batch_mode.get():
            # If this is a FocusOut event and field is empty, set to 1
            if event and event.type == '10':  # FocusOut event
                count_str = self.batch_count.get().strip()
                if not count_str:
                    self.batch_count.set("1")
            self.update_batch_forms()
    
    def update_batch_forms(self, *args):
        """Create or update batch forms based on count"""
        try:
            count_str = self.batch_count.get().strip()
            if not count_str:  # If empty, don't auto-set to 1, just return
                return
            
            count = int(count_str)
            if count < 1:
                count = 1
                self.batch_count.set("1")
            elif count > 10:  # Limit to 10 total (9 batch + 1 main)
                count = 10
                self.batch_count.set("10")
        except ValueError:
            # Don't auto-correct invalid input, let user finish typing
            return
        
        # Check if we're going to count = 1 (special handling needed)
        going_to_one = (count == 1)
        
        # Preserve existing batch form data before clearing
        existing_data = []
        if hasattr(self, 'batch_forms'):
            for i, batch_form in enumerate(self.batch_forms):
                if 'vars' in batch_form:
                    form_data = self.get_batch_form_data(batch_form)
                    existing_data.append(form_data)
        
        # Clear existing batch forms
        self.clear_batch_forms()
        
        if count > 1:
            # Show the batch container BEFORE the Current Notes section
            self.batch_forms_container.pack(fill='x', padx=10, pady=5, before=self.notes_list_frame)
            
            # Hide the Add Note button from main form when in batch mode
            self.add_note_button.grid_remove()
            
            # Get current employee from main form to carry over
            current_employee = self.note_employee.get()
            
            # Create additional forms (count - 1 since we already have the main form)
            self.batch_forms = []
            for i in range(count - 1):
                is_last_form = (i == count - 2)  # Last batch form
                batch_form = self.create_batch_form(i + 2, current_employee, is_last_form)  # Start from 2 since main is 1
                
                # Restore data if it exists for this form
                if i < len(existing_data) and existing_data[i]:
                    self.restore_batch_form_data(batch_form, existing_data[i])
                
                self.batch_forms.append(batch_form)
        else:
            # Hide the batch container when empty (count = 1)
            self.batch_forms_container.pack_forget()
            
            # Show the Add Note button in main form when not in batch mode
            self.add_note_button.grid(row=6, column=1, sticky='w', padx=5, pady=10)
    
    def clear_batch_forms(self):
        """Remove all batch forms"""
        if hasattr(self, 'batch_forms'):
            for form in self.batch_forms:
                form['frame'].destroy()
            self.batch_forms = []
        
        # Show the Add Note button in main form when batch forms are cleared
        if hasattr(self, 'add_note_button'):
            self.add_note_button.grid(row=6, column=1, sticky='w', padx=5, pady=10)
        
        # Force immediate layout refresh using stored references
        self.refresh_notes_layout()
    
    def create_batch_form(self, form_number, default_employee="", is_last_form=False):
        """Create a single batch form"""
        # Create batch form frame in the dedicated container
        batch_frame = ctk.CTkFrame(self.batch_forms_container, corner_radius=8)
        batch_frame.pack(fill='x', padx=10, pady=5)
        
        # Create form fields
        form_vars = {}
        
        # Employee selection
        ctk.CTkLabel(batch_frame, text="Employee:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        form_vars['employee'] = tk.StringVar(value=default_employee)  # Set default employee
        employee_combo = ctk.CTkComboBox(batch_frame, variable=form_vars['employee'], 
                                     width=300, state='normal')
        if hasattr(self, 'all_employees'):
            employee_combo.configure(values=self.all_employees)
        employee_combo.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        # Disable scroll wheel to prevent accidental changes
        employee_combo.bind('<MouseWheel>', lambda e: 'break')
        
        # Date selection
        ctk.CTkLabel(batch_frame, text="Date:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        form_vars['date'] = tk.StringVar()
        date_combo = ctk.CTkComboBox(batch_frame, variable=form_vars['date'], 
                                 width=200, state='readonly')
        if hasattr(self, 'date_combo'):
            try:
                values = self.date_combo.cget('values')
                if values:
                    date_combo.configure(values=values)
            except:
                pass
        date_combo.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        # Disable scroll wheel to prevent accidental changes
        date_combo.bind('<MouseWheel>', lambda e: 'break')
        
        # Override type
        ctk.CTkLabel(batch_frame, text="Override Type:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        # Inherit override type from main form if set, otherwise default to "Missing Punch"
        main_form_type = self.note_type.get() if hasattr(self, 'note_type') and self.note_type.get() else "Missing Punch"
        form_vars['type'] = tk.StringVar(value=main_form_type)
        type_combo = ctk.CTkComboBox(batch_frame, variable=form_vars['type'], width=200, 
                                state='readonly', values=('Time Change', 'Break Change', 'Missing Punch', 'Sick Day', 'Misc', 'ROE'),
                                command=lambda choice: self.on_batch_note_type_change(time_frame, form_vars['type']))
        type_combo.grid(row=2, column=1, sticky='w', padx=5, pady=2)
        # Disable scroll wheel to prevent accidental changes
        type_combo.bind('<MouseWheel>', lambda e: 'break')
        
        # Time fields
        time_frame = ctk.CTkFrame(batch_frame, fg_color="transparent", border_width=1, border_color="gray")
        time_frame.grid(row=3, column=0, columnspan=2, sticky='ew', pady=5)
        
        # Time In
        ctk.CTkLabel(time_frame, text="Time In:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        time_in_frame = ctk.CTkFrame(time_frame, fg_color="transparent")
        time_in_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        form_vars['time_in_time'] = tk.StringVar(value="__:__")
        time_in_entry = ctk.CTkEntry(time_in_frame, textvariable=form_vars['time_in_time'], width=80)
        time_in_entry.pack(side='left', padx=(0,2))
        time_in_entry.bind('<KeyPress>', lambda e: self.on_time_key_press(e, form_vars['time_in_time']))
        time_in_entry.bind('<FocusOut>', lambda e: self.on_time_focus_out_smart_batch(form_vars['time_in_time'], form_vars))
        time_in_entry.bind('<Button-1>', lambda e: self.on_time_click(e, form_vars['time_in_time']))
        
        form_vars['time_in_ampm'] = tk.StringVar(value="AM")
        time_in_ampm = ctk.CTkComboBox(time_in_frame, variable=form_vars['time_in_ampm'], width=50, values=['AM', 'PM'], state='readonly',
                                     command=lambda choice: self.root.after(100, lambda: self.calculate_batch_total_hours(form_vars)))
        time_in_ampm.pack(side='left')
        # Disable scroll wheel to prevent accidental changes
        time_in_ampm.bind('<MouseWheel>', lambda e: 'break')
        
        # Time Out
        ctk.CTkLabel(time_frame, text="Time Out:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        time_out_frame = ctk.CTkFrame(time_frame, fg_color="transparent")
        time_out_frame.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        form_vars['time_out_time'] = tk.StringVar(value="__:__")
        time_out_entry = ctk.CTkEntry(time_out_frame, textvariable=form_vars['time_out_time'], width=80)
        time_out_entry.pack(side='left', padx=(0,2))
        time_out_entry.bind('<KeyPress>', lambda e: self.on_time_key_press(e, form_vars['time_out_time']))
        time_out_entry.bind('<FocusOut>', lambda e: self.on_time_focus_out_smart_batch(form_vars['time_out_time'], form_vars))
        time_out_entry.bind('<Button-1>', lambda e: self.on_time_click(e, form_vars['time_out_time']))
        
        form_vars['time_out_ampm'] = tk.StringVar(value="PM")
        time_out_ampm = ctk.CTkComboBox(time_out_frame, variable=form_vars['time_out_ampm'], width=50, values=['AM', 'PM'], state='readonly',
                                      command=lambda choice: self.root.after(100, lambda: self.calculate_batch_total_hours(form_vars)))
        time_out_ampm.pack(side='left')
        # Disable scroll wheel to prevent accidental changes
        time_out_ampm.bind('<MouseWheel>', lambda e: 'break')
        
        # Break Minutes
        ctk.CTkLabel(time_frame, text="Break Minutes:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        form_vars['break_minutes'] = tk.StringVar()
        ctk.CTkEntry(time_frame, textvariable=form_vars['break_minutes'], width=80).grid(row=2, column=1, sticky='w', padx=5, pady=2)
        
        # Total Hours (calculated field)
        ctk.CTkLabel(time_frame, text="Total Hours:").grid(row=3, column=0, sticky='w', padx=5, pady=2)
        form_vars['total_hours'] = tk.StringVar()
        total_hours_entry = ctk.CTkEntry(time_frame, textvariable=form_vars['total_hours'], width=80, state='readonly')
        total_hours_entry.grid(row=3, column=1, sticky='w', padx=5, pady=2)
        
        # Note text
        ctk.CTkLabel(batch_frame, text="Note:").grid(row=4, column=0, sticky='nw', padx=5, pady=2)
        form_vars['note'] = tk.StringVar()
        ctk.CTkEntry(batch_frame, textvariable=form_vars['note'], width=400).grid(row=4, column=1, sticky='ew', padx=5, pady=2)
        
        # Add the Add Note button to the last batch form
        if is_last_form:
            ctk.CTkButton(batch_frame, text="Add Note", command=self.add_note).grid(row=5, column=1, sticky='w', padx=5, pady=10)
        
        return {
            'frame': batch_frame,
            'vars': form_vars,
            'type_combo': type_combo,
            'time_frame': time_frame
        }
    
    def add_note(self):
        """Add notes from main form and any batch forms"""
        # Check if batch mode is enabled and employee is required
        if self.batch_mode.get() and not self.note_employee.get():
            messagebox.showwarning("Employee Required", 
                                 "Please select an employee before using batch mode.\n" +
                                 "This ensures all batch notes use the same employee for efficient workflow.")
            return
        
        # Collect all forms (main + batch)
        all_forms = [self.get_main_form_data()]
        
        if self.batch_mode.get() and hasattr(self, 'batch_forms'):
            for batch_form in self.batch_forms:
                batch_data = self.get_batch_form_data(batch_form)
                # Check if form has meaningful data before adding
                has_data = (batch_data['employee'] or 
                           batch_data['date'] or 
                           batch_data['type'] or
                           (batch_data['time_in_time'] and batch_data['time_in_time'] != "__:__") or
                           (batch_data['time_out_time'] and batch_data['time_out_time'] != "__:__"))
                if has_data:
                    all_forms.append(batch_data)
        
        notes_added = 0
        validation_errors = []
        
        for form_idx, form_data in enumerate(all_forms):
            form_name = f"Main Form" if form_idx == 0 else f"Note #{form_idx + 1}"
            
            # Validate each form
            missing_fields = self.validate_form_data(form_data)
            if missing_fields:
                validation_errors.append(f"{form_name}: {', '.join(missing_fields)}")
                continue
            
            # Convert display names to internal format
            type_mapping = {
                'Time Change': 'time_override',
                'Break Change': 'break_override',
                'Missing Punch': 'missing_punch_override',
                'Sick Day': 'sick_day',
                'Misc': 'misc',
                'ROE': 'roe'
            }
            
            internal_type = type_mapping.get(form_data['type'], form_data['type'])
            
            # Create note
            note = {
                "employee": form_data['employee'],
                "date": form_data['date'],
                "type": internal_type,
                "note": form_data['note']
            }
            
            # Add additional fields based on type
            if internal_type in ['time_override', 'missing_punch_override']:
                time_in = f"{form_data['time_in_time']} {form_data['time_in_ampm']}"
                time_out = f"{form_data['time_out_time']} {form_data['time_out_ampm']}"
                note["time_in"] = time_in
                note["time_out"] = time_out
                if form_data['break_minutes']:
                    note["break_minutes"] = int(form_data['break_minutes'])
            elif internal_type == 'break_override':
                if form_data['break_minutes']:
                    note["value"] = int(form_data['break_minutes'])
            
            self.current_notes["notes"].append(note)
            notes_added += 1
        
        # Show validation errors if any
        if validation_errors:
            messagebox.showwarning("Validation Errors", 
                                 "Please fix the following issues:\n\n" + 
                                 "\n".join(validation_errors))
            return
        
        if notes_added == 0:
            messagebox.showwarning("No Notes Added", "Please fill in at least one form")
            return
        
        # Clear form after successful add
        self.note_employee.set("")
        self.note_date.set("")
        self.note_type.set("Missing Punch")  # Reset to default
        self.note_text.set("")
        self.note_time_in_time.set("__:__")
        self.note_time_in_ampm.set("AM")
        self.note_time_out_time.set("__:__")
        self.note_time_out_ampm.set("PM")
        self.note_break_minutes.set("")
        self.note_total_hours.set("")
        
        # Clear batch forms
        self.clear_batch_forms()
        self.batch_mode.set(False)
        self.batch_count.set("1")
        self.batch_count_entry.configure(state='disabled')
        
        # Trigger note type change to show correct fields
        self.on_note_type_change("Missing Punch")
        
        # Refresh the notes display and force layout update
        self.refresh_notes_display()
        
        # Force a comprehensive layout refresh for large batch operations
        if notes_added > 4:
            self.force_comprehensive_refresh()
        
        # Show success message
        if notes_added > 1:
            messagebox.showinfo("Batch Notes Added", f"Successfully added {notes_added} notes")
        
    def get_main_form_data(self):
        """Get data from main form"""
        return {
            'employee': self.note_employee.get(),
            'date': self.note_date.get(),
            'type': self.note_type.get(),
            'time_in_time': self.note_time_in_time.get() if self.note_time_in_time.get() != "__:__" else "",
            'time_in_ampm': self.note_time_in_ampm.get(),
            'time_out_time': self.note_time_out_time.get() if self.note_time_out_time.get() != "__:__" else "",
            'time_out_ampm': self.note_time_out_ampm.get(),
            'break_minutes': self.note_break_minutes.get(),
            'note': self.note_text.get()
        }
    
    def get_batch_form_data(self, batch_form):
        """Get data from a batch form"""
        vars_dict = batch_form['vars']
        
        # Always return data (even if empty) for preservation purposes
        return {
            'employee': vars_dict['employee'].get(),
            'date': vars_dict['date'].get(),
            'type': vars_dict['type'].get(),
            'time_in_time': vars_dict['time_in_time'].get(),
            'time_in_ampm': vars_dict['time_in_ampm'].get(),
            'time_out_time': vars_dict['time_out_time'].get(),
            'time_out_ampm': vars_dict['time_out_ampm'].get(),
            'break_minutes': vars_dict['break_minutes'].get(),
            'note': vars_dict['note'].get()
        }
    
    def restore_batch_form_data(self, batch_form, form_data):
        """Restore data to a batch form"""
        vars_dict = batch_form['vars']
        
        # Restore all the data
        vars_dict['employee'].set(form_data.get('employee', ''))
        vars_dict['date'].set(form_data.get('date', ''))
        vars_dict['type'].set(form_data.get('type', ''))
        vars_dict['time_in_time'].set(form_data.get('time_in_time', '__:__'))
        vars_dict['time_in_ampm'].set(form_data.get('time_in_ampm', 'AM'))
        vars_dict['time_out_time'].set(form_data.get('time_out_time', '__:__'))
        vars_dict['time_out_ampm'].set(form_data.get('time_out_ampm', 'PM'))
        vars_dict['break_minutes'].set(form_data.get('break_minutes', ''))
        vars_dict['note'].set(form_data.get('note', ''))
        
        # Trigger the type change to show correct fields
        if 'time_frame' in batch_form and form_data.get('type'):
            self.on_batch_note_type_change(batch_form['time_frame'], vars_dict['type'])
    
    def validate_form_data(self, form_data):
        """Validate form data and return list of missing required fields"""
        missing_fields = []
        
        if not form_data['employee']:
            missing_fields.append("Employee")
        if not form_data['date']:
            missing_fields.append("Date")
        if not form_data['type']:
            missing_fields.append("Override Type")
        
        # Additional validation based on note type
        note_type = form_data['type']
        if note_type in ['Time Change', 'Missing Punch']:
            if not form_data['time_in_time']:
                missing_fields.append("Time In")
            if not form_data['time_out_time']:
                missing_fields.append("Time Out")
        elif note_type == 'Break Change':
            if not form_data['break_minutes']:
                missing_fields.append("Break Minutes")
        
        return missing_fields
        
        # Clear form after successful add
        self.note_employee.set("")
        self.note_date.set("")
        self.note_type.set("")
        self.note_text.set("")
        self.note_time_in_time.set("__:__")
        self.note_time_in_ampm.set("AM")
        self.note_time_out_time.set("__:__")
        self.note_time_out_ampm.set("PM")
        self.note_break_minutes.set("")
        self.note_total_hours.set("")
        
        # Clear batch forms
        self.clear_batch_forms()
        self.batch_mode.set(False)
        self.batch_count.set("1")
        self.batch_count_entry.configure(state='disabled')
        
        # Show success message
        if notes_added > 1:
            messagebox.showinfo("Batch Notes Added", f"Successfully added {notes_added} notes")
    
    def edit_note(self):
        """Edit selected note"""
        selection = self.notes_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a note to edit")
            return
        
        # Get the index of the selected item
        index = self.notes_tree.index(selection[0])
        note = self.current_notes["notes"][index]
        
        # Populate the form with the selected note's data
        self.note_employee.set(note.get("employee", ""))
        self.note_date.set(note.get("date", ""))
        
        # Map internal types to display names
        type_mapping = {
            'time_override': 'Time Change',
            'break_override': 'Break Change',
            'missing_punch_override': 'Missing Punch',
            'sick_day': 'Sick Day'
        }
        display_type = type_mapping.get(note.get("type", ""), note.get("type", ""))
        self.note_type.set(display_type)
        
        # Trigger the type change to show appropriate fields
        self.on_note_type_change(display_type)
        
        # Populate type-specific fields (check both internal types and display types)
        if (note.get("type") in ['time_override', 'missing_punch_override'] or 
            (note.get("type") == 'Missing Punch' and ('time_in' in note or 'time_out' in note))):
            time_in = note.get('time_in', '')
            time_out = note.get('time_out', '')
            
            # Parse time_in (e.g., "8:00 AM")
            if time_in and ' ' in time_in:
                time_part, ampm_part = time_in.rsplit(' ', 1)
                self.note_time_in_time.set(time_part)
                self.note_time_in_ampm.set(ampm_part)
            
            # Parse time_out (e.g., "5:00 PM")
            if time_out and ' ' in time_out:
                time_part, ampm_part = time_out.rsplit(' ', 1)
                self.note_time_out_time.set(time_part)
                self.note_time_out_ampm.set(ampm_part)
            
            # Set break minutes if available
            if 'break_minutes' in note:
                self.note_break_minutes.set(str(note['break_minutes']))
                
        elif note.get("type") == 'break_override':
            if 'value' in note:
                self.note_break_minutes.set(str(note['value']))
        
        # Set the note text
        self.note_text.set(note.get("note", ""))
        
        # Calculate total hours for the loaded note
        self.calculate_note_total_hours()
        
        # Remove the original note (it will be re-added when user clicks "Add Note")
        self.current_notes["notes"].pop(index)
        self.refresh_notes_display()
        
        messagebox.showinfo("Edit Mode", "Note loaded for editing. Make your changes and click 'Add Note' to save.")
    
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
            # Check if note has explicit details field first (from smart parsing)
            if note.get('details'):
                details = note.get('details')
            elif note["type"] in ["time_override", "missing_punch_override"]:
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
        
        # Use the last used notes directory
        initial_dir = getattr(self, 'last_notes_directory', str(self.notes_dir))
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=initial_dir
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(self.current_notes, f, indent=2)
                
                # Track the current notes file path
                self.current_notes_file_path = filename
                
                # Save the directory for next time
                self.last_notes_directory = str(Path(filename).parent)
                self.save_settings()
                
                messagebox.showinfo("Success", "Notes saved successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save notes: {str(e)}")
    
    def load_notes(self):
        """Load notes from file"""
        # Use the last used notes directory
        initial_dir = getattr(self, 'last_notes_directory', str(self.notes_dir))
        
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            initialdir=initial_dir
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    self.current_notes = json.load(f)
                
                # Track the current notes file path
                self.current_notes_file_path = filename
                
                # Save the directory for next time
                self.last_notes_directory = str(Path(filename).parent)
                self.save_settings()
                
                self.pay_period.set(self.current_notes.get("pay_period", ""))
                self.refresh_notes_display()
                messagebox.showinfo("Success", "Notes loaded successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load notes: {str(e)}")
    
    # Processing Methods
    def select_pdf_file(self):
        """Select a single PDF file for processing"""
        file_path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file_path:
            self.selected_pdf_file = file_path
            filename = Path(file_path).name
            self.selected_file_label.configure(text=f"Selected: {filename}", text_color="white")
    
    def clear_pdf_file(self):
        """Clear the selected PDF file"""
        self.selected_pdf_file = None
        self.selected_file_label.configure(text="No file selected", text_color="gray")
    
    def browse_output(self):
        """Browse for output directory"""
        # Use the last selected directory as the initial directory
        initial_dir = getattr(self, 'last_output_directory', str(Path.home() / "Documents" / "Timecard_Output"))
        
        directory = filedialog.askdirectory(initialdir=initial_dir)
        if directory:
            self.output_path.set(directory)
            # Save this as the new default
            self.last_output_directory = directory
            self.save_settings()
    
    def process_files(self):
        """Process the selected file"""
        if not self.current_profile:
            messagebox.showwarning("Warning", "Please load a profile first")
            return
        
        if not self.selected_pdf_file:
            messagebox.showwarning("Warning", "Please select a PDF file to process")
            return
        
        # Get the user-selected output directory
        output_dir = self.output_path.get().strip()
        if not output_dir:
            messagebox.showwarning("Warning", "Please select an output directory")
            return
        
        # Create a copy of the profile with the user-selected output directory
        temp_profile = self.current_profile.copy()
        temp_profile['output_directory'] = output_dir
        
        # Save temporary profile with updated output directory
        temp_config = "temp_config.json"
        with open(temp_config, 'w') as f:
            json.dump(temp_profile, f, indent=2)
        
        # Save current notes temporarily if any
        temp_notes = None
        if self.current_notes["notes"]:
            # If we have a current notes file path, use it directly instead of creating a temp file
            if self.current_notes_file_path and os.path.exists(self.current_notes_file_path):
                temp_notes = self.current_notes_file_path
                print(f"Using existing notes file: {temp_notes}")
            else:
                # Fallback: create temp file from in-memory notes
                temp_notes = "temp_notes.json"
                with open(temp_notes, 'w') as f:
                    json.dump(self.current_notes, f, indent=2)
                print(f"Created temporary notes file: {temp_notes}")
        
        try:
            self.progress_var.set("Processing...")
            self.progress_bar.set(0.1)
            
            # Import processor directly instead of using subprocess
            try:
                # Handle PyInstaller bundled environment
                if hasattr(sys, '_MEIPASS'):
                    # Running as PyInstaller bundle - add the bundled path
                    processor_path = os.path.join(sys._MEIPASS, 'timecard_processing')
                    if processor_path not in sys.path:
                        sys.path.insert(0, processor_path)
                
                # Import the processor module
                from timecard_processing.processor import UniversalTimeCardProcessor
                
                # Process the single file using direct import
                filename = Path(self.selected_pdf_file).name
                self.progress_var.set(f"Processing file: {filename}")
                self.root.update()  # Update GUI to show progress
                
                # Create processor instance and process file directly
                processor = UniversalTimeCardProcessor(temp_config, temp_notes)
                processor.process_file(self.selected_pdf_file)
                
                self.progress_bar.set(1.0)
                self.progress_var.set("Processing complete!")
                messagebox.showinfo("Success", f"File processed successfully!\n\nOutput saved to: {output_dir}")
                
            except ImportError as import_error:
                # Fallback to subprocess method (for development environment)
                self.progress_var.set("Using fallback method...")
                
                filename = Path(self.selected_pdf_file).name
                self.progress_var.set(f"Processing file: {filename}")
                self.root.update()
                
                # Build command for subprocess
                cmd = [sys.executable, "timecard_processing/processor.py", temp_config, self.selected_pdf_file]
                if temp_notes:
                    cmd.extend(["--notes", temp_notes])
                
                # Run processor as subprocess
                result = subprocess.run(cmd, capture_output=True, text=True)
                
                if result.returncode != 0:
                    raise Exception(f"Processing failed for {filename}:\n{result.stderr}")
                
                self.progress_bar.set(1.0)
                self.progress_var.set("Processing complete!")
                messagebox.showinfo("Success", f"File processed successfully!\n\nOutput saved to: {output_dir}")
            
        except Exception as e:
            self.progress_bar.set(0)
            self.progress_var.set("Processing failed")
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
        
        finally:
            # Clean up temporary files
            try:
                if os.path.exists(temp_config):
                    os.remove(temp_config)
                # Only remove temp_notes if it's not our current notes file
                if temp_notes and os.path.exists(temp_notes) and temp_notes != self.current_notes_file_path:
                    os.remove(temp_notes)
            except Exception:
                pass  # Ignore cleanup errors

    def filter_employee_dropdown(self, *args):
        """Filter employee dropdown based on typed text"""
        if not hasattr(self, 'all_employees') or not self.all_employees or not self.current_profile:
            return
        
        typed_text = self.note_employee.get().lower()
        if not typed_text:
            # Show all employees if nothing typed
            self.employee_combo.configure(values=self.all_employees)
        else:
            # Filter employees that contain the typed text
            filtered = [emp for emp in self.all_employees if typed_text in emp.lower()]
            self.employee_combo.configure(values=filtered)
    
    def on_main_employee_change(self, *args):
        """Update batch forms when main employee changes and control batch mode availability"""
        current_employee = self.note_employee.get()
        
        # Enable/disable batch mode based on employee selection
        if hasattr(self, 'batch_mode_checkbox'):
            if current_employee:
                self.batch_mode_checkbox.configure(state='normal')
            else:
                self.batch_mode_checkbox.configure(state='disabled')
                # If no employee selected, disable batch mode
                if self.batch_mode.get():
                    self.batch_mode.set(False)
                    self.toggle_batch_mode()
        
        # Update all batch forms with the new employee
        if self.batch_mode.get() and hasattr(self, 'batch_forms'):
            for batch_form in self.batch_forms:
                if 'vars' in batch_form and 'employee' in batch_form['vars']:
                    batch_form['vars']['employee'].set(current_employee)
    
    def open_employee_dropdown(self, event):
        """Open employee dropdown when Down arrow is pressed"""
        self.employee_combo.event_generate('<Button-1>')
        return 'break'  # Prevent default behavior
    
    def on_employee_click(self, event):
        """Handle employee combobox click - clear selection but allow dropdown"""
        # Only clear selection if clicking on the text area, not the dropdown arrow
        widget_width = self.employee_combo.winfo_width()
        click_x = event.x
        # If click is in the last 20 pixels (arrow area), don't clear selection
        if click_x < widget_width - 20:
            self.employee_combo.selection_clear()
    
    def on_tab_changed(self, event):
        """Remove focus from all widgets when tab changes"""
        # Simply focus the tabview itself to remove focus from any input fields
        self.tabview.focus_set()
    
    def cleanup_old_settings(self):
        """Remove old settings file from current directory if it exists"""
        old_settings_file = Path("timecard_gui_settings.json")
        if old_settings_file.exists():
            try:
                # Try to migrate settings first
                with open(old_settings_file, 'r') as f:
                    old_settings = json.load(f)
                    # Save to new location
                    with open(self.settings_file, 'w') as new_f:
                        json.dump(old_settings, new_f, indent=2)
                # Remove old file
                old_settings_file.unlink()
            except Exception:
                # If migration fails, just remove the old file
                try:
                    old_settings_file.unlink()
                except Exception:
                    pass  # Ignore if we can't remove it
    
    def load_settings(self):
        """Load persistent settings"""
        try:
            if Path(self.settings_file).exists():
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    self.last_output_directory = settings.get('last_output_directory', str(Path.home() / "Documents" / "Timecard_Output"))
                    self.last_notes_directory = settings.get('last_notes_directory', str(self.notes_dir))
            else:
                self.last_output_directory = str(Path.home() / "Documents" / "Timecard_Output")
                self.last_notes_directory = str(self.notes_dir)
        except Exception as e:
            print(f"Error loading settings: {e}")
            self.last_output_directory = str(Path.home() / "Documents" / "Timecard_Output")
            self.last_notes_directory = str(self.notes_dir)
        
        # Set the output path to the loaded directory
        if hasattr(self, 'output_path'):
            self.output_path.set(self.last_output_directory)
    
    def save_settings(self):
        """Save persistent settings"""
        try:
            settings = {
                'last_output_directory': getattr(self, 'last_output_directory', str(Path.home() / "Documents" / "Timecard_Output")),
                'last_notes_directory': getattr(self, 'last_notes_directory', str(self.notes_dir))
            }
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception as e:
            print(f"Error saving settings: {e}")
    
def main():
    root = ctk.CTk()
    app = XelifyGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 
