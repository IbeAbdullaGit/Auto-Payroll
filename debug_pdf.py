#!/usr/bin/env python3
"""
Debug script to examine PDF content
"""

import pdfplumber

def examine_pdf():
    with pdfplumber.open('input_files/TIME_CARD_REPORT_May_12-23-2025.pdf') as pdf:
        print(f"Total pages: {len(pdf.pages)}")
        
        # Get text from first page
        first_page_text = pdf.pages[0].extract_text()
        
        # Clean up the text (remove warning messages)
        lines = first_page_text.split('\n')
        clean_lines = [line for line in lines if not line.startswith('CropBox missing')]
        clean_text = '\n'.join(clean_lines)
        
        print("\n=== CLEANED FIRST PAGE TEXT ===")
        print(clean_text)
        
        print("\n=== LOOKING FOR EMPLOYEE PATTERNS ===")
        if "Name :" in clean_text:
            print("✓ Found 'Name :' pattern")
        if "Employee Group :" in clean_text:
            print("✓ Found 'Employee Group :' pattern")
        if "From :" in clean_text:
            print("✓ Found 'From :' pattern")

if __name__ == "__main__":
    examine_pdf() 