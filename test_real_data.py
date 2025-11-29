"""Test script to process the real PMO Excel file"""

import pandas as pd
from pmo_helpers import process_excel_file
from pmo_report_generator import generate_pdf_report
from word_generator import create_word_report
from llm_integration import format_project_text
import os
from datetime import datetime

def test_real_file():
    """Test with the real PMO Project Tracker file"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ File not found: {excel_file}")
        return
    
    print(f"📂 Processing: {excel_file}")
    
    # Read the file
    with open(excel_file, 'rb') as f:
        file_content = f.read()
    
    # Process the Excel file
    projects, error = process_excel_file(file_content, excel_file)
    
    if error:
        print(f"❌ Error: {error}")
        return
    
    print(f"✅ Successfully extracted {len(projects)} projects")
    
    # Apply LLM formatting if available (optional)
    for i in range(len(projects)):
        projects[i] = format_project_text(projects[i])
    
    # Print summary of first project
    if projects:
        first = projects[0]
        print("\n📊 First Project Summary:")
        print(f"  - Name: {first.get('name', 'Unknown')}")
        print(f"  - Category: {first.get('category', 'Unknown')}")
        print(f"  - Status: {first.get('status', 'Unknown')}")
        print(f"  - Health: {first.get('health', 'Unknown')}")
        print(f"  - Budget Total: {first.get('budget_total', 0):,.2f} SAR")
        print(f"  - Timeline Actual: {first.get('timeline_actual', 0):.1f}%")
        print(f"  - Timeline Planned: {first.get('timeline_planned', 0):.1f}%")
        print(f"  - Schedule Variance: {first.get('schedule_variance', 0):+.1f}%")
        print(f"  - Days Remaining: {first.get('days_remaining', 0)}")
    
    # Generate test reports
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Generate PDF
    pdf_path = f"test_report_{timestamp}.pdf"
    generate_pdf_report(projects[:3], pdf_path)  # First 3 projects only for test
    print(f"\n✅ Generated PDF: {pdf_path}")
    
    # Generate Word
    word_path = f"test_report_{timestamp}.docx"
    create_word_report(projects[:3], word_path)  # First 3 projects only for test
    print(f"✅ Generated Word: {word_path}")
    
    print("\n🎉 Test completed successfully!")


if __name__ == "__main__":
    test_real_file()
