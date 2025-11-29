"""Test PDF font fix"""

from dotenv import load_dotenv
load_dotenv()

import os
from datetime import datetime
from pmo_helpers import process_excel_file
from spld_exact_format import generate_spld_exact_report

def test_pdf_fix():
    """Test that PDF text displays correctly"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🔧 Testing PDF Font Fix")
    print("=" * 60)
    
    # Read the file
    with open(excel_file, 'rb') as f:
        file_content = f.read()
    
    # Process the Excel file
    projects, error = process_excel_file(file_content, excel_file)
    
    if error:
        print(f"❌ Error: {error}")
        return
    
    print(f"✅ Found {len(projects)} projects")
    
    # Generate just first 3 projects for quick test
    test_projects = projects[:3]
    
    # Generate PDF with fixed encoding
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_path = f"Fixed_PDF_{timestamp}.pdf"
    
    print(f"\n📄 Generating PDF with fixed text encoding...")
    generate_spld_exact_report(test_projects, pdf_path)
    
    print(f"\n✅ Generated: {pdf_path}")
    print("\n📝 Check if text displays correctly (not as rectangles)")
    print("   - Project names should be readable")
    print("   - Project Objective section should show text")
    print("   - Activities should be readable")

if __name__ == "__main__":
    test_pdf_fix()
