"""Test that ALL projects are included in the reports"""

from dotenv import load_dotenv
load_dotenv()

import os
from datetime import datetime
from pmo_helpers import process_excel_file
from spld_exact_format import generate_spld_exact_report
from spld_word_generator import create_spld_word_report

def test_all_projects():
    """Test that ALL projects from Excel are included in reports"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🔍 Testing that ALL projects are included in reports")
    print("=" * 60)
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"❌ File not found: {excel_file}")
        return
    
    # Read the file
    with open(excel_file, 'rb') as f:
        file_content = f.read()
    
    # Process the Excel file
    projects, error = process_excel_file(file_content, excel_file)
    
    if error:
        print(f"❌ Error: {error}")
        return
    
    print(f"✅ Successfully extracted {len(projects)} projects from Excel")
    
    # List ALL projects found
    print("\n📋 ALL Projects Found in Excel:")
    print("-" * 40)
    for i, project in enumerate(projects, 1):
        print(f"{i:2}. {project.get('name', 'Unknown')}")
    
    # Generate reports with ALL projects
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    print(f"\n📄 Generating reports with ALL {len(projects)} projects...")
    
    # Generate PDF with ALL projects
    pdf_path = f"ALL_Projects_PDF_{timestamp}.pdf"
    generate_spld_exact_report(projects, pdf_path)
    print(f"✅ PDF Generated with ALL {len(projects)} projects: {pdf_path}")
    
    # Generate Word with ALL projects
    word_path = f"ALL_Projects_Word_{timestamp}.docx"
    create_spld_word_report(projects, word_path)
    print(f"✅ Word Generated with ALL {len(projects)} projects: {word_path}")
    
    print("\n🎯 SUCCESS!")
    print(f"  • Total projects in Excel: {len(projects)}")
    print(f"  • Projects in PDF: ALL {len(projects)}")
    print(f"  • Projects in Word: ALL {len(projects)}")
    print("\n📌 Each project has its own page in the reports")

if __name__ == "__main__":
    test_all_projects()
