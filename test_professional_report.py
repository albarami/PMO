"""Test the professional PDF report generation"""

from dotenv import load_dotenv
load_dotenv()

import os
from datetime import datetime
from pmo_helpers import process_excel_file
from professional_pdf_generator import generate_professional_pdf
from llm_integration import format_project_text

def test_professional_report():
    """Test professional PDF generation"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("📊 Testing Professional PDF Report Generation")
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
    
    print(f"✅ Successfully extracted {len(projects)} projects")
    
    # Apply LLM formatting
    print("🤖 Applying AI text formatting...")
    for i in range(len(projects)):
        projects[i] = format_project_text(projects[i])
    
    # Generate professional PDF
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_path = f"PMO_Professional_Report_{timestamp}.pdf"
    
    print("\n📄 Generating Professional Executive Report...")
    generate_professional_pdf(projects, pdf_path)
    
    print(f"\n✅ Professional PDF generated: {pdf_path}")
    print("\nReport Features:")
    print("  • Professional cover page")
    print("  • Executive summary dashboard")
    print("  • Key metrics overview")
    print("  • Individual project detail pages")
    print("  • Risk and issue tracking")
    print("  • Professional styling with company colors")

if __name__ == "__main__":
    test_professional_report()
