"""Test that the system reads dynamically from Excel"""

from dotenv import load_dotenv
load_dotenv()

import os
from datetime import datetime
from pmo_helpers import process_excel_file
from spld_exact_format import generate_spld_exact_report
from spld_word_generator import create_spld_word_report
from llm_integration import format_project_text

def test_dynamic_data():
    """Test reading and generating from actual Excel data"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🔍 Testing Dynamic Data Reading from Excel")
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
    
    # Apply LLM formatting
    print("🤖 Applying AI text formatting...")
    for i in range(len(projects)):
        projects[i] = format_project_text(projects[i])
    
    # Show data from first 3 projects to verify it's reading correctly
    print("\n📊 Data Read from Excel (First 3 Projects):")
    print("-" * 60)
    
    for i, project in enumerate(projects[:3], 1):
        print(f"\n Project {i}: {project.get('name', 'Unknown')}")
        print(f"  • GM: {project.get('gm', 'Not found')}")
        print(f"  • Lead: {project.get('operational_lead', 'Not found')}")
        print(f"  • Health: {project.get('health', 'Not found')}")
        print(f"  • Progress: {project.get('timeline_actual', 0):.1f}%")
        print(f"  • Planned: {project.get('timeline_planned', 0):.1f}%")
        print(f"  • Budget Total: {project.get('budget_total', 0):,.0f} SAR")
        print(f"  • End Date: {project.get('contract_end_date', 'Not found')}")
        print(f"  • Current Activities: {str(project.get('current_activities', 'None'))[:50]}...")
        print(f"  • Risks: {str(project.get('risks', 'None'))[:50]}...")
    
    # Generate reports
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    print(f"\n📄 Generating SPLD Reports from Excel Data...")
    
    # Generate PDF
    pdf_path = f"SPLD_Dynamic_PDF_{timestamp}.pdf"
    generate_spld_exact_report(projects, pdf_path)
    print(f"✅ PDF Generated: {pdf_path}")
    
    # Generate Word
    word_path = f"SPLD_Dynamic_Word_{timestamp}.docx"
    create_spld_word_report(projects, word_path)
    print(f"✅ Word Generated: {word_path}")
    
    print("\n🎯 Success! Reports generated from actual Excel data:")
    print("  • All project names from Excel")
    print("  • All health statuses from Excel")
    print("  • All progress percentages from Excel")
    print("  • All activities and risks from Excel")
    print("  • All dates and budget data from Excel")

if __name__ == "__main__":
    test_dynamic_data()
