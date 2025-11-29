"""Test the complete PMO Report Generator with all output formats"""

from dotenv import load_dotenv
load_dotenv()

import os
from datetime import datetime
from pmo_helpers import process_excel_file
from pmo_report_generator import generate_pdf_report
from word_generator import create_word_report
from excel_generator import create_excel_report
from llm_integration import format_project_text

def test_complete_system():
    """Test all report generation formats"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🚀 Testing Complete PMO Report Generator System")
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
    
    # Generate test reports
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    test_dir = f"test_output_{timestamp}"
    os.makedirs(test_dir, exist_ok=True)
    
    print("\n📊 Generating Reports...")
    print("-" * 40)
    
    # Generate Excel
    excel_path = os.path.join(test_dir, f"PMO_Summary_{timestamp}.xlsx")
    create_excel_report(projects, excel_path)
    print(f"✅ Excel Report: {excel_path}")
    
    # Generate PDF (first 3 projects for quick test)
    pdf_path = os.path.join(test_dir, f"PMO_Report_{timestamp}.pdf")
    generate_pdf_report(projects[:3], pdf_path)
    print(f"✅ PDF Report: {pdf_path}")
    
    # Generate Word (first 3 projects for quick test)
    word_path = os.path.join(test_dir, f"PMO_Report_{timestamp}.docx")
    create_word_report(projects[:3], word_path)
    print(f"✅ Word Report: {word_path}")
    
    # Show summary
    print("\n📈 Report Summary:")
    print("-" * 40)
    print(f"Total Projects: {len(projects)}")
    
    # Health status breakdown
    on_track = sum(1 for p in projects if 'on track' in str(p.get('health', '')).lower())
    at_risk = sum(1 for p in projects if 'at risk' in str(p.get('health', '')).lower())
    off_track = sum(1 for p in projects if 'off track' in str(p.get('health', '')).lower() or 'delayed' in str(p.get('health', '')).lower())
    
    print(f"  🟢 On Track: {on_track}")
    print(f"  🟡 At Risk: {at_risk}")
    print(f"  🔴 Off Track: {off_track}")
    
    # Budget summary
    total_budget = sum(p.get('budget_total', 0) for p in projects)
    total_spent = sum(p.get('budget_spent', 0) for p in projects)
    
    print(f"\n💰 Budget Summary:")
    print(f"  Total Budget: {total_budget:,.2f} SAR")
    print(f"  Total Spent: {total_spent:,.2f} SAR")
    print(f"  Utilization: {(total_spent/total_budget*100):.1f}%" if total_budget > 0 else "  Utilization: N/A")
    
    print(f"\n✅ All reports generated successfully in '{test_dir}' folder!")
    print("\n📌 Features:")
    print("  • Excel: Complete dashboard with all project data")
    print("  • PDF: Professional landscape reports")
    print("  • Word: Editable documents")
    print("  • AI: Text formatting with OpenAI (if configured)")

if __name__ == "__main__":
    test_complete_system()
