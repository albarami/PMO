"""Test SPLD Health Report Generation"""

from dotenv import load_dotenv
load_dotenv()

import os
from datetime import datetime
from pmo_helpers import process_excel_file
from spld_health_report import generate_spld_health_report
from llm_integration import format_project_text

def test_spld_report():
    """Test SPLD health report generation"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🎯 Testing SPLD Project Health Report")
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
    
    # Generate SPLD report
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_path = f"SPLD_Health_Report_{timestamp}.pdf"
    
    print("\n📊 Generating SPLD Executive Dashboard Report...")
    generate_spld_health_report(projects, pdf_path)
    
    print(f"\n✅ SPLD Report generated: {pdf_path}")
    
    # Show summary statistics
    total = len(projects)
    on_track = sum(1 for p in projects if 'on track' in str(p.get('health', '')).lower())
    at_risk = sum(1 for p in projects if 'at risk' in str(p.get('health', '')).lower())
    off_track = total - on_track - at_risk
    
    print("\n📈 Portfolio Health Summary:")
    print("-" * 40)
    print(f"Total Projects: {total}")
    print(f"  🟢 On Track: {on_track} ({on_track/total*100:.0f}%)")
    print(f"  🟡 At Risk: {at_risk} ({at_risk/total*100:.0f}%)")
    print(f"  🔴 Off Track: {off_track} ({off_track/total*100:.0f}%)")
    
    print("\nReport Features:")
    print("  • SPLD branded executive dashboard")
    print("  • Portfolio health overview with traffic lights")
    print("  • Key performance metrics table")
    print("  • Detailed project list with status indicators")
    print("  • Professional SPLD orange and blue color scheme")

if __name__ == "__main__":
    test_spld_report()
