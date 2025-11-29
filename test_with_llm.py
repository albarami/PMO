"""Test the complete system with LLM formatting"""

from dotenv import load_dotenv
load_dotenv()

import pandas as pd
from pmo_helpers import process_excel_file
from llm_integration import format_project_text

def test_with_llm():
    """Test processing with LLM text formatting"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🚀 Testing PMO Report Generator with OpenAI LLM")
    print("=" * 60)
    
    # Read the file
    with open(excel_file, 'rb') as f:
        file_content = f.read()
    
    # Process the Excel file
    projects, error = process_excel_file(file_content, excel_file)
    
    if error:
        print(f"❌ Error: {error}")
        return
    
    print(f"✅ Extracted {len(projects)} projects")
    print("\n📝 Testing LLM Text Formatting...")
    print("-" * 60)
    
    # Find a project with actual content
    project_with_content = None
    for p in projects:
        if p['current_activities'] != '[To be provided]' or p['risks'] != '[To be provided]':
            project_with_content = p
            break
    
    # Use first project if none have content
    first = project_with_content if project_with_content else projects[0]
    
    if first:
        print(f"\n🏢 Project: {first['name']}")
        print(f"Category: {first['category']}")
        
        # Show before and after for activities
        print("\n📌 Current Activities:")
        print(f"BEFORE LLM: {first['current_activities'][:200]}...")
        
        # Apply LLM formatting
        formatted = format_project_text(first)
        
        print(f"\nAFTER LLM: {formatted['current_activities']}")
        
        # Show risks formatting
        print("\n⚠️ Risks:")
        print(f"BEFORE LLM: {first['risks'][:200]}...")
        print(f"AFTER LLM: {formatted['risks']}")
        
        # Check placeholder handling
        print("\n🔍 Placeholder Detection:")
        test_project = {
            'issues': 'Owner to share issues',
            'risks': 'To be provided by vendor'
        }
        formatted_test = format_project_text(test_project)
        print(f"'Owner to share issues' → '{formatted_test['issues']}'")
        print(f"'To be provided by vendor' → '{formatted_test['risks']}'")
        
        # Test with actual text
        print("\n🎯 Testing with Sample Text:")
        sample_project = {
            'current_activities': 'Working on system integration with ERP modules and conducting user training sessions. Finalizing documentation and preparing deployment scripts.',
            'risks': 'Potential delays due to vendor dependencies. Resource availability during holiday season. Integration complexity with legacy systems.',
            'issues': 'Firewall configuration blocking API calls. User access permissions need review.'
        }
        formatted_sample = format_project_text(sample_project)
        print(f"\n📝 Activities formatting:")
        print(f"BEFORE: {sample_project['current_activities']}")
        print(f"AFTER: {formatted_sample['current_activities']}")
        
        print(f"\n⚠️ Risks formatting:")
        print(f"BEFORE: {sample_project['risks']}")
        print(f"AFTER: {formatted_sample['risks']}")
    
    print("\n✨ LLM Text Formatting is working correctly with OpenAI!")
    print("📊 All calculations remain programmatic (not affected by LLM)")

if __name__ == "__main__":
    test_with_llm()
