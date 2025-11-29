"""Test that activities are being read correctly from Excel"""

from dotenv import load_dotenv
load_dotenv()

from pmo_helpers import process_excel_file

def test_activities():
    """Test reading Current and Future Activities from Excel"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("📋 Testing Activities Reading from Excel")
    print("=" * 60)
    
    # Read the file
    with open(excel_file, 'rb') as f:
        file_content = f.read()
    
    # Process the Excel file
    projects, error = process_excel_file(file_content, excel_file)
    
    if error:
        print(f"❌ Error: {error}")
        return
    
    print(f"✅ Found {len(projects)} projects\n")
    
    # Check activities for first 5 projects
    print("Activities Status:")
    print("-" * 60)
    
    has_current = 0
    has_future = 0
    
    for i, project in enumerate(projects[:10], 1):
        current = project.get('current_activities', 'MISSING')
        future = project.get('future_activities', 'MISSING')
        
        print(f"\n{i}. {project.get('name', 'Unknown')[:40]}...")
        print(f"   Current: {current[:60]}...")
        print(f"   Future: {future[:60]}...")
        
        if current and current != '[To be provided]' and 'owner to share' not in current.lower():
            has_current += 1
        if future and future != '[To be provided]' and 'owner to share' not in future.lower():
            has_future += 1
    
    print(f"\n📊 Summary:")
    print(f"  • Projects with current activities: {has_current}/{len(projects)}")
    print(f"  • Projects with future activities: {has_future}/{len(projects)}")
    
    # Show what the system is looking for
    from pmo_helpers import COLUMN_MAPPING
    print(f"\n🔍 Looking for columns:")
    print(f"  Current: {COLUMN_MAPPING['current_activities']}")
    print(f"  Future: {COLUMN_MAPPING['future_activities']}")

if __name__ == "__main__":
    test_activities()
