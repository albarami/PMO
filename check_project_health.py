"""Check actual project health status from Excel"""

from dotenv import load_dotenv
load_dotenv()

from pmo_helpers import process_excel_file

def check_health_status():
    """Check actual health status of all projects"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🔍 Checking Project Health Status from Excel")
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
    
    # Count health status
    on_track = []
    at_risk = []
    off_track = []
    unknown = []
    
    print("Detailed Health Status:")
    print("-" * 60)
    
    for i, project in enumerate(projects, 1):
        name = project.get('name', 'Unknown')[:40]
        health = str(project.get('health', 'Unknown'))
        
        print(f"{i:2}. {name:40} → {health}")
        
        health_lower = health.lower()
        if 'on track' in health_lower:
            on_track.append((name, health))
        elif 'at risk' in health_lower:
            at_risk.append((name, health))
        elif 'off track' in health_lower or 'delayed' in health_lower:
            off_track.append((name, health))
        else:
            unknown.append((name, health))
    
    print("\n" + "=" * 60)
    print("📊 SUMMARY:")
    print(f"✅ On Track: {len(on_track)} projects")
    print(f"⚠️  At Risk: {len(at_risk)} projects")
    print(f"❌ Off Track: {len(off_track)} projects")
    print(f"❓ Unknown: {len(unknown)} projects")
    
    if at_risk:
        print("\n⚠️ AT RISK PROJECTS:")
        for name, health in at_risk:
            print(f"  - {name}: {health}")
    
    if off_track:
        print("\n❌ OFF TRACK PROJECTS:")
        for name, health in off_track:
            print(f"  - {name}: {health}")
    
    if unknown:
        print("\n❓ UNKNOWN STATUS PROJECTS:")
        for name, health in unknown:
            print(f"  - {name}: {health}")

if __name__ == "__main__":
    check_health_status()
