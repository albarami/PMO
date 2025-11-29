"""Check raw health column data from Excel"""

import pandas as pd

def check_raw_health():
    """Check the raw health column data"""
    excel_file = "PMO_Project_Tracker_Current.xlsx"
    
    print("🔍 Checking RAW Health Column Data from Excel")
    print("=" * 60)
    
    # Read Excel directly with pandas
    df = pd.read_excel(excel_file)
    
    # Find health column
    health_columns = [col for col in df.columns if 'health' in col.lower() or 'track' in col.lower()]
    
    print(f"Found health columns: {health_columns}")
    print()
    
    if health_columns:
        health_col = health_columns[0]
        print(f"Using column: '{health_col}'")
        print("-" * 60)
        
        # Show raw values
        print("\nRAW VALUES in health column:")
        for i, (idx, row) in enumerate(df.iterrows(), 1):
            project = row.get('Project Name', row.get('Name', f'Row {i}'))[:40]
            health_value = row[health_col]
            print(f"{i:2}. {project:40} → '{health_value}'")
        
        # Count unique values
        print("\n" + "=" * 60)
        print("UNIQUE VALUES COUNT:")
        value_counts = df[health_col].value_counts(dropna=False)
        for value, count in value_counts.items():
            if pd.isna(value):
                print(f"  [EMPTY/NaN]: {count} projects")
            else:
                print(f"  '{value}': {count} projects")
    else:
        print("❌ No health column found!")

if __name__ == "__main__":
    check_raw_health()
