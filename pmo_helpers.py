"""Helper functions for PMO Report Generator"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from reportlab.lib import colors

# Required columns mapping (flexible names)
COLUMN_MAPPING = {
    'project_number': ['#', 'No', 'Number', 'Project #', 'ID'],
    'project_name': ['Project Name', 'Name', 'Project Title'],
    'project_category': ['Project Category', 'Category', 'Type'],
    'project_status': ['Project Status', 'Status'],
    'gm': ['GM', 'General Manager', 'Sponsor GM'],
    'director': ['SPLD Director / GM', 'Director', 'SPLD Director', 'Project Director'],
    'operational_lead': ['Project operational Lead', 'Operational Lead', 'Lead', 'Project Lead', 'Project Manager'],
    'contract_end_date': ['Contract End Date', 'End Date', 'Finish Date', 'Deadline'],
    'days_remaining': ['Days Remaining (Until Contract End)', 'Days Remaining', 'Days Left'],
    'budget_spent': ['Budget (Spent)', 'Budget Spent', 'Spent', 'Actual Cost'],
    'budget_remaining': ['Budget Remaining', 'Remaining Budget', 'Budget Left'],
    'timeline_actual': ['timeline Actual', 'Actual Progress', 'Actual %', 'Progress Actual'],
    'timeline_planned': ['timeline planned', 'Planned Progress', 'Planned %', 'Progress Planned'],
    'kpi': ['Service delivery Performance KPI', 'KPI', 'Performance KPI'],
    'service_performance': ['Service delivery Performance', 'Performance', 'Service Performance'],
    'project_health': ['Project health (on track - at risk - off track)', 'Project Health', 'Health', 'Status Health'],
    'issues': ['Issues (From Owner List)', 'Issues', 'Current Issues'],
    'risks': ['Risks', 'Risk', 'Project Risks'],
    'current_activities': ['Current activites', 'Current Activities', 'Current Work'],
    'future_activities': ['Future Activites', 'Future Activities', 'Next Steps', 'Upcoming Activities'],
    'comments': ['Comments  to the owner', 'Comments', 'Notes', 'Owner Comments'],
    'vendor': ['Vendor', 'Supplier', 'Contractor']
}


def find_column(df, column_key):
    """Find the actual column name in dataframe based on possible names."""
    possible_names = COLUMN_MAPPING.get(column_key, [])
    for name in possible_names:
        if name in df.columns:
            return name
        # Case-insensitive search
        for col in df.columns:
            if col.lower().strip() == name.lower().strip():
                return col
    return None


def map_columns(df):
    """Map all columns to standardized names."""
    column_map = {}
    missing = []
    for key in COLUMN_MAPPING.keys():
        found = find_column(df, key)
        if found:
            column_map[key] = found
        else:
            missing.append(key)
    return column_map, missing


def safe_get(row, column_map, key, default=''):
    """Safely get a value from row using mapped column name."""
    col_name = column_map.get(key)
    if col_name and col_name in row.index:
        val = row[col_name]
        if pd.isna(val):
            return default
        return val
    return default


def parse_number(value, default=0):
    """Parse a number from various formats."""
    if pd.isna(value) or value == '' or value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    # Handle string formats like "5,004,225.00SAR" or "12,621,386.75"
    try:
        cleaned = str(value).replace(',', '').replace('SAR', '').replace('$', '').strip()
        return float(cleaned)
    except:
        return default


def parse_percentage(value, default=0):
    """Parse percentage from various formats."""
    if pd.isna(value) or value == '' or value is None:
        return default
    if isinstance(value, (int, float)):
        # If it's already a decimal (0.4) convert to percentage
        if value <= 1:
            return value * 100
        return value
    try:
        cleaned = str(value).replace('%', '').strip()
        val = float(cleaned)
        if val <= 1:
            return val * 100
        return val
    except:
        return default


def parse_date(value):
    """Parse date from various formats."""
    if pd.isna(value) or value == '' or value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    try:
        return pd.to_datetime(value).to_pydatetime()
    except:
        return None


def calculate_days_remaining(end_date):
    """Calculate days remaining until contract end."""
    if end_date is None:
        return 0
    today = datetime.now()
    delta = end_date - today
    return max(0, delta.days)


def calculate_budget_utilization(spent, remaining):
    """Calculate budget utilization percentage."""
    total = spent + remaining
    if total == 0:
        return 0, 0
    spent_pct = (spent / total) * 100
    remaining_pct = (remaining / total) * 100
    return round(spent_pct, 1), round(remaining_pct, 1)


def determine_health_color(health_status):
    """Determine color based on health status."""
    if pd.isna(health_status) or health_status == '':
        return colors.gray
    status = str(health_status).lower().strip()
    if 'on track' in status or 'on-track' in status:
        return colors.Color(0.2, 0.7, 0.3)  # Green
    elif 'at risk' in status or 'at-risk' in status:
        return colors.Color(1, 0.6, 0)  # Orange
    elif 'off track' in status or 'delayed' in status or 'off-track' in status:
        return colors.Color(0.8, 0.2, 0.2)  # Red
    return colors.gray


def clean_text(text, max_length=500):
    """Clean and truncate text for display."""
    if pd.isna(text) or text == '' or text is None:
        return '[To be provided]'
    text = str(text).strip()
    # Remove excessive whitespace and newlines
    text = ' '.join(text.split())
    if len(text) > max_length:
        text = text[:max_length] + '...'
    return text


def extract_project_data(row, column_map):
    """Extract and calculate all project metrics from a row."""
    # Basic info
    project = {
        'number': safe_get(row, column_map, 'project_number', ''),
        'name': safe_get(row, column_map, 'project_name', 'Unnamed Project'),
        'category': safe_get(row, column_map, 'project_category', ''),
        'status': safe_get(row, column_map, 'project_status', ''),
        'gm': safe_get(row, column_map, 'gm', ''),
        'director': safe_get(row, column_map, 'director', ''),
        'operational_lead': safe_get(row, column_map, 'operational_lead', ''),
        'vendor': safe_get(row, column_map, 'vendor', ''),
    }
    
    # Dates - CALCULATED
    end_date = parse_date(safe_get(row, column_map, 'contract_end_date'))
    project['contract_end_date'] = end_date.strftime('%d %b %Y') if end_date else '[TBD]'
    project['days_remaining'] = calculate_days_remaining(end_date)
    
    # If days_remaining is in the Excel, use it if our calculation is 0 but Excel has a value
    excel_days = parse_number(safe_get(row, column_map, 'days_remaining'), 0)
    if project['days_remaining'] == 0 and excel_days > 0:
        project['days_remaining'] = int(excel_days)
    
    # Budget - CALCULATED
    budget_spent = parse_number(safe_get(row, column_map, 'budget_spent'))
    budget_remaining = parse_number(safe_get(row, column_map, 'budget_remaining'))
    project['budget_spent'] = budget_spent
    project['budget_remaining'] = budget_remaining
    project['budget_total'] = budget_spent + budget_remaining
    spent_pct, remaining_pct = calculate_budget_utilization(budget_spent, budget_remaining)
    project['budget_spent_pct'] = spent_pct
    project['budget_remaining_pct'] = remaining_pct
    
    # Timeline - CALCULATED
    actual_progress = parse_percentage(safe_get(row, column_map, 'timeline_actual'))
    planned_progress = parse_percentage(safe_get(row, column_map, 'timeline_planned'))
    project['timeline_actual'] = round(actual_progress, 1)
    project['timeline_planned'] = round(planned_progress, 1)
    
    # Calculate schedule variance
    project['schedule_variance'] = round(actual_progress - planned_progress, 1)
    
    # Performance & Health
    project['kpi'] = clean_text(safe_get(row, column_map, 'kpi'), 200)
    project['service_performance'] = safe_get(row, column_map, 'service_performance', 'TBD')
    project['health'] = safe_get(row, column_map, 'project_health', '')
    project['health_color'] = determine_health_color(project['health'])
    
    # Activities & Risks (for LLM interpretation)
    project['issues'] = clean_text(safe_get(row, column_map, 'issues'), 500)
    project['risks'] = clean_text(safe_get(row, column_map, 'risks'), 500)
    project['current_activities'] = clean_text(safe_get(row, column_map, 'current_activities'), 500)
    project['future_activities'] = clean_text(safe_get(row, column_map, 'future_activities'), 500)
    project['comments'] = clean_text(safe_get(row, column_map, 'comments'), 500)
    
    return project


def process_excel_file(file_content, filename):
    """Process uploaded Excel file and extract project data."""
    try:
        # Read Excel file
        import io
        df = pd.read_excel(io.BytesIO(file_content), sheet_name=0)
        
        # Map columns
        column_map, missing_cols = map_columns(df)
        
        # Check for critical missing columns
        critical_cols = ['project_name']
        critical_missing = [col for col in critical_cols if col in missing_cols]
        if critical_missing:
            return None, f"Missing critical columns: {', '.join(critical_missing)}"
        
        # Extract project data
        projects = []
        for idx, row in df.iterrows():
            # Skip rows where project name is empty
            name = safe_get(row, column_map, 'project_name', '')
            if pd.isna(name) or str(name).strip() == '':
                continue
            
            project = extract_project_data(row, column_map)
            projects.append(project)
        
        if not projects:
            return None, "No valid projects found in the Excel file"
        
        return projects, None
        
    except Exception as e:
        return None, f"Error processing file: {str(e)}"
