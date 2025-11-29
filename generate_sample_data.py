"""Generate sample PMO tracker Excel file for testing"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def generate_sample_pmo_data(num_projects=5):
    """Generate sample PMO tracker data"""
    
    # Sample data pools
    categories = ['Infrastructure', 'Digital Transformation', 'Operations', 'Compliance', 'Innovation']
    statuses = ['In Progress', 'Planning', 'Execution', 'Testing', 'Closing']
    health_statuses = ['on track', 'at risk', 'off track']
    gms = ['Ahmed Al-Rashid', 'Sarah Johnson', 'Mohammed Al-Qahtani', 'Lisa Chen', 'David Smith']
    directors = ['John Williams', 'Fatima Al-Zahrani', 'Robert Brown', 'Aisha Khan', 'Michael Davis']
    leads = ['Tom Wilson', 'Nora Al-Saud', 'Jennifer Lee', 'Ali Hassan', 'Emma Thompson']
    vendors = ['TechCorp Solutions', 'GlobalIT Services', 'Digital Partners', 'Systems Inc', 'CloudTech Pro']
    
    # Sample activities
    current_activities = [
        'Requirements gathering and stakeholder alignment',
        'System architecture design and review',
        'Development of core modules',
        'Integration testing with existing systems',
        'User acceptance testing preparation',
        'Documentation and training material development',
        'Security assessment and compliance review',
        'Performance optimization and tuning'
    ]
    
    future_activities = [
        'Deploy to production environment',
        'Conduct end-user training sessions',
        'Implement monitoring and alerting',
        'Perform post-implementation review',
        'Develop maintenance procedures',
        'Create disaster recovery plan',
        'Establish support processes',
        'Plan for phase 2 enhancements'
    ]
    
    issues = [
        'Resource availability constraints',
        'Scope creep requiring change control',
        'Integration challenges with legacy systems',
        'Vendor delivery delays',
        'Budget approval pending',
        'Technical debt from previous phases',
        'Stakeholder alignment issues',
        'Compliance requirements changes'
    ]
    
    risks = [
        'Potential budget overrun due to scope changes',
        'Key resource dependency - single point of failure',
        'Timeline compression may impact quality',
        'Third-party system compatibility uncertain',
        'Regulatory changes may affect requirements',
        'User adoption challenges anticipated',
        'Data migration complexity underestimated',
        'Performance requirements may not be met'
    ]
    
    kpis = [
        'SLA: 99.9% uptime | Current: 99.95%',
        'Response Time: <2s target | Current: 1.8s',
        'User Satisfaction: >4.0 target | Current: 4.2',
        'Defect Rate: <5% target | Current: 3.2%',
        'On-time Delivery: 95% target | Current: 92%'
    ]
    
    data = []
    
    for i in range(num_projects):
        # Generate random but realistic data
        start_date = datetime.now() - timedelta(days=random.randint(30, 365))
        end_date = start_date + timedelta(days=random.randint(90, 730))
        days_elapsed = (datetime.now() - start_date).days
        total_days = (end_date - start_date).days
        
        # Calculate progress
        planned_progress = min(100, (days_elapsed / total_days) * 100 * random.uniform(0.9, 1.1))
        actual_progress = planned_progress + random.uniform(-15, 10)
        actual_progress = max(0, min(100, actual_progress))
        
        # Budget calculations
        total_budget = random.randint(500000, 10000000)
        budget_spent = total_budget * (actual_progress / 100) * random.uniform(0.8, 1.2)
        budget_remaining = total_budget - budget_spent
        
        # Determine health based on variance
        schedule_variance = actual_progress - planned_progress
        if schedule_variance > -5:
            health = 'on track'
        elif schedule_variance > -10:
            health = 'at risk'
        else:
            health = 'off track'
        
        project = {
            '#': f'PRJ-{1000 + i}',
            'Project Name': f'Project {chr(65 + i)} - {random.choice(categories)} Initiative',
            'Project Category': random.choice(categories),
            'Project Status': random.choice(statuses),
            'GM': random.choice(gms),
            'SPLD Director / GM': random.choice(directors),
            'Project operational Lead': random.choice(leads),
            'Contract End Date': end_date.strftime('%m/%d/%Y'),
            'Days Remaining (Until Contract End)': max(0, (end_date - datetime.now()).days),
            'Budget (Spent)': f'{budget_spent:,.2f}SAR',
            'Budget Remaining': f'{budget_remaining:,.2f}SAR',
            'timeline Actual': f'{actual_progress:.1f}%',
            'timeline planned': f'{planned_progress:.1f}%',
            'Service delivery Performance KPI': random.choice(kpis),
            'Service delivery Performance': f'{random.randint(85, 100)}%',
            'Project health (on track - at risk - off track)': health,
            'Issues (From Owner List)': random.choice(issues),
            'Risks': random.choice(risks),
            'Current activites': random.choice(current_activities),
            'Future Activites': random.choice(future_activities),
            'Comments  to the owner': f'Project progressing as per revised plan. {random.choice(["Weekly steering committee meetings ongoing.", "Awaiting stakeholder approval for next phase.", "Resource augmentation in progress.", "Quality metrics within acceptable range."])}',
            'Vendor': random.choice(vendors)
        }
        
        data.append(project)
    
    return pd.DataFrame(data)


def main():
    """Generate and save sample PMO tracker"""
    print("Generating sample PMO tracker data...")
    
    # Generate data for 10 projects
    df = generate_sample_pmo_data(10)
    
    # Save to Excel
    filename = 'sample_pmo_tracker.xlsx'
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='PMO Tracker', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['PMO Tracker']
        for column in df:
            column_length = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            worksheet.column_dimensions[chr(65 + col_idx)].width = min(50, column_length + 2)
    
    print(f"✅ Sample PMO tracker saved as '{filename}'")
    print(f"   - Contains {len(df)} projects")
    print(f"   - Columns: {', '.join(df.columns[:5])}...")
    print("\nYou can now upload this file to the PMO Report Generator web application.")


if __name__ == '__main__':
    main()
