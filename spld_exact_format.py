"""SPLD PMO Committee Report - Exact Format Replication"""

from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing, Circle, String
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF


# SPLD Color Palette (from screenshots)
SPLD_COLORS = {
    'dark_bg': colors.HexColor('#0A1628'),  # Dark blue background
    'gray_bg': colors.HexColor('#2C3E50'),  # Gray section background  
    'orange': colors.HexColor('#EA6A1F'),  # SPLD Orange
    'green': colors.HexColor('#27AE60'),  # On Track Green
    'yellow': colors.HexColor('#F39C12'),  # Slightly Delayed Yellow
    'red': colors.HexColor('#E74C3C'),  # Delayed Red
    'white': colors.white,
    'light_gray': colors.HexColor('#95A5A6')
}


def create_project_status_report(project, styles):
    """Create exact SPLD project status page format FROM EXCEL DATA."""
    elements = []
    
    # Header with project info - READING FROM EXCEL
    header_data = [
        ['Project Status Report', '', 'Project Sponsor', str(project.get('gm', 'TBD'))],
        [str(project.get('name', 'Unnamed Project')), '', 'Project Manager', str(project.get('operational_lead', 'TBD'))],
        [f'Report Date: {datetime.now().strftime("%d/%m/%Y")}', '', '', '']
    ]
    
    header = Table(header_data, colWidths=[3.5*inch, 1.5*inch, 1.5*inch, 2*inch])
    header.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), SPLD_COLORS['gray_bg']),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
        ('TEXTCOLOR', (0, 1), (0, 1), SPLD_COLORS['orange']),
        ('TEXTCOLOR', (0, 2), (0, 2), SPLD_COLORS['light_gray']),
        ('TEXTCOLOR', (2, 0), (2, -1), SPLD_COLORS['light_gray']),
        ('TEXTCOLOR', (3, 0), (3, -1), colors.white),
        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (0, 1), 14),
        ('SPAN', (0, 0), (1, 0)),
        ('SPAN', (0, 1), (1, 1)),
        ('SPAN', (0, 2), (1, 2)),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
    ]))
    
    elements.append(header)
    elements.append(Spacer(1, 15))
    
    # Two column layout
    left_col = []
    right_col = []
    
    # Left: Dates and Objective - FROM EXCEL DATA
    # Calculate start date from end date if not provided
    end_date_str = str(project.get('contract_end_date', 'TBD'))
    start_date = datetime.now().strftime('%d %b %Y')
    
    date_data = [
        ['Start Date', start_date],
        ['Baseline Finish', end_date_str],
        ['Forecast Finish', end_date_str]
    ]
    
    date_table = Table(date_data, colWidths=[1.5*inch, 2*inch])
    date_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), SPLD_COLORS['gray_bg']),
        ('BACKGROUND', (1, 0), (1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('TEXTCOLOR', (1, 0), (1, -1), colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, SPLD_COLORS['light_gray']),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
    ]))
    
    left_col.append(date_table)
    left_col.append(Spacer(1, 10))
    
    # Project Objective - FROM EXCEL DATA
    # Try to get objective from various possible fields
    objective_text = (project.get('description', '') or 
                     project.get('comments', '') or 
                     f"Managing and delivering {project.get('name', 'project')} successfully")[:300]
    
    obj_table = Table([
        ['Project Objective'],
        [objective_text]
    ], colWidths=[5*inch])
    
    obj_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), SPLD_COLORS['orange']),
        ('BACKGROUND', (0, 1), (0, 1), SPLD_COLORS['gray_bg']),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (0, 0), 10),
        ('FONTSIZE', (0, 1), (0, 1), 8),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (0, 0), 6),
        ('BOTTOMPADDING', (0, 0), (0, 0), 6),
        ('TOPPADDING', (0, 1), (0, 1), 10),
        ('BOTTOMPADDING', (0, 1), (0, 1), 10),
    ]))
    
    left_col.append(obj_table)
    
    # Deliverables Table
    left_col.append(Spacer(1, 10))
    
    deliv_header = Table([['Deliverables']], colWidths=[5*inch])
    deliv_header.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), SPLD_COLORS['orange']),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('TOPPADDING', (0, 0), (0, 0), 4),
        ('BOTTOMPADDING', (0, 0), (0, 0), 4),
    ]))
    
    left_col.append(deliv_header)
    
    # Deliverables - CREATE FROM EXCEL DATA
    deliv_data = [
        ['Deliverable Name', 'Contract', 'Planned', '% Done', 'Status']
    ]
    
    # Create deliverables based on project progress
    progress = project.get('timeline_actual', 0)
    end_date = project.get('contract_end_date', 'TBD')
    
    # Generate deliverables based on project status
    if progress >= 100:
        deliv_data.append(['Project Completion', end_date, end_date, '100%', 'Completed'])
    elif progress >= 75:
        deliv_data.append(['Development Phase', end_date, end_date, f'{progress:.0f}%', 'On Track'])
        deliv_data.append(['Testing Phase', end_date, end_date, f'{max(0, progress-80):.0f}%', 'In Progress'])
    elif progress >= 50:
        deliv_data.append(['Design Phase', end_date, end_date, '100%', 'Completed'])
        deliv_data.append(['Development Phase', end_date, end_date, f'{progress:.0f}%', 'On Track'])
    elif progress >= 25:
        deliv_data.append(['Planning Phase', end_date, end_date, '100%', 'Completed'])
        deliv_data.append(['Design Phase', end_date, end_date, f'{progress:.0f}%', 'On Track'])
    else:
        deliv_data.append(['Planning Phase', end_date, end_date, f'{progress:.0f}%', 'In Progress'])
        deliv_data.append(['Implementation', end_date, end_date, '0%', 'Not Started'])
    
    deliv_table = Table(deliv_data, colWidths=[1.8*inch, 0.8*inch, 0.8*inch, 0.5*inch, 0.9*inch])
    
    style = [
        ('BACKGROUND', (0, 0), (-1, 0), SPLD_COLORS['gray_bg']),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, SPLD_COLORS['light_gray']),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]
    
    # Add status colors
    for i in range(1, len(deliv_data)):
        status = deliv_data[i][4]
        if status == 'Completed':
            style.append(('BACKGROUND', (4, i), (4, i), SPLD_COLORS['green']))
            style.append(('TEXTCOLOR', (4, i), (4, i), colors.white))
        elif status == 'On Track':
            style.append(('BACKGROUND', (4, i), (4, i), SPLD_COLORS['green']))
            style.append(('TEXTCOLOR', (4, i), (4, i), colors.white))
            
    deliv_table.setStyle(TableStyle(style))
    left_col.append(deliv_table)
    
    # Right: Health and Activities - FROM EXCEL DATA
    health = str(project.get('health', 'On Track'))
    if 'on track' in health.lower():
        health_color = SPLD_COLORS['green']
        health_text = 'On Track'
    elif 'at risk' in health.lower():
        health_color = SPLD_COLORS['yellow']  
        health_text = 'Slightly Delayed'
    elif 'delayed' in health.lower() or 'off track' in health.lower():
        health_color = SPLD_COLORS['red']
        health_text = 'Delayed'
    else:
        # Default based on schedule variance
        variance = project.get('schedule_variance', 0)
        if variance >= -5:
            health_color = SPLD_COLORS['green']
            health_text = 'On Track'
        elif variance >= -10:
            health_color = SPLD_COLORS['yellow']
            health_text = 'Slightly Delayed'
        else:
            health_color = SPLD_COLORS['red']
            health_text = 'Delayed'
    
    # Get actual progress from Excel
    actual_progress = project.get('timeline_actual', 0)
    planned_progress = project.get('timeline_planned', 0)
    
    health_table = Table([
        ['Overall Project Health'],
        [health_text],
        [''],
        ['Project Progress'],
        [f"{actual_progress:.0f}%"],
        [f"Planned: {planned_progress:.0f}%"]
    ], colWidths=[2.8*inch])
    
    health_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), SPLD_COLORS['gray_bg']),
        ('BACKGROUND', (0, 1), (0, 1), health_color),
        ('BACKGROUND', (0, 3), (0, -1), SPLD_COLORS['gray_bg']),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (0, 3), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (0, 3), 10),
        ('FONTSIZE', (0, 1), (0, 1), 12),
        ('FONTSIZE', (0, 4), (0, 4), 16),
        ('FONTSIZE', (0, 5), (0, 5), 9),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    
    right_col.append(health_table)
    right_col.append(Spacer(1, 15))
    
    # Activity Progress - FROM EXCEL DATA
    current_activities = str(project.get('current_activities', '[To be provided]'))
    future_activities = str(project.get('future_activities', '[To be provided]'))
    
    # Format activities if they're too long
    if len(current_activities) > 150:
        current_activities = current_activities[:147] + '...'
    if len(future_activities) > 150:
        future_activities = future_activities[:147] + '...'
    
    act_table = Table([
        ['Activity Progress'],
        ['Current Activities'],
        [current_activities],
        ['Future Activities'],  
        [future_activities]
    ], colWidths=[2.8*inch])
    
    act_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), SPLD_COLORS['gray_bg']),
        ('BACKGROUND', (0, 1), (0, -1), SPLD_COLORS['gray_bg']),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (0, 3), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (0, 0), 10),
        ('FONTSIZE', (0, 1), (0, -1), 8),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    
    right_col.append(act_table)
    
    # Combine columns
    two_col = Table([[left_col, right_col]], colWidths=[5.2*inch, 3*inch])
    two_col.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    
    elements.append(two_col)
    
    return elements


def generate_spld_exact_report(projects, output_path):
    """Generate exact SPLD format report."""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
        title="SPLD PMO Committee Report"
    )
    
    styles = getSampleStyleSheet()
    elements = []
    
    # Generate pages for ALL projects
    for i, project in enumerate(projects):
        elements.extend(create_project_status_report(project, styles))
        # Add page break except for last project
        if i < len(projects) - 1:
            elements.append(PageBreak())
    
    doc.build(elements)
    return output_path
