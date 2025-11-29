"""PDF Report Generation Module for PMO Reports"""

from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from reportlab.lib.enums import TA_CENTER


def create_styles():
    """Create custom paragraph styles for the report."""
    styles = getSampleStyleSheet()
    
    # Check if custom styles already exist to avoid duplicates
    if 'ReportTitle' not in styles:
        styles.add(ParagraphStyle(
            name='ReportTitle',
            parent=styles['Heading1'],
            fontSize=18,
            textColor=colors.Color(0.2, 0.4, 0.6),
            spaceAfter=20,
            alignment=TA_CENTER
        ))
    
    if 'ProjectName' not in styles:
        styles.add(ParagraphStyle(
            name='ProjectName',
            parent=styles['Heading2'],
            fontSize=14,
            textColor=colors.Color(0.9, 0.5, 0.2),  # Orange
            spaceBefore=10,
            spaceAfter=5
        ))
    
    if 'SectionHeader' not in styles:
        styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=styles['Heading3'],
            fontSize=11,
            textColor=colors.Color(0.9, 0.5, 0.2),
            spaceBefore=12,
            spaceAfter=6,
            fontName='Helvetica-Bold'
        ))
    
    if 'PMOBodyText' not in styles:
        styles.add(ParagraphStyle(
            name='PMOBodyText',
            parent=styles['Normal'],
            fontSize=9,
            leading=12,
            spaceBefore=3,
            spaceAfter=3
        ))
    
    if 'TableText' not in styles:
        styles.add(ParagraphStyle(
            name='TableText',
            parent=styles['Normal'],
            fontSize=8,
            leading=10
        ))
    
    if 'Label' not in styles:
        styles.add(ParagraphStyle(
            name='Label',
            parent=styles['Normal'],
            fontSize=8,
            textColor=colors.gray,
            spaceBefore=2,
            spaceAfter=1
        ))
    
    if 'Value' not in styles:
        styles.add(ParagraphStyle(
            name='Value',
            parent=styles['Normal'],
            fontSize=9,
            fontName='Helvetica-Bold',
            spaceBefore=0,
            spaceAfter=4
        ))
    
    return styles


def create_project_report_page(project, styles):
    """Create report elements for a single project."""
    elements = []
    
    # Header section with project info
    header_data = [
        [
            Paragraph(f"<b>Project Status Report</b>", styles['ReportTitle']),
            '',
            Paragraph(f"<b>Report Date:</b> {datetime.now().strftime('%d/%m/%Y')}", styles['PMOBodyText'])
        ],
        [
            Paragraph(f"<font color='#E07020'><b>{project['name']}</b></font>", styles['ProjectName']),
            '',
            ''
        ],
        [
            Paragraph(f"<b>Category:</b> {project['category']}", styles['TableText']),
            Paragraph(f"<b>Status:</b> {project['status']}", styles['TableText']),
            Paragraph(f"<b>Vendor:</b> {str(project['vendor'])[:50]}", styles['TableText'])
        ]
    ]
    
    header_table = Table(header_data, colWidths=[4*inch, 2*inch, 3*inch])
    header_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('SPAN', (0, 1), (2, 1)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 10))
    
    # Divider
    elements.append(HRFlowable(width="100%", thickness=2, color=colors.Color(0.9, 0.5, 0.2)))
    elements.append(Spacer(1, 10))
    
    # Key metrics row
    # Determine health color for display
    health_text = str(project['health']) if project['health'] else 'TBD'
    health_display = health_text.replace('on track', 'On Track').replace('On Track', 'On Track')
    
    if 'on track' in health_text.lower():
        health_bg = colors.Color(0.2, 0.7, 0.3)
    elif 'at risk' in health_text.lower():
        health_bg = colors.Color(1, 0.6, 0)
    elif 'delayed' in health_text.lower() or 'off track' in health_text.lower():
        health_bg = colors.Color(0.8, 0.2, 0.2)
    else:
        health_bg = colors.gray
    
    metrics_data = [
        [
            Paragraph("<b>Project Sponsor GM</b>", styles['Label']),
            Paragraph("<b>Project Director</b>", styles['Label']),
            Paragraph("<b>Project Lead</b>", styles['Label']),
            Paragraph("<b>Contract End Date</b>", styles['Label']),
            Paragraph("<b>Days Remaining</b>", styles['Label']),
        ],
        [
            Paragraph(str(project['gm'] or 'TBD'), styles['Value']),
            Paragraph(str(project['director'] or 'TBD'), styles['Value']),
            Paragraph(str(project['operational_lead'] or 'TBD'), styles['Value']),
            Paragraph(str(project['contract_end_date']), styles['Value']),
            Paragraph(str(project['days_remaining']), styles['Value']),
        ]
    ]
    
    metrics_table = Table(metrics_data, colWidths=[1.8*inch, 1.8*inch, 1.8*inch, 1.5*inch, 1.2*inch])
    metrics_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.15, 0.15, 0.15)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, 1), colors.Color(0.25, 0.25, 0.25)),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.4, 0.4, 0.4)),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(metrics_table)
    elements.append(Spacer(1, 15))
    
    # Two-column layout: Left (Progress & Budget) | Right (Health & Activities)
    left_col_data = []
    right_col_data = []
    
    # Left column - Timeline Progress
    left_col_data.append([Paragraph("<font color='#E07020'><b>Project Timeline</b></font>", styles['SectionHeader'])])
    left_col_data.append([Paragraph(f"<b>Actual Progress:</b> {project['timeline_actual']:.1f}%", styles['PMOBodyText'])])
    left_col_data.append([Paragraph(f"<b>Planned Progress:</b> {project['timeline_planned']:.1f}%", styles['PMOBodyText'])])
    
    variance_color = 'green' if project['schedule_variance'] >= 0 else 'red'
    left_col_data.append([Paragraph(f"<b>Schedule Variance:</b> <font color='{variance_color}'>{project['schedule_variance']:+.1f}%</font>", styles['PMOBodyText'])])
    left_col_data.append([Spacer(1, 10)])
    
    # Budget section
    left_col_data.append([Paragraph("<font color='#E07020'><b>Budget Utilization</b></font>", styles['SectionHeader'])])
    left_col_data.append([Paragraph(f"<b>Total Budget:</b> {project['budget_total']:,.2f} SAR", styles['PMOBodyText'])])
    left_col_data.append([Paragraph(f"<b>Spent:</b> {project['budget_spent']:,.2f} SAR ({project['budget_spent_pct']:.1f}%)", styles['PMOBodyText'])])
    left_col_data.append([Paragraph(f"<b>Remaining:</b> {project['budget_remaining']:,.2f} SAR ({project['budget_remaining_pct']:.1f}%)", styles['PMOBodyText'])])
    left_col_data.append([Spacer(1, 10)])
    
    # KPI
    left_col_data.append([Paragraph("<font color='#E07020'><b>Service Delivery KPI</b></font>", styles['SectionHeader'])])
    left_col_data.append([Paragraph(str(project['kpi']), styles['PMOBodyText'])])
    
    # Right column - Health and Activities
    right_col_data.append([Paragraph("<font color='#E07020'><b>Overall Project Health</b></font>", styles['SectionHeader'])])
    
    # Health indicator box
    health_box = Table([[Paragraph(f"<b>{health_display}</b>", styles['Value'])]], colWidths=[2*inch])
    health_box.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), health_bg),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (0, 0), 8),
        ('BOTTOMPADDING', (0, 0), (0, 0), 8),
        ('LEFTPADDING', (0, 0), (0, 0), 10),
        ('RIGHTPADDING', (0, 0), (0, 0), 10),
    ]))
    right_col_data.append([health_box])
    right_col_data.append([Spacer(1, 10)])
    
    # Current Activities
    right_col_data.append([Paragraph("<font color='#E07020'><b>Current Activities</b></font>", styles['SectionHeader'])])
    right_col_data.append([Paragraph(str(project['current_activities']), styles['PMOBodyText'])])
    right_col_data.append([Spacer(1, 8)])
    
    # Future Activities
    right_col_data.append([Paragraph("<font color='#E07020'><b>Future Activities</b></font>", styles['SectionHeader'])])
    right_col_data.append([Paragraph(str(project['future_activities']), styles['PMOBodyText'])])
    
    # Create left and right tables
    left_table = Table(left_col_data, colWidths=[4*inch])
    left_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
    ]))
    
    right_table = Table(right_col_data, colWidths=[4.5*inch])
    right_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
    ]))
    
    # Combine into two-column layout
    main_layout = Table([[left_table, right_table]], colWidths=[4.2*inch, 4.8*inch])
    main_layout.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    elements.append(main_layout)
    elements.append(Spacer(1, 15))
    
    # Risks and Issues section
    elements.append(Paragraph("<font color='#E07020'><b>Risks & Issues</b></font>", styles['SectionHeader']))
    
    risks_data = [
        [
            Paragraph("<b>Type</b>", styles['TableText']),
            Paragraph("<b>Description</b>", styles['TableText']),
            Paragraph("<b>Mitigation / Action</b>", styles['TableText']),
        ],
        [
            Paragraph("Issues", styles['TableText']),
            Paragraph(str(project['issues']), styles['TableText']),
            Paragraph("[To be defined]", styles['TableText']),
        ],
        [
            Paragraph("Risks", styles['TableText']),
            Paragraph(str(project['risks']), styles['TableText']),
            Paragraph("[To be defined]", styles['TableText']),
        ]
    ]
    
    risks_table = Table(risks_data, colWidths=[1*inch, 4*inch, 3.5*inch])
    risks_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.15, 0.15, 0.15)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, -1), colors.Color(0.25, 0.25, 0.25)),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.4, 0.4, 0.4)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(risks_table)
    elements.append(Spacer(1, 15))
    
    # Comments section
    if project['comments'] and project['comments'] != '[To be provided]':
        elements.append(Paragraph("<font color='#E07020'><b>Comments / Notes</b></font>", styles['SectionHeader']))
        elements.append(Paragraph(str(project['comments']), styles['PMOBodyText']))
    
    # Deliverables placeholder
    elements.append(Spacer(1, 15))
    elements.append(Paragraph("<font color='#E07020'><b>Deliverables / Milestones</b></font>", styles['SectionHeader']))
    
    deliverables_data = [
        [
            Paragraph("<b>Deliverable Name</b>", styles['TableText']),
            Paragraph("<b>Contractual Date</b>", styles['TableText']),
            Paragraph("<b>Planned Date</b>", styles['TableText']),
            Paragraph("<b>% Done</b>", styles['TableText']),
            Paragraph("<b>Status</b>", styles['TableText']),
        ],
        [
            Paragraph("[To be added manually]", styles['TableText']),
            Paragraph("-", styles['TableText']),
            Paragraph("-", styles['TableText']),
            Paragraph("-", styles['TableText']),
            Paragraph("-", styles['TableText']),
        ]
    ]
    
    deliverables_table = Table(deliverables_data, colWidths=[3*inch, 1.5*inch, 1.5*inch, 1*inch, 1.5*inch])
    deliverables_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.15, 0.15, 0.15)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, -1), colors.Color(0.25, 0.25, 0.25)),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.4, 0.4, 0.4)),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (2, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(deliverables_table)
    
    return elements


def generate_pdf_report(projects, output_path):
    """Generate a multi-page PDF report for all projects."""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )
    
    styles = create_styles()
    all_elements = []
    
    for i, project in enumerate(projects):
        # Add page content
        page_elements = create_project_report_page(project, styles)
        all_elements.extend(page_elements)
        
        # Add page break except for last project
        if i < len(projects) - 1:
            all_elements.append(PageBreak())
    
    doc.build(all_elements)
    return output_path
