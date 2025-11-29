"""SPLD Project Health Report Generator - Executive Dashboard Style"""

from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, Image, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.graphics.shapes import Drawing, Rect, Circle
from reportlab.graphics import renderPDF
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.pdfgen import canvas


def create_spld_styles():
    """Create SPLD branded styles."""
    styles = getSampleStyleSheet()
    
    # SPLD Title
    styles.add(ParagraphStyle(
        name='SPLDTitle',
        parent=styles['Title'],
        fontSize=28,
        textColor=colors.HexColor('#EA6A1F'),  # SPLD Orange
        spaceAfter=20,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    ))
    
    # Dashboard Header
    styles.add(ParagraphStyle(
        name='DashboardHeader',
        fontSize=16,
        textColor=colors.HexColor('#1B2951'),  # Dark Blue
        spaceAfter=12,
        alignment=TA_LEFT,
        fontName='Helvetica-Bold'
    ))
    
    # Metric Title
    styles.add(ParagraphStyle(
        name='MetricTitle',
        fontSize=10,
        textColor=colors.HexColor('#666666'),
        spaceAfter=4,
        alignment=TA_CENTER
    ))
    
    # Big Number
    styles.add(ParagraphStyle(
        name='BigNumber',
        fontSize=24,
        textColor=colors.HexColor('#1B2951'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    ))
    
    # Status Badge
    styles.add(ParagraphStyle(
        name='StatusBadge',
        fontSize=11,
        textColor=colors.white,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    ))
    
    return styles


def create_health_indicator(health_status):
    """Create a visual health indicator (traffic light)."""
    d = Drawing(30, 30)
    
    # Determine color based on status
    if 'on track' in health_status.lower():
        color = colors.HexColor('#10B981')  # Green
    elif 'at risk' in health_status.lower():
        color = colors.HexColor('#F59E0B')  # Amber
    else:
        color = colors.HexColor('#EF4444')  # Red
    
    # Create circle
    circle = Circle(15, 15, 12)
    circle.fillColor = color
    circle.strokeColor = color
    d.add(circle)
    
    return d


def create_spld_dashboard(projects, styles):
    """Create SPLD executive dashboard page."""
    elements = []
    
    # Title Section
    title_data = [
        [Paragraph("STRATEGIC PROJECTS & LEADERSHIP DEVELOPMENT", styles['SPLDTitle'])],
        [Paragraph("Project Health Dashboard", styles['DashboardHeader'])],
        [Paragraph(f"Report Date: {datetime.now().strftime('%d %B %Y')}", styles['Normal'])]
    ]
    
    title_table = Table(title_data, colWidths=[10*inch])
    title_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (0, 0), 20),
        ('BOTTOMPADDING', (0, 1), (0, 1), 10),
    ]))
    
    elements.append(title_table)
    elements.append(Spacer(1, 20))
    
    # Calculate metrics
    total = len(projects)
    on_track = sum(1 for p in projects if 'on track' in str(p.get('health', '')).lower())
    at_risk = sum(1 for p in projects if 'at risk' in str(p.get('health', '')).lower())
    off_track = total - on_track - at_risk
    
    # Portfolio Health Overview
    health_header = [[Paragraph("PORTFOLIO HEALTH OVERVIEW", styles['DashboardHeader'])]]
    elements.append(Table(health_header, colWidths=[10*inch]))
    elements.append(Spacer(1, 10))
    
    # Health Status Cards
    health_cards_data = [
        [
            Paragraph("Total Projects", styles['MetricTitle']),
            Paragraph("On Track", styles['MetricTitle']),
            Paragraph("At Risk", styles['MetricTitle']),
            Paragraph("Off Track", styles['MetricTitle'])
        ],
        [
            Paragraph(str(total), styles['BigNumber']),
            Paragraph(str(on_track), styles['BigNumber']),
            Paragraph(str(at_risk), styles['BigNumber']),
            Paragraph(str(off_track), styles['BigNumber'])
        ],
        [
            Paragraph("100%", styles['MetricTitle']),
            Paragraph(f"{(on_track/total*100):.0f}%", styles['MetricTitle']),
            Paragraph(f"{(at_risk/total*100):.0f}%", styles['MetricTitle']),
            Paragraph(f"{(off_track/total*100):.0f}%", styles['MetricTitle'])
        ]
    ]
    
    health_table = Table(health_cards_data, colWidths=[2.5*inch]*4)
    health_table.setStyle(TableStyle([
        # Header row
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F3F4F6')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#374151')),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        
        # Numbers row
        ('FONTSIZE', (0, 1), (-1, 1), 28),
        ('TEXTCOLOR', (0, 1), (0, 1), colors.HexColor('#1B2951')),
        ('TEXTCOLOR', (1, 1), (1, 1), colors.HexColor('#10B981')),
        ('TEXTCOLOR', (2, 1), (2, 1), colors.HexColor('#F59E0B')),
        ('TEXTCOLOR', (3, 1), (3, 1), colors.HexColor('#EF4444')),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 10),
        
        # Percentage row
        ('TEXTCOLOR', (0, 2), (-1, 2), colors.HexColor('#6B7280')),
        ('FONTSIZE', (0, 2), (-1, 2), 10),
        
        # Borders and alignment
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#E5E7EB')),
        ('BOX', (0, 0), (-1, -1), 2, colors.HexColor('#EA6A1F')),
        ('TOPPADDING', (0, 0), (-1, -1), 15),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
    ]))
    
    elements.append(health_table)
    elements.append(Spacer(1, 30))
    
    # Key Metrics Summary
    elements.append(Paragraph("KEY PERFORMANCE METRICS", styles['DashboardHeader']))
    elements.append(Spacer(1, 10))
    
    # Calculate budget metrics
    total_budget = sum(p.get('budget_total', 0) for p in projects)
    total_spent = sum(p.get('budget_spent', 0) for p in projects)
    avg_progress = sum(p.get('timeline_actual', 0) for p in projects) / len(projects) if projects else 0
    avg_planned = sum(p.get('timeline_planned', 0) for p in projects) / len(projects) if projects else 0
    
    metrics_data = [
        ['Metric', 'Value', 'Status'],
        ['Total Portfolio Budget', f"{total_budget:,.0f} SAR", "Active"],
        ['Budget Utilized', f"{total_spent:,.0f} SAR ({(total_spent/total_budget*100):.1f}%)", 
         "On Track" if total_spent/total_budget < 0.8 else "Warning"],
        ['Average Project Progress', f"{avg_progress:.1f}%", 
         "On Track" if avg_progress >= avg_planned else "Behind"],
        ['Average Planned Progress', f"{avg_planned:.1f}%", "Baseline"],
        ['Schedule Performance', f"{(avg_progress - avg_planned):+.1f}%", 
         "Good" if avg_progress >= avg_planned else "Needs Attention"]
    ]
    
    metrics_table = Table(metrics_data, colWidths=[3*inch, 4*inch, 2*inch])
    metrics_table.setStyle(TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EA6A1F')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        
        # Data rows
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#374151')),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        
        # Alternating row colors
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9FAFB')]),
        
        # Status column coloring
        ('TEXTCOLOR', (2, 1), (2, -1), colors.HexColor('#059669')),
        ('FONTNAME', (2, 1), (2, -1), 'Helvetica-Bold'),
        
        # Borders
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#E5E7EB')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
    ]))
    
    elements.append(metrics_table)
    
    return elements


def create_project_list_page(projects, styles):
    """Create detailed project list with health indicators."""
    elements = []
    
    elements.append(PageBreak())
    elements.append(Paragraph("PROJECT DETAILS", styles['DashboardHeader']))
    elements.append(Spacer(1, 10))
    
    # Project list table headers
    headers = ['#', 'Project Name', 'Health', 'Progress', 'Budget Used', 'End Date', 'Days Left']
    
    # Build data rows
    data = [headers]
    
    for i, project in enumerate(projects[:15], 1):  # Limit to 15 for single page
        health = project.get('health', 'Unknown')
        
        # Determine health symbol and color
        if 'on track' in health.lower():
            health_symbol = '●'
            health_color = colors.HexColor('#10B981')
        elif 'at risk' in health.lower():
            health_symbol = '●'
            health_color = colors.HexColor('#F59E0B')
        else:
            health_symbol = '●'
            health_color = colors.HexColor('#EF4444')
        
        row = [
            str(i),
            project.get('name', 'Unknown')[:40],
            health_symbol,
            f"{project.get('timeline_actual', 0):.0f}%",
            f"{project.get('budget_spent_pct', 0):.0f}%",
            project.get('contract_end_date', 'TBD'),
            str(project.get('days_remaining', 0))
        ]
        data.append(row)
    
    # Create table
    col_widths = [0.5*inch, 3.5*inch, 0.8*inch, 1*inch, 1*inch, 1.2*inch, 1*inch]
    project_table = Table(data, colWidths=col_widths)
    
    # Apply styling
    table_style = [
        # Header row
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1B2951')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        
        # Data rows
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (2, 1), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        
        # Grid
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9FAFB')]),
        
        # Padding
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]
    
    # Apply health colors
    for i in range(1, len(data)):
        health = projects[i-1].get('health', 'Unknown')
        if 'on track' in health.lower():
            color = colors.HexColor('#10B981')
        elif 'at risk' in health.lower():
            color = colors.HexColor('#F59E0B')
        else:
            color = colors.HexColor('#EF4444')
        
        table_style.append(('TEXTCOLOR', (2, i), (2, i), color))
        table_style.append(('FONTSIZE', (2, i), (2, i), 14))
    
    project_table.setStyle(TableStyle(table_style))
    elements.append(project_table)
    
    # Add legend
    elements.append(Spacer(1, 20))
    legend_data = [
        [Paragraph("<b>Legend:</b>", styles['Normal']), '', '', ''],
        ['', '● On Track', '● At Risk', '● Off Track / Delayed']
    ]
    
    legend_table = Table(legend_data, colWidths=[1*inch, 2*inch, 2*inch, 2*inch])
    legend_table.setStyle(TableStyle([
        ('TEXTCOLOR', (1, 1), (1, 1), colors.HexColor('#10B981')),
        ('TEXTCOLOR', (2, 1), (2, 1), colors.HexColor('#F59E0B')),
        ('TEXTCOLOR', (3, 1), (3, 1), colors.HexColor('#EF4444')),
        ('FONTSIZE', (1, 1), (3, 1), 12),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(legend_table)
    
    return elements


def generate_spld_health_report(projects, output_path):
    """Generate SPLD Project Health Report."""
    # Create document
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        rightMargin=0.75*inch,
        leftMargin=0.75*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch,
        title="SPLD Project Health Report",
        author="Strategic Projects & Leadership Development"
    )
    
    # Get styles
    styles = create_spld_styles()
    
    # Build content
    elements = []
    
    # Dashboard page
    elements.extend(create_spld_dashboard(projects, styles))
    
    # Project list page
    elements.extend(create_project_list_page(projects, styles))
    
    # Build PDF
    doc.build(elements)
    
    return output_path
