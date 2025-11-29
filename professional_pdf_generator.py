"""Professional PDF Report Generation Module for PMO Reports"""

from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape, letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, Image, KeepTogether, PageTemplate,
    Frame, BaseDocTemplate, NextPageTemplate
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.graphics.shapes import Drawing, Line, Rect
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.legends import Legend
from reportlab.graphics import renderPDF
from reportlab.pdfgen import canvas
import os


def create_professional_styles():
    """Create professional paragraph styles for executive report."""
    styles = getSampleStyleSheet()
    
    # Cover page title
    styles.add(ParagraphStyle(
        name='CoverTitle',
        parent=styles['Title'],
        fontSize=32,
        textColor=colors.HexColor('#1B2951'),  # Dark blue
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    ))
    
    # Cover subtitle
    styles.add(ParagraphStyle(
        name='CoverSubtitle',
        parent=styles['Normal'],
        fontSize=18,
        textColor=colors.HexColor('#4A5568'),
        spaceBefore=20,
        spaceAfter=20,
        alignment=TA_CENTER
    ))
    
    # Section title (orange accent)
    styles.add(ParagraphStyle(
        name='SectionTitle',
        parent=styles['Heading1'],
        fontSize=20,
        textColor=colors.HexColor('#E07020'),
        spaceBefore=20,
        spaceAfter=15,
        fontName='Helvetica-Bold',
        borderWidth=2,
        borderColor=colors.HexColor('#E07020'),
        borderPadding=5
    ))
    
    # Subsection header
    styles.add(ParagraphStyle(
        name='SubsectionHeader',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#1B2951'),
        spaceBefore=12,
        spaceAfter=8,
        fontName='Helvetica-Bold'
    ))
    
    # Body text
    styles.add(ParagraphStyle(
        name='ProfessionalBody',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#2D3748'),
        leading=14,
        alignment=TA_JUSTIFY,
        spaceBefore=4,
        spaceAfter=4
    ))
    
    # Metric label
    styles.add(ParagraphStyle(
        name='MetricLabel',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#718096'),
        spaceAfter=2
    ))
    
    # Metric value
    styles.add(ParagraphStyle(
        name='MetricValue',
        parent=styles['Normal'],
        fontSize=14,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#1B2951'),
        spaceAfter=8
    ))
    
    # Table header
    styles.add(ParagraphStyle(
        name='TableHeader',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica-Bold',
        textColor=colors.white,
        alignment=TA_CENTER
    ))
    
    # Risk/Issue text
    styles.add(ParagraphStyle(
        name='RiskText',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#2D3748'),
        leading=11
    ))
    
    return styles


def add_header_footer(canvas, doc):
    """Add professional header and footer to each page."""
    canvas.saveState()
    
    # Header
    canvas.setFillColor(colors.HexColor('#1B2951'))
    canvas.setFont('Helvetica-Bold', 10)
    canvas.drawString(inch, doc.height + inch + 0.5*inch, "PMO Status Report")
    
    canvas.setFont('Helvetica', 9)
    canvas.setFillColor(colors.HexColor('#718096'))
    canvas.drawRightString(doc.width + inch, doc.height + inch + 0.5*inch, 
                           datetime.now().strftime('%B %Y'))
    
    # Header line
    canvas.setStrokeColor(colors.HexColor('#E07020'))
    canvas.setLineWidth(2)
    canvas.line(inch, doc.height + inch + 0.4*inch, 
                doc.width + inch, doc.height + inch + 0.4*inch)
    
    # Footer
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(colors.HexColor('#718096'))
    canvas.drawString(inch, 0.5*inch, "Confidential - Internal Use Only")
    
    # Page number
    page_num = canvas.getPageNumber()
    canvas.drawRightString(doc.width + inch, 0.5*inch, f"Page {page_num}")
    
    # Footer line
    canvas.setStrokeColor(colors.HexColor('#CBD5E0'))
    canvas.setLineWidth(1)
    canvas.line(inch, 0.7*inch, doc.width + inch, 0.7*inch)
    
    canvas.restoreState()


def create_cover_page(styles):
    """Create a professional cover page."""
    elements = []
    
    # Add some space at top
    elements.append(Spacer(1, 2*inch))
    
    # Main title
    elements.append(Paragraph(
        "PMO PROJECT STATUS REPORT",
        styles['CoverTitle']
    ))
    
    # Subtitle with month/year
    elements.append(Paragraph(
        datetime.now().strftime('%B %Y'),
        styles['CoverSubtitle']
    ))
    
    # Divider line
    elements.append(Spacer(1, 0.5*inch))
    elements.append(HRFlowable(
        width="60%",
        thickness=3,
        color=colors.HexColor('#E07020'),
        hAlign='CENTER'
    ))
    elements.append(Spacer(1, 0.5*inch))
    
    # Organization info
    elements.append(Paragraph(
        "Strategic Projects & Leadership Development",
        styles['CoverSubtitle']
    ))
    
    elements.append(Spacer(1, 2*inch))
    
    # Document info box
    info_data = [
        ['Document Type:', 'Executive Summary Report'],
        ['Reporting Period:', datetime.now().strftime('%B %Y')],
        ['Generated:', datetime.now().strftime('%d %B %Y, %H:%M')],
        ['Classification:', 'Internal - Confidential']
    ]
    
    info_table = Table(info_data, colWidths=[2.5*inch, 3*inch])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F7FAFC')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#718096')),
        ('TEXTCOLOR', (1, 0), (1, -1), colors.HexColor('#2D3748')),
        ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
        ('FONT', (1, 0), (1, -1), 'Helvetica', 10),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E0')),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.white, colors.HexColor('#F7FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('LEFTPADDING', (0, 0), (-1, -1), 15),
    ]))
    elements.append(info_table)
    
    elements.append(PageBreak())
    return elements


def create_executive_summary(projects, styles):
    """Create executive summary dashboard."""
    elements = []
    
    elements.append(Paragraph("Executive Summary", styles['SectionTitle']))
    elements.append(Spacer(1, 10))
    
    # Calculate summary metrics
    total_projects = len(projects)
    on_track = sum(1 for p in projects if 'on track' in str(p.get('health', '')).lower())
    at_risk = sum(1 for p in projects if 'at risk' in str(p.get('health', '')).lower())
    off_track = total_projects - on_track - at_risk
    
    total_budget = sum(p.get('budget_total', 0) for p in projects)
    total_spent = sum(p.get('budget_spent', 0) for p in projects)
    avg_progress = sum(p.get('timeline_actual', 0) for p in projects) / len(projects) if projects else 0
    
    # Key Metrics Grid
    metrics_data = [
        [
            Paragraph("Total Projects", styles['MetricLabel']),
            Paragraph("Projects On Track", styles['MetricLabel']),
            Paragraph("Projects At Risk", styles['MetricLabel']),
            Paragraph("Projects Off Track", styles['MetricLabel'])
        ],
        [
            Paragraph(str(total_projects), styles['MetricValue']),
            Paragraph(str(on_track), styles['MetricValue']),
            Paragraph(str(at_risk), styles['MetricValue']),
            Paragraph(str(off_track), styles['MetricValue'])
        ]
    ]
    
    metrics_table = Table(metrics_data, colWidths=[2.25*inch]*4)
    metrics_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F7FAFC')),
        ('BACKGROUND', (0, 1), (-1, 1), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0, 0), (-1, -1), 15),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
    ]))
    elements.append(metrics_table)
    elements.append(Spacer(1, 20))
    
    # Budget Overview
    elements.append(Paragraph("Budget Overview", styles['SubsectionHeader']))
    
    budget_data = [
        ['Total Budget', f"{total_budget:,.0f} SAR"],
        ['Spent to Date', f"{total_spent:,.0f} SAR"],
        ['Remaining', f"{total_budget - total_spent:,.0f} SAR"],
        ['Utilization', f"{(total_spent/total_budget*100):.1f}%" if total_budget > 0 else "0%"]
    ]
    
    budget_table = Table(budget_data, colWidths=[2*inch, 3*inch])
    budget_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#FFF5EB')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#92400E')),
        ('TEXTCOLOR', (1, 0), (1, -1), colors.HexColor('#1B2951')),
        ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
        ('FONT', (1, 0), (1, -1), 'Helvetica', 10),
        ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LINEBELOW', (0, 0), (-1, -2), 0.5, colors.HexColor('#FED7AA')),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (0, -1), 15),
        ('LEFTPADDING', (1, 0), (1, -1), 10),
    ]))
    elements.append(budget_table)
    
    elements.append(PageBreak())
    return elements


def create_project_detail_page(project, styles, index):
    """Create detailed page for individual project."""
    elements = []
    
    # Project header
    elements.append(Paragraph(
        f"Project {index}: {project['name']}",
        styles['SectionTitle']
    ))
    
    # Status badge
    health = str(project.get('health', 'Unknown')).upper()
    health_color = colors.HexColor('#10B981') if 'ON TRACK' in health else \
                  colors.HexColor('#F59E0B') if 'AT RISK' in health else \
                  colors.HexColor('#EF4444')
    
    status_data = [[Paragraph(f"<b>STATUS: {health}</b>", 
                              ParagraphStyle('StatusBadge', 
                                           fontSize=12,
                                           textColor=colors.white,
                                           alignment=TA_CENTER))]]
    
    status_table = Table(status_data, colWidths=[2*inch])
    status_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), health_color),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (0, 0), 8),
        ('BOTTOMPADDING', (0, 0), (0, 0), 8),
        ('LEFTPADDING', (0, 0), (0, 0), 20),
        ('RIGHTPADDING', (0, 0), (0, 0), 20),
    ]))
    elements.append(status_table)
    elements.append(Spacer(1, 15))
    
    # Two-column layout for project details
    details_data = [
        [
            Paragraph("<b>Project Information</b>", styles['SubsectionHeader']),
            Paragraph("<b>Key Metrics</b>", styles['SubsectionHeader'])
        ],
        [
            Table([
                ['Category:', project.get('category', 'N/A')],
                ['Sponsor GM:', project.get('gm', 'TBD')],
                ['Director:', project.get('director', 'TBD')],
                ['Project Lead:', project.get('operational_lead', 'TBD')],
                ['Vendor:', str(project.get('vendor', 'TBD'))[:30]],
            ], colWidths=[1.5*inch, 2.5*inch]),
            
            Table([
                ['End Date:', project.get('contract_end_date', 'TBD')],
                ['Days Remaining:', f"{project.get('days_remaining', 0)} days"],
                ['Progress:', f"Actual: {project.get('timeline_actual', 0):.0f}% | Plan: {project.get('timeline_planned', 0):.0f}%"],
                ['Budget Used:', f"{project.get('budget_spent_pct', 0):.1f}%"],
                ['Budget:', f"{project.get('budget_total', 0):,.0f} SAR"],
            ], colWidths=[1.5*inch, 2.5*inch])
        ]
    ]
    
    # Style for detail tables
    detail_style = TableStyle([
        ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 9),
        ('FONT', (1, 0), (1, -1), 'Helvetica', 9),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#718096')),
        ('TEXTCOLOR', (1, 0), (1, -1), colors.HexColor('#2D3748')),
        ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ])
    
    details_data[1][0].setStyle(detail_style)
    details_data[1][1].setStyle(detail_style)
    
    main_table = Table(details_data, colWidths=[4.5*inch, 4.5*inch])
    main_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#F7FAFC')),
        ('BOX', (0, 1), (-1, 1), 1, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0, 0), (-1, 0), 0),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('LEFTPADDING', (0, 1), (-1, 1), 10),
        ('RIGHTPADDING', (0, 1), (-1, 1), 10),
        ('TOPPADDING', (0, 1), (-1, 1), 10),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 10),
    ]))
    
    elements.append(main_table)
    elements.append(Spacer(1, 20))
    
    # Activities Section
    elements.append(Paragraph("Current Activities", styles['SubsectionHeader']))
    activities_text = project.get('current_activities', '[To be provided]')
    elements.append(Paragraph(activities_text, styles['ProfessionalBody']))
    elements.append(Spacer(1, 15))
    
    # Risks and Issues
    elements.append(Paragraph("Risks & Issues", styles['SubsectionHeader']))
    
    risk_data = [
        [Paragraph("<b>Type</b>", styles['TableHeader']),
         Paragraph("<b>Description</b>", styles['TableHeader']),
         Paragraph("<b>Impact</b>", styles['TableHeader'])],
        ['Issues', project.get('issues', '[None reported]')[:200], 'Medium'],
        ['Risks', project.get('risks', '[None identified]')[:200], 'Low']
    ]
    
    risk_table = Table(risk_data, colWidths=[1*inch, 5.5*inch, 1*inch])
    risk_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1B2951')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#2D3748')),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
        ('FONT', (0, 1), (-1, -1), 'Helvetica', 9),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (2, 1), (2, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
    ]))
    
    elements.append(risk_table)
    
    return elements


def generate_professional_pdf(projects, output_path):
    """Generate professional executive-style PDF report."""
    # Create document
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(letter),
        rightMargin=inch,
        leftMargin=inch,
        topMargin=inch + 0.5*inch,  # Extra space for header
        bottomMargin=inch,
        title="PMO Status Report",
        author="PMO Office"
    )
    
    # Get styles
    styles = create_professional_styles()
    
    # Build content
    elements = []
    
    # Cover page
    elements.extend(create_cover_page(styles))
    
    # Executive summary
    elements.extend(create_executive_summary(projects, styles))
    
    # Individual project pages (limit to first 5 for sample)
    for i, project in enumerate(projects[:5], 1):
        elements.extend(create_project_detail_page(project, styles, i))
        if i < min(5, len(projects)):
            elements.append(PageBreak())
    
    # Build PDF with header/footer
    doc.build(elements, onFirstPage=add_header_footer, onLaterPages=add_header_footer)
    
    return output_path
