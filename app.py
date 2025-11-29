"""
PMO Report Generator
Automatically generates project status reports from Excel PMO tracker.
Uses LLM for interpretation only - all calculations are done programmatically.
"""

import os
import io
import json
import zipfile
from datetime import datetime, timedelta
from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Import helper modules
from pmo_helpers import *
from pmo_report_generator import *

# Flask Routes
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    # Check file extension
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400
    
    try:
        file_content = file.read()
        projects, error = process_excel_file(file_content, file.filename)
        
        if error:
            return jsonify({'error': error}), 400
        
        # Generate reports
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Create output directory
        output_dir = f'/tmp/pmo_reports_{timestamp}'
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate combined PDF
        pdf_path = os.path.join(output_dir, f'PMO_Project_Reports_{timestamp}.pdf')
        generate_pdf_report(projects, pdf_path)
        
        # Generate individual project PDFs
        individual_dir = os.path.join(output_dir, 'individual_reports')
        os.makedirs(individual_dir, exist_ok=True)
        
        for project in projects:
            safe_name = "".join(c for c in str(project['name'])[:50] if c.isalnum() or c in (' ', '-', '_')).strip()
            safe_name = safe_name.replace(' ', '_')
            individual_pdf = os.path.join(individual_dir, f'{safe_name}_Report.pdf')
            generate_pdf_report([project], individual_pdf)
        
        # Create ZIP file with all reports
        zip_path = f'/tmp/PMO_Reports_{timestamp}.zip'
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add combined report
            zipf.write(pdf_path, os.path.basename(pdf_path))
            
            # Add individual reports
            for root, dirs, files in os.walk(individual_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.join('individual_reports', file)
                    zipf.write(file_path, arcname)
        
        # Return success with download info
        return jsonify({
            'success': True,
            'message': f'Successfully generated reports for {len(projects)} projects',
            'project_count': len(projects),
            'download_url': f'/download/{timestamp}'
        })
        
    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500


@app.route('/download/<timestamp>')
def download_file(timestamp):
    zip_path = f'/tmp/PMO_Reports_{timestamp}.zip'
    if os.path.exists(zip_path):
        return send_file(
            zip_path,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'PMO_Reports_{timestamp}.zip'
        )
    return jsonify({'error': 'File not found'}), 404


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
