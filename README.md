# PMO Report Generator

A professional web application that automatically generates comprehensive project status reports from Excel PMO trackers.

## Features

- **Automatic Column Detection**: Intelligently maps Excel columns with flexible naming conventions
- **Professional PDF Reports**: Generates landscape-oriented reports with modern styling
- **Comprehensive Metrics**: 
  - Budget utilization calculations and visualization
  - Timeline progress tracking with variance analysis
  - Days remaining calculations
  - Health status indicators with color coding
- **Batch Processing**: Creates both individual and combined reports
- **Smart Data Parsing**: Handles various formats for numbers, dates, and percentages
- **Visual Indicators**: Color-coded health status (Green: On Track, Orange: At Risk, Red: Off Track)

## Installation

1. Clone or download this repository

2. Install Python dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python app.py
```

4. Open your browser and navigate to: `http://localhost:5000`

## Excel File Requirements

Your Excel file should contain columns for the following information (flexible naming is supported):

### Required Columns
- **Project Name** (Critical - must be present)

### Recommended Columns
- **Project Number**: #, No, Number, Project #, ID
- **Project Category**: Category, Type
- **Project Status**: Status
- **GM**: General Manager, Sponsor GM
- **Director**: SPLD Director, Project Director
- **Operational Lead**: Project Lead, Project Manager
- **Contract End Date**: End Date, Finish Date, Deadline
- **Days Remaining**: Days Left
- **Budget Spent**: Actual Cost
- **Budget Remaining**: Budget Left
- **Timeline Actual**: Actual Progress, Progress Actual (%)
- **Timeline Planned**: Planned Progress, Progress Planned (%)
- **Service Delivery KPI**: KPI, Performance KPI
- **Project Health**: Health, Status Health (on track/at risk/off track)
- **Issues**: Current Issues
- **Risks**: Project Risks
- **Current Activities**: Current Work
- **Future Activities**: Next Steps, Upcoming Activities
- **Comments**: Notes, Owner Comments
- **Vendor**: Supplier, Contractor

## Usage

1. **Prepare Your Excel File**
   - Ensure your PMO tracker Excel file contains project information
   - Column names are flexible - the system will attempt to match them automatically

2. **Upload File**
   - Click the upload area or drag and drop your Excel file
   - Maximum file size: 50MB
   - Supported formats: .xlsx, .xls

3. **Generate Reports**
   - Click "Generate Reports" button
   - Wait for processing to complete

4. **Download Reports**
   - Download the ZIP file containing:
     - Combined PDF with all projects
     - Individual PDF reports for each project

## Report Contents

Each project report includes:

### Header Section
- Project name and category
- Report generation date
- Vendor information

### Key Metrics
- Project Sponsor GM
- Project Director
- Operational Lead
- Contract end date and days remaining

### Project Timeline
- Actual vs Planned progress
- Schedule variance with color indicators

### Budget Utilization
- Total budget
- Spent amount and percentage
- Remaining budget and percentage

### Project Health
- Visual health indicator (color-coded)
- Service delivery KPI

### Activities
- Current activities
- Future/planned activities

### Risks & Issues
- Current issues with descriptions
- Identified risks
- Space for mitigation actions

### Deliverables
- Placeholder section for manual milestone tracking

## Technical Details

- **Backend**: Flask (Python)
- **PDF Generation**: ReportLab
- **Data Processing**: Pandas & NumPy
- **Frontend**: HTML5, CSS3, JavaScript
- **File Handling**: In-memory processing for security

## Troubleshooting

### Common Issues

1. **"Missing critical columns" error**
   - Ensure your Excel file has a "Project Name" column
   - Check column naming matches the supported variations

2. **Date parsing issues**
   - Dates should be in standard Excel date format
   - Supported formats: DD/MM/YYYY, MM/DD/YYYY, YYYY-MM-DD

3. **Number formatting**
   - Numbers can include currency symbols (SAR, $)
   - Thousands separators (,) are automatically handled
   - Percentages can be written as 40% or 0.4

## Security Notes

- Files are processed in memory
- Temporary files are created in system temp directory
- No data is permanently stored on the server
- Maximum file size limit: 50MB

## License

This project is for internal PMO use. All rights reserved.

## Support

For issues or feature requests, please contact your IT administrator.
