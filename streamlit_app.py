"""
SPLD PMO Report Generator - Streamlit App
Upload Excel file to generate professional SPLD reports
"""

import streamlit as st

# Page configuration MUST be first Streamlit command
st.set_page_config(
    page_title="SPLD PMO Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

import pandas as pd
import io
import zipfile
from datetime import datetime
import os
import tempfile
from dotenv import load_dotenv

# Load environment variables (for local development)
load_dotenv()

# For Streamlit Cloud, use st.secrets for API keys (if available)
try:
    if "OPENAI_API_KEY" in st.secrets:
        os.environ['OPENAI_API_KEY'] = st.secrets["OPENAI_API_KEY"]
except (FileNotFoundError, AttributeError):
    # No secrets file - this is normal for local development
    pass

# Import our modules
from pmo_helpers import process_excel_file
from spld_exact_format import generate_spld_exact_report
from spld_word_generator import create_spld_word_report
from excel_generator import create_excel_report
from llm_integration import format_project_text

# Custom CSS for SPLD branding
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #1a1a2e 0%, #0f3460 100%);
    }
    .main-header {
        color: #EA6A1F;
        font-size: 48px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 10px;
    }
    .sub-header {
        color: #ffffff;
        font-size: 20px;
        text-align: center;
        margin-bottom: 30px;
    }
    .feature-box {
        background-color: rgba(42, 42, 50, 0.8);
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        border-left: 4px solid #EA6A1F;
    }
    .stButton > button {
        background-color: #EA6A1F;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        border: none;
        padding: 10px 24px;
        font-size: 16px;
        width: 100%;
    }
    .stButton > button:hover {
        background-color: #C55A1A;
    }
    .success-message {
        background-color: #27AE60;
        color: white;
        padding: 10px;
        border-radius: 5px;
        margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<h1 class="main-header">📊 SPLD PMO Report Generator</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload your PMO Excel tracker to generate executive dashboard reports</p>', unsafe_allow_html=True)

# Create columns for layout
col1, col2 = st.columns([2, 1])

with col1:
    # File upload section
    st.markdown("### 📁 Upload Excel File")
    uploaded_file = st.file_uploader(
        "Choose your PMO tracker Excel file",
        type=['xlsx', 'xls'],
        help="Upload an Excel file containing your PMO project data"
    )
    
    if uploaded_file is not None:
        # Show file info
        st.success(f"✅ File uploaded: **{uploaded_file.name}** ({uploaded_file.size:,} bytes)")
        
        # Process button
        if st.button("🚀 Generate Reports", type="primary"):
            with st.spinner("Processing your Excel file..."):
                try:
                    # Read the uploaded file
                    file_content = uploaded_file.read()
                    
                    # Process the Excel file
                    projects, error = process_excel_file(file_content, uploaded_file.name)
                    
                    if error:
                        st.error(f"❌ Error processing file: {error}")
                    else:
                        st.success(f"✅ Successfully extracted **{len(projects)} projects** from Excel")
                        
                        # Show progress
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # Apply LLM formatting if available
                        status_text.text("🤖 Applying AI text formatting...")
                        progress_bar.progress(20)
                        
                        # Check if LLM is configured
                        from llm_integration import get_llm_formatter
                        llm_formatter = get_llm_formatter()
                        
                        if llm_formatter:
                            st.info(f"✅ AI Active: Using {llm_formatter.provider} ({llm_formatter.model})")
                            for i in range(len(projects)):
                                projects[i] = format_project_text(projects[i])
                        else:
                            st.warning("ℹ️ AI text formatting not configured - using original text")
                        
                        # Generate reports
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        temp_dir = tempfile.mkdtemp()
                        
                        # Generate PDF
                        status_text.text("📄 Generating PDF report...")
                        progress_bar.progress(40)
                        pdf_path = os.path.join(temp_dir, f'SPLD_Report_{timestamp}.pdf')
                        generate_spld_exact_report(projects, pdf_path)
                        
                        # Generate Word
                        status_text.text("📝 Generating Word document...")
                        progress_bar.progress(60)
                        word_path = os.path.join(temp_dir, f'SPLD_Report_{timestamp}.docx')
                        create_spld_word_report(projects, word_path)
                        
                        # Generate Excel
                        status_text.text("📊 Generating Excel dashboard...")
                        progress_bar.progress(80)
                        excel_path = os.path.join(temp_dir, f'SPLD_Dashboard_{timestamp}.xlsx')
                        create_excel_report(projects, excel_path)
                        
                        # Create ZIP file
                        status_text.text("📦 Creating ZIP package...")
                        progress_bar.progress(90)
                        zip_path = os.path.join(temp_dir, f'SPLD_Reports_{timestamp}.zip')
                        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                            zipf.write(pdf_path, os.path.basename(pdf_path))
                            zipf.write(word_path, os.path.basename(word_path))
                            zipf.write(excel_path, os.path.basename(excel_path))
                        
                        progress_bar.progress(100)
                        status_text.text("✅ Reports generated successfully!")
                        
                        # Create download buttons
                        st.markdown("### 📥 Download Your Reports")
                        
                        col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)
                        
                        with col_dl1:
                            with open(pdf_path, 'rb') as f:
                                st.download_button(
                                    label="📄 Download PDF",
                                    data=f.read(),
                                    file_name=f'SPLD_Report_{timestamp}.pdf',
                                    mime='application/pdf'
                                )
                        
                        with col_dl2:
                            with open(word_path, 'rb') as f:
                                st.download_button(
                                    label="📝 Download Word",
                                    data=f.read(),
                                    file_name=f'SPLD_Report_{timestamp}.docx',
                                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                                )
                        
                        with col_dl3:
                            with open(excel_path, 'rb') as f:
                                st.download_button(
                                    label="📊 Download Excel",
                                    data=f.read(),
                                    file_name=f'SPLD_Dashboard_{timestamp}.xlsx',
                                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )
                        
                        with col_dl4:
                            with open(zip_path, 'rb') as f:
                                st.download_button(
                                    label="📦 Download All (ZIP)",
                                    data=f.read(),
                                    file_name=f'SPLD_Reports_{timestamp}.zip',
                                    mime='application/zip'
                                )
                        
                        # Show project summary
                        st.markdown("### 📈 Project Summary")
                        
                        # Calculate metrics
                        total_projects = len(projects)
                        on_track = sum(1 for p in projects if 'on track' in str(p.get('health', '')).lower())
                        at_risk = sum(1 for p in projects if 'at risk' in str(p.get('health', '')).lower())
                        off_track = total_projects - on_track - at_risk
                        
                        # Display metrics
                        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
                        
                        with metric_col1:
                            st.metric("Total Projects", total_projects)
                        with metric_col2:
                            st.metric("On Track", on_track, f"{on_track/total_projects*100:.0f}%")
                        with metric_col3:
                            st.metric("At Risk", at_risk, f"{at_risk/total_projects*100:.0f}%")
                        with metric_col4:
                            st.metric("Off Track", off_track, f"{off_track/total_projects*100:.0f}%")
                        
                        # Show project list
                        with st.expander("📋 View All Projects"):
                            for i, project in enumerate(projects, 1):
                                health = project.get('health', 'Unknown')
                                if 'on track' in health.lower():
                                    status_icon = "🟢"
                                elif 'at risk' in health.lower():
                                    status_icon = "🟡"
                                else:
                                    status_icon = "🔴"
                                st.write(f"{i}. {status_icon} **{project.get('name', 'Unknown')}** - {project.get('timeline_actual', 0):.0f}% complete")
                        
                except Exception as e:
                    st.error(f"❌ An error occurred: {str(e)}")

with col2:
    # AI Configuration (Optional)
    with st.expander("🤖 AI Text Formatting (Optional)"):
        st.info("Enable AI to format activities, risks, and issues into clean bullet points")
        
        api_key = st.text_input("OpenAI API Key", type="password", 
                               help="Enter your OpenAI API key for text formatting")
        
        if api_key:
            os.environ['OPENAI_API_KEY'] = api_key
            st.success("✅ AI key configured for this session")
        else:
            # Check if already in environment
            if os.getenv('OPENAI_API_KEY'):
                st.success("✅ AI configured from .env file")
            else:
                st.caption("AI formatting is optional - reports work without it")
    
    # Features section
    st.markdown("### ✨ Features")
    
    features = [
        "🎯 SPLD Executive Dashboard",
        "📊 Complete Excel Analysis",
        "📄 Professional PDF Reports",
        "📝 Editable Word Documents",
        "🚦 Traffic Light Indicators",
        "💰 Budget Tracking",
        "📈 Progress Monitoring",
        "📋 Risk & Issue Tracking",
        "🤖 AI Text Formatting"
    ]
    
    for feature in features:
        st.markdown(f"""
        <div style="background-color: rgba(234, 106, 31, 0.1); 
                    border-left: 3px solid #EA6A1F; 
                    padding: 8px; 
                    margin: 5px 0;
                    color: white;">
            {feature}
        </div>
        """, unsafe_allow_html=True)
    
    # Instructions
    st.markdown("### 📖 How to Use")
    st.info("""
    1. **Upload** your PMO Excel file
    2. **Click** Generate Reports
    3. **Download** your reports
    
    The system reads any Excel format and 
    generates professional SPLD reports!
    """)
    
    # About section
    st.markdown("### ℹ️ About")
    st.caption("""
    SPLD PMO Report Generator v1.0
    
    Automatically generates executive 
    dashboard reports from PMO Excel data.
    
    • Reads any Excel format
    • Dynamic report generation
    • Professional SPLD styling
    • Multiple output formats
    """)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    <p>SPLD PMO Report Generator | Strategic Projects & Leadership Development</p>
</div>
""", unsafe_allow_html=True)
