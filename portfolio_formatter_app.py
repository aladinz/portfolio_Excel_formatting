"""
Portfolio Formatter Application
A Streamlit app for formatting Excel portfolio files
"""

import streamlit as st
import subprocess
import os
from pathlib import Path
import tempfile
import shutil

# Page configuration
st.set_page_config(
    page_title="Portfolio Formatter",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom styling
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stTitle {
        color: #1F4788;
        text-align: center;
    }
    .success-box {
        background-color: #E2EFDA;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #4472C4;
    }
    .error-box {
        background-color: #F1DCDB;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #C5504B;
    }
    .info-box {
        background-color: #D9E1F2;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1F4788;
    }
    </style>
""", unsafe_allow_html=True)

# Title and description
st.title("📊 Portfolio Formatter")
st.markdown("### Automatically format your Excel portfolio files")
st.markdown("---")

# Get the current working directory where the script is
script_dir = Path(__file__).parent.absolute()
format_script = script_dir / "format_all.py"

# Check if format_all.py exists
if not format_script.exists():
    st.error("❌ Error: format_all.py not found in the same directory")
    st.stop()

# File upload section
st.markdown("### 📁 Upload Your Excel File")
uploaded_file = st.file_uploader(
    "Choose an Excel file to format",
    type=["xlsx", "xls"],
    help="Type A (with Executive Summary) or Type B (with Data sheet) files are supported"
)

# Information columns
col1, col2 = st.columns(2)
with col1:
    st.markdown("""
    **What happens:**
    - ✓ Professional formatting applied
    - ✓ Color-coded sections
    - ✓ Charts added (if Type A)
    - ✓ Optimized layout
    """)

with col2:
    st.markdown("""
    **Supported formats:**
    - Type A: Executive Summary + Monthly Performance
    - Type B: Single Data sheet
    - All portfolio types
    """)

st.markdown("---")

# Processing section
if uploaded_file is not None:
    st.markdown("### ⚙️ Processing")
    
    # Create a temporary file to save the upload
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        tmp_path = tmp_file.name
    
    try:
        # Create a progress placeholder
        progress_placeholder = st.empty()
        status_placeholder = st.empty()
        
        progress_placeholder.info("⏳ Formatting your file... This may take a moment.")
        
        # Run the format_all.py script
        result = subprocess.run(
            [
                str(Path(script_dir.parent) / ".venv" / "Scripts" / "python.exe"),
                str(format_script),
                tmp_path
            ],
            capture_output=True,
            text=True,
            cwd=str(script_dir)
        )
        
        # Check if successful
        if result.returncode == 0:
            progress_placeholder.empty()
            
            # Read the formatted file
            with open(tmp_path, 'rb') as f:
                formatted_data = f.read()
            
            # Display success message
            st.markdown(
                """
                <div class="success-box">
                <b>✅ Success!</b> Your file has been formatted and is ready to download.
                </div>
                """,
                unsafe_allow_html=True
            )
            
            # Show what was applied
            st.markdown("#### Features Applied:")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("✓ **Professional Colors**\nDark blue headers, color-coded sections")
            with col2:
                st.markdown("✓ **Optimization**\nProper spacing, borders, fonts")
            with col3:
                if "Charts added" in result.stdout:
                    st.markdown("✓ **Charts**\nPortfolio growth & monthly returns")
                else:
                    st.markdown("✓ **Structure**\nProfessional layout applied")
            
            # Download button
            st.download_button(
                label="📥 Download Formatted File",
                data=formatted_data,
                file_name=uploaded_file.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Show script output for details
            if result.stdout.strip():
                with st.expander("📋 Details", expanded=False):
                    st.code(result.stdout, language="text")
        
        else:
            # Error handling
            progress_placeholder.empty()
            
            st.markdown(
                """
                <div class="error-box">
                <b>❌ Error during formatting</b>
                </div>
                """,
                unsafe_allow_html=True
            )
            
            st.error("The formatter encountered an issue. Details below:")
            
            if result.stdout:
                st.markdown("**Output:**")
                st.code(result.stdout, language="text")
            
            if result.stderr:
                st.markdown("**Error message:**")
                st.code(result.stderr, language="text")
    
    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")
    
    finally:
        # Clean up temporary file
        try:
            os.unlink(tmp_path)
        except:
            pass

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 0.9rem;">
<p>Portfolio Formatter v2.1 | Python-powered Excel automation</p>
<p>Run this locally with: <code>streamlit run portfolio_formatter_app.py</code></p>
</div>
""", unsafe_allow_html=True)
