import streamlit as st
from pathlib import Path
from docx import Document
import re
import difflib
from PIL import Image
import pytesseract
import pandas as pd
import requests
from io import BytesIO
import PyPDF2
from concurrent.futures import ThreadPoolExecutor
import concurrent.futures
import zipfile
import time
import fitz 
import tempfile
import os
import time
import logging
import json
import math
import numpy as np
from pdf2image import convert_from_path, convert_from_bytes
import cv2

# ========== CONFIGURATION ==========
# Multiple Gemini API Keys - Will be used in rotation
GEMINI_API_KEYS = [
    "AIzaSyBBFKiwVjOlz06hGtjXe_NBa8D4Iyh_k_k",  # Default key
    "AIzaSyCc8zOm1-tjVhnQGaYCBK-vmhDClUAMyQI",  # Add your second API key here
    "AIzaSyBWiEJQQrcqh4bSPHNyOIkbZEeQWVwXsOI",
    "AIzaSyBLDNq9h4ZqLoOEPIwxBOeNkYLzXRyzXnE",
    "AIzaSyAtPq8ltidY48tJMtGBrq527rTbV56W5Qc",
    "AIzaSyA4yzBQ4omQEMDV2BV_bZ9Da0pOWvqXZ2I",
    "AIzaSyBmEhXf9GpC3FOb_qOfl_avzEvXk2tXy24",
    "AIzaSyDWqfVEbw7ix0DtL39qA0Du851VFxiMh8I",
    "AIzaSyAIIFn4dqzq4GYPSK8FMrJ3CneNFlC-36s",
    "AIzaSyDnXW0oe0r-vOO9PIp0EFXQLpsM4FF7cIU"   
]

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize OCR stats in session state
if 'ocr_stats' not in st.session_state:
    st.session_state.ocr_stats = {
        'scanned_count': 0,
        'ocr_success_count': 0,
        'ocr_success_rate': 0,
        'total_processing_time': 0,
        'avg_processing_time': 0
    }

# Function to get the next API key in rotation
def get_next_api_key():
    """
    Rotates to the next available API key.
    Returns the new API key and updates the session state.
    """
    if 'api_key_index' not in st.session_state:
        st.session_state.api_key_index = 0
    
    # Get valid keys (non-empty strings)
    valid_keys = [key for key in GEMINI_API_KEYS if key.strip()]
    if not valid_keys:
        return ""
    
    # Move to the next key
    next_index = (st.session_state.api_key_index + 1) % len(valid_keys)
    st.session_state.api_key_index = next_index
    
    return valid_keys[next_index]

# Function to get the current API key
def get_current_api_key():
    """Get the currently active API key"""
    if 'api_key_index' not in st.session_state:
        st.session_state.api_key_index = 0
    
    # Filter out empty keys
    valid_keys = [key for key in GEMINI_API_KEYS if key.strip()]
    if not valid_keys:
        return ""
    
    current_index = st.session_state.api_key_index % len(valid_keys)
    return valid_keys[current_index]

# Custom CSS to enhance UI
def local_css():
    st.markdown("""
    <style>
        .main {
            padding: 2rem 4rem;
            max-width: 1200px;
            margin: 0 auto;
        }
        .stApp {
            background-color: #f8f9fa;
        }
        h1, h2, h3 {
            color: #1E88E5;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .css-18e3th9 {
            padding-top: 0;
        }
        .css-1d391kg {
            padding-top: 3.5rem;
        }
        .css-12oz5g7 {
            max-width: 1200px;
        }
        .stButton>button {
            background-color: #1E88E5;
            color: white;
            border-radius: 6px;
            padding: 0.5rem 1rem;
            font-weight: 500;
        }
        .stButton>button:hover {
            background-color: #1565C0;
        }
        .stDownloadButton>button {
            background-color: #43A047;
            color: white;
            border-radius: 6px;
            padding: 0.5rem 1rem;
            font-weight: 500;
        }
        .stDownloadButton>button:hover {
            background-color: #2E7D32;
        }
        .stDataFrame {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        div[data-testid="stFileUploadDropzone"] {
            border: 2px dashed #1E88E5;
            border-radius: 8px;
            padding: 20px;
            background-color: #E3F2FD;
        }
        .stProgress .st-bo {
            background-color: #1E88E5;
        }
        .info-box {
            background-color: #E3F2FD;
            border-radius: 8px;
            padding: 1.5rem;
            margin-bottom: 1.5rem;
            border-left: 4px solid #1E88E5;
        }
        .warning-box {
            background-color: #FFF3E0;
            border-radius: 8px;
            padding: 1rem;
            margin-bottom: 1rem;
            border-left: 4px solid #FF9800;
        }
        .success-box {
            background-color: #E8F5E9;
            border-radius: 8px;
            padding: 1rem;
            margin-bottom: 1rem;
            border-left: 4px solid #43A047;
        }
        .css-1y4p8pa {
            padding: 1.5rem 1rem 10rem;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 4px 4px 0 0;
            padding: 10px 20px;
            background-color: #f1f3f4;
        }
        .stTabs [aria-selected="true"] {
            background-color: #1E88E5 !important;
            color: white !important;
        }
        .custom-metric-container {
            display: flex;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 20px;
        }
        .metric-card {
            background-color: white;
            border-radius: 8px;
            padding: 15px;
            flex: 1;
            min-width: 200px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            text-align: center;
        }
        .metric-card h3 {
            margin: 0;
            font-size: 16px;
            color: #555;
        }
        .metric-card p {
            margin: 10px 0 0;
            font-size: 24px;
            font-weight: bold;
            color: #1E88E5;
        }
        .api-key-badge {
            display: inline-block;
            padding: 3px 8px;
            background-color: #E3F2FD;
            border-radius: 12px;
            font-size: 0.8em;
            margin-left: 10px;
            border: 1px solid #1E88E5;
        }
        .api-key-status {
            text-align: center;
            margin: 10px 0;
            padding: 5px;
            border-radius: 4px;
            background-color: #E3F2FD;
            font-size: 0.9em;
        }
        .ocr-badge {
            display: inline-block;
            padding: 3px 8px;
            background-color: #E8F5E9;
            border-radius: 12px;
            font-size: 0.8em;
            margin-left: 10px;
            border: 1px solid #43A047;
        }
        .ocr-disabled-badge {
            display: inline-block;
            padding: 3px 8px;
            background-color: #FFEBEE;
            border-radius: 12px;
            font-size: 0.8em;
            margin-left: 10px;
            border: 1px solid #E57373;
        }
    </style>
    """, unsafe_allow_html=True)

# ========== OCR FUNCTIONS ==========

# Function to check if Tesseract OCR is properly installed
def check_tesseract_installation():
    """
    Checks if Tesseract OCR is properly installed and configured.
    Returns True if working, False otherwise.
    """
    try:
        # Try to get Tesseract version
        version = pytesseract.get_tesseract_version()
        logging.info(f"Tesseract OCR is installed. Version: {version}")
        return True
    except Exception as e:
        logging.error(f"Tesseract OCR is not properly installed: {e}")
        return False

# Function to configure Tesseract path
def configure_tesseract_path(path):
    """
    Configures the path to Tesseract executable.
    """
    pytesseract.pytesseract.tesseract_cmd = path

# Image preprocessing for better OCR results
def preprocess_image_for_ocr(image):
    """
    Preprocess an image for better OCR results.
    
    Args:
        image: PIL Image object
    
    Returns:
        Preprocessed PIL Image object
    """
    try:
        # Get preprocessing level from session state (if set)
        preprocess_level = st.session_state.get('preprocess_level', 'normal')
        
        # Convert PIL image to OpenCV format
        img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        # Grayscale conversion
        img_gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        
        # Apply different preprocessing based on level
        if preprocess_level == 'minimal':
            # Just grayscale and basic noise removal
            img_processed = cv2.fastNlMeansDenoising(img_gray, None, 10, 7, 21)
            
        elif preprocess_level == 'aggressive':
            # Aggressive preprocessing for hard-to-read documents
            # Increase contrast
            img_contrast = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8)).apply(img_gray)
            
            # Denoise
            img_denoised = cv2.fastNlMeansDenoising(img_contrast, None, 30, 7, 21)
            
            # Threshold to make text more distinct
            _, img_thresh = cv2.threshold(img_denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Dilate text slightly to make it more readable
            kernel = np.ones((1, 1), np.uint8)
            img_processed = cv2.dilate(img_thresh, kernel, iterations=1)
            
        else:  # normal preprocessing
            # Moderate preprocessing for most documents
            # Apply adaptive thresholding
            img_thresh = cv2.adaptiveThreshold(
                img_gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
            )
            
            # Denoise
            img_processed = cv2.fastNlMeansDenoising(img_thresh, None, 10, 7, 21)
        
        # Convert back to PIL image
        return Image.fromarray(img_processed)
        
    except Exception as e:
        logging.error(f"Image preprocessing error: {e}")
        # Return original image if preprocessing fails
        return image

# Process scanned resume PDF using OCR
def process_scanned_resume(pdf_file):
    """
    Process a scanned resume PDF by converting pages to images and applying OCR.
    Returns the extracted text.
    """
    start_time = time.time()
    extracted_text = ""
    success = False
    
    # Update OCR stats
    st.session_state.ocr_stats['scanned_count'] += 1
    
    try:
        # Get OCR settings from session state
        ocr_quality = st.session_state.get('ocr_quality', "Balanced")
        ocr_lang = st.session_state.get('ocr_lang', 'eng')
        preprocess_level = st.session_state.get('preprocess_level', 'normal')
        
        # Set DPI based on quality setting
        dpi = 200  # Default/Balanced
        if ocr_quality == "Fast":
            dpi = 150
        elif ocr_quality == "High Quality":
            dpi = 300
        
        # Create a temporary file for the PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            if isinstance(pdf_file, BytesIO):
                temp_pdf.write(pdf_file.read())
                pdf_file.seek(0)  # Reset file pointer
            else:
                temp_pdf.write(pdf_file)
            temp_pdf_path = temp_pdf.name
        
        # Convert PDF to images using pdf2image
        try:
            # Try with pdf2image first (handles scanned PDFs better)
            images = convert_from_path(temp_pdf_path, dpi=dpi)
        except Exception as e:
            logging.warning(f"pdf2image conversion failed: {e}. Falling back to PyMuPDF.")
            # Fall back to PyMuPDF
            images = []
            with fitz.open(temp_pdf_path) as doc:
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    # Set matrix based on quality
                    zoom = 1.5  # Default/Balanced
                    if ocr_quality == "Fast":
                        zoom = 1.0
                    elif ocr_quality == "High Quality":
                        zoom = 2.0
                    
                    matrix = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=matrix)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    images.append(img)
        
        # Process each image with OCR
        for img in images:
            # Preprocess the image based on user settings
            if st.session_state.get('preprocess_images', True):
                img = preprocess_image_for_ocr(img)
            
            # Apply OCR with configured settings
            custom_config = f'--psm 1 --oem 3'  # Default OCR config
            page_text = pytesseract.image_to_string(img, lang=ocr_lang, config=custom_config)
            extracted_text += page_text + "\n\n"  # Add spacing between pages
        
        # Clean up temporary file
        try:
            os.remove(temp_pdf_path)
        except:
            pass
        
        success = True
        end_time = time.time()
        processing_time = end_time - start_time
        
        # Update OCR stats
        if success:
            st.session_state.ocr_stats['ocr_success_count'] += 1
            st.session_state.ocr_stats['total_processing_time'] += processing_time
            total_count = st.session_state.ocr_stats['scanned_count']
            success_count = st.session_state.ocr_stats['ocr_success_count']
            
            # Update success rate and average processing time
            st.session_state.ocr_stats['ocr_success_rate'] = (success_count / total_count) * 100
            st.session_state.ocr_stats['avg_processing_time'] = st.session_state.ocr_stats['total_processing_time'] / success_count
        
        return extracted_text
    
    except Exception as e:
        logging.error(f"Error in OCR processing: {e}")
        end_time = time.time()
        processing_time = end_time - start_time
        return f"Error processing scanned resume: {str(e)}"

# Enhanced function to extract text from PDF with OCR capabilities
def extract_text_from_pdf(pdf_file):
    """
    Enhanced function to extract text from PDF files, including image-based PDFs.
    Uses OCR when necessary based on session state settings.
    """
    # Check if OCR is enabled in session state
    ocr_enabled = st.session_state.get('enable_ocr', True)
    
    # Try regular text extraction first
    text = ""
    try:
        # Try using PyMuPDF (fitz) first for regular text extraction
        with fitz.open(stream=pdf_file, filetype="pdf") as doc:
            # Check if the document has actual text content
            has_text = False
            total_text = ""
            
            # Check the first few pages for text content
            pages_to_check = min(3, len(doc))
            for page_num in range(pages_to_check):
                page = doc[page_num]
                page_text = page.get_text()
                total_text += page_text
                # If we find a reasonable amount of text, consider it a text PDF
                if len(page_text.strip()) > 200:  # Threshold for text detection
                    has_text = True
            
            # If the document appears to have readable text
            if has_text:
                for page in doc:
                    text += page.get_text()
                return text
            
            # If not much text was detected and OCR is enabled, proceed with OCR
            if ocr_enabled:
                logging.info("PDF appears to be image-based. Using OCR processing.")
                
                # Reset file pointer if it's a BytesIO object
                if isinstance(pdf_file, BytesIO):
                    pdf_file.seek(0)
                
                # Use the dedicated OCR function
                return process_scanned_resume(pdf_file)
            else:
                # OCR is disabled, return what little text we found
                logging.warning("PDF appears to be image-based but OCR is disabled.")
                return total_text
                
    except Exception as e:
        logging.error(f"Error using PyMuPDF: {e}")
        
        # Fall back to PyPDF2 if fitz fails
        try:
            if isinstance(pdf_file, BytesIO):
                pdf_file.seek(0)  # Reset file pointer
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            for page in pdf_reader.pages:
                page_text = page.extract_text() or ""
                text += page_text
            
            # If PyPDF2 didn't extract much text and OCR is enabled
            if len(text.strip()) < 200 and ocr_enabled:  # Minimal text threshold
                logging.info("PyPDF2 extraction yielded limited text. Attempting OCR.")
                
                # Reset file pointer
                if isinstance(pdf_file, BytesIO):
                    pdf_file.seek(0)
                
                # Use dedicated OCR function
                return process_scanned_resume(pdf_file)
            
            return text
            
        except Exception as e2:
            logging.error(f"Error extracting text with PyPDF2: {e2}")
            
            # If all other methods fail and OCR is enabled, try pure OCR approach
            if ocr_enabled:
                try:
                    if isinstance(pdf_file, BytesIO):
                        pdf_file.seek(0)
                    return process_scanned_resume(pdf_file)
                except Exception as e3:
                    logging.error(f"Error with OCR fallback: {e3}")
            
            # If everything fails
            return ""

# Setup OCR environment in Streamlit
def setup_ocr_environment():
    """Set up the OCR environment in the Streamlit app"""
    
    # Create OCR settings section
    st.markdown("#### OCR Configuration")
    
    # Check if Tesseract is installed
    tesseract_installed = check_tesseract_installation()
    
    if tesseract_installed:
        st.success("✅ Tesseract OCR is properly installed and configured.")
    else:
        st.error("❌ Tesseract OCR is not installed or not properly configured.")
        
        # Show installation instructions based on OS
        with st.expander("Installation Instructions"):
            st.markdown("""
            ### Tesseract OCR Installation
            
            #### Windows:
            1. Download the installer from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
            2. Run the installer and follow the instructions
            3. Add the Tesseract installation directory to your system PATH or specify the path below
            
            #### macOS:
            ```bash
            brew install tesseract
            ```
            
            #### Ubuntu/Debian:
            ```bash
            sudo apt update
            sudo apt install tesseract-ocr
            sudo apt install libtesseract-dev
            ```
            
            After installation, restart the Streamlit app.
            """)
    
    # Allow manual path configuration
    custom_path = st.text_input(
        "Tesseract executable path (only needed if not in system PATH)",
        value=st.session_state.get('tesseract_path', ''),
        placeholder="e.g., C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        help="Leave empty if Tesseract is in your system PATH"
    )
    
    if custom_path:
        st.session_state['tesseract_path'] = custom_path
        try:
            configure_tesseract_path(custom_path)
            st.success("✅ Custom Tesseract path configured successfully.")
        except Exception as e:
            st.error(f"❌ Failed to configure custom Tesseract path: {str(e)}")

# Modified add_ocr_options_to_ui function to avoid nested expanders
def add_ocr_options_to_ui():
    """Add OCR options to the user interface"""
    st.markdown("<h3>📷 OCR Options for Scanned Resumes</h3>", unsafe_allow_html=True)
    
    # OCR toggle - Default to enabled
    ocr_enabled = st.checkbox(
        "Enable OCR for scanned/image-based resumes", 
        value=st.session_state.get('enable_ocr', True),
        help="Use Optical Character Recognition to extract text from scanned resumes"
    )
    st.session_state['enable_ocr'] = ocr_enabled
    
    if ocr_enabled:
        # OCR quality options
        col1, col2 = st.columns(2)
        
        with col1:
            ocr_quality = st.select_slider(
                "OCR Quality",
                options=["Fast", "Balanced", "High Quality"],
                value=st.session_state.get('ocr_quality', "Balanced"),
                help="Higher quality gives better results but takes longer"
            )
            st.session_state['ocr_quality'] = ocr_quality
        
        with col2:
            if ocr_quality == "Fast":
                st.info("⚡ Fast mode: Quick processing, lower accuracy")
            elif ocr_quality == "Balanced":
                st.info("⚖️ Balanced mode: Recommended for most resumes")
            else:
                st.info("✨ High Quality mode: Detailed analysis, slower processing")
        
        # Use a collapsible section ONLY if this function is not being called inside an expander
        st.markdown("#### Advanced OCR Settings")
        
        # Language selection
        ocr_lang = st.selectbox(
            "OCR Language",
            options=["English", "Multi-language"],
            index=0,
            help="Select language for better OCR results"
        )
        st.session_state['ocr_lang'] = 'eng' if ocr_lang == "English" else 'eng+fra+deu+spa'
        
        # Image preprocessing toggle
        preprocess_images = st.checkbox(
            "Preprocess images for better OCR", 
            value=st.session_state.get('preprocess_images', True),
            help="Apply image enhancement techniques before OCR"
        )
        st.session_state['preprocess_images'] = preprocess_images
        
        # Confidence threshold
        confidence_threshold = st.slider(
            "OCR Confidence Threshold (%)",
            min_value=0,
            max_value=100,
            value=st.session_state.get('confidence_threshold', 60),
            help="Minimum confidence level for OCR results"
        )
        st.session_state['confidence_threshold'] = confidence_threshold
    else:
        st.warning("⚠️ OCR is disabled. Scanned or image-based resumes may not be processed correctly.")

# Modified enhanced_api_key_management_ui function to avoid nested expanders
def enhanced_api_key_management_ui():
    st.markdown("#### API Keys & OCR Configuration")
    
    # Create tabs for API Keys and OCR
    api_tab, ocr_tab = st.tabs(["API Keys", "OCR Settings"])
    
    with api_tab:
        st.info("Configure multiple API keys to handle rate limits")
        
        # Display and allow editing of all API keys
        updated_keys = []
        
        global GEMINI_API_KEYS
        # Initialize with any existing keys
        if 'gemini_api_keys' in st.session_state:
            GEMINI_API_KEYS = st.session_state.gemini_api_keys
        
        # Display text inputs for each key - now supporting 5 keys
        for i in range(10):  # Support up to 5 API keys
            key_label = f"Gemini API Key {i+1}"
            key_value = GEMINI_API_KEYS[i] if i < len(GEMINI_API_KEYS) else ""
            
            new_key = st.text_input(
                key_label, 
                value=key_value, 
                type="password",
                key=f"api_key_{i}"
            )
            updated_keys.append(new_key)
        
        # Add a button to apply changes
        if st.button("Update API Keys"):
            GEMINI_API_KEYS = updated_keys
            st.session_state.gemini_api_keys = updated_keys
            st.session_state.api_key_index = 0  # Reset to first key
            st.session_state.key_resume_counts = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  # Reset counts for all 5 keys
            st.success("API keys updated successfully!")
    
    with ocr_tab:
        # Modified: Don't use expanders here since this is already in an expander
        st.markdown("#### OCR Configuration")
        
        # Check if Tesseract is installed
        tesseract_installed = check_tesseract_installation()
        
        if tesseract_installed:
            st.success("✅ Tesseract OCR is properly installed and configured.")
        else:
            st.error("❌ Tesseract OCR is not installed or not properly configured.")
            
            # Instead of using an expander, just show the installation instructions
            st.markdown("### Tesseract OCR Installation")
            st.markdown("""
            #### Windows:
            1. Download the installer from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
            2. Run the installer and follow the instructions
            3. Add the Tesseract installation directory to your system PATH or specify the path below
            
            #### macOS:
            ```bash
            brew install tesseract
            ```
            
            #### Ubuntu/Debian:
            ```bash
            sudo apt update
            sudo apt install tesseract-ocr
            sudo apt install libtesseract-dev
            ```
            
            After installation, restart the Streamlit app.
            """)
        
        # Allow manual path configuration
        custom_path = st.text_input(
            "Tesseract executable path (only needed if not in system PATH)",
            value=st.session_state.get('tesseract_path', ''),
            placeholder="e.g., C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
            help="Leave empty if Tesseract is in your system PATH"
        )
        
        if custom_path:
            st.session_state['tesseract_path'] = custom_path
            try:
                configure_tesseract_path(custom_path)
                st.success("✅ Custom Tesseract path configured successfully.")
            except Exception as e:
                st.error(f"❌ Failed to configure custom Tesseract path: {str(e)}")

# ========== Existing API Functions ==========

def correct_google_drive_url(url):
    if 'drive.google.com' in url:
        file_id_match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
        if file_id_match:
            file_id = file_id_match.group(1)
            return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url

def extract_clickable_links(pdf_path):
    """Extract clickable links embedded in the PDF."""
    clickable_links = []
    try:
        with fitz.open(pdf_path) as doc:
            for page in doc:
                # Extract all links from the current page
                links = page.get_links()
                for link in links:
                    if 'uri' in link:  # Check if the link contains a URI
                        clickable_links.append(link['uri'])
    except Exception as e:
        print(f"Error extracting clickable links: {e}")
    
    return clickable_links

def extract_github_username(github_url):
    """
    Extract the GitHub username from a given GitHub URL.
    :param github_url: The GitHub URL.
    :return: The GitHub username or None if the URL is invalid.
    """
    # Handle multiple GitHub URL formats
    patterns = [
        r'https?://github\.com/([^/]+)/?',  # Standard GitHub URL
        r'https?://github\.com/([^/]+)/[^/]+',  # GitHub repo URL
        r'https?://([^\.]+)\.github\.io',  # GitHub Pages URL
    ]
    
    for pattern in patterns:
        match = re.match(pattern, github_url)
        if match:
            username = match.group(1)
            # Filter out common non-username paths
            if username not in ['settings', 'login', 'join', 'features', 'marketplace', 'explore']:
                return username
    
    return None  # Not a valid GitHub profile URL

def get_github_repo_count(username):
    """
    Get the public repository count for a GitHub user with enhanced rate limit handling.
    :param username: The GitHub username.
    :return: The repository count (int) or None if the user is not found or an error occurs.
    """
    if not username:
        return None
        
    try:
        # Add a user agent and token-based authentication to avoid GitHub API rate limiting
        headers = {
            'User-Agent': 'Resume-Analyzer-App',
            'Accept': 'application/vnd.github.v3+json'
        }
        
        # Create token pool and rotate through them similar to Gemini API keys
        if 'github_token_index' not in st.session_state:
            st.session_state.github_token_index = 0
            
        github_tokens = [
            "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"  # Add your GitHub tokens here
        ]
        
        # Use authentication if tokens are available
        if any(github_tokens):
            current_token_index = st.session_state.github_token_index
            token = github_tokens[current_token_index % len(github_tokens)]
            
            if token:
                # Use token authentication
                headers['Authorization'] = f'token {token}'
                
                # Rotate to next token for subsequent calls
                st.session_state.github_token_index = (current_token_index + 1) % len(github_tokens)
        
        # Add caching to avoid repeated calls for the same username
        if 'github_cache' not in st.session_state:
            st.session_state.github_cache = {}
            
        # Check if username is in cache
        if username in st.session_state.github_cache:
            return st.session_state.github_cache[username]
        
        # Make request with increased timeout and exponential backoff
        max_retries = 2
        for retry in range(max_retries):
            try:
                response = requests.get(
                    f'https://api.github.com/users/{username}', 
                    headers=headers, 
                    timeout=15
                )
                
                if response.status_code == 200:
                    data = response.json()
                    repo_count = data.get('public_repos', 0)
                    
                    # Cache the result
                    st.session_state.github_cache[username] = repo_count
                    return repo_count
                    
                elif response.status_code == 404:
                    logging.warning(f"GitHub username not found: {username}")
                    st.session_state.github_cache[username] = 0  # Cache not found results
                    return 0
                    
                elif response.status_code == 403:
                    if 'X-RateLimit-Remaining' in response.headers and int(response.headers['X-RateLimit-Remaining']) == 0:
                        logging.warning(f"GitHub API rate limit exceeded for username: {username}")
                        
                        # If we have multiple tokens, try the next one
                        if len(github_tokens) > 1 and retry < max_retries - 1:
                            st.session_state.github_token_index = (st.session_state.github_token_index + 1) % len(github_tokens)
                            time.sleep(1)  # Brief pause before retry
                            continue
                    
                    # Fall back to returning 0 if rate limited
                    logging.warning(f"Returning default value 0 due to GitHub API rate limit")
                    st.session_state.github_cache[username] = 0
                    return 0
                    
                else:
                    logging.warning(f"GitHub API returned status code {response.status_code} for username: {username}")
                    if retry < max_retries - 1:
                        time.sleep(2)  # Wait before retry
                        continue
                    return 0
                    
            except requests.exceptions.Timeout:
                if retry < max_retries - 1:
                    time.sleep(2)  # Wait before retry
                    continue
                logging.error(f"GitHub API request timed out for {username}")
                return 0
                
            except Exception as e:
                logging.error(f"Error fetching GitHub repo count for {username}: {e}")
                return 0
                
    except Exception as e:
        logging.error(f"Error in GitHub API handling for {username}: {e}")
        return 0

# Modified analyze_text_with_gemini function to handle API key rotation
def analyze_text_with_gemini(prompt):
    """
    Use Gemini API to analyze text with automatic key rotation on rate limits.
    """
    # Try all available API keys if needed
    valid_keys = [key for key in GEMINI_API_KEYS if key.strip()]
    if not valid_keys:
        return "Error: No valid API keys configured. Please add at least one API key in Settings."
        
    # Try each key up to the number of valid keys we have
    for attempt in range(len(valid_keys)):
        # Get the current API key
        current_api_key = get_current_api_key()
        if not current_api_key:
            continue
            
        gemini_api_url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={current_api_key}"
        
        # API request payload
        payload = {
            "contents": [{"parts": [{"text": prompt}]}]
        }

        headers = {"Content-Type": "application/json"}

        # Make request to Gemini API with retries
        max_retries = 2  # Reduced number of retries per key
        for retry in range(max_retries):
            try:
                response = requests.post(gemini_api_url, headers=headers, json=payload, timeout=60)
                
                if response.status_code == 200:
                    json_response = response.json()
                    if "candidates" in json_response and json_response["candidates"]:
                        # Extract the raw text response
                        raw_response = json_response["candidates"][0]["content"]["parts"][0]["text"]
                        return raw_response.strip()
                    else:
                        if retry < max_retries - 1:
                            time.sleep(2)  # Short delay before retry
                            continue
                        # Move to next key if no candidates
                        get_next_api_key()
                        break
                
                elif response.status_code == 429:  # Rate limit exceeded
                    # Log the rate limit error
                    logging.warning(f"Rate limit reached for API key {st.session_state.api_key_index + 1}, switching to next key")
                    
                    # Rotate to the next API key
                    next_key = get_next_api_key()
                    
                    # Show a message about key rotation
                    st.toast(f"Rate limit reached. Switching to API key {st.session_state.api_key_index + 1}")
                    
                    # Brief pause before trying with the new key
                    time.sleep(1)
                    
                    # Break the retry loop to try with the new key
                    break
                    
                else:
                    if retry < max_retries - 1:
                        time.sleep(2)  # Short delay before retry
                        continue
                    # Move to next key if persistent error
                    get_next_api_key()
                    break
                    
            except Exception as e:
                if retry < max_retries - 1:
                    time.sleep(2)  # Short delay before retry
                    continue
                # Move to next key if persistent error
                get_next_api_key()
                break
    
    # If we've tried all keys and still failed
    return "Error: Unable to process with any available API keys. Please try again later."

# ========== RESUME PROCESSING FUNCTIONS ==========

def process_resume_with_details(row):
    user_id = row['user_id']
    resume_link = row['Resume link']
    
    # Correct the Google Drive URLs
    resume_link = correct_google_drive_url(resume_link)

    if resume_link.startswith('file:///'):
        local_path = resume_link.replace('file:///', '')
        try:
            with open(local_path, 'rb') as f:
                resume_pdf = f.read()
        except FileNotFoundError:
            logging.error(f"File not found: {local_path}")
            resume_pdf = None
    elif resume_link.startswith('http'):
        response = requests.get(resume_link)
        if response.status_code == 200:
            resume_pdf = BytesIO(response.content)
        else:
            logging.error(f"Failed to retrieve URL: {resume_link}")
            resume_pdf = None
    else:
        resume_pdf = None

    if resume_pdf:
        try:
            # Save the PDF temporarily for clickable link extraction
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(resume_pdf if isinstance(resume_pdf, bytes) else resume_pdf.read())
                temp_pdf_path = temp_pdf.name

            # Extract text from the PDF - Using enhanced OCR-capable function
            resume_text = extract_text_from_pdf(BytesIO(resume_pdf) if isinstance(resume_pdf, bytes) else resume_pdf)
            
            # Single comprehensive prompt to extract all required information
            comprehensive_prompt = f"""
Extract the following information from the resume text. Format your response using the exact structure shown below:

### PROJECT TITLES AND TECHNOLOGIES
Extract project titles and technologies used. Format ALL technologies consistently using the following rules:
1. No periods in technology names (use "ReactJS" not "React.js")
2. For compound names, use camelCase with capital first letters (e.g., "ReactJS", "NodeJS", "PostgreSQL", "MongoDB", "ExpressJS", "TensorFlow")
3. For single-word technologies, capitalize first letter only (e.g., "Python", "Java", "JavaScript")
4. Always use the same format for the same technology (e.g., always use "ReactJS", never "React" or "React.js")

Common technology formats to follow:
- ReactJS (not React.js, React, or react)
- NodeJS (not Node.js or node)
- ExpressJS (not Express.js or express)
- MongoDB (not Mongo DB or mongo)
- PostgreSQL (not Postgres or postgresql)
- JavaScript (not Javascript or javascript)
- TypeScript (not Typescript or typescript)

Each project on a new line followed by its technologies.

Example format:
ProjectName
Technologies: ReactJS, ExpressJS, PostgreSQL, NextJS, MongoDB

### WORK EXPERIENCES
Format: Job Title at Company Name (Duration)
If no experience is mentioned, write "No experience listed"

### LINKS
Extract all valid https links (e.g., GitHub, LinkedIn, personal website).
Each link on a new line. If no valid links, write "No links found"

Resume Text:
{resume_text}
"""
            extracted_info = analyze_text_with_gemini(comprehensive_prompt)
            
            # Parse the extracted information
            sections = extracted_info.split("###")
            
            # Initialize variables
            project_titles = ""
            experiences = None
            text_links = None
            
            # Extract each section
            for section in sections:
                if "PROJECT TITLES AND TECHNOLOGIES" in section:
                    project_titles = section.replace("PROJECT TITLES AND TECHNOLOGIES", "").strip()
                elif "WORK EXPERIENCES" in section:
                    experiences_text = section.replace("WORK EXPERIENCES", "").strip()
                    experiences = experiences_text if experiences_text and "No experience listed" not in experiences_text else None
                elif "LINKS" in section:
                    links_text = section.replace("LINKS", "").strip()
                    text_links = links_text if links_text and "No links found" not in links_text else None
            
            # Extract embedded clickable links
            embedded_links = extract_clickable_links(temp_pdf_path)
            if embedded_links:
                clickable_links = "\n".join(embedded_links)
            else:
                clickable_links = None

            # Combine text-extracted links and embedded clickable links
            all_links = None
            if text_links or clickable_links:
                all_links = "\n".join(filter(None, [text_links, clickable_links])).strip() or None


            # Extract GitHub repository counts for all GitHub links
            github_links = []
            repo_count = 0
    
            if all_links:
                github_links = [link for link in all_links.split("\n") if "github.com" in link]
                logging.info(f"Found {len(github_links)} GitHub links for user {user_id}")
        
                # Extract unique usernames from GitHub links
                usernames = []
                for link in github_links:
                    username = extract_github_username(link)
                    if username:
                        usernames.append(username)
                        logging.info(f"Extracted GitHub username: {username} from {link}")
                    else:
                        logging.warning(f"Could not extract username from GitHub link: {link}")
            
                unique_usernames = set(usernames)
                logging.info(f"Found {len(unique_usernames)} unique GitHub usernames")
            
                # Query the GitHub API for each unique username and sum the repo counts
                for username in unique_usernames:
                    count = get_github_repo_count(username)
                    if count is not None:
                        repo_count += count
                        logging.info(f"Found {count} repositories for GitHub user {username}")
                    else:
                        logging.warning(f"Could not get repository count for GitHub user {username}")
        
            return {
                'User ID': user_id,
                'Resume Link': resume_link,
                'Project Titles': project_titles,
                'Experiences': experiences,
                'Links': all_links,
                'Repo Count': repo_count if repo_count > 0 else None
            }

        except Exception as e:
            logging.error(f"An error occurred during resume processing: {e}")
            return {
                'User ID': user_id,
                'Resume Link': resume_link,
                'Project Titles': 'Failed to process the resume.',
                'Experiences': '',
                'Links': '',
                'Repo Count': None
            }
        finally:
            # Ensure the temporary file is removed
            try:
                if 'temp_pdf_path' in locals():
                    os.remove(temp_pdf_path)
            except Exception as e:
                logging.error(f"Error deleting temporary file: {e}")
    else:
        return {
            'User ID': user_id,
            'Resume Link': resume_link,
            'Project Titles': 'Failed to retrieve or process the resume.',
            'Experiences': '',
            'Links': '',
            'Repo Count': None
        }

def process_resume_skills(row):
    user_id = row['user_id']
    resume_link = row['Resume link']
    
    # Correct the Google Drive URLs
    resume_link = correct_google_drive_url(resume_link)

    if resume_link.startswith('file:///'):
        local_path = resume_link.replace('file:///', '')
        try:
            with open(local_path, 'rb') as f:
                resume_pdf = f.read()
        except FileNotFoundError:
            logging.error(f"File not found: {local_path}")
            resume_pdf = None
    elif resume_link.startswith('http'):
        response = requests.get(resume_link)
        if response.status_code == 200:
            resume_pdf = BytesIO(response.content)
        else:
            logging.error(f"Failed to retrieve URL: {resume_link}")
            resume_pdf = None
    else:
        resume_pdf = None

    if resume_pdf:
        try:
            # Use the enhanced OCR-capable text extraction function
            resume_text = extract_text_from_pdf(BytesIO(resume_pdf) if isinstance(resume_pdf, bytes) else resume_pdf)

            # Updated prompt for consistent skills formatting
            skills_prompt = f"""
Extract only the skills listed in the 'Skills' section from the following resume text. 
Format all technologies and skills in a consistent style, following these rules:
1. Use the format 'React.js', 'Node.js', 'Express.js', etc., where the first letter is capitalized, and the rest are lowercase, with a period for technologies like 'React.js' or 'Node.js'.
2. Separate skills with commas.
3. Include only unique skills; do not repeat any.

For example:
React.js, Node.js, Express.js, PostgreSQL, HTML, CSS, JavaScript

Resume Text:
{resume_text}
"""
            extracted_skills = analyze_text_with_gemini(skills_prompt)

            return {
                'User ID': user_id,
                'Resume Link': resume_link,
                'Skills': extracted_skills
            }
        except Exception as e:
            logging.error(f"Failed to process resume: {e}")
            return {
                'User ID': user_id,
                'Resume Link': resume_link,
                'Skills': 'Failed to process the resume.'
            }
    else:
        return {
            'User ID': user_id,
            'Resume Link': resume_link,
            'Skills': 'Failed to retrieve or process the resume.'
        }

def process_resume_details(row):
    """
    Processes the details for each resume. This includes fetching the resume from a URL or local file,
    extracting the text, and analyzing it to retrieve full name, mobile number, and email.
    """
    user_id = row['user_id']
    resume_link = row['Resume link']
    
    # Correct the Google Drive URLs if necessary
    resume_link = correct_google_drive_url(resume_link)

    # Initialize resume_pdf as None
    resume_pdf = None

    try:
        # Check if the link is a local file or URL
        if resume_link.startswith('file:///'):
            local_path = resume_link.replace('file:///', '')
            with open(local_path, 'rb') as f:
                resume_pdf = f.read()
        elif resume_link.startswith('http'):
            response = requests.get(resume_link)
            if response.status_code == 200:
                resume_pdf = BytesIO(response.content)
            else:
                # Log the error internally, not displaying on frontend
                logging.error(f"Failed to retrieve URL: {resume_link}, status code: {response.status_code}")
        else:
            # Log the unsupported link type error
            logging.error(f"Unsupported resume link format: {resume_link}")
        
        # If we have the PDF, continue processing
        if resume_pdf:
            # Use the enhanced OCR-capable text extraction function
            resume_text = extract_text_from_pdf(resume_pdf)

            # Construct the refined prompt to avoid triggering content filtering
            details_prompt = f"""
Extract the full name, mobile number, and email ID from the following resume text. 
If any of the details are missing, leave the field empty. 
Please return the results in this format:

Full Name: <Full Name>
Mobile Number: <Mobile Number>
Email ID: <Email ID>

Resume Text:
{resume_text}
"""
            # Try to analyze the resume text using Gemini API
            extracted_details = analyze_text_with_gemini(details_prompt)
            
            # Parse the extracted details
            return parse_extracted_details(extracted_details, user_id, resume_link)
        else:
            return create_error_result(user_id, resume_link)

    except Exception as e:
        # Log any exception that occurs, but do not display it on the frontend
        logging.error(f"Error processing resume for {user_id}: {str(e)}")
        return create_error_result(user_id, resume_link)

def create_error_result(user_id, resume_link):
    """Creates a standardized error result when an issue occurs."""
    return {
        'User ID': user_id,
        'Resume Link': resume_link,
        'Full Name': 'Failed to retrieve or process the resume',
        'Mobile Number': 'Failed to retrieve or process the resume',
        'Email ID': 'Failed to retrieve or process the resume',
    }

def parse_extracted_details(extracted_details, user_id, resume_link):
    """Parses the extracted details into the proper fields."""
    details = extracted_details.split('\n')
    full_name = None
    mobile_number = None
    email_id = None
    
    # Parse the details into corresponding fields
    for line in details:
        if line.lower().startswith('full name'):
            full_name = line.split(":", 1)[-1].strip() if ":" in line else None
        elif line.lower().startswith('mobile number'):
            mobile_number = line.split(":", 1)[-1].strip() if ":" in line else None
        elif line.lower().startswith('email id'):
            email_id = line.split(":", 1)[-1].strip() if ":" in line else None
    
    # Return default values if some details are missing
    return {
        'User ID': user_id,
        'Resume Link': resume_link,
        'Full Name': full_name if full_name else 'N/A',
        'Mobile Number': mobile_number if mobile_number else 'N/A',
        'Email ID': email_id if email_id else 'N/A',
    }

# ========== BATCH PROCESSING FUNCTIONS ==========

def process_resumes_in_batches(df, process_function, batch_size=5):
    """
    Process resumes in batches, rotating API keys after every 10 resumes.
    """
    total_resumes = len(df)
    results = []
    
    # Create progress components
    progress_container = st.container()
    with progress_container:
        progress_text = st.empty()
        progress_bar = st.progress(0)
        api_key_info = st.empty()
        
        # Add OCR status information
        ocr_status = st.empty()
        if st.session_state.get('enable_ocr', True):
            ocr_status.markdown("""
            <div style="margin: 10px 0; padding: 8px; background-color: #E8F5E9; border-radius: 8px; text-align: center;">
                <span class="ocr-badge">OCR Enabled</span> 
                <small>Image-based PDFs will be processed using OCR</small>
            </div>
            """, unsafe_allow_html=True)
        else:
            ocr_status.markdown("""
            <div style="margin: 10px 0; padding: 8px; background-color: #FFEBEE; border-radius: 8px; text-align: center;">
                <span class="ocr-disabled-badge">OCR Disabled</span> 
                <small>Image-based PDFs may not be processed correctly</small>
            </div>
            """, unsafe_allow_html=True)
    
    # Initialize counters for each key if they don't exist
    if 'key_resume_counts' not in st.session_state:
        st.session_state.key_resume_counts = [0, 0, 0, 0, 0]  # Track resumes processed by each key (now 5 keys)
    
    # Process in batches
    processed_count = 0
    
    for batch_start in range(0, total_resumes, batch_size):
        batch_end = min(batch_start + batch_size, total_resumes)
        current_batch = df.iloc[batch_start:batch_end]
        batch_size_actual = len(current_batch)
        
        # Check if we need to rotate the key before processing this batch
        current_key_index = st.session_state.get('api_key_index', 0)
        current_key_count = st.session_state.key_resume_counts[current_key_index]
        
        # If this key has already processed 10 or more resumes, rotate to next key
        if current_key_count >= 10:
            next_key = get_next_api_key()
            st.session_state.key_resume_counts[st.session_state.api_key_index] = 0  # Reset count for this key
            st.toast(f"Rotated to API Key {st.session_state.api_key_index + 1} after processing 10 resumes")
        
        # Update progress info
        batch_progress = f"Processing batch {batch_start//batch_size + 1}/{math.ceil(total_resumes/batch_size)}"
        progress_text.markdown(f"<div style='text-align: center; font-size: 1.2rem;'><b>{batch_progress}</b> (Resumes {batch_start+1}-{batch_end} of {total_resumes})</div>", unsafe_allow_html=True)
        
        # Update API key info with count
        valid_keys = [key for key in GEMINI_API_KEYS if key.strip()]
        current_key_index = st.session_state.get('api_key_index', 0)
        current_count = st.session_state.key_resume_counts[current_key_index]
        api_key_info.markdown(f"""
        <div class='api-key-status'>
            Using API Key {current_key_index + 1} of {len(valid_keys)} 
            <span class='api-key-badge'>Active</span>
            <br>
            <small>Resumes processed by this key: {current_count}/10</small>
        </div>
        """, unsafe_allow_html=True)
        
        # Process batch with ThreadPoolExecutor
        batch_results = []
        with ThreadPoolExecutor(max_workers=min(batch_size, 3)) as executor:
            future_to_row = {executor.submit(process_function, row): idx for idx, row in current_batch.iterrows()}
            
            for future in concurrent.futures.as_completed(future_to_row):
                try:
                    result = future.result()
                    batch_results.append(result)
                    
                    # Increment the count for the current key
                    current_key_index = st.session_state.get('api_key_index', 0)
                    st.session_state.key_resume_counts[current_key_index] += 1
                    
                except Exception as exc:
                    logging.error(f'An error occurred: {exc}')
                    idx = future_to_row[future]
                    row = df.loc[idx]
                    # Create error result based on function type
                    if process_function == process_resume_with_details:
                        error_result = {
                            'User ID': row['user_id'],
                            'Resume Link': row['Resume link'],
                            'Project Titles': 'Error processing resume',
                            'Experiences': '',
                            'Links': '',
                            'Repo Count': None
                        }
                    elif process_function == process_resume_skills:
                        error_result = {
                            'User ID': row['user_id'],
                            'Resume Link': row['Resume link'],
                            'Skills': 'Error processing resume'
                        }
                    elif process_function == process_resume_details:
                        error_result = create_error_result(row['user_id'], row['Resume link'])
                    
                    batch_results.append(error_result)
        
        # Add batch results to overall results
        results.extend(batch_results)
        processed_count += batch_size_actual
        
        # Update progress bar
        progress_bar.progress(processed_count / total_resumes)
        
        # Pause briefly between batches to manage rate limits
        if batch_end < total_resumes:
            with st.spinner(f"Pausing briefly between batches ({batch_end}/{total_resumes} resumes processed)..."):
                time.sleep(2)  # Brief pause between batches
    
    result_df = pd.DataFrame(results)
    if processed_count == total_resumes:
        progress_container.success("✅ Processing complete!")
    
    return result_df

def extract_and_analyze_resume_projects_with_details(csv_file):
    # Handle both file objects and path strings
    if isinstance(csv_file, str):
        df = pd.read_csv(csv_file)
    else:
        df = pd.read_csv(csv_file)
    
    # Get batch size from session state or use default
    batch_size = st.session_state.get('batch_size', 5)
    
    # Process in batches
    return process_resumes_in_batches(df, process_resume_with_details, batch_size)

def extract_and_analyze_resume_skills(csv_file):
    # Handle both file objects and path strings
    if isinstance(csv_file, str):
        df = pd.read_csv(csv_file)
    else:
        df = pd.read_csv(csv_file)
    
    # Get batch size from session state or use default
    batch_size = st.session_state.get('batch_size', 5)
    
    # Process in batches
    return process_resumes_in_batches(df, process_resume_skills, batch_size)

def extract_and_analyze_resume_details(csv_file):
    # Read the uploaded CSV file into a pandas DataFrame
    if isinstance(csv_file, str):
        df = pd.read_csv(csv_file)
    else:
        df = pd.read_csv(csv_file)
    
    # Get batch size from session state or use default
    batch_size = st.session_state.get('batch_size', 5)
    
    # Process in batches
    return process_resumes_in_batches(df, process_resume_details, batch_size)

def parse_text_input(text_input):
    """
    Parse text input containing user IDs and resume links.
    Expected format: Each line contains a user ID followed by a resume link, separated by a tab or multiple spaces.
    Returns a pandas DataFrame with 'user_id' and 'Resume link' columns.
    """
    lines = text_input.strip().split('\n')
    data = []
    
    for line in lines:
        # Skip empty lines
        if not line.strip():
            continue
            
        # Split by tab or multiple spaces
        parts = re.split(r'\t|\s{2,}', line.strip())
        
        # Need at least two parts (user_id and resume link)
        if len(parts) >= 2:
            user_id = parts[0].strip()
            # The resume link might contain spaces, so join the rest
            resume_link = parts[1].strip()
            data.append({'user_id': user_id, 'Resume link': resume_link})
    
    return pd.DataFrame(data)

# Display metrics with OCR statistics
def display_metrics(df, title, count_label="Total Resumes"):
    """Display metrics in a visually appealing way"""
    st.markdown(f"<h3 style='text-align: center; margin-bottom: 15px;'>{title}</h3>", unsafe_allow_html=True)
    
    # Calculate metrics
    total_count = len(df)
    
    # Check which columns exist in the DataFrame to determine metric calculation
    if 'Project Titles' in df.columns:
        processed_count = df[~df['Project Titles'].str.contains('Failed', na=False)].shape[0]
    elif 'Skills' in df.columns:
        processed_count = df[~df['Skills'].str.contains('Failed', na=False)].shape[0]
    elif 'Full Name' in df.columns:
        processed_count = df[~df['Full Name'].str.contains('Failed', na=False)].shape[0]
    else:
        processed_count = 0
        
    success_rate = (processed_count / total_count * 100) if total_count > 0 else 0
    
    # Create metrics container
    st.markdown("""
    <div class="custom-metric-container">
        <div class="metric-card">
            <h3>{}</h3>
            <p>{}</p>
        </div>
        <div class="metric-card">
            <h3>Successfully Processed</h3>
            <p>{}</p>
        </div>
        <div class="metric-card">
            <h3>Success Rate</h3>
            <p>{:.1f}%</p>
        </div>
    </div>
    """.format(count_label, total_count, processed_count, success_rate), unsafe_allow_html=True)
    
    # Add OCR stats if OCR is enabled and we have processed some data
    if st.session_state.get('enable_ocr', True) and st.session_state.ocr_stats['scanned_count'] > 0:
        stats = st.session_state.ocr_stats
        st.markdown("""
        <h4 style='text-align: center; margin: 20px 0 10px 0;'>OCR Processing Statistics</h4>
        <div class="custom-metric-container">
            <div class="metric-card">
                <h3>Scanned PDFs</h3>
                <p>{}</p>
            </div>
            <div class="metric-card">
                <h3>OCR Success Rate</h3>
                <p>{:.1f}%</p>
            </div>
            <div class="metric-card">
                <h3>Avg OCR Time</h3>
                <p>{:.1f}s</p>
            </div>
        </div>
        """.format(
            stats.get('scanned_count', 0),
            stats.get('ocr_success_rate', 0),
            stats.get('avg_processing_time', 0)
        ), unsafe_allow_html=True)

# ========== MAIN UI FUNCTIONS ==========

def resume_projects_matching():
    local_css()
    
    # Header with logo and title
    col1, col2 = st.columns([1, 5])
    with col1:
        st.image("https://res.cloudinary.com/dg8n2jeur/image/upload/v1742020733/colvg2dikleueyuwnia3.webp", width=70)
    with col2:
        st.markdown("<h1 style='margin-bottom: 0px;'>Resume Analysis Dashboard</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #666; margin-top: 0px;'>Extract, analyze, and search through candidate resumes</p>", unsafe_allow_html=True)
    
    st.markdown("<hr style='margin: 1rem 0;'>", unsafe_allow_html=True)

    # Introduction and instructions
    with st.expander("ℹ️ About This Tool", expanded=False):
        st.markdown("### Resume Analysis Tool")
        st.markdown("This tool helps you analyze candidate resumes by extracting projects, skills, and personal details. Upload a CSV file with user IDs and resume links or paste them directly in the text box to get started.")
        
        st.markdown("#### Features:")
        st.markdown("- **Projects Analysis:** Extract project titles and technologies used")
        st.markdown("- **Skills Extraction:** Identify technical and soft skills from resumes")
        st.markdown("- **Personal Details:** Extract contact information")
        st.markdown("- **GitHub Integration:** Count repositories from GitHub profiles")
        st.markdown("- **OCR Processing:** Extract text from scanned or image-based PDFs")
        st.markdown("- **Multiple API Keys:** Automatically rotates between API keys to handle rate limits")
        st.markdown("- **Batch Processing:** Processes resumes in small batches to avoid exceeding quota limits")

    # Display current API key info
    valid_keys = [key for key in GEMINI_API_KEYS if key.strip()]
    if valid_keys:
        current_key_index = st.session_state.get('api_key_index', 0) % len(valid_keys)
        st.markdown(f"""
        <div class="api-key-status">
            Currently using API Key {current_key_index + 1} of {len(valid_keys)}
            <span class="api-key-badge">Active</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.warning("⚠️ No valid API keys configured. Please add API keys in the Settings section.")
    
    # Data input section with tabs for CSV and Text Input
    st.markdown("<h2>📁 Input Resume Data</h2>", unsafe_allow_html=True)
    
    input_method = st.radio(
        "Choose input method:",
        ["Upload CSV", "Paste Text Input"],
        key="input_method"
    )
    
    uploaded_file = None
    text_input_df = None
    
    if input_method == "Upload CSV":
        # Add sample template link and description
        st.markdown("""
        <div style="margin-bottom: 20px;">
            Upload a CSV file containing User IDs and Resume Links. 
            <a href="https://res.cloudinary.com/dho0r1xpj/raw/upload/v1732302092/Resume.csv" target="_blank">Download Template</a>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("Upload CSV file", type="csv", key="resume", label_visibility="collapsed")
        
    else:  # Text Input option
        st.markdown("""
        <div style="margin-bottom: 10px;">
            Paste user IDs and resume links below. Format: one entry per line, with user ID followed by resume link, separated by tab or multiple spaces.
            Press Ctrl+Enter to submit.
        </div>
        """, unsafe_allow_html=True)
        
        # Example format to show to the user
        example = """user123    https://example.com/resume1.pdf
user456    https://example.com/resume2.pdf"""
        
        text_input = st.text_area(
            "Enter User IDs and Resume Links:", 
            height=150,
            placeholder=example,
            key="text_input"
        )
        
        if text_input:
            text_input_df = parse_text_input(text_input)
            
            # Show preview of parsed data
            if not text_input_df.empty:
                st.markdown("<h4>Parsed Input Preview:</h4>", unsafe_allow_html=True)
                st.dataframe(text_input_df, use_container_width=True)
                
                if st.button("✅ Confirm and Proceed with Analysis", key="confirm_text_input"):
                    st.session_state['text_input_confirmed'] = True
            else:
                st.error("Could not parse any valid entries from the text input. Please check the format and try again.")
    
    # Select analysis type and run analysis
    input_data_ready = (uploaded_file is not None) or (text_input_df is not None and st.session_state.get('text_input_confirmed', False))
    
    if input_data_ready:
        # Add OCR options (only once)
        st.markdown("<h3>📷 OCR Settings for Scanned Resumes</h3>", unsafe_allow_html=True)
        with st.expander("Configure OCR", expanded=False):
            add_ocr_options_to_ui()
        
        st.markdown("<h2>📊 Select Analysis Type</h2>", unsafe_allow_html=True)
        
        # Add a dropdown to select the analysis type
        analysis_type = st.selectbox(
            "Select type of analysis to perform:",
            ["Select an option...", "Project Analysis", "Skills Analysis", "Personal Details Analysis"],
            key="analysis_selection"
        )
        
        # Initialize session state for tracking if analysis has been run
        if 'analysis_run' not in st.session_state:
            st.session_state['analysis_run'] = False
            st.session_state['current_analysis'] = None
        
        # Check if analysis type has changed
        if analysis_type != "Select an option..." and (not st.session_state['analysis_run'] or st.session_state['current_analysis'] != analysis_type):
            # Update session state
            st.session_state['analysis_run'] = True
            st.session_state['current_analysis'] = analysis_type
            
            # Clear previous results if switching analysis types
            if analysis_type == "Project Analysis":
                st.session_state.pop('resume_analysis_results', None)
            elif analysis_type == "Skills Analysis":
                st.session_state.pop('skills_analysis_results', None)
            elif analysis_type == "Personal Details Analysis":
                st.session_state.pop('personal_details_results', None)
        
        # Get the input data (either from CSV or text input)
        input_data = uploaded_file if uploaded_file is not None else text_input_df
        
        # Process data based on the selected analysis type
        if analysis_type == "Project Analysis":
            # Projects Tab
            st.markdown("<h2>Project Analysis</h2>", unsafe_allow_html=True)
            
            if 'resume_analysis_results' not in st.session_state:
                with st.spinner('⏳ Analyzing resumes for projects...'):
                    if uploaded_file is not None:
                        result_df = extract_and_analyze_resume_projects_with_details(uploaded_file)
                    else:
                        # Save the DataFrame to a temporary CSV file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp:
                            text_input_df.to_csv(tmp.name, index=False)
                            tmp_path = tmp.name
                        
                        # Process the temporary CSV file
                        result_df = extract_and_analyze_resume_projects_with_details(tmp_path)
                        
                        # Clean up
                        os.unlink(tmp_path)
                    
                    st.session_state['resume_analysis_results'] = result_df
            
            result_df = st.session_state['resume_analysis_results']
            
            # Display metrics
            display_metrics(result_df, "Project Analysis Results")
            
            # Show the data
            st.dataframe(result_df, use_container_width=True)
            
            # Separator
            st.markdown("<hr style='margin: 1.5rem 0;'>", unsafe_allow_html=True)
            
            # Project search section
            st.markdown("<h3>🔍 Search Projects by Technologies</h3>", unsafe_allow_html=True)
            input_technologies = st.text_input("Enter technologies to match (comma-separated)", 
                                              placeholder="e.g. React.js, Node.js, PostgreSQL", 
                                              key="tech_input")
            
            if input_technologies:
                tech_list = [tech.strip().lower() for tech in input_technologies.split(',')]

                # Display selected technologies as pills/tags
                tech_html = " ".join([f'<span style="background-color: #E3F2FD; padding: 5px 10px; border-radius: 15px; margin-right: 5px; font-size: 0.9em;">{tech}</span>' for tech in tech_list])
                st.markdown(f"<div style='margin: 10px 0 20px 0;'>Selected technologies: {tech_html}</div>", unsafe_allow_html=True)

                def match_technologies(project_technologies, tech_list):
                    if not isinstance(project_technologies, str):
                        return False
                    return all(tech in project_technologies.lower() for tech in tech_list)

                matched_df = result_df[result_df['Project Titles'].apply(lambda x: match_technologies(x, tech_list))]

                matched_count = len(matched_df)
                
                if not matched_df.empty:
                    st.markdown(f"<div class='success-box'><b>✅ Found {matched_count} matches</b> for your selected technologies</div>", unsafe_allow_html=True)
                    st.dataframe(matched_df[['User ID', 'Resume Link', 'Project Titles', 'Experiences', 'Links', 'Repo Count']], use_container_width=True)
                    
                    # Export options
                    output_filename_csv = "Matched_Projects.csv"
                    os.makedirs('output', exist_ok=True)
                    output_path_csv = f'output/{output_filename_csv}'
                    matched_df.to_csv(output_path_csv, index=False)
                    
                    with open(output_path_csv, "rb") as file:
                        st.download_button(
                            label="📥 Download Matched Projects as CSV",
                            data=file,
                            file_name=output_filename_csv,
                            mime="text/csv"
                        )
                else:
                    st.markdown("<div class='warning-box'>⚠️ No projects matched all the entered technologies.</div>", unsafe_allow_html=True)
            else:
                st.info("Enter technologies above to search through projects")

        elif analysis_type == "Skills Analysis":
            # Skills Tab
            st.markdown("<h2>Skills Analysis</h2>", unsafe_allow_html=True)
            
            if 'skills_analysis_results' not in st.session_state:
                with st.spinner('⏳ Analyzing resumes for skills...'):
                    if uploaded_file is not None:
                        skills_df = extract_and_analyze_resume_skills(uploaded_file)
                    else:
                        # Save the DataFrame to a temporary CSV file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp:
                            text_input_df.to_csv(tmp.name, index=False)
                            tmp_path = tmp.name
                        
                        # Process the temporary CSV file
                        skills_df = extract_and_analyze_resume_skills(tmp_path)
                        
                        # Clean up
                        os.unlink(tmp_path)
                    
                    st.session_state['skills_analysis_results'] = skills_df
            
            skills_df = st.session_state['skills_analysis_results']
            
            # Display metrics
            display_metrics(skills_df, "Skills Analysis Results")
            
            # Show the data
            st.dataframe(skills_df, use_container_width=True)
            
            # Separator
            st.markdown("<hr style='margin: 1.5rem 0;'>", unsafe_allow_html=True)
            
            # Skills search section
            st.markdown("<h3>🔍 Search Candidates by Skills</h3>", unsafe_allow_html=True)
            input_skills = st.text_input("Enter skills to match (comma-separated)", 
                                        placeholder="e.g. Python, Machine Learning, SQL", 
                                        key="skills_input")
            
            if input_skills:
                skill_list = [skill.strip().lower() for skill in input_skills.split(',')]
                
                # Display selected skills as pills/tags
                skill_html = " ".join([f'<span style="background-color: #E8F5E9; padding: 5px 10px; border-radius: 15px; margin-right: 5px; font-size: 0.9em;">{skill}</span>' for skill in skill_list])
                st.markdown(f"<div style='margin: 10px 0 20px 0;'>Selected skills: {skill_html}</div>", unsafe_allow_html=True)

                def match_skills(resume_skills, skill_list):
                    if not resume_skills or not isinstance(resume_skills, str):
                        return False
                        
                    skills = [skill.strip().lower() for skill in resume_skills.split(',')]
                    matched = []

                    for skill in skill_list:
                        skill_lower = skill.strip().lower()
                        if skill_lower == "java":
                            if "java" in skills and not any("javascript" in s for s in skills):
                                matched.append(True)
                            else:
                                matched.append(False)
                        else:
                            matched.append(any(skill_lower in s for s in skills))

                    return all(matched)
                
                matched_skills_df = skills_df[skills_df['Skills'].apply(lambda x: match_skills(x, skill_list))]

                matched_count = len(matched_skills_df)
                
                if not matched_skills_df.empty:
                    st.markdown(f"<div class='success-box'><b>✅ Found {matched_count} matches</b> for your selected skills</div>", unsafe_allow_html=True)
                    st.dataframe(matched_skills_df[['User ID', 'Resume Link', 'Skills']], use_container_width=True)
                    
                    # Export options
                    output_filename_csv = "Matched_Skills.csv"
                    os.makedirs('output', exist_ok=True)
                    output_path_csv = f'output/{output_filename_csv}'
                    matched_skills_df.to_csv(output_path_csv, index=False)
                    
                    with open(output_path_csv, "rb") as file:
                        st.download_button(
                            label="📥 Download Matched Skills as CSV",
                            data=file,
                            file_name=output_filename_csv,
                            mime="text/csv"
                        )
                else:
                    st.markdown("<div class='warning-box'>⚠️ No resumes matched all the entered skills.</div>", unsafe_allow_html=True)
            else:
                st.info("Enter skills above to search through resumes")

        elif analysis_type == "Personal Details Analysis":
            # Personal Details Tab
            st.markdown("<h2>Personal Details Analysis</h2>", unsafe_allow_html=True)
            
            if 'personal_details_results' not in st.session_state:
                with st.spinner('⏳ Extracting personal details from resumes...'):
                    if uploaded_file is not None:
                        details_df = extract_and_analyze_resume_details(uploaded_file)
                    else:
                        # Save the DataFrame to a temporary CSV file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp:
                            text_input_df.to_csv(tmp.name, index=False)
                            tmp_path = tmp.name
                        
                        # Process the temporary CSV file
                        details_df = extract_and_analyze_resume_details(tmp_path)
                        
                        # Clean up
                        os.unlink(tmp_path)
                    
                    st.session_state['personal_details_results'] = details_df
            
            details_df = st.session_state['personal_details_results']
            
            # Display metrics
            display_metrics(details_df, "Personal Details Results", "Total Candidates")
            
            # Show the data
            st.dataframe(details_df, use_container_width=True)
            
            # Separator
            st.markdown("<hr style='margin: 1.5rem 0;'>", unsafe_allow_html=True)
            
            # Export options
            output_filename_csv = "Resume_Details.csv"
            os.makedirs('output', exist_ok=True)
            output_path_csv = f'output/{output_filename_csv}'
            details_df.to_csv(output_path_csv, index=False)
            
            with open(output_path_csv, "rb") as file:
                st.download_button(
                    label="📥 Download All Details as CSV",
                    data=file,
                    file_name=output_filename_csv,
                    mime="text/csv"
                )
        
        elif analysis_type == "Select an option...":
            st.info("Please select an analysis type from the dropdown above to begin processing.")

def main():
    # Set page configuration
    st.set_page_config(
        page_title="Smart Resume Analyzer",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Apply custom CSS for dark mode toggle
    st.markdown("""
    <style>
    .dark-mode-toggle {
        position: fixed;
        top: 20px;
        right: 20px;
        z-index: 9999;
        background-color: #f1f3f4;
        border-radius: 20px;
        padding: 5px 15px;
        font-size: 0.8rem;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        display: flex;
        align-items: center;
        cursor: pointer;
    }
    .dark-mode-toggle:hover {
        background-color: #e0e0e0;
    }
    </style>
    
    <div class="dark-mode-toggle" onclick="toggleDarkMode()">
        <span id="dark-mode-icon">🌙</span>
        <span id="dark-mode-text" style="margin-left: 5px;">Dark mode</span>
    </div>
    
    <script>
    function toggleDarkMode() {
        const body = document.querySelector('body');
        const isDark = body.classList.toggle('dark-theme');
        const icon = document.getElementById('dark-mode-icon');
        const text = document.getElementById('dark-mode-text');
        
        if (isDark) {
            icon.innerText = '☀️';
            text.innerText = 'Light mode';
            document.documentElement.style.setProperty('--background-color', '#1e1e1e');
            document.documentElement.style.setProperty('--text-color', '#f0f0f0');
            document.documentElement.style.setProperty('--card-background', '#2d2d2d');
        } else {
            icon.innerText = '🌙';
            text.innerText = 'Dark mode';
            document.documentElement.style.setProperty('--background-color', '#f8f9fa');
            document.documentElement.style.setProperty('--text-color', '#212529');
            document.documentElement.style.setProperty('--card-background', '#ffffff');
        }
    }
    </script>
    """, unsafe_allow_html=True)
    
    # Sidebar with info and options
    with st.sidebar:
        st.image("https://res.cloudinary.com/dg8n2jeur/image/upload/v1742020733/colvg2dikleueyuwnia3.webp", width=70)
        st.markdown("<h2>Smart Resume Analyzer</h2>", unsafe_allow_html=True)
        st.markdown("<p>AI-powered resume analysis and candidate matching</p>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Navigation
        st.markdown("<h3>Navigation</h3>", unsafe_allow_html=True)
        page = st.radio("", ["Resume Analysis", "About"], label_visibility="collapsed")
        
        st.markdown("---")
        
        # Settings
        with st.expander("⚙️ Settings"):
            # Add the enhanced API key management UI with OCR settings
            enhanced_api_key_management_ui()
            
            # Batch size setting
            st.markdown("#### Batch Processing")
            batch_size = st.slider("Resumes per batch", min_value=1, max_value=10, value=5, 
                                  help="Process resumes in batches to manage API rate limits")
            st.session_state['batch_size'] = batch_size
        
        st.markdown("---")
        st.markdown("© 2025 Smart Resume Analyzer")
    
    # Main content
    if page == "Resume Analysis":
        resume_projects_matching()
    else:
        # About page
        st.title("About Smart Resume Analyzer")
        
        st.info("""
        ### Welcome to Smart Resume Analyzer
        
        This application helps recruiters and hiring managers analyze candidate resumes at scale using AI.
        
        #### Key features:
        * **Project Extraction:** Identify projects and technologies from resumes
        * **Skills Analysis:** Extract and categorize skills
        * **Candidate Details:** Get contact information
        * **GitHub Integration:** View repository counts
        * **OCR Processing:** Extract text from scanned or image-based PDFs
        * **Batch Processing:** Process multiple resumes simultaneously
        * **Advanced Filtering:** Find candidates matching specific criteria
        * **Multiple API Keys:** Automatically rotate between keys to handle rate limits
        
        #### How to use:
        1. Upload a CSV file with user IDs and resume links or paste them directly
        2. Configure OCR settings if working with scanned resumes
        3. Select the analysis type (Projects, Skills, or Details)
        4. Wait for the analysis to complete
        5. Search and filter the results as needed
        6. Download the filtered results for further use
        """)
        
        # How It Works section
        st.subheader("How It Works")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
            <div style="padding: 20px; background-color: white; border-radius: 8px; height: 200px; display: flex; flex-direction: column; justify-content: center; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <img src="https://img.icons8.com/fluency/48/upload-to-cloud.png" style="width: 48px; margin: 0 auto;">
                <h4>1. Upload CSV</h4>
                <p>Upload a CSV file with user IDs and resume links</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div style="padding: 20px; background-color: white; border-radius: 8px; height: 200px; display: flex; flex-direction: column; justify-content: center; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <img src="https://img.icons8.com/fluency/48/process.png" style="width: 48px; margin: 0 auto;">
                <h4>2. Process Resumes</h4>
                <p>AI analyzes and extracts key information</p>
            </div>
            """, unsafe_allow_html=True)
            
        with col3:
            st.markdown("""
            <div style="padding: 20px; background-color: white; border-radius: 8px; height: 200px; display: flex; flex-direction: column; justify-content: center; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <img src="https://img.icons8.com/fluency/48/search-client.png" style="width: 48px; margin: 0 auto;">
                <h4>3. Search & Filter</h4>
                <p>Find candidates matching your criteria</p>
            </div>
            """, unsafe_allow_html=True)
        
        # OCR capabilities section
        st.subheader("OCR Capabilities")
        st.markdown("""
        This application can extract text from scanned or image-based PDFs using Optical Character Recognition (OCR) technology:
        
        1. **Automatic Detection**: The system automatically detects if a resume is text-based or image-based
        2. **Smart Image Processing**: Applies image preprocessing to improve text recognition
        3. **Multiple Quality Levels**: Choose between fast processing or high-quality recognition
        4. **Configurable Settings**: Fine-tune OCR settings for different types of documents
        
        You can adjust OCR settings before running the analysis to optimize results for your specific resumes.
        """)
        
        # Multi-API Key section
        st.subheader("Smart Rate Limit Handling")
        st.markdown("""
        This application uses multiple API keys to handle Gemini's rate limits intelligently:
        
        1. **Automatic Key Rotation**: When one API key hits the rate limit, the system automatically switches to the next available key
        2. **Batch Processing**: Resumes are processed in small batches to avoid hitting quota limits
        3. **Error Resilience**: If errors occur with one key, the system tries with another key
        4. **Configurable Batch Size**: Adjust the batch size in settings to balance speed and API usage
        
        You can configure up to 5 different Gemini API keys in the Settings panel to maximize your processing capacity and handle even larger batches of resumes.
        """)

if __name__ == "__main__":
    main()