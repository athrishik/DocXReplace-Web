#!/usr/bin/env python
# coding: utf-8

"""
DocXReplace v3.0 - Professional Document Replacement Tool (Streamlit Edition)
Copyright 2025 Hrishik Kunduru. All rights reserved.

Professional document replacement with legal token migration support and modern UI.
"""

import streamlit as st
import tempfile
import os
import re
import zipfile
import pandas as pd
from datetime import datetime
from docx import Document
import json
import shutil
from pathlib import Path
import threading
from io import BytesIO
import base64
import time
import glob
import platform
import traceback
from typing import Dict, List, Tuple, Optional, Any

# Configure Streamlit page
st.set_page_config(
    page_title="DocXReplace v3.0 Web",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Modern CSS styling - EXACT MATCH to DocXScan
def load_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&display=swap');
    
    :root {
        --bg-primary: #0A0E1A;
        --bg-secondary: #151922;
        --bg-card: #1E2532;
        --accent-blue: #2563EB;
        --accent-green: #10B981;
        --accent-orange: #F59E0B;
        --accent-red: #EF4444;
        --accent-purple: #8B5CF6;
        --text-primary: #FFFFFF;
        --text-secondary: #E5E7EB;
        --text-muted: #9CA3AF;
        --border-color: #374151;
        --success: #10B981;
        --warning: #F59E0B;
        --error: #EF4444;
        --shadow-sm: 0 2px 4px rgba(0, 0, 0, 0.1);
        --shadow-md: 0 4px 12px rgba(0, 0, 0, 0.2);
        --shadow-lg: 0 8px 24px rgba(0, 0, 0, 0.3);
        --shadow-xl: 0 12px 32px rgba(0, 0, 0, 0.4);
    }
    
    .stApp {
        background: linear-gradient(135deg, var(--bg-primary) 0%, var(--bg-secondary) 100%);
        color: var(--text-primary);
        font-family: 'Inter', sans-serif;
    }
    
    /* ============= HEADER STYLING ============= */
    .main-header {
        background: linear-gradient(145deg, var(--bg-card), #2A3441);
        border: 2px solid var(--accent-green);
        border-radius: 16px;
        padding: 2rem;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: var(--shadow-lg);
        position: relative;
        overflow: hidden;
    }
    
    .main-header::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: linear-gradient(45deg, transparent 30%, rgba(16, 185, 129, 0.05) 50%, transparent 70%);
        animation: shimmer 3s infinite;
    }
    
    @keyframes shimmer {
        0% { transform: translateX(-100%); }
        100% { transform: translateX(100%); }
    }
    
    .main-title {
        font-size: 2.5rem;
        font-weight: 800;
        color: var(--accent-green);
        margin-bottom: 0.5rem;
        text-shadow: 0 0 20px rgba(16, 185, 129, 0.3);
        position: relative;
        z-index: 1;
    }
    
    .main-subtitle {
        font-size: 1.1rem;
        color: var(--text-secondary);
        font-weight: 500;
        margin: 0;
        position: relative;
        z-index: 1;
    }
    
    /* ============= CARD SYSTEM ============= */
    .modern-card {
        background: linear-gradient(145deg, var(--bg-card), #242B3D);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: var(--shadow-md);
        transition: all 0.3s ease;
    }
    
    .modern-card:hover {
        border-color: var(--accent-green);
        box-shadow: var(--shadow-lg);
    }
    
    .card-title {
        font-size: 1.2rem;
        font-weight: 600;
        color: var(--text-primary);
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    .status-indicator {
        width: 8px;
        height: 8px;
        background: var(--success);
        border-radius: 50%;
        margin-left: auto;
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0% { opacity: 1; }
        50% { opacity: 0.5; }
        100% { opacity: 1; }
    }
    
    /* ============= TAB SYSTEM ENHANCEMENT ============= */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: transparent !important;
        border-bottom: none !important;
        margin-bottom: 1rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: linear-gradient(145deg, var(--bg-card), #2A3441) !important;
        color: var(--text-secondary) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 10px 20px !important;
        margin: 0 !important;
        transition: all 0.3s ease !important;
        box-shadow: var(--shadow-sm) !important;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background: linear-gradient(145deg, #2A3441, var(--bg-card)) !important;
        border-color: var(--accent-green) !important;
        transform: translateY(-1px) !important;
        box-shadow: var(--shadow-md) !important;
    }
    
    .stTabs [aria-selected="true"][data-baseweb="tab"] {
        background: linear-gradient(135deg, var(--accent-green), var(--accent-blue)) !important;
        color: white !important;
        border-color: var(--accent-green) !important;
        box-shadow: var(--shadow-md) !important;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3) !important;
    }
    
    .stTabs [data-baseweb="tab-panel"] {
        padding: 0 !important;
    }
    
    /* ============= CONSOLE STYLING ============= */
    .console-area {
        background: #000000;
        color: #00FF88;
        font-family: 'JetBrains Mono', monospace;
        border: 2px solid #00FF88;
        border-radius: 8px;
        padding: 1rem;
        height: 300px;
        overflow-y: auto;
        font-size: 0.9rem;
        line-height: 1.4;
        box-shadow: var(--shadow-md);
        position: relative;
    }
    
    .console-area::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: linear-gradient(45deg, transparent 30%, rgba(0, 255, 136, 0.05) 50%, transparent 70%);
        animation: scan-line 2s infinite;
        pointer-events: none;
    }
    
    @keyframes scan-line {
        0% { transform: translateY(-100%); }
        100% { transform: translateY(100%); }
    }
    
    /* ============= METRICS AND CARDS ============= */
    .metrics-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin: 1rem 0;
    }
    
    .metric-card {
        background: linear-gradient(145deg, var(--bg-card), #242B3D);
        border: 1px solid var(--border-color);
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        transition: all 0.3s ease;
        box-shadow: var(--shadow-sm);
    }
    
    .metric-card:hover {
        transform: translateY(-4px);
        border-color: var(--accent-green);
        box-shadow: var(--shadow-md);
    }
    
    .metric-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: var(--accent-green);
        margin-bottom: 0.25rem;
        text-shadow: 0 0 10px rgba(16, 185, 129, 0.3);
    }
    
    .metric-label {
        font-size: 0.875rem;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 500;
    }
    
    /* ============= SIDEBAR STYLING ============= */
    .css-1d391kg, section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, var(--bg-secondary) 0%, #1A2332 100%) !important;
        border-right: 2px solid var(--accent-green) !important;
        box-shadow: var(--shadow-lg) !important;
    }
    
    /* Sidebar Text - High Visibility */
    .css-1d391kg *, section[data-testid="stSidebar"] *,
    section[data-testid="stSidebar"] .stMarkdown *,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stTextInput label,
    section[data-testid="stSidebar"] .stTextArea label,
    section[data-testid="stSidebar"] .stFileUploader label {
        color: var(--text-primary) !important;
        font-weight: 500 !important;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.8) !important;
    }
    
    /* Sidebar Headers */
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] h4,
    .css-1d391kg h1, .css-1d391kg h2, .css-1d391kg h3, .css-1d391kg h4 {
        color: var(--accent-green) !important;
        font-weight: 700 !important;
        text-shadow: 0 0 10px rgba(16, 185, 129, 0.5) !important;
        margin-bottom: 1rem !important;
    }
    
    /* ============= INPUT FIELD ENHANCEMENTS ============= */
    /* Text Input Fields */
    .stTextInput > div > div > input,
    section[data-testid="stSidebar"] .stTextInput > div > div > input {
        background-color: #1A1A1A !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--accent-green) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        padding: 12px !important;
        outline: none !important;
        box-shadow: var(--shadow-sm) !important;
        transition: all 0.3s ease !important;
        font-family: 'JetBrains Mono', monospace !important;
    }
    
    .stTextInput > div > div > input::placeholder {
        color: var(--text-muted) !important;
        opacity: 0.8 !important;
        font-style: italic !important;
    }
    
    .stTextInput > div > div > input:focus {
        outline: none !important;
        box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.2) !important;
        border: 2px solid var(--accent-green) !important;
        transform: translateY(-1px) !important;
    }
    
    /* Text Area Fields */
    .stTextArea > div > div > textarea,
    section[data-testid="stSidebar"] .stTextArea > div > div > textarea {
        background-color: #1A1A1A !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--accent-green) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        padding: 12px !important;
        outline: none !important;
        box-shadow: var(--shadow-sm) !important;
        font-family: 'JetBrains Mono', monospace !important;
        transition: all 0.3s ease !important;
    }
    
    .stTextArea > div > div > textarea::placeholder {
        color: var(--text-muted) !important;
        opacity: 0.8 !important;
        font-style: italic !important;
    }
    
    .stTextArea > div > div > textarea:focus {
        outline: none !important;
        box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.2) !important;
        border: 2px solid var(--accent-green) !important;
    }
    
    /* ============= SELECTBOX COMPLETE STYLING ============= */
    /* Reset all selectbox styling first */
    .stSelectbox,
    .stSelectbox *,
    .stSelectbox > div,
    .stSelectbox > div > div,
    section[data-testid="stSidebar"] .stSelectbox,
    section[data-testid="stSidebar"] .stSelectbox *,
    section[data-testid="stSidebar"] .stSelectbox > div,
    section[data-testid="stSidebar"] .stSelectbox > div > div {
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        background: transparent !important;
    }
    
    /* Style the main selectbox container */
    .stSelectbox > div > div,
    section[data-testid="stSidebar"] .stSelectbox > div > div {
        background-color: #1A1A1A !important;
        border: 2px solid var(--accent-green) !important;
        border-radius: 8px !important;
        padding: 8px 12px !important;
        color: var(--text-primary) !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.3s ease !important;
        min-height: 40px !important;
    }
    
    /* Hover and focus states */
    .stSelectbox > div > div:hover,
    section[data-testid="stSidebar"] .stSelectbox > div > div:hover {
        border-color: var(--accent-green) !important;
        box-shadow: 0 0 0 2px rgba(16, 185, 129, 0.1) !important;
    }
    
    .stSelectbox > div > div:focus-within,
    section[data-testid="stSidebar"] .stSelectbox > div > div:focus-within {
        border-color: var(--accent-green) !important;
        box-shadow: 0 0 0 2px rgba(16, 185, 129, 0.2) !important;
    }
    
    /* Style the selected value text */
    .stSelectbox [data-baseweb="select"] > div,
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div {
        background: transparent !important;
        border: none !important;
        color: var(--text-primary) !important;
        font-weight: 600 !important;
        padding: 0 !important;
    }
    
    /* Remove all visual artifacts */
    .stSelectbox [data-baseweb="select"],
    .stSelectbox div[role="button"],
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"],
    section[data-testid="stSidebar"] .stSelectbox div[role="button"] {
        background: transparent !important;
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        color: var(--text-primary) !important;
    }
    
    /* Dropdown arrow styling */
    .stSelectbox svg,
    section[data-testid="stSidebar"] .stSelectbox svg {
        color: var(--accent-green) !important;
        fill: var(--accent-green) !important;
    }
    
    /* Dropdown options panel */
    .stSelectbox [role="listbox"] {
        background-color: #1A1A1A !important;
        border: 2px solid var(--accent-green) !important;
        border-radius: 8px !important;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3) !important;
        margin-top: 4px !important;
        z-index: 9999 !important;
    }
    
    /* Individual dropdown options */
    .stSelectbox [role="option"] {
        background-color: transparent !important;
        color: var(--text-primary) !important;
        padding: 12px 16px !important;
        font-weight: 500 !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }
    
    .stSelectbox [role="option"]:hover {
        background-color: var(--accent-green) !important;
        color: white !important;
    }
    
    /* Remove all focus artifacts */
    .stSelectbox *:focus,
    .stSelectbox *:active,
    .stSelectbox *:focus-visible,
    section[data-testid="stSidebar"] .stSelectbox *:focus,
    section[data-testid="stSidebar"] .stSelectbox *:active,
    section[data-testid="stSidebar"] .stSelectbox *:focus-visible {
        outline: none !important;
        box-shadow: none !important;
        border: none !important;
    }
    
    /* ============= FILE UPLOADER STYLING ============= */
    .stFileUploader > div,
    .stFileUploader section,
    .stFileUploader [data-testid="stFileUploader"],
    section[data-testid="stSidebar"] .stFileUploader > div,
    section[data-testid="stSidebar"] .stFileUploader section,
    section[data-testid="stSidebar"] .stFileUploader [data-testid="stFileUploader"] {
        background-color: #1A1A1A !important;
        border: 3px dashed var(--accent-green) !important;
        border-radius: 12px !important;
        padding: 20px !important;
        transition: all 0.3s ease !important;
        outline: none !important;
        box-shadow: var(--shadow-sm) !important;
    }
    
    .stFileUploader > div:hover,
    section[data-testid="stSidebar"] .stFileUploader > div:hover {
        border-color: var(--accent-green) !important;
        background-color: #222222 !important;
        transform: translateY(-2px) !important;
        box-shadow: var(--shadow-md) !important;
    }
    
    /* File uploader text */
    .stFileUploader span, 
    .stFileUploader p,
    .stFileUploader div,
    .stFileUploader label,
    section[data-testid="stSidebar"] .stFileUploader span,
    section[data-testid="stSidebar"] .stFileUploader p,
    section[data-testid="stSidebar"] .stFileUploader div,
    section[data-testid="stSidebar"] .stFileUploader label {
        color: var(--text-primary) !important;
        font-weight: 600 !important;
        background-color: transparent !important;
    }
    
    /* File uploader button */
    .stFileUploader button,
    section[data-testid="stSidebar"] .stFileUploader button {
        background: linear-gradient(135deg, var(--accent-green), var(--accent-blue)) !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        outline: none !important;
        box-shadow: var(--shadow-sm) !important;
        transition: all 0.3s ease !important;
        padding: 8px 16px !important;
    }
    
    .stFileUploader button:hover,
    section[data-testid="stSidebar"] .stFileUploader button:hover {
        transform: translateY(-2px) !important;
        box-shadow: var(--shadow-md) !important;
    }
    
    /* ============= BUTTON ENHANCEMENTS ============= */
    .stButton > button {
        background: linear-gradient(145deg, var(--bg-card), #374151) !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--border-color) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        padding: 10px 20px !important;
        transition: all 0.3s ease !important;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.8) !important;
        outline: none !important;
        box-shadow: var(--shadow-sm) !important;
        position: relative !important;
        overflow: hidden !important;
    }
    
    .stButton > button::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.1), transparent);
        transition: left 0.5s ease;
    }
    
    .stButton > button:hover {
        background: linear-gradient(145deg, var(--accent-green), var(--accent-blue)) !important;
        border-color: var(--accent-green) !important;
        transform: translateY(-2px) !important;
        box-shadow: var(--shadow-md) !important;
        color: white !important;
    }
    
    .stButton > button:hover::before {
        left: 100%;
    }
    
    .stButton > button:focus {
        outline: none !important;
        box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.3) !important;
    }
    
    .stButton > button:active {
        transform: translateY(0) !important;
    }
    
    /* Sidebar Buttons */
    section[data-testid="stSidebar"] .stButton > button {
        width: 100% !important;
        margin: 4px 0 !important;
    }
    
    /* Disabled buttons */
    .stButton > button:disabled {
        background: linear-gradient(145deg, #2A2A2A, #1A1A1A) !important;
        color: var(--text-muted) !important;
        border-color: var(--text-muted) !important;
        transform: none !important;
        box-shadow: none !important;
        cursor: not-allowed !important;
    }
    
    /* ============= PROGRESS BAR STYLING ============= */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, var(--accent-green), var(--accent-blue)) !important;
        border-radius: 6px !important;
        box-shadow: 0 0 10px rgba(16, 185, 129, 0.5) !important;
        position: relative !important;
        overflow: hidden !important;
    }
    
    .stProgress > div > div > div::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: linear-gradient(45deg, transparent 30%, rgba(255, 255, 255, 0.2) 50%, transparent 70%);
        animation: progress-shine 2s infinite;
    }
    
    @keyframes progress-shine {
        0% { transform: translateX(-100%); }
        100% { transform: translateX(100%); }
    }
    
    /* ============= ALERTS AND NOTIFICATIONS ============= */
    .stWarning, .stError, .stSuccess, .stInfo {
        border-radius: 8px !important;
        font-weight: 500 !important;
        border: none !important;
        box-shadow: var(--shadow-sm) !important;
        position: relative !important;
        overflow: hidden !important;
    }
    
    .stSuccess {
        background: linear-gradient(145deg, var(--success), #059669) !important;
        color: white !important;
    }
    
    .stError {
        background: linear-gradient(145deg, var(--error), #DC2626) !important;
        color: white !important;
    }
    
    .stWarning {
        background: linear-gradient(145deg, var(--warning), #D97706) !important;
        color: white !important;
    }
    
    .stInfo {
        background: linear-gradient(145deg, var(--accent-blue), #1D4ED8) !important;
        color: white !important;
    }
    
    /* ============= LABEL OVERRIDES ============= */
    .stSelectbox > label,
    .stSelectbox label,
    section[data-testid="stSidebar"] .stSelectbox > label,
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] label,
    .css-1d391kg label,
    .css-1d391kg .stSelectbox label,
    .css-1d391kg .stSelectbox > label {
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        background: transparent !important;
        color: var(--text-primary) !important;
        font-weight: 600 !important;
        padding: 0 !important;
        margin: 0 0 8px 0 !important;
        -webkit-appearance: none !important;
        -moz-appearance: none !important;
        appearance: none !important;
        font-size: 14px !important;
    }
    
    /* Hide Streamlit elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}
    </style>
    """, unsafe_allow_html=True)

class SessionState:
    """Manage session state variables"""
    @staticmethod
    def init():
        if 'replacement_map' not in st.session_state:
            st.session_state.replacement_map = {}
        if 'loaded_files' not in st.session_state:
            st.session_state.loaded_files = []
        if 'temp_directories' not in st.session_state:
            st.session_state.temp_directories = []
        if 'backup_history' not in st.session_state:
            st.session_state.backup_history = []
        if 'console_messages' not in st.session_state:
            st.session_state.console_messages = ["[READY] DocXReplace v3.0 initialized", 
                                                "[READY] Load files and replacement patterns to begin"]
        if 'process_progress' not in st.session_state:
            st.session_state.process_progress = 0
        if 'process_status' not in st.session_state:
            st.session_state.process_status = "Ready to process"
        if 'process_running' not in st.session_state:
            st.session_state.process_running = False
        if 'results' not in st.session_state:
            st.session_state.results = []
        if 'regex_mode' not in st.session_state:
            st.session_state.regex_mode = False

class DocumentProcessor:
    """Handle document processing operations"""
    
    @staticmethod
    def validate_replacement_map(replacement_map: Dict[str, str]) -> List[str]:
        """Validate replacement map for common issues"""
        errors = []
        
        for old_text, new_text in replacement_map.items():
            if not old_text.strip():
                errors.append("Empty search pattern found")
            if len(old_text) > 1000:
                errors.append(f"Search pattern too long: {old_text[:50]}...")
            if not isinstance(new_text, str):
                errors.append(f"Invalid replacement type for '{old_text}': expected string")
        
        return errors
    
    @staticmethod
    def perform_replacement_in_doc(doc: Document, file_path: str, 
                                 replacement_map: Dict[str, str], 
                                 regex_mode: bool = False) -> Tuple[int, List[Dict]]:
        """Perform replacements in a document with enhanced error handling"""
        replacements_made = 0
        replacement_details = []
        
        try:
            # Process paragraphs
            for para_idx, para in enumerate(doc.paragraphs):
                original_text = para.text
                modified_text = original_text
                
                for old_text, new_text in replacement_map.items():
                    try:
                        if regex_mode:
                            if "{{match}}" in new_text:
                                def replace_with_match(match):
                                    matched_groups = match.groups()
                                    replacement = new_text
                                    if matched_groups:
                                        for i, group in enumerate(matched_groups):
                                            replacement = replacement.replace("{{match}}", group, 1)
                                    else:
                                        replacement = replacement.replace("{{match}}", match.group(0))
                                    return replacement
                                modified_text = re.sub(old_text, replace_with_match, modified_text)
                            else:
                                modified_text = re.sub(old_text, new_text, modified_text)
                        else:
                            if old_text in modified_text:
                                modified_text = modified_text.replace(old_text, new_text)
                    except re.error as e:
                        log_message(f"⚠️ Invalid regex pattern '{old_text}': {e}")
                        continue
                    except Exception as e:
                        log_message(f"❌ Error processing pattern '{old_text}': {e}")
                        continue
                
                if modified_text != original_text:
                    try:
                        para.clear()
                        para.add_run(modified_text)
                    except Exception:
                        para.text = modified_text
                    
                    replacements_made += 1
                    replacement_details.append({
                        'location': f'paragraph_{para_idx}',
                        'original': original_text[:100] + '...' if len(original_text) > 100 else original_text,
                        'modified': modified_text[:100] + '...' if len(modified_text) > 100 else modified_text
                    })
            
            # Process tables
            for table_idx, table in enumerate(doc.tables):
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        for para_idx, para in enumerate(cell.paragraphs):
                            original_text = para.text
                            modified_text = original_text
                            
                            for old_text, new_text in replacement_map.items():
                                try:
                                    if regex_mode:
                                        if "{{match}}" in new_text:
                                            def replace_with_match(match):
                                                matched_groups = match.groups()
                                                replacement = new_text
                                                if matched_groups:
                                                    for i, group in enumerate(matched_groups):
                                                        replacement = replacement.replace("{{match}}", group, 1)
                                                else:
                                                    replacement = replacement.replace("{{match}}", match.group(0))
                                                return replacement
                                            modified_text = re.sub(old_text, replace_with_match, modified_text)
                                        else:
                                            modified_text = re.sub(old_text, new_text, modified_text)
                                    else:
                                        if old_text in modified_text:
                                            modified_text = modified_text.replace(old_text, new_text)
                                except re.error:
                                    continue
                                except Exception as e:
                                    log_message(f"❌ Error in table processing: {e}")
                                    continue
                            
                            if modified_text != original_text:
                                try:
                                    para.clear()
                                    para.add_run(modified_text)
                                except Exception:
                                    para.text = modified_text
                                
                                replacements_made += 1
                                replacement_details.append({
                                    'location': f'table_{table_idx}_row_{row_idx}_cell_{cell_idx}_para_{para_idx}',
                                    'original': original_text[:50] + '...' if len(original_text) > 50 else original_text,
                                    'modified': modified_text[:50] + '...' if len(modified_text) > 50 else modified_text
                                })
        
        except Exception as e:
            log_message(f"❌ Critical error processing document {file_path}: {e}")
            raise
        
        return replacements_made, replacement_details

def log_message(message, console_placeholder=None):
    """Add message to console log"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    formatted_msg = f"[{timestamp}] {message}"
    st.session_state.console_messages.append(formatted_msg)
    
    # Keep only last 30 messages
    if len(st.session_state.console_messages) > 30:
        st.session_state.console_messages = st.session_state.console_messages[-30:]
    
    # Update console display if placeholder provided
    if console_placeholder:
        console_text = '\n'.join(st.session_state.console_messages)
        console_placeholder.markdown(
            f'<div class="console-area">{console_text}</div>',
            unsafe_allow_html=True
        )

def clear_console():
    """Clear console messages"""
    st.session_state.console_messages = ["[READY] Console cleared"]

def create_replacement_template():
    """Create token replacement template JSON"""
    template = {
        "<<FileService.": "<<NewFileService.",
        "</ff>": "<<PAGE_BREAK>>",
        "</pp>": "<<HARD_RETURN>>",
        "<backspace>": "<<BACKSPACE>>",
        "<<STNDRDTH": "<<STANDARD_TH",
        "<c>": "<<CENTER>>",
        "<u>": "<<UNDERLINE>>",
        "<i>": "<<ITALIC>>",
        "<bold>": "<<BOLD>>",
        "</nobullet>": "<<NO_BULLET>>",
        "[[MCOMPUTEINTO(<<": "<<MCOMPUTE_INTO(",
        "[[SCOMPUTEINTO(": "<<SCOMPUTE_INTO(",
        "[[ABORTIIF": "<<ABORT_IF",
        "PROMTINTO(": "<<PROMPT_INTO(",
        "PROMTINTOIIF(": "<<PROMPT_INTO_IF(",
        "<<Checklist.": "<<CHECKLIST.",
        "TABLE(": "<<TABLE(",
        "<<jfig": "<<JFIG",
        "{ATTY": "<<ESIGN_ATTORNEY"
    }
    
    return json.dumps(template, indent=2)

def create_regex_template():
    """Create regex replacement template JSON"""
    template = {
        r"<<FileService\.\\w*": "<<NewFileService.{{match}}",
        r"<<BLTO\\d+": "<<BULLET_ORDERED_{{match}}>>",
        r"<<BLT#\\d+": "<<BULLET_NUMBERED_{{match}}>>",
        r"\\[\\[(\\w+)COMPUTEINTO\\(": "<<{{match}}_INTO(",
        r"PROMT(\\w*)\\(": "<<PROMPT_{{match}}(",
        r"<<Special\\.(\\w*)": "<<SPECIAL.{{match}}",
        r"<<Tracker\\.(\\w+)>>~(\\w+):": "<<TRACKER.{{match}}_FORMAT>>",
        r"\\{ATTY(\\w*)": "<<ESIGN_{{match}}",
        r"<(\\w+)>": "<<{{match}}>>",
        r"</(\\w+)>": "<<END_{{match}}>>",
        r"[+-]\\d+\\|<<Special\\.ToDay:": "<<SPECIAL.DATE_OFFSET_{{match}}>>"
    }
    
    return json.dumps(template, indent=2)

def load_files_from_folder(folder_path: str):
    """Load files from folder"""
    st.session_state.loaded_files = []
    if not os.path.exists(folder_path):
        return
    
    file_count = 0
    try:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.docx') and not file.startswith('~'):
                    full_path = os.path.join(root, file)
                    if os.path.exists(full_path):
                        st.session_state.loaded_files.append(full_path)
                        file_count += 1
        
        log_message(f"📁 Loaded {file_count} files from folder: {os.path.basename(folder_path)}")
        
    except Exception as e:
        log_message(f"❌ Error scanning folder: {str(e)}")

def load_files_from_zip(zip_path: str):
    """Load files from ZIP archive"""
    st.session_state.loaded_files = []
    
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            temp_dir = tempfile.mkdtemp(prefix='docx_replace_')
            st.session_state.temp_directories.append(temp_dir)
            
            # Extract .docx files
            for zip_item in zip_ref.namelist():
                if zip_item.endswith('.docx') and not zip_item.startswith('~') and not zip_item.endswith('/'):
                    file_data = zip_ref.read(zip_item)
                    filename = os.path.basename(zip_item)
                    extracted_path = os.path.join(temp_dir, filename)
                    
                    with open(extracted_path, 'wb') as f:
                        f.write(file_data)
                    
                    st.session_state.loaded_files.append(extracted_path)
            
            log_message(f"📦 Extracted {len(st.session_state.loaded_files)} files from ZIP: {os.path.basename(zip_path)}")
            log_message(f"📁 Temporary folder: {temp_dir}")
            
    except Exception as e:
        st.error(f"Failed to load ZIP file: {str(e)}")
        log_message(f"❌ Error loading ZIP: {str(e)}")

def load_files_from_excel(excel_path: str):
    """Load files from Excel scan results"""
    try:
        df = pd.read_excel(excel_path)
        if 'File Path' in df.columns:
            all_paths = df['File Path'].tolist()
            st.session_state.loaded_files = [path for path in all_paths if os.path.exists(path)]
            missing_count = len(all_paths) - len(st.session_state.loaded_files)
            
            log_message(f"📊 Loaded {len(st.session_state.loaded_files)} files from Excel: {os.path.basename(excel_path)}")
            
            if missing_count > 0:
                log_message(f"⚠️ {missing_count} files from Excel list were not found")
        else:
            st.error("Excel file must contain 'File Path' column")
    except Exception as e:
        st.error(f"Failed to load Excel file: {str(e)}")

def cleanup_temp_files():
    """Clean up temporary extracted files"""
    if not st.session_state.temp_directories:
        log_message("ℹ️ No temporary directories to clean up")
        return
    
    cleaned_count = 0
    for temp_dir in st.session_state.temp_directories:
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
                cleaned_count += 1
                log_message(f"🧹 Cleaned up temporary folder: {temp_dir}")
        except Exception as e:
            log_message(f"⚠️ Could not clean up {temp_dir}: {str(e)}")
    
    st.session_state.temp_directories = []
    if cleaned_count > 0:
        log_message(f"✅ Cleaned up {cleaned_count} temporary directories")
        st.session_state.loaded_files = []

def create_output_copy(file_path: str, output_root: str, session_timestamp: str = None) -> Tuple[str, str]:
    """Create a copy of the file in output directory for modification"""
    if session_timestamp is None:
        session_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    output_dir = os.path.join(output_root, f"modified_{session_timestamp}")
    
    try:
        os.makedirs(output_dir, exist_ok=True)
    except PermissionError:
        user_docs = os.path.expanduser("~/Documents")
        output_dir = os.path.join(user_docs, "DocReplace_Output", f"modified_{session_timestamp}")
        os.makedirs(output_dir, exist_ok=True)
        log_message(f"⚠️ Using fallback output location: {output_dir}")
    
    output_filename = os.path.basename(file_path)
    base_name, ext = os.path.splitext(output_filename)
    counter = 1
    output_path = os.path.join(output_dir, output_filename)
    
    while os.path.exists(output_path):
        output_filename = f"{base_name}_{counter}{ext}"
        output_path = os.path.join(output_dir, output_filename)
        counter += 1
    
    try:
        shutil.copy2(file_path, output_path)
        log_message(f"📋 Created working copy: {os.path.basename(file_path)} → {output_filename}")
    except Exception as e:
        log_message(f"❌ Copy failed for {os.path.basename(file_path)}: {str(e)}")
        raise
    
    return output_dir, output_path

def process_documents(mode: str, output_folder: str = None, progress_placeholder=None, console_placeholder=None):
    """Main document processing function"""
    if not st.session_state.loaded_files:
        st.error("Please load files to process first.")
        return
    
    if not st.session_state.replacement_map:
        st.error("Please load a replacement file first.")
        return
    
    if mode != "Dry Run (preview only)" and "Modified Copies" in mode and not output_folder:
        st.error("Please select an output folder.")
        return
    
    total_files = len(st.session_state.loaded_files)
    processed_files = 0
    modified_files = 0
    total_replacements = 0
    current_output_dir = None
    session_timestamp = None
    start_time = time.time()
    
    log_message(f"🚀 Starting {mode} on {total_files} files...", console_placeholder)
    
    for i, file_path in enumerate(st.session_state.loaded_files):
        if not os.path.exists(file_path):
            log_message(f"⚠️ File not found: {file_path}", console_placeholder)
            continue
        
        try:
            # Update progress
            progress = int((i / total_files) * 100)
            st.session_state.process_progress = progress
            st.session_state.process_status = f"Processing {os.path.basename(file_path)[:20]}..."
            
            if progress_placeholder:
                progress_placeholder.progress(progress / 100)
            
            # Load document
            doc = Document(file_path)
            
            # Perform replacements
            replacements_made, replacement_details = DocumentProcessor.perform_replacement_in_doc(
                doc, file_path, st.session_state.replacement_map, st.session_state.regex_mode)
            
            if replacements_made > 0:
                if mode == "Dry Run (preview only)":
                    log_message(f"🔍 Would modify {os.path.basename(file_path)}: {replacements_made} replacements", console_placeholder)
                    modified_files += 1
                else:
                    if "Modified Copies" in mode:
                        # Create output copy
                        if current_output_dir is None:
                            session_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            current_output_dir, output_file_path = create_output_copy(
                                file_path, output_folder, session_timestamp)
                        else:
                            _, output_file_path = create_output_copy(
                                file_path, output_folder, session_timestamp)
                        
                        # Save modified document to output copy
                        doc.save(output_file_path)
                        modified_files += 1
                        total_replacements += replacements_made
                        
                        log_message(f"✅ Created modified copy of {os.path.basename(file_path)}: {replacements_made} replacements", console_placeholder)
                    else:
                        # In-place replacement
                        doc.save(file_path)
                        modified_files += 1
                        total_replacements += replacements_made
                        
                        log_message(f"✅ Modified {os.path.basename(file_path)}: {replacements_made} replacements", console_placeholder)
            else:
                log_message(f"➖ No changes needed: {os.path.basename(file_path)}", console_placeholder)
            
            processed_files += 1
            
        except Exception as e:
            log_message(f"❌ Error processing {os.path.basename(file_path)}: {str(e)}", console_placeholder)
        finally:
            # Clean up document from memory
            if 'doc' in locals():
                del doc
    
    # Add to backup history if outputs were created
    if current_output_dir and mode != "Dry Run (preview only)":
        st.session_state.backup_history.append({
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'backup_dir': current_output_dir,
            'files_modified': modified_files,
            'total_replacements': total_replacements,
            'mode': mode
        })
    
    # Final summary
    elapsed_total = time.time() - start_time
    
    if mode == "Dry Run (preview only)":
        log_message(f"\n📋 Dry Run Complete:", console_placeholder)
        log_message(f"   • Files processed: {processed_files}", console_placeholder)
        log_message(f"   • Files that would be modified: {modified_files}", console_placeholder)
        log_message(f"   • Time elapsed: {elapsed_total:.1f}s", console_placeholder)
        st.success(f"Dry run completed! {modified_files} files would be modified")
    else:
        log_message(f"\n🎉 Replacement Complete:", console_placeholder)
        log_message(f"   • Files processed: {processed_files}", console_placeholder)
        log_message(f"   • Files modified: {modified_files}", console_placeholder)
        log_message(f"   • Total replacements: {total_replacements}", console_placeholder)
        log_message(f"   • Time elapsed: {elapsed_total:.1f}s", console_placeholder)
        if current_output_dir:
            log_message(f"   • Output folder: {current_output_dir}", console_placeholder)
        st.success(f"Replacement completed! {modified_files} files modified with {total_replacements} total replacements")
    
    # Update final progress
    st.session_state.process_progress = 100
    st.session_state.process_status = "Processing completed!"
    if progress_placeholder:
        progress_placeholder.progress(1.0)
    
    # Store results
    st.session_state.results = {
        'processed_files': processed_files,
        'modified_files': modified_files,
        'total_replacements': total_replacements,
        'output_dir': current_output_dir,
        'mode': mode
    }

def create_zip_download(output_dir: str, zip_name: str = "replaced_files"):
    """Create ZIP file for download"""
    try:
        zip_buffer = BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add all files from output directory
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, output_dir)
                    zipf.write(file_path, arcname)
        
        zip_buffer.seek(0)
        return zip_buffer.getvalue()
        
    except Exception as e:
        st.error(f"Error creating ZIP: {str(e)}")
        return None

def format_file_size(size_bytes):
    """Format file size in human readable format"""
    if size_bytes == 0:
        return "0 B"
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1
    return f"{size_bytes:.1f} {size_names[i]}"

def get_files_stats():
    """Get total size and file count of loaded files"""
    total_size = 0
    file_count = len(st.session_state.loaded_files)
    
    for file_path in st.session_state.loaded_files:
        try:
            size = os.path.getsize(file_path)
            total_size += size
        except OSError:
            continue
    
    size_mb = round(total_size / 1024 / 1024, 1)
    return f"{file_count} files ({size_mb} MB total)"

def main():
    """Main application"""
    load_css()
    SessionState.init()
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1 class="main-title">🔄 DocXReplace v3.0</h1>
        <p class="main-subtitle">Professional document replacement with legal token migration support</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar Configuration
    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        
        # File Loading Section
        st.markdown("#### 📂 File Loading")
        
        # File source tabs
        file_tab1, file_tab2, file_tab3 = st.tabs(["📁 Folder", "📊 Excel", "📦 ZIP"])
        
        with file_tab1:
            if st.button("📁 Browse Folder", use_container_width=True, key="browse_folder_btn"):
                # This would require a folder picker component in real deployment
                st.info("In a real deployment, this would open a folder picker dialog")
            
            folder_path = st.text_input(
                "Folder Path",
                placeholder="Enter folder path containing .docx files",
                help="Path to folder containing documents to process",
                key="folder_path_input"
            )
            
            if folder_path and os.path.exists(folder_path):
                if st.button("Load from Folder", use_container_width=True, key="load_folder_btn"):
                    load_files_from_folder(folder_path)
                    st.rerun()
        
        with file_tab2:
            excel_file = st.file_uploader(
                "Upload Excel from DocXScan",
                type=['xlsx'],
                help="Excel file from DocXScan with file paths",
                key="excel_uploader"
            )
            
            if excel_file is not None:
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(excel_file.read())
                    temp_excel_path = tmp_file.name
                
                try:
                    load_files_from_excel(temp_excel_path)
                    st.success(f"✅ Loaded files from Excel: {excel_file.name}")
                finally:
                    os.unlink(temp_excel_path)
        
        with file_tab3:
            zip_file = st.file_uploader(
                "Upload ZIP Archive",
                type=['zip'],
                help="ZIP file containing .docx files to process",
                key="zip_uploader"
            )
            
            if zip_file is not None:
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_file:
                    tmp_file.write(zip_file.read())
                    temp_zip_path = tmp_file.name
                
                try:
                    load_files_from_zip(temp_zip_path)
                    st.success(f"✅ Loaded files from ZIP: {zip_file.name}")
                finally:
                    os.unlink(temp_zip_path)
        
        # Show loaded files status
        if st.session_state.loaded_files:
            files_stats = get_files_stats()
            st.success(f"📄 Loaded: {files_stats}")
        else:
            st.info("📄 No files loaded")
        
        st.markdown("---")
        
        # Replacement Configuration
        st.markdown("#### 🔄 Replacement Configuration")
        
        # Replacement file upload
        replacement_file = st.file_uploader(
            "Upload Replacement JSON",
            type=['json'],
            help="JSON file containing find/replace patterns",
            key="replacement_uploader"
        )
        
        if replacement_file is not None:
            try:
                replacement_data = json.load(replacement_file)
                st.session_state.replacement_map = replacement_data
                st.success(f"✅ Loaded {len(replacement_data)} replacement patterns")
                log_message(f"✅ Loaded {len(replacement_data)} replacements from {replacement_file.name}")
            except Exception as e:
                st.error(f"❌ Error loading replacement file: {str(e)}")
                log_message(f"❌ Replacement load error: {str(e)}")
        
        # Regex mode
        st.session_state.regex_mode = st.checkbox(
            "Enable Regex Mode",
            value=st.session_state.regex_mode,
            help="Enable regular expression pattern matching",
            key="regex_mode_checkbox"
        )
        
        # Template creation
        template_col1, template_col2 = st.columns(2)
        
        with template_col1:
            if st.button("📄 Standard Template", use_container_width=True, key="standard_template_btn"):
                template_json = create_replacement_template()
                st.download_button(
                    label="💾 Download Standard",
                    data=template_json,
                    file_name="standard_replacement_template.json",
                    mime="application/json",
                    use_container_width=True,
                    key="download_standard_template_btn"
                )
                log_message("📄 Standard replacement template created")
        
        with template_col2:
            if st.button("🔧 Regex Template", use_container_width=True, key="regex_template_btn"):
                regex_template_json = create_regex_template()
                st.download_button(
                    label="💾 Download Regex",
                    data=regex_template_json,
                    file_name="regex_replacement_template.json",
                    mime="application/json",
                    use_container_width=True,
                    key="download_regex_template_btn"
                )
                log_message("🔧 Regex replacement template created")
        
        # Show replacement status
        if st.session_state.replacement_map:
            st.success(f"🔄 Loaded: {len(st.session_state.replacement_map)} patterns")
        else:
            st.info("🔄 No replacement patterns loaded")
        
        st.markdown("---")
        
        # Processing Options
        st.markdown("#### ⚙️ Processing Options")
        
        # Processing mode
        processing_mode = st.selectbox(
            "Processing Mode",
            ["Dry Run (preview only)", 
             "Create Modified Copies (originals untouched)", 
             "In-place Replace (modify originals)"],
            help="Select how to process the documents",
            key="processing_mode_selector"
        )
        
        # Output folder (only for modified copies mode)
        output_folder = None
        if "Modified Copies" in processing_mode:
            output_folder = st.text_input(
                "Output Folder",
                placeholder="Enter output folder path",
                help="Folder where modified copies will be saved",
                key="output_folder_input"
            )
            
            if output_folder and not os.path.exists(output_folder):
                st.warning("⚠️ Output folder does not exist")
            elif output_folder:
                st.success("✅ Output folder is valid")
        
        st.markdown("---")
        
        # Utility Functions
        st.markdown("#### 🛠️ Utilities")
        
        if st.button("🧹 Clean Temp Files", use_container_width=True, key="clean_temp_btn"):
            cleanup_temp_files()
            st.rerun()
        
        if st.button("📋 Clear Console", use_container_width=True, key="clear_console_btn"):
            clear_console()
            st.rerun()
    
    # Main content area
    col1, col2 = st.columns([2, 1], gap="medium")
    
    with col1:
        # Processing Controls
        st.markdown("""
        <div class="modern-card">
            <div class="card-title">🚀 Processing Controls <div class="status-indicator"></div></div>
        </div>
        """, unsafe_allow_html=True)
        
        # Progress display
        progress_placeholder = st.empty()
        status_placeholder = st.empty()
        
        if st.session_state.process_progress > 0:
            progress_placeholder.progress(st.session_state.process_progress / 100)
            status_placeholder.info(f"Status: {st.session_state.process_status}")
        
        # Check if processing can be started
        can_process = (
            st.session_state.loaded_files and 
            st.session_state.replacement_map and 
            not st.session_state.process_running
        )
        
        # Additional validation for output folder
        if "Modified Copies" in processing_mode and not output_folder:
            can_process = False
        
        col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])
        
        with col_btn1:
            if st.button("🚀 Start Processing", disabled=not can_process, use_container_width=True, key="start_process_btn"):
                if can_process:
                    st.session_state.process_running = True
                    
                    # Create placeholders for real-time updates
                    console_placeholder = st.empty()
                    
                    # Execute processing
                    try:
                        process_documents(
                            processing_mode,
                            output_folder,
                            progress_placeholder,
                            console_placeholder
                        )
                        
                        if st.session_state.results:
                            if st.session_state.results['modified_files'] > 0:
                                st.success(f"🎉 Processing completed! Modified {st.session_state.results['modified_files']} files")
                                st.balloons()
                            else:
                                st.info("ℹ️ Processing completed but no files were modified")
                    
                    except Exception as e:
                        st.error(f"❌ Processing failed: {str(e)}")
                        log_message(f"❌ Processing error: {str(e)}")
                    
                    finally:
                        st.session_state.process_running = False
        
        with col_btn2:
            if st.button("📊 Results", use_container_width=True, key="results_btn"):
                if st.session_state.results:
                    st.info(f"📊 Last run: {st.session_state.results['modified_files']} files modified")
                else:
                    st.warning("No processing results available")
        
        with col_btn3:
            if st.button("🔄 Reset", use_container_width=True, key="reset_btn"):
                st.session_state.loaded_files = []
                st.session_state.replacement_map = {}
                st.session_state.process_progress = 0
                st.session_state.process_status = "Ready to process"
                st.session_state.results = []
                clear_console()
                st.rerun()
        
        # Validation and Help Section
        st.markdown("""
        <div class="modern-card">
            <div class="card-title">📖 Help & Validation <div class="status-indicator"></div></div>
        </div>
        """, unsafe_allow_html=True)
        
        help_tab1, help_tab2, help_tab3 = st.tabs(["🔧 Pattern Validation", "📚 Token Help", "🔍 Regex Help"])
        
        with help_tab1:
            if st.button("🔍 Validate Patterns", use_container_width=True, key="validate_patterns_btn"):
                if not st.session_state.replacement_map:
                    st.warning("No replacement patterns loaded to validate")
                else:
                    errors = DocumentProcessor.validate_replacement_map(st.session_state.replacement_map)
                    
                    # Validate regex patterns if regex mode is enabled
                    if st.session_state.regex_mode:
                        regex_errors = []
                        for pattern in st.session_state.replacement_map.keys():
                            try:
                                re.compile(pattern)
                            except re.error as e:
                                regex_errors.append(f"'{pattern}': {e}")
                        errors.extend(regex_errors)
                    
                    if errors:
                        st.error(f"Found {len(errors)} validation errors:")
                        for error in errors[:5]:  # Show first 5 errors
                            st.write(f"• {error}")
                        if len(errors) > 5:
                            st.write(f"... and {len(errors)-5} more errors")
                    else:
                        st.success("✅ All patterns are valid!")
        
        with help_tab2:
            st.markdown("""
            **Common Legal Token Patterns:**
            
            **Text Formatting:**
            - `<<FileService.` → Fileservice  
            - `<c>` → Center
            - `<u>` → Underline
            - `<i>` → Italic
            - `<bold>` → Bold
            
            **Document Structure:**
            - `</ff>` → Page Break
            - `</pp>` → Hard Return
            - `<backspace>` → Backspace
            
            **Computation Tokens:**
            - `[[MCOMPUTEINTO(<<` → MCOMPUTE INTO
            - `[[SCOMPUTEINTO(` → SCOMPUTE INTO
            - `[[ABORTIIF` → Abort if
            
            **Special Functions:**
            - `<<Checklist.` → CHECKLIST
            - `TABLE(` → TABLE
            - `<<jfig` → JFIG
            - `{ATTY` → ESIGN
            
            **Date/Tracker Tokens:**
            - `+91|<<Special.ToDay:` → Add 91 days
            - `<<Tracker.MortDate>>~MMMM dd:` → Date format
            """)
        
        with help_tab3:
            st.markdown("""
            **Regex Patterns for Legal Tokens:**
            
            **Basic Patterns:**
            - `\\<` and `\\>` - Escape angle brackets
            - `\\[` and `\\]` - Escape square brackets  
            - `\\w+` - One or more word characters
            - `\\d+` - One or more digits
            
            **Legal Token Examples:**
            - `<<FileService\\.\\w*` → `<<NewFileService.{{match}}`
            - `<<BLT(#|O)(\\d+)` → `<<BULLET_{{match}}_{{match}}>>`
            - `\\[\\[(\\w+)COMPUTEINTO\\(` → `<<{{match}}_COMPUTE_INTO(`
            - `<(\\w+)>` → `<<{{match}}>>`
            
            **Special Replacement:**
            - `{{match}}` - Replaced with captured group
            - Use multiple `{{match}}` for multiple groups
            - Groups replaced in order
            
            **Tips:**
            - Always escape special characters: `< > [ ] . ( )`
            - Test with simple tokens first
            - Use Dry Run to preview changes
            """)
        
        # Results and Downloads Section
        if st.session_state.results:
            st.markdown("""
            <div class="modern-card">
                <div class="card-title">📊 Processing Results <div class="status-indicator"></div></div>
            </div>
            """, unsafe_allow_html=True)
            
            # Metrics
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            
            with col_m1:
                st.metric("📄 Files Processed", st.session_state.results['processed_files'])
            
            with col_m2:
                st.metric("✅ Files Modified", st.session_state.results['modified_files'])
            
            with col_m3:
                st.metric("🔄 Total Replacements", st.session_state.results['total_replacements'])
            
            with col_m4:
                st.metric("📁 Mode", st.session_state.results['mode'][:10] + "...")
            
            # Download options
            if st.session_state.results.get('output_dir') and os.path.exists(st.session_state.results['output_dir']):
                col_dl1, col_dl2 = st.columns(2)
                
                with col_dl1:
                    zip_data = create_zip_download(st.session_state.results['output_dir'], "replaced_files")
                    if zip_data:
                        st.download_button(
                            label="📦 Download Modified Files",
                            data=zip_data,
                            file_name="replaced_files.zip",
                            mime="application/zip",
                            use_container_width=True,
                            key="download_results_btn"
                        )
                
                with col_dl2:
                    # Create summary report
                    summary_data = {
                        'Processing Summary': [
                            f"Mode: {st.session_state.results['mode']}",
                            f"Files Processed: {st.session_state.results['processed_files']}",
                            f"Files Modified: {st.session_state.results['modified_files']}",
                            f"Total Replacements: {st.session_state.results['total_replacements']}",
                            f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                        ]
                    }
                    
                    summary_text = '\n'.join(summary_data['Processing Summary'])
                    
                    st.download_button(
                        label="📋 Download Summary",
                        data=summary_text,
                        file_name="processing_summary.txt",
                        mime="text/plain",
                        use_container_width=True,
                        key="download_summary_btn"
                    )
        
        # Backup History
        if st.session_state.backup_history:
            with st.expander("📚 Processing History", expanded=False):
                for i, backup in enumerate(reversed(st.session_state.backup_history)):
                    st.write(f"**{i+1}.** {backup['timestamp']} - {backup['mode']}")
                    st.write(f"   Files: {backup['files_modified']}, Replacements: {backup['total_replacements']}")
                    st.write(f"   Location: `{backup['backup_dir']}`")
                    st.write("---")
    
    with col2:
        # Console section
        st.markdown("""
        <div class="modern-card">
            <div class="card-title">📋 Console Output <div class="status-indicator"></div></div>
        </div>
        """, unsafe_allow_html=True)
        
        # Console display
        console_text = '\n'.join(st.session_state.console_messages[-20:])  # Show last 20 messages
        st.markdown(
            f'<div class="console-area">{console_text}</div>',
            unsafe_allow_html=True
        )
        
        # System information
        st.markdown("### 💻 System Status")
        
        status_info = {
            "🐍 Python Version": f"{os.sys.version_info.major}.{os.sys.version_info.minor}.{os.sys.version_info.micro}",
            "⏰ Current Time": datetime.now().strftime('%H:%M:%S'),
            "📊 Console Lines": len(st.session_state.console_messages),
            "🔄 Patterns Loaded": len(st.session_state.replacement_map),
            "📄 Files Loaded": len(st.session_state.loaded_files),
            "📁 Temp Directories": len(st.session_state.temp_directories)
        }
        
        for label, value in status_info.items():
            col_info1, col_info2 = st.columns([2, 1])
            with col_info1:
                st.caption(label)
            with col_info2:
                st.caption(f"**{value}**")
        
        # Quick Actions
        st.markdown("### ⚡ Quick Actions")
        
        if st.button("🔄 Reload Data", use_container_width=True, key="reload_data_btn"):
            # Force a rerun to refresh all data
            st.rerun()
        
        if st.button("📁 Open Temp Folder", use_container_width=True, key="open_temp_btn"):
            if st.session_state.temp_directories:
                temp_dir = st.session_state.temp_directories[-1]
                if os.path.exists(temp_dir):
                    st.info(f"Temp folder: {temp_dir}")
                else:
                    st.warning("Temp folder no longer exists")
            else:
                st.info("No temp directories created")
        
        if st.session_state.results and st.session_state.results.get('output_dir'):
            if st.button("📂 Open Output Folder", use_container_width=True, key="open_output_btn"):
                output_dir = st.session_state.results['output_dir']
                if os.path.exists(output_dir):
                    st.info(f"Output folder: {output_dir}")
                else:
                    st.warning("Output folder no longer exists")
    
    # Footer
    st.markdown("---")
    st.markdown(
        '<div style="text-align: center; color: var(--text-muted); font-size: 0.875rem; padding: 1rem;">'
        '© 2025 Hrishik Kunduru • DocXReplace v3.0 Professional • All Rights Reserved'
        '</div>',
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
