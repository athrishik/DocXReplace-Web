# README.md
# DocXReplace v3.0 - Professional Document Replacement Tool

🔄 **Modern web-based document replacement with legal token migration support**

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://DocXReplace-Web.streamlit.app)

## 🌟 Features

- **Professional Document Processing**: Handle multiple .docx files with advanced replacement patterns
- **Legal Token Migration**: Specialized support for legal document token transformation
- **Multiple Input Methods**: Load from folders, Excel reports, or ZIP archives
- **Regex Support**: Advanced pattern matching with capture groups
- **Safe Processing**: Dry run preview and backup creation options
- **Modern Interface**: Dark theme with glassmorphism design
- **Real-time Progress**: Live console output and progress tracking

## 🚀 Quick Start

### Online Access
Visit the live app: [DocXReplace v3.0](https://DocXReplace-Web.streamlit.app)

### Local Installation
```bash
git clone https://github.com/athrishik/docxreplace-web.git
cd docxreplace-streamlit
pip install -r requirements.txt
streamlit run app.py
```

## 📖 Usage Guide

### 1. Load Files
- **Folder**: Browse and select a folder containing .docx files
- **Excel**: Upload DocXScan Excel reports with file paths
- **ZIP**: Upload archives containing document files

### 2. Configure Replacements
- Upload JSON replacement files
- Enable regex mode for advanced patterns
- Use built-in templates for common legal tokens

### 3. Process Documents
- **Dry Run**: Preview changes without modification
- **Modified Copies**: Create new files (originals untouched)
- **In-place**: Modify original files directly

### 4. Download Results
- Download ZIP of processed files
- Export processing summaries
- View detailed operation logs

## 🔧 Replacement Patterns

### Standard Tokens
```json
{
  "<<FileService.": "<<NewFileService.",
  "</ff>": "<<PAGE_BREAK>>",
  "<bold>": "<<BOLD>>",
  "[[MCOMPUTEINTO(<<": "<<MCOMPUTE_INTO("
}
```

### Regex Patterns
```json
{
  "<<FileService\\.\\w*": "<<NewFileService.{{match}}",
  "<(\\w+)>": "<<{{match}}>>",
  "\\[\\[(\\w+)COMPUTEINTO\\(": "<<{{match}}_INTO("
}
```
## 📜 License

Copyright © 2025 Hrishik Kunduru. All rights reserved.

This project is proprietary software. Unauthorized copying, distribution, or modification is strictly prohibited.

## 🆘 Support

- **Documentation**: [Document Tools Suite](https://github.com/athrishik/docxsuite/readme.md)
- **Contact**: hrishik.kunduru@gmail.com
  
## 🔗 Related Projects

- [DocXScan v3.0](https://github.com/athrishik/docxscan-web) - Professional document scanner - web version
- [Document Tools Suite](https://github.com/athrishik/docxsuite) - Complete document suite toolkit
