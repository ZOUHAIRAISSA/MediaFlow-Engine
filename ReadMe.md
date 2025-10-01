# MediaFlow Engine 🎬📊

**An intelligent media processing platform that combines batch video/image processing with Google Drive integration and automated metadata management.**

## 🌟 Key Features

### 🎥 **Advanced Media Processing**
- **Batch Video Processing**: FFmpeg-powered encoding with customizable parameters (CRF, presets, codecs)
- **Image Optimization**: HEIC to JPG conversion with quality preservation
- **Smart Metadata Management**: Automated title, tags, and rating assignment
- **Multi-format Support**: MP4, MOV, M4V, MKV, AVI, WMV, FLV, WebM, HEIC, JPEG

### ☁️ **Google Cloud Integration**
- **Google Drive API**: Automated folder upload/download with progress tracking
- **Google Sheets Integration**: Real-time data synchronization and status updates
- **CSV-based Workflows**: Bulk operations with spreadsheet-driven automation
- **OAuth2 Authentication**: Secure Google services access

### 🔧 **Professional Tools**
- **Modern GUI**: Dark-themed interface with real-time progress monitoring
- **Folder Management**: Intelligent merging and organization tools
- **Error Handling**: Comprehensive logging and error recovery
- **Cross-platform**: Windows executable with portable dependencies

## 🛠️ Technical Stack

- **Backend**: Python 3.8+, FFmpeg, ExifTool
- **APIs**: Google Drive API v3, Google Sheets API v4
- **GUI**: Tkinter with modern styling
- **Media Processing**: FFmpeg, Pillow, Mutagen
- **Authentication**: OAuth2, Google Credentials

## 🚀 Use Cases

- **Content Creators**: Batch process video/image libraries
- **E-commerce**: Automated product media optimization
- **Data Engineers**: Media pipeline automation
- **Marketing Teams**: Bulk content preparation for platforms

## 📦 Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/mediaflow-engine.git

# Install dependencies
pip install -r requirements.txt

# Setup Google credentials
# Place credentials.json in project root

# Run the application
python dirvecopy.py
```

## 🔑 Prerequisites

- Python 3.8+
- Google Cloud Project with Drive & Sheets APIs enabled
- FFmpeg (included in dist/)
- ExifTool (included in dist/)

## 📊 Project Structure
