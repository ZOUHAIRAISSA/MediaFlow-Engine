# MediaFlow Engine 🎬📊
## E-commerce Media Management Platform

**A comprehensive solution for e-commerce businesses to automate media processing, organize product photos, and manage Google Drive workflows with CSV-driven batch operations.**

---

## 🎯 Project Overview

This project is designed for **e-commerce businesses** (specifically Etsy sellers) to streamline their product media management workflow. It combines automated image/video processing with Google Drive integration to create an efficient pipeline for handling product photos and videos at scale.

## 🌟 Core Functionality

### 🛍️ **E-commerce Workflow Automation**
- **CSV-Driven Processing**: Process product folders based on spreadsheet data (SKU mapping, titles, tags)
- **Product Photo Organization**: Automatic folder creation and naming based on SKU codes
- **Batch Media Processing**: Handle hundreds of product images/videos simultaneously
- **Etsy Integration**: Designed specifically for Etsy seller workflows

### 🎥 **Media Processing Pipeline**
- **HEIC to JPG Conversion**: Convert iPhone photos to web-friendly JPG format
- **Video Optimization**: FFmpeg-powered video encoding for web delivery
- **Metadata Management**: Automatic title, tags, and rating assignment from CSV data
- **Quality Preservation**: Maintain image quality while optimizing file sizes

### ☁️ **Google Drive Integration**
- **Automated Upload**: Upload processed product folders to Google Drive
- **Google Sheets Sync**: Update product spreadsheets with Drive folder URLs
- **CSV Download**: Download product folders from Drive based on CSV links
- **Folder Management**: Organize product photos in structured folder hierarchies

### 🔧 **Business Tools**
- **Modern Interface**: User-friendly GUI for non-technical users
- **Progress Monitoring**: Real-time tracking of batch operations
- **Error Handling**: Comprehensive logging and recovery mechanisms
- **Folder Merging**: Combine product folders from different sources

## 🛠️ Technical Implementation

### **Core Technologies**
- **Python 3.8+**: Main development language
- **FFmpeg**: Video processing and encoding
- **ExifTool**: Metadata extraction and manipulation
- **Pillow**: Image processing and HEIC conversion
- **Google APIs**: Drive and Sheets integration

### **Key Components**
- **CSV Parser**: Reads product data from spreadsheets
- **Media Processor**: Handles HEIC/JPG conversion and video encoding
- **Drive Manager**: Manages Google Drive uploads/downloads
- **Sheets Sync**: Updates product spreadsheets with Drive URLs
- **GUI Interface**: Modern Tkinter-based user interface

## 🚀 Real-World Use Cases

### **Etsy Sellers**
- Process iPhone photos (HEIC) to JPG for product listings
- Organize product photos by SKU codes
- Upload product folders to Google Drive automatically
- Update product spreadsheets with Drive folder links

### **E-commerce Businesses**
- Batch process product catalogs
- Standardize image formats and quality
- Automate media workflow from photoshoot to listing
- Manage product media across multiple platforms

### **Product Photographers**
- Convert client photos from HEIC to JPG
- Apply consistent metadata and organization
- Deliver processed photos via Google Drive
- Maintain organized client folders

## 📦 Installation

```bash
# Clone the repository


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



## 🔄 Typical Workflow

### **1. Product Photo Processing**
1. Take photos with iPhone (HEIC format)
2. Organize photos in folders by product SKU
3. Run batch processor to convert HEIC → JPG
4. Apply metadata (titles, tags) from CSV data

### **2. Google Drive Integration**
1. Upload processed product folders to Google Drive
2. Update Google Sheets with Drive folder URLs
3. Download product folders based on CSV links
4. Sync product data between local and cloud storage

### **3. E-commerce Listing**
1. Use processed JPG images for product listings
2. Reference Google Drive folders for additional photos
3. Maintain organized product catalog
4. Automate media workflow for new products

## 🎯 Business Benefits

- **Time Savings**: 80% reduction in manual photo processing time
- **Consistency**: Standardized image formats and quality
- **Organization**: Automatic folder structure based on SKU codes
- **Scalability**: Handle hundreds of products simultaneously
- **Integration**: Seamless Google Drive and Sheets workflow

