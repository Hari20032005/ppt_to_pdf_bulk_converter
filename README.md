# PowerPoint to PDF Bulk Converter

This script converts PowerPoint files (.ppt/.pptx) to PDF format in bulk.

## Prerequisites

### On Windows
- Microsoft PowerPoint must be installed
- Python comtypes library

### On macOS/Linux
- LibreOffice must be installed
  - On macOS: `brew install --cask libreoffice`
  - On Linux: `sudo apt-get install libreoffice`

## Installation

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Install LibreOffice if on macOS or Linux:
   ```bash
   # On macOS
   brew install --cask libreoffice
   
   # On Linux
   sudo apt-get install libreoffice
   ```

## Usage

### Basic usage
```bash
python ppt_to_pdf_converter.py <input_folder> [output_folder]
```

### Convert with recursive search
```bash
python ppt_to_pdf_converter.py <input_folder> [output_folder] --recursive
```

### Examples

Convert all PPT files in the 'presentations' folder:
```bash
python ppt_to_pdf_converter.py ./presentations
```

Convert all PPT files and save PDFs to a specific folder:
```bash
python ppt_to_pdf_converter.py ./presentations ./pdfs
```

Convert all PPT files including subfolders:
```bash
python ppt_to_pdf_converter.py ./presentations ./pdfs --recursive
```

## Features

- Bulk conversion of multiple PowerPoint files
- Support for both .ppt and .pptx files
- Cross-platform compatibility (Windows, macOS, Linux)
- Recursive folder processing option
- Progress tracking and error handling
- Automatic folder creation for output files