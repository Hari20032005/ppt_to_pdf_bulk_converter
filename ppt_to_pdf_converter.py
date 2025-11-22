import os
import sys
import glob
from pathlib import Path

def convert_ppt_to_pdf(input_path, output_path=None):
    """
    Convert a PowerPoint file (.ppt/.pptx) to PDF
    This function works differently based on the operating system:
    - On Windows: Uses COM interface to PowerPoint
    - On macOS: Uses LibreOffice (if available)
    - On Linux: Uses LibreOffice (if available)
    
    Args:
        input_path (str): Path to the input PowerPoint file
        output_path (str): Path to the output PDF file (optional)
    
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    import platform
    system = platform.system()
    
    # If no output path provided, create one based on input path
    if output_path is None:
        output_path = str(Path(input_path).with_suffix('.pdf'))
    
    try:
        if system == "Windows":
            return convert_ppt_to_pdf_windows(input_path, output_path)
        else:
            return convert_ppt_to_pdf_libreoffice(input_path, output_path)
    except Exception as e:
        print(f"Error converting {input_path}: {str(e)}")
        return False

def convert_ppt_to_pdf_windows(input_path, output_path):
    """Convert PPT to PDF on Windows using COM interface"""
    import comtypes.client
    
    # Create PowerPoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = True  # Make PowerPoint visible
    
    # Open the PowerPoint file
    presentation = powerpoint.Presentations.Open(input_path)
    
    # Export as PDF (format type 32 = PDF format)
    presentation.Export(output_path, "PDF")
    
    # Close the presentation
    presentation.Close()
    
    # Quit PowerPoint application
    powerpoint.Quit()
    
    print(f"Successfully converted: {input_path} -> {output_path}")
    return True

def convert_ppt_to_pdf_libreoffice(input_path, output_path):
    """Convert PPT to PDF using LibreOffice (cross-platform)"""
    import subprocess
    
    # Check if LibreOffice is installed
    try:
        subprocess.run(['soffice', '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except FileNotFoundError:
        print("LibreOffice is not installed or not in PATH. Please install LibreOffice to convert PPT to PDF on macOS/Linux.")
        print("On macOS: brew install --cask libreoffice")
        print("On Linux: sudo apt-get install libreoffice")
        return False
    
    # Create directory for output if needed
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # Run LibreOffice in headless mode to convert to PDF
    cmd = [
        'soffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', os.path.dirname(output_path),
        input_path
    ]
    
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    
    if result.returncode == 0:
        # LibreOffice creates the PDF in the same directory as the input file with the same name
        # We need to move it to the desired location if different
        input_dir = os.path.dirname(input_path)
        expected_output = os.path.join(input_dir, Path(input_path).stem + '.pdf')
        
        if expected_output != output_path:
            if os.path.exists(expected_output):
                os.rename(expected_output, output_path)
        
        print(f"Successfully converted: {input_path} -> {output_path}")
        return True
    else:
        print(f"Error converting with LibreOffice: {result.stderr.decode()}")
        return False

def convert_bulk_ppt_to_pdf(input_folder, output_folder=None, recursive=False):
    """
    Convert all PowerPoint files in a folder to PDF
    
    Args:
        input_folder (str): Path to the folder containing PowerPoint files
        output_folder (str): Path to the folder for output PDF files (optional)
        recursive (bool): Whether to process subfolders recursively
    
    Returns:
        list: List of successfully converted files
    """
    # If no output folder provided, create one with '_pdf' suffix
    if output_folder is None:
        output_folder = input_folder + "_pdf"
    
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    # Find all PowerPoint files
    ppt_extensions = ['*.ppt', '*.pptx']
    ppt_files = []
    
    for ext in ppt_extensions:
        if recursive:
            ppt_files.extend(glob.glob(os.path.join(input_folder, '**', ext), recursive=True))
        else:
            ppt_files.extend(glob.glob(os.path.join(input_folder, ext)))
    
    if not ppt_files:
        print(f"No PowerPoint files found in {input_folder}")
        return []
    
    print(f"Found {len(ppt_files)} PowerPoint files to convert")
    
    successful_conversions = []
    
    for ppt_file in ppt_files:
        try:
            # Get the filename without extension
            filename = Path(ppt_file).stem
            output_pdf_path = os.path.join(output_folder, f"{filename}.pdf")
            
            print(f"Converting: {ppt_file}")
            
            # Convert the file
            success = convert_ppt_to_pdf(ppt_file, output_pdf_path)
            
            if success:
                successful_conversions.append((ppt_file, output_pdf_path))
            
        except Exception as e:
            print(f"Error processing {ppt_file}: {str(e)}")
    
    print(f"\nConversion completed! Successfully converted {len(successful_conversions)} out of {len(ppt_files)} files.")
    return successful_conversions

def main():
    if len(sys.argv) < 2:
        print("Usage: python ppt_to_pdf_converter.py <input_folder> [output_folder] [--recursive]")
        print("Example: python ppt_to_pdf_converter.py ./powerpoints ./pdfs --recursive")
        return
    
    input_folder = sys.argv[1]
    output_folder = sys.argv[2] if len(sys.argv) > 2 else None
    recursive = "--recursive" in sys.argv or "-r" in sys.argv
    
    if not os.path.isdir(input_folder):
        print(f"Error: {input_folder} is not a valid directory")
        return
    
    convert_bulk_ppt_to_pdf(input_folder, output_folder, recursive)

if __name__ == "__main__":
    main()