#!/usr/bin/env python3
"""
Pandoc Converter Script
Converts markdown files to PDF using pandoc with custom CSS

Standalone usage:
  ./pandoc_converter.py
  ./pandoc_converter.py -i custom.md -o custom.pdf -c styles/custom.css
  python3 pandoc_converter.py --help
"""

import subprocess
import os
import sys
import argparse

__version__ = "1.0.0"

# Default values
DEFAULT_INPUT = "cv.md"
DEFAULT_OUTPUT = "cv_euge_stumm.pdf"
DEFAULT_CSS = os.path.join("pandoc", "harvard_cv.css")

def check_dependency(command, name):
    """Check if a dependency is installed"""
    try:
        subprocess.run([command, '--version'], capture_output=True, check=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print(f"Warning: {name} ({command}) not found")
        return False

def get_available_engines():
    """Get list of available PDF engines"""
    available_engines = []
    engines_to_check = [
        ('wkhtmltopdf', 'wkhtmltopdf'),
        ('weasyprint', 'weasyprint'),
        ('prince', 'prince'),
        ('pdflatex', 'LaTeX pdflatex'),
        ('xelatex', 'LaTeX xelatex')
    ]
    
    for engine_cmd, engine_name in engines_to_check:
        if check_dependency(engine_cmd, engine_name):
            available_engines.append(engine_cmd)
    
    return available_engines

def convert_md_to_pdf(md_file, pdf_file, css_file, engine=None):
    """Convert markdown file to PDF using pandoc"""
    
    # Check if pandoc is installed
    if not check_dependency('pandoc', 'pandoc'):
        print("Error: pandoc is required but not installed or not found in PATH")
        print("Install pandoc from: https://pandoc.org/installing.html")
        return False
    
    # Check if input files exist
    if not os.path.exists(md_file):
        print(f"Error: Input file '{md_file}' not found")
        return False
    
    if not os.path.exists(css_file):
        print(f"Error: CSS file '{css_file}' not found")
        return False
    
    # Build output directory if it doesn't exist
    output_dir = os.path.dirname(pdf_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # If specific engine requested, check if it's available
    if engine:
        if not check_dependency(engine, f"PDF engine '{engine}'"):
            print(f"Error: Requested engine '{engine}' is not available")
            return False
        available_engines = [engine]
    else:
        # Get available engines in order of preference
        available_engines = get_available_engines()
    
    if not available_engines:
        print("Error: No PDF engines found. Please install one of:")
        print("  - wkhtmltopdf (recommended): sudo apt-get install wkhtmltopdf")
        print("  - weasyprint: pip install weasyprint")
        print("  - prince: https://www.princexml.com/download/")
        print("  - LaTeX: sudo apt-get install texlive-latex-base")
        return False
    
    # Try available engines
    for engine in available_engines:
        try:
            print(f"Using {engine} as PDF engine...")
            
            pandoc_cmd = [
                'pandoc',
                md_file,
                '-o', pdf_file,
                '--css', css_file,
                f'--pdf-engine={engine}',
                '--variable', 'papersize=a4',
                '--standalone'
            ]
            
            # Add engine-specific options
            if engine in ['pdflatex', 'xelatex']:
                pandoc_cmd.extend(['--variable', 'geometry:margin=1in'])
            
            result = subprocess.run(pandoc_cmd, capture_output=True, text=True, check=True)
            print(f"✓ Successfully converted '{md_file}' to '{pdf_file}' using {engine}")
            return True
            
        except subprocess.CalledProcessError as e:
            print(f"✗ Conversion failed with {engine}: {e}")
            if e.stderr:
                print(f"   Error details: {e.stderr.strip()}")
            continue
    
    print("Error: All available PDF engines failed to convert the file")
    return False

def main():
    parser = argparse.ArgumentParser(
        description="Convert markdown files to PDF using pandoc with custom CSS",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s                       # Uses defaults: cv.md -> cv.pdf with pandoc/harvard_cv.css
  %(prog)s -i resume.md -o resume.pdf -c styles/resume.css
  %(prog)s -e wkhtmltopdf        # Use specific engine with defaults
  %(prog)s --version
        
Available PDF engines (auto-detected):
  wkhtmltopdf, weasyprint, prince, pdflatex, xelatex
        """
    )
    
    parser.add_argument('-i', '--input', default=DEFAULT_INPUT,
                       help=f'Input markdown file (default: {DEFAULT_INPUT})')
    parser.add_argument('-o', '--output', default=DEFAULT_OUTPUT,
                       help=f'Output PDF file (default: {DEFAULT_OUTPUT})')
    parser.add_argument('-c', '--css', default=DEFAULT_CSS,
                       help=f'CSS stylesheet file (default: {DEFAULT_CSS})')
    parser.add_argument('-e', '--engine', help='Specify PDF engine to use')
    parser.add_argument('-v', '--version', action='version', 
                       version=f'%(prog)s {__version__}')
    
    args = parser.parse_args()
    
    # Show default values when no arguments are provided
    if len(sys.argv) == 1:
        print(f"Using default values:")
        print(f"  Input:  {args.input}")
        print(f"  Output: {args.output}")
        print(f"  CSS:    {args.css}")
        print()
    
    # Perform conversion
    success = convert_md_to_pdf(args.input, args.output, args.css, args.engine)
    
    if success:
        print(f"\n✓ Conversion completed successfully!")
        print(f"  Input:  {args.input}")
        print(f"  Output: {args.output}")
        print(f"  CSS:    {args.css}")
        
        # Show file size if successful
        if os.path.exists(args.output):
            file_size = os.path.getsize(args.output)
            print(f"  Size:   {file_size / 1024:.1f} KB")
    else:
        print(f"\n✗ Conversion failed!")
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()