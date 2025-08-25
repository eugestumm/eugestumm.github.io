#!/usr/bin/env python3
"""
CV Generator Script for Academic Applications
Converts Markdown CV to PDF using Pandoc with LaTeX template
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

def check_dependencies():
    """Check if required dependencies are installed"""
    dependencies = {
        'pandoc': 'Pandoc is required for document conversion',
        'xelatex': 'XeLaTeX is required for PDF generation'
    }
    
    missing = []
    for cmd, description in dependencies.items():
        if shutil.which(cmd) is None:
            missing.append(f"  - {cmd}: {description}")
    
    if missing:
        print("❌ Missing dependencies:")
        for dep in missing:
            print(dep)
        print("\nPlease install the missing dependencies and try again.")
        return False
    
    print("✅ All dependencies found")
    return True

def ensure_directory_exists(directory):
    """Create directory if it doesn't exist"""
    Path(directory).mkdir(parents=True, exist_ok=True)
    print(f"📁 Directory ensured: {directory}")

def copy_cv_to_assets():
    """Copy cv_euge_stumm.pdf to assets directory"""
    print("\n📋 Copying cv_euge_stumm.pdf to assets directory...")
    cv_pdf = 'cv_euge_stumm.pdf'
    assets_dir = Path.cwd().parent / 'assets'
    
    # Create assets directory if it doesn't exist
    ensure_directory_exists(assets_dir)
    
    if Path(cv_pdf).exists():
        dest_path = assets_dir / cv_pdf
        try:
            shutil.copy2(cv_pdf, dest_path)
            print(f"✅ Successfully copied {cv_pdf} to assets directory")
            return True
        except Exception as e:
            print(f"❌ Error copying {cv_pdf} to assets: {e}")
            return False
    else:
        print(f"❌ Error: {cv_pdf} not found - cannot copy to assets")
        return False

def escape_latex_chars(text):
    """Escape special LaTeX characters"""
    if not isinstance(text, str):
        return str(text)
    
    # Don't escape URLs, LaTeX commands, or markdown links
    if (text.startswith('http') or 
        text.startswith('\\') or 
        '[' in text and '](' in text or
        text.startswith('*') or
        text.startswith('#')):
        return text
    
    # Dictionary of characters to escape
    latex_chars = {
        '&': r'\&',
        '%': r'\%',
        '$': r'\$',
        '#': r'\#',
        '^': r'\textasciicircum{}',
        '_': r'\_',
        '{': r'\{',
        '}': r'\}',
        '~': r'\textasciitilde{}',
        '\\': r'\textbackslash{}'
    }
    
    result = text
    for char, escape in latex_chars.items():
        result = result.replace(char, escape)
    
    return result

def preprocess_markdown(input_file, output_file):
    """Preprocess the markdown to make it more LaTeX-friendly"""
    print(f"🔄 Preprocessing {input_file}...")
    
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Extract YAML front matter variables
    yaml_vars = {}
    if content.startswith('---'):
        parts = content.split('---', 2)
        if len(parts) >= 3:
            yaml_content = parts[1].strip()
            content = parts[2].strip()
            
            # Parse basic YAML variables
            for line in yaml_content.split('\n'):
                if ':' in line and not line.strip().startswith('#'):
                    key, value = line.split(':', 1)
                    yaml_vars[key.strip()] = value.strip().strip('"').strip("'")
    
    # Process content for better LaTeX formatting
    lines = content.split('\n')
    processed_lines = []
    in_pub_list = False
    in_teaching_section = False
    current_institution = None
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Skip empty lines but preserve some spacing
        if not line.strip():
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            processed_lines.append("")
            i += 1
            continue
        
        # Handle section headers
        if line.startswith('# '):
            title = line[2:].strip()
            # Skip the main title (first h1)
            if 'Euge Helyantus Stumm' not in title:
                processed_lines.append(f"\\section{{{escape_latex_chars(title)}}}")
                in_teaching_section = ('Teaching' in title)
        elif line.startswith('## '):
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            title = line[3:].strip()
            processed_lines.append(f"\\section{{{escape_latex_chars(title)}}}")
            in_teaching_section = ('Teaching' in title)
        elif line.startswith('### '):
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            title = line[4:].strip()
            processed_lines.append(f"\\subsection{{{escape_latex_chars(title)}}}")
        elif line.startswith('#### '):
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            title = line[5:].strip()
            # In teaching section, these are institution names
            if in_teaching_section:
                current_institution = title
                processed_lines.append(f"\\subsection{{{escape_latex_chars(title)}}}")
            else:
                processed_lines.append(f"\\subsubsection{{{escape_latex_chars(title)}}}")
        
        # Handle numbered publication lists
        elif line.strip() and len(line) > 2 and line.strip()[0].isdigit() and '. ' in line[:5]:
            if not in_pub_list:
                processed_lines.append("\\begin{publicationlist}")
                in_pub_list = True
            # Remove the number and period, let LaTeX handle numbering
            content_part = line.strip()
            dot_index = content_part.find('. ')
            if dot_index != -1:
                content_part = content_part[dot_index + 2:]
            processed_lines.append(f"\\item {content_part}")
        
        # Handle course/position entries with em dash or hyphen
        elif line.startswith('**') and '**' in line[2:] and ('—' in line or ' - ' in line):
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            # Split on em dash or regular dash
            separator = '—' if '—' in line else ' - '
            parts = line.split(separator, 1)
            if len(parts) == 2:
                title = parts[0].strip().replace('**', '')
                date = parts[1].strip()
                processed_lines.append(f"\\CVitem{{{escape_latex_chars(title)}}}{{{escape_latex_chars(date)}}}")
                
                # Look ahead for additional info (co-taught, supervised, etc.)
                j = i + 1
                while j < len(lines) and lines[j].strip() and not lines[j].startswith('**') and not lines[j].startswith('#'):
                    next_line = lines[j].strip()
                    if next_line.startswith('*') and next_line.endswith('*'):
                        # Italic line (like co-taught info)
                        processed_lines.append(f"\\hspace{{1em}}\\textit{{{escape_latex_chars(next_line[1:-1])}}}")
                    else:
                        processed_lines.append(f"\\hspace{{1em}}{escape_latex_chars(next_line)}")
                    j += 1
                i = j - 1  # Adjust index to skip processed lines
            else:
                processed_lines.append(escape_latex_chars(line))
        
        # Handle degree entries (for Education section)
        elif line.startswith('**') and line.count('**') >= 2 and ('Ph.D.' in line or 'M.A.' in line or 'B.A.' in line):
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            degree = line.replace('**', '').strip()
            
            # Look ahead for university, advisor, thesis info
            university = ""
            advisor_info = ""
            thesis_info = ""
            additional_info = []
            
            j = i + 1
            while j < len(lines) and lines[j].strip() and not lines[j].startswith('**') and not lines[j].startswith('#'):
                next_line = lines[j].strip()
                if 'University' in next_line or 'Federal' in next_line:
                    university = next_line
                elif next_line.startswith('Advisor:'):
                    advisor_info = next_line
                elif next_line.startswith('Thesis:'):
                    thesis_info = next_line
                elif next_line.startswith('Graduate Certificate'):
                    additional_info.append(next_line)
                j += 1
            
            processed_lines.append(f"\\CVdegree{{{escape_latex_chars(degree)}}}{{}}{{{escape_latex_chars(university)}}}{{}}")
            if thesis_info:
                processed_lines.append(f"\\thesis{{{escape_latex_chars(thesis_info[8:])}}}")
            if advisor_info:
                processed_lines.append(f"\\advisor{{{escape_latex_chars(advisor_info[9:])}}}")
            for info in additional_info:
                processed_lines.append(f"{escape_latex_chars(info)}\\\\")
            
            i = j - 1  # Skip processed lines
        
        # Handle regular bold text for positions/awards
        elif line.startswith('**') and line.count('**') >= 2:
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            bold_text = line.replace('**', '').strip()
            
            # Check if this is followed by additional info
            if i + 1 < len(lines) and lines[i + 1].strip() and not lines[i + 1].startswith('**') and not lines[i + 1].startswith('#'):
                next_line = lines[i + 1].strip()
                # If next line looks like a date or amount, format as CVitem
                if any(word in next_line.lower() for word in ['amount:', '(202', '— 202', 'fall', 'spring', 'summer']):
                    processed_lines.append(f"\\CVitem{{{escape_latex_chars(bold_text)}}}{{}}")
                else:
                    processed_lines.append(f"\\textbf{{{escape_latex_chars(bold_text)}}}")
            else:
                processed_lines.append(f"\\textbf{{{escape_latex_chars(bold_text)}}}")
        
        # Handle italic text
        elif line.strip().startswith('*') and line.strip().endswith('*') and line.count('*') == 2:
            if in_pub_list:
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            italic_text = line.strip()[1:-1]  # Remove * markers
            processed_lines.append(f"\\textit{{{escape_latex_chars(italic_text)}}}")
        
        # Handle regular lines
        else:
            # End publication list if we hit a non-list item
            if in_pub_list and not (line.strip() and len(line) > 2 and line.strip()[0].isdigit() and '. ' in line[:5]):
                processed_lines.append("\\end{publicationlist}")
                in_pub_list = False
            
            processed_lines.append(escape_latex_chars(line))
        
        i += 1

    # Close any open lists
    if in_pub_list:
        processed_lines.append("\\end{publicationlist}")
    
    # Create YAML front matter with extracted variables
    yaml_header = "---\n"
    yaml_header += "author: Euge Helyantus Stumm\n"
    yaml_header += "subtitle: Ph.D. Candidate in Literary, Cultural, and Linguistic Studies\n"
    yaml_header += "email: ehs89@miami.edu\n"
    yaml_header += "website: https://eugestumm.github.io\n"
    yaml_header += "orcid: https://orcid.org/0000-0001-9087-4198\n"
    yaml_header += "github: https://github.com/eugestumm\n"
    
    # Add any other YAML variables that were in the original
    for key, value in yaml_vars.items():
        if key not in ['author', 'subtitle', 'email', 'website', 'orcid', 'github']:
            yaml_header += f"{key}: {value}\n"
    
    yaml_header += "---\n\n"
    
    final_content = yaml_header + '\n'.join(processed_lines)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_content)
    
    print(f"✅ Preprocessed markdown saved to {output_file}")

def convert_cv_to_pdf():
    """Convert cv.md to cv_euge_stumm.pdf using Pandoc with LaTeX template"""
    print("\n📄 Converting CV to PDF using Pandoc (LaTeX template)...")
    
    cv_md = 'cv.md'
    processed_md = 'cv_processed.md'
    cv_pdf = 'cv_euge_stumm.pdf'
    template_file = Path('pandoc') / 'harvard_cv.tex'
    
    # Check if input files exist
    if not Path(cv_md).exists():
        print(f"❌ Error: {cv_md} not found")
        return False
    
    if not template_file.exists():
        print(f"❌ Error: LaTeX template not found at {template_file}")
        print("Please ensure the template file exists in the pandoc/ directory")
        return False
    
    try:
        # Preprocess the markdown
        preprocess_markdown(cv_md, processed_md)
        
        # First attempt with XeLaTeX and template
        pandoc_cmd = [
            'pandoc', processed_md,
            '-o', cv_pdf,
            '--from', 'markdown',
            '--template', str(template_file),
            '--pdf-engine=xelatex',
            '--variable', 'geometry:margin=1in',
            '--variable', 'fontsize=11pt',
            '--variable', 'papersize=letter',
            '--variable', 'date:August 2025'
        ]
        
        print(f"🔄 Running: {' '.join(pandoc_cmd)}")
        result = subprocess.run(pandoc_cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0 and Path(cv_pdf).exists():
            print("✅ PDF conversion completed successfully with XeLaTeX")
            # Clean up processed markdown
            if Path(processed_md).exists():
                Path(processed_md).unlink()
            return True
        
        # If XeLaTeX failed, try with pdflatex
        print("⚠️  XeLaTeX failed, trying with pdflatex...")
        
        # Clean up any partial files
        if Path(cv_pdf).exists():
            Path(cv_pdf).unlink()
        
        pandoc_cmd[pandoc_cmd.index('--pdf-engine=xelatex')] = '--pdf-engine=pdflatex'
        
        print(f"🔄 Running: {' '.join(pandoc_cmd)}")
        result = subprocess.run(pandoc_cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0 and Path(cv_pdf).exists():
            print("✅ PDF conversion completed successfully with pdflatex")
            # Clean up processed markdown
            if Path(processed_md).exists():
                Path(processed_md).unlink()
            return True
        
        # If both fail, try without template (basic conversion)
        print("⚠️  Template conversion failed, trying basic conversion...")
        
        # Clean up any partial files
        if Path(cv_pdf).exists():
            Path(cv_pdf).unlink()
        
        basic_cmd = [
            'pandoc', cv_md,  # Use original markdown
            '-o', cv_pdf,
            '--from', 'markdown',
            '--pdf-engine=xelatex',
            '--variable', 'geometry:margin=1in',
            '--variable', 'fontsize=11pt',
            '--variable', 'papersize=letter'
        ]
        
        print(f"🔄 Running: {' '.join(basic_cmd)}")
        result = subprocess.run(basic_cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0 and Path(cv_pdf).exists():
            print("✅ PDF conversion completed with basic formatting")
            return True
        
        # All attempts failed
        print("❌ All PDF conversion attempts failed")
        print(f"Last error: {result.stderr}")
        if result.stdout:
            print(f"Output: {result.stdout}")
        return False
            
    except subprocess.TimeoutExpired:
        print("❌ PDF conversion timed out")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False
    finally:
        # Clean up processed markdown if it exists
        if Path(processed_md).exists():
            try:
                Path(processed_md).unlink()
                print("🧹 Cleaned up temporary files")
            except:
                pass

def copy_to_pages():
    """Copy files to _pages directory if it exists"""
    pages_dir = Path.cwd().parent / '_pages'
    if pages_dir.exists():
        print(f"\n📋 Copying files to {pages_dir}...")
        try:
            # Copy markdown file
            if Path('cv.md').exists():
                shutil.copy2('cv.md', pages_dir / 'cv.md')
                print("✅ Copied cv.md to _pages")
            
            # Copy PDF if it exists
            if Path('cv_euge_stumm.pdf').exists():
                shutil.copy2('cv_euge_stumm.pdf', pages_dir / 'cv_euge_stumm.pdf')
                print("✅ Copied cv_euge_stumm.pdf to _pages")
                
        except Exception as e:
            print(f"❌ Error copying to _pages: {e}")
    else:
        print("📁 _pages directory not found, skipping copy")

def main():
    """Main function to orchestrate CV generation"""
    print("🎓 Academic CV Generator")
    print("=" * 40)
    
    # Check dependencies
    if not check_dependencies():
        sys.exit(1)
    
    # Ensure pandoc directory exists
    ensure_directory_exists('pandoc')
    
    # Convert cv.md to cv_euge_stumm.pdf
    if convert_cv_to_pdf():
        print("\n✅ CV conversion successful!")
        
        # Copy to assets directory
        copy_cv_to_assets()
        
        # Copy to _pages directory if it exists
        copy_to_pages()
        
        print("\n🎉 All tasks completed successfully!")
        print("📄 Your CV is ready: cv_euge_stumm.pdf")
    else:
        print("\n❌ CV conversion failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()