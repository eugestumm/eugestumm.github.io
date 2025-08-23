import subprocess
import os
import shutil

def run_updater(script_name):
    print(f"\nRunning {script_name}...")
    subprocess.run(['python3', script_name])

def show_file_content(file_name):
    print(f"\nContent of {file_name}:")
    try:
        with open(file_name, 'r') as file:
            print(file.read())
    except Exception as e:
        print(f"Error reading {file_name}: {e}")

def copy_to_pages():
    print("\nCopying files to _pages directory...")
    parent_pages_dir = os.path.join(os.path.dirname(os.getcwd()), '_pages')
    
    # Create _pages directory if it doesn't exist
    os.makedirs(parent_pages_dir, exist_ok=True)
    
    files_to_copy = ['projects.md', 'publications.md', 'teaching.md', 'conferences.md', 'cv.md', 'cv_euge_stumm.pdf']
    for file in files_to_copy:
        if os.path.exists(file):  # Only copy files that exist
            dest_path = os.path.join(parent_pages_dir, file)
            try:
                shutil.copy2(file, dest_path)
                print(f"Successfully copied {file} to _pages directory")
            except Exception as e:
                print(f"Error copying {file}: {e}")

def copy_cv_to_assets():
    """Copy cv_euge_stumm.pdf to assets directory"""
    print("\nCopying cv_euge_stumm.pdf to assets directory...")
    
    cv_pdf = 'cv_euge_stumm.pdf'
    assets_dir = os.path.join(os.path.dirname(os.getcwd()), 'assets')
    
    # Create assets directory if it doesn't exist
    os.makedirs(assets_dir, exist_ok=True)
    
    if os.path.exists(cv_pdf):
        dest_path = os.path.join(assets_dir, cv_pdf)
        try:
            shutil.copy2(cv_pdf, dest_path)
            print(f"Successfully copied {cv_pdf} to assets directory")
        except Exception as e:
            print(f"Error copying {cv_pdf} to assets: {e}")
    else:
        print(f"Error: {cv_pdf} not found - cannot copy to assets")

def convert_cv_to_pdf():
    """Convert cv.md to cv_euge_stumm.pdf using the separate pandoc converter script"""
    print("\nConverting cv.md to cv_euge_stumm.pdf using pandoc converter...")
    
    cv_md = 'cv.md'
    cv_pdf = 'cv_euge_stumm.pdf'
    css_file = os.path.join('pandoc', 'harvard_cv.css')
    
    # Check if input files exist
    if not os.path.exists(cv_md):
        print(f"Error: {cv_md} not found")
        return False
    
    if not os.path.exists(css_file):
        print(f"Error: CSS file not found at {css_file}")
        return False
    
    # Run the pandoc converter script
    try:
        result = subprocess.run([
            'python3', 'pandoc_converter.py'
        ], capture_output=True, text=True, check=True)
        
        print("PDF conversion completed successfully")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"Error during PDF conversion: {e}")
        if e.stderr:
            print(f"Stderr: {e.stderr}")
        return False

def main():
    # Run all updater scripts
    updater_scripts = [
        'updater_projects.py',
        'updater_publications.py',
        'updater_teaching.py',
        'updater_conferences.py',
        'updater_cv.py'
    ]
    
    for script in updater_scripts:
        run_updater(script)
    
    # Show content of generated files
    files_to_show = ['projects.md', 'publications.md', 'teaching.md', 'conferences.md', 'cv.md']
    for file in files_to_show:
        show_file_content(file)
    
    # Convert cv.md to cv_euge_stumm.pdf
    if convert_cv_to_pdf():
        # Copy cv_euge_stumm.pdf to assets directory only if conversion was successful
        copy_cv_to_assets()
    
    # Copy all files to _pages directory (including the generated PDF)
    copy_to_pages()

if __name__ == "__main__":
    main()