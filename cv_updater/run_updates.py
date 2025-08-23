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
    
    files_to_copy = ['projects.md', 'publications.md', 'teaching.md', 'conferences.md', 'cv.md']
    for file in files_to_copy:
        dest_path = os.path.join(parent_pages_dir, file)
        try:
            # Use replace instead of copy to overwrite existing files
            shutil.copy2(file, dest_path)
            print(f"Successfully replaced {file} in _pages directory")
        except Exception as e:
            print(f"Error copying {file}: {e}")

def main():
    # Run all updater scripts
    updater_scripts = [
        'updater_projects.py',
        'updater_publications.py',
        'updater_teaching.py',
        'updater_conferences.py',
        'updater_cv.py'  # Added the CV updater script
    ]
    
    for script in updater_scripts:
        run_updater(script)
    
    # Show content of generated files
    files_to_show = ['projects.md', 'publications.md', 'teaching.md', 'conferences.md', 'cv.md']
    for file in files_to_show:
        show_file_content(file)
    
    # Copy files to _pages directory
    copy_to_pages()

if __name__ == "__main__":
    main()