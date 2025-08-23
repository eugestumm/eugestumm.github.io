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
    
    files_to_copy = ['projects.md', 'publications.md', 'teaching.md']
    for file in files_to_copy:
        dest_path = os.path.join(parent_pages_dir, file)
        try:
            # Use replace instead of copy to overwrite existing files
            shutil.copy2(file, dest_path)
            print(f"Successfully replaced {file} in _pages directory")
        except Exception as e:
            print(f"Error copying {file}: {e}")

def check_git_available():
    try:
        # Try to run 'git --version' to check if git is available
        subprocess.run(['which', 'git'], check=True, capture_output=True)
        return True
    except subprocess.CalledProcessError:
        print("Error: Git is not installed or not in PATH")
        return False

def git_operations(operation='pull'):
    if not check_git_available():
        return

    try:
        # Get the full path to git executable
        git_path = subprocess.run(['which', 'git'], check=True, capture_output=True, text=True).stdout.strip()
        
        if operation == 'pull':
            print("\nPulling latest changes from remote...")
            subprocess.run([git_path, 'pull', 'origin', 'main'], check=True)
        elif operation == 'push':
            print("\nCommitting and pushing changes...")
            subprocess.run([git_path, 'add', '.'], check=True)
            subprocess.run([git_path, 'commit', '-m', 'Update CV content'], check=True)
            subprocess.run([git_path, 'push', 'origin', 'main'], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Git operation failed: {e}")
    except Exception as e:
        print(f"Unexpected error during git operation: {e}")

def main():
    # Run all updater scripts
    updater_scripts = [
        'updater_projects.py',
        'updater_publications.py',
        'updater_teaching.py'
    ]
    
    for script in updater_scripts:
        run_updater(script)
    
    # Show content of generated files
    files_to_show = ['projects.md', 'publications.md', 'teaching.md']
    for file in files_to_show:
        show_file_content(file)
    
    # Pull latest changes before copying to _pages
    git_operations('pull')
    
    # Copy files to _pages directory
    copy_to_pages()

    # Push changes to remote repository
    git_operations('push')

if __name__ == "__main__":
    main()
