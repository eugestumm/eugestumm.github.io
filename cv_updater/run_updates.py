import subprocess
import os
import shutil
import sys

# The website's _pages folder (one level up from this script's folder)
PAGES_DIR = os.path.join(os.path.dirname(os.getcwd()), '_pages')


def run_updater(script_name, *args):
    """Run one script; report if it fails instead of carrying on silently."""
    print(f"\nRunning {script_name}...")
    result = subprocess.run([sys.executable, script_name, *args])
    if result.returncode != 0:
        print(f"⚠️  {script_name} exited with an error (code {result.returncode})")
    return result.returncode == 0


def show_file_content(file_name):
    print(f"\nContent of {file_name}:")
    try:
        with open(file_name, 'r') as file:
            print(file.read())
    except Exception as e:
        print(f"Error reading {file_name}: {e}")


def copy_to_pages():
    """Copy the website pages to ../_pages. The CV (cv.md and
    cv_euge_stumm.pdf) is published by create_cv_latex.py instead."""
    print("\nCopying files to _pages directory...")
    files_to_copy = ['projects.md', 'publications.md', 'teaching.md', 'conferences.md']
    for file in files_to_copy:
        if os.path.exists(file):  # Only copy files that exist
            dest_path = os.path.join(PAGES_DIR, file)
            try:
                shutil.copy2(file, dest_path)
                print(f"Successfully copied {file} to _pages directory")
            except Exception as e:
                print(f"Error copying {file}: {e}")


def main():
    # Create _pages first, so create_cv_latex.py can publish the CV into it
    os.makedirs(PAGES_DIR, exist_ok=True)

    # Run all updater scripts. updater_cv.py writes both CVs as Markdown:
    # cv.md (research) and cv_teaching.md (teaching).
    updater_scripts = [
        'updater_projects.py',
        'updater_teaching.py',
        'updater_conferences.py',
        'updater_cv.py',
        'updater_publications.py'
    ]

    for script in updater_scripts:
        run_updater(script)

    # Show content of generated files
    files_to_show = ['projects.md', 'publications.md', 'teaching.md', 'conferences.md', 'cv.md']
    for file in files_to_show:
        show_file_content(file)

    # Build both CV PDFs (cv_euge_stumm.pdf, cv_euge_stumm_teaching.pdf) and
    # publish the research CV: PDF to ../assets and ../_pages, and a
    # website-ready cv.md to ../_pages.
    run_updater('create_cv_latex.py')

    # Copy the other website pages to _pages
    copy_to_pages()


if __name__ == "__main__":
    main()