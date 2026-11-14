#!/usr/bin/env python3
"""
Extract the Publications section from the research CV (cv.md, written by
updater_cv.py) and create the Jekyll publications page.

Run after updater_cv.py (run_updates.py already does this in that order).
"""
import re
import sys


def extract_publications(cv_file='cv.md', output_file='publications.md'):
    """
    Extract the Publications section from a CV markdown file
    and create a Jekyll-formatted publications page.

    Args:
        cv_file: Path to the CV markdown file
        output_file: Path to the output Jekyll page

    Returns True on success, False otherwise.
    """
    try:
        with open(cv_file, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print(f"{cv_file} not found. Run updater_cv.py first.")
        return False

    # Find the Publications section: from "## Publications" (any
    # capitalization, e.g. "## PUBLICATIONS") up to the next "## " section.
    # "### " sub-headings inside the section don't end it.
    pattern = r'^## publications[ \t]*\n(.*?)(?=^## |\Z)'
    match = re.search(pattern, content, re.DOTALL | re.MULTILINE | re.IGNORECASE)

    if not match:
        print(f"Publications section not found in {cv_file}")
        return False

    publications_content = match.group(1).strip()
    if not publications_content:
        print(f"The Publications section in {cv_file} is empty")
        return False

    # Sub-headings (### Peer-Reviewed Journal Articles, ...) become the
    # page's main headings (##)
    publications_content = re.sub(r'^### ', '## ', publications_content, flags=re.MULTILINE)

    jekyll_header = """---
layout: archive
title: "Publications"
permalink: /publications/
author_profile: true
---

"""

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(jekyll_header + publications_content + "\n")

    subsections = len(re.findall(r'^## ', publications_content, flags=re.MULTILINE))
    print(f"Publications section extracted successfully to {output_file}")
    print(f"Found {subsections} subsections (converted from h3 to h2)")
    return True


if __name__ == "__main__":
    # You can customize the file paths here.
    # A non-zero exit lets run_updates.py report the failure.
    sys.exit(0 if extract_publications('cv.md', 'publications.md') else 1)