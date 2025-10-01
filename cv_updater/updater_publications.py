#!/usr/bin/env python3
"""
Extract Publications section from CV markdown and create Jekyll page
"""
import re

def extract_publications(cv_file='cv.md', output_file='publications.md'):
    """
    Extract the Publications section from a CV markdown file
    and create a Jekyll-formatted publications page.
    
    Args:
        cv_file: Path to the CV markdown file
        output_file: Path to the output Jekyll page
    """
    # Read the CV file
    with open(cv_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find the Publications section
    # Match from "## Publications" until the next "##" (next major section)
    pattern = r'## Publications\s*\n(.*?)(?=\n## [A-Z]|\Z)'
    match = re.search(pattern, content, re.DOTALL)
    
    if not match:
        print("Publications section not found in CV")
        return
    
    # Store publications content in a variable
    publications_content = match.group(1).strip()
    
    # Process the content: convert h3 (###) to h2 (##)
    publications_content = re.sub(r'^### ', '## ', publications_content, flags=re.MULTILINE)
    
    # Create the Jekyll front matter
    jekyll_header = """---
layout: archive
title: "Publications"
permalink: /publications/
author_profile: true
---

"""
    
    # Combine header with processed publications content
    output_content = jekyll_header + publications_content
    
    # Write to output file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(output_content)
    
    print(f"Publications section extracted successfully to {output_file}")
    print(f"Found {publications_content.count('##')} subsections (converted from h3 to h2)")

if __name__ == "__main__":
    # You can customize the file paths here
    extract_publications('cv.md', 'publications.md')