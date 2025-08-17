#!/usr/bin/env python3
"""
Publications Generator from BibTeX
Reads BibTeX entries from an ODS spreadsheet and generates a Jekyll-compatible markdown file.
"""

import pandas as pd
import requests
from io import BytesIO
import re
from collections import defaultdict

def download_and_read_ods(url):
    """
    Download ODS file from URL and read the "Publications" tab
    """
    try:
        response = requests.get(url)
        response.raise_for_status()
        
        # Read ODS file from bytes, specifically the "Publications" sheet
        print("📋 Looking for 'Publications' sheet in the spreadsheet...")
        df = pd.read_excel(BytesIO(response.content), engine='odf', sheet_name='Publications')
        print("✅ Successfully found and read 'Publications' sheet")
        return df
    except Exception as e:
        print(f"❌ Error reading 'Publications' sheet: {e}")
        print("⚠️  Attempting to read all sheets to check available sheet names...")
        
        # Try to read all sheets to show available options
        try:
            response = requests.get(url)
            response.raise_for_status()
            all_sheets = pd.read_excel(BytesIO(response.content), engine='odf', sheet_name=None)
            available_sheets = list(all_sheets.keys())
            print(f"📊 Available sheets in the spreadsheet: {available_sheets}")
            
            # Try some common alternatives
            alternatives = ['publications', 'PUBLICATIONS', 'Publication', 'Pubs', 'Sheet1']
            for alt_name in alternatives:
                if alt_name in available_sheets:
                    print(f"🔄 Trying alternative sheet name: '{alt_name}'")
                    return all_sheets[alt_name]
            
            # If no alternatives work, use the first sheet
            if available_sheets:
                first_sheet = available_sheets[0]
                print(f"🔄 Using first available sheet: '{first_sheet}'")
                return all_sheets[first_sheet]
                
        except Exception as sheet_error:
            print(f"❌ Could not read any sheets from spreadsheet: {sheet_error}")
        
        return None

def parse_bibtex_entry(bibtex_str):
    """
    Parse a BibTeX entry string and extract key information
    """
    try:
        if not bibtex_str or pd.isna(bibtex_str):
            print("⚠️  Warning: Empty or invalid BibTeX entry")
            return None
            
        bibtex_str = str(bibtex_str).strip()
        if not bibtex_str.startswith('@'):
            print(f"⚠️  Warning: Invalid BibTeX format: {bibtex_str[:50]}...")
            return None
        
        # Extract entry type and key
        type_match = re.match(r'@(\w+)\{([^,]+),', bibtex_str)
        if not type_match:
            print(f"⚠️  Warning: Could not parse BibTeX entry type: {bibtex_str[:50]}...")
            return None
            
        entry_type = type_match.group(1).lower()
        entry_key = type_match.group(2)
        
        # Extract fields using regex
        fields = {}
        
        # Common field patterns
        field_patterns = [
            (r'author\s*=\s*\{([^}]+)\}', 'author'),
            (r'title\s*=\s*\{([^}]+)\}', 'title'),
            (r'journal\s*=\s*\{([^}]+)\}', 'journal'),
            (r'booktitle\s*=\s*\{([^}]+)\}', 'booktitle'),
            (r'year\s*=\s*\{([^}]+)\}', 'year'),
            (r'volume\s*=\s*\{([^}]+)\}', 'volume'),
            (r'number\s*=\s*\{([^}]+)\}', 'number'),
            (r'pages\s*=\s*\{([^}]+)\}', 'pages'),
            (r'publisher\s*=\s*\{([^}]+)\}', 'publisher'),
            (r'editor\s*=\s*\{([^}]+)\}', 'editor'),
            (r'doi\s*=\s*\{([^}]+)\}', 'doi'),
            (r'url\s*=\s*\{([^}]+)\}', 'url'),
            (r'isbn\s*=\s*\{([^}]+)\}', 'isbn'),
            (r'address\s*=\s*\{([^}]+)\}', 'address'),
        ]
        
        for pattern, field_name in field_patterns:
            match = re.search(pattern, bibtex_str, re.IGNORECASE)
            if match:
                fields[field_name] = match.group(1).strip()
        
        return {
            'type': entry_type,
            'key': entry_key,
            'fields': fields
        }
        
    except Exception as e:
        print(f"⚠️  Warning: Error parsing BibTeX entry: {e}")
        return None

def format_citation_mla(entry):
    """
    Format a BibTeX entry as MLA citation with clickable links and better formatting
    """
    try:
        if not entry or 'fields' not in entry:
            return "Citation formatting error"
        
        fields = entry['fields']
        citation_parts = []
        
        # Author - bold for better visibility
        if 'author' in fields:
            # Clean up author field - handle "and" separators
            authors = fields['author'].replace(' and ', ', ')
            # Highlight your name in the author list if present
            authors_highlighted = re.sub(r'\bStumm,?\s*E\.?\s*H?\.?', '**Stumm, E. H.**', authors)
            citation_parts.append(f"{authors_highlighted}.")
        
        # Title - make it clickable if URL/DOI available, and more prominent
        title_text = ""
        if 'title' in fields:
            title = fields['title']
            # Create clickable title
            if 'url' in fields:
                if entry['type'] in ['article', 'inproceedings']:
                    title_text = f'["{title}".]({fields["url"]})'
                else:
                    title_text = f'[*{title}.*]({fields["url"]})'
            elif 'doi' in fields:
                doi_url = f"https://doi.org/{fields['doi']}"
                if entry['type'] in ['article', 'inproceedings']:
                    title_text = f'["{title}".]({doi_url})'
                else:
                    title_text = f'[*{title}.*]({doi_url})'
            else:
                # No link available
                if entry['type'] in ['article', 'inproceedings']:
                    title_text = f'"{title}."'
                else:
                    title_text = f'*{title}.*'
            
            citation_parts.append(title_text)
        
        # Journal/Book information
        if entry['type'] == 'article' and 'journal' in fields:
            journal_info = f"*{fields['journal']}*"
            
            # Add volume and number
            if 'volume' in fields:
                journal_info += f", vol. {fields['volume']}"
            if 'number' in fields:
                journal_info += f", no. {fields['number']}"
                
            # Add year
            if 'year' in fields:
                journal_info += f", {fields['year']}"
                
            # Add pages
            if 'pages' in fields:
                pages = fields['pages'].replace('--', '-')  # Clean up page ranges
                journal_info += f", pp. {pages}"
                
            citation_parts.append(f"{journal_info}.")
            
        elif entry['type'] in ['book', 'incollection', 'inbook']:
            book_info = ""
            
            # For book chapters
            if entry['type'] in ['incollection', 'inbook']:
                if 'booktitle' in fields:
                    book_info = f"*{fields['booktitle']}*"
                elif 'journal' in fields:  # Sometimes booktitle is stored as journal
                    book_info = f"*{fields['journal']}*"
            
            # Add editor
            if 'editor' in fields:
                editors = fields['editor'].replace(' and ', ', ')
                book_info += f", edited by {editors}"
            
            # Add publisher
            if 'publisher' in fields:
                book_info += f", {fields['publisher']}"
            
            # Add year
            if 'year' in fields:
                book_info += f", {fields['year']}"
                
            # Add pages for chapters
            if 'pages' in fields and entry['type'] in ['incollection', 'inbook']:
                pages = fields['pages'].replace('--', '-')
                book_info += f", pp. {pages}"
            
            if book_info:
                citation_parts.append(f"{book_info}.")
        
        # Add access information as separate elements (not embedded in title)
        access_info = []
        if 'doi' in fields:
            access_info.append(f"[DOI: {fields['doi']}](https://doi.org/{fields['doi']})")
        
        # Only add URL if no DOI (to avoid duplication)
        if 'url' in fields and 'doi' not in fields:
            # Clean URL for display
            display_url = fields['url']
            if len(display_url) > 50:
                display_url = display_url[:47] + "..."
            access_info.append(f"[Web]({fields['url']})")
        
        # Add language note if it seems non-English
        if any(indicator in str(fields.get('title', '')).lower() for indicator in ['português', 'spanish', 'em português']):
            access_info.append("*(in Portuguese)*")
        
        # Combine citation
        main_citation = " ".join(citation_parts)
        
        if access_info:
            return f"{main_citation} {' | '.join(access_info)}"
        else:
            return main_citation
        
    except Exception as e:
        print(f"⚠️  Warning: Error formatting citation: {e}")
        return "Citation formatting error"

def categorize_publications(entries):
    """
    Categorize publications by type and sort by year (most recent first)
    """
    categories = {
        'Peer-reviewed journal articles': [],
        'Book chapters': [],
        'Books': [],
        'Conference proceedings': [],
        'Other publications': []
    }
    
    for entry in entries:
        if not entry or 'type' not in entry:
            continue
            
        entry_type = entry['type'].lower()
        year = int(entry['fields'].get('year', 0)) if entry['fields'].get('year', '').isdigit() else 0
        
        # Add year for sorting
        entry['sort_year'] = year
        
        if entry_type == 'article':
            categories['Peer-reviewed journal articles'].append(entry)
        elif entry_type in ['incollection', 'inbook']:
            categories['Book chapters'].append(entry)
        elif entry_type == 'book':
            categories['Books'].append(entry)
        elif entry_type in ['inproceedings', 'conference']:
            categories['Conference proceedings'].append(entry)
        else:
            categories['Other publications'].append(entry)
    
    # Sort each category by year (most recent first)
    for category in categories:
        categories[category].sort(key=lambda x: x.get('sort_year', 0), reverse=True)
        if categories[category]:
            print(f"✓ Found {len(categories[category])} entries for: {category}")
    
    # Remove empty categories
    categories = {k: v for k, v in categories.items() if v}
    
    return categories

def generate_jekyll_markdown(categories):
    """
    Generate Jekyll-compatible markdown with enhanced formatting for academic CVs
    """
    markdown_lines = [
        "---",
        "layout: archive",
        "title: \"Publications\"",
        "permalink: /publications/",
        "author_profile: true",
        "---",
        ""
    ]
    
    if not categories:
        markdown_lines.extend([
            "## No publications available",
            "",
            "*This section will be populated when publication data is added to the spreadsheet.*"
        ])
        print("⚠️  Warning: No publication categories found")
        return "\n".join(markdown_lines)
    
    total_publications = 0
    
    # Add summary at the top for impressive CVs
    total_pubs = sum(len(entries) for entries in categories.values())
    years_range = []
    for entries in categories.values():
        for entry in entries:
            year = entry.get('sort_year', 0)
            if year > 0:
                years_range.append(year)
    
    if years_range and total_pubs > 0:
        min_year, max_year = min(years_range), max(years_range)
        markdown_lines.extend([
            f"*{total_pubs} publications ({min_year}–{max_year})*",
            "",
            "---",
            ""
        ])
    
    for category, entries in categories.items():
        if not entries:
            continue
            
        markdown_lines.append(f"## {category}")
        markdown_lines.append("")
        
        for i, entry in enumerate(entries, 1):
            try:
                citation = format_citation_mla(entry)
                # Add some visual separation between entries
                markdown_lines.append(f"{i}. {citation}")
                markdown_lines.append("")
                total_publications += 1
            except Exception as e:
                print(f"⚠️  Warning: Error processing entry in {category}: {e}")
                continue
        
        # Add spacing between sections
        markdown_lines.append("---")
        markdown_lines.append("")
    
    # Enhanced footer section
    markdown_lines.extend([
        "## Additional Academic Profiles",
        "",
        "{% if author.googlescholar %}",
        "- [Google Scholar Profile]({{author.googlescholar}}) - *Complete publication metrics and citations*",
        "{% endif %}",
        "",
        "{% if author.orcid %}",
        "- [ORCID Profile](https://orcid.org/{{author.orcid}}) - *Verified academic identity and works*",
        "{% endif %}",
        "",
        "{% if author.researchgate %}",
        "- [ResearchGate Profile]({{author.researchgate}}) - *Academic networking and collaboration*",
        "{% endif %}",
        "",
        "{% if author.academia %}",
        "- [Academia.edu Profile]({{author.academia}}) - *Independent research platform*",
        "{% endif %}",
        "",
        "---",
        "",
        "*Last updated: {{ site.time | date: '%B %Y' }}*",
        "",
        "{% include base_path %}"
    ])
    
    if total_publications > 0:
        print(f"✅ Successfully processed {total_publications} publications total")
    else:
        print("⚠️  Warning: No publications were successfully processed")
    
    return "\n".join(markdown_lines)

def main():
    """
    Main function to generate publications Jekyll page from BibTeX data
    """
    # URL of the ODS file - you'll need to replace this with your actual URL
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS4TyMkL-aPWhYseDOToCruWUmoiM72tPAzGWvb_DauEtXZZxuHy3AVXFXAQ6DbEuU-T5S5yS9lt2xS/pub?output=ods"
    
    print("🔄 Starting Publications page generation...")
    print("📥 Downloading and reading spreadsheet...")
    
    try:
        df = download_and_read_ods(url)
        
        if df is None or df.empty:
            print("❌ Failed to read spreadsheet or spreadsheet is empty")
            # Create a minimal publications.md file
            with open("publications.md", 'w', encoding='utf-8') as f:
                f.write("---\nlayout: archive\ntitle: \"Publications\"\npermalink: /publications/\nauthor_profile: true\n---\n\nNo publication data available. Please check the spreadsheet URL and format.\n")
            return
        
        print(f"✅ Successfully read {len(df)} rows from spreadsheet")
        print(f"📊 Columns found: {list(df.columns)}")
        
        # Check for bibtex column
        if 'bibtex' not in df.columns:
            print("❌ 'bibtex' column not found in spreadsheet")
            print(f"Available columns: {list(df.columns)}")
            return
        
        # Parse BibTeX entries
        print("📚 Parsing BibTeX entries...")
        entries = []
        for index, row in df.iterrows():
            bibtex_str = row.get('bibtex', '')
            if bibtex_str and not pd.isna(bibtex_str):
                entry = parse_bibtex_entry(bibtex_str)
                if entry:
                    entries.append(entry)
                    print(f"✓ Parsed entry: {entry.get('key', 'Unknown')}")
                else:
                    print(f"⚠️  Warning: Failed to parse BibTeX entry at row {index + 1}")
        
        if not entries:
            print("❌ No valid BibTeX entries found")
            with open("publications.md", 'w', encoding='utf-8') as f:
                f.write("---\nlayout: archive\ntitle: \"Publications\"\npermalink: /publications/\nauthor_profile: true\n---\n\nNo valid BibTeX entries found in spreadsheet.\n")
            return
        
        print(f"✅ Successfully parsed {len(entries)} BibTeX entries")
        
        # Categorize publications
        print("📂 Categorizing publications...")
        categories = categorize_publications(entries)
        
        # Generate markdown
        print("📝 Generating Jekyll markdown...")
        markdown_content = generate_jekyll_markdown(categories)
        
        # Write to file
        output_file = "publications.md"
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            print(f"✅ Successfully generated {output_file}")
            
            print(f"\n📋 Preview of generated content:")
            print("-" * 50)
            preview_length = min(800, len(markdown_content))
            print(markdown_content[:preview_length])
            if len(markdown_content) > preview_length:
                print("...")
            print("-" * 50)
            
        except Exception as e:
            print(f"❌ Error writing file: {e}")
            print("⚠️  Printing content to console instead:")
            print(markdown_content)
            
    except Exception as e:
        print(f"❌ Unexpected error during publications generation: {e}")
        print("⚠️  Creating minimal publications.md file...")
        try:
            with open("publications.md", 'w', encoding='utf-8') as f:
                f.write(f"---\nlayout: archive\ntitle: \"Publications\"\npermalink: /publications/\nauthor_profile: true\n---\n\nError generating publications: {e}\n\nPlease check the spreadsheet format and try again.\n")
        except Exception as file_error:
            print(f"❌ Could not even create error file: {file_error}")
    
    print("🏁 Publications generation completed.")

if __name__ == "__main__":
    # Required packages: pandas, requests, odfpy
    # Install with: pip install pandas requests odfpy
    main()