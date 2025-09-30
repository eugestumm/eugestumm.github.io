import pandas as pd
import requests
from io import BytesIO
import re
import numpy as np
from datetime import datetime

def fetch_sheet_data(url):
    """Fetch data from Google Sheets URL and return a dictionary of DataFrames"""
    try:
        response = requests.get(url)
        response.raise_for_status()
        
        # Read the ODS file
        ods_data = BytesIO(response.content)
        xl = pd.ExcelFile(ods_data)
        
        # Create a dictionary of all sheets
        sheets_dict = {}
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            # Replace NaN values with empty strings
            df = df.replace({np.nan: None})
            sheets_dict[sheet_name] = df
        
        return sheets_dict
    except Exception as e:
        print(f"Error fetching data: {e}")
        return None

def parse_bibtex(bibtex):
    """Parse bibtex entry and return a dictionary of fields"""
    if not bibtex:
        return {}
    
    entry = {}
    
    # Extract entry type
    entry_type_match = re.search(r'@(\w+){', bibtex)
    if entry_type_match:
        entry['type'] = entry_type_match.group(1)
    
    # Extract key
    key_match = re.search(r'@\w+{([^,]+),', bibtex)
    if key_match:
        entry['key'] = key_match.group(1)
    
    # Extract all fields
    field_pattern = r'(\w+)\s*=\s*{([^}]*)}'
    fields = re.findall(field_pattern, bibtex)
    
    for field, value in fields:
        entry[field] = value
    
    return entry

def format_author_name(author_str):
    """Format author names in Harvard style"""
    if not author_str:
        return ""
    
    authors = author_str.split(' and ')
    formatted_authors = []
    
    for author in authors:
        # Handle case where author is already in "Last, First" format
        if ',' in author:
            parts = author.split(',')
            if len(parts) >= 2:
                last = parts[0].strip()
                first = parts[1].strip()
                formatted_authors.append(f"{last}, {first}")
            else:
                formatted_authors.append(author.strip())
        else:
            # Convert "First Last" to "Last, First"
            parts = author.strip().split(' ')
            if len(parts) >= 2:
                last = parts[-1]
                first = ' '.join(parts[:-1])
                formatted_authors.append(f"{last}, {first}")
            else:
                formatted_authors.append(author.strip())
    
    # Join authors with semicolons
    return '; '.join(formatted_authors)

def format_date(year):
    """Format year, handling ranges and single years"""
    if not year:
        return ""
    
    if isinstance(year, (int, float)):
        return str(int(year))
    
    if isinstance(year, str):
        if ' to ' in year:
            return year.replace(' to ', '–')
        return year
    
    return ""

def format_currency_amount(amount):
    """Format currency amounts, handling different currencies"""
    if not amount:
        return ""
    
    amount_str = str(amount).strip()
    
    # Common currency symbols and codes
    currency_patterns = {
        'R$': 'BRL',
        'USD': 'USD',
        'US$': 'USD',
        '$': 'USD',
        '€': 'EUR',
        '£': 'GBP'
    }
    
    # If amount already has currency info, return as is
    for symbol in currency_patterns.keys():
        if symbol in amount_str:
            return amount_str
    
    # If it's just a number, assume it needs formatting
    try:
        # Try to parse as number
        float(amount_str.replace(',', ''))
        # If successful, format with proper thousand separators
        if ',' in amount_str or '.' in amount_str:
            return amount_str  # Already formatted
        else:
            # Add thousand separators for large numbers
            num = float(amount_str)
            if num >= 1000:
                return f"{num:,.0f}"
            else:
                return amount_str
    except ValueError:
        # Not a number, return as is
        return amount_str

def format_date_range(start_date, end_date):
    """Format date ranges for multi-year positions"""
    if not start_date:
        return ""
    
    start_formatted = format_date(start_date)
    
    if not end_date or end_date == start_date:
        return start_formatted
    
    end_formatted = format_date(end_date)
    return f"{start_formatted}–{end_formatted}"

def init_cv():
    """Initialize CV with YAML front matter and header"""
    cv_content = "---\n"
    #cv_content += "layout: archive\n"
    cv_content += "author: Euge Stumm\n"
    #cv_content += "fontsize: 12pt\n"
    #cv_content += "papersize: \"letter\"\n"
    cv_content += "geometry: \"margin=1in\"\n"
    #cv_content += "mainfont: \"Times New Roman\"\n"
    #cv_content += "title: \"Curriculum Vitae\"\n"
    #cv_content += "permalink: /cv/\n"
    #cv_content += "author_profile: true\n"
    cv_content += "---\n"
    
    return cv_content

def generate_header_section(data):
    """
    Generate the CV header with name, title, contact info, and research statement.
    Expects a sheet named 'Header' or 'Profile'.
    """
    header_sheet_name = None
    for sheet in ['Header', 'Profile']:
        if sheet in data:
            header_sheet_name = sheet
            break
    
    if not header_sheet_name:
        # Fallback to a minimal header if the sheet doesn't exist
        return "# Your Name\n\n*Your Title* \n**Email:** [email@example.com](mailto:email@example.com) | **Website:** [yourwebsite.com](https://yourwebsite.com) | **ORCID:** [0000-1234-5678-9012](https://orcid.org/your-orcid) | **GitHub:** [@yourusername](https://github.com/yourusername)\n\n---\n\n"
    
    header_df = data[header_sheet_name]
    # We assume the first row contains the data
    row = header_df.iloc[0] if not header_df.empty else {}
    
    name = row.get('name', 'Your Name')
    title = row.get('title', 'Your Title')
    email = row.get('email', '')
    website = row.get('website', '')
    orcid = row.get('orcid', '')
    github = row.get('github', '')
    linkedin = row.get('linkedin', '')
    research_statement = row.get('research_statement', '')
    address = row.get('address', '')
    
    content = f"# {name}\n\n"
    if title:
        content += f"*{title}*\n\n"
    
    # Build the contact info line with labels and clean URLs
    contact_info = []
    
    if email:
        contact_info.append(f"**Email:** [{email}](mailto:{email})")
    
    if website:
        # Clean up website URL for display - remove protocol
        website_clean = website.replace('https://', '').replace('http://', '').strip('/')
        # Ensure the actual link has https://
        website_url = website if website.startswith(('http://', 'https://')) else f"https://{website_clean}"
        contact_info.append(f"**Website:** [{website_clean}]({website_url})")
    
    if orcid:
        # Format ORCID nicely: extract just the ID from a full URL if needed
        if 'orcid.org' in orcid:
            orcid_id = orcid.split('/')[-1]
        else:
            orcid_id = orcid
        contact_info.append(f"**ORCID:** [{orcid_id}](https://orcid.org/{orcid_id})")
    
    if github:
        # Clean GitHub URL - remove protocol and extract username
        if github.startswith('http'):
            github_clean = github.replace('https://github.com/', '').replace('http://github.com/', '').strip('/')
        else:
            github_clean = github.replace('github.com/', '').replace('@', '').strip('/')
        contact_info.append(f"**GitHub:** [@{github_clean}](https://github.com/{github_clean})")
    
    if linkedin:
        # Clean LinkedIn URL - remove protocol and extract username
        if linkedin.startswith('http'):
            linkedin_clean = linkedin.replace('https://linkedin.com/in/', '').replace('http://linkedin.com/in/', '')
            linkedin_clean = linkedin_clean.replace('https://www.linkedin.com/in/', '').replace('http://www.linkedin.com/in/', '')
            linkedin_clean = linkedin_clean.strip('/')
        else:
            linkedin_clean = linkedin.replace('linkedin.com/in/', '').replace('@', '').strip('/')
        contact_info.append(f"**LinkedIn:** [{linkedin_clean}](https://linkedin.com/in/{linkedin_clean})")
    
    if address:
        contact_info.append(address)  # Physical address usually isn't linked
    
    if contact_info:
        content += ' | '.join(contact_info) + "\n\n"
    
    # Research statement is crucial for interdisciplinary fields
    if research_statement:
        content += f"\n{research_statement}\n"
    
    content += "\n---\n\n"
    return content

def generate_education_section(data):
    """Generate Education section with bullet lists Pandoc recognizes"""
    import datetime
    current_year = datetime.datetime.now().year

    content = "## Education\n\n"

    if "Training" in data:
        training_df = data["Training"]

        # Separate PhDs and graduate certificates
        phd_entries = []
        other_entries = []
        graduate_certificates = []

        for _, row in training_df.iterrows():
            if row.get("degree") and row.get("title"):
                degree = row["degree"].replace("_", " ").lower()
                if degree == "doctorate":
                    phd_entries.append(row)
                elif degree == "graduate certificate":
                    graduate_certificates.append(row)
                else:
                    other_entries.append(row)

        def format_degree_line(degree_label, title, university, year, gpa):
            """Return the bold degree line with year and GPA inline"""
            line = f"**{degree_label} in {title}**"
            if year:
                try:
                    year_int = int(year)
                    if year_int > current_year:
                        line += f", Expected {year}"
                    else:
                        line += f", {year}"
                except ValueError:
                    line += f", {year}"
            if university:
                line += f", {university}"
            if gpa:
                line += f", GPA: {gpa}"
            return line

        def format_committee(row):
            members = []
            for i in range(1, 6):
                member = row.get(f"committee_member{i}", "")
                member_url = row.get(f"committee_member{i}_url", "")
                if member:
                    if member_url:
                        members.append(f"[{member}]({member_url})")
                    else:
                        members.append(member)
            return members

        # Process PhD entries first (with nested certificates)
        for row in phd_entries:
            title = row["title"].replace("Ph.D. in ", "")
            degree_line = format_degree_line(
                "Ph.D.", title, row.get("university", ""), row.get("year", ""), row.get("gpa", "")
            )
            content += f"{degree_line}\n\n"  # blank line before bullets

            # Thesis
            if row.get("thesis"):
                content += f"- Thesis: *{row['thesis']}*\n"

            # Advisor
            if row.get("advisor"):
                if row.get("advisor_link"):
                    content += f"- Advisor: [{row['advisor']}]({row['advisor_link']})\n"
                else:
                    content += f"- Advisor: {row['advisor']}\n"

            # Committee
            committee_members = format_committee(row)
            if committee_members:
                content += f"- Committee: {', '.join(committee_members)}\n"

            # Related graduate certificates
            phd_university = row.get("university", "")
            for cert_row in graduate_certificates:
                cert_university = cert_row.get("university", "")
                if cert_university == phd_university or not cert_university:
                    cert_line = f"**Graduate Certificate in {cert_row.get('title','')}**"
                    cert_year = cert_row.get("year", "")
                    if cert_year:
                        try:
                            cert_year_int = int(cert_year)
                            if cert_year_int > current_year:
                                cert_line += f", Expected {cert_year}"
                            else:
                                cert_line += f", {cert_year}"
                        except ValueError:
                            cert_line += f", {cert_year}"
                    content += f"- {cert_line}\n"

            content += "\n"

        # Process other degree entries (M.A., B.A., etc.)
        for row in other_entries:
            degree = row["degree"].replace("_", " ").title()
            title = row["title"].replace("M.A. in ", "").replace("B.A. in ", "")
            degree_label = (
                "M.A." if degree.lower() == "masters"
                else "B.A." if degree.lower() == "bachelors"
                else degree
            )
            degree_line = format_degree_line(
                degree_label, title, row.get("university", ""), row.get("year", ""), row.get("gpa", "")
            )
            content += f"{degree_line}\n\n"  # blank line before bullets

            if row.get("thesis"):
                content += f"- Thesis: *{row['thesis']}*\n"

            if row.get("advisor"):
                if row.get("advisor_link"):
                    content += f"- Advisor: [{row['advisor']}]({row['advisor_link']})\n"
                else:
                    content += f"- Advisor: {row['advisor']}\n"

            committee_members = format_committee(row)
            if committee_members:
                content += f"- Committee: {', '.join(committee_members)}\n"

            content += "\n"

    content += "---\n\n"
    return content


def generate_publications_section(data):
    """Generate Publications section"""
    content = "## Publications\n\n"
    
    if 'Publications' in data:
        publications_df = data['Publications']
        
        # Separate publication types
        journal_articles = []
        book_chapters = []
        other_publications = []
        
        for _, row in publications_df.iterrows():
            if row.get('bibtex'):
                bibtex_entry = parse_bibtex(row['bibtex'])
                
                if bibtex_entry.get('type') == 'article':
                    journal_articles.append(bibtex_entry)
                elif bibtex_entry.get('type') in ['incollection', 'inproceedings']:
                    book_chapters.append(bibtex_entry)
                else:
                    other_publications.append(bibtex_entry)
        
        # Sort by year (descending)
        def sort_by_year(entries):
            return sorted(entries, key=lambda x: int(x.get('year', 0) or 0), reverse=True)
        
        journal_articles = sort_by_year(journal_articles)
        book_chapters = sort_by_year(book_chapters)
        other_publications = sort_by_year(other_publications)
        
        # Helper function to format title with translation
        def format_title_with_translation(title, titleaddon):
            if not title:
                return ""
            
            title_clean = title.strip('"{}')
            formatted_title = f'"{title_clean}"'
            
            # Add translation in brackets if available
            if titleaddon:
                titleaddon_clean = titleaddon.strip('"{}')
                formatted_title += f" [{titleaddon_clean}]"
            
            # Add period after the title
            formatted_title += "."
            
            return formatted_title + " "
        
        # Journal Articles
        if journal_articles:
            content += "### Peer-Reviewed Journal Articles\n\n"
            
            for i, article in enumerate(journal_articles, 1):
                authors = format_author_name(article.get('author', ''))
                title = article.get('title', '')
                titleaddon = article.get('titleaddon', '')
                journal = article.get('journal', '')
                year = article.get('year', '')
                volume = article.get('volume', '')
                number = article.get('number', '')
                pages = article.get('pages', '')
                doi = article.get('doi', '')
                url = article.get('url', '')
                
                content += f"{i}. {authors} ({year}). "
                
                # Format title with translation
                content += format_title_with_translation(title, titleaddon)
                
                if journal:
                    content += f"*{journal}*"
                
                if volume:
                    content += f", {volume}"
                    if number:
                        content += f"({number})"
                
                if pages:
                    content += f": {pages}."
                else:
                    content += "."
                
                # Add URL or DOI at the end
                if url:
                    content += f" [{url}]({url})"
                elif doi:
                    content += f" DOI: [{doi}](https://doi.org/{doi})"
                
                content += "\n\n"
        
        # Book Chapters
        if book_chapters:
            content += "### Book Chapters and Edited Volumes\n\n"
            
            for i, chapter in enumerate(book_chapters, 1):
                authors = format_author_name(chapter.get('author', ''))
                title = chapter.get('title', '')
                titleaddon = chapter.get('titleaddon', '')
                booktitle = chapter.get('booktitle', '')
                year = chapter.get('year', '')
                pages = chapter.get('pages', '')
                publisher = chapter.get('publisher', '')
                editor = chapter.get('editor', '')
                url = chapter.get('url', '')
                
                content += f"{i}. {authors} ({year}). "
                
                # Format title with translation
                content += format_title_with_translation(title, titleaddon)
                
                if booktitle:
                    content += f"In *{booktitle}*"
                
                if editor:
                    content += f", edited by {editor}"
                
                if pages:
                    content += f", pp. {pages}"
                
                if publisher:
                    content += f". {publisher}."
                else:
                    content += "."
                
                # Add URL at the end
                if url:
                    content += f" [{url}]({url})"
                
                content += "\n\n"
        
        # Other Publications
        if other_publications:
            content += "### Other Publications\n\n"
            
            for i, pub in enumerate(other_publications, 1):
                authors = format_author_name(pub.get('author', ''))
                title = pub.get('title', '')
                titleaddon = pub.get('titleaddon', '')
                year = pub.get('year', '')
                publisher = pub.get('publisher', '')
                url = pub.get('url', '')
                
                content += f"{i}. {authors} ({year}). "
                
                # Format title with translation
                content += format_title_with_translation(title, titleaddon)
                
                if publisher:
                    content += f"{publisher}."
                
                # Add URL at the end
                if url:
                    content += f" [{url}]({url})"
                
                content += "\n\n"
    
    content += "---\n\n"
    return content

def generate_teaching_section(data):
    """Generate Teaching Experience section with support for multi-year positions and numbered courses"""
    content = "## Teaching Experience\n\n"
    if 'Teaching' in data:
        teaching_df = data['Teaching']
        # Exclude workshops from main teaching section
        regular_teaching = teaching_df[teaching_df['category'] != 'Workshop']
        # Group by institution first
        institutions = [inst for inst in regular_teaching['institution'].unique() if inst]
        for institution in institutions:
            content += f"### {institution}\n\n"
            inst_df = regular_teaching[regular_teaching['institution'] == institution]
            # Group by category within institution
            categories = [cat for cat in inst_df['category'].unique() if cat]
            for category in categories:
                content += f"**{category}**\n\n" # Changed from #### to **bold**
                category_df = inst_df[inst_df['category'] == category]
                # Sort by year (descending)
                category_df = category_df.sort_values('year', ascending=False, na_position='last')
                
                # Initialize course counter for this specific category
                course_number = 1
                
                for _, row in category_df.iterrows():
                    if row.get('course_title'):
                        course_code = row.get('course_code', '')
                        course_title = row['course_title']
                        semester = row.get('semester', '')
                        year = row.get('year', '')
                        end_year = row.get('end_year', '') # New field for multi-year positions
                        supervisor = row.get('supervisor', '')
                        co_instructor = row.get('co_instructor', '')
                        special_notes = row.get('special_notes', '')
                        
                        # CHANGED: Use number instead of bullet
                        content += f"{course_number}. **{course_title}**"
                        course_number += 1 # Increment the counter
                        
                        if course_code:
                            content += f" ({course_code})"
                        # Handle date ranges
                        if year:
                            if end_year and end_year != year:
                                date_range = format_date_range(year, end_year)
                                if semester:
                                    content += f" — {semester} {date_range}"
                                else:
                                    content += f" — {date_range}"
                            else:
                                if semester:
                                    content += f" — {semester} {format_date(year)}"
                                else:
                                    content += f" — {format_date(year)}"
                        content += "\n"
                        
                        details = []
                        if supervisor:
                            details.append(f"Supervised by {supervisor}")
                        if co_instructor:
                            details.append(f"Co-taught with {co_instructor}")
                        if special_notes:
                            details.append(special_notes)
                        if details:
                            for detail in details:
                                content += f"    - {detail}\n" # Indented bullet for details
                        content += "\n"
                content += "\n" # Extra space after each category
            content += "\n" # Extra space after each institution
        
        # Add separator only at the end of the entire section
        content += "---\n\n"
    
    return content

def generate_conferences_section(data):
    """Generate Conference Presentations section in Harvard style - compact single line format"""
    content = "## Conference Presentations\n\n"
    
    if 'Conferences' in data:
        # Add this right after getting the conferences_df
        conferences_df = data['Conferences'].copy()

        # Convert year to numeric, coercing errors to NaN
        conferences_df['year'] = pd.to_numeric(conferences_df['year'], errors='coerce')

        # NOW sort by year (descending)
        conferences_df = conferences_df.sort_values('year', ascending=False, na_position='last')
        # Sort by year (descending)
        conferences_df = conferences_df.sort_values('year', ascending=False, na_position='last')
        
        # Define section mapping based on role
        sections = {
            'Graduate Presentations': 'presenter_graduate',
            'Undergraduate Presentations': 'presenter_undergraduate'
        }
        
        for section_title, role_value in sections.items():
            # Filter by role
            section_df = conferences_df[conferences_df['role'] == role_value]
            
            if not section_df.empty:
                content += f"### {section_title}\n\n"
                
                for _, row in section_df.iterrows():
                    # Get all necessary fields
                    title = row.get('title', '')
                    event_name = row.get('event_name', '')
                    display_title = title or event_name
                    
                    if not display_title:
                        continue  # Skip if no title or event name
                    
                    author = row.get('author', '')
                    co_authors = row.get('co_authors', '')
                    event_theme = row.get('event_theme', '')
                    conference_number = row.get('conference_number', '')
                    institution = row.get('institution', '')
                    city = row.get('city', '')
                    country = row.get('country', '')
                    day = row.get('day', '')
                    month = row.get('month', '')
                    year = row.get('year', '')
                    
                    # Build CV formatted entry: "Title." *Conference Name*, Month Year, Location.
                    line_parts = []
                    
                    # Title in quotes
                    line_parts.append(f'"{display_title}."')
                    
                    # Conference name in italics (using markdown)
                    conference_name = ""
                    if event_name and title and event_name != title:
                        conference_name = event_name
                    elif event_name:
                        conference_name = event_name
                    
                    if conference_name:
                        if event_theme and conference_name != event_theme:
                            conference_name += f": {event_theme}"
                        if conference_number:
                            conference_name += f" ({conference_number})"
                        line_parts.append(f"*{conference_name}*,")
                    elif event_theme:
                        theme_name = event_theme
                        if conference_number:
                            theme_name += f" ({conference_number})"
                        line_parts.append(f"*{theme_name}*,")
                    
                    # Date in "Month Year" format only
                    date_str = f"{month} {year}" if month and year else year or ""
                    
                    # Location
                    location_parts = [part for part in [institution, city, country] if part]
                    location_str = ", ".join(location_parts)
                    
                    # Combine date and location
                    if date_str and location_str:
                        line_parts.append(f"{date_str}, {location_str}.")
                    elif date_str:
                        line_parts.append(f"{date_str}.")
                    elif location_str:
                        line_parts.append(f"{location_str}.")
                    
                    # Add co-authors if present (since you're the main author)
                    if co_authors:
                        line_parts.append(f"With {co_authors}.")
                    
                    # Join all parts with single space and add to content
                    content += " ".join(line_parts) + "\n\n"  # Each entry on its own line with spacing
                
                content += "\n\n"
    
    return content

def generate_awards_section(data):
    """Generate Awards and Honors section with proper currency handling"""
    if 'Awards' not in data or data['Awards'].empty:
        return ""
    
    awards_df = data['Awards']
    
    # Check if there are any actual awards
    has_awards = any(row.get('award_name') for _, row in awards_df.iterrows())
    if not has_awards:
        return ""
    
    # --- Local helpers ---
    def format_date(year):
        """Format year for display"""
        if not year:
            return ""
        try:
            return str(int(year))
        except (ValueError, TypeError):
            return str(year)

    def format_currency_amount(amount):
        """Format currency amount when no currency is specified"""
        if not amount:
            return ""
        try:
            return f"${float(amount):,.2f}"
        except (ValueError, TypeError):
            return str(amount)
    # ----------------------

    content = "## Awards and Honors\n\n"
    
    # Sort by year (descending)
    awards_df = awards_df.sort_values('year', ascending=False, na_position='last')
    
    for _, row in awards_df.iterrows():
        if row.get('award_name'):
            award_name = row['award_name']
            institution = row.get('institution', '')
            year = row.get('year', '')
            amount = row.get('amount', '')
            currency = row.get('currency', '')
            description = row.get('description', '')
            
            # Award name and year
            content += f"**{award_name}**"
            if year:
                content += f" ({format_date(year)})"
            content += "  \n"  # Two spaces at end for markdown line break
            
            # Institution
            if institution:
                content += f"*{institution}*  \n"  # Two spaces at end
            
            # Amount with currency
            if amount and str(amount).strip():
                if currency:
                    if currency.upper() == 'BRL':
                        content += f"Amount: R$ {amount:,.2f}  \n"  # Two spaces at end
                    elif currency.upper() == 'USD':
                        content += f"Amount: ${amount:,.2f}  \n"  # Two spaces at end
                    else:
                        content += f"Amount: {amount} {currency}  \n"  # Two spaces at end
                else:
                    formatted_amount = format_currency_amount(amount)
                    content += f"Amount: {formatted_amount}  \n"  # Two spaces at end
            
            # Description
            if description and str(description).strip():
                content += f"{description}  \n"  # Two spaces at end
            
            content += "\n"  # Empty line between entries
    
    return content

def generate_funded_research_section(data):
    """Generate Funded Research section for fellowships and grants"""
    print("Checking for FundedResearch sheet...")
    
    if 'FundedResearch' not in data:
        print("FundedResearch sheet not found in data")
        print("Available sheets:", list(data.keys()))
        return ""
    
    funded_df = data['FundedResearch']
    print(f"FundedResearch sheet found with {len(funded_df)} rows")
    
    # Check if there are any valid funded entries
    has_entries = any(row.get('title') for _, row in funded_df.iterrows())
    if not has_entries:
        print("No valid entries found in FundedResearch sheet")
        return ""
    
    content = "## Funded Research\n\n"
    
    # Sort by start_year (descending)
    funded_df = funded_df.sort_values('start_year', ascending=False, na_position='last')
    
    for _, row in funded_df.iterrows():
        title = row.get('title', '')
        institution = row.get('institution', '')
        start_year = row.get('start_year', '')
        end_year = row.get('end_year', '')
        amount = row.get('amount', '')
        currency = row.get('currency', '')
        pi = row.get('pi', '')
        
        if not title:
            continue
        
        # Title and date range
        content += f"**{title}**"
        if start_year:
            if end_year and end_year != start_year:
                content += f" ({format_date(start_year)}–{format_date(end_year)})"
            else:
                content += f" ({format_date(start_year)})"
        content += "  \n"  # Two spaces at end for markdown line break
        
        # Institution
        if institution:
            content += f"*{institution}*  \n"  # Two spaces at end
        
        # Amount with currency
        if amount and str(amount).strip():
            if currency:
                if currency.upper() == 'BRL':
                    content += f"Amount: R$ {float(amount):,.2f}  \n"  # Two spaces at end
                elif currency.upper() == 'USD':
                    content += f"Amount: ${float(amount):,.2f}  \n"  # Two spaces at end
                else:
                    content += f"Amount: {amount} {currency}  \n"  # Two spaces at end
            else:
                content += f"Amount: {format_currency_amount(amount)}  \n"  # Two spaces at end
        
        # PI
        if pi:
            content += f"Supervisor: {pi}  \n"  # Two spaces at end
        
        content += "\n"  # Empty line between entries
    
    return content

def generate_research_section(data):
    """Generate Projects section"""
    content = "## Projects\n\n"
    if 'Projects' in data:
        projects_df = data['Projects']
        # Sort projects by most recent first
        # Convert start_date to datetime for proper sorting, handling potential None values
        projects_df['start_date_parsed'] = pd.to_datetime(projects_df['start_date'], errors='coerce')
        projects_df_sorted = projects_df.sort_values('start_date_parsed', ascending=False, na_position='last')
        
        # Check if project_type column exists and has values
        if 'project_type' in projects_df_sorted.columns and projects_df_sorted['project_type'].notna().any():
            # Group by project_type
            # Get unique project types, prioritizing Digital Scholarship first
            unique_types = projects_df_sorted['project_type'].dropna().unique()
            
            # Sort so Digital Scholarship comes first, then Creative Work, then others
            type_order = []
            for t in ['Digital Scholarship', 'Creative Work']:
                if t in unique_types:
                    type_order.append(t)
            # Add any other types
            for t in unique_types:
                if t not in type_order:
                    type_order.append(t)
            
            for project_type in type_order:
                project_group = projects_df_sorted[projects_df_sorted['project_type'] == project_type]
                
                if not project_group.empty:
                    content += f"### {project_type}\n\n"
                    
                    for _, row in project_group.iterrows():
                        if row.get('title'):
                            short_title = row.get('short_title', '')
                            title = row['title']
                            start_date = row.get('start_date', '')
                            end_date = row.get('end_date', '')
                            url = row.get('url', '')
                            
                            display_title = short_title if short_title else title
                            
                            content += f'"{display_title}" '

                            if start_date:
                                date_range = format_date(start_date)
                                if end_date and end_date != start_date:
                                    date_range += f"–{format_date(end_date)}"
                                content += f"({date_range})"
                            else:
                                content += f"({date_range})"

                            
                            # Add URL as markdown link if available
                            if pd.notna(url):
                                content += f" [{url}]({url})"
                            
                            content += "\n\n"
        else:
            # If no project_type column or all values are NaN, just list all projects
            for _, row in projects_df_sorted.iterrows():
                if row.get('title'):
                    short_title = row.get('short_title', '')
                    title = row['title']
                    start_date = row.get('start_date', '')
                    end_date = row.get('end_date', '')
                    url = row.get('url', '')
                    
                    display_title = short_title if short_title else title
                    
                    content += f'"{display_title}"'
                    
                    # Use end_date if it exists, otherwise show start_date-present
                    if pd.notna(end_date):
                        end_year = pd.to_datetime(end_date, errors='coerce').year
                        if pd.notna(end_year):
                            content += f" ({int(end_year)})"
                    elif pd.notna(start_date):
                        start_year = pd.to_datetime(start_date, errors='coerce').year
                        if pd.notna(start_year):
                            content += f" ({int(start_year)}-present)"
                    
                    if pd.notna(url):
                        content += f", [{url}]({url})"
                    
                    content += ".\n\n"
        
        # Clean up the temporary column
        projects_df.drop('start_date_parsed', axis=1, inplace=True)
    
    return content

def generate_workshops_section(data):
    """Generate Workshops and Training section"""
    content = "## Workshops and Professional Development\n\n"
    
    if 'Teaching' in data:
        teaching_df = data['Teaching']
        workshops_df = teaching_df[teaching_df['category'] == 'Workshop']
        
        if not workshops_df.empty:
            # Sort by year (descending)
            workshops_df = workshops_df.sort_values('year', ascending=False, na_position='last')
            
            for _, row in workshops_df.iterrows():
                if row.get('course_title'):
                    title = row['course_title']
                    institution = row.get('institution', '')
                    year = row.get('year', '')
                    supervisor = row.get('supervisor', '')
                    co_instructor = row.get('co_instructor', '')
                    description = row.get('description', '')
                    
                    content += f"**{title}**"
                    if institution and year:
                        content += f" — {institution}, {format_date(year)}"
                    elif institution:
                        content += f" — {institution}"
                    elif year:
                        content += f" — {format_date(year)}"
                    content += "\n"
                    
                    if description:
                        content += f"{description}\n"
                    
                    details = []
                    if supervisor:
                        details.append(f"Supervised by {supervisor}")
                    if co_instructor:
                        details.append(f"Co-facilitated with {co_instructor}")
                    
                    if details:
                        content += f"*{'; '.join(details)}*\n"
                    
                    content += "\n"
        else:
            content += "*No workshops recorded*\n\n"
    else:
        content += "*No workshops recorded*\n\n"
    
    content += "---\n\n"
    return content

def generate_languages_section(data):
    """Generate Languages section"""
    if 'Languages' not in data or data['Languages'].empty:
        return ""
    
    languages_df = data['Languages']
    
    # Check if there are any valid language entries
    has_valid_entries = any(row.get('language') for _, row in languages_df.iterrows())
    if not has_valid_entries:
        return ""
    
    content = "## Languages\n\n"
    
    for _, row in languages_df.iterrows():
        if row.get('language'):
            language = row['language']
            proficiency = row.get('proficiency', '')
            
            content += f"**{language}**"
            if proficiency:
                content += f": {proficiency}"
            content += "\n\n"  # Changed to double newline for spacing between languages
    
    content += "---\n\n"
    return content

def generate_service_section(data):
    """Generate Academic Service section"""
    if 'Service' not in data or data['Service'].empty:
        return ""
    
    service_df = data['Service']
    
    # Check if there are any valid service entries
    has_valid_entries = any(row.get('position') for _, row in service_df.iterrows())
    if not has_valid_entries:
        return ""
    
    content = "## Academic Service\n\n"
    
    # Group by service type
    service_types = [stype for stype in service_df['service_type'].unique() if stype]
    
    for service_type in service_types:
        content += f"### {service_type}\n\n"
        
        type_df = service_df[service_df['service_type'] == service_type]
        
        for _, row in type_df.iterrows():
            if row.get('position'):
                position = row['position']
                organization = row.get('organization', '')
                start_year = row.get('start_year', '')
                end_year = row.get('end_year', '')
                description = row.get('description', '')
                
                content += f"**{position}**"
                if organization:
                    content += f", {organization}"
                
                if start_year:
                    if end_year and end_year != start_year:
                        content += f" ({format_date(start_year)}–{format_date(end_year)})"
                    else:
                        content += f" ({format_date(start_year)})"
                content += "\n"
                
                if description:
                    content += f"{description}\n"
                
                content += "\n"
    
    content += "---\n\n"
    return content

def generate_memberships_section(data):
    """Generate Professional Memberships section"""
    if 'Memberships' not in data or data['Memberships'].empty:
        return ""
    
    memberships_df = data['Memberships']
    
    # Check if there are any valid membership entries
    has_valid_entries = any(row.get('organization') for _, row in memberships_df.iterrows())
    if not has_valid_entries:
        return ""
    
    content = "## Professional Memberships\n\n"
    
    for _, row in memberships_df.iterrows():
        if row.get('organization'):
            organization = row['organization']
            start_year = row.get('start_year', '')
            end_year = row.get('end_year', '')
            membership_type = row.get('membership_type', '')
            
            content += f"**{organization}**"
            
            if membership_type:
                content += f" ({membership_type})"
            
            if start_year:
                if end_year and end_year != start_year:
                    content += f", {format_date(start_year)}–{format_date(end_year)}"
                else:
                    content += f", {format_date(start_year)}–present"
            
            content += "\n"
    
    content += "---\n\n"
    return content

def generate_cv_markdown(data):
    """Generate complete Harvard CV in markdown format from the spreadsheet data"""
    
    # Initialize CV with header
    cv_content = init_cv()
    
    # Generate each section (sections will return empty string if no data)
    cv_content += generate_header_section(data)  
    cv_content += generate_education_section(data)
    cv_content += generate_awards_section(data)
    cv_content += generate_publications_section(data)
    cv_content += generate_teaching_section(data)
    cv_content += generate_conferences_section(data)
    cv_content += generate_funded_research_section(data)
    cv_content += generate_research_section(data)
    cv_content += generate_workshops_section(data)
    cv_content += generate_languages_section(data)
    cv_content += generate_service_section(data)
    cv_content += generate_memberships_section(data)

    # Add last updated date
    current_date_last_updated = datetime.now().strftime("%B %Y")
    cv_content += f"*Last updated: {current_date_last_updated}*\n\n"
    
    return cv_content

def main():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS4TyMkL-aPWhYseDOToCruWUmoiM72tPAzGWvb_DauEtXZZxuHy3AVXFXAQ6DbEuU-T5S5yS9lt2xS/pub?output=ods"
    output_file = "cv.md"
    
    # Fetch data from Google Sheets
    print("Fetching data from Google Sheets...")
    data = fetch_sheet_data(url)
    
    if data is None:
        print("Failed to fetch data. Exiting.")
        return
    
    print("Available sheets:", list(data.keys()))
    
    # Generate CV content
    print("Generating comprehensive academic CV...")
    cv_content = generate_cv_markdown(data)
    
    # Save to file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(cv_content)
    
    print(f"Academic CV successfully generated and saved to {output_file}")
    print("\nCV includes the following sections:")
    print("1. Header")
    print("2. Education")
    print("3. Awards and Honors")
    print("4. Publications")
    print("5. Teaching Experience")
    print("6. Conference Presentations")
    print("7. Funded Research")
    print("8. Projects")
    print("9. Workshops and Professional Development")
    print("10. Languages")
    print("11. Academic Service")
    print("12. Professional Memberships")

if __name__ == "__main__":
    main()