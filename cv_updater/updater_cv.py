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
    #cv_content += "geometry: \"margin=1in\"\n"
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
        return "# Your Name\n\n*Your Title*  \n[email@example.com](mailto:email@example.com) | [Website](https://yourwebsite.com) | [ORCID](https://orcid.org/your-orcid) | [GitHub](https://github.com/yourusername)\n\n---\n\n"
    
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
    #content = ""
    
    if title:
        content += f"*{title}*  \n"
    
    # Build the contact info line
    contact_info = []
    if email:
        contact_info.append(f"[{email}](mailto:{email})")
    if website:
        contact_info.append(f"[Website]({website})")
    if orcid:
        # Format ORCID nicely: usually just the last 4 digits are shown in links
        orcid_id = orcid.split('/')[-1]  # Extract just the ID from a full URL
        contact_info.append(f"[ORCID](https://orcid.org/{orcid_id})")
    if github:
        contact_info.append(f"[GitHub]({github})")
    if linkedin:
        contact_info.append(f"[LinkedIn]({linkedin})")
    if address:
        contact_info.append(address) # Physical address usually isn't linked

    if contact_info:
        content += ' | '.join(contact_info) + "\n"
    
    # Research statement is crucial for interdisciplinary fields
    if research_statement:
        content += f"\n{research_statement}\n"
    
    content += "\n---\n\n"
    
    return content

def generate_education_section(data):
    """Generate Education section with nested graduate certificates"""
    content = "## Education\n\n"
    
    if 'Training' in data:
        training_df = data['Training']
        
        # Separate PhDs and graduate certificates
        phd_entries = []
        other_entries = []
        graduate_certificates = []
        
        for _, row in training_df.iterrows():
            if row.get('degree') and row.get('title'):
                degree = row['degree'].replace('_', ' ').lower()
                if degree == 'doctorate':
                    phd_entries.append(row)
                elif degree == 'graduate certificate':  # This should be 'graduate certificate' not 'graduate_certificate'
                    graduate_certificates.append(row)
                else:
                    other_entries.append(row)
        
        # Process PhD entries first, then nest certificates
        for row in phd_entries:
            status = row.get('status', '')
            degree = row['degree'].replace('_', ' ').title()
            title = row['title']
            university = row.get('university', '')
            thesis = row.get('thesis', '')
            advisor = row.get('advisor', '')
            advisor_link = row.get('advisor_link', '')
            
            # Format PhD title
            if "Ph.D." in title:
                display_degree = title
            else:
                display_degree = f"Ph.D. in {title.replace('Ph.D. in ', '')}"
            
            content += f"**{display_degree}**"
            if status and status.lower() != 'finished':
                content += f" ({status})"
            content += "\n"  # Line break after degree
            
            if university:
                content += f"{university}\n"  # University on its own line
            
            if thesis:
                content += f"\nThesis: *{thesis}*\n"  # Added line break before thesis
            
            if advisor:
                if advisor_link:
                    content += f"\nAdvisor: [{advisor}]({advisor_link})\n"
                else:
                    content += f"\nAdvisor: {advisor}\n"
            
            # Handle committee members
            committee_members = []
            for i in range(1, 4):  # committee_member1, committee_member2, committee_member3
                member = row.get(f'commitee_member{i}', '')  # Note: typo in column name
                member_url = row.get(f'commitee_member{i}_url', '')
                
                if member:
                    if member_url:
                        committee_members.append(f"[{member}]({member_url})")
                    else:
                        committee_members.append(member)
            
            if committee_members:
                content += f"\nCommittee: {', '.join(committee_members)}\n"
            
            # Add related graduate certificates - simple nested lines
            phd_university = row.get('university', '')
            if graduate_certificates:  # Only add line break if there are certificates
                content += "\n"  # Add line break before certificates
            for cert_row in graduate_certificates:
                cert_university = cert_row.get('university', '')
                # Match by university, or if certificate has no university, assume it belongs to this PhD
                if cert_university == phd_university or not cert_university:
                    cert_title = cert_row.get('title', '')
                    cert_status = cert_row.get('status', '')
                    if cert_title:
                        content += f"Graduate Certificate in {cert_title}"
                        if cert_status and cert_status.lower() != 'finished':
                            content += f" ({cert_status})"
                        content += "\n\n"
            
            content += "\n"  # Extra line break between entries
        
        # Process other degree entries (but NOT graduate certificates - they're already nested under PhDs)
        for row in other_entries:
            status = row.get('status', '')
            degree = row['degree'].replace('_', ' ').title()
            title = row['title']
            university = row.get('university', '')
            thesis = row.get('thesis', '')
            advisor = row.get('advisor', '')
            advisor_link = row.get('advisor_link', '')
            
            # Format degree titles
            if degree.lower() == 'masters':
                if "M.A." in title:
                    display_degree = title
                else:
                    display_degree = f"M.A. in {title.replace('M.A. in ', '')}"
            elif degree.lower() == 'bachelors':
                if "B.A." in title:
                    display_degree = title
                else:
                    display_degree = f"B.A. in {title.replace('B.A. in ', '')}"
            else:
                display_degree = f"{degree} in {title}"
            
            content += f"**{display_degree}**"
            if status and status.lower() != 'finished':
                content += f" ({status})"
            content += "\n"  # Line break after degree
            
            if university:
                content += f"{university}\n"  # University on its own line
            
            if thesis:
                content += f"\nThesis: *{thesis}*\n"  # Added line break before thesis
            
            if advisor:
                if advisor_link:
                    content += f"\nAdvisor: [{advisor}]({advisor_link})\n"  # Added line break before advisor
                else:
                    content += f"\nAdvisor: {advisor}\n"
            
            content += "\n"  # Extra line break between entries
    
    content += "---\n\n"
    return content

def generate_publications_section(data):
    """Generate Publications section"""
    content = "## Publications\n\n"
    
    if 'Publications' in data:
        publications_df = data['Publications']
        pub_count = sum(1 for _, row in publications_df.iterrows() if row.get('bibtex'))
        
        # Get year range
        years = []
        for _, row in publications_df.iterrows():
            if row.get('bibtex'):
                bibtex_entry = parse_bibtex(row['bibtex'])
                year = bibtex_entry.get('year', '')
                if year:
                    try:
                        years.append(int(year))
                    except ValueError:
                        pass
        
        year_range = ""
        if years:
            min_year = min(years)
            max_year = max(years)
            if min_year == max_year:
                year_range = f" ({min_year})"
            else:
                year_range = f" ({min_year}–{max_year})"
        
        content += f"*{pub_count} publications{year_range}*\n\n"
        
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
        
        # Journal Articles
        if journal_articles:
            content += "### Peer-Reviewed Journal Articles\n\n"
            
            for i, article in enumerate(journal_articles, 1):
                authors = format_author_name(article.get('author', ''))
                title = article.get('title', '')
                journal = article.get('journal', '')
                year = article.get('year', '')
                volume = article.get('volume', '')
                number = article.get('number', '')
                pages = article.get('pages', '')
                doi = article.get('doi', '')
                url = article.get('url', '')
                
                content += f"{i}. {authors} ({year}). "
                
                if title:
                    title_clean = title.strip('"{}')
                    if url:
                        content += f'["{title_clean}".]({url}) '
                    elif doi:
                        content += f'["{title_clean}".](https://doi.org/{doi}) '
                    else:
                        content += f'"{title_clean}". '
                
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
                
                if doi:
                    content += f" DOI: [{doi}](https://doi.org/{doi})"
                
                content += "\n\n"
        
        # Book Chapters
        if book_chapters:
            content += "### Book Chapters and Edited Volumes\n\n"
            
            for i, chapter in enumerate(book_chapters, 1):
                authors = format_author_name(chapter.get('author', ''))
                title = chapter.get('title', '')
                booktitle = chapter.get('booktitle', '')
                year = chapter.get('year', '')
                pages = chapter.get('pages', '')
                publisher = chapter.get('publisher', '')
                editor = chapter.get('editor', '')
                url = chapter.get('url', '')
                
                content += f"{i}. {authors} ({year}). "
                
                if title:
                    title_clean = title.strip('"{}')
                    if url:
                        content += f'["{title_clean}".]({url}) '
                    else:
                        content += f'"{title_clean}". '
                
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
                
                content += "\n\n"
        
        # Other Publications
        if other_publications:
            content += "### Other Publications\n\n"
            
            for i, pub in enumerate(other_publications, 1):
                authors = format_author_name(pub.get('author', ''))
                title = pub.get('title', '')
                year = pub.get('year', '')
                publisher = pub.get('publisher', '')
                url = pub.get('url', '')
                
                content += f"{i}. {authors} ({year}). "
                
                if title:
                    title_clean = title.strip('"{}')
                    if url:
                        content += f'["{title_clean}".]({url}) '
                    else:
                        content += f'"{title_clean}". '
                
                if publisher:
                    content += f"{publisher}."
                
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
    """Generate Conference Presentations section with proper Markdown line breaks"""
    content = "## Conference Presentations\n\n"
    
    if 'Conferences' in data:
        conferences_df = data['Conferences']
        
        # Count presentations
        presentation_count = len(conferences_df)
        content += f"*{presentation_count} presentations*\n\n"
        
        # Sort by year (descending)
        conferences_df = conferences_df.sort_values('year', ascending=False, na_position='last')
        
        # Group by role
        roles = [role for role in conferences_df['role'].unique() if role]
        
        for role in roles:
            content += f"### {role.title()}\n\n"
            role_df = conferences_df[conferences_df['role'] == role]
            
            for _, row in role_df.iterrows():
                # Determine title to display
                title = row.get('title', '')
                event_name = row.get('event_name', '')
                display_title = title or event_name
                if not display_title:
                    continue  # Skip if no title or event name
                
                event_theme = row.get('event_theme', '')
                conference_number = row.get('conference_number', '')
                institution = row.get('institution', '')
                city = row.get('city', '')
                country = row.get('country', '')
                month = row.get('month', '')
                year = row.get('year', '')
                co_authors = row.get('co_authors', '')
                
                # Title
                content += f"**{display_title}**  \n"  # Two spaces before newline to force line break
                
                # Event details
                if event_name and title and event_name != title:
                    content += f"{event_name}"
                    if event_theme:
                        content += f": *{event_theme}*"
                    if conference_number:
                        content += f" ({conference_number})"
                    content += "  \n"
                elif event_theme:
                    content += f"*{event_theme}*"
                    if conference_number:
                        content += f" ({conference_number})"
                    content += "  \n"
                
                # Location and date
                location_parts = [part for part in [institution, city, country] if part]
                location_str = ", ".join(location_parts)
                date_str = f"{month} {format_date(year)}" if month and year else year or ""
                
                if location_str and date_str:
                    content += f"{location_str}, {date_str}.  \n"
                elif location_str:
                    content += f"{location_str}.  \n"
                elif date_str:
                    content += f"{date_str}.  \n"
                
                # Co-authors
                if co_authors:
                    content += f"*With {co_authors}.*  \n"
                
                content += "\n"  # Extra blank line between entries
        
        content += "---\n\n"
    
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
        
        for _, row in projects_df_sorted.iterrows():
            if row.get('title'):
                short_title = row.get('short_title', '')
                title = row['title']
                status = row.get('status', '')
                start_date = row.get('start_date', '')
                end_date = row.get('end_date', '')
                description = row.get('description', '')
                url = row.get('url', '')
                technologies = row.get('technologies', '')
                collaborators = row.get('collaborators', '')
                project_type = row.get('project_type', '')
                
                display_title = short_title if short_title else title
                
                if url:
                    content += f"### [{display_title}]({url})\n\n"
                else:
                    content += f"### {display_title}\n\n"
                
                # Project details
                details = []
                if start_date:
                    date_range = format_date(start_date)
                    if end_date and end_date != start_date:
                        date_range += f"–{format_date(end_date)}"
                    details.append(f"**Years:** {date_range}")
                
                if project_type:
                    details.append(f"**Type:** {project_type}")
                
                if status:
                    details.append(f"**Status:** {status}")
                
                if details:
                    content += " | ".join(details) + "\n\n"
                
                if description:
                    content += f"{description}\n\n"
                
                if technologies:
                    content += f"**Technologies/Methods:** {technologies}\n\n"
                
                if collaborators:
                    content += f"**Collaborators:** {collaborators}\n\n"
                
                content += "---\n\n"
        
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
    cv_content += generate_publications_section(data)
    cv_content += generate_teaching_section(data)
    cv_content += generate_conferences_section(data)
    cv_content += generate_awards_section(data)
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
    print("1. Education")
    print("2. Publications") 
    print("3. Teaching Experience")
    print("4. Conference Presentations")
    print("5. Awards and Honors")
    print("6. Research Projects")
    print("7. Workshops and Professional Development")
    print("8. Languages")
    print("9. Academic Service")
    print("10. Professional Memberships")

if __name__ == "__main__":
    main()