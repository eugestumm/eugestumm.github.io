#!/usr/bin/env python3
"""
Projects Generator - Test Version
Can work with local CSV or Google Sheets URL
"""

import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
import os

def test_url_access(url):
    """Test if the URL is accessible and returns data"""
    try:
        print(f"🔗 Testing URL access: {url[:80]}...")
        response = requests.get(url, timeout=10)
        print(f"📊 Response status: {response.status_code}")
        print(f"📏 Content length: {len(response.content)} bytes")
        print(f"📋 Content type: {response.headers.get('content-type', 'Unknown')}")
        
        if response.status_code == 200 and len(response.content) > 0:
            return True
        else:
            return False
    except Exception as e:
        print(f"❌ URL access error: {e}")
        return False

def read_data_source():
    """Try to read data from various sources"""
    print("🔍 Checking for data sources...")
    
    # Option 1: Local CSV file
    if os.path.exists("projects.csv"):
        print("📁 Found local projects.csv file")
        try:
            df = pd.read_csv("projects.csv")
            print(f"✅ Successfully read {len(df)} rows from projects.csv")
            return df
        except Exception as e:
            print(f"❌ Error reading projects.csv: {e}")
    
    # Option 2: Get URL from user
    print("🌐 No local CSV found. Please provide Google Sheets URL.")
    print("\n📋 To get the URL:")
    print("   1. Open your Google Sheets document")
    print("   2. File → Share → Publish to web")
    print("   3. Select 'Projects' sheet and 'OpenDocument Spreadsheet (.ods)' format")
    print("   4. Click Publish and copy the URL")
    

    # Use fixed Google Sheets ODS URL
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS4TyMkL-aPWhYseDOToCruWUmoiM72tPAzGWvb_DauEtXZZxuHy3AVXFXAQ6DbEuU-T5S5yS9lt2xS/pub?output=ods"
    
    # Test URL access first
    if not test_url_access(url):
        print("❌ Cannot access the provided URL")
        return create_sample_data()
    
    # Try to read from URL
    try:
        response = requests.get(url)
        response.raise_for_status()
        
        # Try ODS first
        try:
            df = pd.read_excel(BytesIO(response.content), engine='odf', sheet_name='Projects')
            print("✅ Successfully read Projects sheet from ODS")
            return df
        except:
            print("⚠️  Projects sheet not found, trying all sheets...")
            all_sheets = pd.read_excel(BytesIO(response.content), engine='odf', sheet_name=None)
            print(f"📊 Available sheets: {list(all_sheets.keys())}")
            
            # Try first sheet
            first_sheet_name = list(all_sheets.keys())[0]
            print(f"🔄 Using first sheet: {first_sheet_name}")
            return all_sheets[first_sheet_name]
            
    except Exception as e:
        print(f"❌ Error reading from URL: {e}")
        return create_sample_data()

def create_sample_data():
    """Create sample data to test the script"""
    print("📝 Creating sample data for testing...")
    
    sample_data = {
        'project_key': ['crusader_kings_2024', 'non_binary_binary_2024', 'transylvania_witch_2024'],
        'title': [
            'Sexual Violence in Video Game Social Media: A Sentiment Analysis of Rape Speech in a Subreddit of Crusader Kings',
            'Non-Binary in Binary',
            'Transylvania Witch'
        ],
        'status': ['Ongoing', 'Completed', 'Completed'],
        'start_date': ['2024-01-01', '2023-06-01', '2023-12-01'],
        'end_date': ['', '2024-03-01', '2024-01-15'],
        'project_type': ['Computational Analysis', 'Digital Art Installation', 'Interactive Media'],
        'description': [
            'Computational analysis of discourse patterns in gaming communities using PRAW Reddit API and VADER sentiment analysis.',
            'Digital installation displaying non-binary gender identities converted to binary code, examining tensions between identity and computational logic.',
            'Interactive digital spell and participatory performance exploring gender transformation themes.'
        ],
        'methods': ['PRAW API, VADER sentiment analysis, descriptive statistics', 'Binary conversion, web installation', 'Interactive web design, performance art'],
        'technologies': ['Python, Reddit API, VADER', 'HTML, CSS, JavaScript, Tumblr', 'Web technologies, Tumblr'],
        'url': [
            'https://eugestumm.github.io/sexualviolence_crusaderkings',
            'https://non-binary-in-binary.tumblr.com/',
            'https://transylvania-witch.tumblr.com/'
        ],
        'exhibition_venues': ['', 'Museum of Contemporary Art of Rio Grande do Sul, Brazil', ''],
        'keywords': [
            'gender studies, gaming, sexual violence, sentiment analysis',
            'gender studies, non-binary, digital art, binary code',
            'transgender, digital ritual, performance art'
        ]
    }
    
    df = pd.DataFrame(sample_data)
    
    # Save as CSV for future use
    df.to_csv('projects_sample.csv', index=False)
    print("💾 Saved sample data as projects_sample.csv for future testing")
    
    return df

# Keep all the other functions from the original script
def clean_and_parse_date(date_str):
    """Parse date strings in various formats and return a datetime object"""
    if not date_str or pd.isna(date_str):
        return None
    
    date_str = str(date_str).strip()
    if not date_str:
        return None
    
    formats = ['%Y-%m-%d', '%Y-%m', '%Y', '%m/%d/%Y', '%d/%m/%Y', '%B %Y', '%b %Y']
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    return None

def process_project_entry(row):
    """Process a single project row and extract all relevant information"""
    try:
        project = {}
        
        # Basic information with safe defaults
        project['key'] = str(row.get('project_key', '')).strip()
        project['title'] = str(row.get('title', '')).strip()
        project['short_title'] = str(row.get('short_title', project['title'])).strip()
        project['status'] = str(row.get('status', 'Unknown')).strip()
        project['project_type'] = str(row.get('project_type', 'Research Project')).strip()
        
        # Dates
        project['start_date'] = clean_and_parse_date(row.get('start_date'))
        project['end_date'] = clean_and_parse_date(row.get('end_date'))
        project['last_updated'] = clean_and_parse_date(row.get('last_updated'))
        
        # Content with safe defaults
        for field in ['primary_author', 'collaborators', 'institution_affiliation', 
                     'description', 'methods', 'technologies', 'keywords', 
                     'theoretical_framework', 'url', 'additional_urls',
                     'exhibition_venues', 'exhibition_dates', 'press_coverage', 
                     'outputs', 'funding_source', 'impact_metrics']:
            project[field] = str(row.get(field, '')).strip()
        
        # Calculate sort date
        if project['last_updated']:
            project['sort_date'] = project['last_updated']
        elif project['start_date']:
            project['sort_date'] = project['start_date']
        else:
            project['sort_date'] = datetime.min
            
        return project
        
    except Exception as e:
        print(f"⚠️  Warning: Error processing project entry: {e}")
        return None

def format_project_markdown(project):
    """Format a single project as markdown"""
    try:
        markdown_parts = []
        
        # Project header as clickable link if URL is present (no emoji)
        title = project['title'] or "Untitled Project"
        if project.get('url'):
            header = f"## [{title}]({project['url']})"
        else:
            header = f"## {title}"
        markdown_parts.append(header)
        markdown_parts.append("")
        
        # Metadata
        metadata_parts = []
        if project['start_date']:
            year = project['start_date'].year
            metadata_parts.append(f"**Year**: {year}")
        if project['project_type']:
            metadata_parts.append(f"**Type**: {project['project_type']}")

        if metadata_parts:
            markdown_parts.append("; ".join(metadata_parts))
            markdown_parts.append("")
        
        # Description
        if project['description']:
            markdown_parts.append(project['description'])
            markdown_parts.append("")
        
        # Methods and technologies
        if project['methods'] or project['technologies']:
            tech_parts = []
            if project['methods']:
                tech_parts.append(f"**Methods**: {project['methods']}")
            if project['technologies']:
                tech_parts.append(f"**Technologies**: {project['technologies']}")
            markdown_parts.append("; ".join(tech_parts))
            markdown_parts.append("")
        
        # Collaborators (skip if empty, blank, or 'nan')
        collaborators = project.get('collaborators', '').strip()
        if collaborators and collaborators.lower() != 'nan':
            markdown_parts.append(f"**Collaborators**: {collaborators}")
            markdown_parts.append("")

        # Exhibition info
        if project['exhibition_venues']:
            markdown_parts.append(f"**Exhibited at**: {project['exhibition_venues']}")
            markdown_parts.append("")
        
        # Keywords
        if project['keywords']:
            markdown_parts.append(f"*Keywords*: {project['keywords']}")
            markdown_parts.append("")
        
        return "\n".join(markdown_parts)
        
    except Exception as e:
        print(f"⚠️  Warning: Error formatting project: {e}")
        return f"### {project.get('title', 'Project')} \n\n*Error formatting this project*\n\n"

def main():
    """Main function"""
    print("🚀 Projects Generator - Test Version")
    print("=" * 50)
    
    # Try to read data
    df = read_data_source()
    
    if df is None or df.empty:
        print("❌ No data available to process")
        return
    
    print(f"✅ Successfully loaded {len(df)} rows")
    print(f"📊 Columns: {list(df.columns)}")
    
    # Process projects
    projects = []
    for index, row in df.iterrows():
        project = process_project_entry(row)
        if project and project.get('title'):
            projects.append(project)

    print(f"✅ Processed {len(projects)} valid projects")

    # Sort projects by most recent first (descending by sort_date)
    projects.sort(key=lambda p: p.get('sort_date', None) or datetime.min, reverse=True)

    # Generate simple markdown
    markdown_lines = [
        "---",
        "layout: archive", 
        "title: \"Projects\"",
        "permalink: /projects/",
        "author_profile: true",
        "---",
        ""
    ]
    
    for project in projects:
        project_md = format_project_markdown(project)
        markdown_lines.append(project_md)
        markdown_lines.append("---")
        markdown_lines.append("")
    
    # Write output
    output = "\n".join(markdown_lines)
    
    try:
        with open("projects.md", 'w', encoding='utf-8') as f:
            f.write(output)
        print(f"✅ Generated projects.md successfully!")
        
        # Show preview
        print("\n📋 Preview:")
        print("-" * 50)
        print(output[:800])
        print("...")
        
    except Exception as e:
        print(f"❌ Error writing file: {e}")
        print("\n📋 Generated content:")
        print(output)

if __name__ == "__main__":
    main()