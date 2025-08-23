#!/usr/bin/env python3
"""
Teaching CV Generator - FIXED VERSION
Reads teaching experience from an ODS spreadsheet and generates a markdown CV section.
Fixed to avoid pipe characters being interpreted as table delimiters.
"""

import pandas as pd
import requests
from io import BytesIO
from collections import defaultdict

def download_and_read_ods(url):
    """
    Download ODS file from URL and read it with pandas, specifically from the "Teaching" tab
    """
    try:
        response = requests.get(url)
        response.raise_for_status()
        
        # Read ODS file from bytes, specifically the "Teaching" sheet
        print("📋 Looking for 'Teaching' sheet in the spreadsheet...")
        df = pd.read_excel(BytesIO(response.content), engine='odf', sheet_name='Teaching')
        print("✅ Successfully found and read 'Teaching' sheet")
        return df
    except Exception as e:
        print(f"❌ Error reading 'Teaching' sheet: {e}")
        print("⚠️  Attempting to read all sheets to check available sheet names...")
        
        # Try to read all sheets to show available options
        try:
            all_sheets = pd.read_excel(BytesIO(response.content), engine='odf', sheet_name=None)
            available_sheets = list(all_sheets.keys())
            print(f"📊 Available sheets in the spreadsheet: {available_sheets}")
            
            # Try some common alternatives if "Teaching" doesn't exist
            alternatives = ['teaching', 'TEACHING', 'Teaching Experience', 'Sheet1']
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

def clean_and_validate_data(df):
    """
    Clean the dataframe and handle missing values safely
    """
    if df is None or df.empty:
        print("⚠️  Warning: Spreadsheet is empty or could not be read")
        return pd.DataFrame()
    
    # Print original columns for debugging
    print(f"📊 Original columns in spreadsheet: {list(df.columns)}")
    
    # Ensure all expected columns exist
    expected_columns = ['category', 'institution', 'course_code', 'course_title', 
                       'semester', 'year', 'supervisor', 'co_instructor', 'special_notes', 'description', 'sort_order']
    
    for col in expected_columns:
        if col not in df.columns:
            df[col] = ''
            print(f"⚠️  Warning: Column '{col}' not found in spreadsheet, using empty values")
    
    # Fill NaN values with empty strings for text columns
    text_columns = ['category', 'institution', 'course_code', 'course_title', 
                   'semester', 'supervisor', 'co_instructor', 'special_notes', 'description']
    for col in text_columns:
        df[col] = df[col].fillna('').astype(str).str.strip()
    
    # Handle year column - convert to int where possible
    df['year'] = pd.to_numeric(df['year'], errors='coerce')
    
    # Handle sort_order - fill with 999 for missing values to sort them last
    df['sort_order'] = pd.to_numeric(df['sort_order'], errors='coerce').fillna(999)
    
    # Debug: Print unique categories found
    unique_categories = df['category'].dropna().str.strip().unique()
    print(f"🏷️  Categories found in data: {list(unique_categories)}")
    
    # Remove completely empty rows
    essential_cols = ['category', 'course_title']
    df = df.dropna(subset=essential_cols, how='all')
    df = df[~(df[essential_cols] == '').all(axis=1)]
    
    if df.empty:
        print("⚠️  Warning: No valid teaching entries found after cleaning data")
    else:
        print(f"✅ Found {len(df)} valid entries after cleaning")
    
    return df

def format_course_entry(row):
    """
    Format a single course entry - FIXED to avoid pipe table issues and institution mixing
    Uses alternative separators and formatting that won't trigger Markdown tables
    IMPORTANT: Does NOT include institution in the entry (institution is handled as section headers)
    """
    try:
        # Get basic information (EXCLUDING institution - that's handled separately)
        course_code = str(row.get('course_code', '')).strip()
        course_title = str(row.get('course_title', '')).strip()
        category = str(row.get('category', '')).strip().lower()
        semester = str(row.get('semester', '')).strip()
        year = row.get('year')
        supervisor = str(row.get('supervisor', '')).strip()
        co_instructor = str(row.get('co_instructor', '')).strip()
        special_notes = str(row.get('special_notes', '')).strip()
        description = str(row.get('description', '')).strip()
        
        # Build main course information using em-dash instead of pipes
        course_info_parts = []
        
        # Course title and code ONLY (no institution here!)
        if course_title and course_code:
            course_info_parts.append(f"{course_title} ({course_code})")
        elif course_title:
            course_info_parts.append(course_title)
        elif course_code:
            course_info_parts.append(course_code)
        else:
            if category == 'workshop':
                course_info_parts.append("Workshop")
            else:
                course_info_parts.append("Course information not available")
        
        # Add semester and year (handle missing semester gracefully)
        date_info = []
        if semester:
            date_info.append(semester)
        if pd.notna(year) and year != '' and year != 0:
            try:
                year_str = str(int(year))
                date_info.append(year_str)
            except (ValueError, TypeError):
                pass
        elif not semester:  # If no semester and no year, don't add empty date info
            pass
        
        if date_info:
            course_info_parts.append(" ".join(date_info))
        
        # Use em-dash as separator instead of pipe to avoid table formatting
        main_line = " — ".join(course_info_parts)
        
        # Build additional information lines
        additional_lines = []
        
        # Supervision information
        if supervisor:
            additional_lines.append(f"*Under supervision of {supervisor}*")
        
        if co_instructor:
            additional_lines.append(f"*(Co-taught with {co_instructor})*")
        
        # Description/notes (check that these don't accidentally contain institution names)
        notes_content = description if description else special_notes
        if notes_content:
            # Clean notes to avoid accidental institution inclusion
            notes_content = notes_content.strip()
            if notes_content and not notes_content.startswith('University') and not notes_content.startswith('Federal'):
                additional_lines.append(f"*{notes_content}*")
        
        # Combine everything with proper line breaks
        if additional_lines:
            return main_line + "  \n  " + "  \n  ".join(additional_lines)
        else:
            return main_line
            
    except Exception as e:
        print(f"⚠️  Warning: Error formatting course entry: {e}")
        return "Course entry could not be formatted"

def format_workshop_entry(row):
    """
    Format a workshop entry - FIXED to avoid pipe table issues
    """
    try:
        course_code = str(row.get('course_code', '')).strip()
        course_title = str(row.get('course_title', '')).strip()
        institution = str(row.get('institution', '')).strip()
        semester = str(row.get('semester', '')).strip()
        year = row.get('year')
        supervisor = str(row.get('supervisor', '')).strip()
        co_instructor = str(row.get('co_instructor', '')).strip()
        description = str(row.get('description', '')).strip()
        special_notes = str(row.get('special_notes', '')).strip()
        
        # Build main workshop information using em-dash
        workshop_parts = []
        
        # Workshop title
        if course_title and course_code:
            workshop_parts.append(f'"{course_title}" ({course_code})')
        elif course_title:
            workshop_parts.append(f'"{course_title}"')
        elif course_code:
            workshop_parts.append(course_code)
        else:
            workshop_parts.append("Workshop")
        
        # Institution
        if institution:
            workshop_parts.append(institution)
        
        # Date
        date_info = []
        if semester:
            date_info.append(semester)
        if pd.notna(year) and year != '' and year != 0:
            try:
                date_info.append(str(int(year)))
            except (ValueError, TypeError):
                pass
        
        if date_info:
            workshop_parts.append(" ".join(date_info))
        
        # Use em-dash separator
        main_line = " — ".join(workshop_parts)
        
        # Additional information
        additional_lines = []
        
        # Supervision
        supervision_parts = []
        if supervisor:
            supervision_parts.append(f"Under supervision of {supervisor}")
        if co_instructor:
            supervision_parts.append(f"co-taught with {co_instructor}")
        
        if supervision_parts:
            additional_lines.append(f"*{', '.join(supervision_parts)}*")
        
        # Description
        notes_content = description if description else special_notes
        if notes_content:
            additional_lines.append(f"*{notes_content}*")
        
        # Combine
        if additional_lines:
            return main_line + "  \n  " + "  \n  ".join(additional_lines)
        else:
            return main_line
            
    except Exception as e:
        print(f"⚠️  Warning: Error formatting workshop entry: {e}")
        return "Workshop entry could not be formatted"

def group_and_sort_data(df):
    """
    Group data by category and institution, sort appropriately - safe for empty data
    Enhanced to better handle Workshop category
    """
    if df.empty:
        print("⚠️  Warning: No data to group and sort")
        return {}
    
    # Define the desired order of categories - Workshop is included
    category_order = ['Instructor', 'Mentored Teaching', 'Teaching Assistant', 'Guest Lectures', 'Workshop', 'Tutoring']
    
    grouped_data = {}
    
    try:
        # Get unique categories, handling missing/empty values
        unique_categories = df['category'].dropna().str.strip().unique()
        unique_categories = [cat for cat in unique_categories if cat != '']
        
        print(f"🏷️  Processing categories: {list(unique_categories)}")
        
        if not unique_categories:
            print("⚠️  Warning: No valid categories found in data")
            return {}
        
        # Process predefined categories first - using case-insensitive matching
        for category in category_order:
            # Case-insensitive matching
            category_data = df[df['category'].str.strip().str.lower() == category.lower()].copy()
            if not category_data.empty:
                try:
                    # Sort by year (descending - most recent first), then by sort_order
                    # Handle NaN years by treating them as very old (0)
                    category_data['year_for_sorting'] = category_data['year'].fillna(0)
                    category_data = category_data.sort_values(['year_for_sorting', 'sort_order'], 
                                                            ascending=[False, True])
                    grouped_data[category] = category_data
                    print(f"✓ Found {len(category_data)} entries for category: {category}")
                except Exception as e:
                    print(f"⚠️  Warning: Error sorting data for category '{category}': {e}")
                    grouped_data[category] = category_data
        
        # Add any categories not in our predefined list
        processed_categories_lower = [cat.lower() for cat in category_order]
        for category in unique_categories:
            if category.lower() not in processed_categories_lower:
                try:
                    category_data = df[df['category'].str.strip().str.lower() == category.lower()].copy()
                    if not category_data.empty:
                        category_data = category_data.sort_values(['year', 'sort_order'], 
                                                                ascending=[False, True], 
                                                                na_position='last')
                        grouped_data[category.title()] = category_data
                        print(f"✓ Found {len(category_data)} entries for additional category: {category}")
                except Exception as e:
                    print(f"⚠️  Warning: Error processing category '{category}': {e}")
        
        if not grouped_data:
            print("⚠️  Warning: No valid teaching categories found after grouping")
        
    except Exception as e:
        print(f"⚠️  Warning: Error during grouping and sorting: {e}")
        return {}
    
    return grouped_data

def generate_markdown(grouped_data):
    """
    Generate markdown content from grouped data - FIXED to avoid table formatting issues
    """
    # Start with Jekyll front matter
    markdown_lines = [
        "---",
        "layout: archive",
        'title: "Teaching"',
        "permalink: /teaching/",
        "author_profile: true",
        "---",
        "",
        ""
    ]
    
    if not grouped_data:
        markdown_lines.extend([
            "No teaching experience data available.",
            "",
            "*This section will be populated when teaching data is added to the spreadsheet.*"
        ])
        print("⚠️  Warning: No teaching data available for markdown generation")
        return "\n".join(markdown_lines)
    
    entries_added = 0
    
    for category, data in grouped_data.items():
        if data.empty:
            print(f"⚠️  Warning: No entries found for category: {category}")
            continue
            
        try:
            # Add category header (NO BOLD FORMATTING - this should be a markdown header)
            markdown_lines.append(f"## {category}")
            markdown_lines.append("")
            
            # Special handling for Workshop category - no institution nesting
            if category.lower() == 'workshop':
                category_entries = 0
                
                for _, row in data.iterrows():
                    try:
                        course_entry = format_workshop_entry(row)
                        markdown_lines.append(f"- {course_entry}")
                        category_entries += 1
                        entries_added += 1
                        
                    except Exception as e:
                        print(f"⚠️  Warning: Error processing workshop entry: {e}")
                        continue
                
                if category_entries == 0:
                    print(f"⚠️  Warning: No valid entries processed for category: {category}")
                    # Remove the category header if no entries were added
                    markdown_lines = markdown_lines[:-2]
                else:
                    markdown_lines.append("")  # Add blank line after category
                    print(f"✓ Added {category_entries} workshop entries")
            
            else:
                # Fixed institution grouping - group by institution, then sort within each group
                category_entries = 0
                
                # Group by institution first
                institutions_in_category = {}
                
                for _, row in data.iterrows():
                    institution = str(row.get('institution', '')).strip()
                    if not institution:
                        institution = "Other"
                    
                    if institution not in institutions_in_category:
                        institutions_in_category[institution] = []
                    institutions_in_category[institution].append(row)
                
                # Sort institutions by the most recent course in each institution
                def get_latest_year(institution_rows):
                    years = []
                    for row in institution_rows:
                        year = row.get('year')
                        if pd.notna(year) and year != '' and year != 0:
                            years.append(float(year))
                    return max(years) if years else 0
                
                # Sort institutions by their latest year (descending)
                sorted_institutions = sorted(institutions_in_category.items(), 
                                           key=lambda x: get_latest_year(x[1]), 
                                           reverse=True)
                
                # Now process each institution group
                for institution, institution_rows in sorted_institutions:
                    # Add institution header
                    markdown_lines.append(f"**{institution}**")
                    
                    # Sort courses within this institution by year (most recent first)
                    institution_df = pd.DataFrame(institution_rows)
                    try:
                        institution_df['year_for_sorting'] = institution_df['year'].fillna(0)
                        institution_df = institution_df.sort_values(['year_for_sorting', 'sort_order'], 
                                                                   ascending=[False, True])
                    except:
                        pass
                    
                    # Add all courses for this institution
                    for _, row in institution_df.iterrows():
                        try:
                            course_entry = format_course_entry(row)
                            markdown_lines.append(f"- {course_entry}")
                            category_entries += 1
                            entries_added += 1
                        except Exception as e:
                            print(f"⚠️  Warning: Error processing entry: {e}")
                            continue
                
                if category_entries == 0:
                    print(f"⚠️  Warning: No valid entries processed for category: {category}")
                    # Remove the category header if no entries were added
                    markdown_lines = markdown_lines[:-2]
                else:
                    markdown_lines.append("")  # Add blank line after each category
                    print(f"✓ Added {category_entries} entries for category: {category}")
                
        except Exception as e:
            print(f"⚠️  Warning: Error processing category '{category}': {e}")
            continue
    
    if entries_added == 0:
        markdown_lines = ["No valid teaching entries could be processed from the spreadsheet.", ""]
        print("⚠️  Warning: No teaching entries were successfully processed")
    else:
        print(f"✓ Successfully processed {entries_added} teaching entries total")
    
    return "\n".join(markdown_lines)

def main():
    """
    Main function to generate teaching CV - safe execution with comprehensive error handling
    """
    # URL of the ODS file
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS4TyMkL-aPWhYseDOToCruWUmoiM72tPAzGWvb_DauEtXZZxuHy3AVXFXAQ6DbEuU-T5S5yS9lt2xS/pub?output=ods"
    
    print("🔄 Starting Teaching CV generation...")
    print("📥 Downloading and reading spreadsheet...")
    
    try:
        df = download_and_read_ods(url)
        
        if df is None or df.empty:
            print("❌ Failed to read spreadsheet or spreadsheet is empty")
            # Create a minimal teaching.md file
            with open("teaching.md", 'w', encoding='utf-8') as f:
                f.write("No teaching data available. Please check the spreadsheet URL and format.\n")
            return
        
        print(f"✅ Successfully read {len(df)} rows from spreadsheet")
        print(f"📊 Columns found: {list(df.columns)}")
        
        # Clean and validate data
        print("🧹 Cleaning and validating data...")
        df = clean_and_validate_data(df)
        
        if df.empty:
            print("❌ No valid data found after cleaning")
            with open("teaching.md", 'w', encoding='utf-8') as f:
                f.write("No valid teaching entries found in spreadsheet.\n")
            return
        
        print(f"✅ Data cleaned successfully. {len(df)} valid entries found.")
        
        # Group and sort data
        print("📂 Grouping and sorting data by categories...")
        grouped_data = group_and_sort_data(df)
        
        if not grouped_data:
            print("❌ No valid categories found after grouping")
            with open("teaching.md", 'w', encoding='utf-8') as f:
                f.write("No valid teaching categories found in spreadsheet.\n")
            return
        
        # Generate markdown
        print("📝 Generating markdown content...")
        markdown_content = generate_markdown(grouped_data)
        
        # Write to file
        output_file = "teaching.md"
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            print(f"✅ Successfully generated {output_file}")
            
            print(f"\n📋 Preview of generated content:")
            print("-" * 50)
            preview_length = min(500, len(markdown_content))
            print(markdown_content[:preview_length])
            if len(markdown_content) > preview_length:
                print("...")
            print("-" * 50)
            
        except Exception as e:
            print(f"❌ Error writing file: {e}")
            print("⚠️  Printing content to console instead:")
            print(markdown_content)
            
    except Exception as e:
        print(f"❌ Unexpected error during CV generation: {e}")
        print("⚠️  Creating minimal teaching.md file...")
        try:
            with open("teaching.md", 'w', encoding='utf-8') as f:
                f.write(f"Error generating teaching CV: {e}\n\nPlease check the spreadsheet format and try again.\n")
        except Exception as file_error:
            print(f"❌ Could not even create error file: {file_error}")
    
    print("🏁 Teaching CV generation completed.")

if __name__ == "__main__":
    # Required packages: pandas, requests, odfpy
    # Install with: pip install pandas requests odfpy
    main()