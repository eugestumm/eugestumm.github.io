#!/usr/bin/env python3
"""
Teaching CV Generator
Reads teaching experience from an ODS spreadsheet and generates a markdown CV section.
Enhanced to properly handle Workshop category and descriptions.
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
    Format a single course entry based on available data - safe for missing values
    Enhanced to better handle Workshop entries and descriptions
    """
    try:
        # Start with course title - handle missing data safely
        entry_parts = []
        
        # Course title and code - check if values exist and are not empty
        course_code = str(row.get('course_code', '')).strip()
        course_title = str(row.get('course_title', '')).strip()
        category = str(row.get('category', '')).strip().lower()
        
        # For workshops, prioritize the title even if there's no course code
        if course_title and course_code:
            course_info = f"{course_title} ({course_code})"
        elif course_title:
            course_info = course_title
        elif course_code:
            course_info = course_code
        else:
            # For workshops, this might be more descriptive
            if category == 'workshop':
                course_info = "Workshop"
            else:
                course_info = "Course information not available"
                print(f"⚠️  Warning: Missing course title and code for entry in category: {category}")
        
        entry_parts.append(course_info)
        
        # Date information - handle safely
        date_parts = []
        semester = str(row.get('semester', '')).strip()
        year = row.get('year')
        
        if semester:
            date_parts.append(semester)
        if pd.notna(year) and year != '' and year != 0:
            try:
                date_parts.append(str(int(year)))
            except (ValueError, TypeError):
                print(f"⚠️  Warning: Invalid year format: {year}")
        
        if date_parts:
            entry_parts.append(" | ".join(date_parts))
        
        # Additional information (supervisor, co-instructor, notes) - handle safely
        additional_info = []
        
        supervisor = str(row.get('supervisor', '')).strip()
        co_instructor = str(row.get('co_instructor', '')).strip()
        special_notes = str(row.get('special_notes', '')).strip()
        description = str(row.get('description', '')).strip()
        
        # Enhanced handling of description and special_notes
        # Prioritize description column if it exists and has content
        notes_content = ""
        if description:
            notes_content = description
            print(f"ℹ️   Added description: {description[:50]}{'...' if len(description) > 50 else ''}")
        elif special_notes:
            notes_content = special_notes
            print(f"ℹ️   Added special_notes: {special_notes[:50]}{'...' if len(special_notes) > 50 else ''}")
        
        # Combine all parts with better formatting
        main_entry = " | ".join(entry_parts) if entry_parts else "Entry information incomplete"
        
        # Format additional information with line breaks for better readability
        if supervisor or co_instructor or notes_content:
            formatted_additional = []
            
            # Add supervisor and co-instructor info first
            if supervisor:
                formatted_additional.append(f"*Under supervision of {supervisor}*")
            if co_instructor:
                formatted_additional.append(f"*(co-taught with {co_instructor})*")
            
            # Add description/notes on a separate line for workshops and when content is long
            if notes_content:
                if category.lower() == 'workshop' or len(notes_content) > 50:
                    # For workshops or long descriptions, put on separate line
                    if formatted_additional:
                        return f"{main_entry}  \n  {' '.join(formatted_additional)}  \n  *{notes_content}*"
                    else:
                        return f"{main_entry}  \n  *{notes_content}*"
                else:
                    # For shorter descriptions, keep inline
                    formatted_additional.append(f"*{notes_content}*")
            
            if formatted_additional:
                return f"{main_entry}  \n  {' '.join(formatted_additional)}"
        
        return main_entry
            
    except Exception as e:
        print(f"⚠️  Warning: Error formatting course entry: {e}")
        return "Course entry could not be formatted"

def format_workshop_entry(row):
    """
    Format a workshop entry specifically for academic CV standards
    Includes institution within the entry rather than as a header
    """
    try:
        entry_parts = []
        
        # Course title and code - check if values exist and are not empty
        course_code = str(row.get('course_code', '')).strip()
        course_title = str(row.get('course_title', '')).strip()
        institution = str(row.get('institution', '')).strip()
        
        # Build the main workshop title
        if course_title and course_code:
            workshop_title = f'"{course_title}" ({course_code})'
        elif course_title:
            workshop_title = f'"{course_title}"'
        elif course_code:
            workshop_title = course_code
        else:
            workshop_title = "Workshop"
            print(f"⚠️  Warning: Missing workshop title and code")
        
        entry_parts.append(workshop_title)
        
        # Add institution if available
        if institution:
            entry_parts.append(institution)
        
        # Date information
        date_parts = []
        semester = str(row.get('semester', '')).strip()
        year = row.get('year')
        
        if semester:
            date_parts.append(semester)
        if pd.notna(year) and year != '' and year != 0:
            try:
                date_parts.append(str(int(year)))
            except (ValueError, TypeError):
                print(f"⚠️  Warning: Invalid year format: {year}")
        
        if date_parts:
            entry_parts.append(" ".join(date_parts))
        
        # Main entry line
        main_entry = " | ".join(entry_parts) if entry_parts else "Workshop information incomplete"
        
        # Additional information
        additional_lines = []
        
        supervisor = str(row.get('supervisor', '')).strip()
        co_instructor = str(row.get('co_instructor', '')).strip()
        description = str(row.get('description', '')).strip()
        special_notes = str(row.get('special_notes', '')).strip()
        
        # Build additional info lines
        supervision_info = []
        if supervisor:
            supervision_info.append(f"Under supervision of {supervisor}")
        if co_instructor:
            supervision_info.append(f"co-taught with {co_instructor}")
        
        if supervision_info:
            additional_lines.append(f"*{', '.join(supervision_info)}*")
        
        # Add description/notes
        notes_content = description if description else special_notes
        if notes_content:
            additional_lines.append(f"*{notes_content}*")
        
        # Combine everything
        if additional_lines:
            return f"{main_entry}  \n  " + "  \n  ".join(additional_lines)
        else:
            return main_entry
            
    except Exception as e:
        print(f"⚠️  Warning: Error formatting workshop entry: {e}")
        return "Workshop entry could not be formatted"
    """
    Format a single course entry based on available data - safe for missing values
    Enhanced to better handle Workshop entries and descriptions
    """
    try:
        # Start with course title - handle missing data safely
        entry_parts = []
        
        # Course title and code - check if values exist and are not empty
        course_code = str(row.get('course_code', '')).strip()
        course_title = str(row.get('course_title', '')).strip()
        category = str(row.get('category', '')).strip().lower()
        
        # For workshops, prioritize the title even if there's no course code
        if course_title and course_code:
            course_info = f"{course_title} ({course_code})"
        elif course_title:
            course_info = course_title
        elif course_code:
            course_info = course_code
        else:
            # For workshops, this might be more descriptive
            if category == 'workshop':
                course_info = "Workshop"
            else:
                course_info = "Course information not available"
                print(f"⚠️  Warning: Missing course title and code for entry in category: {category}")
        
        entry_parts.append(course_info)
        
        # Date information - handle safely
        date_parts = []
        semester = str(row.get('semester', '')).strip()
        year = row.get('year')
        
        if semester:
            date_parts.append(semester)
        if pd.notna(year) and year != '' and year != 0:
            try:
                date_parts.append(str(int(year)))
            except (ValueError, TypeError):
                print(f"⚠️  Warning: Invalid year format: {year}")
        
        if date_parts:
            entry_parts.append(" | ".join(date_parts))
        
        # Additional information (supervisor, co-instructor, notes) - handle safely
        additional_info = []
        
        supervisor = str(row.get('supervisor', '')).strip()
        co_instructor = str(row.get('co_instructor', '')).strip()
        special_notes = str(row.get('special_notes', '')).strip()
        description = str(row.get('description', '')).strip()
        
        if supervisor:
            additional_info.append(f"*Under supervision of {supervisor}*")
        
        if co_instructor:
            additional_info.append(f"*(co-taught with {co_instructor})*")
        
        # Enhanced handling of description and special_notes
        # Prioritize description column if it exists and has content
        notes_content = ""
        if description:
            notes_content = description
            print(f"ℹ️   Added description: {description[:50]}{'...' if len(description) > 50 else ''}")
        elif special_notes:
            notes_content = special_notes
            print(f"ℹ️   Added special_notes: {special_notes[:50]}{'...' if len(special_notes) > 50 else ''}")
        
        # Combine all parts with better formatting
        main_entry = " | ".join(entry_parts) if entry_parts else "Entry information incomplete"
        
        # Format additional information with line breaks for better readability
        if additional_info or notes_content:
            formatted_additional = []
            
            # Add supervisor and co-instructor info first
            if supervisor:
                formatted_additional.append(f"*Under supervision of {supervisor}*")
            if co_instructor:
                formatted_additional.append(f"*(co-taught with {co_instructor})*")
            
            # Add description/notes on a separate line for workshops and when content is long
            if notes_content:
                if category.lower() == 'workshop' or len(notes_content) > 50:
                    # For workshops or long descriptions, put on separate line
                    if formatted_additional:
                        return f"{main_entry}  \n  {' '.join(formatted_additional)}  \n  *{notes_content}*"
                    else:
                        return f"{main_entry}  \n  *{notes_content}*"
                else:
                    # For shorter descriptions, keep inline
                    formatted_additional.append(f"*{notes_content}*")
            
            if formatted_additional:
                return f"{main_entry}  \n  {' '.join(formatted_additional)}"
        
        return main_entry
            
    except Exception as e:
        print(f"⚠️  Warning: Error formatting course entry: {e}")
        return "Course entry could not be formatted"

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
                    # Sort by year (descending), then by sort_order
                    category_data = category_data.sort_values(['year', 'sort_order'], 
                                                            ascending=[False, True], 
                                                            na_position='last')
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
    Generate markdown content from grouped data - safe for empty/missing data
    Special handling for Workshop category to list chronologically without institution nesting
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
        "# Teaching",
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
            markdown_lines.append(f"## {category}")
            markdown_lines.append("")
            
            # Special handling for Workshop category - no institution nesting
            if category.lower() == 'workshop':
                category_entries = 0
                
                for _, row in data.iterrows():
                    try:
                        # Add course entry directly without institution grouping
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
                # Standard handling for other categories with institution nesting
                current_institution = None
                category_entries = 0
                
                for _, row in data.iterrows():
                    try:
                        institution = str(row.get('institution', '')).strip()
                        
                        # Add institution header if it's new and not empty
                        if institution and institution != current_institution:
                            markdown_lines.append(f"**{institution}**")
                            current_institution = institution
                        elif not institution:
                            print(f"⚠️  Warning: Missing institution for entry in category: {category}")
                        
                        # Add course entry
                        course_entry = format_course_entry(row)
                        markdown_lines.append(f"- {course_entry}")
                        category_entries += 1
                        entries_added += 1
                        
                    except Exception as e:
                        print(f"⚠️  Warning: Error processing entry in category '{category}': {e}")
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
        markdown_lines = ["# Teaching", "", "No valid teaching entries could be processed from the spreadsheet.", ""]
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
                f.write("# Teaching\n\nNo teaching data available. Please check the spreadsheet URL and format.\n")
            return
        
        print(f"✅ Successfully read {len(df)} rows from spreadsheet")
        print(f"📊 Columns found: {list(df.columns)}")
        
        # Clean and validate data
        print("🧹 Cleaning and validating data...")
        df = clean_and_validate_data(df)
        
        if df.empty:
            print("❌ No valid data found after cleaning")
            with open("teaching.md", 'w', encoding='utf-8') as f:
                f.write("# Teaching\n\nNo valid teaching entries found in spreadsheet.\n")
            return
        
        print(f"✅ Data cleaned successfully. {len(df)} valid entries found.")
        
        # Group and sort data
        print("📂 Grouping and sorting data by categories...")
        grouped_data = group_and_sort_data(df)
        
        if not grouped_data:
            print("❌ No valid categories found after grouping")
            with open("teaching.md", 'w', encoding='utf-8') as f:
                f.write("# Teaching\n\nNo valid teaching categories found in spreadsheet.\n")
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
                f.write(f"# Teaching\n\nError generating teaching CV: {e}\n\nPlease check the spreadsheet format and try again.\n")
        except Exception as file_error:
            print(f"❌ Could not even create error file: {file_error}")
    
    print("🏁 Teaching CV generation completed.")

if __name__ == "__main__":
    # Required packages: pandas, requests, odfpy
    # Install with: pip install pandas requests odfpy
    main()