#!/usr/bin/env python3
"""
Conferences CV Generator
Reads conference presentations from an ODS spreadsheet (tab: "Conferences")
and generates a Jekyll-friendly Markdown page in Harvard CV style.

Output uses compact single-line format: "Title." *Conference Name*, Month Year, Location.

Required packages:
    pip install pandas requests odfpy
"""

import pandas as pd
import requests
from io import BytesIO
from calendar import month_name

# ----------------------------
# Config
# ----------------------------
SHEET_NAME = "Conferences"
URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS4TyMkL-aPWhYseDOToCruWUmoiM72tPAzGWvb_DauEtXZZxuHy3AVXFXAQ6DbEuU-T5S5yS9lt2xS/pub?output=ods"
OUTPUT_FILE = "conferences.md"

# Roles to include + order (Graduate first, then Undergraduate)
ROLE_ORDER = ["presenter_graduate", "presenter_undergraduate"]
ROLE_HEADINGS = {
    "presenter_graduate": "Graduate Presentations",
    "presenter_undergraduate": "Undergraduate Presentations"
}


# ----------------------------
# Helpers
# ----------------------------
def download_and_read_ods(url: str, sheet_name: str) -> pd.DataFrame | None:
    """Download ODS and read a specific sheet, with graceful fallbacks."""
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        return pd.read_excel(BytesIO(resp.content), engine="odf", sheet_name=sheet_name)
    except Exception as e:
        print(f"❌ Error reading '{sheet_name}' sheet: {e}")
        # Try to list available sheets for debugging
        try:
            all_sheets = pd.read_excel(BytesIO(resp.content), engine="odf", sheet_name=None)
            print(f"📊 Available sheets: {list(all_sheets.keys())}")
            if sheet_name in all_sheets:
                return all_sheets[sheet_name]
            # Fallback: first sheet
            first = next(iter(all_sheets.values()), None)
            if first is not None:
                print("🔄 Using first available sheet as a fallback.")
                return first
        except Exception as ee:
            print(f"❌ Could not read any sheets: {ee}")
    return None


def normalize_month(value) -> tuple[int, str]:
    """
    Return (month_index, month_display).
    Accepts integers 1–12, full names, 3-letter abbreviations, or free text.
    Unknown → (0, original_string_stripped or "")
    """
    if pd.isna(value):
        return 0, ""
    s = str(value).strip()
    if not s:
        return 0, ""
    # numeric?
    try:
        m = int(float(s))
        if 1 <= m <= 12:
            return m, month_name[m]
    except ValueError:
        pass
    # textual
    s_lower = s.lower()
    # map common abbreviations
    abbrev = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12
    }
    if s_lower in abbrev:
        m = abbrev[s_lower]
        return m, month_name[m]
    # full names?
    for i in range(1, 13):
        if s_lower == month_name[i].lower():
            return i, month_name[i]
    # unknown free text; keep as-is for display, 0 for sorting
    return 0, s


def clean_and_validate_data(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure expected columns exist, normalize data, and add sort helpers."""
    if df is None or df.empty:
        print("⚠️  Spreadsheet empty or unreadable.")
        return pd.DataFrame()

    expected = [
        "title", "event_name", "event_theme", "conference_number",
        "institution", "city", "country", "month", "year",
        "role", "co_authors", "special_notes", "sort_order"
    ]
    for col in expected:
        if col not in df.columns:
            df[col] = ""

    # Trim strings
    for col in expected:
        if col != "year" and col != "sort_order":
            df[col] = df[col].fillna("").astype(str).str.strip()

    # Numeric fields
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df["sort_order"] = pd.to_numeric(df["sort_order"], errors="coerce").fillna(999)

    # Normalize role to lowercase
    df["role"] = df["role"].astype(str).str.strip().str.lower()

    # Normalize month (store index for sorting and display text for rendering)
    month_idx, month_disp = [], []
    for val in df["month"]:
        idx, disp = normalize_month(val)
        month_idx.append(idx)
        month_disp.append(disp)
    df["month_index"] = month_idx
    df["month_display"] = month_disp

    # Drop rows without a title
    df = df[df["title"].astype(str).str.strip() != ""]

    # Informative log
    roles_found = sorted(df["role"].dropna().unique())
    print(f"🏷️  Roles found: {roles_found}")
    extra_roles = [r for r in roles_found if r not in ROLE_ORDER]
    if extra_roles:
        print(f"ℹ️  Note: roles not included in output (by design): {extra_roles}")

    return df


def format_conference_entry(row: pd.Series) -> str:
    """
    Harvard CV style formatting - compact single line:
    "Title." *Conference Name*, Month Year, Location. With Co-authors.
    """
    title = row.get("title", "").strip()
    event_name = row.get("event_name", "").strip()
    display_title = title or event_name
    
    if not display_title:
        return ""
    
    co_authors = row.get("co_authors", "").strip()
    event_theme = row.get("event_theme", "").strip()
    conference_number = row.get("conference_number", "").strip()
    institution = row.get("institution", "").strip()
    city = row.get("city", "").strip()
    country = row.get("country", "").strip()
    month_disp = row.get("month_display", "").strip()
    year = row.get("year")
    
    line_parts = []
    
    # Title in quotes
    line_parts.append(f'"{display_title}."')
    
    # Conference name in italics
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
    year_str = str(int(year)) if pd.notna(year) else ""
    date_str = f"{month_disp} {year_str}".strip() if month_disp and year_str else year_str
    
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
    
    # Add co-authors if present
    if co_authors:
        line_parts.append(f"With {co_authors}.")
    
    # Join all parts with single space
    return " ".join(line_parts)


def generate_markdown(df: pd.DataFrame) -> str:
    """
    Build full Markdown with Jekyll front matter, then role sections as H2.
    Uses Harvard CV style formatting.
    """
    lines = [
        "---",
        "layout: archive",
        'title: "Conferences"',
        "permalink: /conferences/",
        "author_profile: true",
        "---",
        "",
    ]

    if df.empty:
        lines.append("_No conference data available._")
        return "\n".join(lines)

    # Sort within each role: Year ↓, Month ↓, sort_order ↑
    def sort_block(block: pd.DataFrame) -> pd.DataFrame:
        return block.sort_values(
            by=["year", "month_index", "sort_order"],
            ascending=[False, False, True],
            na_position="last"
        )

    for role in ROLE_ORDER:
        block = df[df["role"] == role].copy()
        if block.empty:
            continue

        heading = ROLE_HEADINGS.get(role, role.replace("_", " ").title())
        lines.append(f"## {heading}")
        lines.append("")

        block = sort_block(block)
        for _, row in block.iterrows():
            entry = format_conference_entry(row)
            if entry:  # Only add non-empty entries
                lines.append(entry)
                lines.append("")  # blank line between items

    # Remove trailing blank line if present
    if lines and lines[-1] == "":
        lines = lines[:-1]
    
    return "\n".join(lines)


def main():
    print("🔄 Generating conferences.md ...")
    df = download_and_read_ods(URL, SHEET_NAME)
    df = clean_and_validate_data(df)

    md = generate_markdown(df)

    try:
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"✅ Successfully generated {OUTPUT_FILE}")
    except Exception as e:
        print(f"❌ Error writing file: {e}")
        print("⚠️  Printing content to stdout instead:\n")
        print(md)


if __name__ == "__main__":
    main()