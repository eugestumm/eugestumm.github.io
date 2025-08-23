#!/usr/bin/env python3
"""
Conferences CV Generator
Reads conference presentations from an ODS spreadsheet (tab: "Conferences")
and generates a Jekyll-friendly Markdown page grouped by role.

Output body has role sections as H2 (## Presenter, ## Organizer) and NO extra
"Conferences" heading in the content. Front matter is included for Jekyll.

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

# Roles to include + order (Presenter first, then Organizer)
ROLE_ORDER = ["presenter", "organizer"]


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


def _join_nonempty(parts, sep=", "):
    return sep.join([p for p in parts if p])


def format_conference_entry(row: pd.Series) -> str:
    """
    Digital Humanities–style formatting with clean line breaks:
    - **Title** (with co-authors)
    - Number + Event + : "Theme".
    - Institution, City, Country. Month Year.
    - *Special notes* (optional)
    """
    title = row.get("title", "").strip()
    co_authors = row.get("co_authors", "").strip()
    number = row.get("conference_number", "").strip()
    event = row.get("event_name", "").strip()
    theme = row.get("event_theme", "").strip()
    institution = row.get("institution", "").strip()
    city = row.get("city", "").strip()
    country = row.get("country", "").strip()
    month_disp = row.get("month_display", "").strip()
    year = row.get("year")
    notes = row.get("special_notes", "").strip()

    # Line 1: Title (+ co-authors)
    main = f"**{title}**"
    if co_authors:
        main += f" (with {co_authors})"

    # Line 2: Event string
    event_lead = _join_nonempty(
        [f"{number} {event}".strip() if number and event else (number or event)]
    , sep="")

    event_line = event_lead
    if theme:
        # Colon + italicized theme
        event_line += f': *{theme}*'
    if event_line:
        event_line += "."

    # Line 3: Venue + date
    venue_parts = _join_nonempty([institution, _join_nonempty([city, country])])
    date_parts = " ".join([p for p in [month_disp, str(int(year)) if pd.notna(year) else ""] if p])
    venue_date_line = _join_nonempty([venue_parts, date_parts], sep=". ") + "."

    # Optional notes
    if notes:
        return f"{main}  \n  {event_line}  \n  {venue_date_line}  \n  *{notes}*"
    else:
        return f"{main}  \n  {event_line}  \n  {venue_date_line}"


def generate_markdown(df: pd.DataFrame) -> str:
    """
    Build full Markdown with Jekyll front matter, then role sections as H2.
    No extra "Conferences" heading in the body.
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

        heading = role.replace("_", " ").title()  # presenter → Presenter
        lines.append(f"## {heading}")
        lines.append("")

        block = sort_block(block)
        for _, row in block.iterrows():
            entry = format_conference_entry(row)
            lines.append(f"- {entry}")
            lines.append("")  # blank line between items

    # If nothing matched ROLE_ORDER, still say something
    if lines[-1] == "":
        # remove trailing blank line
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
