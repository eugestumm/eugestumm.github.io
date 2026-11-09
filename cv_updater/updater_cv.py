"""
CV updater — writes both CVs as Markdown from the same Google Sheet:

  research CV  ->  cv.md
  teaching CV  ->  cv_teaching.md

This script only organizes the content. The look, the PDFs and the website
copy are handled by create_cv_latex.py, which you run afterwards:

  python updater_cv.py && python create_cv_latex.py

Both share all the formatting code. What differs (which sections, in what
order, and whether teaching items are highlighted) is set per CV in `CVS`
below. The teaching CV puts Teaching right after Education, adds course
development, pedagogical training, workshops and mentoring sections, and
lists teaching awards, pedagogy publications and pedagogy talks first.

Optional columns/sheets (used when present, ignored otherwise):

  Header sheet:       teaching_statement, teaching_interests, research_interests
  Teaching sheet:     level, modality, enrollment, evaluation_score,
                      evaluation_scale, description, development
  CourseDevelopment:  course_title, course_code, institution, semester, year,
                      role, level, description
  PedagogicalTraining: title, provider, month, year, end_year, hours, description
  Awards sheet:       award_type ("Teaching" marks a teaching award)
  Conferences sheet:  topic ("Teaching"/"Pedagogy" marks a pedagogy talk)
  Publications:       BibTeX keywords = {pedagogy} (or note = {pedagogy})
  Service sheet:      service_type "Mentoring" / "Advising" / "Pedagogical Training"

Usage:
  python updater_cv.py                     # both CVs
  python updater_cv.py research            # only cv.md
  python updater_cv.py teaching            # only cv_teaching.md
  python updater_cv.py --data my_cv.ods    # read a local file instead of the sheet
"""

import argparse
import re
from io import BytesIO
from datetime import datetime

import numpy as np
import pandas as pd
import requests

# =====================================================================
# SETTINGS
# =====================================================================

SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS4TyMkL-aPWhYseDOToCruWUmoiM72tPAzGWvb_DauEtXZZxuHy3AVXFXAQ6DbEuU-T5S5yS9lt2xS/pub?output=ods"

# The two CVs. For each: output file, the sections in order (reorder or
# delete freely), and whether teaching items are highlighted.
#
# teaching_focus = True:
#   - header uses teaching_statement / teaching_interests when present
#   - teaching awards, pedagogy publications and pedagogy talks come first
#     in their sections
#
# Service rows of type Thesis Supervision / Professional Development /
# Invited Workshops stay under SERVICE unless the CV includes the
# "mentoring", "pedagogical_training" or "workshops" section, where they
# are listed instead.
CVS = {
    "research": {
        "markdown": "cv.md",
        "teaching_focus": False,
        "grants_heading": "FUNDED RESEARCH",
        "sections": [
            "header",
            "education",
            "awards",
            "publications",
            "teaching",
            "conferences",
            "grants",
            "projects",
            "service",
            "memberships",
            "languages",
        ],
    },
    "teaching": {
        "markdown": "cv_teaching.md",
        "teaching_focus": True,
        "grants_heading": "GRANTS AND FELLOWSHIPS",
        "sections": [
            "header",
            "education",
            "teaching",              # workshops/guest lectures go to "workshops"
            "course_development",    # CourseDevelopment sheet + Teaching 'development' notes
            "pedagogical_training",  # PedagogicalTraining sheet + Service 'Professional Development'
            "workshops",             # Teaching 'Workshop' rows + Service 'Invited Workshops'
            "mentoring",             # Service 'Thesis Supervision' / 'Mentoring' / 'Advising'
            "awards",
            "publications",
            "conferences",
            "grants",
            "projects",
            "service",
            "memberships",
            "languages",
        ],
    },
}

# One-line overview at the top of Teaching Experience
# ("9 courses at 2 institutions, 410 students in total"). Off by default.
SHOW_TEACHING_SUMMARY = False

# Teaching categories are shown in this order (first matching pattern wins,
# lowercase regex). Categories that match nothing come last, in sheet order.
TEACHING_CATEGORY_PRIORITY = [
    r"instructor of record|primary instructor|course director|lecturer",
    r"(?<!co-)(?<!co )instructor",
    r"co-?\s?instructor|co-?teach",
    r"teaching fellow",
    r"teaching assistant|\bta\b",
    r"grader|tutor|lab",
]

# Service types listed in the mentoring / pedagogical training / workshops
# sections (and left out of SERVICE) when a CV includes those sections.
MENTORING_TYPES = ["Thesis Supervision", "Mentoring", "Advising",
                   "Student Mentoring", "Student Advising"]
PED_TRAINING_TYPES = ["Professional Development", "Pedagogical Training"]
WORKSHOP_TYPES = ["Invited Workshops", "Workshops", "Guest Lectures"]

# What counts as teaching-related when a sheet has a type/topic column.
TEACHING_PATTERN = r"teach|pedagog|sotl|classroom|curricul|instruct"


# =====================================================================
# GENERIC HELPERS
# =====================================================================

def fetch_sheet_data(source):
    """Read every sheet from a Google Sheets URL or a local .ods/.xlsx file."""
    try:
        if re.match(r"https?://", source):
            response = requests.get(source)
            response.raise_for_status()
            xl = pd.ExcelFile(BytesIO(response.content))
        else:
            xl = pd.ExcelFile(source)

        sheets_dict = {}
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            df = df.astype(object).where(pd.notnull(df), None)
            sheets_dict[sheet_name] = df
        return sheets_dict
    except Exception as e:
        print(f"Error fetching data: {e}")
        return None


def is_blank(v):
    """True for None, NaN/NaT and empty strings."""
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def as_text(v):
    """Cell value as clean text: '' for blanks, 2023.0 -> '2023'."""
    if is_blank(v):
        return ""
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


def val(row, key):
    """Text of row[key], or '' when the column is missing or the cell is blank."""
    try:
        return as_text(row.get(key))
    except AttributeError:
        return ""


def has_rows(data, sheet):
    return sheet in data and data[sheet] is not None and not data[sheet].empty


def parse_bibtex(bibtex):
    """Parse bibtex entry and return a dictionary of fields"""
    if not bibtex:
        return {}

    entry = {}

    entry_type_match = re.search(r'@(\w+){', bibtex)
    if entry_type_match:
        entry['type'] = entry_type_match.group(1)

    key_match = re.search(r'@\w+{([^,]+),', bibtex)
    if key_match:
        entry['key'] = key_match.group(1)

    field_pattern = r'(\w+)\s*=\s*{([^}]*)}'
    for field, value in re.findall(field_pattern, bibtex):
        entry[field] = value

    return entry


def format_date(year):
    """Format a year (int, float, string, date or range) for display."""
    if is_blank(year):
        return ""
    if isinstance(year, datetime):  # also covers pandas Timestamps
        return str(year.year)
    if isinstance(year, (int, np.integer)):
        return str(int(year))
    if isinstance(year, (float, np.floating)):
        return str(int(year)) if float(year).is_integer() else str(year)
    s = str(year).strip()
    if re.fullmatch(r"\d{4}\.0", s):
        s = s[:-2]
    return s.replace(' to ', '–')


def format_date_range(start_date, end_date):
    """Format date ranges for multi-year positions"""
    start = format_date(start_date)
    if not start:
        return ""
    end = format_date(end_date)
    if not end or end == start:
        return start
    return f"{start}–{end}"


def format_currency_amount(amount):
    """Format an amount that has no currency column (keeps symbols as typed)."""
    amount_str = as_text(amount)
    if not amount_str:
        return ""
    if any(sym in amount_str for sym in ('R$', 'USD', 'US$', '$', '€', '£')):
        return amount_str
    try:
        num = float(amount_str.replace(',', ''))
    except ValueError:
        return amount_str
    if ',' in amount_str or '.' in amount_str:
        return amount_str
    return f"{num:,.0f}" if num >= 1000 else amount_str


def format_money(amount, currency):
    """'R$ 1,000.00' / '$1,000.00' / '1,000.00 EUR'. Never crashes on text."""
    amount_str = as_text(amount)
    cur = as_text(currency).upper()
    try:
        num = float(amount_str.replace(',', ''))
    except ValueError:
        return f"{amount_str} {cur}".strip()
    if cur == 'BRL':
        return f"R$ {num:,.2f}"
    if cur == 'USD':
        return f"${num:,.2f}"
    return f"{num:,.2f} {cur}"


def year_key(value):
    """Numeric sort key for any year-like value; 'present' sorts newest."""
    if is_blank(value):
        return np.nan
    if isinstance(value, datetime):
        return float(value.year)
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    s = str(value).lower()
    if 'present' in s or 'current' in s or 'ongoing' in s:
        return 9999.0
    m = re.search(r"\d{4}", s)
    return float(m.group()) if m else np.nan


def semester_rank(semester):
    s = as_text(semester).lower()
    for word, rank in (('fall', 4), ('autumn', 4), ('summer', 3),
                       ('spring', 2), ('winter', 1)):
        if word in s:
            return rank
    return 0


def row_year_key(row, cols):
    keys = [year_key(row.get(c)) for c in cols]
    keys = [k for k in keys if not np.isnan(k)]
    return max(keys) if keys else np.nan


def sort_by_year(df, cols=('year',), semester_col=None):
    """Newest first (using the latest of `cols`), then Fall > Summer > Spring.
    Undated rows go last; ties keep spreadsheet order. Never raises on
    mixed types (numbers, text, blanks) the way sort_values('year') can."""
    if df is None or df.empty:
        return df
    df = df.copy()
    df['_yk'] = [row_year_key(r, cols) for _, r in df.iterrows()]
    by, asc = ['_yk'], [False]
    if semester_col and semester_col in df.columns:
        df['_sk'] = [semester_rank(v) for v in df[semester_col]]
        by.append('_sk')
        asc.append(False)
    df = df.sort_values(by, ascending=asc, na_position='last', kind='stable')
    return df.drop(columns=[c for c in ('_yk', '_sk') if c in df.columns])


def newest_first(items):
    """Sort (year_key, markdown) pairs newest first, undated last, stable."""
    def k(item):
        y = item[0]
        return (1, 0) if np.isnan(y) else (0, -y)
    return [md for _, md in sorted(items, key=k)]


def entry_block(line, details):
    """A bold entry line followed by a real bullet list of details.
    (A list must be preceded by a blank line, or pandoc merges the bullets
    into the line above as plain ' - ' text.)"""
    out = line + "\n\n"
    details = [d for d in details if d]
    if details:
        out += "".join(f"- {d}\n" for d in details) + "\n"
    return out


def matches(pattern, text):
    return bool(re.search(pattern, as_text(text).lower()))



# =====================================================================
# SERVICE-STYLE ENTRIES (shared by Service, Mentoring, Training, Workshops)
# =====================================================================

def service_date(row):
    """Date string from single_* or start_*/end_* columns."""
    single_year = row.get('single_year')
    if not is_blank(single_year):
        month, day = val(row, 'single_month'), val(row, 'single_day')
        if month:
            date_str = month
            if day:
                try:
                    date_str += f" {int(float(day))},"
                except ValueError:
                    date_str += f" {day},"
            return f"{date_str} {format_date(single_year)}"
        return format_date(single_year)

    start_year = row.get('start_year')
    if is_blank(start_year):
        return ""
    start_month, end_month = val(row, 'start_month'), val(row, 'end_month')
    sy, ey = format_date(start_year), format_date(row.get('end_year'))

    date_str = f"{start_month} {sy}" if start_month else sy
    if ey and ey != sy:
        date_str += f"–{end_month} {ey}" if end_month else f"–{ey}"
    elif end_month and end_month != start_month and ey == sy:
        date_str += f"–{end_month} {ey}"
    elif not ey and not end_month:
        date_str += "–present"

    additional_years = val(row, 'additional_years')
    if additional_years:
        date_str += f", {additional_years}"
    return date_str


def service_entry(row, mentoring=False):
    """Markdown for one Service-sheet row ('' if it has no position/role)."""
    position = val(row, 'position') or val(row, 'role')
    if not position:
        return ""

    line = f"**{val(row, 'workshop_title') or position}**"
    student = val(row, 'student_name')
    if mentoring and student:
        line += f" for {student}"
    if val(row, 'organization'):
        line += f", {val(row, 'organization')}"

    paren = [p for p in (service_date(row),
                         f"Chair: {val(row, 'chair')}" if mentoring and val(row, 'chair') else "")
             if p]
    if paren:
        line += f" ({', '.join(paren)})"

    details = []
    if mentoring and val(row, 'thesis_title'):
        details.append(f"*{val(row, 'thesis_title')}*")
    details.append(val(row, 'description'))
    return entry_block(line, details)


def service_rows_of_type(data, types):
    """Service rows whose service_type is one of `types` (case-insensitive)."""
    if not has_rows(data, 'Service') or 'service_type' not in data['Service'].columns:
        return pd.DataFrame()
    wanted = {t.lower() for t in types}
    df = data['Service']
    return df[[val(r, 'service_type').lower() in wanted for _, r in df.iterrows()]]


def service_row_key(row):
    return row_year_key(row, ('single_year', 'end_year', 'start_year'))


# =====================================================================
# CV SECTIONS
# =====================================================================

def init_cv():
    """YAML front matter. Page size, margins, fonts and spacing are set by
    create_cv_latex.py, not here."""
    cv_content = "---\n"
    cv_content += "author: Euge Stumm\n"
    # Lets create_cv_latex.py check it isn't converting an older cv.md
    cv_content += "generator: updater_cv.py\n"
    cv_content += "---\n"
    return cv_content


def generate_header_section(data, cv):
    """Name, title, contact info, interests and statement. The teaching CV
    prefers `teaching_statement`; both fall back to `research_statement`."""
    header_sheet_name = next((s for s in ('Header', 'Profile') if has_rows(data, s)), None)
    if not header_sheet_name:
        return "# Your Name\n\n*Your Title*  \n**Email:** [email@example.com](mailto:email@example.com) | **Website:** [yourwebsite.com](https://yourwebsite.com)\n\n"

    row = data[header_sheet_name].iloc[0]
    name = val(row, 'name') or 'Your Name'
    title = val(row, 'title')
    email = val(row, 'email')
    website = val(row, 'website')
    orcid = val(row, 'orcid')
    github = val(row, 'github')
    linkedin = val(row, 'linkedin')
    address = val(row, 'address')

    content = f"# {name}\n\n"
    if title:
        content += f"*{title}*\n\n"

    contact_info = []
    if email:
        contact_info.append(f"**Email:** [{email}](mailto:{email})")
    if website:
        website_clean = website.replace('https://', '').replace('http://', '').strip('/')
        website_url = website if website.startswith(('http://', 'https://')) else f"https://{website_clean}"
        contact_info.append(f"**Website:** [{website_clean}]({website_url})")
    if orcid:
        orcid_id = orcid.split('/')[-1] if 'orcid.org' in orcid else orcid
        contact_info.append(f"**ORCID:** [{orcid_id}](https://orcid.org/{orcid_id})")
    if github:
        if github.startswith('http'):
            github_clean = github.replace('https://github.com/', '').replace('http://github.com/', '').strip('/')
        else:
            github_clean = github.replace('github.com/', '').replace('@', '').strip('/')
        contact_info.append(f"**GitHub:** [@{github_clean}](https://github.com/{github_clean})")
    if linkedin:
        linkedin_clean = re.sub(r'^(https?://)?(www\.)?linkedin\.com/in/', '', linkedin)
        linkedin_clean = linkedin_clean.replace('@', '').strip('/')
        contact_info.append(f"**LinkedIn:** [{linkedin_clean}](https://linkedin.com/in/{linkedin_clean})")
    if address:
        contact_info.append(address)

    if contact_info:
        lines = [' | '.join(contact_info[i:i + 3]) for i in range(0, len(contact_info), 3)]
        content += '  \n'.join(lines) + "\n\n"

    focus = cv["teaching_focus"]
    interests = []
    if focus and val(row, 'teaching_interests'):
        interests.append(f"**Teaching Interests:** {val(row, 'teaching_interests')}")
    if val(row, 'research_interests'):
        interests.append(f"**Research Interests:** {val(row, 'research_interests')}")
    if interests:
        content += '  \n'.join(interests) + "\n\n"

    statement = (val(row, 'teaching_statement') if focus else '') or val(row, 'research_statement')
    if statement:
        content += f"{statement}\n\n"

    return content


def generate_education_section(data, cv):
    """Generate Education section with bullet lists Pandoc recognizes"""
    if not has_rows(data, "Training"):
        return ""

    current_year = datetime.now().year
    content = "## EDUCATION\n\n"
    training_df = data["Training"]

    phd_entries = []
    other_entries = []
    graduate_certificates = []

    for _, row in training_df.iterrows():
        if val(row, "degree") and val(row, "title"):
            degree = val(row, "degree").replace("_", " ").lower()
            if degree == "doctorate":
                phd_entries.append(row)
            elif degree == "graduate certificate":
                graduate_certificates.append(row)
            else:
                other_entries.append(row)

    def year_text(year):
        y = format_date(year)
        if not y:
            return ""
        try:
            return f"Expected {y}" if int(y) > current_year else y
        except ValueError:
            return y

    def format_degree_line(degree_label, title, university, year, gpa):
        line = f"**{degree_label} in {title}**"
        if year_text(year):
            line += f", {year_text(year)}"
        if university:
            line += f", {university}"
        if gpa:
            line += f", GPA: {gpa}"
        return line

    def format_committee(row):
        members = []
        for i in range(1, 6):
            member = val(row, f"committee_member{i}")
            member_url = val(row, f"committee_member{i}_url")
            if member:
                members.append(f"[{member}]({member_url})" if member_url else member)
        return members

    def advisor_line(row):
        if not val(row, "advisor"):
            return ""
        if val(row, "advisor_link"):
            return f"- Advisor: [{val(row, 'advisor')}]({val(row, 'advisor_link')})\n"
        return f"- Advisor: {val(row, 'advisor')}\n"

    for row in phd_entries:
        title = val(row, "title").replace("Ph.D. in ", "")
        content += format_degree_line("Ph.D.", title, val(row, "university"),
                                      row.get("year"), val(row, "gpa")) + "\n\n"
        if val(row, "thesis"):
            content += f"- Dissertation: *{val(row, 'thesis')}*\n"
        content += advisor_line(row)
        committee = format_committee(row)
        if committee:
            content += f"- Committee: {', '.join(committee)}\n"

        # Graduate certificates (e.g., in College Teaching) nest under the Ph.D.
        phd_university = val(row, "university")
        for cert_row in graduate_certificates:
            cert_university = val(cert_row, "university")
            if cert_university == phd_university or not cert_university:
                cert_line = f"*Graduate Certificate in {val(cert_row, 'title')}*"
                if year_text(cert_row.get("year")):
                    cert_line += f", {year_text(cert_row.get('year'))}"
                content += f"- {cert_line}\n"
        content += "\n"

    for row in other_entries:
        degree = val(row, "degree").replace("_", " ").title()
        title = val(row, "title").replace("M.A. in ", "").replace("B.A. in ", "")
        degree_label = ("M.A." if degree.lower() == "masters"
                        else "B.A." if degree.lower() == "bachelors"
                        else degree)
        content += format_degree_line(degree_label, title, val(row, "university"),
                                      row.get("year"), val(row, "gpa")) + "\n\n"
        if val(row, "thesis"):
            content += f"- Thesis: *{val(row, 'thesis')}*\n"
        content += advisor_line(row)
        committee = format_committee(row)
        if committee:
            content += f"- Committee: {', '.join(committee)}\n"
        content += "\n"

    return content + "\n\n"


# ---------------------------------------------------------------------
# Teaching
# ---------------------------------------------------------------------

def is_workshop_category(category):
    return matches(r"workshop|guest lectur", category)


def teaching_category_rank(category):
    c = as_text(category).lower()
    for i, pattern in enumerate(TEACHING_CATEGORY_PRIORITY):
        if re.search(pattern, c):
            return i
    return len(TEACHING_CATEGORY_PRIORITY)


def course_date(row):
    """'Fall 2024', '2021–2023', 'Spring 2022–2023' …"""
    year = row.get('year')
    if is_blank(year):
        return ""
    end_year = row.get('end_year')
    date = format_date_range(year, end_year) if not is_blank(end_year) else format_date(year)
    semester = val(row, 'semester')
    return f"{semester} {date}" if semester else date


def teaching_rows(data):
    """Teaching rows that have a course title, workshops excluded."""
    if not has_rows(data, 'Teaching') or 'course_title' not in data['Teaching'].columns:
        return pd.DataFrame()
    df = data['Teaching']
    keep = [bool(val(r, 'course_title')) and not is_workshop_category(r.get('category'))
            for _, r in df.iterrows()]
    return df[keep]


def teaching_summary(df):
    """'*9 courses at 2 institutions, 410 students in total.*'"""
    courses, institutions = set(), set()
    students, have_students = 0, False
    for _, r in df.iterrows():
        inst = val(r, 'institution')
        courses.add((inst, (val(r, 'course_code') or val(r, 'course_title')).lower()))
        if inst:
            institutions.add(inst)
        n = re.sub(r"[^\d.]", "", val(r, 'enrollment'))
        if n:
            try:
                students += float(n)
                have_students = True
            except ValueError:
                pass

    text = f"{len(courses)} course{'s' if len(courses) != 1 else ''}"
    if institutions:
        text += f" at {len(institutions)} institution{'s' if len(institutions) != 1 else ''}"
    if have_students:
        text += f", {int(students):,} students in total"
    return f"*{text}.*\n\n"


def format_teaching_entry(row):
    """One course: title line, details indented underneath (no numbering)."""
    out = f"- **{val(row, 'course_title')}**"
    if val(row, 'course_code'):
        out += f" ({val(row, 'course_code')})"
    if course_date(row):
        out += f" — {course_date(row)}"
    out += "\n"

    details = []

    # e.g. "Undergraduate, Online, 32 students"
    facts = [val(row, 'level'), val(row, 'modality')]
    enrollment = val(row, 'enrollment')
    if enrollment:
        facts.append(enrollment if re.search(r"[a-zA-Z]", enrollment) else f"{enrollment} students")
    facts = [f for f in facts if f]
    if facts:
        details.append(", ".join(facts))

    score = val(row, 'evaluation_score')
    if score:
        scale = val(row, 'evaluation_scale')
        details.append(f"Student evaluations: {score}/{scale}" if scale and '/' not in score
                       else f"Student evaluations: {score}")

    if val(row, 'supervisor'):
        details.append(f"Supervised by {val(row, 'supervisor')}")
    if val(row, 'co_instructor'):
        details.append(f"Co-taught with {val(row, 'co_instructor')}")
    details.append(val(row, 'description'))
    details.append(val(row, 'special_notes'))

    for d in details:
        if d:
            out += f"    - {d}\n"
    return out


def generate_teaching_section(data, cv):
    """Teaching Experience: newest institution first; within each, instructor-
    of-record categories before assistant roles; courses (unnumbered) with level,
    modality, enrollment and evaluation details when those columns exist."""
    df = teaching_rows(data)
    if df.empty:
        return ""
    df = sort_by_year(df, ('year', 'end_year'), 'semester')

    content = "## TEACHING EXPERIENCE\n\n"
    if SHOW_TEACHING_SUMMARY:
        content += teaching_summary(df)

    institutions = []
    for _, r in df.iterrows():
        inst = val(r, 'institution')
        if inst not in institutions:
            institutions.append(inst)
    institutions.sort(key=lambda i: i == "")  # rows without an institution go last

    for institution in institutions:
        inst_df = df[[val(r, 'institution') == institution for _, r in df.iterrows()]]
        if institution:
            content += f"### {institution}\n\n"

        categories = []
        for _, r in inst_df.iterrows():
            cat = val(r, 'category')
            if cat not in categories:
                categories.append(cat)
        categories.sort(key=lambda c: (c == "", teaching_category_rank(c)))

        for category in categories:
            if category:
                # Role in plain italic, so it reads as a label under the
                # institution heading (which is in bold).
                content += f"*{category}*\n\n"
            cat_df = inst_df[[val(r, 'category') == category for _, r in inst_df.iterrows()]]
            for _, row in cat_df.iterrows():
                content += format_teaching_entry(row)
            content += "\n"
        content += "\n"

    return content + "\n"


def generate_course_development_section(data, cv):
    """Courses designed or redesigned: CourseDevelopment sheet plus Teaching
    rows that have a 'development' note (each course listed once)."""
    items, seen = [], set()

    def add(row, title, note):
        inst = val(row, 'institution')
        key = (title.lower(), inst.lower())
        if key in seen:
            return
        seen.add(key)
        line = f"- **{title}**"
        if val(row, 'course_code'):
            line += f" ({val(row, 'course_code')})"
        extras = [p for p in (inst, course_date(row)) if p]
        if extras:
            line += f", {', '.join(extras)}"
        details = [note, val(row, 'level'), val(row, 'description')]
        line += "\n" + "".join(f"    - {d}\n" for d in details if d)
        items.append((row_year_key(row, ('year', 'end_year')), line))

    if has_rows(data, 'CourseDevelopment'):
        dev_df = sort_by_year(data['CourseDevelopment'], ('year', 'end_year'), 'semester')
        for _, r in dev_df.iterrows():
            title = val(r, 'course_title') or val(r, 'title')
            if title:
                add(r, title, val(r, 'role') or val(r, 'development'))

    tdf = teaching_rows(data)
    if not tdf.empty and 'development' in tdf.columns:
        for _, r in sort_by_year(tdf, ('year', 'end_year'), 'semester').iterrows():
            if val(r, 'development'):
                add(r, val(r, 'course_title'), val(r, 'development'))

    if not items:
        return ""
    return "## COURSE AND CURRICULUM DEVELOPMENT\n\n" + "".join(newest_first(items)) + "\n"


def generate_pedagogical_training_section(data, cv):
    """PedagogicalTraining sheet + Service rows of type Professional Development."""
    items = []

    if has_rows(data, 'PedagogicalTraining'):
        for _, r in data['PedagogicalTraining'].iterrows():
            title = val(r, 'title') or val(r, 'program')
            if not title:
                continue
            line = f"**{title}**"
            provider = val(r, 'provider') or val(r, 'organization') or val(r, 'institution')
            if provider:
                line += f", {provider}"
            if not is_blank(r.get('single_year')) or not is_blank(r.get('start_year')):
                date = service_date(r)
            else:
                date = format_date_range(r.get('year'), r.get('end_year'))
                if date and val(r, 'month'):
                    date = f"{val(r, 'month')} {date}"
            if date:
                line += f" ({date})"
            hours = val(r, 'hours')
            if hours and not re.search(r"[a-zA-Z]", hours):
                hours = f"{hours} hours"
            items.append((row_year_key(r, ('year', 'end_year', 'single_year', 'start_year')),
                          entry_block(line, [hours, val(r, 'description')])))

    for _, r in service_rows_of_type(data, PED_TRAINING_TYPES).iterrows():
        block = service_entry(r)
        if block:
            items.append((service_row_key(r), block))

    if not items:
        return ""
    return "## PEDAGOGICAL TRAINING AND PROFESSIONAL DEVELOPMENT\n\n" + "".join(newest_first(items)) + "\n"


def generate_workshops_section(data, cv):
    """Workshops and guest lectures: Teaching rows in a Workshop/Guest Lecture
    category (left out of the main teaching list) + Service 'Invited Workshops'."""
    items = []

    if has_rows(data, 'Teaching') and 'course_title' in data['Teaching'].columns:
        for _, r in data['Teaching'].iterrows():
            if not (val(r, 'course_title') and is_workshop_category(r.get('category'))):
                continue
            line = f"**{val(r, 'course_title')}**"
            if val(r, 'course_code'):
                line += f" ({val(r, 'course_code')})"
            if val(r, 'institution'):
                line += f", {val(r, 'institution')}"
            if course_date(r):
                line += f" ({course_date(r)})"
            details = [f"Co-led with {val(r, 'co_instructor')}" if val(r, 'co_instructor') else "",
                       val(r, 'description'), val(r, 'special_notes')]
            items.append((row_year_key(r, ('year', 'end_year')), entry_block(line, details)))

    for _, r in service_rows_of_type(data, WORKSHOP_TYPES).iterrows():
        block = service_entry(r)
        if block:
            items.append((service_row_key(r), block))

    if not items:
        return ""
    return "## WORKSHOPS AND GUEST LECTURES\n\n" + "".join(newest_first(items)) + "\n"


def generate_mentoring_section(data, cv):
    """Thesis supervision, mentoring and advising (moved out of Service)."""
    rows = service_rows_of_type(data, MENTORING_TYPES)
    if rows.empty:
        return ""

    by_type = {}
    for _, r in rows.iterrows():
        block = service_entry(r, mentoring=True)
        if block:
            by_type.setdefault(val(r, 'service_type'), []).append((service_row_key(r), block))
    if not by_type:
        return ""

    content = "## STUDENT MENTORING AND ADVISING\n\n"
    ordered = sorted(by_type, key=lambda t: [m.lower() for m in MENTORING_TYPES].index(t.lower()))
    for stype in ordered:
        if len(by_type) > 1:
            content += f"### {stype}\n\n"
        content += "".join(newest_first(by_type[stype]))
    return content + "\n"


# ---------------------------------------------------------------------
# Awards
# ---------------------------------------------------------------------

def is_teaching_award(row):
    for col in ('award_type', 'category', 'type'):
        v = val(row, col)
        if v:
            return matches(TEACHING_PATTERN, v)
    return matches(TEACHING_PATTERN, val(row, 'award_name'))


def format_award(row):
    content = f"**{val(row, 'award_name')}**"
    year = format_date(row.get('year'))
    if year:
        content += f" ({year})"
    content += "  \n"

    if val(row, 'institution'):
        content += f"*{val(row, 'institution')}*  \n"

    amount, currency = val(row, 'amount'), val(row, 'currency')
    if amount:
        if currency.upper() in ('BRL', 'USD'):
            content += f"{format_money(amount, currency)}  \n"
        elif currency:
            content += f"Amount: {amount} {currency}  \n"
        else:
            content += f"Amount: {format_money(amount, 'USD')}  \n"

    if val(row, 'description'):
        content += f"{val(row, 'description')}  \n"
    return content + "\n"


def generate_awards_section(data, cv):
    """Awards and Honors, with teaching awards in their own list first."""
    if not has_rows(data, 'Awards'):
        return ""
    awards_df = sort_by_year(data['Awards'], ('year',))
    rows = [r for _, r in awards_df.iterrows() if val(r, 'award_name')]
    if not rows:
        return ""

    if cv["teaching_focus"]:
        teaching = [r for r in rows if is_teaching_award(r)]
        other = [r for r in rows if not is_teaching_award(r)]
    else:
        teaching, other = [], rows

    content = "## AWARDS AND HONORS\n\n"
    if teaching:
        content += "### Teaching Awards\n\n" + "".join(format_award(r) for r in teaching)
        if other:
            content += "### Other Honors and Awards\n\n"
    content += "".join(format_award(r) for r in other)
    return content + "\n\n"


# ---------------------------------------------------------------------
# Publications (MLA 9)
# ---------------------------------------------------------------------

def generate_publications_section(data, cv):
    """Generate Publications section in MLA 9 style. Entries tagged as
    pedagogy (BibTeX keywords/note containing pedagogy, teaching or SoTL)
    are listed first under 'Teaching and Pedagogical Scholarship'."""
    if not has_rows(data, 'Publications'):
        return ""

    # MLA 9 lists only the first author + "et al." when there are 3 or more.
    # Set to True to list every co-author (common on CVs).
    LIST_ALL_AUTHORS = True

    def clean(s):
        """Strip BibTeX braces and stray surrounding quotes."""
        if s is None:
            return ""
        return re.sub(r'[{}]', '', str(s)).strip().strip('"').strip()

    def end_period(s):
        s = s.rstrip()
        if not s:
            return s
        return s if s.endswith(('.', '?', '!')) else s + "."

    def first_last(name):
        if ',' in name:
            last, first = [p.strip() for p in name.split(',', 1)]
        else:
            parts = name.split()
            last, first = parts[-1], ' '.join(parts[:-1])
        return first, last

    def mla_authors(author_str):
        names = [n.strip() for n in re.split(r'\s+and\s+', clean(author_str)) if n.strip()]
        if not names:
            return ""
        first, last = first_last(names[0])
        lead = f"{last}, {first}" if first else last
        others = [' '.join(p for p in first_last(n) if p) for n in names[1:]]
        if not others:
            return lead
        if len(others) == 1:
            return f"{lead}, and {others[0]}"
        if LIST_ALL_AUTHORS:
            return f"{lead}, " + ", ".join(others[:-1]) + f", and {others[-1]}"
        return f"{lead}, et al."

    def author_part(entry):
        a = mla_authors(entry.get('author', ''))
        return end_period(a) + " " if a else ""

    def mla_editors(editor_str):
        eds = []
        for ed in re.split(r'\s+and\s+', clean(editor_str)):
            ed = ed.strip()
            if ed:
                eds.append(' '.join(p for p in first_last(ed) if p))
        if not eds:
            return ""
        if len(eds) == 1:
            return eds[0]
        if len(eds) == 2:
            return f"{eds[0]} and {eds[1]}"
        return f"{eds[0]} et al."

    def mla_title(title, titleaddon='', italic=False):
        t = clean(title)
        if not t:
            return ""
        addon = clean(titleaddon)
        if italic:
            s = f"*{t}*"
            if addon:
                return f"{s} [{addon}]. "
            return (s if t.endswith(('.', '?', '!')) else s + ".") + " "
        if addon:
            return f'"{t}" [{addon}]. '
        return f'"{end_period(t)}" '

    def mla_pages(pages):
        p = clean(pages).replace('--', '–').replace('-', '–')
        if not p:
            return ""
        return f"pp. {p}" if ('–' in p or ',' in p) else f"p. {p}"

    def mla_link(entry):
        doi = re.sub(r'^https?://(dx\.)?doi\.org/', '', clean(entry.get('doi')), flags=re.I)
        url = clean(entry.get('url'))
        if doi:
            return f" [https://doi.org/{doi}](https://doi.org/{doi})."
        if url:
            return f" [{url}]({url})."
        return ""

    def note_of(entry):
        return clean(entry.get('note')).lower()

    def is_submitted(note):
        return 'submitted' in note or 'under review' in note

    def is_forthcoming(note):
        return 'accepted' in note or 'forthcoming' in note

    def is_in_prep(note):
        return 'in preparation' in note or 'in prep' in note

    def is_review(note):
        stripped = re.sub(r'under review|peer[- ]?reviewed', '', note)
        return 'review' in stripped

    def is_pedagogical(entry):
        tags = f"{clean(entry.get('keywords'))} {clean(entry.get('note'))}".lower()
        return bool(re.search(r'pedagog|sotl|teaching', tags))

    def entry_year(entry):
        m = re.search(r'\d{4}', clean(entry.get('year')))
        return int(m.group()) if m else 0

    def by_year(entries):
        return sorted(entries, key=entry_year, reverse=True)

    def format_article(e):
        note = note_of(e)
        year = clean(e.get('year'))
        out = author_part(e) + mla_title(e.get('title'), e.get('titleaddon'))
        parts = []
        if clean(e.get('journal')):
            parts.append(f"*{clean(e.get('journal'))}*")
        if clean(e.get('volume')):
            parts.append(f"vol. {clean(e.get('volume'))}")
        if clean(e.get('number')):
            parts.append(f"no. {clean(e.get('number'))}")
        if is_forthcoming(note):
            parts.append(f"**forthcoming {year}**" if year else "**forthcoming**")
        else:
            if year:
                parts.append(year)
            pages = mla_pages(e.get('pages'))
            if pages:
                parts.append(pages)
        if parts:
            out += end_period(", ".join(parts))
        return out + mla_link(e)

    def format_chapter(e):
        note = note_of(e)
        year = clean(e.get('year'))
        out = author_part(e) + mla_title(e.get('title'), e.get('titleaddon'))
        parts = []
        if clean(e.get('booktitle')):
            parts.append(f"*{clean(e.get('booktitle'))}*")
        editors = mla_editors(e.get('editor'))
        if editors:
            parts.append(f"edited by {editors}")
        if clean(e.get('volume')):
            parts.append(f"vol. {clean(e.get('volume'))}")
        if clean(e.get('publisher')):
            parts.append(clean(e.get('publisher')))
        if is_forthcoming(note):
            parts.append(f"**forthcoming {year}**" if year else "**forthcoming**")
        else:
            if year:
                parts.append(year)
            pages = mla_pages(e.get('pages'))
            if pages:
                parts.append(pages)
        if parts:
            out += end_period(", ".join(parts))
        return out + mla_link(e)

    def format_digital(e):
        out = author_part(e) + mla_title(e.get('title'), e.get('titleaddon'), italic=True)
        parts = []
        publisher = clean(e.get('organization') or e.get('publisher'))
        if publisher:
            parts.append(publisher)
        if clean(e.get('year')):
            parts.append(clean(e.get('year')))
        url = clean(e.get('url'))
        if url:
            parts.append(f"[{url}]({url})")
        if parts:
            out += ", ".join(parts) + "."
        return out

    def format_software(e):
        out = author_part(e) + mla_title(e.get('title'), italic=True).rstrip()
        details = []
        if clean(e.get('version')):
            details.append(f"Version {clean(e.get('version'))}")
        if clean(e.get('year')):
            details.append(clean(e.get('year')))
        if details:
            out += " " + ", ".join(details) + "."
        if clean(e.get('license')):
            out += f" {clean(e.get('license'))} License."
        if clean(e.get('note')):
            out += f" {end_period(clean(e.get('note')))}"
        url = clean(e.get('url'))
        repo = clean(e.get('repository'))
        if url:
            out += f" [{url}]({url})"
        if repo and repo != url:
            out += f" Code: [{repo}]({repo})"
        return out

    def format_manuscript(e):
        note = note_of(e)
        out = author_part(e) + mla_title(e.get('title'), e.get('titleaddon'))
        if is_review(note):
            out += "**[Review]**"
        elif e.get('type') in ['incollection', 'inproceedings']:
            out += "**[Book Chapter]**"
        else:
            out += "**[Peer-Reviewed Journal Article]**"
        return out

    def format_other(e):
        standalone = e.get('type') in ['book', 'phdthesis', 'mastersthesis', 'thesis', 'report']
        out = author_part(e) + mla_title(e.get('title'), e.get('titleaddon'), italic=standalone)
        parts = []
        publisher = clean(e.get('publisher') or e.get('school') or e.get('institution'))
        if publisher:
            parts.append(publisher)
        if clean(e.get('year')):
            parts.append(clean(e.get('year')))
        if parts:
            out += end_period(", ".join(parts))
        return out + mla_link(e)

    def format_any(e):
        etype = e.get('type', '').lower()
        if etype == 'article':
            return format_article(e)
        if etype in ('incollection', 'inproceedings'):
            return format_chapter(e)
        return format_other(e)

    # Classify entries
    pedagogical = []
    journal_articles, journal_articles_review = [], []
    invited_publications = []
    reviews, reviews_review = [], []
    book_chapters, book_chapters_review = [], []
    digital_scholarship, software = [], []
    in_preparation, other_publications = [], []

    for _, row in data['Publications'].iterrows():
        if not val(row, 'bibtex'):
            continue
        entry = parse_bibtex(row['bibtex'])
        note = note_of(entry)
        etype = entry.get('type', '').lower()

        if etype == 'online':
            digital_scholarship.append(entry)
        elif etype == 'software':
            software.append(entry)
        elif is_in_prep(note):
            in_preparation.append(entry)
        elif (cv["teaching_focus"] and is_pedagogical(entry)
              and not is_submitted(note) and not is_review(note)
              and etype in ('article', 'incollection', 'inproceedings', 'book')):
            pedagogical.append(entry)
        elif etype == 'article':
            if is_review(note):
                (reviews_review if is_submitted(note) else reviews).append(entry)
            elif 'invited' in note:
                invited_publications.append(entry)
            elif is_submitted(note):
                journal_articles_review.append(entry)
            else:
                journal_articles.append(entry)
        elif etype in ['incollection', 'inproceedings']:
            (book_chapters_review if is_submitted(note) else book_chapters).append(entry)
        else:
            other_publications.append(entry)

    manuscripts_under_review = journal_articles_review + book_chapters_review + reviews_review

    sections = [
        ("Teaching and Pedagogical Scholarship", pedagogical, format_any),
        ("Peer-Reviewed Journal Articles", journal_articles, format_article),
        ("Invited Publications", invited_publications, format_article),
        ("Book Chapters and Edited Volumes", book_chapters, format_chapter),
        ("Reviews", reviews, format_article),
        ("Digital Scholarship", digital_scholarship, format_digital),
        ("Software", software, format_software),
        ("Manuscripts Under Review", manuscripts_under_review, format_manuscript),
        ("Manuscripts in Preparation", in_preparation, format_manuscript),
        ("Other Publications", other_publications, format_other),
    ]

    content = "## PUBLICATIONS\n\n"
    for heading, entries, formatter in sections:
        if not entries:
            continue
        content += f"### {heading}\n\n"
        # Unnumbered; the PDF gives each entry a hanging indent instead
        for entry in by_year(entries):
            content += f"- {formatter(entry)}\n\n"
    return content + "\n\n"


# ---------------------------------------------------------------------
# Conferences
# ---------------------------------------------------------------------

def is_teaching_talk(row):
    for col in ('topic', 'category', 'focus'):
        v = val(row, col)
        if v and matches(TEACHING_PATTERN, v):
            return True
    return matches(r"pedagog|teaching|classroom", val(row, 'title'))


def format_conference(row):
    title = val(row, 'title')
    event_name = val(row, 'event_name')
    display_title = title or event_name
    if not display_title:
        return ""

    line_parts = [f'"{display_title}."']

    event_theme = val(row, 'event_theme')
    conference_number = val(row, 'conference_number')
    if event_name:
        conference_name = event_name
        if event_theme and conference_name != event_theme:
            conference_name += f": {event_theme}"
        if conference_number:
            conference_name += f" ({conference_number})"
        line_parts.append(f"*{conference_name}*,")
    elif event_theme:
        theme_name = event_theme + (f" ({conference_number})" if conference_number else "")
        line_parts.append(f"*{theme_name}*,")

    year, month = format_date(row.get('year')), val(row, 'month')
    date_str = f"{month} {year}" if month and year else year

    location_str = ", ".join(p for p in (val(row, 'institution'), val(row, 'city'),
                                          val(row, 'country')) if p)
    if date_str and location_str:
        line_parts.append(f"{date_str}, {location_str}.")
    elif date_str or location_str:
        line_parts.append(f"{date_str or location_str}.")

    if val(row, 'co_authors'):
        line_parts.append(f"With {val(row, 'co_authors')}.")

    return " ".join(line_parts) + "\n\n"


def generate_conferences_section(data, cv):
    """Conference Presentations, with teaching/pedagogy talks and invited
    talks listed before the graduate and undergraduate presentations."""
    if not has_rows(data, 'Conferences') or 'role' not in data['Conferences'].columns:
        return ""
    df = sort_by_year(data['Conferences'], ('year',))

    def role(r):
        return val(r, 'role').lower()

    listed_roles = ('presenter_graduate', 'presenter_undergraduate')
    rows = [r for _, r in df.iterrows() if role(r) in listed_roles or 'invited' in role(r)]

    def pedagogy(r):
        return cv["teaching_focus"] and is_teaching_talk(r)

    groups = [
        ("Teaching and Pedagogy Presentations", [r for r in rows if pedagogy(r)]),
        ("Invited Talks", [r for r in rows if not pedagogy(r) and 'invited' in role(r)]),
        ("Graduate Presentations", [r for r in rows if not pedagogy(r)
                                    and role(r) == 'presenter_graduate']),
        ("Undergraduate Presentations", [r for r in rows if not pedagogy(r)
                                         and role(r) == 'presenter_undergraduate']),
    ]

    content = ""
    for heading, group in groups:
        entries = "".join(format_conference(r) for r in group)
        if entries:
            content += f"### {heading}\n\n{entries}\n"
    if not content:
        return ""
    return "## CONFERENCE PRESENTATIONS\n\n" + content + "\n"


# ---------------------------------------------------------------------
# Grants, projects, service, memberships, languages
# ---------------------------------------------------------------------

def generate_funded_research_section(data, cv):
    """Grants and fellowships, under the heading set in CVS."""
    if not has_rows(data, 'FundedResearch'):
        return ""
    funded_df = sort_by_year(data['FundedResearch'], ('start_year', 'end_year'))
    rows = [r for _, r in funded_df.iterrows() if val(r, 'title')]
    if not rows:
        return ""

    content = f"## {cv.get('grants_heading', 'GRANTS AND FELLOWSHIPS')}\n\n"
    for row in rows:
        content += f"**{val(row, 'title')}**"
        dates = format_date_range(row.get('start_year'), row.get('end_year'))
        if dates:
            content += f" ({dates})"
        content += "  \n"

        if val(row, 'institution'):
            content += f"*{val(row, 'institution')}*  \n"

        amount, currency = val(row, 'amount'), val(row, 'currency')
        if amount:
            if currency:
                content += f"Amount: {format_money(amount, currency)}  \n"
            else:
                content += f"Amount: {format_currency_amount(amount)}  \n"

        if val(row, 'pi'):
            content += f"Supervisor: {val(row, 'pi')}  \n"
        content += "\n"

    return content + "\n"


def generate_research_section(data, cv):
    """Projects (Digital Scholarship projects excluded, as before)."""
    if not has_rows(data, 'Projects'):
        return ""
    df = data['Projects'].copy()
    if 'project_type' in df.columns:
        df = df[[val(r, 'project_type') != 'Digital Scholarship' for _, r in df.iterrows()]]
    df = df[[bool(val(r, 'title')) for _, r in df.iterrows()]]
    if df.empty:
        return ""
    df = sort_by_year(df, ('start_date',))

    def project_line(row):
        display_title = val(row, 'short_title') or val(row, 'title')
        line = f'"{display_title}"'
        dates = format_date_range(row.get('start_date'), row.get('end_date'))
        if dates:
            if is_blank(row.get('end_date')):
                dates += "–present"
            line += f" ({dates})"
        if val(row, 'url'):
            line += f" [{val(row, 'url')}]({val(row, 'url')})"
        return line + "\n\n"

    content = "## PROJECTS\n\n"
    types = []
    if 'project_type' in df.columns:
        for _, r in df.iterrows():
            t = val(r, 'project_type')
            if t and t not in types:
                types.append(t)

    if types:
        for project_type in types:
            content += f"### {project_type}\n\n"
            for _, row in df.iterrows():
                if val(row, 'project_type') == project_type:
                    content += project_line(row)
        untyped = [r for _, r in df.iterrows() if not val(r, 'project_type')]
        content += "".join(project_line(r) for r in untyped)
    else:
        content += "".join(project_line(r) for _, r in df.iterrows())

    return content + "\n"


def generate_service_section(data, cv):
    """Service, grouped by type. Types that have their own section in this
    CV (mentoring, pedagogical training, workshops) are left out here."""
    if not has_rows(data, 'Service'):
        return ""
    service_df = data['Service']

    if 'service_type' not in service_df.columns or not any(
            val(r, 'service_type') for _, r in service_df.iterrows()):
        blocks = [(service_row_key(r), service_entry(r, mentoring=True))
                  for _, r in service_df.iterrows()]
        blocks = [b for b in blocks if b[1]]
        return ("## SERVICE\n\n" + "".join(newest_first(blocks)) + "\n") if blocks else ""

    moved_types = []
    if "mentoring" in cv["sections"]:
        moved_types += MENTORING_TYPES
    if "pedagogical_training" in cv["sections"]:
        moved_types += PED_TRAINING_TYPES
    if "workshops" in cv["sections"]:
        moved_types += WORKSHOP_TYPES
    moved = {t.lower() for t in moved_types}
    mentoring_types = {t.lower() for t in MENTORING_TYPES}
    category_order = [
        'Thesis Supervision',
        'University Service',
        'Departmental Service',
        'Professional Development',
        'Invited Workshops',
        'Professional Service',
        'Community Engagement',
        'Editorial Service',
        'Peer Review',
    ]

    groups = {}
    for _, r in service_df.iterrows():
        stype = val(r, 'service_type')
        if not stype or stype.lower() in moved:
            continue
        block = service_entry(r, mentoring=stype.lower() in mentoring_types)
        if block:
            groups.setdefault(stype, []).append((service_row_key(r), block))
    if not groups:
        return ""

    ordered = [t for t in category_order if t in groups] + \
              [t for t in groups if t not in category_order]

    content = "## SERVICE\n\n"
    for stype in ordered:
        content += f"### {stype}\n\n" + "".join(newest_first(groups[stype])) + "\n"
    return content + "\n"


def generate_memberships_section(data, cv):
    """Professional Memberships with support for multiple periods."""
    if not has_rows(data, 'Memberships'):
        return ""

    memberships_by_org = {}
    for _, row in data['Memberships'].iterrows():
        organization = val(row, 'organization')
        if not organization:
            continue
        section = val(row, 'section')
        org = memberships_by_org.setdefault(organization, {'main': None, 'sections': {}})
        period = {'start': row.get('start_year'), 'end': row.get('end_year')}

        if section:
            if section not in org['sections']:
                org['sections'][section] = {'membership_type': val(row, 'membership_type'),
                                            'periods': [period]}
            else:
                org['sections'][section]['periods'].append(period)
        else:
            if org['main'] is None:
                org['main'] = {'membership_type': val(row, 'membership_type'),
                               'periods': [period]}
            else:
                org['main']['periods'].append(period)

    if not memberships_by_org:
        return ""

    def periods_text(details):
        out = []
        for p in details['periods']:
            start, end = format_date(p['start']), format_date(p['end'])
            if start:
                out.append(f"{start}–{end}" if end and end != start else f"{start}–present")
        return f", {', '.join(out)}" if out else ""

    content = "## PROFESSIONAL MEMBERSHIPS\n\n"
    for organization, d in memberships_by_org.items():
        content += f"**{organization}**"
        if d['main']:
            if d['main']['membership_type']:
                content += f" ({d['main']['membership_type']})"
            content += periods_text(d['main'])
        content += "\n\n"

        for section, details in d['sections'].items():
            content += f"- {section}"
            if details['membership_type']:
                content += f" ({details['membership_type']})"
            content += periods_text(details) + "\n"
        content += "\n"

    return content + "\n\n"


def generate_languages_section(data, cv):
    """Languages"""
    if not has_rows(data, 'Languages'):
        return ""
    rows = [r for _, r in data['Languages'].iterrows() if val(r, 'language')]
    if not rows:
        return ""

    content = "## LANGUAGES\n\n"
    for row in rows:
        content += f"**{val(row, 'language')}**"
        if val(row, 'proficiency'):
            content += f": {val(row, 'proficiency')}"
        content += "\n\n"
    return content + "\n\n"


# =====================================================================
# ASSEMBLY
# =====================================================================

SECTION_FUNCTIONS = {
    "header": generate_header_section,
    "education": generate_education_section,
    "teaching": generate_teaching_section,
    "course_development": generate_course_development_section,
    "pedagogical_training": generate_pedagogical_training_section,
    "workshops": generate_workshops_section,
    "mentoring": generate_mentoring_section,
    "awards": generate_awards_section,
    "publications": generate_publications_section,
    "conferences": generate_conferences_section,
    "grants": generate_funded_research_section,
    "projects": generate_research_section,
    "service": generate_service_section,
    "memberships": generate_memberships_section,
    "languages": generate_languages_section,
}


def generate_cv_markdown(data, cv):
    """Build one CV (an entry of CVS) in the order of its 'sections'."""
    cv_content = init_cv()
    for name in cv["sections"]:
        cv_content += SECTION_FUNCTIONS[name](data, cv)

    current_date_last_updated = datetime.now().strftime("%B %Y")
    cv_content += f"*Last updated: {current_date_last_updated}*\n\n"
    return cv_content


def main():
    parser = argparse.ArgumentParser(description="Write the research and/or teaching CV as Markdown.")
    parser.add_argument("which", nargs="?", default="all", choices=["all", *CVS],
                        help="which CV to write (default: all)")
    parser.add_argument("--data", default=SHEET_URL,
                        help="Google Sheets URL or a local .ods/.xlsx file")
    args = parser.parse_args()

    print("Fetching CV data...")
    data = fetch_sheet_data(args.data)
    if data is None:
        print("Failed to fetch data. Exiting.")
        return
    print("Available sheets:", list(data.keys()))

    names = list(CVS) if args.which == "all" else [args.which]
    for name in names:
        cv = CVS[name]
        print(f"\nGenerating {name} CV...")
        with open(cv["markdown"], 'w', encoding='utf-8') as f:
            f.write(generate_cv_markdown(data, cv))
        print(f"Saved {cv['markdown']}")

    print("\nNext: python create_cv_latex.py  (PDFs + website copy)")


if __name__ == "__main__":
    main()