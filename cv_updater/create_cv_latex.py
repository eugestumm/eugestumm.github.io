#!/usr/bin/env python3
"""
CV PDF builder and publisher.

Run after updater_cv.py (which writes the Markdown):

  python updater_cv.py && python create_cv_latex.py

What it does:
  1. Converts each CV's Markdown to PDF with Pandoc:
         cv.md           ->  cv_euge_stumm.pdf
         cv_teaching.md  ->  cv_euge_stumm_teaching.pdf
     The look comes from pandoc/harvard_cv.css whenever a CSS-capable PDF
     engine is installed (wkhtmltopdf, WeasyPrint or Prince). LaTeX engines
     can't read CSS, so if none of those is installed the PDF is made with
     LaTeX using LATEX_STYLE below, which mirrors harvard_cv.css.
  2. Publishes the research CV to the website:
         cv_euge_stumm.pdf           ->  ../assets/
         cv.md (website-clean copy)  ->  ../_pages/   (if that folder exists)
         cv_euge_stumm.pdf           ->  ../_pages/   (if that folder exists)
     The teaching CV is not published; it's for applications.

Usage:
  python create_cv_latex.py             # both PDFs + website copy
  python create_cv_latex.py teaching    # only the teaching PDF
  python create_cv_latex.py research    # only the research PDF + website copy
"""

import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# =====================================================================
# SETTINGS
# =====================================================================

# Markdown (from updater_cv.py) -> PDF
CVS = {
    "research": ("cv.md", "cv_euge_stumm.pdf"),
    "teaching": ("cv_teaching.md", "cv_euge_stumm_teaching.pdf"),
}

# Only the research CV goes to the website.
PUBLISH = "research"

# The CV stylesheet, used by the CSS-capable engines.
CSS_FILE = Path("pandoc") / "harvard_cv.css"

PAPER_SIZE = "letter"     # US letter; use "a4" for A4
FONT_SIZE = "10pt"        # LaTeX only (the CSS sets its own font size)
PAGE_MARGIN = "0.75in"

# Tried in order; the first one that works is used. CSS engines come first
# so harvard_cv.css is used whenever possible.
CSS_ENGINES = ["wkhtmltopdf", "weasyprint", "prince"]
LATEX_ENGINES = ["pdflatex", "xelatex"]
PDF_ENGINES = CSS_ENGINES + LATEX_ENGINES

# LaTeX version of the harvard_cv.css look, added to the LaTeX preamble.
# Lists have no markers; each item starts 1.5em in and wrapped lines hang
# 1em further (a hanging indent, so long publications stay easy to tell apart).
# Each package is only used if it is installed, so a missing one never
# breaks the PDF.
LATEX_STYLE = r"""
\setlength{\parskip}{2.5pt plus 0.5pt}
\IfFileExists{mathptmx.sty}{\usepackage{mathptmx}}{}
\IfFileExists{enumitem.sty}{
  \usepackage{enumitem}
  \setlist{nosep}
  \setlist[itemize]{label={}, labelwidth=0pt, labelsep=0pt, leftmargin=2.5em, itemindent=-1em, topsep=0pt, partopsep=0pt, before=\vspace{-\parskip}}
  \setlist[itemize,2]{leftmargin=1.5em, itemindent=-1em}
  \setlist[enumerate]{leftmargin=2em, topsep=0pt, partopsep=0pt, itemsep=1.5pt}
}{}
\IfFileExists{titlesec.sty}{
  \usepackage{titlesec}
  \titleformat{\section}{\centering\Large\bfseries}{}{0pt}{}
  \titlespacing*{\section}{0pt}{0pt}{2pt}
  \titleformat{\subsection}{\normalsize\bfseries}{}{0pt}{}[\vspace{0.5pt}\titlerule\vspace{2pt}]
  \titlespacing*{\subsection}{0pt}{8pt}{3pt}
  \titleformat{\subsubsection}{\normalsize\bfseries}{}{0pt}{}
  \titlespacing*{\subsubsection}{0pt}{6.5pt}{0.5pt}
}{}
"""

# Pandoc Lua filter for the LaTeX engines: centers the title and contact
# lines under the name, and the "Last updated" line. (The CSS does the same
# for the CSS engines.) Doing it here means cv.md needs no LaTeX of its own.
CENTER_FILTER = r"""
local function walk(block, filter)
  if block.walk then return block:walk(filter) end
  return pandoc.walk_block(block, filter)
end

-- Title line (all italic) or contact line (has an email/ORCID/GitHub/LinkedIn link)
local function is_title_or_contact(block)
  if block.t ~= "Para" then return false end
  if #block.content == 1 and block.content[1].t == "Emph" then return true end
  local found = false
  walk(block, {Link = function(link)
    local t = link.target
    if t:match("^mailto:") or t:match("orcid%.org") or t:match("github%.com")
       or t:match("linkedin%.com") then found = true end
  end})
  return found
end

function Pandoc(doc)
  local blocks, out, i, done = doc.blocks, {}, 1, false
  while i <= #blocks do
    local b = blocks[i]
    out[#out + 1] = b
    i = i + 1
    if not done and b.t == "Header" and b.level == 1 then
      done = true
      local group = {}
      while i <= #blocks and is_title_or_contact(blocks[i]) do
        group[#group + 1] = blocks[i]
        i = i + 1
      end
      if #group > 0 then
        out[#out + 1] = pandoc.RawBlock("latex", "{\\centering")
        for _, g in ipairs(group) do out[#out + 1] = g end
        out[#out + 1] = pandoc.RawBlock("latex", "\\par}")
      end
    end
  end
  local last = out[#out]
  if last and last.t == "Para" and pandoc.utils.stringify(last):match("^Last updated") then
    out[#out] = pandoc.RawBlock("latex", "\\vspace{8pt}{\\centering\\small")
    out[#out + 1] = last
    out[#out + 1] = pandoc.RawBlock("latex", "\\par}")
  end
  doc.blocks = out
  return doc
end
"""


# =====================================================================
# HELPERS
# =====================================================================

def check_dependencies():
    """Pandoc and at least one LaTeX engine are required."""
    missing = []
    if shutil.which('pandoc') is None:
        missing.append("  - pandoc: Pandoc is required for document conversion")
    if not any(shutil.which(e) for e in PDF_ENGINES):
        missing.append(f"  - {' or '.join(PDF_ENGINES)}: a LaTeX engine is required for PDF generation")

    if missing:
        print("❌ Missing dependencies:")
        for dep in missing:
            print(dep)
        print("\nPlease install the missing dependencies and try again.")
        return False

    print("✅ All dependencies found")
    return True


def ensure_directory_exists(directory):
    """Create directory if it doesn't exist"""
    Path(directory).mkdir(parents=True, exist_ok=True)
    print(f"📁 Directory ensured: {directory}")


def written_by_updater(md_text):
    """True if the Markdown's front matter has the tag updater_cv.py writes."""
    if not md_text.startswith("---"):
        return False
    front_matter = md_text.split("---", 2)[1]
    return any(line.strip() == "generator: updater_cv.py" for line in front_matter.splitlines())


def pandoc_command(md_file, pdf_file, engine, style_file, filter_file):
    """The pandoc call for a CSS engine (uses harvard_cv.css) or a LaTeX
    engine (uses LATEX_STYLE and CENTER_FILTER)."""
    cmd = ['pandoc', md_file, '-o', pdf_file, '--from', 'markdown',
           f'--pdf-engine={engine}', '--variable', f'papersize={PAPER_SIZE}']
    if engine in CSS_ENGINES:
        cmd += ['--css', str(CSS_FILE),
                '--metadata', 'pagetitle=Curriculum Vitae']
        # wkhtmltopdf takes its page margins from these, not from the CSS
        for side in ('top', 'bottom', 'left', 'right'):
            cmd += ['--variable', f'margin-{side}={PAGE_MARGIN}']
    else:
        cmd += ['--include-in-header', style_file,
                '--lua-filter', filter_file,
                '--variable', f'fontsize={FONT_SIZE}',
                '--variable', f'geometry:margin={PAGE_MARGIN}',
                '--variable', 'colorlinks=true',
                '--variable', 'urlcolor=blue',
                '--variable', 'linkcolor=blue']
    return cmd


def convert_to_pdf(md_file, pdf_file):
    """Convert one CV's Markdown to PDF, with harvard_cv.css when possible."""
    print(f"\n📄 Converting {md_file} → {pdf_file}...")

    if not Path(md_file).exists():
        print(f"❌ {md_file} not found — run updater_cv.py first")
        return False

    if not written_by_updater(Path(md_file).read_text(encoding='utf-8')):
        print(f"⚠️  {md_file} was not written by the current updater_cv.py — it may come from")
        print("   an older script, so the PDF can show old formatting (e.g. numbered lists).")
        print("   Run `python updater_cv.py` first.")

    engines = [e for e in PDF_ENGINES if shutil.which(e)]
    if not CSS_FILE.exists():
        print(f"⚠️  {CSS_FILE} not found, so the CSS engines are skipped")
        engines = [e for e in engines if e not in CSS_ENGINES]
    elif not any(e in CSS_ENGINES for e in engines):
        print("ℹ️  No CSS-capable engine installed (wkhtmltopdf, WeasyPrint or Prince),")
        print("   so LaTeX is used with the matching style instead of harvard_cv.css.")

    with tempfile.NamedTemporaryFile('w', suffix='.tex', delete=False, encoding='utf-8') as f:
        f.write(LATEX_STYLE)
        style_file = f.name
    with tempfile.NamedTemporaryFile('w', suffix='.lua', delete=False, encoding='utf-8') as f:
        f.write(CENTER_FILTER)
        filter_file = f.name

    try:
        last_error = ""
        for engine in engines:
            cmd = pandoc_command(md_file, pdf_file, engine, style_file, filter_file)
            using = f"with {CSS_FILE.name}" if engine in CSS_ENGINES else "with the LaTeX style"
            print(f"🔄 Running pandoc with {engine} ({using})...")
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            except subprocess.TimeoutExpired:
                last_error = f"{engine} timed out"
                print(f"⚠️  {last_error}")
                continue

            if result.returncode == 0 and Path(pdf_file).exists():
                print(f"✅ Created {pdf_file} with {engine}")
                return True

            last_error = (result.stderr or result.stdout).strip()
            print(f"⚠️  {engine} failed, trying the next engine...")

        print(f"❌ Could not create {pdf_file}")
        if last_error:
            print(f"Last error: {last_error}")
        return False
    finally:
        Path(style_file).unlink(missing_ok=True)
        Path(filter_file).unlink(missing_ok=True)


def website_markdown(md_text):
    """cv.md without the LaTeX-only blocks (header/"Last updated" centering),
    which would otherwise show up as code on the website."""
    return re.sub(r"```\{=latex\}\n.*?\n```\n\n?", "", md_text, flags=re.S)


def copy_cv_to_assets(pdf_file):
    """Copy the research PDF to ../assets"""
    print(f"\n📋 Copying {pdf_file} to assets directory...")
    assets_dir = Path.cwd().parent / 'assets'
    ensure_directory_exists(assets_dir)

    try:
        shutil.copy2(pdf_file, assets_dir / pdf_file)
        print(f"✅ Successfully copied {pdf_file} to assets directory")
        return True
    except Exception as e:
        print(f"❌ Error copying {pdf_file} to assets: {e}")
        return False


def copy_to_pages(md_file, pdf_file):
    """Copy a website-clean cv.md and the research PDF to ../_pages if it exists"""
    pages_dir = Path.cwd().parent / '_pages'
    if not pages_dir.exists():
        print("📁 _pages directory not found, skipping copy")
        return

    if pages_dir.resolve() == Path.cwd().resolve():
        print("📁 _pages is this folder, so the files are already there")
        return

    print(f"\n📋 Copying files to {pages_dir}...")
    try:
        if Path(md_file).exists():
            text = Path(md_file).read_text(encoding='utf-8')
            (pages_dir / md_file).write_text(website_markdown(text), encoding='utf-8')
            print(f"✅ Copied {md_file} to _pages (website version)")
        if Path(pdf_file).exists():
            shutil.copy2(pdf_file, pages_dir / pdf_file)
            print(f"✅ Copied {pdf_file} to _pages")
    except Exception as e:
        print(f"❌ Error copying to _pages: {e}")


# =====================================================================
# MAIN
# =====================================================================

def main():
    """Build the PDFs, then publish the research CV."""
    print("🎓 Academic CV Generator")
    print("=" * 40)

    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    if which != "all" and which not in CVS:
        print(f"Unknown CV '{which}'. Use one of: all, {', '.join(CVS)}")
        sys.exit(1)

    if not check_dependencies():
        sys.exit(1)

    names = list(CVS) if which == "all" else [which]
    results = {name: convert_to_pdf(*CVS[name]) for name in names}

    if results.get(PUBLISH):
        md_file, pdf_file = CVS[PUBLISH]
        copy_cv_to_assets(pdf_file)
        copy_to_pages(md_file, pdf_file)

    print()
    for name, ok in results.items():
        print(f"{'✅' if ok else '❌'} {CVS[name][1]}")

    if all(results.values()):
        print("\n🎉 All tasks completed successfully!")
    else:
        print("\n❌ Some CVs could not be converted (see above).")
        sys.exit(1)


if __name__ == "__main__":
    main()