#!/usr/bin/env python3
"""
Virtual Job Fair — Employer Profile Generator
=============================================
Reads employer data from a .xlsx or .csv spreadsheet and generates
lightweight, self-contained HTML profile pages for embedding in
Matterport spaces (or any iframe-capable platform).

Usage:
    python generate_profiles.py <spreadsheet_file>

Outputs:
    profiles/               — One HTML file per employer
    employer_index.html     — Master index with all employers + file links
    employer_manifest.csv   — CSV manifest (employer name → filename)
"""

import sys
import os
import re
import csv
import html
import unicodedata
from pathlib import Path

# ---------------------------------------------------------------------------
# Optional: openpyxl for .xlsx support (falls back to csv-only if missing)
# ---------------------------------------------------------------------------
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ---------------------------------------------------------------------------
# Column name mapping — flexible matching so slight header variations work
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    "employer_name":   ["employer name", "employer", "company name", "company", "organization", "org name", "business name"],
    "website":         ["employer website", "website", "web", "url", "company website", "site"],
    "logo":            ["logo", "logo url", "logo file", "logo image", ".png logo", "logo png", "logo path", "image"],
    "phone":           ["phone", "phone number", "telephone", "tel", "contact phone"],
    "email":           ["email", "email address", "contact email", "e-mail"],
    "contact_name":    ["contact name", "contact", "contact person", "representative", "rep name", "recruiter", "recruiter name"],
    "documents":       ["documents", "document", "downloadable documents", "downloads", "files", "attachments", "resources", "document links", "docs"],
    "links":           ["links", "additional links", "extra links", "other links", "related links"],
    "description":     ["description", "about", "bio", "summary", "company description", "employer description", "overview"],
    "location":        ["location", "city", "address", "headquarters", "hq"],
    "industry":        ["industry", "sector", "field", "category"],
    "positions":       ["positions", "open positions", "jobs", "openings", "roles", "hiring for"],
}


def slugify(text: str) -> str:
    """Convert text to a clean filename-safe slug."""
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^\w\s-]", "", text.lower())
    text = re.sub(r"[-\s]+", "-", text).strip("-")
    return text or "employer"


def match_column(header: str) -> str | None:
    """Return the canonical field name for a given header, or None."""
    h = header.strip().lower()
    for canonical, aliases in COLUMN_ALIASES.items():
        if h in aliases or h == canonical:
            return canonical
    return None


def read_spreadsheet(filepath: str) -> list[dict]:
    """Read a .xlsx or .csv file and return a list of row dicts with canonical keys."""
    ext = Path(filepath).suffix.lower()

    if ext in (".xlsx", ".xlsm", ".xls") :
        if not HAS_OPENPYXL:
            print("ERROR: openpyxl is required to read .xlsx files.")
            print("Install it with:  pip install openpyxl")
            sys.exit(1)
        return _read_xlsx(filepath)
    elif ext in (".csv", ".tsv"):
        delimiter = "\t" if ext == ".tsv" else ","
        return _read_csv(filepath, delimiter)
    else:
        print(f"ERROR: Unsupported file type '{ext}'. Use .xlsx, .csv, or .tsv")
        sys.exit(1)


def _read_xlsx(filepath: str) -> list[dict]:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h else "" for h in rows[0]]
    mapping = {}
    for i, h in enumerate(headers):
        canonical = match_column(h)
        if canonical:
            mapping[i] = canonical
        else:
            # Keep unmapped columns under their original header
            mapping[i] = h.lower().strip()

    results = []
    for row in rows[1:]:
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue
        entry = {}
        for i, cell in enumerate(row):
            if i in mapping:
                entry[mapping[i]] = str(cell).strip() if cell is not None else ""
        results.append(entry)
    return results


def _read_csv(filepath: str, delimiter: str) -> list[dict]:
    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=delimiter)
        headers = next(reader, None)
        if not headers:
            return []
        mapping = {}
        for i, h in enumerate(headers):
            canonical = match_column(h)
            if canonical:
                mapping[i] = canonical
            else:
                mapping[i] = h.lower().strip()

        results = []
        for row in reader:
            if all(cell.strip() == "" for cell in row):
                continue
            entry = {}
            for i, cell in enumerate(row):
                if i in mapping:
                    entry[mapping[i]] = cell.strip()
            results.append(entry)
    return results


# ---------------------------------------------------------------------------
# HTML Template — Modern Dark Theme, fully self-contained, lightweight
# ---------------------------------------------------------------------------
def build_profile_html(emp: dict) -> str:
    """Generate a complete self-contained HTML page for one employer."""

    name = html.escape(emp.get("employer_name", "Employer"))
    website = emp.get("website", "")
    logo = emp.get("logo", "")
    phone = html.escape(emp.get("phone", ""))
    email = emp.get("email", "")
    contact = html.escape(emp.get("contact_name", ""))
    documents = emp.get("documents", "")
    links = emp.get("links", "")
    description = html.escape(emp.get("description", ""))
    location = html.escape(emp.get("location", ""))
    industry = html.escape(emp.get("industry", ""))
    positions = html.escape(emp.get("positions", ""))

    # --- Build dynamic sections ---
    logo_block = ""
    if logo:
        logo_esc = html.escape(logo)
        logo_block = f'<img class="logo" src="{logo_esc}" alt="{name} logo" onerror="this.style.display=\'none\'">'

    website_block = ""
    if website:
        w = website if website.startswith("http") else "https://" + website
        website_block = f'<a class="btn website-btn" href="{html.escape(w)}" target="_blank" rel="noopener">&#127760; Visit Website</a>'

    contact_block = ""
    contact_items = []
    if contact:
        contact_items.append(f'<div class="contact-item"><span class="label">Contact</span><span class="value">{contact}</span></div>')
    if email:
        email_esc = html.escape(email)
        contact_items.append(f'<div class="contact-item"><span class="label">Email</span><a class="value link" href="mailto:{email_esc}">{email_esc}</a></div>')
    if phone:
        phone_href = re.sub(r"[^\d+]", "", phone)
        contact_items.append(f'<div class="contact-item"><span class="label">Phone</span><a class="value link" href="tel:{phone_href}">{phone}</a></div>')
    if location:
        contact_items.append(f'<div class="contact-item"><span class="label">Location</span><span class="value">{location}</span></div>')
    if industry:
        contact_items.append(f'<div class="contact-item"><span class="label">Industry</span><span class="value">{industry}</span></div>')
    if contact_items:
        contact_block = '<div class="contact-grid">' + "\n".join(contact_items) + '</div>'

    description_block = ""
    if description:
        description_block = f'<div class="section"><h2>About Us</h2><p>{description}</p></div>'

    positions_block = ""
    if positions:
        pos_items = [html.escape(p.strip()) for p in positions.split(",") if p.strip()]
        if pos_items:
            tags = "".join(f'<span class="tag">{p}</span>' for p in pos_items)
            positions_block = f'<div class="section"><h2>Open Positions</h2><div class="tags">{tags}</div></div>'

    # Parse documents and links (semicolon or comma separated, format: "Label|URL" or just "URL")
    resource_items = []
    for field in [documents, links]:
        if not field:
            continue
        parts = re.split(r"[;\n]+", field)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if "|" in part:
                label, url = part.split("|", 1)
                label = html.escape(label.strip())
                url = url.strip()
            else:
                url = part
                label = html.escape(Path(url).stem.replace("-", " ").replace("_", " ").title() if "/" in url or "." in url else part)
            if not url.startswith("http"):
                url = "https://" + url if "." in url else url
            url_esc = html.escape(url)
            icon = "&#128196;" if any(url.lower().endswith(ext) for ext in [".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".zip"]) else "&#128279;"
            resource_items.append(f'<a class="resource-link" href="{url_esc}" target="_blank" rel="noopener">{icon} {label}</a>')

    resources_block = ""
    if resource_items:
        resources_block = '<div class="section"><h2>Resources & Documents</h2><div class="resources">' + "\n".join(resource_items) + '</div></div>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{name} — Employer Profile</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{height:100%;overflow-x:hidden}}
body{{
  font-family:'Segoe UI',system-ui,-apple-system,sans-serif;
  background:linear-gradient(145deg,#0a0e17 0%,#131a2b 50%,#0d1321 100%);
  color:#e2e8f0;
  line-height:1.6;
  padding:24px;
}}
.card{{
  max-width:520px;
  margin:0 auto;
  background:linear-gradient(160deg,rgba(30,41,66,0.95),rgba(20,28,50,0.98));
  border:1px solid rgba(99,135,210,0.18);
  border-radius:16px;
  padding:32px 28px;
  box-shadow:0 8px 32px rgba(0,0,0,0.4),0 0 60px rgba(59,93,180,0.07);
}}
.header{{text-align:center;margin-bottom:24px}}
.logo{{
  max-width:140px;max-height:80px;
  margin-bottom:16px;
  border-radius:8px;
  object-fit:contain;
  filter:drop-shadow(0 2px 8px rgba(0,0,0,0.3));
}}
.header h1{{
  font-size:1.5rem;font-weight:700;
  background:linear-gradient(135deg,#60a5fa,#a78bfa);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;
  background-clip:text;
  margin-bottom:4px;
}}
.divider{{
  height:1px;
  background:linear-gradient(90deg,transparent,rgba(99,135,210,0.3),transparent);
  margin:20px 0;
}}
.contact-grid{{display:flex;flex-direction:column;gap:10px;margin-bottom:8px}}
.contact-item{{
  display:flex;justify-content:space-between;align-items:center;
  padding:8px 12px;
  background:rgba(255,255,255,0.03);
  border-radius:8px;
  border:1px solid rgba(99,135,210,0.08);
}}
.contact-item .label{{
  font-size:0.75rem;text-transform:uppercase;letter-spacing:0.08em;
  color:#64748b;font-weight:600;
}}
.contact-item .value{{font-size:0.9rem;color:#cbd5e1}}
.link{{color:#60a5fa;text-decoration:none}}
.link:hover{{text-decoration:underline;color:#93bbfd}}
.section{{margin-top:20px}}
.section h2{{
  font-size:0.85rem;text-transform:uppercase;letter-spacing:0.1em;
  color:#64748b;margin-bottom:10px;font-weight:600;
}}
.section p{{font-size:0.92rem;color:#94a3b8;line-height:1.7}}
.tags{{display:flex;flex-wrap:wrap;gap:6px}}
.tag{{
  padding:5px 12px;border-radius:20px;font-size:0.8rem;
  background:rgba(96,165,250,0.12);color:#60a5fa;
  border:1px solid rgba(96,165,250,0.2);
}}
.resources{{display:flex;flex-direction:column;gap:8px}}
.resource-link{{
  display:block;padding:10px 14px;
  background:rgba(255,255,255,0.03);
  border:1px solid rgba(99,135,210,0.1);
  border-radius:8px;
  color:#cbd5e1;text-decoration:none;font-size:0.88rem;
  transition:all 0.2s;
}}
.resource-link:hover{{
  background:rgba(96,165,250,0.08);
  border-color:rgba(96,165,250,0.3);
  color:#60a5fa;
}}
.btn{{
  display:inline-block;padding:10px 24px;
  border-radius:8px;text-decoration:none;
  font-weight:600;font-size:0.9rem;
  transition:all 0.25s;margin-top:8px;
}}
.website-btn{{
  background:linear-gradient(135deg,#3b5cb8,#6366f1);
  color:#fff;border:none;
}}
.website-btn:hover{{
  background:linear-gradient(135deg,#4a6bd4,#818cf8);
  box-shadow:0 4px 16px rgba(99,102,241,0.35);
  transform:translateY(-1px);
}}
.footer{{
  text-align:center;margin-top:24px;
  font-size:0.7rem;color:#334155;
}}
</style>
</head>
<body>
<div class="card">
  <div class="header">
    {logo_block}
    <h1>{name}</h1>
  </div>
  {f'<div class="divider"></div>' if contact_items else ''}
  {contact_block}
  {description_block}
  {positions_block}
  {f'<div class="divider"></div>' if resource_items else ''}
  {resources_block}
  {f'<div class="divider"></div>' if website else ''}
  <div style="text-align:center">{website_block}</div>
  <div class="footer">Virtual Job Fair &mdash; Employer Profile</div>
</div>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Index page — lists all employers with links to their profiles
# ---------------------------------------------------------------------------
def build_index_html(employers: list[tuple[str, str, str]]) -> str:
    """Generate a master index page. employers = [(name, filename, logo_url), ...]"""
    rows = ""
    for i, (name, fname, logo) in enumerate(employers, 1):
        logo_img = f'<img src="{html.escape(logo)}" class="idx-logo" onerror="this.style.display=\'none\'">' if logo else '<div class="idx-logo-placeholder">&#127970;</div>'
        rows += f"""
        <a class="emp-row" href="profiles/{html.escape(fname)}" target="_blank">
          <span class="emp-num">{i:02d}</span>
          {logo_img}
          <span class="emp-name">{html.escape(name)}</span>
          <span class="emp-file">{html.escape(fname)}</span>
        </a>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Employer Profiles — Index</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{
  font-family:'Segoe UI',system-ui,-apple-system,sans-serif;
  background:linear-gradient(145deg,#0a0e17,#131a2b);
  color:#e2e8f0;padding:32px 24px;min-height:100vh;
}}
.container{{max-width:700px;margin:0 auto}}
h1{{
  font-size:1.6rem;text-align:center;margin-bottom:8px;
  background:linear-gradient(135deg,#60a5fa,#a78bfa);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;
  background-clip:text;
}}
.subtitle{{text-align:center;color:#64748b;font-size:0.9rem;margin-bottom:32px}}
.emp-row{{
  display:flex;align-items:center;gap:14px;
  padding:14px 16px;margin-bottom:8px;
  background:rgba(30,41,66,0.7);
  border:1px solid rgba(99,135,210,0.12);
  border-radius:10px;text-decoration:none;color:#e2e8f0;
  transition:all 0.2s;
}}
.emp-row:hover{{
  background:rgba(96,165,250,0.08);
  border-color:rgba(96,165,250,0.3);
  transform:translateX(4px);
}}
.emp-num{{color:#64748b;font-size:0.8rem;font-weight:600;min-width:28px}}
.idx-logo{{width:36px;height:36px;border-radius:6px;object-fit:contain}}
.idx-logo-placeholder{{
  width:36px;height:36px;border-radius:6px;
  background:rgba(99,135,210,0.12);
  display:flex;align-items:center;justify-content:center;
  font-size:1.1rem;
}}
.emp-name{{flex:1;font-weight:600;font-size:0.95rem}}
.emp-file{{color:#64748b;font-size:0.78rem;font-family:monospace}}
.stats{{
  text-align:center;margin-top:24px;padding:16px;
  background:rgba(30,41,66,0.5);border-radius:10px;
  color:#64748b;font-size:0.85rem;
}}
</style>
</head>
<body>
<div class="container">
  <h1>Virtual Job Fair — Employer Profiles</h1>
  <p class="subtitle">{len(employers)} employer profile{'' if len(employers)==1 else 's'} generated</p>
  {rows}
  <div class="stats">
    Copy each file from the <code>profiles/</code> folder into your Matterport media library.<br>
    Use the filename as the billboard media-on-click target.
  </div>
</div>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_profiles.py <spreadsheet.xlsx|.csv>")
        print("\nThis tool reads your employer spreadsheet and generates")
        print("individual HTML profile pages for your Matterport space.")
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.isfile(filepath):
        print(f"ERROR: File not found: {filepath}")
        sys.exit(1)

    print(f"\n{'='*55}")
    print("  VIRTUAL JOB FAIR — Employer Profile Generator")
    print(f"{'='*55}\n")
    print(f"Reading: {filepath}")

    employers = read_spreadsheet(filepath)
    if not employers:
        print("ERROR: No employer data found in the spreadsheet.")
        sys.exit(1)

    print(f"Found {len(employers)} employer(s)\n")

    # Create output directory
    out_dir = Path(filepath).parent / "profiles"
    out_dir.mkdir(exist_ok=True)

    index_data = []
    manifest_rows = [["Employer Name", "Filename", "Filepath"]]
    used_slugs = set()

    for emp in employers:
        name = emp.get("employer_name", "Unknown Employer")
        slug = slugify(name)

        # Ensure unique filenames
        if slug in used_slugs:
            counter = 2
            while f"{slug}-{counter}" in used_slugs:
                counter += 1
            slug = f"{slug}-{counter}"
        used_slugs.add(slug)

        filename = f"{slug}.html"
        profile_html = build_profile_html(emp)
        profile_path = out_dir / filename

        with open(profile_path, "w", encoding="utf-8") as f:
            f.write(profile_html)

        size_kb = os.path.getsize(profile_path) / 1024
        print(f"  ✓ {name:<35} → {filename}  ({size_kb:.1f} KB)")

        index_data.append((name, filename, emp.get("logo", "")))
        manifest_rows.append([name, filename, str(profile_path)])

    # Write index page
    index_path = Path(filepath).parent / "employer_index.html"
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(build_index_html(index_data))

    # Write CSV manifest
    manifest_path = Path(filepath).parent / "employer_manifest.csv"
    with open(manifest_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(manifest_rows)

    print(f"\n{'─'*55}")
    print(f"  ✓ Generated {len(employers)} profile(s) in: {out_dir}/")
    print(f"  ✓ Index page:  {index_path}")
    print(f"  ✓ Manifest:    {manifest_path}")
    print(f"{'─'*55}")
    print("\nNext steps:")
    print("  1. Upload each HTML file from profiles/ to Matterport")
    print("  2. Set each billboard's media-on-click to the uploaded file")
    print("  3. Use employer_manifest.csv to track which file = which employer\n")


if __name__ == "__main__":
    main()
