"""
merge_to_pdf.py
---------------
Merges all NetHelp HTML files (in TOC order) into a single PDF with:
  - PDF bookmarks that mirror the TOC hierarchy exactly (collapsible nesting)
  - Working internal hyperlinks (cross-page topic links)
  - Embedded images (BMP files converted to PNG on-the-fly)
  - A clickable Table of Contents page at the front

Usage (run inside your GitHub Codespace):
    pip install weasyprint beautifulsoup4 lxml pillow
    sudo apt-get install -y libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz0b
    python merge_to_pdf.py

Adjust the CONFIG block below if your paths differ.
"""

import os
import re
import base64
import xml.etree.ElementTree as ET
from io import BytesIO
from bs4 import BeautifulSoup

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE_DIR   = "/workspaces/MDX-NetHelp-Documentation/NetHelp"
DOCS_DIR   = os.path.join(BASE_DIR, "Documents")
MEDIA_DIR  = os.path.join(BASE_DIR, "Media")
TOC_PATH   = os.path.join(BASE_DIR, "toc.xml")
OUTPUT_PDF = "/workspaces/MDX-NetHelp-Documentation/MDX_NetHelp.pdf"
# ─────────────────────────────────────────────────────────────────────────────


# ── 1. Parse toc.xml into a TREE ─────────────────────────────────────────────

class TocNode:
    def __init__(self, title, url, depth):
        self.title    = title
        self.url      = url
        self.depth    = depth
        self.anchor   = fname_to_anchor(os.path.basename(url)) if url else None
        self.children = []

    @property
    def is_section(self):
        """True = folder/group node with no HTML file of its own."""
        return self.url is None


def fname_to_anchor(fname):
    """'GAGEH.htm' → 'GAGEH'  (safe HTML id / PDF anchor)"""
    return re.sub(r'[^A-Za-z0-9_\-]', '_', os.path.splitext(fname)[0])


def parse_toc(toc_path):
    """Return list of top-level TocNode objects (each may have .children)."""
    tree = ET.parse(toc_path)
    root = tree.getroot()

    def walk(xml_node, depth):
        url      = xml_node.get('url')            # e.g. "Documents/GAGEH.htm"
        title_el = xml_node.find('title')
        title    = (title_el.text or '').strip() if title_el is not None else '(untitled)'
        node     = TocNode(title, url, depth)
        for child_xml in xml_node:
            if child_xml.tag == 'item':
                node.children.append(walk(child_xml, depth + 1))
        return node

    return [walk(child, 0) for child in root if child.tag == 'item']


def flatten(nodes):
    """Depth-first flattening of a TocNode tree."""
    result = []
    for n in nodes:
        result.append(n)
        result.extend(flatten(n.children))
    return result


# ── 2. Image helpers ──────────────────────────────────────────────────────────

def image_to_data_uri(img_path):
    """Load image (BMP or otherwise), convert to PNG, return base64 data URI."""
    if not os.path.exists(img_path):
        return None
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            buf = BytesIO()
            im.save(buf, format='PNG')
            b64 = base64.b64encode(buf.getvalue()).decode('ascii')
            return f"data:image/png;base64,{b64}"
    except Exception:
        ext  = os.path.splitext(img_path)[1].lower().lstrip('.')
        mime = {'bmp': 'image/bmp', 'png': 'image/png', 'jpg': 'image/jpeg',
                'jpeg': 'image/jpeg', 'gif': 'image/gif', 'svg': 'image/svg+xml'
                }.get(ext, 'image/png')
        with open(img_path, 'rb') as f:
            b64 = base64.b64encode(f.read()).decode('ascii')
        return f"data:{mime};base64,{b64}"


# ── 3. Process a single HTML file ─────────────────────────────────────────────

def process_html(filepath):
    """
    Read one .htm file; return cleaned inner HTML (no <html>/<head>/<body>).
      - Images:       src rewritten to embedded base64 data URIs
      - Topic links:  href="GAGEH.htm" → href="#GAGEH"
      - Scripts:      removed
    """
    if not os.path.exists(filepath):
        return f'<p><em>[Missing file: {os.path.basename(filepath)}]</em></p>'

    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        soup = BeautifulSoup(f.read(), 'lxml')

    for img in soup.find_all('img'):
        src = img.get('src', '')
        if not src.startswith('data:') and not src.startswith('http'):
            img_path = os.path.join(MEDIA_DIR, os.path.basename(src))
            data_uri = image_to_data_uri(img_path)
            img['src'] = data_uri if data_uri else ''

    for a in soup.find_all('a', href=True):
        href = a['href']
        if href.startswith('http') or href.startswith('mailto') or href.startswith('#'):
            continue
        base = os.path.basename(href.split('#')[0])
        if base.lower().endswith(('.htm', '.html')):
            frag      = '#' + href.split('#')[1] if '#' in href else ''
            a['href'] = f'#{fname_to_anchor(base)}{frag}'

    for tag in soup.find_all('script'):
        tag.decompose()

    body = soup.find('body')
    return body.decode_contents() if body else str(soup)


# ── 4. Render nodes → HTML with CSS bookmark annotations ─────────────────────
#
# WeasyPrint creates PDF bookmarks from CSS properties on heading elements:
#
#   bookmark-level  : integer 1-6 — nesting depth in bookmark panel
#   bookmark-label  : string      — text shown in bookmark panel
#   bookmark-state  : open|closed — whether the node is expanded by default
#
# We assign:
#   depth 0 → bookmark-level:1  (top-level groups, e.g. "General Information")
#   depth 1 → bookmark-level:2  (pages or sub-groups inside top-level)
#   depth 2 → bookmark-level:3  (deeper nesting, e.g. input reference subsections)
#   … capped at 6
#
# Section headers (no file) get a styled heading + their children follow.
# Leaf pages (with file) get their content injected after the bookmark heading.

def bml(depth):
    return min(depth + 1, 6)


def safe_label(title):
    """Escape double-quotes for use inside CSS string value."""
    return title.replace('"', '\\"')


def render_node(node, out):
    """
    Recursively emit HTML for one TocNode into the list `out`.
    """
    level = bml(node.depth)
    label = safe_label(node.title)
    esc   = node.title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    if node.is_section:
        # ── Section header ── parent bookmark, no page content of its own
        sec_id = "section_" + re.sub(r'[^A-Za-z0-9_]', '_', node.title)
        out.append(f'<div class="section-block" id="{sec_id}">')
        out.append(
            f'<h{level} class="bm-heading bm-section" '
            f'style="bookmark-level:{level}; bookmark-label:\\"{label}\\"; bookmark-state:open">'
            f'{esc}</h{level}>'
        )
        for child in node.children:
            render_node(child, out)
        out.append('</div>')

    else:
        # ── Content page ── leaf bookmark with real HTML content
        out.append(f'<div class="page-block" id="{node.anchor}">')
        out.append(
            f'<h{level} class="bm-heading bm-page" '
            f'style="bookmark-level:{level}; bookmark-label:\\"{label}\\"; bookmark-state:open">'
            f'{esc}</h{level}>'
        )

        filepath = os.path.join(DOCS_DIR, os.path.basename(node.url))
        content  = process_html(filepath)

        # Remove the page's own <h1> if it duplicates the title we just emitted
        cs = BeautifulSoup(content, 'lxml')
        first_h1 = cs.find('h1')
        if first_h1 and first_h1.get_text(strip=True).lower() == node.title.lower():
            first_h1.decompose()
        out.append(str(cs))

        for child in node.children:
            render_node(child, out)

        out.append('</div>')


# ── 5. Build the clickable TOC page ──────────────────────────────────────────

def build_toc_page(roots):
    """
    Generates the front Table of Contents page.
    Section headers are bold labels; leaf pages are clickable links.
    Nesting is shown via indentation.
    """
    lines = [
        '<div class="toc-page" id="__toc__">',
        # The TOC page itself gets a top-level bookmark
        '<h1 class="toc-title" '
        'style="bookmark-level:1; bookmark-label:\\"Table of Contents\\"; bookmark-state:open">'
        'Table of Contents</h1>',
        '<nav><ul class="toc-root">',
    ]

    def toc_walk(node):
        indent = node.depth * 18  # px indent per depth level
        esc    = node.title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        if node.is_section:
            sec_id = "section_" + re.sub(r'[^A-Za-z0-9_]', '_', node.title)
            lines.append(
                f'<li class="toc-section" style="margin-left:{indent}px">'
                f'<a class="toc-section-link" href="#{sec_id}">{esc}</a>'
            )
        else:
            lines.append(
                f'<li class="toc-leaf" style="margin-left:{indent}px">'
                f'<a href="#{node.anchor}">{esc}</a>'
            )

        if node.children:
            lines.append('<ul>')
            for child in node.children:
                toc_walk(child)
            lines.append('</ul>')

        lines.append('</li>')

    for root_node in roots:
        toc_walk(root_node)

    lines += ['</ul></nav></div>']
    return '\n'.join(lines)


# ── 6. CSS ────────────────────────────────────────────────────────────────────

CSS = """
@page {
    margin: 2cm 2.5cm;
    size: A4;
}

body {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 10pt;
    color: #111;
}

/* ── Bookmark heading elements ── */
.bm-heading {
    color: #1a3a5c;
    margin-top: 0;
    padding-bottom: 3pt;
    /* WeasyPrint reads the bookmark-* properties from the style attribute */
}
.bm-section {
    font-size: 15pt;
    border-bottom: 1.5px solid #b0bec5;
    margin-bottom: 6pt;
}
.bm-page {
    font-size: 12pt;
    border-bottom: 1px solid #e0e0e0;
    margin-bottom: 4pt;
}

/* ── Page breaks ── */
.toc-page      { page-break-after: always; }
.section-block { page-break-before: always; }
.page-block    { page-break-before: always; }

/* Keep the very first section from getting a redundant break */
.toc-page + .section-block,
.toc-page + .page-block {
    page-break-before: always;
}

/* ── TOC page ── */
.toc-title {
    font-size: 20pt;
    color: #1a3a5c;
    border-bottom: 2px solid #1a3a5c;
    padding-bottom: 6pt;
    margin-bottom: 14pt;
}
.toc-root, .toc-root ul {
    list-style: none;
    padding-left: 0;
    margin: 0;
}
.toc-root li {
    margin-bottom: 2pt;
    line-height: 1.6;
}
.toc-section-link {
    font-weight: bold;
    color: #1a3a5c !important;
    text-decoration: none;
}
.toc-leaf a {
    color: #1155cc;
    text-decoration: none;
}
.toc-root ul {
    margin-top: 2pt;
    margin-bottom: 5pt;
}

/* ── General content ── */
a             { color: #1155cc; text-decoration: none; }
img           { max-width: 100%; height: auto; }
p.Note        { background: #fffbe6; border-left: 4px solid #f0c040;
                padding: 6pt 8pt; margin: 8pt 0; }
p.RelatedHead { font-weight: bold; margin-top: 12pt; }
table         { border-collapse: collapse; width: 100%; margin: 8pt 0; }
td, th        { border: 1px solid #ccc; padding: 4pt 6pt; font-size: 9pt; }
th            { background: #e8eef6; font-weight: bold; }
"""


# ── 7. Main ───────────────────────────────────────────────────────────────────

def main():
    print("Parsing TOC …")
    roots     = parse_toc(TOC_PATH)
    all_nodes = flatten(roots)
    file_nodes = [n for n in all_nodes if not n.is_section]
    print(f"  {len(all_nodes)} total nodes  |  {len(file_nodes)} with HTML files")

    print("Building merged HTML …")
    parts = [build_toc_page(roots)]

    for root_node in roots:
        render_node(root_node, parts)

    html_doc = (
        '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
        '  <meta charset="utf-8"/>\n'
        '  <title>MDX NetHelp Documentation</title>\n'
        f'  <style>{CSS}</style>\n'
        '</head>\n<body>\n'
        + ''.join(parts)
        + '\n</body>\n</html>'
    )

    # Save merged HTML for debugging
    merged_html_path = OUTPUT_PDF.replace('.pdf', '_merged.html')
    with open(merged_html_path, 'w', encoding='utf-8') as f:
        f.write(html_doc)
    print(f"  Merged HTML saved → {merged_html_path}")

    # Convert to PDF
    print("Converting to PDF — may take several minutes for large docs …")
    try:
        from weasyprint import HTML
        HTML(string=html_doc, base_url=DOCS_DIR).write_pdf(OUTPUT_PDF)
        print(f"\n✅  Done!  PDF saved → {OUTPUT_PDF}")
    except ImportError:
        print("\n⚠️  WeasyPrint not installed. Run:")
        print("    pip install weasyprint pillow beautifulsoup4 lxml")
        print("    sudo apt-get install -y libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz0b")
        print("Then re-run this script.")


if __name__ == '__main__':
    main()