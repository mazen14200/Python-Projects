import re
import sys
from pathlib import Path
from html import escape
from playwright.sync_api import sync_playwright


# ============================================================
# Configuration
# ============================================================

MERMAID_CDN = "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"


# ============================================================
# File Reading
# ============================================================

def read_markdown_file(file_path):
    encodings = [
        "utf-8-sig",
        "utf-8",
        "cp1256",
        "windows-1256",
        "utf-16"
    ]

    for enc in encodings:
        try:
            with open(file_path, "r", encoding=enc) as f:
                return f.read().lstrip("\ufeff")
        except UnicodeDecodeError:
            continue

    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
        return f.read().lstrip("\ufeff")


# ============================================================
# HTML Helpers
# ============================================================

def html_escape(text):
    return escape(text, quote=False)


def format_text(text):
    text = html_escape(text)

    # Bold
    text = re.sub(
        r"\*\*([^*]+)\*\*",
        r"<strong>\1</strong>",
        text
    )

    # Italic
    text = re.sub(
        r"\*([^*]+)\*",
        r"<em>\1</em>",
        text
    )

    # Inline code
    text = re.sub(
        r"`([^`]+)`",
        r"<code>\1</code>",
        text
    )

    # Markdown links
    def make_link(match):
        label = match.group(1)
        url = match.group(2)

        url = (
            url
            .replace("&amp;", "&")
            .replace("&lt;", "<")
            .replace("&gt;", ">")
        )

        return f'<a href="{url}">{label}</a>'

    text = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        make_link,
        text
    )

    return text


def slugify(text):
    text = (text or "").strip().lower()

    text = re.sub(
        r"[^\w\s-]",
        "",
        text,
        flags=re.UNICODE
    )

    text = re.sub(
        r"[\s_]+",
        "-",
        text
    )

    return text.strip("-")


# ============================================================
# Mermaid Validation
# ============================================================

_MERMAID_DIAGRAM_KEYWORDS = (
    "graph",
    "flowchart",
    "sequenceDiagram",
    "classDiagram",
    "stateDiagram",
    "erDiagram",
    "journey",
    "gantt",
    "pie",
    "gitGraph",
    "mindmap",
    "timeline",
    "quadrantChart",
    "requirementDiagram",
    "C4Context",
    "C4Container",
    "C4Component",
    "C4Dynamic",
    "sankey-beta",
    "block-beta",
    "xychart-beta",
)

_MERMAID_SUSPECT_BRACE_RE = re.compile(
    r"(?<![|{}\w])\{[A-Za-z_][A-Za-z0-9_]*\}"
)


def validate_mermaid(raw_code):
    warnings = []

    lines = [
        line
        for line in raw_code.strip().split("\n")
        if line.strip()
    ]

    if not lines:
        warnings.append("Mermaid diagram is empty.")
        return warnings

    first_line = lines[0].strip()

    diagram_type = (
        first_line.split()[0]
        if first_line.split()
        else ""
    )

    if not any(
        first_line.startswith(keyword)
        for keyword in _MERMAID_DIAGRAM_KEYWORDS
    ):
        warnings.append(
            f"First line '{first_line[:50]}' "
            f"does not look like a known Mermaid diagram type."
        )

    bracket_checks = [
        ("[", "]"),
        ("(", ")")
    ]

    if diagram_type not in ("erDiagram",):
        bracket_checks.append(("{", "}"))

    for open_ch, close_ch in bracket_checks:

        n_open = raw_code.count(open_ch)
        n_close = raw_code.count(close_ch)

        if n_open != n_close:
            warnings.append(
                f"Number of '{open_ch}' ({n_open}) "
                f"does not match '{close_ch}' ({n_close})."
            )

    if raw_code.count('"') % 2 != 0:
        warnings.append(
            'Odd number of double quotes detected.'
        )

    if diagram_type in ("graph", "flowchart"):

        suspects = sorted(
            set(
                _MERMAID_SUSPECT_BRACE_RE.findall(raw_code)
            )
        )

        if suspects:
            examples = ", ".join(suspects[:5])

            warnings.append(
                f"Possible Mermaid placeholder detected: {examples}"
            )

    return warnings


# ============================================================
# Lists
# ============================================================

def parse_list(lines, start, n):

    i = start

    first_line = lines[i].strip()

    is_ordered = bool(
        re.match(r"^\d+\. ", first_line)
    )

    tag = "ol" if is_ordered else "ul"

    items = []

    while i < n:

        line = lines[i]
        stripped = line.strip()

        if not stripped:
            break

        if stripped.startswith("- ") or stripped.startswith("* "):

            items.append(
                f"<li>{format_text(stripped[2:])}</li>"
            )

        elif re.match(r"^\d+\. ", stripped):

            content = stripped[
                stripped.find(".") + 2:
            ]

            items.append(
                f"<li>{format_text(content)}</li>"
            )

        elif line.startswith("  ") or line.startswith("\t"):

            if items:

                last = items[-1]

                if last.endswith("</li>"):

                    items[-1] = (
                        last[:-5]
                        + " "
                        + format_text(stripped)
                        + "</li>"
                    )

        else:
            break

        i += 1

    return (
        f"<{tag}>\n"
        + "\n".join(items)
        + f"\n</{tag}>"
    )


# ============================================================
# Tables
# ============================================================

def convert_table(table_lines):

    html = [
        '<div class="table-wrapper">',
        "<table>"
    ]

    rows = []

    for line in table_lines:

        cells = [
            cell.strip()
            for cell in line.split("|")
            if cell.strip()
        ]

        if cells:
            rows.append(cells)

    if len(rows) >= 2:

        html.append("<thead>")
        html.append("<tr>")

        for cell in rows[0]:

            html.append(
                f"<th>{format_text(cell)}</th>"
            )

        html.append("</tr>")
        html.append("</thead>")

        html.append("<tbody>")

        for row in rows[2:]:

            html.append("<tr>")

            for cell in row:

                html.append(
                    f"<td>{format_text(cell)}</td>"
                )

            html.append("</tr>")

        html.append("</tbody>")

    html.append("</table>")
    html.append("</div>")

    return "\n".join(html)


# ============================================================
# Markdown → HTML
# ============================================================

def convert_markdown_to_html(markdown_content):

    html_parts = []

    html_parts.append(
        '<section class="documentation-module">'
    )

    lines = markdown_content.split("\n")

    current_section = None
    doc_title = ""

    used_section_ids = {}

    i = 0
    n = len(lines)

    while i < n:

        line = lines[i]
        stripped_line = line.strip()

        # ----------------------------------------------------
        # H1
        # ----------------------------------------------------

        if (
            stripped_line.startswith("# ")
            and not stripped_line.startswith("## ")
        ):

            doc_title = stripped_line[2:].strip()

            html_parts.append(
                f"<h1>{format_text(doc_title)}</h1>"
            )

        # ----------------------------------------------------
        # H2
        # ----------------------------------------------------

        elif (
            stripped_line.startswith("## ")
            and not stripped_line.startswith("### ")
        ):

            if current_section:
                html_parts.append("</section>")

            section_name = stripped_line[3:].strip()

            base_slug = slugify(section_name) or "section"

            section_id = f"module-{base_slug}"

            if section_id in used_section_ids:

                used_section_ids[section_id] += 1

                section_id = (
                    f"{section_id}-"
                    f"{used_section_ids[section_id]}"
                )

            else:

                used_section_ids[section_id] = 0

            html_parts.append(
                f'<section class="module-section" '
                f'id="{section_id}">'
            )

            html_parts.append(
                f"<h2>{format_text(section_name)}</h2>"
            )

            current_section = section_name

        # ----------------------------------------------------
        # H3
        # ----------------------------------------------------

        elif (
            stripped_line.startswith("### ")
            and not stripped_line.startswith("#### ")
        ):

            html_parts.append(
                f"<h3>{format_text(stripped_line[4:])}</h3>"
            )

        # ----------------------------------------------------
        # H4
        # ----------------------------------------------------

        elif (
            stripped_line.startswith("#### ")
            and not stripped_line.startswith("##### ")
        ):

            html_parts.append(
                f"<h4>{format_text(stripped_line[5:])}</h4>"
            )

        # ----------------------------------------------------
        # H5
        # ----------------------------------------------------

        elif stripped_line.startswith("##### "):

            html_parts.append(
                f"<h5>{format_text(stripped_line[6:])}</h5>"
            )

        # ----------------------------------------------------
        # Horizontal Rule
        # ----------------------------------------------------

        elif stripped_line == "---":

            html_parts.append("<hr>")

        # ----------------------------------------------------
        # Code Block
        # ----------------------------------------------------

        elif stripped_line.startswith("```"):

            language = (
                stripped_line[3:].strip()
                or "text"
            )

            code_lines = []

            i += 1

            while (
                i < n
                and not lines[i].strip().startswith("```")
            ):

                code_lines.append(lines[i])

                i += 1

            if i >= n:

                print(
                    f"WARNING: Unterminated code fence "
                    f"({language})"
                )

            raw_code = "\n".join(code_lines).strip()

            # Mermaid
            if language.lower() == "mermaid":

                warnings = validate_mermaid(raw_code)

                if warnings:

                    print(
                        "WARNING: Mermaid issues detected:"
                    )

                    for warning in warnings:
                        print(f"  - {warning}")

                code_content = html_escape(raw_code)

                html_parts.append(
                    f'<pre class="mermaid">'
                    f'{code_content}'
                    f'</pre>'
                )

            else:

                code_content = html_escape(raw_code)

                if (
                    "folder" in language.lower()
                    or "structure" in language.lower()
                ):

                    html_parts.append(
                        '<pre class="folder-structure">'
                        f"{code_content}"
                        "</pre>"
                    )

                else:

                    html_parts.append(
                        "<pre>"
                        f'<code class="language-{language}">'
                        f"{code_content}"
                        "</code>"
                        "</pre>"
                    )

        # ----------------------------------------------------
        # Table
        # ----------------------------------------------------

        elif (
            "|" in stripped_line
            and i + 1 < n
            and "|" in lines[i + 1]
            and set(lines[i + 1].strip())
            - set("|-: ") == set()
        ):

            table_lines = []

            while i < n and "|" in lines[i]:

                table_lines.append(lines[i])

                i += 1

            html_parts.append(
                convert_table(table_lines)
            )

            continue

        # ----------------------------------------------------
        # Blockquote
        # ----------------------------------------------------

        elif stripped_line.startswith("> "):

            quote_lines = []

            while (
                i < n
                and lines[i].strip().startswith("> ")
            ):

                quote_lines.append(
                    lines[i].strip()[2:]
                )

                i += 1

            html_parts.append(
                "<blockquote>"
                f"<p>{format_text(' '.join(quote_lines))}</p>"
                "</blockquote>"
            )

            continue

        # ----------------------------------------------------
        # List
        # ----------------------------------------------------

        elif (
            stripped_line.startswith("- ")
            or stripped_line.startswith("* ")
            or re.match(
                r"^\d+\. ",
                stripped_line
            )
        ):

            html_parts.append(
                parse_list(lines, i, n)
            )

            while i < n:

                current = lines[i].strip()

                if not current:
                    break

                if (
                    current.startswith("- ")
                    or current.startswith("* ")
                    or re.match(
                        r"^\d+\. ",
                        current
                    )
                    or lines[i].startswith("  ")
                ):

                    i += 1

                else:
                    break

            continue

        # ----------------------------------------------------
        # Info Box
        # ----------------------------------------------------

        elif stripped_line in [
            "Important Notes",
            "Warnings",
            "Tips",
            "Recommendations"
        ]:

            html_parts.append(
                '<div class="info-box">'
            )

            html_parts.append(
                f"<strong>{stripped_line}</strong>"
            )

            i += 1

            while i < n:

                box_line = lines[i].strip()

                if not box_line:
                    break

                html_parts.append(
                    f"<p>{format_text(box_line)}</p>"
                )

                i += 1

            html_parts.append("</div>")

            continue

        # ----------------------------------------------------
        # Regular Text
        # ----------------------------------------------------

        elif stripped_line:

            html_parts.append(
                f"<p>{format_text(stripped_line)}</p>"
            )

        i += 1

    if current_section:
        html_parts.append("</section>")

    html_parts.append("</section>")

    body_content = "\n".join(html_parts)

    return wrap_in_full_html(
        body_content,
        doc_title
    )


# ============================================================
# Full HTML
# ============================================================

def wrap_in_full_html(
    body_content,
    title="Documentation"
):

    return f"""<!DOCTYPE html>

<html lang="ar" dir="rtl">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>{escape(title)}</title>

<style>

*,
*::before,
*::after {{
    box-sizing: border-box;
}}

html {{
    direction: rtl;
}}

body {{

    font-family:
        "Segoe UI",
        Tahoma,
        Arial,
        sans-serif;

    background: #f4f6f9;

    color: #2c3e50;

    line-height: 1.8;

    padding: 2rem;

    direction: rtl;
}}

.documentation-module {{

    max-width: 960px;

    margin: 0 auto;

    background: #fff;

    border-radius: 12px;

    box-shadow:
        0 2px 16px rgba(0,0,0,0.07);

    padding: 2.5rem 3rem;
}}


/* ==========================================================
   Headings
   ========================================================== */

h1 {{

    font-size: 2rem;

    color: #1a5276;

    border-bottom:
        3px solid #2980b9;

    padding-bottom: 0.6rem;

    margin-bottom: 1.5rem;
}}

h2 {{

    font-size: 1.5rem;

    color: #21618c;

    margin-top: 2rem;

    margin-bottom: 1rem;

    padding-bottom: 0.4rem;

    border-bottom:
        2px solid #aed6f1;
}}

h3 {{

    font-size: 1.25rem;

    color: #2e86c1;

    margin-top: 1.5rem;

    margin-bottom: 0.75rem;
}}

h4 {{

    font-size: 1.1rem;

    color: #3498db;

    margin-top: 1.2rem;

    margin-bottom: 0.5rem;
}}

h5 {{

    font-size: 1rem;

    color: #5dade2;

    margin-top: 1rem;

    margin-bottom: 0.4rem;
}}


/* ==========================================================
   Paragraph
   ========================================================== */

p {{
    margin-bottom: 0.8rem;
}}


/* ==========================================================
   Links
   ========================================================== */

a {{

    color: #2980b9;

    text-decoration: none;
}}


/* ==========================================================
   Lists
   ========================================================== */

ul,
ol {{

    margin:
        0.8rem 1.5rem 0.8rem 0;

    padding-right: 1.5rem;
}}

li {{
    margin-bottom: 0.4rem;
}}


/* ==========================================================
   Tables
   ========================================================== */

.table-wrapper {{

    overflow-x: auto;

    margin: 1rem 0;
}}

table {{

    width: 100%;

    border-collapse: collapse;

    font-size: 0.95rem;
}}

thead {{

    background: #2980b9;

    color: #fff;
}}

th,
td {{

    padding:
        0.6rem 1rem;

    border:
        1px solid #d5dbdb;

    text-align: right;

    vertical-align: top;
}}

tbody tr:nth-child(even) {{
    background: #eaf2f8;
}}


/* ==========================================================
   Code
   ========================================================== */

code {{

    background: #eaf2f8;

    color: #c0392b;

    padding:
        0.15rem 0.4rem;

    border-radius: 4px;

    font-family:
        Consolas,
        "Courier New",
        monospace;

    font-size: 0.9em;
}}

pre {{

    background: #2c3e50;

    color: #ecf0f1;

    padding: 1.2rem;

    border-radius: 8px;

    overflow-x: auto;

    margin: 1rem 0;

    direction: ltr;

    text-align: left;

    white-space: pre-wrap;

    word-wrap: break-word;
}}

pre code {{

    background: transparent;

    color: inherit;

    padding: 0;
}}


/* ==========================================================
   Mermaid
   ========================================================== */

pre.mermaid {{

    background: #fff;

    color: #2c3e50;

    text-align: center;

    direction: ltr;

    margin: 1rem auto;

    padding: 1rem;

    border-radius: 8px;

    border:
        1px solid #e2e8f0;

    overflow: visible;
}}


/* ==========================================================
   Folder Structure
   ========================================================== */

.folder-structure {{

    background: #1b2631;

    color: #58d68d;

    font-size: 0.9rem;
}}


/* ==========================================================
   Blockquote
   ========================================================== */

blockquote {{

    border-right:
        4px solid #2980b9;

    padding:
        0.8rem 1.2rem;

    margin: 1rem 0;

    background: #eaf2f8;

    border-radius:
        0 8px 8px 0;
}}


/* ==========================================================
   Info Box
   ========================================================== */

.info-box {{

    background: #fef9e7;

    border:
        1px solid #f9e79f;

    border-right:
        4px solid #f39c12;

    padding:
        1rem 1.2rem;

    border-radius:
        0 8px 8px 0;

    margin: 1rem 0;
}}

.info-box strong {{

    color: #d68910;

    display: block;

    margin-bottom: 0.4rem;
}}


/* ==========================================================
   Sections
   ========================================================== */

.module-section {{

    margin-bottom: 2rem;

    padding: 1rem 0;
}}


/* ==========================================================
   Horizontal Rule
   ========================================================== */

hr {{

    border: none;

    border-top:
        2px solid #d5dbdb;

    margin: 1.5rem 0;
}}


/* ==========================================================
   Print
   ========================================================== */

@media print {{

    @page {{

        size: A4 portrait;

        margin:
            1.2cm
            1cm
            1.5cm
            1cm;
    }}

    html,
    body {{

        background: #fff !important;

        color: #000 !important;

        padding: 0 !important;

        margin: 0 !important;

        font-size: 10pt;

        width: 100% !important;

        height: auto !important;
    }}

    .documentation-module {{

        box-shadow: none !important;

        padding: 0 !important;

        max-width: 100% !important;

        margin: 0 !important;

        border: none !important;

        border-radius: 0 !important;
    }}

    h1,
    h2,
    h3,
    h4,
    h5 {{

        break-after: avoid !important;

        page-break-after: avoid !important;

        break-inside: avoid !important;

        page-break-inside: avoid !important;
    }}

    .module-section {{

        break-inside: auto;

        page-break-inside: auto;
    }}

    pre {{

        white-space: pre-wrap !important;

        word-wrap: break-word !important;

        overflow: visible !important;

        font-size: 8pt !important;

        padding: 0.6rem !important;

        margin: 0.6rem 0 !important;

        break-inside: avoid !important;

        page-break-inside: avoid !important;
    }}

    pre.mermaid {{

        background: transparent !important;

        border: none !important;

        margin: 0.5rem auto !important;

        padding: 0 !important;

        width: 100% !important;

        max-width: 100% !important;

        overflow: visible !important;

        break-inside: avoid !important;

        page-break-inside: avoid !important;
    }}

    pre.mermaid svg {{

        max-width: 100% !important;

        height: auto !important;

        max-height: 22cm !important;

        display: block !important;

        margin: 0 auto !important;
    }}

    .table-wrapper {{

        overflow: visible !important;

        width: 100% !important;
    }}

    table {{

        width: 100% !important;

        table-layout: fixed !important;

        font-size: 8.5pt !important;
    }}

    th,
    td {{

        word-break: break-word !important;

        padding:
            0.4rem
            0.6rem !important;
    }}

    tr {{

        break-inside: avoid !important;

        page-break-inside: avoid !important;
    }}

    thead {{

        display:
            table-header-group !important;
    }}

    img {{

        max-width: 100% !important;

        height: auto !important;

        break-inside: avoid !important;

        page-break-inside: avoid !important;
    }}
}}

</style>

</head>

<body>

{body_content}

<script src="{MERMAID_CDN}"></script>

<script>

mermaid.initialize({{
    startOnLoad: false,
    securityLevel: "loose",
    theme: "default"
}});

async function renderMermaid() {{

    const elements =
        document.querySelectorAll(
            ".mermaid"
        );

    for (let i = 0; i < elements.length; i++) {{

        const element = elements[i];

        const code =
            element.textContent;

        try {{

            const result =
                await mermaid.render(
                    "mermaid-" + i,
                    code
                );

            element.innerHTML =
                result.svg;

        }} catch (error) {{

            console.error(
                "Mermaid rendering error:",
                error
            );

            element.innerHTML =
                "<pre>" +
                code.replace(
                    /</g,
                    "&lt;"
                ) +
                "</pre>";
        }}
    }}
}}

window.renderMermaid =
    renderMermaid;

</script>

</body>

</html>
"""


# ============================================================
# Convert One Markdown File
# ============================================================

def convert_file(md_file, output_dir, browser):

    print()
    print("=" * 70)
    print(f"Converting: {md_file}")
    print("=" * 70)

    try:

        markdown_content = read_markdown_file(
            md_file
        )

        html_content = convert_markdown_to_html(
            markdown_content
        )

        output_dir.mkdir(
            parents=True,
            exist_ok=True
        )

        pdf_path = (
            output_dir
            / f"{md_file.stem}.pdf"
        )

        page = browser.new_page()

        page.set_content(
            html_content,
            wait_until="networkidle"
        )

        # Wait for Mermaid CDN
        try:

            page.wait_for_function(
                """
                () => typeof mermaid !== 'undefined'
                """,
                timeout=15000
            )

        except Exception:

            print(
                "WARNING: Mermaid library "
                "did not load."
            )

        # Render Mermaid
        try:

            page.evaluate(
                """
                async () => {
                    if (window.renderMermaid) {
                        await window.renderMermaid();
                    }
                }
                """
            )

            page.wait_for_timeout(1000)

        except Exception as ex:

            print(
                f"WARNING: Mermaid rendering failed: {ex}"
            )

        # PDF
        page.pdf(
            path=str(pdf_path),
            format="A4",
            print_background=True,
            prefer_css_page_size=True,
            margin={
                "top": "0",
                "right": "0",
                "bottom": "0",
                "left": "0"
            }
        )

        page.close()

        print(
            f"PDF saved: {pdf_path}"
        )

        return True

    except Exception as ex:

        print(
            f"ERROR converting {md_file}:"
        )

        print(ex)

        return False


# ============================================================
# Get Markdown Files
# ============================================================

def get_markdown_files(input_path):

    input_path = Path(input_path)

    if input_path.is_file():

        if input_path.suffix.lower() != ".md":

            print(
                "ERROR: Input file must be a .md file."
            )

            return []

        return [input_path]

    if input_path.is_dir():

        return sorted(
            input_path.rglob("*.md")
        )

    print(
        f"ERROR: Path does not exist:"
        f"\n{input_path}"
    )

    return []


# ============================================================
# Main
# ============================================================

def main():

    if len(sys.argv) != 2:

        print()
        print(
            "Usage:"
        )

        print(
            '  python generatePdfFromMd.py "PATH"'
        )

        print()

        print("Examples:")

        print(
            '  python generatePdfFromMd.py "D:\\Docs\\file.md"'
        )

        print(
            '  python generatePdfFromMd.py "D:\\Docs"'
        )

        sys.exit(1)

    input_path = Path(
        sys.argv[1]
    ).resolve()

    markdown_files = get_markdown_files(
        input_path
    )

    if not markdown_files:

        print(
            "No Markdown files found."
        )

        sys.exit(1)

    # --------------------------------------------------------
    # Output Directory
    # --------------------------------------------------------

    if input_path.is_file():

        output_dir = (
            input_path.parent
            / "PDF"
        )

    else:

        output_dir = (
            input_path
            / "PDF"
        )

    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    print()
    print("=" * 70)
    print("Markdown → PDF Generator")
    print("=" * 70)

    print(
        f"Input : {input_path}"
    )

    print(
        f"Output: {output_dir}"
    )

    print(
        f"Files : {len(markdown_files)}"
    )

    print("=" * 70)

    success_count = 0
    failed_count = 0

    # --------------------------------------------------------
    # Playwright
    # --------------------------------------------------------

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True
        )

        try:

            for md_file in markdown_files:

                # Don't convert files inside PDF folder
                if output_dir in md_file.parents:
                    continue

                success = convert_file(
                    md_file,
                    output_dir,
                    browser
                )

                if success:
                    success_count += 1
                else:
                    failed_count += 1

        finally:

            browser.close()

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("Conversion Finished")
    print("=" * 70)

    print(
        f"Successful: {success_count}"
    )

    print(
        f"Failed    : {failed_count}"
    )

    print(
        f"Output    : {output_dir}"
    )

    print("=" * 70)


# ============================================================
# Entry Point
# ============================================================

if __name__ == "__main__":
    main()