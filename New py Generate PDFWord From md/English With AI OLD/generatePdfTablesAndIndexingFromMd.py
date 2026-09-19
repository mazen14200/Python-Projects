import re
import os
import sys
from pathlib import Path


def read_markdown_file(file_path):
    encodings = ['utf-8-sig', 'utf-8', 'cp1256', 'windows-1256', 'utf-16']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                return f.read().lstrip('\ufeff')
        except UnicodeDecodeError:
            continue
    # If all encodings fail, try with errors='replace'
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    # Strip BOM if present
    return content.lstrip('\ufeff')


def write_html_file(output_dir, file_name, html_content):
    os.makedirs(output_dir, exist_ok=True)
    output_path = Path(output_dir) / f"{Path(file_name).stem}.html"
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    return str(output_path)


def html_escape(text):
    """Escape HTML special characters (used for code blocks AND plain text)."""
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    return text


def format_text(text):
    # Escape raw HTML special chars FIRST so stray < > & in Arabic/technical
    # text (e.g. "أقل من <100") never breaks the page or swallows content.
    text = html_escape(text)
    # Handle bold
    text = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', text)
    # Handle italic
    text = re.sub(r'\*([^*]+)\*', r'<em>\1</em>', text)
    # Handle inline code
    text = re.sub(r'`([^`]+)`', r'<code>\1</code>', text)
    # Handle links [text](url) -- run AFTER escaping, so re-unescape the URL
    # itself since it must not contain &amp; etc. for href to work right.
    def _make_link(m):
        label, url = m.group(1), m.group(2)
        url = url.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
        return f'<a href="{url}">{label}</a>'
    text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', _make_link, text)
    return text


def slugify(text):
    """
    Build a URL-safe id fragment that supports Arabic/Unicode text.
    The old version stripped everything except a-z0-9, which erases
    Arabic text entirely and makes every heading collapse to the same
    (empty) id -> duplicate ids -> sidebar links jump to the wrong place.
    """
    text = (text or '').strip().lower()
    # \w with Python's default (unicode) flags keeps Arabic letters/digits
    text = re.sub(r'[^\w\s-]', '', text, flags=re.UNICODE)
    text = re.sub(r'[\s_]+', '-', text).strip('-')
    return text


# ===================== Mermaid handling =====================
#
# Mermaid special characters that break parsing when they occur as
# literal DATA inside a label (as opposed to Mermaid's own syntax):
#   {  }   -> node-shape delimiters (decision/rhombus, hexagon...)
#   {  }   -> node-shape delimiters (decision/rhombus, hexagon...) in
#             flowcharts, BUT part of relationship syntax in erDiagram
#             ("||--|{"), BUT class-body delimiters in classDiagram...
#   <  >   -> relationship arrows (classDiagram "<|--"), <br/> line breaks
#   |      -> edge-label delimiters in flowcharts, BUT relationship
#             cardinality tokens in erDiagram ("}|--||")...
#
# Earlier versions of this script tried to auto-escape these characters
# with regex heuristics (e.g. turning a literal "{reportId}" inside a
# label into a safe placeholder). In practice this kept causing NEW
# breakage: the same character means something different in each
# diagram type, so a rule that is correct for flowchart silently
# corrupts erDiagram, and the next diagram type would break something
# else again. There is no way to fully disambiguate this with regex --
# it needs Mermaid's real grammar, which we don't have in Python.
#
# The reliable choice is therefore: NEVER rewrite the Mermaid source.
# Pass it through byte-for-byte exactly as written in the .md file (the
# only transformation applied is the standard HTML escape of & < >,
# which is fully transparent -- browsers decode those entities back to
# the original characters before Mermaid.js ever reads them, so Mermaid
# always sees exactly what was written in the markdown). Instead of
# guessing and possibly corrupting the diagram, validate_mermaid() below
# only WARNS about patterns that are known to sometimes cause problems,
# so they can be reviewed and fixed by hand in the source .md file.

_MERMAID_DIAGRAM_KEYWORDS = (
    'graph', 'flowchart', 'sequenceDiagram', 'classDiagram', 'stateDiagram',
    'erDiagram', 'journey', 'gantt', 'pie', 'gitGraph', 'mindmap',
    'timeline', 'quadrantChart', 'requirementDiagram', 'C4Context',
    'C4Container', 'C4Component', 'C4Dynamic', 'sankey-beta', 'block-beta',
    'xychart-beta',
)

# A bare {word} that shows up right after plain text (not immediately
# after a node id, and not part of a known relationship token like "|{"
# or "}|") is the pattern most likely to have caused real breakage in
# practice (REST-style placeholders such as "/users/{id}" typed straight
# into a flowchart label). We only use this to raise a warning -- never
# to rewrite the diagram.
_MERMAID_SUSPECT_BRACE_RE = re.compile(r'(?<![|{}\w])\{[A-Za-z_][A-Za-z0-9_]*\}')


def validate_mermaid(raw_code):
    """
    Lightweight structural sanity check (not a full Mermaid parser).
    Catches the most common causes of a diagram failing to render so
    problems surface as a warning during conversion instead of silently
    breaking in the browser. Returns a list of warning strings.
    Never modifies the diagram -- warnings only, so review is manual.
    """
    warnings = []
    lines = [l for l in raw_code.strip().split('\n') if l.strip()]
    if not lines:
        warnings.append("مخطط Mermaid فارغ")
        return warnings

    first_line = lines[0].strip()
    diagram_type = first_line.split()[0] if first_line.split() else ''
    if not any(first_line.startswith(kw) for kw in _MERMAID_DIAGRAM_KEYWORDS):
        warnings.append(
            f"السطر الأول '{first_line[:50]}' لا يبدأ بنوع مخطط معروف "
            f"(graph / flowchart / sequenceDiagram / classDiagram / ...)"
        )

    bracket_checks = [('[', ']'), ('(', ')')]
    if diagram_type not in ('erDiagram',):
        # erDiagram's crow's-foot cardinality tokens (e.g. "||--|{",
        # "}|--||") legitimately use single unpaired '{' / '}' characters
        # as part of a 2-char relationship symbol, not as matching
        # brackets -- so this check would false-positive on valid ER
        # diagrams and is skipped for that diagram type.
        bracket_checks.append(('{', '}'))

    for open_ch, close_ch in bracket_checks:
        n_open, n_close = raw_code.count(open_ch), raw_code.count(close_ch)
        if n_open != n_close:
            warnings.append(
                f"عدد '{open_ch}' ({n_open}) لا يطابق عدد '{close_ch}' ({n_close}) "
                f"-- على الأرجح قوس مش متقفل"
            )

    if raw_code.count('"') % 2 != 0:
        warnings.append('عدد علامات الاقتباس " فردي -- على الأرجح اقتباس غير مغلق')

    if diagram_type in ('graph', 'flowchart'):
        suspects = sorted(set(_MERMAID_SUSPECT_BRACE_RE.findall(raw_code)))
        if suspects:
            examples = '، '.join(suspects[:5])
            warnings.append(
                f"لقيت نص شبه {{placeholder}} زي {examples} -- Mermaid ممكن "
                f"يفهمه غلط كشكل rhombus. لو المخطط ما ظهرش صح، حط النص ده "
                f"جوه علامات اقتباس داخل القوس، مثلاً: [\"...{examples.split('،')[0]}...\"]"
            )

    return warnings


def convert_markdown_to_html(markdown_content):
    html_parts = []
    html_parts.append('<section class="documentation-module">')

    lines = markdown_content.split('\n')
    current_section = None
    doc_title = ""
    used_section_ids = {}  # guarantees every H2 in this doc gets a unique id

    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped_line = line.strip()

        # H1
        if stripped_line.startswith('# ') and not stripped_line.startswith('## '):
            doc_title = stripped_line[2:].strip()
            html_parts.append(f'<h1>{format_text(doc_title)}</h1>')

        # H2 - start of module-section
        elif stripped_line.startswith('## ') and not stripped_line.startswith('### '):
            if current_section:
                html_parts.append('</section>')

            section_name = stripped_line[3:].strip()
            base_slug = slugify(section_name) or 'section'
            section_id = f'module-{base_slug}'
            # If this slug was already used in this document (e.g. two
            # Arabic headings that slugify the same way, or both empty),
            # append a counter so ids never collide.
            if section_id in used_section_ids:
                used_section_ids[section_id] += 1
                section_id = f'{section_id}-{used_section_ids[section_id]}'
            else:
                used_section_ids[section_id] = 0
            html_parts.append(f'<section class="module-section" id="{section_id}">')
            html_parts.append(f'<h2>{format_text(section_name)}</h2>')
            current_section = section_name

        # H3
        elif stripped_line.startswith('### ') and not stripped_line.startswith('#### '):
            html_parts.append(f'<h3>{format_text(stripped_line[4:])}</h3>')

        # H4
        elif stripped_line.startswith('#### ') and not stripped_line.startswith('##### '):
            html_parts.append(f'<h4>{format_text(stripped_line[5:])}</h4>')

        # H5
        elif stripped_line.startswith('##### '):
            html_parts.append(f'<h5>{format_text(stripped_line[6:])}</h5>')

        # Horizontal rule
        elif stripped_line == '---':
            html_parts.append('<hr>')

        # Code block
        elif stripped_line.startswith('```'):
            language = stripped_line[3:].strip() or 'text'
            code_lines = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            if i >= n:
                # Unterminated ``` fence: without this, the rest of the
                # file (including later ## headings) silently gets
                # swallowed into one giant code block and disappears
                # from the sidebar/navigation.
                print(f"  WARNING: unterminated code fence (```{language}) "
                      f"-- rest of file may have been swallowed as code")

            raw_code = '\n'.join(code_lines).strip()

            if language == 'mermaid':
                mermaid_warnings = validate_mermaid(raw_code)
                if mermaid_warnings:
                    print(f"  WARNING: Mermaid diagram issues detected:")
                    for w in mermaid_warnings:
                        print(f"    - {w}")
                # Pass the diagram through EXACTLY as written in the .md
                # file. The only transformation is the standard HTML
                # escape of & < > (needed to be valid inside <pre>) --
                # this is fully transparent since browsers decode those
                # entities back to the original characters before Mermaid.js
                # reads the text, so Mermaid always sees byte-for-byte
                # what was authored in the markdown. See the comment
                # above validate_mermaid() for why we no longer try to
                # auto-rewrite the diagram content.
                code_content = html_escape(raw_code)
                html_parts.append(f'<pre class="mermaid">{code_content}</pre>')
            else:
                code_content = html_escape(raw_code)
                if 'folder' in language or 'structure' in language:
                    html_parts.append(f'<pre class="folder-structure">{code_content}</pre>')
                else:
                    html_parts.append(f'<pre><code class="language-{language}">{code_content}</code></pre>')

        # Table
        elif '|' in stripped_line and i + 1 < n and '|' in lines[i + 1] and set(lines[i + 1].strip()) - set('|-: ') == set():
            table_lines = []
            while i < n and '|' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            html_parts.append(convert_table(table_lines))
            continue  # don't increment i again

        # Blockquote
        elif stripped_line.startswith('> '):
            quote_lines = []
            while i < n and lines[i].strip().startswith('> '):
                quote_lines.append(lines[i].strip()[2:])
                i += 1
            html_parts.append(f'<blockquote><p>{format_text(" ".join(quote_lines))}</p></blockquote>')
            continue

        # List (unordered or ordered) - with nested list support
        elif stripped_line.startswith('- ') or stripped_line.startswith('* ') or re.match(r'^\d+\. ', stripped_line):
            html_parts.append(parse_list(lines, i, n))
            # Advance past all list lines
            while i < n:
                l = lines[i].strip()
                if not l:
                    break
                if l.startswith('- ') or l.startswith('* ') or re.match(r'^\d+\. ', l) or lines[i].startswith('  '):
                    i += 1
                else:
                    break
            continue

        # Info box (Important Notes, Warnings, Tips, Recommendations)
        elif stripped_line in ['Important Notes', 'Warnings', 'Tips', 'Recommendations']:
            html_parts.append('<div class="info-box">')
            html_parts.append(f'<strong>{stripped_line}</strong>')
            i += 1
            while i < n:
                box_line = lines[i].strip()
                if not box_line:
                    break
                html_parts.append(f'<p>{format_text(box_line)}</p>')
                i += 1
            html_parts.append('</div>')
            continue

        # Regular text
        elif stripped_line:
            html_parts.append(f'<p>{format_text(stripped_line)}</p>')

        i += 1

    if current_section:
        html_parts.append('</section>')

    html_parts.append('</section>')

    body_content = '\n'.join(html_parts)
    return wrap_in_full_html(body_content, doc_title)


def parse_list(lines, start, n):
    """Parse a list block with basic nesting support."""
    i = start
    first_line = lines[i].strip()
    is_ordered = bool(re.match(r'^\d+\. ', first_line))
    tag = 'ol' if is_ordered else 'ul'
    items = []

    while i < n:
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            break

        # Check for list item markers
        if stripped.startswith('- ') or stripped.startswith('* '):
            items.append(f'<li>{format_text(stripped[2:])}</li>')
        elif re.match(r'^\d+\. ', stripped):
            content = stripped[stripped.find('.') + 2:]
            items.append(f'<li>{format_text(content)}</li>')
        elif line.startswith('  ') or line.startswith('\t'):
            # Indented continuation of previous item
            if items:
                last = items[-1]
                if last.endswith('</li>'):
                    items[-1] = last[:-5] + ' ' + format_text(stripped) + '</li>'
        else:
            break
        i += 1

    return f'<{tag}>\n' + '\n'.join(items) + f'\n</{tag}>'


def convert_table(table_lines):
    html = ['<div class="table-wrapper">', '<table>']
    rows = []

    for line in table_lines:
        cells = [cell.strip() for cell in line.split('|') if cell.strip()]
        if cells:
            rows.append(cells)

    if len(rows) >= 2:
        html.append('<thead>')
        html.append('<tr>')
        for cell in rows[0]:
            html.append(f'<th>{format_text(cell)}</th>')
        html.append('</tr>')
        html.append('</thead>')

        html.append('<tbody>')
        for row in rows[2:]:  # skip separator row
            html.append('<tr>')
            for cell in row:
                html.append(f'<td>{format_text(cell)}</td>')
            html.append('</tr>')
        html.append('</tbody>')

    html.append('</table>')
    html.append('</div>')
    return '\n'.join(html)


def wrap_in_full_html(body_content, title="Documentation"):
    """Wrap content in a complete HTML document with embedded CSS for standalone viewing."""
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        /* ===== Reset & Base ===== */
        *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

        body {{
            font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
            background: #f4f6f9;
            color: #2c3e50;
            line-height: 1.8;
            padding: 2rem;
            direction: rtl;
        }}

        /* ===== Main Container ===== */
        .documentation-module {{
            max-width: 960px;
            margin: 0 auto;
            background: #fff;
            border-radius: 12px;
            box-shadow: 0 2px 16px rgba(0,0,0,0.07);
            padding: 2.5rem 3rem;
        }}

        /* ===== Headings ===== */
        h1 {{
            font-size: 2rem;
            color: #1a5276;
            border-bottom: 3px solid #2980b9;
            padding-bottom: 0.6rem;
            margin-bottom: 1.5rem;
        }}
        h2 {{
            font-size: 1.5rem;
            color: #21618c;
            margin-top: 2rem;
            margin-bottom: 1rem;
            padding-bottom: 0.4rem;
            border-bottom: 2px solid #aed6f1;
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

        /* ===== Paragraphs ===== */
        p {{
            margin-bottom: 0.8rem;
        }}

        /* ===== Links ===== */
        a {{
            color: #2980b9;
            text-decoration: none;
        }}
        a:hover {{
            text-decoration: underline;
        }}

        /* ===== Lists ===== */
        ul, ol {{
            margin: 0.8rem 1.5rem 0.8rem 0;
            padding-right: 1.5rem;
        }}
        li {{
            margin-bottom: 0.4rem;
        }}

        /* ===== Tables ===== */
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
        th, td {{
            padding: 0.6rem 1rem;
            border: 1px solid #d5dbdb;
            text-align: right;
        }}
        tbody tr:nth-child(even) {{
            background: #eaf2f8;
        }}
        tbody tr:hover {{
            background: #d4e6f1;
        }}

        /* ===== Code ===== */
        code {{
            background: #eaf2f8;
            color: #c0392b;
            padding: 0.15rem 0.4rem;
            border-radius: 4px;
            font-family: Consolas, 'Courier New', monospace;
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
        }}
        pre code {{
            background: transparent;
            color: inherit;
            padding: 0;
        }}

        /* ===== Mermaid Diagrams ===== */
        pre.mermaid {{
            background: #fff;
            color: #2c3e50;
            text-align: center;
            direction: ltr;
            margin: 1rem auto;
            padding: 1rem;
            border-radius: 8px;
            border: 1px solid #e2e8f0;
            overflow-x: auto;
        }}
        pre.mermaid svg {{
            max-width: 100% !important;
            height: auto !important;
        }}

        /* ===== Folder Structure ===== */
        .folder-structure {{
            background: #1b2631;
            color: #58d68d;
            font-size: 0.9rem;
        }}

        /* ===== Blockquote ===== */
        blockquote {{
            border-right: 4px solid #2980b9;
            padding: 0.8rem 1.2rem;
            margin: 1rem 0;
            background: #eaf2f8;
            border-radius: 0 8px 8px 0;
        }}

        /* ===== Info Box ===== */
        .info-box {{
            background: #fef9e7;
            border: 1px solid #f9e79f;
            border-right: 4px solid #f39c12;
            padding: 1rem 1.2rem;
            border-radius: 0 8px 8px 0;
            margin: 1rem 0;
        }}
        .info-box strong {{
            color: #d68910;
            display: block;
            margin-bottom: 0.4rem;
        }}

        /* ===== Sections ===== */
        .module-section {{
            margin-bottom: 2rem;
            padding: 1rem 0;
        }}

        /* ===== HR ===== */
        hr {{
            border: none;
            border-top: 2px solid #d5dbdb;
            margin: 1.5rem 0;
        }}

        /* ===== Print ===== */
        @media print {{
            @page {{
                size: A4 portrait;
                margin: 1.2cm 1cm 1.5cm 1cm;
            }}
            
            html, body {{ 
                background: #fff !important; 
                color: #000 !important;
                padding: 0 !important;
                margin: 0 !important;
                font-size: 10pt;
                width: 100% !important;
                height: auto !important;
                overflow: visible !important;
            }}
            
            .documentation-module {{ 
                box-shadow: none !important;
                padding: 0 !important;
                max-width: 100% !important;
                margin: 0 0 1.5rem 0 !important;
                border: none !important;
                border-radius: 0 !important;
                page-break-inside: auto !important;
                break-inside: auto !important;
            }}
            
            .module-section {{
                margin-bottom: 1.5rem !important;
                padding: 0.5rem 0 !important;
                page-break-inside: auto !important;
                break-inside: auto !important;
            }}
            
            /* Prevent orphan headings */
            h1, h2, h3, h4, h5 {{
                page-break-after: avoid !important;
                break-after: avoid !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }}
            
            /* Mermaid Diagrams - Precise print spacing & fitting */
            pre.mermaid {{
                background: transparent !important;
                border: none !important;
                margin: 0.5rem auto !important;
                padding: 0 !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
                text-align: center !important;
                overflow: visible !important;
                max-width: 100% !important;
                width: 100% !important;
                box-shadow: none !important;
            }}
            
            pre.mermaid svg {{
                max-width: 100% !important;
                height: auto !important;
                max-height: 22cm !important;
                display: block !important;
                margin: 0 auto !important;
                overflow: visible !important;
            }}
            
            /* Tables - prevent horizontal cutoff */
            .table-wrapper {{
                overflow: visible !important;
                margin: 0.8rem 0 !important;
                width: 100% !important;
            }}
            
            table {{
                font-size: 8.5pt !important;
                width: 100% !important;
                table-layout: fixed !important;
                word-wrap: break-word !important;
            }}
            
            th, td {{
                word-break: break-word !important;
                padding: 0.4rem 0.6rem !important;
            }}
            
            tr {{
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }}
            
            thead {{
                display: table-header-group !important;
            }}
            
            /* Code blocks - prevent clipping */
            pre {{
                white-space: pre-wrap !important;
                word-wrap: break-word !important;
                word-break: break-word !important;
                font-size: 8pt !important;
                padding: 0.6rem !important;
                margin: 0.6rem 0 !important;
                border: 1px solid #d5dbdb !important;
                background: #f8f9fa !important;
                color: #2c3e50 !important;
                overflow: visible !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }}
            
            /* Images */
            img {{
                max-width: 100% !important;
                height: auto !important;
                page-break-inside: avoid !important;
            }}
        }}

        /* ===== Table of Contents ===== */
        .toc-page {{
            max-width: 960px;
            margin: 0 auto 2rem;
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            padding: 2rem;
            page-break-after: always;
        }}

        .toc-heading {{
            text-align: center;
            font-size: 2rem;
            font-weight: 700;
            color: #1a5276;
            margin-bottom: 0.3rem;
            padding-bottom: 0.7rem;
            border-bottom: 3px solid #2980b9;
        }}

        .toc-document-title {{
            text-align: center;
            color: #5d6d7e;
            font-size: 1.05rem;
            margin-bottom: 1.5rem;
        }}

        .toc-table-wrapper {{
            overflow-x: auto;
            border-radius: 12px;
            border: 1px solid #d6eaf8;
        }}

        .toc-table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            font-size: 0.98rem;
        }}

        .toc-table thead th {{
            background: #2980b9;
            color: #fff;
            padding: 0.9rem 1rem;
            text-align: center;
            font-weight: 700;
            border: none;
        }}

        .toc-table tbody td {{
            padding: 0.75rem 1rem;
            border-bottom: 1px solid #eaf2f8;
            vertical-align: middle;
        }}

        .toc-table tbody tr:last-child td {{
            border-bottom: none;
        }}

        .toc-table tbody tr {{
            cursor: pointer;
            transition: 0.2s ease;
        }}

        .toc-table tbody tr:hover {{
            background: #eaf2f8 !important;
        }}

        .toc-number-cell {{
            width: 18%;
            text-align: center;
            direction: ltr;
            font-weight: 700;
            color: #1a5276;
        }}

        .toc-title-cell {{
            width: 67%;
            text-align: right;
        }}

        .toc-page-cell {{
            width: 15%;
            text-align: center;
            font-weight: 700;
            color: #21618c;
        }}

        .toc-row-level-2 {{
            background: #f8fbfd;
            font-weight: 700;
        }}

        .toc-row-level-2 .toc-number-cell {{
            background: #d6eaf8;
        }}

        .toc-row-level-3 {{
            background: #ffffff;
        }}

        .toc-row-level-3 .toc-title-cell {{
            padding-right: 2rem;
        }}

        .toc-row-level-4 {{
            background: #fcfcfc;
            font-size: 0.95em;
        }}

        .toc-row-level-4 .toc-title-cell {{
            padding-right: 3rem;
        }}

        .toc-row-level-5 {{
            background: #f7f9f9;
            font-size: 0.92em;
        }}

        .toc-row-level-5 .toc-title-cell {{
            padding-right: 4rem;
        }}

        @media print {{
            .toc-page {{
                box-shadow: none !important;
                padding: 0 !important;
                margin: 0 !important;
                border-radius: 0 !important;
                page-break-after: always !important;
            }}

            .toc-table-wrapper {{
                overflow: visible !important;
            }}

            .toc-table {{
                font-size: 9pt !important;
            }}

            .toc-table thead th {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}

            .toc-row-level-2,
            .toc-row-level-3,
            .toc-row-level-4,
            .toc-row-level-5,
            .toc-number-cell {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
        }}

    </style>
</head>
<body>
{body_content}
</body>
</html>"""


# ===================== PDF / TOC pipeline =====================

def extract_headings(markdown_content):
    """Extract H1-H5 headings using 0.01 hierarchical numbering."""
    headings = []
    counters = [0, 0, 0, 0, 0]

    for line in markdown_content.splitlines():
        m = re.match(r'^\s*(#{1,5})\s+(.+?)\s*$', line)
        if not m:
            continue

        level = len(m.group(1))
        title = m.group(2).strip()

        # H1 is the document title and is not included as a numbered section.
        if level == 1:
            continue

        index = level - 2
        counters[index] += 1
        for j in range(index + 1, 4):
            counters[j] = 0

        # Main sections: 0.01, 0.02 ...
        # Subsections: 0.01.01, 0.01.02 ...
        parts = ["0"] + [f"{counters[j]:02d}" for j in range(index + 1)]
        number = ".".join(parts)

        headings.append({
            "level": level,
            "number": number,
            "title": title,
            "slug": slugify(title) or "section"
        })

    # Make IDs unique.
    used = {}
    for h in headings:
        base_slug = f"toc-{h['number'].replace('.', '-')}-{h['slug']}"
        count = used.get(base_slug, 0)
        used[base_slug] = count + 1
        h["id"] = base_slug if count == 0 else f"{base_slug}-{count + 1}"

    return headings

def extract_title(markdown_content, fallback):
    for line in markdown_content.splitlines():
        m = re.match(r'^\s*#\s+(.+?)\s*$', line)
        if m:
            return m.group(1).strip()
    return fallback


def add_numbering_to_html(html_content):
    """
    Add generated heading numbers without changing the original Markdown.
    H2-H5 receive numbers through CSS counters/HTML data attributes.
    """
    counters = [0, 0, 0, 0]

    def replace_heading(match):
        tag = match.group(1)
        attrs = match.group(2) or ""
        content = match.group(3)

        level = int(tag[1:])
        if level == 1:
            return match.group(0)

        index = level - 2
        counters[index] += 1
        for j in range(index + 1, len(counters)):
            counters[j] = 0

        number = ".".join(["0"] + [f"{counters[j]:02d}" for j in range(index + 1)])

        # Give every heading a deterministic anchor.
        clean = re.sub(r"<[^>]+>", "", content)
        anchor = slugify(clean) or "section"
        anchor = f"content-{number.replace('.', '-')}-{anchor}"

        return (
            f'<{tag}{attrs} id="{anchor}" '
            f'data-section-number="{number}">'
            f'<span class="section-number">{number}</span> {content}'
            f'</{tag}>'
        )

    return re.sub(
        r'<(h[2-5])([^>]*)>(.*?)</\1>',
        replace_heading,
        html_content,
        flags=re.DOTALL | re.IGNORECASE
    )


def add_table_numbering(html_content):
    """
    Add Table N captions and row indexes to every generated table.
    If the document has no tables at all, this is a harmless no-op --
    the regex below simply finds zero matches and the HTML is returned
    unchanged, so files without tables still convert normally.
    """
    table_counter = 0

    def replace_table(match):
        nonlocal table_counter
        table_counter += 1

        table_html = match.group(0)

        # Add row numbering to tbody rows.
        row_number = 0

        def row_replace(row_match):
            nonlocal row_number
            row_number += 1
            row = row_match.group(0)

            # Avoid duplicating an index if this pipeline is run twice.
            if '<td class="row-index">' in row:
                return row

            return re.sub(
                r'<tr(\s*)>',
                r'<tr\1><td class="row-index">' + str(row_number) + '</td>',
                row,
                count=1,
                flags=re.IGNORECASE
            )

        table_html = re.sub(
            r'<tr(?:\s*)>.*?</tr>',
            row_replace,
            table_html,
            flags=re.DOTALL | re.IGNORECASE
        )

        # Add index header cell.
        table_html = re.sub(
            r'(<thead>\s*<tr[^>]*>)',
            r'\1<th class="row-index-header">No.</th>',
            table_html,
            count=1,
            flags=re.IGNORECASE
        )

        caption = (
            f'<div class="table-caption">'
            f'<span>Table {table_counter}</span>'
            f'</div>'
        )

        return caption + table_html

    return re.sub(
        r'<div class="table-wrapper">.*?</div>',
        replace_table,
        html_content,
        flags=re.DOTALL | re.IGNORECASE
    )


def build_toc_html(headings, title):
    """Build a clean, colored table-based Table of Contents."""
    if not headings:
        return ""

    rows = []
    for h in headings:
        level_class = f"toc-row-level-{h['level']}"
        rows.append(
            f'<tr class="{level_class}" data-target="{h["id"]}">'
            f'<td class="toc-number-cell">{h["number"]}</td>'
            f'<td class="toc-title-cell">{format_text(h["title"])}</td>'
            f'<td class="toc-page-cell"><span class="toc-page-number">?</span></td>'
            f'</tr>'
        )

    return f"""
    <section class="toc-page" id="table-of-contents">
        <div class="toc-heading">الفهرس</div>
        <div class="toc-document-title">{format_text(title)}</div>

        <div class="toc-table-wrapper">
            <table class="toc-table">
                <thead>
                    <tr>
                        <th class="toc-col-number">الرقم</th>
                        <th>عنوان القسم</th>
                        <th class="toc-col-page">الصفحة</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(rows)}
                </tbody>
            </table>
        </div>
    </section>
    """

def build_document_html(markdown_content, source_name):
    title = extract_title(markdown_content, Path(source_name).stem)
    headings = extract_headings(markdown_content)

    # Existing converter creates the document body.
    html = convert_markdown_to_html(markdown_content)

    # Number headings and tables after conversion.
    # (add_table_numbering() is a safe no-op when the file has no tables,
    # so files without any table still get the rest of their content
    # converted normally.)
    html = add_numbering_to_html(html)
    html = add_table_numbering(html)

    # The generated heading IDs above are based on content numbering.
    # Synchronize TOC target IDs with those actual IDs.
    for h in headings:
        expected = (
            f'content-{h["number"].replace(".", "-")}-'
            f'{slugify(h["title"]) or "section"}'
        )
        h["id"] = expected

    toc = build_toc_html(headings, title)

    # Only inject a TOC page if one was actually built (i.e. the
    # document has at least one heading). Otherwise the document is
    # rendered in full without any TOC page, exactly as requested.
    if toc:
        html = html.replace(
            "<body>",
            f"<body>\n{toc}\n",
            1
        )

    # JavaScript calculates actual page positions and fills TOC page
    # numbers for every entry (H2 through H5). This only does anything
    # when a TOC page/items actually exist; on documents without
    # headings it simply finds zero .toc-item elements and does nothing.
    toc_script = r"""
<script>
(async () => {
    function updateTocPageNumbers() {
        const tocItems = document.querySelectorAll(".toc-table tbody tr[data-target]");

        tocItems.forEach(item => {
            const target = document.getElementById(item.dataset.target);
            const pageElement = item.querySelector(".toc-page-number");

            if (!target || !pageElement) return;

            // Chromium print layout is A4 at 96 CSS px/in.
            // PDF @page margin is 1.2cm top + 1.5cm bottom.
            // We use the rendered vertical position to calculate the page.
            const y = target.getBoundingClientRect().top + window.scrollY;

            // Approximate printable page height in CSS pixels.
            // A4 = 297mm = 1122.52px. Margins = 1.2cm + 1.5cm.
            const pageHeight = 1122.52 - ((12 + 15) * 96 / 25.4);

            // Account for the TOC being physically before the content.
            // Page numbers are finalized by the PDF generation pass below.
            const page = Math.max(1, Math.floor(y / pageHeight) + 1);

            pageElement.textContent = page;
        });
    }

    if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
    }

    updateTocPageNumbers();
    window.__TOC_READY__ = true;
})();

    document.querySelectorAll(".toc-table tbody tr[data-target]").forEach(row => {
        row.addEventListener("click", () => {
            const target = document.getElementById(row.dataset.target);
            if (target) target.scrollIntoView({behavior: "smooth", block: "start"});
        });
    });
</script>
"""

    html = html.replace("</body>", toc_script + "\n</body>", 1)
    return html, headings


def create_pdf(pdf_path, html_content):
    from playwright.sync_api import sync_playwright

    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(
            viewport={"width": 794, "height": 1123},
            device_scale_factor=1
        )

        # Set the complete HTML directly.
        page.set_content(html_content, wait_until="networkidle")

        # Render Mermaid diagrams in Chromium.
        page.add_script_tag(
            url="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"
        )

        page.wait_for_function(
            """() => typeof mermaid !== 'undefined'"""
        )

        page.evaluate("""
        async () => {
            try {
                mermaid.initialize({
                    startOnLoad: false,
                    securityLevel: 'loose',
                    theme: 'default'
                });

                const nodes = document.querySelectorAll('pre.mermaid');
                let i = 0;

                for (const node of nodes) {
                    const source = node.textContent;
                    const id = 'mermaid-generated-' + (++i);

                    try {
                        const result = await mermaid.render(id, source);
                        node.innerHTML = result.svg;
                        node.removeAttribute('data-processed');
                    } catch (error) {
                        console.warn('Mermaid render failed:', error);
                    }
                }
            } catch (error) {
                console.warn('Mermaid initialization failed:', error);
            }
        }
        """)

        # Give fonts, images and SVGs time to settle.
        page.evaluate("""
        async () => {
            if (document.fonts && document.fonts.ready) {
                await document.fonts.ready;
            }

            const images = Array.from(document.images);
            await Promise.all(images.map(img => {
                if (img.complete) return Promise.resolve();
                return new Promise(resolve => {
                    img.addEventListener('load', resolve, {once:true});
                    img.addEventListener('error', resolve, {once:true});
                });
            }));
        }
        """)

        # Wait for Mermaid SVGs to be present.
        page.wait_for_timeout(500)

        # Update TOC page labels after final layout (no-op if there is
        # no TOC on this document).
        page.evaluate("""
        () => {
            const tocItems = document.querySelectorAll(".toc-table tbody tr[data-target]");
            const pageHeight = 1122.52 - ((12 + 15) * 96 / 25.4);

            tocItems.forEach(item => {
                const target = document.getElementById(item.dataset.target);
                const pageElement = item.querySelector(".toc-page-number");

                if (!target || !pageElement) return;

                const y = target.getBoundingClientRect().top + window.scrollY;
                const page = Math.max(1, Math.floor(y / pageHeight) + 1);
                pageElement.textContent = page;
            });
        }
        """)

        page.pdf(
            path=str(pdf_path),
            format="A4",
            print_background=True,
            prefer_css_page_size=True,
            margin={
                "top": "1.2cm",
                "right": "1cm",
                "bottom": "1.5cm",
                "left": "1cm"
            },
            display_header_footer=False
        )

        browser.close()


def process_file(input_file):
    input_file = Path(input_file)

    if input_file.suffix.lower() != ".md":
        raise ValueError(f"Not a Markdown file: {input_file}")

    output_dir = input_file.parent / "PDF"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nConverting: {input_file}")

    md_content = read_markdown_file(input_file)

    # Every .md file is converted regardless of whether it contains a
    # table or headings. build_document_html() already tolerates both
    # cases: add_table_numbering() is a no-op with no tables, and
    # build_toc_html() returns "" (no TOC page injected) with no
    # headings -- either way the full document content still renders.
    html_content, headings = build_document_html(
        md_content,
        input_file.name
    )

    output_pdf = output_dir / f"{input_file.stem}.pdf"

    create_pdf(output_pdf, html_content)

    print(f"  Sections: {len(headings)}"
          + ("" if headings else "  (no headings -> TOC page skipped)"))
    print(f"  PDF saved to: {output_pdf}")

    return output_pdf


def main(input_path):
    input_path = Path(input_path)

    if not input_path.exists():
        print(f"ERROR: Path does not exist: {input_path}")
        sys.exit(1)

    if input_path.is_file():
        if input_path.suffix.lower() != ".md":
            print("ERROR: Input file must be a .md file.")
            sys.exit(1)

        process_file(input_path)
        return

    if input_path.is_dir():
        md_files = sorted(
            p for p in input_path.rglob("*.md")
            if p.is_file()
        )

        if not md_files:
            print(f"No .md files found under: {input_path}")
            return

        print(f"Found {len(md_files)} Markdown files.")

        for md_file in md_files:
            try:
                process_file(md_file)
            except Exception as ex:
                print(f"  ERROR processing {md_file}: {ex}")

        print("\nDone.")
        return

    print("ERROR: Input path must be a file or folder.")
    sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(
            "Usage:\n"
            "  py generatePdfTablesAndIndexingFromMd.py <markdown-file-or-folder>\n\n"
            "Examples:\n"
            r'  py generatePdfTablesAndIndexingFromMd.py "E:\mazen\English\Authentication.md"' "\n"
            r'  py generatePdfTablesAndIndexingFromMd.py "E:\mazen\English"'
        )
        sys.exit(1)

    main(sys.argv[1])