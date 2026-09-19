import re
import os
import sys
import html
import tempfile
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from urllib.parse import quote
import fitz  # PyMuPDF
import win32com.client

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT


# ============================================================
# Configuration
# ============================================================
BLUE_ENGLISH = '3662AE'
BROWN_ENGLISH = 'AE6A36'
BODY_COLOR = '2C3E50'
HEADING_COLORS = {
    1: '1A5276',
    2: '21618C',
    3: '2E86C1',
    4: '3498DB',
}
TABLE_HEADER_COLOR = '2980B9'
TABLE_BORDER_COLOR = 'D5DBDB'
TOC_HEADER_COLOR = '2980B9'

ARABIC_RE = re.compile(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]')
LATIN_RE = re.compile(r'[A-Za-z]')
URL_RE = re.compile(r'https?://\S+$', re.IGNORECASE)
MARKDOWN_LINK_RE = re.compile(r'^\[([^\]]+)\]\(([^)]+)\)$')
INLINE_LINK_RE = re.compile(r'\[([^\]]+)\]\((https?://[^)]+)\)', re.IGNORECASE)

# Pronounce / YouGlish configuration
PRONOUNCE_TEXT = 'pronounce'
PRONOUNCE_COLOR = '5D6D7E'
PRONOUNCE_SIZE = 8.5
PRONOUNCE_TABLE_SIZE = 7.5
PRONOUNCE_SEPARATOR = '  '
YOUGLISH_BASE_URL = 'https://youglish.com/pronounce/'
YOUGLISH_SUFFIX = '/english/us'
MAX_PRONOUNCE_WORDS = 9

# Question labels such as Q1:, q1 :, Q2 : and Q12:
# Q/q + one or more digits + optional whitespace + colon.
QUESTION_MARKER_RE = re.compile(r'\b[qQ]\d+\s*:')


def add_colored_text_run(paragraph, text, *, size=11, bold=False, italic=False,
                         default_color=BODY_COLOR):
    """Add text while forcing Q<number>: markers to the configured blue.

    The question marker color is independent of the surrounding English
    color rule and works in both English and mixed Arabic/English text.
    """
    if not text:
        return

    pos = 0
    for match in QUESTION_MARKER_RE.finditer(text):
        if match.start() > pos:
            chunk = text[pos:match.start()]
            if chunk:
                run = paragraph.add_run(chunk)
                set_run_font(run, size=size, bold=bold, italic=italic, color=default_color)
        run = paragraph.add_run(match.group(0))
        set_run_font(run, size=size, bold=bold, italic=italic, color=BLUE_ENGLISH)
        pos = match.end()

    if pos < len(text):
        chunk = text[pos:]
        run = paragraph.add_run(chunk)
        set_run_font(run, size=size, bold=bold, italic=italic, color=default_color)


def read_markdown_file(file_path):
    encodings = ['utf-8-sig', 'utf-8', 'cp1256', 'windows-1256', 'utf-16']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                return f.read().lstrip('\ufeff')
        except UnicodeDecodeError:
            continue
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        return f.read().lstrip('\ufeff')


def html_escape(text):
    return html.escape(str(text), quote=False)


def normalize_url(url: str) -> str:
    return html.unescape(url).strip()


def is_english_only_text(text):
    if not text or not text.strip():
        return False
    if ARABIC_RE.search(text):
        return False
    return bool(LATIN_RE.search(text))


def english_color(text):
    return BLUE_ENGLISH if '?' in text else BROWN_ENGLISH


def add_bookmark(paragraph, name: str, bookmark_id: int):
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bookmark_id))
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def add_internal_hyperlink(paragraph, text: str, anchor: str, *, bold=False, color='2980B9', underline=False):
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), anchor)

    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    if bold:
        b = OxmlElement('w:b')
        rPr.append(b)
    color_el = OxmlElement('w:color')
    color_el.set(qn('w:val'), color)
    rPr.append(color_el)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Segoe UI')
    rFonts.set(qn('w:hAnsi'), 'Segoe UI')
    rFonts.set(qn('w:cs'), 'Arial')
    rPr.append(rFonts)

    run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_external_hyperlink(paragraph, text: str, url: str, *, color=None, underline=False, bold=False):
    url = normalize_url(url)
    rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), rid)

    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if bold:
        b = OxmlElement('w:b')
        rPr.append(b)
    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Segoe UI')
    rFonts.set(qn('w:hAnsi'), 'Segoe UI')
    rFonts.set(qn('w:cs'), 'Arial')
    rPr.append(rFonts)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_field(paragraph, instruction: str, text_fallback: str = ''):
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = instruction
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    text = OxmlElement('w:t')
    text.text = text_fallback
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(text)
    run._r.append(fld_end)
    return run


def set_run_font(run, font_name='Segoe UI', size=None, bold=None, italic=None, color=None):
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), font_name)
    rfonts.set(qn('w:hAnsi'), font_name)
    rfonts.set(qn('w:cs'), 'Arial')
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)


def set_cell_border(cell, color=TABLE_BORDER_COLOR, sz='6'):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        tag = 'w:' + edge
        el = borders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            borders.append(el)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), sz)
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), 'true')
    trPr.append(tblHeader)


def set_update_fields_on_open(doc):
    settings = doc.settings.element
    update = settings.find(qn('w:updateFields'))
    if update is None:
        update = OxmlElement('w:updateFields')
        settings.append(update)
    update.set(qn('w:val'), 'true')


def set_section_page_number_restart(section, start=1):
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    pgNumType.set(qn('w:start'), str(start))


def add_footer_page_number(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.clear()
    run = add_field(p, 'PAGE', '1')
    set_run_font(run, size=18, bold=True, color='000000')


def configure_styles(doc: Document):
    styles = doc.styles
    normal = styles['Normal']
    normal.font.name = 'Segoe UI'
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor.from_string(BODY_COLOR)
    normal._element.rPr.rFonts.set(qn('w:ascii'), 'Segoe UI')
    normal._element.rPr.rFonts.set(qn('w:hAnsi'), 'Segoe UI')
    normal._element.rPr.rFonts.set(qn('w:cs'), 'Arial')

    for level in range(1, 5):
        style = styles[f'Heading {level}']
        style.font.name = 'Segoe UI'
        style.font.size = Pt({1: 18, 2: 15, 3: 13, 4: 11}[level])
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(HEADING_COLORS[level])
        style._element.rPr.rFonts.set(qn('w:ascii'), 'Segoe UI')
        style._element.rPr.rFonts.set(qn('w:hAnsi'), 'Segoe UI')
        style._element.rPr.rFonts.set(qn('w:cs'), 'Arial')

    if 'Code Block' not in styles:
        style = styles.add_style('Code Block', WD_STYLE_TYPE.PARAGRAPH)
    else:
        style = styles['Code Block']
    style.font.name = 'Consolas'
    style.font.size = Pt(9)
    style.font.color.rgb = RGBColor.from_string(BODY_COLOR)
    style._element.rPr.rFonts.set(qn('w:ascii'), 'Consolas')
    style._element.rPr.rFonts.set(qn('w:hAnsi'), 'Consolas')
    style._element.rPr.rFonts.set(qn('w:cs'), 'Arial')


def add_mixed_text_runs(paragraph, text: str, *, force_color=None):
    """Add text while coloring English chunks outside code runs.

    English chunks containing ? are blue; other English chunks are brown.
    Q<number>: markers are always blue, both in English and mixed text.
    """
    text = str(text)
    chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', text)
    for chunk in chunks:
        if not chunk:
            continue

        if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
            color = force_color or english_color(chunk)

            # Split Q<number>: markers out so they are always blue.
            pos = 0
            for match in QUESTION_MARKER_RE.finditer(chunk):
                if match.start() > pos:
                    before = chunk[pos:match.start()]
                    run = paragraph.add_run(before)
                    set_run_font(run, color=color, bold=True, size=11)
                run = paragraph.add_run(match.group(0))
                set_run_font(run, color=BLUE_ENGLISH, bold=True, size=11)
                pos = match.end()

            if pos < len(chunk):
                run = paragraph.add_run(chunk[pos:])
                set_run_font(run, color=color, bold=True, size=11)
        else:
            # A marker can appear inside mixed text, so handle it here too.
            pos = 0
            for match in QUESTION_MARKER_RE.finditer(chunk):
                if match.start() > pos:
                    run = paragraph.add_run(chunk[pos:match.start()])
                    set_run_font(run, color=BODY_COLOR, size=11)
                run = paragraph.add_run(match.group(0))
                set_run_font(run, color=BLUE_ENGLISH, bold=False, size=11)
                pos = match.end()

            if pos < len(chunk):
                run = paragraph.add_run(chunk[pos:])
                set_run_font(run, color=BODY_COLOR, size=11)

def parse_inline_segments(text: str):
    """Return sequence of (kind, value, optional_url)."""
    text = text.replace('\u00a0', ' ')
    segments = []
    pos = 0

    # Markdown links first.
    combined = re.compile(r'\[([^\]]+)\]\((https?://[^)]+)\)|`([^`]+)`|\*\*([^*]+)\*\*|\*([^*]+)\*', re.IGNORECASE)
    for m in combined.finditer(text):
        if m.start() > pos:
            segments.append(('text', text[pos:m.start()], None))
        if m.group(1) is not None:
            segments.append(('link', m.group(1), normalize_url(m.group(2))))
        elif m.group(3) is not None:
            segments.append(('code', m.group(3), None))
        elif m.group(4) is not None:
            segments.append(('bold', m.group(4), None))
        else:
            segments.append(('italic', m.group(5), None))
        pos = m.end()
    if pos < len(text):
        segments.append(('text', text[pos:], None))
    return segments


def visible_text_for_pronounce(text: str) -> str:
    """Return the visible text that should be used to build a YouGlish query.

    Markdown syntax and existing link URLs are excluded; only visible labels/content
    remain. The original text is never modified.
    """
    parts = []
    for kind, value, _ in parse_inline_segments(str(text)):
        if kind in ('text', 'link', 'code', 'bold', 'italic'):
            parts.append(value)
    return ' '.join(parts)


def build_pronounce_query(text: str) -> str:
    """Build the YouGlish query from words only, capped at 9 words.
    
    IMPORTANT: Remove ALL URLs before extracting words to ensure the query
    contains only natural English words, no URL fragments or domains.
    """
    visible = visible_text_for_pronounce(text)

    # Q-number labels are document markers, not part of the spoken sentence.
    visible = QUESTION_MARKER_RE.sub(' ', visible)

    # Remove ALL URLs before extracting words to prevent any URL fragments
    # from entering the query. This removes the entire URL including:
    # - http:// and https://
    # - domain names (example.com)
    # - path components
    # - query parameters
    visible = re.sub(r'https?://\S+', ' ', visible)
    
    # Also remove URLs that might appear without http:// (just domains)
    visible = re.sub(r'\b[a-zA-Z0-9-]+\.[a-zA-Z]{2,}\b', ' ', visible)

    # Keep English words only. This removes quotes, punctuation and other symbols
    # from the URL version while leaving the displayed document text unchanged.
    words = re.findall(r'[A-Za-z]+', visible)
    if not words:
        return ''

    # Additional safety: filter out any remaining words that might be URL-related
    words = [word for word in words if word.lower() not in ['http', 'https', 'www']]
    if not words:
        return ''

    words = words[:MAX_PRONOUNCE_WORDS]
    query = '_'.join(words)
    
    # Final safety check: ensure the query doesn't contain http/https
    if 'http' in query.lower() or 'https' in query.lower():
        return ''
    
    return quote(query, safe='_')


def build_pronounce_url(text: str) -> Optional[str]:
    """Return the YouGlish URL for text, or None when text is not a candidate."""
    visible = visible_text_for_pronounce(text)
    if not visible.strip() or not is_english_only_text(visible):
        return None

    query = build_pronounce_query(visible)
    if not query:
        return None

    return f'{YOUGLISH_BASE_URL}{query}{YOUGLISH_SUFFIX}'


def add_pronounce_hyperlink(paragraph, source_text: str, *, size=PRONOUNCE_SIZE,
                            color=PRONOUNCE_COLOR, bold=False) -> bool:
    """Pronounce links disabled - function does nothing."""
    return False


def add_external_hyperlink_before(paragraph, text: str, url: str, before_element=None,
                                  *, color=None, underline=False, bold=False):
    """Create an external hyperlink and insert it before a specific XML element."""
    url = normalize_url(url)
    rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), rid)

    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if bold:
        b = OxmlElement('w:b')
        rPr.append(b)
    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Segoe UI')
    rFonts.set(qn('w:hAnsi'), 'Segoe UI')
    rFonts.set(qn('w:cs'), 'Arial')
    rPr.append(rFonts)

    run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hyperlink.append(run)

    if before_element is not None:
        paragraph._p.insert(paragraph._p.index(before_element), hyperlink)
    else:
        paragraph._p.append(hyperlink)
    return hyperlink


def add_pronounce_after_existing_hyperlink(paragraph, source_text: str, *, size=PRONOUNCE_SIZE,
                                            color=PRONOUNCE_COLOR, bold=False) -> bool:
    """Pronounce links disabled - function does nothing."""
    return False


def add_rich_paragraph(doc, text: str, *, style=None, alignment=WD_ALIGN_PARAGRAPH.RIGHT, hyperlink_url: Optional[str] = None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_after = Pt(7)

    if hyperlink_url:
        # The visible label is the entire previous sentence/text.
        label = text.strip()
        add_external_hyperlink(p, label, hyperlink_url, color=english_color(label), bold=is_english_only_text(label))
        add_pronounce_after_existing_hyperlink(
            p, label,
            bold=False
        )
        return p

    # Check if the text contains "= URL" pattern first
    url_pattern = re.compile(r'\s*=\s*(https?://[^\s]+)', re.IGNORECASE)
    url_pattern_no_space = re.compile(r'\s*=(https?://[^\s]+)', re.IGNORECASE)
    
    url_matches = list(url_pattern.finditer(text))
    
    # If no matches with space, try without space
    if not url_matches:
        url_matches = list(url_pattern_no_space.finditer(text))
    
    if url_matches:
        # Handle "= URL" pattern in paragraphs
        previous_url_end = 0
        
        for match in url_matches:
            url = normalize_url(match.group(1))
            
            # Text before the "= URL"
            label = text[previous_url_end:match.start()].strip()
            
            if label:
                # Check if the entire label is English only
                if is_english_only_text(label):
                    # English text gets hyperlink
                    add_external_hyperlink(p, label, url, color=english_color(label), underline=False, bold=True)
                else:
                    # Split label into English and Arabic parts
                    chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', label)
                    
                    for chunk in chunks:
                        if not chunk:
                            continue
                        
                        if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                            # English text gets hyperlink
                            add_external_hyperlink(p, chunk, url, color=english_color(chunk), underline=False, bold=True)
                        else:
                            # Arabic text remains normal
                            add_mixed_text_runs(p, chunk)
            
            previous_url_end = match.end()
        
        # Any text after the final URL
        remaining = text[previous_url_end:].strip()
        if remaining:
            add_mixed_text_runs(p, remaining)
    else:
        # Handle normal markdown links and text
        for kind, value, url in parse_inline_segments(text):
            if kind == 'text':
                add_mixed_text_runs(p, value)
            elif kind == 'link':
                # IMPORTANT: Only English text gets the hyperlink
                if is_english_only_text(value):
                    color = english_color(value)
                    add_external_hyperlink(p, value, url, color=color, underline=False, bold=True)
                else:
                    # Split the link text into English and Arabic parts
                    chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', value)
                    
                    for chunk in chunks:
                        if not chunk:
                            continue
                        
                        if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                            # English text gets hyperlink
                            color = english_color(chunk)
                            add_external_hyperlink(p, chunk, url, color=color, underline=False, bold=True)
                        else:
                            # Arabic text remains normal
                            add_mixed_text_runs(p, chunk)
            elif kind == 'code':
                r = p.add_run(value)
                set_run_font(r, 'Consolas', size=9, color='C0392B')
            elif kind == 'bold':
                for subkind, subvalue, _ in parse_inline_segments(value):
                    if subkind == 'text':
                        chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', subvalue)
                        for chunk in chunks:
                            if not chunk:
                                continue
                            color = english_color(chunk) if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk) else BODY_COLOR
                            r = p.add_run(chunk)
                            set_run_font(r, color=color, bold=True, size=11)
                    else:
                        # fallback for nested code/link
                        r = p.add_run(subvalue)
                        set_run_font(r, bold=True, size=11, color=BODY_COLOR)
            elif kind == 'italic':
                chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', value)
                for chunk in chunks:
                    if not chunk:
                        continue
                    color = english_color(chunk) if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk) else BODY_COLOR
                    r = p.add_run(chunk)
                    set_run_font(r, color=color, italic=True, size=11)

    # add_pronounce_hyperlink(p, text)  # Pronounce disabled
    return p


def make_heading_number(counters, level):
    idx = level - 2
    counters[idx] += 1
    for j in range(idx + 1, len(counters)):
        counters[j] = 0
    return '.'.join(['0'] + [f'{counters[j]:02d}' for j in range(idx + 1)])


def extract_headings(markdown_content):
    headings = []
    counters = [0, 0, 0, 0]
    used = {}
    for line in markdown_content.splitlines():
        m = re.match(r'^\s*(#{2,5})\s+(.+?)\s*$', line)
        if not m:
            continue
        level = len(m.group(1))
        title = m.group(2).strip()
        number = make_heading_number(counters, level)
        base = f'heading-{number.replace(".", "-")}-{slugify(title) or "section"}'
        count = used.get(base, 0) + 1
        used[base] = count
        bookmark = base if count == 1 else f'{base}-{count}'
        headings.append({'level': level, 'number': number, 'title': title, 'bookmark': bookmark})
    return headings


def slugify(text):
    text = (text or '').strip().lower()
    text = re.sub(r'[^\w\s-]', '', text, flags=re.UNICODE)
    return re.sub(r'[\s_]+', '-', text).strip('-')


def parse_table_block(lines):
    rows = []
    for line in lines:
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        rows.append(cells)
    return rows


def add_table_mixed_text_runs(paragraph, text: str, *, bold=False, size=9, default_color=BODY_COLOR):
    """Render table text with slightly smaller/lighter English styling.

    English text uses the same configured blue/brown colors as the document;
    Q<number>: markers are always blue. Arabic remains the normal body color.
    """
    text = str(text)
    chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', text)

    for chunk in chunks:
        if not chunk:
            continue

        is_english = bool(LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk))
        base_color = english_color(chunk) if is_english else default_color

        pos = 0
        for match in QUESTION_MARKER_RE.finditer(chunk):
            if match.start() > pos:
                before = chunk[pos:match.start()]
                run = paragraph.add_run(before)
                set_run_font(
                    run,
                    size=size,
                    bold=bold if not is_english else False,
                    color=base_color
                )
            run = paragraph.add_run(match.group(0))
            set_run_font(run, size=size, bold=bold if not is_english else False, color=BLUE_ENGLISH)
            pos = match.end()

        if pos < len(chunk):
            run = paragraph.add_run(chunk[pos:])
            set_run_font(
                run,
                size=size,
                bold=bold if not is_english else False,
                color=base_color
            )


def add_table(doc, rows, table_index):
    if not rows:
        return
    # Skip markdown separator row.
    body_rows = rows
    if len(rows) >= 2 and all(set(cell.strip()) <= set('-: ') for cell in rows[1]):
        body_rows = [rows[0]] + rows[2:]

    cols = max(len(r) for r in body_rows)
    table = doc.add_table(rows=len(body_rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    table.autofit = True

    caption = doc.add_paragraph()
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = caption.add_run(f'Table {table_index}')
    set_run_font(r, size=10, bold=True, color='1A5276')

    for ri, row in enumerate(body_rows):
        for ci in range(cols):
            cell = table.cell(ri, ci)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_border(cell)
            if ri == 0:
                set_cell_shading(cell, TABLE_HEADER_COLOR)
            else:
                if ri % 2 == 0:
                    set_cell_shading(cell, 'EAF2F8')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.clear()
            value = row[ci] if ci < len(row) else ''
            
            # First, check if the cell contains "= URL" pattern
            url_pattern = re.compile(r'\s*=\s*(https?://[^\s]+)', re.IGNORECASE)
            url_pattern_no_space = re.compile(r'\s*=(https?://[^\s]+)', re.IGNORECASE)
            
            url_matches = list(url_pattern.finditer(value))
            
            # If no matches with space, try without space
            if not url_matches:
                url_matches = list(url_pattern_no_space.finditer(value))
            
            if url_matches:
                # Handle "= URL" pattern in table cells
                previous_url_end = 0
                
                for match in url_matches:
                    url = normalize_url(match.group(1))
                    
                    # Text before the "= URL"
                    label = value[previous_url_end:match.start()].strip()
                    
                    if label:
                        # Check if the entire label is English only
                        if is_english_only_text(label):
                            # English text gets hyperlink
                            link_color = 'FFFFFF' if ri == 0 else (
                                BLUE_ENGLISH if QUESTION_MARKER_RE.search(label)
                                else english_color(label)
                            )
                            add_external_hyperlink(
                                p,
                                label,
                                url,
                                color=link_color,
                                underline=False,
                                bold=(ri == 0)
                            )
                        else:
                            # Split label into English and Arabic parts
                            chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', label)
                            
                            for chunk in chunks:
                                if not chunk:
                                    continue
                                
                                if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                                    # English text gets hyperlink
                                    link_color = 'FFFFFF' if ri == 0 else (
                                        BLUE_ENGLISH if QUESTION_MARKER_RE.search(chunk)
                                        else english_color(chunk)
                                    )
                                    add_external_hyperlink(
                                        p,
                                        chunk,
                                        url,
                                        color=link_color,
                                        underline=False,
                                        bold=(ri == 0)
                                    )
                                else:
                                    # Arabic text remains normal
                                    add_table_mixed_text_runs(
                                        p,
                                        chunk,
                                        bold=(ri == 0),
                                        size=9,
                                        default_color='FFFFFF' if ri == 0 else BODY_COLOR
                                    )
                    
                    previous_url_end = match.end()
                
                # Any text after the final URL
                remaining = value[previous_url_end:].strip()
                if remaining:
                    add_table_mixed_text_runs(
                        p,
                        remaining,
                        bold=(ri == 0),
                        size=9,
                        default_color='FFFFFF' if ri == 0 else BODY_COLOR
                    )
            else:
                # Handle normal markdown links
                for kind, val, url in parse_inline_segments(value):
                    if kind == 'text':
                        add_table_mixed_text_runs(
                            p,
                            val,
                            bold=(ri == 0),
                            size=9,
                            default_color='FFFFFF' if ri == 0 else BODY_COLOR
                        )
                    elif kind == 'link':
                        # IMPORTANT: Only English text gets the hyperlink
                        if is_english_only_text(val):
                            link_color = 'FFFFFF' if ri == 0 else (
                                BLUE_ENGLISH if QUESTION_MARKER_RE.search(val)
                                else english_color(val)
                            )
                            add_external_hyperlink(
                                p,
                                val,
                                url,
                                color=link_color,
                                underline=False,
                                bold=(ri == 0)
                            )
                        else:
                            # Split the link text into English and Arabic parts
                            chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', val)
                            
                            for chunk in chunks:
                                if not chunk:
                                    continue
                                
                                if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                                    # English text gets hyperlink
                                    link_color = 'FFFFFF' if ri == 0 else (
                                        BLUE_ENGLISH if QUESTION_MARKER_RE.search(chunk)
                                        else english_color(chunk)
                                    )
                                    add_external_hyperlink(
                                        p,
                                        chunk,
                                        url,
                                        color=link_color,
                                        underline=False,
                                        bold=(ri == 0)
                                    )
                                else:
                                    # Arabic text remains normal
                                    add_table_mixed_text_runs(
                                        p,
                                        chunk,
                                        bold=(ri == 0),
                                        size=9,
                                        default_color='FFFFFF' if ri == 0 else BODY_COLOR
                                    )
                    elif kind == 'code':
                        rr = p.add_run(val)
                        set_run_font(rr, 'Consolas', size=8, color='FFFFFF' if ri == 0 else BODY_COLOR)
                    else:
                        add_table_mixed_text_runs(
                            p,
                            val,
                            bold=(ri == 0),
                            size=9,
                            default_color='FFFFFF' if ri == 0 else BODY_COLOR
                        )
            # Add pronounce once per cell, after all inline segments are rendered.
            # add_pronounce_hyperlink disabled - entire function call removed
    set_repeat_table_header(table.rows[0])
    return table


def add_info_box(doc, title, body_lines):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    cell = table.cell(0, 0)
    set_cell_shading(cell, 'FEF9E7')
    set_cell_border(cell, color='F9E79F', sz='8')
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(title)
    set_run_font(r, size=11, bold=True, color='D68910')
    for line in body_lines:
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        add_mixed_text_runs(p, line)
        add_pronounce_hyperlink(p, line)
    doc.add_paragraph()


def add_quote(doc, text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_shading(cell, 'EAF2F8')
    set_cell_border(cell, color='2980B9', sz='10')
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_mixed_text_runs(p, text)
    # add_pronounce_hyperlink(p, text)  # Pronounce disabled
    doc.add_paragraph()


def add_list_item(doc, text, ordered=False, level=0):
    style = 'List Number' if ordered else 'List Bullet'
    p = doc.add_paragraph(style=style)
    p.paragraph_format.left_indent = Inches(0.25 * level)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # Check if the text contains "= URL" pattern
    url_pattern = re.compile(r'\s*=\s*(https?://[^\s]+)', re.IGNORECASE)
    url_pattern_no_space = re.compile(r'\s*=(https?://[^\s]+)', re.IGNORECASE)
    
    url_matches = list(url_pattern.finditer(text))
    
    # If no matches with space, try without space
    if not url_matches:
        url_matches = list(url_pattern_no_space.finditer(text))
    
    if url_matches:
        # Handle "= URL" pattern in list items
        previous_url_end = 0
        
        for match in url_matches:
            url = normalize_url(match.group(1))
            
            # Text before the "= URL"
            label = text[previous_url_end:match.start()].strip()
            
            if label:
                # Check if the entire label is English only
                if is_english_only_text(label):
                    # English text gets hyperlink
                    add_external_hyperlink(p, label, url, color=english_color(label), underline=False, bold=True)
                else:
                    # Split label into English and Arabic parts
                    chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', label)
                    
                    for chunk in chunks:
                        if not chunk:
                            continue
                        
                        if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                            # English text gets hyperlink
                            add_external_hyperlink(p, chunk, url, color=english_color(chunk), underline=False, bold=True)
                        else:
                            # Arabic text remains normal
                            add_mixed_text_runs(p, chunk)
            
            previous_url_end = match.end()
        
        # Any text after the final URL
        remaining = text[previous_url_end:].strip()
        if remaining:
            add_mixed_text_runs(p, remaining)
    else:
        # Handle normal text without "= URL" pattern
        add_mixed_text_runs(p, text)
    
    # add_pronounce_hyperlink(p, text)  # Pronounce disabled
    return p


def add_code_block(doc, code, language='text'):
    """
    Add code block WITHOUT pronounce hyperlink.
    
    Code blocks should not have pronounce links as they contain
    programming code/syntax, not natural language text.
    """
    p = doc.add_paragraph(style='Code Block')
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(7)
    r = p.add_run(code)
    set_run_font(r, 'Consolas', size=8.5, color=BODY_COLOR)
    # Add a subtle shading/border to the paragraph.
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'F8F9FA')
    pPr.append(shd)
    return p


def render_mermaid_to_png(raw_code, output_png):
    """Best-effort Mermaid -> PNG using Playwright + Mermaid CDN.
    The Word conversion remains functional when rendering fails; a readable
    fallback code block is inserted instead.
    """
    try:
        from playwright.sync_api import sync_playwright
        import cairosvg

        html_doc = f'''<!doctype html><html><head><meta charset="utf-8"></head><body>
        <pre class="mermaid">{html.escape(raw_code)}</pre>
        <script src="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"></script>
        </body></html>'''
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={'width': 1400, 'height': 1000})
            page.set_content(html_doc, wait_until='networkidle')
            page.wait_for_function("() => typeof mermaid !== 'undefined'")
            page.evaluate('''async () => {
              mermaid.initialize({startOnLoad:false, securityLevel:'loose', theme:'default'});
              const node = document.querySelector('.mermaid');
              const result = await mermaid.render('word-mermaid', node.textContent);
              node.innerHTML = result.svg;
            }''')
            svg = page.locator('svg').evaluate('(el) => el.outerHTML')
            browser.close()
        cairosvg.svg2png(bytestring=svg.encode('utf-8'), write_to=str(output_png), output_width=1400)
        return True
    except Exception as exc:
        print(f'  WARNING: Mermaid could not be rendered into Word image: {exc}')
        return False


def add_toc_page(doc, title, headings):
    # Not a Heading style: should not pollute the Navigation pane.
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('الفهرس')
    set_run_font(r, size=22, bold=True, color='1A5276')

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    set_run_font(r, size=11, color='5D6D7E')

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    headers = [('الرقم', 0), ('عنوان القسم', 1), ('الصفحة', 2)]
    for cell, (label, _) in zip(hdr, headers):
        set_cell_shading(cell, TOC_HEADER_COLOR)
        set_cell_border(cell, color=TOC_HEADER_COLOR, sz='8')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.clear()
        r = p.add_run(label)
        set_run_font(r, size=10, bold=True, color='FFFFFF')

    level_to_indent = {2: 0, 3: 1, 4: 2, 5: 3}
    for h in headings:
        cells = table.add_row().cells
        for c in cells:
            set_cell_border(c)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        # number
        p = cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.clear()
        r = p.add_run(h['number'])
        set_run_font(r, size=9, bold=True, color='1A5276')

        # title and internal link
        p = cells[1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.clear()
        indent = level_to_indent.get(h['level'], 0)
        p.paragraph_format.right_indent = Inches(0.35 * indent)
        add_internal_hyperlink(p, h['title'], h['bookmark'], bold=(h['level'] == 2), color='21618C')

        # PAGE field referencing bookmark
        p = cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.clear()
        run = add_field(p, f'PAGEREF {h["bookmark"]} \\h', '?')
        set_run_font(run, size=9, bold=True, color='21618C')

    doc.add_paragraph()


def build_document(markdown_content, source_name):
    title = extract_title(markdown_content, Path(source_name).stem)
    headings = extract_headings(markdown_content)
    doc = Document()
    configure_styles(doc)
    set_update_fields_on_open(doc)

    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    section.different_first_page_header_footer = False

    # First page: document title only (TOC removed).
    # TOC table is removed while navigation/bookmarks still work.
    if headings:
        # add_toc_page(doc, title, headings)  # TOC removed
        pass

        # Configure margins for the main content section (no new page needed)
        content_section = doc.sections[0]
        content_section.top_margin = Inches(0.65)
        content_section.bottom_margin = Inches(0.75)
        content_section.left_margin = Inches(0.7)
        content_section.right_margin = Inches(0.7)
        content_section.footer.is_linked_to_previous = False
        set_section_page_number_restart(content_section, 1)
        add_footer_page_number(content_section)

        # Restore footer page numbering (TOC removed but footer numbering needed)
        doc.sections[0].footer.is_linked_to_previous = False
        add_footer_page_number(doc.sections[0])
    else:
        # No TOC: render the document in one section and start numbering at 1.
        content_section = doc.sections[0]
        set_section_page_number_restart(content_section, 1)
        add_footer_page_number(content_section)

    # Rendering state
    heading_iter = iter(headings)
    counters = [0, 0, 0, 0]
    table_index = 0
    lines = markdown_content.splitlines()
    i = 0
    n = len(lines)
    bookmark_id = 100
    current_heading_index = 0

    while i < n:
        raw = lines[i]
        stripped = raw.strip()

        if not stripped:
            i += 1
            continue

        # H1 is document title and not repeated in body.
        if re.match(r'^#\s+', stripped) and not re.match(r'^##\s+', stripped):
            i += 1
            continue

        # Heading H2-H5
        hm = re.match(r'^(#{2,5})\s+(.+?)\s*$', stripped)
        if hm:
            level = len(hm.group(1))
            text = hm.group(2).strip()
            number = make_heading_number(counters, level)
            if current_heading_index < len(headings):
                h = headings[current_heading_index]
                current_heading_index += 1
            else:
                h = {'level': level, 'number': number, 'title': text, 'bookmark': f'heading-{bookmark_id}'}
            style_name = f'Heading {level - 1}'
            p = doc.add_paragraph(style=style_name)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(12 if level == 2 else 8)
            p.paragraph_format.space_after = Pt(5)
            add_bookmark(p, h['bookmark'], bookmark_id)
            bookmark_id += 1
            rn = p.add_run(h['number'] + ' ')
            set_run_font(rn, size={2: 15, 3: 13, 4: 11, 5: 10}[level], bold=True, color=HEADING_COLORS[level - 1])
            add_mixed_text_runs(p, text, force_color=None)
            # add_pronounce_hyperlink disabled
            i += 1
            continue

        # Code block
        if stripped.startswith('```'):
            language = stripped[3:].strip() or 'text'
            code_lines = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            raw_code = '\n'.join(code_lines).strip('\n')
            if language.lower() == 'mermaid':
                with tempfile.TemporaryDirectory() as td:
                    png = Path(td) / 'mermaid.png'
                    if render_mermaid_to_png(raw_code, png):
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p.add_run()
                        r.add_picture(str(png), width=Inches(6.4))
                    else:
                        add_code_block(doc, raw_code, language)
            else:
                add_code_block(doc, raw_code, language)
            i += 1
            continue

        # Markdown table
        if '|' in stripped and i + 1 < n and '|' in lines[i + 1] and set(lines[i + 1].strip()) - set('|-: ') == set():
            table_lines = []
            while i < n and '|' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            table_index += 1
            add_table(doc, parse_table_block(table_lines), table_index)
            continue

        # Blockquote
        if stripped.startswith('> '):
            quote_lines = []
            while i < n and lines[i].strip().startswith('> '):
                quote_lines.append(lines[i].strip()[2:])
                i += 1
            add_quote(doc, ' '.join(quote_lines))
            continue

        # Lists
        if stripped.startswith(('- ', '* ')) or re.match(r'^\d+\. ', stripped):
            ordered = bool(re.match(r'^\d+\. ', stripped))
            while i < n:
                candidate = lines[i]
                cs = candidate.strip()
                if not cs:
                    break
                m = re.match(r'^[-*]\s+(.+)$', cs) if not ordered else re.match(r'^\d+\.\s+(.+)$', cs)
                if m:
                    add_list_item(doc, m.group(1), ordered=ordered)
                    i += 1
                elif candidate.startswith('  ') or candidate.startswith('\t'):
                    # Continuation line: append as normal indented bullet text.
                    add_list_item(doc, cs, ordered=ordered, level=1)
                    i += 1
                else:
                    break
            continue

        # Info boxes
        if stripped in ['Important Notes', 'Warnings', 'Tips', 'Recommendations']:
            box_title = stripped
            body = []
            i += 1
            while i < n and lines[i].strip():
                body.append(lines[i].strip())
                i += 1
            add_info_box(doc, box_title, body)
            continue

        # Multiple hyperlinks in ONE line: Label = URL Label = URL
        if re.search(r'=\s*https?://[^\s]+', stripped, re.IGNORECASE) or re.search(r'=https?://[^\s]+', stripped, re.IGNORECASE):
            # Handle multiple URLs in one line with English/Arabic distinction
            url_pattern = re.compile(r'\s*=\s*(https?://[^\s]+)', re.IGNORECASE)
            url_pattern_no_space = re.compile(r'\s*=(https?://[^\s]+)', re.IGNORECASE)
            
            url_matches = list(url_pattern.finditer(stripped))
            
            # If no matches with space, try without space
            if not url_matches:
                url_matches = list(url_pattern_no_space.finditer(stripped))
            
            if url_matches:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_after = Pt(7)
                
                previous_url_end = 0
                
                for index, match in enumerate(url_matches):
                    url = normalize_url(match.group(1))
                    
                    # Text between previous URL and current "= URL"
                    label = stripped[previous_url_end:match.start()].strip()
                    
                    if label:
                        # Check if the entire label is English only
                        if is_english_only_text(label):
                            # English text gets hyperlink
                            if index > 0:
                                spacer = p.add_run(' ')
                                set_run_font(spacer, size=11, color=BODY_COLOR)
                            
                            add_external_hyperlink(p, label, url, color=english_color(label), underline=False, bold=True)
                        else:
                            # Split label into English and Arabic parts
                            chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', label)
                            
                            first_chunk = True
                            for chunk in chunks:
                                if not chunk:
                                    continue
                                
                                if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                                    # English text gets hyperlink
                                    if not first_chunk:
                                        spacer = p.add_run(' ')
                                        set_run_font(spacer, size=11, color=BODY_COLOR)
                                    
                                    add_external_hyperlink(p, chunk, url, color=english_color(chunk), underline=False, bold=True)
                                    first_chunk = False
                                else:
                                    # Arabic text remains normal
                                    if not first_chunk:
                                        spacer = p.add_run(' ')
                                        set_run_font(spacer, size=11, color=BODY_COLOR)
                                    add_mixed_text_runs(p, chunk)
                                    first_chunk = False
                    
                    previous_url_end = match.end()
                
                # Any text after the final URL
                remaining = stripped[previous_url_end:].strip()
                if remaining:
                    spacer = p.add_run(' ')
                    set_run_font(spacer, size=11, color=BODY_COLOR)
                    add_mixed_text_runs(p, remaining)
                
                # add_pronounce_hyperlink disabled
                i += 1
                continue

        # Explicit hyperlink syntax: sentence = URL or previous line + = URL.
        standalone = re.match(r'^=\s*(https?://[^\s]+)\s*$', stripped, re.IGNORECASE)
        if standalone:
            url = normalize_url(standalone.group(1))
            # Attach to immediately previous paragraph while preserving any existing
            # pronounce hyperlink already appended to that paragraph.
            if doc.paragraphs:
                previous = doc.paragraphs[-1]
                direct_runs = list(previous.runs)
                visible = ''.join(r.text or '' for r in direct_runs)
                if visible.strip():
                    pronounce_hyperlink = None
                    for child in previous._p:
                        if child.tag == qn('w:hyperlink'):
                            texts = [t.text or '' for t in child.iter(qn('w:t'))]
                            if ''.join(texts).strip().lower() == PRONOUNCE_TEXT.lower():
                                pronounce_hyperlink = child
                                break

                    for r in direct_runs:
                        r._element.getparent().remove(r._element)

                    # IMPORTANT: Only English text gets the hyperlink
                    if is_english_only_text(visible):
                        add_external_hyperlink_before(
                            previous,
                            visible,
                            url,
                            before_element=pronounce_hyperlink,
                            color=english_color(visible),
                            bold=True
                        )
                    else:
                        # Split visible text into English and Arabic parts
                        chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', visible)
                        
                        for chunk in chunks:
                            if not chunk:
                                continue
                            
                            if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                                # English text gets hyperlink
                                add_external_hyperlink_before(
                                    previous,
                                    chunk,
                                    url,
                                    before_element=pronounce_hyperlink,
                                    color=english_color(chunk),
                                    bold=True
                                )
                            else:
                                # Arabic text remains normal
                                add_mixed_text_runs(previous, chunk)
                else:
                    add_rich_paragraph(doc, stripped)
            else:
                add_rich_paragraph(doc, stripped)
            i += 1
            continue

        # Case 2b: No space after =
        standalone_no_space = re.match(r'^=(https?://[^\s]+)\s*$', stripped, re.IGNORECASE)
        if standalone_no_space:
            url = normalize_url(standalone_no_space.group(1))
            # Attach to immediately previous paragraph while preserving any existing
            # pronounce hyperlink already appended to that paragraph.
            if doc.paragraphs:
                previous = doc.paragraphs[-1]
                direct_runs = list(previous.runs)
                visible = ''.join(r.text or '' for r in direct_runs)
                if visible.strip():
                    pronounce_hyperlink = None
                    for child in previous._p:
                        if child.tag == qn('w:hyperlink'):
                            texts = [t.text or '' for t in child.iter(qn('w:t'))]
                            if ''.join(texts).strip().lower() == PRONOUNCE_TEXT.lower():
                                pronounce_hyperlink = child
                                break

                    for r in direct_runs:
                        r._element.getparent().remove(r._element)

                    # IMPORTANT: Only English text gets the hyperlink
                    if is_english_only_text(visible):
                        add_external_hyperlink_before(
                            previous,
                            visible,
                            url,
                            before_element=pronounce_hyperlink,
                            color=english_color(visible),
                            bold=True
                        )
                    else:
                        # Split visible text into English and Arabic parts
                        chunks = re.split(r'([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)', visible)
                        
                        for chunk in chunks:
                            if not chunk:
                                continue
                            
                            if LATIN_RE.search(chunk) and not ARABIC_RE.search(chunk):
                                # English text gets hyperlink
                                add_external_hyperlink_before(
                                    previous,
                                    chunk,
                                    url,
                                    before_element=pronounce_hyperlink,
                                    color=english_color(chunk),
                                    bold=True
                                )
                            else:
                                # Arabic text remains normal
                                add_mixed_text_runs(previous, chunk)
                else:
                    add_rich_paragraph(doc, stripped)
            else:
                add_rich_paragraph(doc, stripped)
            i += 1
            continue

        inline_hl = re.match(r'^(.+?)\s*=\s*(https?://[^\s]+)\s*$', stripped, re.IGNORECASE)
        if inline_hl:
            label = inline_hl.group(1).strip()
            url = normalize_url(inline_hl.group(2))
            add_rich_paragraph(doc, label, hyperlink_url=url)
            i += 1
            continue

        # Case for inline without space: label=URL
        inline_hl_no_space = re.match(r'^(.+?)=(https?://[^\s]+)\s*$', stripped, re.IGNORECASE)
        if inline_hl_no_space:
            label = inline_hl_no_space.group(1).strip()
            url = normalize_url(inline_hl_no_space.group(2))
            add_rich_paragraph(doc, label, hyperlink_url=url)
            i += 1
            continue

        # Regular paragraph.
        add_rich_paragraph(doc, stripped)
        i += 1

        # Regular paragraph.
        add_rich_paragraph(doc, stripped)
        i += 1

    # Add Top bookmark at start of first content paragraph and Bottom at end.
    if doc.paragraphs:
        add_bookmark(doc.paragraphs[0], 'ContentTop', 9000)
        add_bookmark(doc.paragraphs[-1], 'ContentBottom', 9001)

    # Make each content section navigation-capable in Word's Navigation Pane.
    for sec in doc.sections[1:]:
        add_footer_page_number(sec)

    # Remove empty paragraphs created by the section break if possible.
    return doc, headings


def extract_title(markdown_content, fallback):
    for line in markdown_content.splitlines():
        m = re.match(r'^\s*#\s+(.+?)\s*$', line)
        if m:
            return m.group(1).strip()
    return fallback


def save_docx_safely(doc, output_docx):
    """
    Save DOCX safely.
    If the target already exists, replace it after removing the old file.
    """
    output_docx = Path(output_docx)
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    temp_docx = output_docx.with_name(
        f'.{output_docx.stem}_temp_{os.getpid()}.docx'
    )

    try:
        # Save to a temporary file first.
        doc.save(temp_docx)

        # Remove existing file if present.
        if output_docx.exists():
            try:
                output_docx.unlink()
            except PermissionError:
                raise PermissionError(
                    f'Cannot replace existing file because it is locked:\n'
                    f'{output_docx}\n\n'
                    f'Close the DOCX in Microsoft Word/OneDrive and run again.'
                )

        # Move temporary file into place.
        temp_docx.replace(output_docx)

        return True

    except PermissionError:
        if temp_docx.exists():
            try:
                temp_docx.unlink()
            except Exception:
                pass
        raise

    except Exception:
        if temp_docx.exists():
            try:
                temp_docx.unlink()
            except Exception:
                pass
        raise

def convert_docx_to_pdf_with_navigation(docx_path, pdf_path, headings):
    """
    Convert DOCX to PDF using Microsoft Word, then rebuild the PDF
    outline/bookmarks using PyMuPDF.

    This guarantees that PDF Navigation/Bookmarks exist even when
    Word does not transfer the DOCX navigation structure correctly.

    NOTE (fix): CreateBookmarks is now requested from Word itself
    (wdExportCreateHeadingBookmarks = 1) as well, so that even a manual
    "File > Save As > PDF" from inside Word on this DOCX produces a
    navigable PDF using the Heading styles. The PyMuPDF pass below still
    runs afterwards and OVERWRITES the outline with the exact
    numbered/bookmarked headings, since Word's own heading-based bookmarks
    don't know about our custom numbering (e.g. "0.01") or our bookmark
    names, and can occasionally miss/duplicate entries.
    """

    word = None
    document = None

    try:
        docx_path = Path(docx_path).resolve()
        pdf_path = Path(pdf_path).resolve()

        print('  Converting DOCX to PDF...')

        # --------------------------------------------------------
        # STEP 1: DOCX -> PDF using Microsoft Word
        # --------------------------------------------------------
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        document = word.Documents.Open(
            str(docx_path),
            ReadOnly=True
        )

        # Make sure PAGE / PAGEREF fields (TOC page numbers, footers) are
        # current before export, since ReadOnly opening does not always
        # trigger Word's own field refresh even though updateFields is set
        # in the document settings.
        try:
            document.Fields.Update()
        except Exception:
            pass

        # Word:
        # 17 = wdExportFormatPDF
        # 0  = Print quality
        # CreateBookmarks:
        #   0 = wdExportCreateNoBookmarks
        #   1 = wdExportCreateHeadingBookmarks   <-- now enabled
        #   2 = wdExportCreateWordBookmarks
        #
        # This makes the PDF produced by a manual "Save As PDF" (which
        # reuses the same export path when a user does it from the Word
        # UI) also come out with a Navigation pane, based on the document's
        # Heading 1-4 styles. Our own PyMuPDF pass below still rebuilds the
        # outline afterwards with the correct custom numbering/titles.
        document.ExportAsFixedFormat(
            OutputFileName=str(pdf_path),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            From=0,
            To=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

        document.Close(False)
        document = None

        word.Quit()
        word = None

        print(f'  PDF created: {pdf_path}')

        # --------------------------------------------------------
        # STEP 2: Build PDF Navigation / Outline with PyMuPDF
        # --------------------------------------------------------
        add_pdf_navigation_bookmarks(
            pdf_path,
            headings
        )

        print('  PDF Navigation/Bookmarks: ENABLED')

        return pdf_path

    except Exception as ex:
        print(f'  ERROR converting DOCX to PDF: {ex}')
        return None

    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass

        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

def _normalize_pdf_cell_text(value: object) -> str:
    """Normalize PDF table cell text for reliable Arabic/RTL matching."""
    import unicodedata

    s = '' if value is None else str(value)
    s = s.replace('\u00ad', '').replace('\u200b', '').replace('\ufeff', '')
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _normalize_digits(value: object) -> str:
    """Convert Arabic-Indic and Eastern Arabic-Indic digits to ASCII digits."""
    s = _normalize_pdf_cell_text(value)
    trans = str.maketrans(
        '٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹',
        '01234567890123456789'
    )
    return s.translate(trans)


def _text_variants(value: object) -> List[str]:
    """Return matching variants for PDF text, including a reversed RTL variant."""
    s = _normalize_pdf_cell_text(value)
    variants = []
    if s:
        variants.append(s)
        if ARABIC_RE.search(s):
            rev = s[::-1]
            if rev != s:
                variants.append(rev)
    return variants


def _compact_number_text(value: object) -> str:
    """Normalize a TOC numbering cell such as ٠.٠١ / 0.01."""
    s = _normalize_digits(value)
    return re.sub(r'\s+', '', s).replace('٫', '.').replace(',', '.')


def _extract_page_number_from_page_cell(value: object) -> Optional[int]:
    """Extract a page number from text already isolated to the PDF TOC page cell."""
    s = _normalize_digits(value)
    if not s:
        return None

    # The page cell is expected to contain exactly one logical page number.
    # Do NOT interpret heading/section numbers here.
    matches = re.findall(r'(?<!\d)\d{1,4}(?!\d)', s)
    if not matches:
        return None
    try:
        return int(matches[0])
    except ValueError:
        return None


def _cell_text_from_pdf_rect(page, bbox) -> str:
    """Extract normalized text from one geometric PDF rectangle."""
    if not bbox:
        return ''
    try:
        rect = fitz.Rect(bbox)
        words = page.get_text('words', clip=rect, sort=True) or []
        return _normalize_pdf_cell_text(' '.join(str(w[4]) for w in words if len(w) >= 5))
    except Exception:
        return ''


def _is_page_header_text(value: object) -> bool:
    """Return True when text looks like the PDF TOC page-column header."""
    s = _normalize_pdf_cell_text(value)
    variants = set(_text_variants(s))
    compact = re.sub(r'\s+', '', s)
    variants.add(compact)
    variants.add(compact[::-1])
    targets = {
        'الصفحة', 'صفحة', 'ةحفصلا', 'ةحفص', 'page', 'pages',
    }
    return any(v.lower() in targets for v in variants)


def _find_toc_page_column_bbox(page, table, page_idx: Optional[int] = None):
    """
    Find the *physical* PDF rectangle used by the TOC page column.

    Important: PyMuPDF can return a logical table row with many extracted
    columns while the real geometry has fewer columns (especially with RTL
    Word tables). Therefore the extraction index (for example 7) must NEVER
    be used directly as a geometric cell index.

    We identify the page header inside the actual header-cell rectangles and
    return that physical x-range.
    """
    if table is None:
        return None

    # Preferred: the real physical header cells exposed by PyMuPDF.
    try:
        header = getattr(table, 'header', None)
        header_cells = list(getattr(header, 'cells', None) or [])
    except Exception:
        header_cells = []

    if header_cells:
        for bbox in header_cells:
            txt = _cell_text_from_pdf_rect(page, bbox)
            if _is_page_header_text(txt):
                return bbox

        # Some PDFs expose header.names separately from header.cells.  Use the
        # name index only as a *hint*, then map it to the closest physical cell.
        try:
            names = list(getattr(header, 'names', None) or [])
        except Exception:
            names = []
        if page_idx is not None and names and 0 <= page_idx < len(names):
            # Map extraction-column x-center (when available) to the physical
            # header cell whose x-range is closest to the logical page-column
            # text found on the page.
            name = names[page_idx]
            if _is_page_header_text(name):
                # Try the page header's visible words first.
                try:
                    words = page.get_text('words', clip=fitz.Rect(table.bbox), sort=True) or []
                    candidates = []
                    for w in words:
                        if len(w) < 5:
                            continue
                        wt = _normalize_pdf_cell_text(w[4])
                        if _is_page_header_text(wt):
                            candidates.append(((float(w[0]) + float(w[2])) / 2.0, w))
                    if candidates:
                        center = candidates[0][0]
                        return min(
                            header_cells,
                            key=lambda b: abs(((float(b[0]) + float(b[2])) / 2.0) - center)
                        )
                except Exception:
                    pass

    # Secondary robust path: locate the visible header text in the table bbox,
    # then associate its x-center with the nearest physical header cell.
    try:
        table_rect = fitz.Rect(table.bbox)
        words = page.get_text('words', clip=table_rect, sort=True) or []
        candidates = []
        for w in words:
            if len(w) < 5:
                continue
            wt = _normalize_pdf_cell_text(w[4])
            if _is_page_header_text(wt):
                candidates.append(((float(w[0]) + float(w[2])) / 2.0, w))

        if candidates and header_cells:
            center = candidates[0][0]
            return min(
                header_cells,
                key=lambda b: abs(((float(b[0]) + float(b[2])) / 2.0) - center)
            )
    except Exception:
        pass

    return None


def _get_table_row_bbox(table, row_idx: int):
    """Return a row's physical bbox without confusing logical/extracted indexes."""
    if table is None:
        return None

    try:
        rows = list(getattr(table, 'rows', None) or [])
        if 0 <= row_idx < len(rows):
            bbox = getattr(rows[row_idx], 'bbox', None)
            if bbox:
                return bbox
            cells = list(getattr(rows[row_idx], 'cells', None) or [])
            valid = [c for c in cells if c]
            if valid:
                return (
                    min(float(c[0]) for c in valid),
                    min(float(c[1]) for c in valid),
                    max(float(c[2]) for c in valid),
                    max(float(c[3]) for c in valid),
                )
    except Exception:
        pass

    return None


def _extract_page_number_from_toc_geometry(page, table, row_idx: int, page_header_bbox) -> Optional[int]:
    """
    Read the page number only from the intersection of:
      1) the TOC page-column x-range, and
      2) the exact TOC row y-range.

    This is the key fix for RTL Word-generated PDFs where table.extract() can
    report 9 logical columns even though the real table has only 3 physical
    columns.
    """
    if not page_header_bbox:
        return None

    row_bbox = _get_table_row_bbox(table, row_idx)
    if not row_bbox:
        return None

    try:
        header_rect = fitz.Rect(page_header_bbox)
        row_rect = fitz.Rect(row_bbox)

        # Keep x strictly in the physical page column.  A very small tolerance
        # handles glyphs touching a border without reaching the title column.
        clip = fitz.Rect(
            max(0, header_rect.x0 - 1),
            max(0, row_rect.y0 - 1),
            header_rect.x1 + 1,
            row_rect.y1 + 1,
        )

        words = page.get_text('words', clip=clip, sort=True) or []
        tokens = [str(w[4]).strip() for w in words if len(w) >= 5 and str(w[4]).strip()]
        text = ' '.join(tokens)
        number = _extract_page_number_from_page_cell(text)
        if number is not None:
            return number

        # Token-by-token retry for RTL/direction-mark cases.
        for token in tokens:
            number = _extract_page_number_from_page_cell(token)
            if number is not None:
                return number

    except Exception:
        return None

    return None


def _extract_page_number_from_pdf_cell_bbox(page, cell_bbox) -> Optional[int]:
    """
    Read only the visible text inside one PDF TOC page-column cell.

    This deliberately does NOT search the heading or the physical PDF page.
    """
    if not cell_bbox:
        return None

    try:
        rect = fitz.Rect(cell_bbox)
        # Small tolerance for glyphs sitting on the cell border.
        clip = fitz.Rect(
            max(0, rect.x0 - 1),
            max(0, rect.y0 - 1),
            rect.x1 + 1,
            rect.y1 + 1,
        )
        words = page.get_text('words', clip=clip, sort=True) or []
        text = ' '.join(
            str(w[4]) for w in words
            if len(w) >= 5 and str(w[4]).strip()
        )

        page_number = _extract_page_number_from_page_cell(text)
        if page_number is not None:
            return page_number

        # Some RTL/Word PDFs return the digit as a separate token but the
        # normal text extraction can contain direction marks.  Inspect each
        # token independently as a final cell-local attempt.
        for word in words:
            if len(word) < 5:
                continue
            candidate = _extract_page_number_from_page_cell(str(word[4]))
            if candidate is not None:
                return candidate

    except Exception:
        return None

    return None



def _extract_page_number_from_table_row_by_column(page, table, row_idx: int, page_idx: int) -> Optional[int]:
    """
    Last cell-local fallback: derive the row y-band from any available table
    cell and the page-column x-band from table geometry, then read numeric words
    only from that intersection.
    """
    bbox = _get_table_cell_bbox(table, row_idx, page_idx)
    if bbox:
        return _extract_page_number_from_pdf_cell_bbox(page, bbox)
    return None


def _detect_toc_columns(header_row: List[object]) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    """Detect number/title/page columns from the actual PDF table headers."""
    header_cells = [_normalize_pdf_cell_text(cell) for cell in header_row]

    number_idx = None
    title_idx = None
    page_idx = None

    for idx, cell in enumerate(header_cells):
        variants = _text_variants(cell)
        joined = ' '.join(variants)
        compact = re.sub(r'\s+', '', joined)

        if page_idx is None and (
            'الصفحة' in joined or 'صفحة' in joined or
            'ةحفصلا' in joined or 'ةحفص' in compact or
            'page' in joined.lower()
        ):
            page_idx = idx

        if title_idx is None and (
            'عنوان القسم' in joined or 'عنوان' in joined or
            'مسقلا ناونع' in joined or 'title' in joined.lower()
        ):
            title_idx = idx

        if number_idx is None and (
            'الرقم' in joined or 'رقم' in joined or 'م قرلا' in joined or
            'number' in joined.lower() or 'no' == joined.lower()
        ):
            number_idx = idx

    return number_idx, title_idx, page_idx


def _find_heading_for_pdf_row(row: List[object], headings: List[Dict],
                               number_idx: Optional[int], title_idx: Optional[int]) -> Optional[Dict]:
    """Match a PDF TOC row to the correct Markdown heading."""
    row_number = _compact_number_text(row[number_idx]) if number_idx is not None and number_idx < len(row) else ''
    row_title_variants = set(_text_variants(row[title_idx])) if title_idx is not None and title_idx < len(row) else set()

    # Strongest match: TOC number.
    if row_number:
        for heading in headings:
            if _compact_number_text(heading.get('number', '')) == row_number:
                return heading

    # Second choice: exact title match, including RTL-reversed extraction.
    if row_title_variants:
        for heading in headings:
            heading_variants = set(_text_variants(heading.get('title', '')))
            if row_title_variants & heading_variants:
                return heading

    # Last resort: normalized containment, useful when PDF extraction inserts
    # or removes line-break spacing inside Arabic titles.
    row_title = _normalize_pdf_cell_text(row[title_idx]) if title_idx is not None and title_idx < len(row) else ''
    if row_title:
        row_compact = re.sub(r'\s+', '', row_title)
        for heading in headings:
            for variant in _text_variants(heading.get('title', '')):
                variant_compact = re.sub(r'\s+', '', variant)
                if row_compact and variant_compact and (
                    row_compact == variant_compact or
                    row_compact in variant_compact or
                    variant_compact in row_compact
                ):
                    return heading

    return None


def extract_page_numbers_from_pdf(pdf_path, headings):
    """
    Extract the page numbers shown in the PDF TOC's visible **page column**.

    The PDF is treated as the source of truth.  No heading-search fallback is
    used here, because a heading search returns the PDF's physical page index
    and can accidentally find the heading inside the TOC itself.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        print(f'  WARNING: PDF file not found: {pdf_path}')
        return {}

    # Removed DEBUG: Reading page numbers ONLY from the PDF TOC page column
    page_numbers: Dict[str, int] = {}

    try:
        pdf = fitz.open(str(pdf_path))
        # Removed DEBUG: PDF has {len(pdf)} pages

        toc_tables = []
        scan_pages = min(12, len(pdf))

        for page_num in range(scan_pages):
            page = pdf[page_num]
            try:
                tables = page.find_tables()
            except Exception as table_error:
                print(f'  WARNING: Could not inspect tables on PDF page {page_num + 1}: {table_error}')
                continue

            for table in tables:
                try:
                    table_data = table.extract()
                except Exception as extract_error:
                    print(f'  WARNING: Could not extract a PDF table on page {page_num + 1}: {extract_error}')
                    continue

                if not table_data:
                    continue

                header_row = table_data[0]
                number_idx, title_idx, page_idx = _detect_toc_columns(header_row)

                print(
                    # Removed DEBUG: PDF page table headers
                    f'{[_normalize_pdf_cell_text(c) for c in header_row]}'
                )
                print(
                    # Removed DEBUG: Detected TOC columns
                    f'title={title_idx}, page={page_idx}'
                )

                page_header_bbox = _find_toc_page_column_bbox(page, table, page_idx)

                if page_header_bbox and (title_idx is not None or number_idx is not None):
                    print(
                        # Removed DEBUG: Physical PDF page-column bbox
                    )
                    toc_tables.append(
                        (page_num, table, table_data, number_idx, title_idx, page_idx, page_header_bbox)
                    )

        if not toc_tables:
            print('  WARNING: Could not locate the PDF TOC page-column geometry; no page numbers will be invented.')
            pdf.close()
            return {}

        matched_bookmarks = set()
        used_heading_indices = set()

        for page_num, table, table_data, number_idx, title_idx, page_idx, page_header_bbox in toc_tables:
            page = pdf[page_num]

            # We deliberately ignore row[page_idx] as a source of page numbers.
            # In RTL Word PDFs that logical index may not correspond to the
            # physical page column. Geometry is the source of truth.
            for row_idx, row in enumerate(table_data[1:], start=1):
                page_number = _extract_page_number_from_toc_geometry(
                    page, table, row_idx, page_header_bbox
                )

                if page_number is None:
                    print(
                        f'  WARNING: Could not read logical page number from the PDF TOC page column '
                        f'on PDF page {page_num + 1}, row {row_idx}; skipping this row.'
                    )
                    continue

                heading = _find_heading_for_pdf_row(row, headings, number_idx, title_idx)

                # Safe row-order fallback: only among headings not already
                # matched, and only after a real page-column number was read.
                if heading is None:
                    for idx, candidate in enumerate(headings):
                        if idx not in used_heading_indices and candidate['bookmark'] not in matched_bookmarks:
                            heading = candidate
                            print(
                                # Removed DEBUG: Row-order match used
                                f'-> {candidate["number"]} {candidate["title"][:50]}...'
                            )
                            break

                if heading is None:
                    continue

                bookmark = heading['bookmark']
                if bookmark in matched_bookmarks:
                    continue

                page_numbers[bookmark] = page_number
                matched_bookmarks.add(bookmark)
                try:
                    used_heading_indices.add(headings.index(heading))
                except ValueError:
                    pass

                print(
                    # Removed DEBUG: PDF TOC bookmark and page
                    f'(read from physical page-column geometry {page_header_bbox})'
                )

        pdf.close()

        missing = [h for h in headings if h['bookmark'] not in page_numbers]
        if missing:
            print(
                f'  WARNING: {len(missing)} headings were not matched to a readable PDF TOC page-column value. '
                'They are left unchanged; no physical-page fallback is used.'
            )

        # Removed DEBUG: Final page-number map
        return page_numbers

    except Exception as ex:
        print(f'  ERROR extracting page numbers from PDF TOC: {ex}')
        import traceback
        traceback.print_exc()
        return {}


def extract_page_numbers_from_pdf_heading_search(pdf_path, headings):
    """
    Fallback method: extract page numbers by searching for headings in the PDF.
    This is used only when a PDF TOC page-column value is unavailable.
    """
    if not pdf_path.exists():
        print(f'  WARNING: PDF file not found: {pdf_path}')
        return {}

    # Removed DEBUG: Fallback heading search
    page_numbers = {}

    try:
        pdf = fitz.open(str(pdf_path))
        # Removed DEBUG: PDF has {len(pdf)} pages
        last_search_page = 0

        for h in headings:
            bookmark = h['bookmark']
            title = f"{h['number']} {h['title']}"
            # Removed DEBUG: Searching for heading

            matches = []
            for page_index in range(last_search_page, len(pdf)):
                page = pdf[page_index]
                found = page.search_for(title)
                if found:
                    matches = found
                    last_search_page = page_index
                    break

            if not matches:
                for page_index in range(last_search_page, len(pdf)):
                    page = pdf[page_index]
                    found = page.search_for(h['title'])
                    if found:
                        matches = found
                        last_search_page = page_index
                        break

            if matches:
                page_number = last_search_page + 1
                # Removed DEBUG: Fallback found at PDF physical page
                page_numbers[bookmark] = page_number
            else:
                print(f'  WARNING: Heading not located in PDF: {title}')

        pdf.close()
        # Removed DEBUG: Fallback extracted page numbers
        return page_numbers

    except Exception as ex:
        print(f'  ERROR in heading-search fallback: {ex}')
        import traceback
        traceback.print_exc()
        return {}


def update_word_toc_with_page_numbers(docx_path, page_numbers, headings):
    """
    Update the Word TOC page column with page numbers read from the PDF TOC.

    The Word TOC itself has a fixed three-column layout generated by this
    script, so the Word page column is the third cell.  Only the value coming
    from page_numbers is written; Word is not asked to recalculate the TOC.
    """
    docx_path = Path(docx_path)
    if not docx_path.exists():
        print(f'  WARNING: DOCX file not found: {docx_path}')
        return False

        # Removed DEBUG: PDF page-number map and headings count

    try:
        doc = Document(str(docx_path))
        # Removed DEBUG: Found tables in Word document

        for table_idx, table in enumerate(doc.tables):
            if not table.rows:
                continue

            cells_text = [cell.text.strip() for cell in table.rows[0].cells]
            # Removed DEBUG: Word table headers

            if not ('الرقم' in cells_text and 'عنوان القسم' in cells_text and 'الصفحة' in cells_text):
                continue

            # Removed DEBUG: Found Word TOC table
            updated_count = 0

            # Word TOC rows are generated from headings in the same order.
            for row_idx, row in enumerate(table.rows[1:]):
                heading_index = row_idx
                if len(row.cells) < 3 or heading_index >= len(headings):
                    continue

                heading = headings[heading_index]
                bookmark = heading['bookmark']
                if bookmark not in page_numbers:
                    print(f'  WARNING: No PDF page number for bookmark {bookmark}; leaving Word cell unchanged.')
                    continue

                new_page = str(page_numbers[bookmark])
                page_cell = row.cells[2]
                old_page = page_cell.text.strip()

                # Replace the visible content while keeping the table cell valid.
                paragraph = page_cell.paragraphs[0]
                # Remove all existing paragraph XML children except paragraph properties.
                for child in list(paragraph._p):
                    if child.tag != qn('w:pPr'):
                        paragraph._p.remove(child)

                new_run = paragraph.add_run(new_page)
                set_run_font(new_run, size=9, bold=True, color='21618C')
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Remove extra paragraphs inside the page cell.
                while len(page_cell.paragraphs) > 1:
                    page_cell._element.remove(page_cell.paragraphs[-1]._element)

                print(
                    # Removed DEBUG: Word TOC row update
                    f'{old_page!r} -> {new_page!r} ({bookmark})'
                )
                updated_count += 1

            # Save once after all cells are updated.
            doc.save(str(docx_path))
            print(f'  Updated Word TOC with {updated_count} PDF-derived page numbers.')

            # Optional Word COM pass: unlink any residual PAGEREF fields and
            # write the same PDF-derived number while preserving the cell marker.
            try:
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                doc_word = word.Documents.Open(str(docx_path.absolute()))
                try:
                    for table in doc_word.Tables:
                        if table.Rows.Count == 0:
                            continue

                        first_row_cells = [cell.Range.Text.strip('\r\x07 ').strip() for cell in table.Rows(1).Cells]
                        if not ('الرقم' in first_row_cells and 'عنوان القسم' in first_row_cells and 'الصفحة' in first_row_cells):
                            continue

                        for row_idx in range(2, table.Rows.Count + 1):
                            heading_index = row_idx - 2
                            if heading_index >= len(headings):
                                continue

                            heading = headings[heading_index]
                            bookmark = heading['bookmark']
                            if bookmark not in page_numbers:
                                continue

                            page_cell = table.Cell(row_idx, 3)
                            page_cell.Range.Fields.Unlink()

                            # Exclude Word's end-of-cell marker from the replacement range.
                            cell_range = page_cell.Range.Duplicate
                            if cell_range.End > cell_range.Start:
                                cell_range.End -= 1
                            cell_range.Text = str(page_numbers[bookmark])

                    doc_word.Save()
                finally:
                    doc_word.Close(SaveChanges=True)
                    word.Quit()

                print('  Word COM verification/update completed.')
            except Exception as e:
                print(f'  Note: Word COM pass skipped: {e}')

            return updated_count > 0

        print('  WARNING: Word TOC table not found.')
        return False

    except Exception as ex:
        print(f'  ERROR updating Word TOC: {ex}')
        import traceback
        traceback.print_exc()
        return False


def add_pdf_navigation_bookmarks(pdf_path, headings):
    """
    Add a real PDF Outline / Bookmark tree.

    Each Markdown heading becomes a PDF bookmark.
    Heading hierarchy:
        H2 -> PDF level 1
        H3 -> PDF level 2
        H4 -> PDF level 3
        H5 -> PDF level 4
    """

    pdf_path = Path(pdf_path)

    pdf = fitz.open(str(pdf_path))

    try:
        if not headings:
            print('  No headings found. PDF navigation was not created.')
            return

        toc = []

        # Keep track of the last page used for each heading.
        last_search_page = 0

        for h in headings:

            # PDF hierarchy levels must start at 1, not 0
            # Map Word heading levels to PDF outline levels:
            # H1 (level 1) -> PDF level 1
            # H2 (level 2) -> PDF level 1  
            # H3 (level 3) -> PDF level 2
            # H4 (level 4) -> PDF level 3
            # H5 (level 5) -> PDF level 4
            if h['level'] == 1:
                level = 1  # H1 -> PDF level 1
            elif h['level'] == 2:
                level = 1  # H2 -> PDF level 1 (main sections)
            elif h['level'] == 3:
                level = 2  # H3 -> PDF level 2
            elif h['level'] == 4:
                level = 3  # H4 -> PDF level 3
            else:
                level = 4  # H5+ -> PDF level 4
            
            title = f"{h['number']} {h['title']}"

            # ----------------------------------------------------
            # Search exact numbered heading first.
            # ----------------------------------------------------
            matches = []

            for page_index in range(last_search_page, len(pdf)):
                page = pdf[page_index]

                found = page.search_for(title)

                if found:
                    matches = found
                    last_search_page = page_index
                    break

            # ----------------------------------------------------
            # If exact search fails, search title only.
            # ----------------------------------------------------
            if not matches:

                for page_index in range(last_search_page, len(pdf)):
                    page = pdf[page_index]

                    found = page.search_for(h['title'])

                    if found:
                        matches = found
                        last_search_page = page_index
                        break

            # ----------------------------------------------------
            # Final fallback:
            # keep the bookmark alive even if text search fails.
            # ----------------------------------------------------
            if matches:
                page_number = last_search_page + 1
            else:
                page_number = 1

                print(
                    f'  WARNING: Heading not located in PDF: {title}'
                )

            toc.append([
                level,
                title,
                page_number,
            ])
            
            # Removed DEBUG: TOC entry

        # --------------------------------------------------------
        # Validate TOC before setting it
        # --------------------------------------------------------
        # Removed DEBUG: Total TOC entries
        
        # PDF requires first item to be level 1
        if toc and toc[0][0] != 1:
            # Removed DEBUG: Fixing first TOC entry level
            toc[0][0] = 1
        
        for idx, entry in enumerate(toc):
            if entry[0] < 1:
                print(f'  ERROR: TOC entry {idx} has invalid level {entry[0]}: {entry[1]}')
                entry[0] = 1  # Fix invalid level

        # --------------------------------------------------------
        # Set the complete PDF Outline.
        # --------------------------------------------------------
        pdf.set_toc(toc)

        # --------------------------------------------------------
        # Save PDF safely.
        # --------------------------------------------------------
        temp_pdf = pdf_path.with_name(
            f'.{pdf_path.stem}_navigation_temp.pdf'
        )

        pdf.save(
            str(temp_pdf),
            garbage=4,
            deflate=True
        )

        pdf.close()

        if pdf_path.exists():
            pdf_path.unlink()

        temp_pdf.replace(pdf_path)

    except Exception:
        try:
            pdf.close()
        except Exception:
            pass
        raise
                
def process_file(input_file):
    input_file = Path(input_file)

    if input_file.suffix.lower() != '.md':
        raise ValueError(f'Not a Markdown file: {input_file}')

    output_dir = input_file.parent / 'WORD'
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f'\nConverting to Word: {input_file}')

    md_content = read_markdown_file(input_file)

    doc, headings = build_document(
        md_content,
        input_file.name
    )

    # --------------------------------------------------------
    # Save DOCX
    # --------------------------------------------------------
    output_docx = output_dir / f'{input_file.stem}.docx'

    try:
        save_docx_safely(
            doc,
            output_docx
        )
    except PermissionError as ex:
        print(f'  ERROR: {ex}')
        return None

    print(f'  Sections: {len(headings)}')
    print(f'  DOCX saved to: {output_docx}')

    # --------------------------------------------------------
    # Convert DOCX -> PDF
    # --------------------------------------------------------
    output_pdf = output_dir / f'{input_file.stem}.pdf'

    pdf_result = convert_docx_to_pdf_with_navigation(
        output_docx,
        output_pdf,
        headings
    )

    if pdf_result:
        print(f'  PDF saved to: {output_pdf}')
        print('  PDF Navigation/Bookmarks: ENABLED')
        
        # --------------------------------------------------------
        # TOC table removed - no page number extraction needed
        # Navigation/bookmarks still work in PDF
        # --------------------------------------------------------
        print('  TOC table removed - navigation/bookmarks still work in PDF')

    return output_docx


def main(input_path):
    input_path = Path(input_path)
    if not input_path.exists():
        print(f'ERROR: Path does not exist: {input_path}')
        sys.exit(1)

    if input_path.is_file():
        if input_path.suffix.lower() != '.md':
            print('ERROR: Input file must be a .md file.')
            sys.exit(1)
        process_file(input_path)
        return

    if input_path.is_dir():
        md_files = sorted(p for p in input_path.rglob('*.md') if p.is_file())
        if not md_files:
            print(f'No .md files found under: {input_path}')
            return
        print(f'Found {len(md_files)} Markdown files.')
        for md_file in md_files:
            try:
                process_file(md_file)
            except Exception as ex:
                print(f'  ERROR processing {md_file}: {ex}')
        print('\nDone.')
        return

    print('ERROR: Input path must be a file or folder.')
    sys.exit(1)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(
            'Usage:\n'
            '  py generatePdfTablesAndIndexingFromMd.py <markdown-file-or-folder>\n\n'
            'Examples:\n'
            r'  py generatePdfTablesAndIndexingFromMd.py "E:\\mazen\\English\\Authentication.md"' '\n'
            r'  py generatePdfTablesAndIndexingFromMd.py "E:\\mazen\\English"'
        )
        sys.exit(1)
    main(sys.argv[1])