"""
report_generator.py
-------------------
Core docx generation logic — imported by both the CLI script and the web app.
Call generate_report_bytes(content: str) -> bytes to get the .docx as bytes.
"""
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement, ns


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

def _create_element(name):
    return OxmlElement(name)

def _create_attribute(element, name, value):
    element.set(ns.qn(name), value)

def _add_page_number(run):
    fldChar1 = _create_element('w:fldChar')
    _create_attribute(fldChar1, 'w:fldCharType', 'begin')
    instrText = _create_element('w:instrText')
    _create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"
    fldChar2 = _create_element('w:fldChar')
    _create_attribute(fldChar2, 'w:fldCharType', 'separate')
    fldChar3 = _create_element('w:fldChar')
    _create_attribute(fldChar3, 'w:fldCharType', 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def _add_field(paragraph, field_code):
    """Insert a Word field (TOC, LOF, etc.) into a paragraph."""
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(ns.qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(ns.qn('xml:space'), 'preserve')
    instrText.text = field_code
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(ns.qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(ns.qn('w:fldCharType'), 'end')
    r = run._r
    r.append(fldChar)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)


# ---------------------------------------------------------------------------
# Main generator
# ---------------------------------------------------------------------------

def generate_report_bytes(content: str) -> bytes:
    """
    Parse `content` (formatted text) and return the generated .docx as bytes.
    """
    doc = Document()

    # Margins — 1 inch all sides
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)

    # --- Setup Front Matter Numbering (i, ii, iii) ---
    first_section = doc.sections[0]
    first_section.footer_distance = Inches(0.5)
    footer = first_section.footer
    # Ensure footer has a paragraph
    if not footer.paragraphs:
        footer.add_paragraph()
    fp = footer.paragraphs[0]
    fp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    fr = fp.add_run()
    fr.font.name = 'Times New Roman'
    fr.font.size = Pt(12)
    _add_page_number(fr)

    # Set numbering format to upperRoman for first section
    sectPr = first_section._sectPr
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(ns.qn('w:fmt'), 'upperRoman')
    sectPr.append(pgNumType)

    styles = doc.styles

    # Normal style
    ns_style = styles['Normal']
    ns_style.font.name = 'Times New Roman'
    ns_style.font.size = Pt(12)
    ns_style.paragraph_format.line_spacing      = 1.5
    ns_style.paragraph_format.line_spacing_rule = 1
    ns_style.paragraph_format.space_after       = Pt(12)
    ns_style.paragraph_format.alignment         = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    # Heading 1
    h1 = styles['Heading 1']
    h1.font.name = 'Times New Roman'
    h1.font.size = Pt(18)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 0, 0)
    h1.paragraph_format.alignment   = WD_PARAGRAPH_ALIGNMENT.LEFT
    h1.paragraph_format.space_before = Pt(24)
    h1.paragraph_format.space_after  = Pt(24)

    # Heading 2
    h2 = styles['Heading 2']
    h2.font.name = 'Times New Roman'
    h2.font.size = Pt(16)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0, 0, 0)
    h2.paragraph_format.space_before = Pt(18)
    h2.paragraph_format.space_after  = Pt(12)

    # Heading 3
    h3 = styles['Heading 3']
    h3.font.name = 'Times New Roman'
    h3.font.size = Pt(14)
    h3.font.bold = True
    h3.font.color.rgb = RGBColor(0, 0, 0)
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after  = Pt(10)

    # --- Ensure 'Caption' style exists ---
    try:
        cap_style = styles['Caption']
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        cap_style = styles.add_style('Caption', WD_STYLE_TYPE.PARAGRAPH)
        cap_style.base_style = styles['Normal']
        cap_style.font.italic = True
        cap_style.font.size = Pt(11)
        cap_style.font.name = 'Times New Roman'
        cap_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        cap_style.paragraph_format.space_after = Pt(12)

    # -----------------------------------------------------------------------
    # Front matter placeholder
    # -----------------------------------------------------------------------
    note = doc.add_paragraph()
    note.add_run("--- Manual Pages Placeholder ---").bold = True
    doc.add_paragraph("1. Title Page\n2. Certificate\n3. Acknowledgement\n4. Abstract")
    doc.add_page_break()

    # Table of Contents
    doc.add_paragraph("Table of Contents", style='Heading 1')
    _add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
    doc.add_paragraph(
        "[Update in Word: Ctrl+A → F9 → Update entire table]"
    ).italic = True
    doc.add_page_break()

    # List of Figures
    doc.add_paragraph("List of Figures", style='Heading 1')
    # \t "Caption" tells word to pick up every paragraph with 'Caption' style
    _add_field(doc.add_paragraph(), 'TOC \\h \\z \\t "Caption"')
    doc.add_paragraph(
        "[Update in Word: Ctrl+A → F9 → Update entire table]"
    ).italic = True
    doc.add_page_break()

    # -----------------------------------------------------------------------
    # Parse content line by line
    # -----------------------------------------------------------------------
    for line in content.split('\n'):
        line = line.strip()
        if not line:
            continue

        # ---- Skip pure structural markers ----
        if (line.startswith('===') or
                line.startswith('NOTE TO GENERATOR:') or
                line.startswith('FakeShield project')):
            continue

        # ---- CHAPTER ----
        if line.startswith('CHAPTER:'):
            text = line.replace('CHAPTER:', '').strip()

            if text.lower().startswith('chapter 1 '):
                new_section = doc.add_section(2)  # 2 = WD_SECTION_START.NEW_PAGE
                new_section.footer.is_linked_to_previous = False
                fp = new_section.footer.paragraphs[0]
                fp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                fr = fp.add_run()
                fr.font.name = 'Times New Roman'
                fr.font.size = Pt(12)
                _add_page_number(fr)
                
                # Restart page numbering at 1 AND switch to decimal (1, 2, 3)
                sectPr = new_section._sectPr
                pgNumType = OxmlElement('w:pgNumType')
                pgNumType.set(ns.qn('w:start'), "1")
                pgNumType.set(ns.qn('w:fmt'), 'decimal')
                sectPr.append(pgNumType)
            else:
                doc.add_page_break()

            if text.lower().startswith('chapter '):
                parts = text.split(' ', 2)
                if len(parts) == 3 and parts[1].isdigit():
                    text = f"Chapter {parts[1]}\n{parts[2]}"

            doc.add_heading(text, level=1)

        # ---- HEADING ----
        elif line.startswith('HEADING:'):
            doc.add_heading(line.replace('HEADING:', '').strip(), level=2)

        # ---- SUBHEADING ----
        elif line.startswith('SUBHEADING:'):
            doc.add_heading(line.replace('SUBHEADING:', '').strip(), level=3)

        # ---- PARA ----
        elif line.startswith('PARA:'):
            doc.add_paragraph(line.replace('PARA:', '').strip(), style='Normal')

        # ---- FIGURE ----
        elif line.startswith('FIGURE:'):
            text = line.replace('FIGURE:', '').strip()

            # Spacer paragraph — keep_with_next glues it to the caption
            spacer = doc.add_paragraph()
            spacer.paragraph_format.space_before    = Pt(0)
            spacer.paragraph_format.space_after     = Inches(3.5)
            spacer.paragraph_format.keep_with_next  = True   # ← KEY: never split from caption

            # Caption paragraph (Caption style → appears in List of Figures)
            try:
                fig_para = doc.add_paragraph(style='Caption')
            except KeyError:
                fig_para = doc.add_paragraph()

            fig_para.paragraph_format.space_before = Pt(0)
            fig_para.paragraph_format.space_after  = Pt(12)
            fig_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

            fig_run = fig_para.add_run(text)
            fig_run.italic = True
            fig_run.font.name = 'Times New Roman'
            fig_run.font.size = Pt(11)

        # ---- Unrecognised line (pass-through) ----
        else:
            doc.add_paragraph(line)

    # -----------------------------------------------------------------------
    # Save to bytes buffer and return
    # -----------------------------------------------------------------------
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
