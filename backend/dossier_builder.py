import sys
import os
import json
import re
import shutil
import stat
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docxcompose.composer import Composer
from pdf2docx import Converter
import win32com.client


# ---------------------------------------------------------------------------
# PDF → DOCX conversion
# ---------------------------------------------------------------------------
def convert_pdf_to_docx(pdf_path, docx_path):
    print(f"Converting PDF {pdf_path} to DOCX...")
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()
    return docx_path


# ---------------------------------------------------------------------------
# Template text replacement
# ---------------------------------------------------------------------------
def replace_text_in_document(doc, old_text, new_text):
    for p in doc.paragraphs:
        if old_text in p.text:
            for run in p.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
            if old_text in p.text:
                p.text = p.text.replace(old_text, new_text)


# ---------------------------------------------------------------------------
# Section heading formatter
# ---------------------------------------------------------------------------
def format_heading(paragraph, heading_type, heading_text):
    # Apply Heading 1 style so Word's TOC field (TOC \o "1-3") can index
    # this paragraph. Visual overrides below preserve the exact appearance.
    try:
        paragraph.style = paragraph.part.document.styles['Heading 1']
    except Exception:
        pass
    paragraph.text = ""
    cyan_run = paragraph.add_run(f"SECTION {heading_type} \n")
    cyan_run.font.name = 'Arial'
    cyan_run.font.size = Pt(10)
    cyan_run.font.bold = True
    cyan_run.font.color.rgb = RGBColor(0, 126, 197)

    plum_run = paragraph.add_run(heading_text)
    plum_run.font.name = 'Arial'
    plum_run.font.size = Pt(16)
    plum_run.font.bold = True
    plum_run.font.color.rgb = RGBColor(108, 16, 91)

    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(12)


# ---------------------------------------------------------------------------
# Strip header/footer noise paragraphs from converted DOCX
# ---------------------------------------------------------------------------
HEADER_FOOTER_PATTERNS = [
    re.compile(r"^Page\s+\d+", re.IGNORECASE),
    re.compile(r"^\d+\s+of\s+\d+", re.IGNORECASE),
    re.compile(r"^\d+\s*/\s*\d+", re.IGNORECASE),
    re.compile(r"JMCP\.org", re.IGNORECASE),
    re.compile(r"Volume\s+\d+", re.IGNORECASE),
    re.compile(r"Number\s+\d+\S*", re.IGNORECASE),
    re.compile(r"\.{4,}", re.IGNORECASE),
]


# ---------------------------------------------------------------------------
# Collapse runs of blank paragraphs — keeps at most max_consecutive in a row.
# Skips paragraphs that carry an explicit page break or inline drawings so
# intentional layout is preserved.
# ---------------------------------------------------------------------------
_WNS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def _collapse_blank_paragraphs(body, max_consecutive=1):
    blank_run = []
    to_remove = []
    for el in list(body):
        if el.tag.endswith('}p') or el.tag.endswith('p'):
            text = ''.join(el.itertext()).strip()
            has_drawing = bool(el.findall(f'.//{_WNS}drawing'))
            brs = el.findall(f'.//{_WNS}br')
            has_page_break = any(
                br.get(f'{_WNS}type') == 'page' for br in brs
            )
            is_blank = not text and not has_drawing and not has_page_break
            if is_blank:
                blank_run.append(el)
            else:
                if len(blank_run) > max_consecutive:
                    to_remove.extend(blank_run[max_consecutive:])
                blank_run = []
    # Trailing blank run
    if len(blank_run) > max_consecutive:
        to_remove.extend(blank_run[max_consecutive:])
    for el in to_remove:
        try:
            body.remove(el)
        except Exception:
            pass


def prepare_section_document(source_doc_path, temp_path):
    doc = Document(source_doc_path)
    body = doc.element.body
    # Pass 1 — remove header/footer noise lines
    for el in list(body):
        if el.tag.endswith('p') or el.tag.endswith('}p'):
            text = ''.join(el.itertext()).strip()
            for cp in HEADER_FOOTER_PATTERNS:
                if cp.search(text):
                    body.remove(el)
                    break
    # Pass 2 — collapse excessive blank paragraphs left by PDF conversion
    _collapse_blank_paragraphs(body, max_consecutive=1)
    doc.save(temp_path)


# ---------------------------------------------------------------------------
# Inject a DOCX at a placeholder inside the master template
# ---------------------------------------------------------------------------
def insert_document_at_placeholder(master_doc, placeholder_text, filepath_to_insert):
    body = master_doc.element.body
    target_idx = -1
    for i, el in enumerate(body):
        if el.tag.endswith('p') and placeholder_text in ''.join(el.itertext()):
            target_idx = i
            break
    if target_idx == -1:
        print(f"Warning: Placeholder '{placeholder_text}' not found — skipping.")
        return

    tail_elements = list(body[target_idx + 1:-1])
    final_sectPr = body[-1] if body[-1].tag.endswith('sectPr') else None

    for el in tail_elements:
        body.remove(el)
    if final_sectPr is not None:
        body.remove(final_sectPr)

    try:
        composer = Composer(master_doc)
        composer.append(Document(filepath_to_insert))
    except Exception as e:
        print(f"Warning: Could not append {filepath_to_insert}: {e}")

    body.remove(body[target_idx])

    for el in tail_elements:
        body.append(el)
    if final_sectPr is not None:
        body.append(final_sectPr)


# ---------------------------------------------------------------------------
# Remove unused placeholder paragraphs from the template
# ---------------------------------------------------------------------------
def remove_unused_section(master_doc, placeholder_text):
    body = master_doc.element.body
    for i, el in enumerate(body):
        if el.tag.endswith('p') and placeholder_text in ''.join(el.itertext()):
            to_remove = [el]
            if i > 0 and body[i - 1].tag.endswith('p'):
                to_remove.append(body[i - 1])
            if i + 1 < len(body) and body[i + 1].tag.endswith('p'):
                to_remove.append(body[i + 1])
            for rem in to_remove:
                try:
                    body.remove(rem)
                except Exception:
                    pass
            break


# ---------------------------------------------------------------------------
# Remove consecutive duplicate page-break paragraphs (blank page prevention)
# ---------------------------------------------------------------------------
def remove_duplicate_page_breaks(docx_path):
    doc = Document(docx_path)
    body = doc.element.body
    # Pass 1 — remove consecutive duplicate explicit page-break paragraphs
    to_remove = []
    last_was_break = False
    for el in list(body):
        if el.tag.endswith('p') or el.tag.endswith('}p'):
            brs = el.findall(f'.//{_WNS}br')
            is_page_break = any(br.get(f'{_WNS}type') == 'page' for br in brs)
            text = ''.join(el.itertext()).strip()
            has_content = bool(text) or bool(el.findall(f'.//{_WNS}drawing'))
            if is_page_break and not has_content:
                if last_was_break:
                    to_remove.append(el)
                else:
                    last_was_break = True
            else:
                if not is_page_break:
                    last_was_break = False
    for el in to_remove:
        try:
            body.remove(el)
        except Exception:
            pass
    # Pass 2 — collapse runs of blank paragraphs (implicit blank pages)
    _collapse_blank_paragraphs(body, max_consecutive=2)
    doc.save(docx_path)


# ---------------------------------------------------------------------------
# Optional: update TOC/fields via Word COM — best effort, never blocks
# ---------------------------------------------------------------------------
def try_update_toc(docx_path):
    abs_path = str(Path(docx_path).resolve())
    doc_name = Path(docx_path).name
    lock_path = str(Path(docx_path).parent / ("~$" + doc_name))
    # Remove stale lock file
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass

    temp_path = abs_path.replace('.docx', '_toc_tmp.docx')
    try:
        shutil.copy2(abs_path, temp_path)
        os.chmod(temp_path, stat.S_IRUSR | stat.S_IWUSR | stat.S_IRGRP | stat.S_IWGRP | stat.S_IROTH | stat.S_IWOTH)
    except Exception:
        return  # Can't even copy — skip silently

    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(temp_path, ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False)
        if not doc.ReadOnly:
            doc.Fields.Update()
            for toc in doc.TablesOfContents:
                toc.Update()
            # FileFormat=16 = wdFormatXMLDocument (.docx) — prevents
            # Word from showing an interactive format dialog on Windows.
            doc.SaveAs(abs_path, FileFormat=16)
            print("TOC updated successfully.")
        else:
            print("WARNING: Could not update TOC (read-only). Skipping.")
        doc.Close(SaveChanges=False)
    except Exception as e:
        print(f"WARNING: TOC update skipped ({e})")
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Main compilation entry point
# ---------------------------------------------------------------------------
def compile_dossier(config_path, output_path):
    with open(config_path, 'r') as f:
        config = json.load(f)

    metadata = config.get('metadata', {})
    sections = config.get('sections', {})
    dossier_type = config.get('dossierType', 'format-a')

    # Load master AMCP template
    template_path = os.path.join(os.path.dirname(__file__), 'amcp_template.docx')
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Master template not found: {template_path}")

    master_doc = Document(template_path)

    # Inject cover page metadata
    replace_text_in_document(master_doc, '[BRAND_NAME]', metadata.get('brandName', ''))
    replace_text_in_document(master_doc, '[GENERIC_NAME]', metadata.get('genericName', ''))
    replace_text_in_document(master_doc, '[MANUFACTURER]', metadata.get('manufacturer', ''))
    replace_text_in_document(master_doc, '[VERSION_NUMBER]', metadata.get('versionNumber', ''))
    replace_text_in_document(master_doc, '[COMPILATION_DATE]', metadata.get('compilationDate', ''))

    # Format section headings
    fmt = 'B' if dossier_type == 'format-b' else 'A'
    prod = 'Approved Product' if dossier_type == 'format-b' else 'Unapproved Product / Unapproved Use'
    headings_map = {
        '[HEADING_EXEC_SUMMARY]':     (f"1.0{fmt}", f"1.0{fmt} Executive Summary: Clinical and Economic Value of the {prod}"),
        '[HEADING_PRODUCT_INFO]':     (f"2.0{fmt}", f"2.0{fmt} Product Information and Disease Description"),
        '[HEADING_CLINICAL_EVIDENCE]':(f"3.0{fmt}", f"3.0{fmt} Clinical Evidence"),
        '[HEADING_ECONOMIC_MODEL]':   (f"4.0{fmt}", f"4.0{fmt} Economic Value and Modeling Report"),
        '[HEADING_SUPPORTING_INFO]':  (f"5.0{fmt}", f"5.0{fmt} Additional Supporting Evidence"),
    }
    for p in master_doc.paragraphs:
        txt = p.text.strip()
        if txt in headings_map:
            format_heading(p, *headings_map[txt])

    sections_mapping = {
        'exec_summary':     '[INSERT_EXEC_SUMMARY]',
        'product_info':     '[INSERT_PRODUCT_INFO]',
        'clinical_evidence':'[INSERT_CLINICAL_EVIDENCE]',
        'economic_model':   '[INSERT_ECONOMIC_MODEL]',
        'supporting_info':  '[INSERT_SUPPORTING_INFO]',
    }

    mapped_sections = set()
    temp_files = []

    for section_key, file_path in sections.items():
        if not section_key or section_key not in sections_mapping:
            continue
        if not file_path or not os.path.exists(file_path):
            print(f"Skipping missing file: {file_path}")
            continue

        # Convert PDF → DOCX if needed
        if file_path.lower().endswith('.pdf'):
            docx_path = file_path.replace('.pdf', '_converted.docx')
            try:
                convert_pdf_to_docx(file_path, docx_path)
            except Exception as e:
                print(f"WARNING: PDF conversion failed for {file_path}: {e} — skipping section.")
                continue
            file_path = docx_path

        # Strip header/footer noise and save temp section
        temp_sec = file_path.replace('.docx', '_sec_tmp.docx')
        try:
            prepare_section_document(file_path, temp_sec)
        except Exception as e:
            print(f"WARNING: Section prep failed for {file_path}: {e} — using raw file.")
            shutil.copy2(file_path, temp_sec)
        temp_files.append(temp_sec)

        # Inject into template
        placeholder = sections_mapping[section_key]
        print(f"Injecting {section_key} -> {placeholder}")
        try:
            insert_document_at_placeholder(master_doc, placeholder, temp_sec)
            mapped_sections.add(section_key)
        except Exception as e:
            print(f"WARNING: Injection failed for {section_key}: {e}")

    # Remove placeholders for sections that had no file
    for section_key, placeholder in sections_mapping.items():
        if section_key not in mapped_sections:
            print(f"Removing empty section: {section_key}")
            remove_unused_section(master_doc, placeholder)

    # Link headers/footers across sections
    for idx, section in enumerate(master_doc.sections):
        if idx > 0:
            try:
                section.header.is_linked_to_previous = True
                section.footer.is_linked_to_previous = True
                section.even_page_header.is_linked_to_previous = True
                section.even_page_footer.is_linked_to_previous = True
            except Exception:
                pass

    # Save DOCX
    print(f"Saving dossier to {output_path}")
    master_doc.save(output_path)

    # Remove duplicate page breaks (blank page prevention)
    try:
        remove_duplicate_page_breaks(output_path)
    except Exception as e:
        print(f"WARNING: Page break cleanup failed: {e}")

    # Best-effort TOC/fields update via Word — never blocks generation
    try_update_toc(output_path)

    # Cleanup temp files
    for tf in temp_files:
        try:
            if os.path.exists(tf):
                os.remove(tf)
        except Exception:
            pass

    print("Dossier generation complete.")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python dossier_builder.py <config.json> <output.docx>")
        sys.exit(1)
    compile_dossier(sys.argv[1], sys.argv[2])
