from dataclasses import dataclass
import os
import sys
import docx
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import win32com.client as win32

# Hardcoded title properties
TITLE_FONT_FAMILIES = [".VnCentury SchoolbookH", ".VnTimeH"]
TITLE_FONT_SIZE = 13.0

# jesus fuck, all the LLMs couldn't type this fucking wikipedia table in
TCVN_UNICODE_MAP = [chr(i) if i < 128 else "" for i in range(256)]
TCVN_UNICODE_MAP[1:3] = ["Ú","Ụ"]
TCVN_UNICODE_MAP[4:7] = ["Ừ","Ử","Ữ"]
TCVN_UNICODE_MAP[0x11:0x18] = ["Ứ","Ự","Ỳ","Ỷ","Ỹ","Ý","Ỵ"]
TCVN_UNICODE_MAP[0x80:0xA0] = [
    "À","Ả","Ã","Á","Ạ","Ặ","Ậ","È","Ẻ","Ẽ","É","Ẹ","Ệ","Ì","Ỉ","Ĩ",
    "Í","Ị","Ò","Ỏ","Õ","Ó","Ọ","Ộ","Ờ","Ở","Ỡ","Ớ","Ợ","Ù","Ủ","Ũ"
]
TCVN_UNICODE_MAP[0xA1:0x100] = [
        "Ă","Â","Ê","Ô","Ơ","Ư","Đ","ă","â","ê","ô","ơ","ư","đ","Ằ",
    "`","?","~","'",".","à","ả","ã","á","ạ","Ẳ","ằ","ẳ","ẵ","ắ","Ẵ",
    "Ắ","Ầ","Ẩ","Ẫ","Ấ","Ề","ặ","ầ","ẩ","ẫ","ấ","ậ","è","Ể","ẻ","ẽ",
    "é","ẹ","ề","ể","ễ","ế","ệ","ì","ỉ","Ễ","Ế","Ồ","ĩ","í","ị","ò",
    "Ổ","ỏ","õ","ó","ọ","ồ","ổ","ỗ","ó","ộ","ờ","ở","ỡ","ớ","ợ","ù",
    "Ỗ","ủ","ũ","ú","ụ","ừ","ử","ữ","ứ","ự","ỳ","ỷ","ỹ","ý","ỵ","Ố",
]

all_fonts = set()

def TCVN_Unicode(text):
    return "".join(TCVN_UNICODE_MAP[ord(c)] if ord(c) < 256 else c for c in text)

def build_style_cache(doc):
    """
    Pre-resolves font attributes for all styles in the document.
    Traverses base_style chains so style property lookups are O(1).
    """
    cache = {}

    for style in doc.styles:
        name, size, bold, superscript = None, None, None, None
        curr = style

        # Walk up the base_style chain to extract fallback properties
        while curr is not None:
            # API check
            font = getattr(curr, 'font', None)
            if font:
                if name is None and font.name:
                    name = font.name
                if size is None and font.size is not None:
                    size = font.size.pt
                if bold is None and font.bold is not None:
                    bold = font.bold
                if superscript is None and font.superscript is not None:
                    superscript = font.superscript

            # XML check on style element
            if hasattr(curr, '_element') and curr._element is not None:
                if bold is None:
                    bold = get_xml_bold_flag(curr._element)

            curr = getattr(curr, 'base_style', None)

        cache[style.style_id] = {
            'name': name,
            'size': size,
            'bold': bold if bold is not None else False,
            'superscript': superscript if superscript is not None else False
        }

    return cache

def get_xml_bold_flag(element):
    """
    Helper to extract bold state from any element containing w:rPr
    (works for both w:r and w:pPr/w:p).
    """
    if element is None:
        return None

    # Locate rPr container
    rPr = element.find(qn('w:rPr'))
    if rPr is None and element.tag.endswith('rPr'):
        rPr = element

    if rPr is not None:
        b = rPr.find(qn('w:b'))
        if b is not None:
            val = b.get(qn('w:val'))
            # Explicitly turned off (<w:b w:val="0"/> or <w:b w:val="false"/>)
            if val in ('0', 'false'):
                return False
            # Tag exists with no val or val="1"/"true" -> Bold IS active
            return True

    return None

def process_paragraph(paragraph, style_cache, results):
    """Processes paragraph runs and appends formatted run dictionaries to results."""
    p_style_id = paragraph.style.style_id if paragraph.style else None
    p_info = style_cache.get(p_style_id, {})

    for run in paragraph.runs:
        r_font = run.font
        r_style_id = run.style.style_id if run.style else None
        r_info = style_cache.get(r_style_id, {})

        # Resolve effective font name
        name = r_font.name or r_info.get('name') or p_info.get('name')

        # Resolve effective font size
        size = r_font.size.pt if r_font.size else (r_info.get('size') or p_info.get('size'))

        # Resolve effective bold
        bold = r_font.bold
        if bold is None:
            bold = r_info.get('bold', p_info.get('bold', False))

        # Resolve effective superscript
        superscript = r_font.superscript
        if superscript is None:
            superscript = r_info.get('superscript', p_info.get('superscript', False))

        results.append({
            'text': run.text,
            'font_name': name,
            'font_size': size,
            'bold': bold,
            'superscript': superscript
        })


def process_container(container, style_cache, results):
    """Recursively traverses document blocks (paragraphs and tables)."""
    for block in container.iter_inner_content():
        if hasattr(block, 'runs'):  # Paragraph
            process_paragraph(block, style_cache, results)
        elif hasattr(block, 'rows'):  # Table
            for row in block.rows:
                for cell in row.cells:
                    process_container(cell, style_cache, results)


def get_all_runs(doc):
    """
    Extracts all runs across paragraphs and tables in a single pass.
    Returns a list of dictionaries with 5 keys: text, font_name, font_size, bold, superscript.
    """
    style_cache = build_style_cache(doc)
    results = []
    process_container(doc, style_cache, results)
    return results


@dataclass
class EndnoteRange:
    start: int = 0
    end: int = 0
    def __str__(self) -> str:
        if self.start == self.end:
            return str(self.start)
        return f'{self.start}-{self.end}'

    def merge(self, other):
        if not isinstance(other, EndnoteRange):
            return False
        if other.start == self.end + 1:
            self.end = other.end
            return True
        return False


@dataclass
class TitleText:
    text: str = ""
    def __str__(self) -> str:
        return "\n" + self.text

    def merge(self, other):
        if not isinstance(other, TitleText):
            return False
        self.text += " " + other.text
        return True


def markers_append(ms, m):
    if len(ms) == 0 or not ms[-1].merge(m):
        ms.append(m)


def extract_markers(docx_path):
    """Parses a .docx file using python-docx to quickly extract markers."""
    markers = []
    doc = docx.Document(docx_path)

    for run in get_all_runs(doc):
        all_fonts.add(run['font_name'])
        text = run['text'].strip()
        if run['superscript'] and text.isdigit():
            num = int(text.replace(" ", ""))
            markers_append(markers, EndnoteRange(num, num))
            continue

        if (run['font_size'] == TITLE_FONT_SIZE
            and run['font_name'] in TITLE_FONT_FAMILIES):
            markers_append(markers, TitleText(TCVN_Unicode(text)))
            continue

    return markers


def process_single_doc(word, abs_doc_path):
    try:
        base_name = os.path.basename(abs_doc_path)
        temp_docx_path = os.path.splitext(abs_doc_path)[0] + "_temp_converted.docx"

        # Fast conversion via Word COM
        doc = word.Documents.Open(abs_doc_path)
        doc.SaveAs2(temp_docx_path, FileFormat=16) # 16 = wdFormatXMLDocument
        doc.Close()

        # Parsing using python-docx (COM is slow)
        markers = extract_markers(temp_docx_path)

        # Print result
        matches_str = ", ".join(str(x) for x in markers)
        print(f"\n{base_name}: {matches_str}")
    except Exception as e:
        print(f"\n{base_name}: Error processing file ({e})")

    finally:
        if os.path.exists(temp_docx_path):
            try:
                os.remove(temp_docx_path)
            except Exception:
                print(f"Failed to remove temp file {temp_docx_path}")


def process_path(target_path):
    """Handles both individual file paths and folder paths."""
    abs_path = os.path.abspath(target_path)

    if not os.path.exists(abs_path):
        print(f"Error: Path '{abs_path}' does not exist.")
        return

    # Collect .doc files to process
    doc_files = []
    if os.path.isfile(abs_path):
        if abs_path.lower().endswith(".doc") and not abs_path.lower().endswith(".docx"):
            doc_files.append(abs_path)
        else:
            print(f"Error: Provided file '{abs_path}' is not a .doc file.")
            return
    elif os.path.isdir(abs_path):
        for root, _, files in os.walk(abs_path):
            for file in files:
                # Filter for .doc (excluding .docx and temporary ~$ files)
                if file.lower().endswith(".doc") and not file.startswith("~$"):
                    doc_files.append(os.path.join(root, file))

    if not doc_files:
        print("No .doc files found to process.")
        return

    doc_files.sort()

    # Start Word application once for all conversions
    word = None
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False

    # Suppress pop-ups, warnings, and macro prompts
    word.DisplayAlerts = 0  # 0 = wdAlertsNone
    word.AutomationSecurity = 3  # 3 = msoAutomationSecurityForceDisable


    try:
        for file_path in doc_files:
            process_single_doc(word, file_path)
    except Exception as e:
        print(f"Word Automation Error: {e}")
    finally:
        print("\n----------\nAll fonts:")
        print("".join([str(x)+"\n" for x in all_fonts]))
        if word is not None:
            word.Quit()


if __name__ == "__main__":
    # Example 1: Pass path as command line argument, e.g. python script.py "C:\MyDocs"
    if len(sys.argv) > 1:
        process_path(sys.argv[1])
    else:
        print("Syntax: `markers.py a.doc` or `markers.py docfolder`")
