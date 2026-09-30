"""
Either:
  - Split a double-column PDF into single-column pages, or
  - Stitch (extract a page range from input and append it to another PDF).

Usage examples:
    # Split mode
    python pdfsplit.py input.pdf -o output.pdf
    python pdfsplit.py input.pdf -o output.pdf -ra -rb

    # Stitch mode
    python pdfsplit.py input.pdf -o output.pdf -rr 12 13 other.pdf
"""

import argparse
import sys
import os
import subprocess
from pathlib import Path
import pymupdf
import shutil

# ============================================================
# Hardcoded extra margin cuts (in centimetres)
# Applied only when doing double-column splitting
# ============================================================
MARGIN_TOP_CM    = 1.2   # cut from top
MARGIN_BOTTOM_CM = 1.0   # cut from bottom
MARGIN_INNER_CM  = 0.0   # cut from the inner edge
                         # (right side of left half-page, left side of right half-page)
MARGIN_OUTER_CM  = 1.2   # cut from the outer edge
                         # (left side of left half-page, right side of right half-page)
# ============================================================


def cm_to_pt(cm: float) -> float:
    """Convert centimetres to PDF points (1 inch = 2.54 cm = 72 pt)."""
    return cm * 72 / 2.54


def make_temp_path(target: str) -> str:
    """Create a temporary filename in the same directory as target."""
    p = Path(target)
    return str(p.with_name(p.stem + ".tmp" + p.suffix))


def find_ghostscript_executable() -> str:
    """
    Finds the Ghostscript executable path.
    Checks Windows default command names/paths first, then Linux/macOS standard executable.
    """
    # Candidate executables in order of preference:
    # 1. Windows 64-bit CLI executable
    # 2. Windows 32-bit CLI executable
    # 3. Linux/macOS/Standard PATH executable
    candidates = ["gswin64c", "gswin32c", "gs"]

    for executable in candidates:
        path = shutil.which(executable)
        if path:
            return path

    raise FileNotFoundError(
        "Ghostscript executable not found. Looked for: " + ", ".join(candidates)
    )


def run_ghostscript(input_pdf: str, output_pdf: str):
    """Aggressively clean/optimize PDF with Ghostscript."""
    try:
        gs_path = find_ghostscript_executable()
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    cmd = [
        gs_path,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/prepress",
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        "-dDetectDuplicateImages=true",
        "-dNOPAUSE",
        "-dBATCH",
        "-dQUIET",
        f"-sOutputFile={output_pdf}",
        input_pdf,
    ]

    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error: Ghostscript failed with code {e.returncode}", file=sys.stderr)
        sys.exit(1)


def split_double_column(input_pdf: str, output_pdf: str,
                        remove_first_left: bool = False,
                        remove_last_right: bool = False):
    src = pymupdf.open(input_pdf)
    dst = pymupdf.open()

    total = len(src)

    top_pt    = cm_to_pt(MARGIN_TOP_CM)
    bottom_pt = cm_to_pt(MARGIN_BOTTOM_CM)
    inner_pt  = cm_to_pt(MARGIN_INNER_CM)
    outer_pt  = cm_to_pt(MARGIN_OUTER_CM)

    for page_num in range(total):
        page = src[page_num]
        rect = page.rect
        mid = rect.width / 2

        left = pymupdf.Rect(
            rect.x0 + outer_pt,
            rect.y0 + top_pt,
            mid     - inner_pt,
            rect.y1 - bottom_pt
        )

        right = pymupdf.Rect(
            mid     + inner_pt,
            rect.y0 + top_pt,
            rect.x1 - outer_pt,
            rect.y1 - bottom_pt
        )

        keep_left  = not (remove_first_left  and page_num == 0)
        keep_right = not (remove_last_right and page_num == total - 1)

        if keep_left and left.width > 0 and left.height > 0:
            new_page = dst.new_page(width=left.width, height=left.height)
            new_page.show_pdf_page(new_page.rect, src, page_num, clip=left)

        if keep_right and right.width > 0 and right.height > 0:
            new_page = dst.new_page(width=right.width, height=right.height)
            new_page.show_pdf_page(new_page.rect, src, page_num, clip=right)

    if len(dst) == 0:
        print("Error: no pages left after applying -ra / -rb", file=sys.stderr)
        src.close()
        dst.close()
        sys.exit(1)

    tmp_name = make_temp_path(output_pdf)
    dst.save(tmp_name, garbage=4, deflate=True)
    src.close()
    dst.close()

    os.replace(tmp_name, output_pdf)
    print(f"Saved: {output_pdf}")


def stitch(input_pdf: str, left: int, right: int, appendto: str, output_pdf: str):
    """
    Take pages left..right (1-based) from input_pdf,
    append them to the content of appendto,
    and write the result to output_pdf.
    Never modifies input_pdf or appendto.
    """
    src = pymupdf.open(input_pdf)
    total = len(src)

    start = left - 1
    end   = right - 1

    if start < 0 or end >= total or start > end:
        print(f"Error: invalid page range {left}-{right}. "
              f"Document has {total} pages (1-based).", file=sys.stderr)
        src.close()
        sys.exit(1)

    # Start from a copy of appendto (or empty if it doesn't exist)
    if os.path.exists(appendto):
        dst = pymupdf.open(appendto)
    else:
        dst = pymupdf.open()

    # Append the requested pages from input
    dst.insert_pdf(src, from_page=start, to_page=end)

    tmp_name = make_temp_path(output_pdf)
    dst.save(tmp_name, garbage=4, deflate=True)
    src.close()
    dst.close()

    os.replace(tmp_name, output_pdf)
    print(f"Stitched pages {left}-{right} from {input_pdf} + {appendto} → {output_pdf}")


def main():
    parser = argparse.ArgumentParser(
        description="Split double-column PDFs or stitch a page range onto another PDF."
    )
    parser.add_argument("input", help="Input PDF file")
    parser.add_argument("-o", "--output", required=True, help="Output PDF file")
    parser.add_argument("-ra", action="store_true",
                        help="Remove the left half of the first page (split mode only)")
    parser.add_argument("-rb", action="store_true",
                        help="Remove the right half of the last page (split mode only)")
    parser.add_argument("-rr", nargs=3, metavar=("LEFT", "RIGHT", "APPENDTO"),
                        help="Stitch mode: extract pages LEFT-RIGHT (1-based) from input "
                             "and append them to APPENDTO, writing result to -o")
    parser.add_argument("--nogs", action="store_true",
                        help="Disable Ghostscript cleanup (stitch mode only)")

    args = parser.parse_args()

    # ---------- Stitch mode ----------
    if args.rr is not None:
        try:
            left  = int(args.rr[0])
            right = int(args.rr[1])
        except ValueError:
            print("Error: LEFT and RIGHT must be integers", file=sys.stderr)
            sys.exit(1)

        appendto = args.rr[2]
        stitch(args.input, left, right, appendto, args.output)

        if not args.nogs:
            tmp_name = make_temp_path(args.output)
            run_ghostscript(args.output, tmp_name)
            os.replace(tmp_name, args.output)
            print(f"Ghostscript cleaned: {args.output}")
        return

    # ---------- Split mode ----------
    split_double_column(
        input_pdf=args.input,
        output_pdf=args.output,
        remove_first_left=args.ra,
        remove_last_right=args.rb,
    )


if __name__ == "__main__":
    main()