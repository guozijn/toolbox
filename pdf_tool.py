"""
Command line utilities for working with PDF files.

Current sub-commands:
    merge      Merge all PDFs in the target directory (non-recursive) into a single file.
    extract    Extract the specified pages from a single PDF into a new PDF.
"""
from __future__ import annotations

import argparse
import sys
import shutil
import subprocess
import tempfile
import re
from pathlib import Path
from typing import Iterable, List

from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas


def natural_key(value: str) -> List[object]:
    """
    Split text into case-insensitive chunks so numbers are sorted numerically.
    """
    return [
        int(chunk) if chunk.isdigit() else chunk.lower()
        for chunk in re.split(r"(\d+)", value)
        if chunk
    ]


def find_pdfs(directory: Path) -> List[Path]:
    """
    Return a sorted list of PDF files directly inside `directory`.
    """
    pdfs = [
        path
        for path in directory.iterdir()
        if path.is_file() and path.suffix.lower() == ".pdf"
    ]
    return sorted(pdfs, key=lambda path: natural_key(path.name))


def find_pptx(directory: Path) -> List[Path]:
    """
    Return a sorted list of PPTX files directly inside `directory`.
    """
    pptx = [
        path
        for path in directory.iterdir()
        if path.is_file() and path.suffix.lower() == ".pptx"
    ]
    return sorted(pptx, key=lambda path: natural_key(path.name))


def merge_pdfs(input_files: Iterable[Path], output_path: Path) -> Path:
    """
    Merge `input_files` into `output_path`.
    """
    writer = PdfWriter()
    for pdf_path in input_files:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as out_handle:
        writer.write(out_handle)
    writer.close()

    return output_path


def convert_pptx_files(
    pptx_files: Iterable[Path], output_dir: Path, soffice_path: str | None
) -> List[tuple[str, Path]]:
    """
    Convert PPTX files to PDFs using LibreOffice/soffice.
    Returns a list of tuples (original_name, converted_pdf_path).
    """
    soffice_binary = (
        soffice_path
        or shutil.which("soffice")
        or shutil.which("libreoffice")
    )
    if not soffice_binary:
        raise RuntimeError(
            "LibreOffice (soffice) executable not found. "
            "Install LibreOffice or pass --soffice-path."
        )

    converted: List[tuple[str, Path]] = []
    for pptx_path in pptx_files:
        target_path = output_dir / f"{pptx_path.stem}.pdf"
        command = [
            soffice_binary,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(pptx_path),
        ]
        try:
            subprocess.run(
                command,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
        except subprocess.CalledProcessError as exc:
            stderr = exc.stderr.decode(errors="ignore").strip()
            raise RuntimeError(
                f"Failed to convert {pptx_path.name}: {stderr or exc}"
            ) from exc

        if not target_path.exists():
            raise RuntimeError(
                f"Conversion did not produce output for {pptx_path.name}"
            )

        converted.append((pptx_path.name, target_path))
    return converted


def create_title_page(title: str, destination: Path) -> Path:
    """
    Create a single-page PDF containing `title` centered on the page.
    """
    def wrap_text(
        text: str, font_name: str, font_size: float, max_width: float
    ) -> List[str]:
        lines: List[str] = []
        words = text.split()
        if not words:
            return ["Untitled"]

        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            width = pdfmetrics.stringWidth(candidate, font_name, font_size)
            if width <= max_width or not current:
                current = candidate
            else:
                lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines or ["Untitled"]

    title_text = title.strip() or "Untitled"
    c = canvas.Canvas(str(destination), pagesize=letter)
    width, height = letter
    font_name = "Helvetica-Bold"
    font_size = 24
    max_width = width * 0.8
    margin = height * 0.15

    lines = wrap_text(title_text, font_name, font_size, max_width)

    def recompute_line_metrics(
        current_lines: List[str], current_font_size: float
    ) -> tuple[float, float]:
        max_line_width = max(
            pdfmetrics.stringWidth(line, font_name, current_font_size)
            for line in current_lines
        )
        leading = current_font_size * 1.2
        total_height = len(current_lines) * leading
        return max_line_width, total_height

    max_line_width, total_height = recompute_line_metrics(lines, font_size)
    available_height = height - 2 * margin

    while (
        (max_line_width > max_width or total_height > available_height)
        and font_size > 10
    ):
        font_size -= 2
        lines = wrap_text(title_text, font_name, font_size, max_width)
        max_line_width, total_height = recompute_line_metrics(lines, font_size)

    leading = font_size * 1.2
    start_y = max(
        margin + total_height - leading,
        (height + total_height) / 2 - leading,
    )

    c.setFont(font_name, font_size)
    for index, line in enumerate(lines):
        y = start_y - index * leading
        if y < margin:
            break
        c.drawCentredString(width / 2, y, line)
    c.showPage()
    c.save()
    return destination


def parse_page_ranges(spec: str, total_pages: int | None = None) -> List[int]:
    """
    Parse a page selection specification like "1,3-5" into a list of 1-based page numbers.
    """
    if not spec:
        raise ValueError("Page specification cannot be empty")

    pages: List[int] = []
    for raw_part in spec.split(","):
        part = raw_part.strip()
        if not part:
            continue
        if "-" in part:
            start_str, end_str = part.split("-", maxsplit=1)
            try:
                start = int(start_str)
                end = int(end_str)
            except ValueError as exc:
                raise ValueError(f"Invalid page range: {part}") from exc
            if start <= 0 or end <= 0:
                raise ValueError("Page numbers must be positive integers")
            if start > end:
                raise ValueError(f"Range start greater than end: {part}")
            for page in range(start, end + 1):
                pages.append(page)
        else:
            try:
                page = int(part)
            except ValueError as exc:
                raise ValueError(f"Invalid page number: {part}") from exc
            if page <= 0:
                raise ValueError("Page numbers must be positive integers")
            pages.append(page)

    if not pages:
        raise ValueError("No valid pages found in specification")

    if total_pages is not None:
        for page in pages:
            if page > total_pages:
                raise ValueError(
                    f"Requested page {page} exceeds document length ({total_pages} pages)"
                )

    return pages


def handle_extract(args: argparse.Namespace) -> int:
    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() != ".pdf":
        raise ValueError("Input file must be a PDF")

    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = input_path.parent / output_path
    output_path = output_path.resolve()

    reader = PdfReader(str(input_path))
    total_pages = len(reader.pages)
    selected_pages = parse_page_ranges(args.pages, total_pages=total_pages)

    writer = PdfWriter()
    for page_number in selected_pages:
        writer.add_page(reader.pages[page_number - 1])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as handle:
        writer.write(handle)
    writer.close()

    print(
        f"Extracted {len(selected_pages)} pages from {input_path.name} into {output_path}"
    )
    return 0


def handle_merge(args: argparse.Namespace) -> int:
    target_dir = Path(args.directory).expanduser().resolve()
    if not target_dir.exists():
        raise FileNotFoundError(f"Directory not found: {target_dir}")
    if not target_dir.is_dir():
        raise NotADirectoryError(f"Not a directory: {target_dir}")

    input_pdfs = find_pdfs(target_dir)
    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = target_dir / output_path
    output_path = output_path.resolve()

    # Avoid reading the output file as an input if it already exists.
    input_pdfs = [path for path in input_pdfs if path.resolve() != output_path]

    documents: List[tuple[str, Path]] = [(pdf.name, pdf) for pdf in input_pdfs]

    temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None
    temp_path: Path | None = None
    total_documents = 0
    try:
        if args.include_pptx or args.prepend_titles:
            temp_dir_obj = tempfile.TemporaryDirectory()
            temp_path = Path(temp_dir_obj.name)

        if args.include_pptx:
            pptx_files = find_pptx(target_dir)
            if pptx_files:
                if temp_path is None:
                    temp_dir_obj = tempfile.TemporaryDirectory()
                    temp_path = Path(temp_dir_obj.name)
                converted = convert_pptx_files(
                    pptx_files,
                    temp_path,
                    args.soffice_path,
                )
                documents.extend(converted)

        if not documents:
            raise ValueError(
                f"No PDF files found in {target_dir}"
                + (" or PPTX files to convert" if args.include_pptx else "")
            )

        documents.sort(key=lambda item: natural_key(item[0]))
        total_documents = len(documents)
        merge_inputs: List[Path] = []
        for index, (name, path) in enumerate(documents, start=1):
            if args.prepend_titles:
                if temp_path is None:
                    temp_dir_obj = tempfile.TemporaryDirectory()
                    temp_path = Path(temp_dir_obj.name)
                title_path = temp_path / f"title_{index:04d}.pdf"
                create_title_page(name, title_path)
                merge_inputs.append(title_path)
            merge_inputs.append(path)

        merged_path = merge_pdfs(merge_inputs, output_path)
    finally:
        if temp_dir_obj is not None:
            temp_dir_obj.cleanup()

    print(f"Merged {total_documents} documents into {merged_path}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Utility CLI for common PDF operations."
    )
    subparsers = parser.add_subparsers(dest="command")

    merge_parser = subparsers.add_parser(
        "merge",
        help="Merge all PDF files in the given directory (non-recursive).",
    )
    merge_parser.add_argument(
        "-d",
        "--directory",
        default=".",
        help="Directory to scan for PDF files (default: current directory).",
    )
    merge_parser.add_argument(
        "-o",
        "--output",
        default="merged.pdf",
        help="Output PDF file path (default: merged.pdf in the target directory).",
    )
    merge_parser.add_argument(
        "--include-pptx",
        action="store_true",
        help="Include PPTX files by converting them to PDF (requires LibreOffice/soffice).",
    )
    merge_parser.add_argument(
        "--prepend-titles",
        action="store_true",
        help="Insert a title page showing the source filename before each merged document.",
    )
    merge_parser.add_argument(
        "--soffice-path",
        help="Explicit path to the LibreOffice 'soffice' executable.",
    )
    merge_parser.set_defaults(func=handle_merge)

    extract_parser = subparsers.add_parser(
        "extract",
        help="Extract specific pages from a PDF and save them to a new file.",
    )
    extract_parser.add_argument(
        "input",
        help="Path to the source PDF file.",
    )
    extract_parser.add_argument(
        "-p",
        "--pages",
        required=True,
        help="Page selection (1-based) using commas and ranges, e.g. '1,3-5,8'.",
    )
    extract_parser.add_argument(
        "-o",
        "--output",
        default="extracted.pdf",
        help="Output PDF filepath (default: extracted.pdf next to the source file).",
    )
    extract_parser.set_defaults(func=handle_extract)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not hasattr(args, "func"):
        parser.print_help()
        return 0

    try:
        return args.func(args)
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
