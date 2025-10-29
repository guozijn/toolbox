"""
Command line utilities for working with PDF files.

Current sub-commands:
    merge    Merge all PDFs in the target directory (non-recursive) into a single file.
"""
from __future__ import annotations

import argparse
import sys
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Iterable, List

from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


def find_pdfs(directory: Path) -> List[Path]:
    """
    Return a sorted list of PDF files directly inside `directory`.
    """
    pdfs = [
        path
        for path in directory.iterdir()
        if path.is_file() and path.suffix.lower() == ".pdf"
    ]
    return sorted(pdfs, key=lambda path: path.name.lower())


def find_pptx(directory: Path) -> List[Path]:
    """
    Return a sorted list of PPTX files directly inside `directory`.
    """
    pptx = [
        path
        for path in directory.iterdir()
        if path.is_file() and path.suffix.lower() == ".pptx"
    ]
    return sorted(pptx, key=lambda path: path.name.lower())


def merge_pdfs(input_files: Iterable[Path], output_path: Path) -> Path:
    """
    Merge `input_files` into `output_path`.
    """
    writer = PdfWriter()
    for pdf_path in input_files:
        with pdf_path.open("rb") as handle:
            reader = PdfReader(handle)
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
    title_text = title.strip() or "Untitled"
    c = canvas.Canvas(str(destination), pagesize=letter)
    width, height = letter
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(width / 2, height / 2, title_text)
    c.showPage()
    c.save()
    return destination


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

        documents.sort(key=lambda item: item[0].lower())
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
