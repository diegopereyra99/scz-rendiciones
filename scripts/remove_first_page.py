import argparse
from pathlib import Path

from PyPDF2 import PdfReader, PdfWriter


def split_pdf_without_first_page(pdf_path: Path, output_root: Path) -> None:
    reader = PdfReader(pdf_path)
    if len(reader.pages) <= 1:
        print(f"Skipping {pdf_path.name}: only {len(reader.pages)} page(s).")
        return

    pdf_output_dir = output_root / pdf_path.stem
    pdf_output_dir.mkdir(parents=True, exist_ok=True)

    # Save each page after the first one as its own single-page PDF.
    for page_idx in range(1, len(reader.pages)):
        writer = PdfWriter()
        writer.add_page(reader.pages[page_idx])

        output_file = pdf_output_dir / f"{pdf_path.stem}_page_{page_idx}.pdf"
        with output_file.open("wb") as f:
            writer.write(f)

    print(f"Processed {pdf_path.name} -> {pdf_output_dir}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Split PDFs in a folder into single-page PDFs, skipping the first page of each."
        )
    )
    parser.add_argument("input_dir", help="Folder containing the source PDFs.")
    parser.add_argument(
        "output_root",
        help="Root folder where per-PDF subfolders with single-page PDFs will be created.",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_root = Path(args.output_root)

    if not input_dir.is_dir():
        raise SystemExit(f"Input directory not found: {input_dir}")

    output_root.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(p for p in input_dir.iterdir() if p.suffix.lower() == ".pdf")
    if not pdf_files:
        print(f"No PDFs found in {input_dir}.")
        return

    for pdf_path in pdf_files:
        split_pdf_without_first_page(pdf_path, output_root)


if __name__ == "__main__":
    main()
