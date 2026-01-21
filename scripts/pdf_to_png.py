#!/usr/bin/env python3
"""Convierte PDFs de una página a PNG con el lado mayor fijado."""

import argparse
from pathlib import Path


def convert_pdf(pdf_path: Path, out_dir: Path, max_side: int) -> Path:
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:  # noqa: ICN001
        raise RuntimeError("PyMuPDF no está instalado") from exc

    with fitz.open(pdf_path) as doc:
        if doc.page_count < 1:
            raise ValueError(f"Sin páginas: {pdf_path}")
        page = doc[0]
        width, height = page.rect.width, page.rect.height
        scale = max_side / max(width, height)
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)

    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{pdf_path.stem}.png"
    pix.save(out_path)
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convierte PDFs de una página a PNG con resolución fija.",
    )
    parser.add_argument(
        "-i",
        "--input-dir",
        default="estados/pdf/",
        help="Carpeta con PDFs de entrada (default: estados/pdf/)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="estados/png/",
        help="Carpeta destino para PNGs (default: estados/png/)",
    )
    parser.add_argument(
        "-m",
        "--max-side",
        type=int,
        default=2048,
        help="Lado mayor en píxeles del PNG (default: 2048)",
    )
    args = parser.parse_args()

    in_dir = Path(args.input_dir)
    out_dir = Path(args.output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {in_dir}")

    pdfs = sorted(in_dir.glob("*.pdf"))
    if not pdfs:
        print(f"No se encontraron PDFs en {in_dir}")
        return

    for pdf in pdfs:
        try:
            out_path = convert_pdf(pdf, out_dir, args.max_side)
            print(f"[OK] {pdf.name} -> {out_path}")
        except Exception as exc:  # noqa: BLE001
            print(f"[ERR] {pdf.name}: {exc}")


if __name__ == "__main__":
    main()
